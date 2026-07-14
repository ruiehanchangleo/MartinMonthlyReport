#!/usr/bin/env python3
"""
send_via_mail.py — send email via macOS Mail.app (osascript) AND verify delivery.

Reliable replacement for the Outlook AppleScript send path, which silently parks
messages in Drafts when Outlook is in "New Outlook" mode. Mail.app actually
delivers. Verification is real: after sending we confirm the message count in
Mail's *Sent* mailbox INCREASED (rc=0 from osascript alone does NOT prove a send
left the machine — it can sit in the Outbox).

Canonical copy lives in the XTM-glossary repo; an identical copy is dropped into
each mail-sending project so each stays self-contained. Keep copies in sync.

Import:
    from send_via_mail import send_mail
    ok, detail = send_mail(["a@x.com"], "Subject", "Body text",
                           sender="leo.chang@familysearch.org",
                           attachments=["/path/report.xlsx"])

CLI:
    send_via_mail.py --to a@x.com [--to b@y.com] --subject "S" --body-file body.txt \
        [--sender leo.chang@familysearch.org] [--attach /path ...] [--verify-timeout 60]
    echo "body" | send_via_mail.py --to a@x.com --subject "S"      # body from stdin
Exits 0 only on a verified send.
"""
import subprocess, os, sys, time, tempfile, argparse

DEFAULT_SENDER = "leo.chang@familysearch.org"


def _esc(s):
    return str(s).replace("\\", "\\\\").replace('"', '\\"')


def _sent_count(subject, window_secs=900):
    """Count recent messages in Mail's Sent mailbox with this exact subject."""
    q = ('tell application "Mail"\n'
         f'  set cutoff to (current date) - {int(window_secs)}\n'
         f'  return count (messages of sent mailbox whose subject is "{_esc(subject)}" and date sent > cutoff)\n'
         'end tell')
    try:
        r = subprocess.run(["osascript", "-e", q], capture_output=True, text=True, timeout=30)
    except subprocess.TimeoutExpired:
        return None
    out = (r.stdout or "").strip()
    if r.returncode == 0 and out.lstrip("-").isdigit():
        return int(out)
    return None


def send_mail(to, subject, body, sender=DEFAULT_SENDER, attachments=None,
              verify_timeout=60, draft_only=False):
    """Send via Mail.app and verify it reached Sent. Returns (ok: bool, detail: str).

    draft_only=True composes a visible draft and does NOT send or verify (dry-run/review).
    """
    if isinstance(to, str):
        to = [to]
    to = [a for a in to if a]
    if not to:
        return False, "no recipients"
    attachments = attachments or []

    baseline = 0 if draft_only else _sent_count(subject)
    if baseline is None:
        baseline = 0  # can't read Sent; fall back to "any appears" semantics below

    bf = tempfile.NamedTemporaryFile("w", suffix=".txt", delete=False, encoding="utf-8")
    bf.write(body)
    bf.close()

    rcpts = "\n".join(
        f'        make new to recipient at end of to recipients with properties {{address:"{_esc(a)}"}}'
        for a in to
    )
    atts = ""
    if attachments:
        lines = [
            f'        make new attachment with properties {{file name:(POSIX file "{_esc(os.path.abspath(p))}")}} at after the last paragraph'
            for p in attachments
        ]
        atts = "    tell content of m\n" + "\n".join(lines) + "\n    end tell\n"

    script = (
        'tell application "Mail"\n'
        f'    set theBody to (read (POSIX file "{_esc(bf.name)}") as «class utf8»)\n'
        f'    set m to make new outgoing message with properties {{subject:"{_esc(subject)}", content:theBody, visible:{"true" if draft_only else "false"}}}\n'
        '    tell m\n'
        f'        set sender to "{_esc(sender)}"\n'
        f'{rcpts}\n'
        '    end tell\n'
        f'{atts}'
        f'{"" if draft_only else "    tell m to send" + chr(10)}'
        'end tell'
    )
    sf = tempfile.NamedTemporaryFile("w", suffix=".applescript", delete=False, encoding="utf-8")
    sf.write(script)
    sf.close()

    try:
        r = subprocess.run(["osascript", sf.name], capture_output=True, text=True, timeout=90)
        if r.returncode != 0:
            return False, f"osascript send failed: {(r.stderr or '').strip()[:200]}"
    except subprocess.TimeoutExpired:
        return False, "osascript send timed out (Mail wedged?)"
    finally:
        for f in (bf.name, sf.name):
            try:
                os.unlink(f)
            except OSError:
                pass

    if draft_only:
        return True, "draft composed for review (not sent)"

    deadline = time.time() + verify_timeout
    while time.time() < deadline:
        cur = _sent_count(subject)
        if cur is not None and cur > baseline:
            return True, "sent + verified in Sent mailbox"
        time.sleep(3)
    return False, (f"send issued but not confirmed in Sent within {verify_timeout}s "
                   "(likely stuck in Mail Outbox — check the account's send/auth)")


def _main(argv=None):
    ap = argparse.ArgumentParser(description="Send email via Mail.app and verify delivery.")
    ap.add_argument("--to", action="append", required=True, help="recipient (repeatable)")
    ap.add_argument("--subject", required=True)
    ap.add_argument("--body-file", help="path to body text file; if omitted, read stdin")
    ap.add_argument("--sender", default=DEFAULT_SENDER)
    ap.add_argument("--attach", action="append", default=[], help="attachment path (repeatable)")
    ap.add_argument("--verify-timeout", type=int, default=60)
    ap.add_argument("--draft-only", action="store_true", help="compose a draft for review; do not send")
    a = ap.parse_args(argv)
    body = open(a.body_file, encoding="utf-8").read() if a.body_file else sys.stdin.read()
    ok, detail = send_mail(a.to, a.subject, body, sender=a.sender,
                           attachments=a.attach, verify_timeout=a.verify_timeout,
                           draft_only=a.draft_only)
    print(("OK: " if ok else "FAIL: ") + detail)
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(_main())
