#!/usr/bin/env python3
"""
send_via_mail.py — send email via macOS Mail.app (osascript) AND verify delivery.

Reliable replacement for the Outlook AppleScript send path, which silently parks
messages in Drafts when Outlook is in "New Outlook" mode. Mail.app actually
delivers. Verification is real: after sending we confirm the message count in
Mail's *Sent* mailbox INCREASED (rc=0 from osascript alone does NOT prove a send
left the machine — it can sit in the Outbox).

SELF-HEALING (added 2026-07-15): every send first probes Mail.app with a short
AppleScript. If Mail is wedged (a hung process makes ALL Apple Events time out —
this silently broke every scheduled email job for ~5 days), we force-restart
Mail and retry once. A send that times out mid-flight also triggers one restart
+ retry. This is capped at a single restart per send call to avoid restart
storms. Mail is only ever killed when it is already unresponsive; a healthy Mail
is never touched.

Canonical copy lives in the XTM-glossary repo; an identical copy is dropped into
each mail-sending project so each stays self-contained. Keep copies in sync
(all four are byte-identical — deploy changes to every copy).

Import:
    from send_via_mail import send_mail
    ok, detail = send_mail(["a@x.com"], "Subject", "Body text",
                           sender="leo.chang@familysearch.org",
                           attachments=["/path/report.xlsx"])

CLI:
    send_via_mail.py --to a@x.com [--to b@y.com] --subject "S" --body-file body.txt \
        [--sender leo.chang@familysearch.org] [--attach /path ...] [--verify-timeout 60]
    echo "body" | send_via_mail.py --to a@x.com --subject "S"      # body from stdin
    send_via_mail.py --health-check    # probe Mail.app; restart if wedged; exit 0 if healthy
Exits 0 only on a verified send.
"""
import subprocess, os, sys, time, tempfile, argparse

DEFAULT_SENDER = "leo.chang@familysearch.org"


def _esc(s):
    return str(s).replace("\\", "\\\\").replace('"', '\\"')


def mail_status(timeout=12):
    """Probe Mail.app with a trivial Apple Event. Returns (status, detail):

      "healthy" — answered normally.
      "denied"  — TCC Automation permission missing/revoked (osascript -1743 /
                  "Not authorized to send Apple events"). A RESTART WILL NOT FIX
                  THIS — the user must re-approve in System Settings > Privacy &
                  Security > Automation. macOS shows this prompt once per client
                  binary; a background launchd job can't answer it, so it must be
                  granted from an interactive run first.
      "wedged"  — hung (no reply within timeout) or another transient error; a
                  restart is worth trying.
    """
    try:
        r = subprocess.run(
            ["osascript", "-e", 'tell application "Mail" to get name of first account'],
            capture_output=True, text=True, timeout=timeout)
    except subprocess.TimeoutExpired:
        return "wedged", f"no reply within {timeout}s"
    if r.returncode == 0 and (r.stdout or "").strip():
        return "healthy", (r.stdout or "").strip()
    err = (r.stderr or "").strip()
    low = err.lower()
    if "-1743" in err or "not authorized" in low or "assistive access" in low:
        return "denied", err[:200] or "not authorized to send Apple events"
    return "wedged", err[:200] or f"unexpected rc={r.returncode}"


def mail_healthy(timeout=12):
    """Back-compat bool: True only if Mail answered normally."""
    return mail_status(timeout)[0] == "healthy"


def restart_mail(wait=45):
    """Force-restart Mail.app and wait until it answers AppleScript. Returns bool.

    Graceful quit first (short timeout — a wedged Mail ignores it), then SIGKILL,
    then relaunch and poll mail_healthy(). Only call when Mail is already unhealthy.
    """
    try:
        subprocess.run(["osascript", "-e", 'tell application "Mail" to quit'],
                       capture_output=True, text=True, timeout=8)
    except subprocess.TimeoutExpired:
        pass
    subprocess.run(["pkill", "-9", "-x", "Mail"], capture_output=True)
    time.sleep(2)
    subprocess.run(["open", "-a", "Mail"], capture_output=True)
    deadline = time.time() + wait
    while time.time() < deadline:
        if mail_healthy():
            return True
        time.sleep(3)
    return mail_healthy()


def _ensure_mail_healthy():
    """Probe Mail; restart once ONLY if wedged. Returns (ok, status, detail, restarted).

    A "denied" status is a permission problem a restart can't fix, so we return
    immediately without killing the user's Mail.
    """
    status, detail = mail_status()
    if status in ("healthy", "denied"):
        return status == "healthy", status, detail, False
    # wedged -> restart once
    healthy = restart_mail()
    if healthy:
        return True, "healthy", "recovered by restart", True
    status, detail = mail_status()
    return status == "healthy", status, detail, True


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


def _run_send(sf_name):
    """Run the compiled send AppleScript. Returns (status, detail).

    status is one of: "ok", "timeout" (Mail wedged — recoverable via restart),
    "error" (script rejected — a restart won't help).
    """
    try:
        r = subprocess.run(["osascript", sf_name], capture_output=True, text=True, timeout=90)
    except subprocess.TimeoutExpired:
        return "timeout", "osascript send timed out (Mail wedged?)"
    if r.returncode != 0:
        return "error", f"osascript send failed: {(r.stderr or '').strip()[:200]}"
    return "ok", ""


def send_mail(to, subject, body, sender=DEFAULT_SENDER, attachments=None,
              verify_timeout=60, draft_only=False):
    """Send via Mail.app and verify it reached Sent. Returns (ok: bool, detail: str).

    Self-heals a wedged Mail.app (see module docstring). draft_only=True composes
    a visible draft and does NOT send or verify (dry-run/review).
    """
    if isinstance(to, str):
        to = [to]
    to = [a for a in to if a]
    if not to:
        return False, "no recipients"
    attachments = attachments or []

    # Pre-flight: make sure Mail can actually be driven before we compute a
    # baseline or issue the send. Recovers the "wedged since boot" case, and
    # reports a permission denial distinctly (a restart can't fix that).
    ok, status, detail, restarted = _ensure_mail_healthy()
    if not ok:
        if status == "denied":
            return False, ("Mail.app Automation permission denied (" + detail + "). Grant the "
                           "sending process control of Mail in System Settings > Privacy & "
                           "Security > Automation, from an interactive run (background jobs "
                           "can't answer the prompt).")
        return False, "Mail.app unresponsive and restart failed (" + detail + ")"

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
        status, detail = _run_send(sf.name)
        # A mid-flight hang means Mail wedged after our pre-flight check. Restart
        # once (unless we already did) and retry the send.
        if status == "timeout" and not restarted:
            restarted = restart_mail()
            if restarted:
                status, detail = _run_send(sf.name)
        if status != "ok":
            return False, detail
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
    ap.add_argument("--to", action="append", help="recipient (repeatable)")
    ap.add_argument("--subject")
    ap.add_argument("--body-file", help="path to body text file; if omitted, read stdin")
    ap.add_argument("--sender", default=DEFAULT_SENDER)
    ap.add_argument("--attach", action="append", default=[], help="attachment path (repeatable)")
    ap.add_argument("--verify-timeout", type=int, default=60)
    ap.add_argument("--draft-only", action="store_true", help="compose a draft for review; do not send")
    ap.add_argument("--health-check", action="store_true",
                    help="probe Mail.app (restart if wedged) and exit; 0=healthy, 2=permission denied, 1=wedged")
    a = ap.parse_args(argv)

    if a.health_check:
        ok, status, detail, _ = _ensure_mail_healthy()
        msg = {
            "healthy": "Mail.app healthy",
            "denied": "Mail.app Automation permission DENIED — re-grant in System Settings > "
                      "Privacy & Security > Automation (a restart won't fix this)",
            "wedged": "Mail.app UNRESPONSIVE (restart failed)",
        }.get(status, "Mail.app status: " + status)
        if status != "healthy" and detail:
            msg += " (" + detail + ")"
        print(msg)
        return 0 if ok else (2 if status == "denied" else 1)

    if not a.to or not a.subject:
        ap.error("--to and --subject are required unless --health-check")
    body = open(a.body_file, encoding="utf-8").read() if a.body_file else sys.stdin.read()
    ok, detail = send_mail(a.to, a.subject, body, sender=a.sender,
                           attachments=a.attach, verify_timeout=a.verify_timeout,
                           draft_only=a.draft_only)
    print(("OK: " if ok else "FAIL: ") + detail)
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(_main())
