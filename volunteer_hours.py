"""
Volunteer time-in-XTM, derived from the PM-GUI login/logout history.

XTM's REST API exposes no usable time data (the only time field, manualTime, is
empty because volunteers never log manual time). Real time-on-platform lives in
the GUI backend's getUserLoginHistory.serv, which is fetched by the Playwright
helper `fetch_login_history.js` (it auto-logs-in with the browser profile's
saved credentials and pages through all records).

This module runs that fetcher, pairs LOGIN/LOGOUT events per user, and computes
"active hours" = LOGIN -> LAST_ACTION_IN_SESSION (excludes idle time before an
explicit logout; nearly equal to the full session span in practice).

Request dates are DD-MM-YYYY; the response DATE field is MM-DD-YYYY (an XTM
quirk, verified empirically). See memory: xtm-volunteer-login-history.
"""
from __future__ import annotations

import json
import logging
import os
import subprocess
from collections import defaultdict
from datetime import datetime, date, timedelta
from pathlib import Path
from typing import Dict, List, Optional

logger = logging.getLogger("xtm_report")

_ROOT = Path(__file__).parent
_FETCHER = _ROOT / "fetch_login_history.js"
_CACHE_DIR = _ROOT / ".cache" / "login_history"
_NODE = os.environ.get("XTM_NODE", "node")

# Sessions longer than this are treated as a left-open tab, not real work, and
# dropped so a single stale session can't dominate a volunteer's total.
_MAX_SESSION_HOURS = 16.0


def _ddmmyyyy(d: date) -> str:
    return f"{d.day:02d}-{d.month:02d}-{d.year}"


def _parse_response_dt(s: Optional[str]) -> Optional[datetime]:
    """Response DATE / LAST_ACTION_IN_SESSION are MM-DD-YYYY HH:MM."""
    if not s:
        return None
    try:
        return datetime.strptime(s, "%m-%d-%Y %H:%M")
    except ValueError:
        return None


def fetch_login_history(period_start: date, period_end: date,
                        out_path: Path, headless: bool = False,
                        timeout: int = 300) -> bool:
    """Run the Playwright fetcher for [period_start, period_end] (padded by a
    day on each side so boundary sessions pair correctly). Returns True on a
    successful write."""
    if not _FETCHER.exists():
        logger.warning("Login-history fetcher not found: %s", _FETCHER)
        return False
    date_from = _ddmmyyyy(period_start - timedelta(days=1))
    date_to = _ddmmyyyy(period_end + timedelta(days=1))
    env = dict(os.environ)
    if headless:
        env["XTM_HEADLESS"] = "1"
    out_path.parent.mkdir(parents=True, exist_ok=True)
    try:
        proc = subprocess.run(
            [_NODE, str(_FETCHER), date_from, date_to, str(out_path)],
            capture_output=True, text=True, timeout=timeout, env=env,
        )
    except subprocess.TimeoutExpired:
        logger.warning("Login-history fetch timed out after %ss", timeout)
        return False
    if proc.returncode != 0:
        logger.warning("Login-history fetch failed (rc=%s): %s",
                       proc.returncode, (proc.stderr or "").strip()[-300:])
        return False
    return out_path.exists()


def aggregate_from_file(path: Path, period_start: date, period_end: date,
                        excluded_users: List[str]) -> Dict:
    """Pair sessions and total active hours per volunteer over the period.

    A session counts toward the period if its LOGIN falls within
    [period_start 00:00, period_end 23:59]. Excluded users are dropped
    (case-insensitive). Returns a structured summary dict.
    """
    by_user: Dict[str, dict] = {}
    for user, login_dt, active_seconds in _iter_sessions(path, excluded_users):
        if not (datetime.combine(period_start, datetime.min.time()) <= login_dt
                <= datetime.combine(period_end, datetime.max.time())):
            continue
        e = by_user.setdefault(user, {"sessions": 0, "active_seconds": 0})
        e["sessions"] += 1
        e["active_seconds"] += active_seconds
    for e in by_user.values():
        e["active_hours"] = round(e["active_seconds"] / 3600.0, 2)

    total_seconds = sum(v["active_seconds"] for v in by_user.values())
    total_sessions = sum(v["sessions"] for v in by_user.values())
    return {
        "by_user": by_user,
        "total_seconds": total_seconds,
        "total_hours": round(total_seconds / 3600.0, 2),
        "total_sessions": total_sessions,
        "volunteer_count": len(by_user),
        "period_start": period_start.isoformat(),
        "period_end": period_end.isoformat(),
    }


def _iter_sessions(path, excluded_users):
    """Yield (username, login_datetime, active_seconds) for each paired
    LOGIN/LOGOUT session in the file, dropping excluded users and sessions
    longer than _MAX_SESSION_HOURS. Active seconds = login -> last action."""
    excl = {u.lower() for u in (excluded_users or [])}
    try:
        raw = json.loads(Path(path).read_text())
    except Exception as e:
        logger.warning("Could not read login history %s: %s", path, e)
        return

    by_user_events: Dict[str, List[dict]] = defaultdict(list)
    for r in raw.get("data", []):
        u = (r.get("USERNAME") or "").strip()
        if not u or u.lower() in excl:
            continue
        by_user_events[u].append(r)

    for user, events in by_user_events.items():
        events.sort(key=lambda r: _parse_response_dt(r.get("DATE")) or datetime.min)
        open_login = None
        for r in events:
            action = r.get("ACTION")
            if action == "LOGIN":
                open_login = _parse_response_dt(r.get("DATE"))
            elif action == "LOGOUT" and open_login is not None:
                last = _parse_response_dt(r.get("LAST_ACTION_IN_SESSION")) or \
                    _parse_response_dt(r.get("DATE"))
                if last:
                    seconds = (last - open_login).total_seconds()
                    if 0 <= seconds <= _MAX_SESSION_HOURS * 3600:
                        yield user, open_login, int(round(seconds))
                open_login = None


def aggregate_monthly_breakdown(path, months: List[str],
                                excluded_users: List[str]) -> Dict:
    """Break active time down by month per volunteer, for a YTD line chart.

    `months` is a list of 'YYYY-MM' strings. Each session's active time is
    attributed to the month of its LOGIN. Returns
    {'months': [...], 'by_user': {login: {'months': {ym: seconds}, ...}}}.
    """
    month_set = set(months)
    by_user: Dict[str, dict] = {}
    for user, login_dt, active_seconds in _iter_sessions(path, excluded_users):
        ym = f"{login_dt.year:04d}-{login_dt.month:02d}"
        if ym not in month_set:
            continue
        e = by_user.setdefault(user, {"months": {}, "sessions": 0, "active_seconds": 0})
        e["months"][ym] = e["months"].get(ym, 0) + active_seconds
        e["sessions"] += 1
        e["active_seconds"] += active_seconds
    for e in by_user.values():
        e["active_hours"] = round(e["active_seconds"] / 3600.0, 2)
    return {"months": list(months), "by_user": by_user}


def format_hms(seconds) -> str:
    """Format a duration in seconds as H:MM:SS (hours can exceed 24)."""
    seconds = int(round(seconds or 0))
    h, rem = divmod(seconds, 3600)
    m, s = divmod(rem, 60)
    return f"{h}:{m:02d}:{s:02d}"


def _empty_summary() -> Dict:
    return {"by_user": {}, "total_seconds": 0, "total_hours": 0.0,
            "total_sessions": 0, "volunteer_count": 0, "period_start": None,
            "period_end": None, "source_file": None, "unavailable": True}


def get_volunteer_hours(period_start: date, period_end: date,
                        excluded_users: List[str], refresh: bool = True,
                        headless: bool = False) -> Dict:
    """Fetch (unless refresh=False and a cache exists) and aggregate volunteer
    hours for the period. Never raises — returns an empty/unavailable summary
    on any failure so report generation is unaffected. The returned summary
    includes 'source_file' so callers can re-aggregate sub-windows (e.g. the
    current month) from the same fetched data without a second fetch."""
    out_path = _CACHE_DIR / f"login_history_{period_start:%Y%m%d}_{period_end:%Y%m%d}.json"
    try:
        if refresh or not out_path.exists():
            ok = fetch_login_history(period_start, period_end, out_path, headless=headless)
            if not ok and not out_path.exists():
                logger.warning("Volunteer hours unavailable (fetch failed, no cache)")
                return _empty_summary()
        summary = aggregate_from_file(out_path, period_start, period_end, excluded_users)
        summary["source_file"] = str(out_path)
        return summary
    except Exception as e:
        logger.warning("Volunteer hours aggregation failed: %s", e)
        return _empty_summary()


if __name__ == "__main__":
    # Quick manual test against an existing fetched file or a fresh fetch.
    import argparse
    logging.basicConfig(level=logging.INFO)
    ap = argparse.ArgumentParser()
    ap.add_argument("--start", required=True, help="period start YYYY-MM-DD")
    ap.add_argument("--end", required=True, help="period end YYYY-MM-DD")
    ap.add_argument("--file", help="use an existing fetched JSON instead of fetching")
    ap.add_argument("--no-refresh", action="store_true")
    args = ap.parse_args()
    ps = datetime.strptime(args.start, "%Y-%m-%d").date()
    pe = datetime.strptime(args.end, "%Y-%m-%d").date()
    excl = ["leo.chang@familysearch.org", "LeoAdmin",
            "Robert.Sena@churchofjesuschrist.org", "MartinADMIN", "Tester BSP BSP"]
    if args.file:
        summary = aggregate_from_file(Path(args.file), ps, pe, excl)
    else:
        summary = get_volunteer_hours(ps, pe, excl, refresh=not args.no_refresh)
    print(json.dumps({k: v for k, v in summary.items() if k != "by_user"}, indent=2))
    print("total active time:", format_hms(summary.get("total_seconds", 0)))
    top = sorted(summary["by_user"].items(), key=lambda x: -x[1]["active_seconds"])[:15]
    for u, v in top:
        print(f"  {u[:40]:40} {format_hms(v['active_seconds']):>10}  ({v['sessions']} sessions)")
