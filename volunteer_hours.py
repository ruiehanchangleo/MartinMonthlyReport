"""
Volunteer translation time in XTM, from the REST statistics API.

Source: the async aggregated-statistics report
  POST /projects/statistics?reportType=USER&startDate=..&endDate=..&linguistIds=..
    -> 202 {processId}
  GET  /projects/statistics?processId=..   (404 "Unavailable data" until ready)
    -> aggregatedByUser[] -> projectStatistics[] -> stepStatistics[]
       -> jobStatistics[] -> targetStatistics.totalTime   (milliseconds)

Summing targetStatistics.totalTime per user gives real tracked translation
time (what the UI shows as "Translation time [hh:mm:ss]"). This replaces the
old login/logout-session estimate and needs no browser, so it runs unattended.

All functions are best-effort: they return an 'unavailable' summary on any
failure so report generation is never blocked.
"""
from __future__ import annotations

import logging
import time
from collections import defaultdict
from concurrent.futures import ThreadPoolExecutor
from datetime import date
from typing import Dict, List, Optional

import requests

logger = logging.getLogger("xtm_report")

# Report generation time grows with batch size, so smaller batches finish
# faster server-side; we lean on parallelism (not big batches) for throughput.
# (Batches of 100 were observed to exceed the poll window and drop data.)
_BATCH = 50            # linguistIds per async report request
_POLL_TRIES = 40       # GET polls before giving up on one batch
_POLL_MAX = 150        # seconds total to wait for one report before giving up
_MAX_WORKERS = 10      # parallel (month, batch) report requests


def format_hms(seconds) -> str:
    """Format a duration in seconds as H:MM:SS (hours may exceed 24)."""
    seconds = int(round(seconds or 0))
    h, rem = divmod(seconds, 3600)
    m, s = divmod(rem, 60)
    return f"{h}:{m:02d}:{s:02d}"


def _chunks(lst, n):
    for i in range(0, len(lst), n):
        yield lst[i:i + n]


def _linguist_ids(base_url: str, headers: Dict, excluded_users: List[str]) -> List[int]:
    """All user IDs from /users, minus EXCLUDED_USERS (matched by username)."""
    excl = {u.lower() for u in (excluded_users or [])}
    ids: List[int] = []
    page = 1
    while True:
        batch = requests.get(f"{base_url}/users", headers=headers,
                             params={"page": page, "pageSize": 1000}, timeout=90).json()
        if not isinstance(batch, list) or not batch:
            break
        for u in batch:
            if u.get("id") and (u.get("username", "").lower() not in excl):
                ids.append(u["id"])
        if len(batch) < 1000:
            break
        page += 1
    return ids


def _run_report(base_url: str, headers: Dict, start_iso: str, end_iso: str,
                linguist_batch: List[int]) -> Dict[str, dict]:
    """Run one async USER statistics report for a batch of linguists and return
    the compact per-user sums {username: {active_seconds, jobs, languages:set}}.

    Summing inside the worker lets the large raw response be freed immediately
    (only the small summary crosses back to the caller). Empty dict on failure.
    """
    params = [("reportType", "USER"), ("startDate", start_iso), ("endDate", end_iso)]
    params += [("linguistIds", i) for i in linguist_batch]
    try:
        r = requests.post(f"{base_url}/projects/statistics", headers=headers,
                          params=params, timeout=120)
        if r.status_code != 202:
            logger.warning("stats POST failed %s: %s", r.status_code, r.text[:120])
            return {}
        process_id = r.json().get("processId")
        # Adaptive polling: check quickly at first (reports are often ready in a
        # few seconds), then back off. Cap total wait at _POLL_MAX.
        waited, delay = 0.0, 1.0
        for _ in range(_POLL_TRIES):
            g = requests.get(f"{base_url}/projects/statistics", headers=headers,
                            params={"processId": process_id}, timeout=120)
            if g.status_code == 200 and g.content and len(g.content) > 5:
                return _sum_user_seconds(g.json().get("aggregatedByUser", []) or [])
            if waited >= _POLL_MAX:
                break
            time.sleep(delay)
            waited += delay
            delay = min(delay * 1.4, 5.0)
        logger.warning("stats report timed out for batch of %d linguists", len(linguist_batch))
    except Exception as e:
        logger.warning("stats report error: %s", e)
    return {}


def _sum_user_seconds(aggregated: List[dict]) -> Dict[str, dict]:
    """Sum targetStatistics.totalTime (ms) per user from an aggregatedByUser
    list into {username: {'active_seconds', 'jobs', 'languages': set}}."""
    out: Dict[str, dict] = {}
    for u in aggregated:
        uname = u.get("userName")
        if not uname:
            continue
        e = out.setdefault(uname, {"active_seconds": 0.0, "jobs": 0, "languages": set()})
        for ps in u.get("projectStatistics", []):
            lang = ps.get("targetLanguage")
            if lang:
                e["languages"].add(lang)
            for st in ps.get("stepStatistics", []):
                for js in st.get("jobStatistics", []):
                    ms = (js.get("targetStatistics") or {}).get("totalTime", 0) or 0
                    e["active_seconds"] += ms / 1000.0
                    e["jobs"] += 1
    return out


def _finalize(by_user: Dict[str, dict]) -> Dict:
    for e in by_user.values():
        e["active_seconds"] = int(round(e["active_seconds"]))
        e["active_hours"] = round(e["active_seconds"] / 3600.0, 2)
        if isinstance(e.get("languages"), set):
            e["languages"] = sorted(e["languages"])
    total_seconds = sum(e["active_seconds"] for e in by_user.values())
    return {
        "by_user": by_user,
        "total_seconds": total_seconds,
        "total_hours": round(total_seconds / 3600.0, 2),
        "total_jobs": sum(e["jobs"] for e in by_user.values()),
        "volunteer_count": len([e for e in by_user.values() if e["active_seconds"] > 0]),
    }


def _empty_summary() -> Dict:
    return {"by_user": {}, "total_seconds": 0, "total_hours": 0.0, "total_jobs": 0,
            "volunteer_count": 0, "unavailable": True}


def get_translation_time(base_url: str, headers: Dict, period_start: date,
                         period_end: date, excluded_users: List[str],
                         linguist_ids: Optional[List[int]] = None) -> Dict:
    """Total tracked translation time per volunteer for [period_start, period_end]."""
    try:
        if linguist_ids is None:
            linguist_ids = _linguist_ids(base_url, headers, excluded_users)
        if not linguist_ids:
            return _empty_summary()
        start_iso, end_iso = period_start.isoformat(), period_end.isoformat()
        batches = list(_chunks(linguist_ids, _BATCH))
        merged: Dict[str, dict] = {}
        with ThreadPoolExecutor(max_workers=_MAX_WORKERS) as ex:
            for summed in ex.map(lambda b: _run_report(base_url, headers, start_iso, end_iso, b), batches):
                for uname, e in summed.items():
                    m = merged.setdefault(uname, {"active_seconds": 0, "jobs": 0, "languages": set()})
                    m["active_seconds"] += e["active_seconds"]
                    m["jobs"] += e["jobs"]
                    m["languages"] |= e["languages"]
        return _finalize(merged)
    except Exception as e:
        logger.warning("get_translation_time failed: %s", e)
        return _empty_summary()


def get_translation_time_monthly(base_url: str, headers: Dict, months: List[str],
                                 excluded_users: List[str],
                                 linguist_ids: Optional[List[int]] = None) -> Dict:
    """Per-month tracked translation time per volunteer, for a YTD trend chart.

    `months` is a list of 'YYYY-MM'. Returns
    {'months': [...], 'by_user': {username: {'months': {ym: seconds},
     'active_seconds', 'jobs', 'languages': [..]}}}.
    Also returns 'unavailable': True if nothing could be fetched.
    """
    try:
        if linguist_ids is None:
            linguist_ids = _linguist_ids(base_url, headers, excluded_users)
        if not linguist_ids:
            return {"months": list(months), "by_user": {}, "unavailable": True}
        batches = list(_chunks(linguist_ids, _BATCH))

        # Build one task per (month, batch); each returns (ym, aggregatedByUser).
        tasks = []
        for ym in months:
            y, mo = int(ym[:4]), int(ym[5:7])
            start = date(y, mo, 1)
            end = (date(y + (mo == 12), (mo % 12) + 1, 1))  # first of next month
            for b in batches:
                tasks.append((ym, start.isoformat(), end.isoformat(), b))

        by_user: Dict[str, dict] = {}
        got_any = False
        with ThreadPoolExecutor(max_workers=_MAX_WORKERS) as ex:
            results = ex.map(
                lambda t: (t[0], _run_report(base_url, headers, t[1], t[2], t[3])), tasks)
            for ym, summed in results:
                if summed:
                    got_any = True
                for uname, e in summed.items():
                    m = by_user.setdefault(uname, {"months": defaultdict(int),
                                                   "months_jobs": defaultdict(int), "jobs": 0,
                                                   "active_seconds": 0, "languages": set()})
                    m["months"][ym] += e["active_seconds"]
                    m["months_jobs"][ym] += e["jobs"]
                    m["jobs"] += e["jobs"]
                    m["active_seconds"] += e["active_seconds"]
                    m["languages"] |= e["languages"]
        # NOTE: end date is first-of-next-month; XTM's range is treated as
        # [start, end) by creation date, so months don't double-count.
        for e in by_user.values():
            e["months"] = {k: int(round(v)) for k, v in e["months"].items()}
            e["months_jobs"] = dict(e["months_jobs"])
            e["active_seconds"] = int(round(e["active_seconds"]))
            e["active_hours"] = round(e["active_seconds"] / 3600.0, 2)
            e["languages"] = sorted(e["languages"]) if isinstance(e["languages"], set) else e["languages"]
        result = {"months": list(months), "by_user": by_user}
        if not got_any:
            result["unavailable"] = True
        return result
    except Exception as e:
        logger.warning("get_translation_time_monthly failed: %s", e)
        return {"months": list(months), "by_user": {}, "unavailable": True}


def summary_from_breakdown(breakdown: Dict, month: Optional[str] = None) -> Dict:
    """Derive a period summary from a monthly breakdown. If `month` is given,
    summarize just that 'YYYY-MM'; otherwise summarize the whole range."""
    if not breakdown or breakdown.get("unavailable"):
        return _empty_summary()
    by_user: Dict[str, dict] = {}
    for uname, e in breakdown.get("by_user", {}).items():
        if month is None:
            secs = e.get("active_seconds", 0)
            jobs = e.get("jobs", 0)
        else:
            secs = e.get("months", {}).get(month, 0)
            jobs = e.get("months_jobs", {}).get(month, 0)
        if secs > 0:
            by_user[uname] = {"active_seconds": secs,
                              "active_hours": round(secs / 3600.0, 2),
                              "jobs": jobs,
                              "languages": e.get("languages", [])}
    total_seconds = sum(v["active_seconds"] for v in by_user.values())
    return {
        "by_user": by_user,
        "total_seconds": total_seconds,
        "total_hours": round(total_seconds / 3600.0, 2),
        "total_jobs": sum(v["jobs"] for v in by_user.values()),
        "volunteer_count": len(by_user),
    }


if __name__ == "__main__":
    import argparse, json
    logging.basicConfig(level=logging.INFO)
    ap = argparse.ArgumentParser()
    ap.add_argument("--start", required=True)
    ap.add_argument("--end", required=True)
    args = ap.parse_args()
    cfg = json.load(open("xtm_config.json"))
    base = cfg["base_url"]
    hdrs = {"Authorization": f"{cfg['auth_type']} {cfg['auth_token']}", "Content-Type": "application/json"}
    excl = ["leo.chang@familysearch.org", "LeoAdmin",
            "Robert.Sena@churchofjesuschrist.org", "MartinADMIN", "Tester BSP BSP"]
    from datetime import datetime
    ps = datetime.strptime(args.start, "%Y-%m-%d").date()
    pe = datetime.strptime(args.end, "%Y-%m-%d").date()
    s = get_translation_time(base, hdrs, ps, pe, excl)
    print("total:", format_hms(s["total_seconds"]), "| volunteers:", s["volunteer_count"], "| jobs:", s["total_jobs"])
    for u, e in sorted(s["by_user"].items(), key=lambda x: -x[1]["active_seconds"])[:12]:
        print(f"  {u[:38]:38} {format_hms(e['active_seconds']):>10}  jobs={e['jobs']}  {','.join(e['languages'][:3])}")
