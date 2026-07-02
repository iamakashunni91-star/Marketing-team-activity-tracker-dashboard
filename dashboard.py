"""
Way.com Marketing Dashboard Updater — GitHub Actions Version
==============================================================
This version runs on GitHub's servers (not your laptop).
It downloads the Excel from a PUBLIC SharePoint "view" link
(no login needed), rebuilds all task data, and updates the
embedded JSON inside dashboard.html.

This file lives in the root of your GitHub repository.
GitHub Actions runs it automatically every hour.
"""

import os
import re
import json
import time
import logging
import tempfile
from datetime import datetime
from pathlib import Path

import pandas as pd
import requests

# ════════════════════════════════════════════════════════════════
# CONFIG
# ════════════════════════════════════════════════════════════════

# The dashboard file in this repo (relative path — do not change)
DASHBOARD_HTML = "index.html"

# The PUBLIC SharePoint "anyone with the link can view" URL.
# This is stored as a GitHub Secret (SHAREPOINT_URL) for security —
# it is NOT hardcoded here. GitHub injects it as an environment variable.
SHAREPOINT_URL = os.environ.get("SHAREPOINT_URL", "")

# Which sheet inside the Excel to read
SHEET_NAME = "Master activity matrix - June"

# Date window: include everything from the program start date onward.
# No upper cap — as the team adds July, August, etc. rows to the same
# sheet, they flow into the dashboard automatically. Leadership filters
# to any custom range (weekly / MTD) using the date picker in the page.
DATE_START = "2026-06-01"
DATE_END   = None   # None = no upper limit (include all future dates)

# ════════════════════════════════════════════════════════════════
# LOGGING
# ════════════════════════════════════════════════════════════════

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)s  %(message)s",
)
log = logging.getLogger(__name__)


# ════════════════════════════════════════════════════════════════
# STEP 1 — DOWNLOAD EXCEL FROM PUBLIC SHAREPOINT LINK
# ════════════════════════════════════════════════════════════════

def convert_to_download_url(share_url: str) -> str:
    """
    SharePoint 'anyone with the link can view' URLs need '?download=1'
    appended to force a file download instead of opening the web viewer.
    """
    if "download=1" in share_url:
        return share_url
    sep = "&" if "?" in share_url else "?"
    return f"{share_url}{sep}download=1"


def download_excel() -> Path:
    if not SHAREPOINT_URL:
        raise RuntimeError(
            "SHAREPOINT_URL secret is not set. "
            "Add it in GitHub repo Settings → Secrets and variables → Actions."
        )

    download_url = convert_to_download_url(SHAREPOINT_URL)
    log.info("Downloading Excel from SharePoint public link...")

    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                      "AppleWebKit/537.36 (KHTML, like Gecko) "
                      "Chrome/120.0.0.0 Safari/537.36",
    }

    resp = requests.get(download_url, headers=headers, timeout=60, allow_redirects=True)

    if resp.status_code != 200:
        raise RuntimeError(f"Download failed — HTTP {resp.status_code}")

    content_type = resp.headers.get("Content-Type", "")
    if "html" in content_type.lower() and len(resp.content) < 50_000:
        raise RuntimeError(
            "Got an HTML page instead of the Excel file. "
            "Check that the SharePoint link is set to "
            "'Anyone with the link can view' (not 'People in organization')."
        )

    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
    tmp.write(resp.content)
    tmp.close()

    log.info(f"Downloaded {len(resp.content)/1024:.1f} KB")
    return Path(tmp.name)


# ════════════════════════════════════════════════════════════════
# STEP 2 — COMPUTE TASK DATA (same logic as the local version)
# ════════════════════════════════════════════════════════════════


def _clean(s):
    if not isinstance(s, str):
        return s
    return ''.join(c if c >= ' ' else ' ' for c in s).strip()

def fmt_date(d):
    try:
        return pd.to_datetime(d).strftime("%d %b")
    except Exception:
        return str(d)


def iso_date(d):
    try:
        return pd.to_datetime(d).strftime("%Y-%m-%d")
    except Exception:
        return None


def extract_count(deliverable, activity=""):
    raw = str(deliverable).strip()
    if not raw or raw in ["nan", "—", ""]:
        return 1
    if raw.startswith('"') or raw.startswith('\u201c'):
        return 1
    d = raw.strip('"').strip('\u201c').strip('\u201d').strip()
    dl = d.lower()

    patterns = [
        r'(\d+)\s*screens?', r'(\d+)\s*emails?', r'(\d+)\s*templates?',
        r'(\d+)\s*campaign', r'(\d+)\s*posts?', r'(\d+)\s*stor(?:y|ies)',
        r'(\d+)\s*reels?', r'(\d+)\s*videos?', r'(\d+)\s*pages?',
        r'(\d+)\s*keywords?', r'(\d+)\s*markets?', r'(\d+)\s*reports?',
        r'(\d+)\s*leads?', r'(\d+)\s*tickets?', r'(\d+)\s*variants?',
        r'(\d+)\s*qr codes?', r'(\d+)\s*banners?', r'(\d+)\s*proposals?',
        r'(\d+)\s*wireframes?', r'(\d+)\s*designs?', r'\(total (\d+)\)',
        r'(\d+)\s*comments?',
        r'(\d+)\s*(?:city parking|airport parking|all vertical|all parking|event parking)\s*emails?',
        r'correspondence emails[-–]\s*(\d+)',
        r'(\d+)\s*[Bb]log',
    ]
    for p in patterns:
        m = re.search(p, d, re.IGNORECASE)
        if m:
            return int(m.group(1))

    written = {r'\btwo\b': 2, r'\bthree\b': 3, r'\bfour\b': 4, r'\bfive\b': 5,
               r'\bsix\b': 6, r'\bseven\b': 7, r'\beight\b': 8,
               r'\bnine\b': 9, r'\bten\b': 10}
    for p, val in written.items():
        if re.search(p, dl):
            return val

    return 1


def build_task_data(df: pd.DataFrame) -> dict:
    """
    Builds the full task_data dict for all marketing sub-teams.
    Mirrors the logic used in the local dashboard builder —
    one bucket assignment per row, no gaps, no double-counting.
    """
    all_data = {}

    def add_entries(key, subset):
        rows = []
        for _, r in subset.iterrows():
            cnt = extract_count(
                str(r.get('Work Output / Deliverable', '')),
                str(r.get('Activity', ''))
            )
            mins = int(r['Time Spent (in Mins)']) if pd.notna(r['Time Spent (in Mins)']) else 0
            rows.append({
                "date": fmt_date(r['Date']),
                "iso": iso_date(r['Date']),
                "activity": _clean(str(r.get('Activity', '')))[:200],
                "deliverable": _clean(str(r.get('Work Output / Deliverable', '')))[:300],
                "mins": mins,
                "freq": str(r.get('Frequency', '—')).strip() if pd.notna(r.get('Frequency')) else "—",
                "impact": str(r.get('Impact Bucket', '—')).strip() if pd.notna(r.get('Impact Bucket')) else "—",
                "del_count": cnt,
                "mins_per_del": round(mins / cnt) if cnt and cnt > 0 and mins > 0 else None,
            })
        all_data[key] = rows

    def A(df_, person, pat):
        sub = df_[df_['Employee name'] == person]
        return sub[sub['Activity'].str.contains(pat, na=False, regex=True, case=False)]

    # ── CONTENT TEAM ─────────────────────────────────────────────
    for person, pkey in [("Archa Ullas", "Archa"), ("Anna Mary", "Anna")]:
        sub = df[df['Employee name'] == person]
        add_entries(f"{pkey}_omni_email", sub[sub['Activity'].str.contains(r'[Oo]mni[Cc]hannel email|Email [Cc]ontent.*[Oo]mni', na=False, regex=True)])
        add_entries(f"{pkey}_micro_email", sub[sub['Activity'].str.contains(r'[Mm]icrosite [Ee]mail', na=False, regex=True)])
        add_entries(f"{pkey}_ad_copy", sub[sub['Activity'].str.contains(r'[Aa]d copy|[Gg]oogle search ad|[Mm]eta ad', na=False, regex=True)])
        add_entries(f"{pkey}_video", sub[sub['Activity'].str.contains(r'[Vv]ideo script', na=False, regex=True)])
        add_entries(f"{pkey}_push", sub[sub['Activity'].str.contains(r'[Pp]ush notification', na=False, regex=True)])
        add_entries(f"{pkey}_process", sub[sub['Activity'].str.contains(r'[Ww]orkflow|[Ii]ntake [Ff]orm|[Jj]ira', na=False, regex=True)])
        add_entries(f"{pkey}_interview", sub[sub['Activity'].str.contains(r'[Ii]nterv|[Ii]nvigilat', na=False, regex=True)])
        add_entries(f"{pkey}_meetings", sub[sub['Activity'].str.contains(r'[Mm]eeting', na=False, regex=True)])

    for person, pkey in [("Gautham S", "Gautham"), ("Gokul Nath", "Gokul"), ("Sreejith SL", "Sreejith")]:
        sub = df[df['Employee name'] == person]
        add_entries(f"{pkey}_way_pages", sub[sub['Activity'].str.contains(r'[Ww]ay static page|[Ww]ay airport', na=False, regex=True)])
        add_entries(f"{pkey}_gap_pages", sub[sub['Activity'].str.contains(r'[Gg][Aa][Pp].*[Pp]age', na=False, regex=True)])
        add_entries(f"{pkey}_cruise_pages", sub[sub['Activity'].str.contains(r'[Cc]ruise.*[Pp]arking|[Cc]ruise.*[Pp]age', na=False, regex=True)])
        add_entries(f"{pkey}_city_pages", sub[sub['Activity'].str.contains(r'[Cc]ity.*[Pp]arking.*[Pp]age|NRG Stadium', na=False, regex=True)])
        add_entries(f"{pkey}_reddit", sub[sub['Activity'].str.contains(r'[Rr]eddit', na=False, regex=True)])
        add_entries(f"{pkey}_qa", sub[sub['Activity'].str.contains(r'[Rr]eview|QA', na=False, regex=True)])
        add_entries(f"{pkey}_meetings", sub[sub['Activity'].str.contains(r'[Mm]eeting|[Aa]ll [Hh]ands', na=False, regex=True)])

    for person, pkey in [("Shilpa Sara", "Shilpa"), ("Arun Mahadev", "Arun")]:
        sub = df[df['Employee name'] == person]
        add_entries(f"{pkey}_carwash", sub[sub['Activity'].str.contains(r'[Cc]ar [Ww]ash', na=False, regex=True) | sub['Work Output / Deliverable'].str.contains(r'[Cc]ar [Ww]ash', na=False, regex=True)])
        add_entries(f"{pkey}_wayplus", sub[sub['Activity'].str.contains(r'[Ww]ay\+|[Rr]etention', na=False, regex=True) | sub['Work Output / Deliverable'].str.contains(r'[Ww]ay\+|[Rr]etention', na=False, regex=True)])
        add_entries(f"{pkey}_mileage", sub[sub['Activity'].str.contains(r'[Mm]ileage [Tt]racker|[Tt]rackmile', na=False, regex=True) | sub['Work Output / Deliverable'].str.contains(r'[Mm]ileage [Tt]racker|[Tt]rackmile', na=False, regex=True)])
        add_entries(f"{pkey}_rm", sub[sub['Activity'].str.contains(r'R&M|[Rr]epair.*[Mm]anagement', na=False, regex=True) | sub['Work Output / Deliverable'].str.contains(r'R&M|[Ww]orkboard|[Rr]oadside', na=False, regex=True)])
        add_entries(f"{pkey}_hyundai", sub[sub['Activity'].str.contains(r'[Hh]yundai|[Pp]leos', na=False, regex=True)])
        add_entries(f"{pkey}_event", sub[sub['Activity'].str.contains(r'[Ee]vent [Pp]arking|[Ss]elect [Pp]arking|[Cc]ity [Pp]arking|[Ss]MS|[Cc]ancellation', na=False, regex=True)])
        add_entries(f"{pkey}_qa_doc", sub[sub['Activity'].str.contains(r'QA|[Dd]ocument|[Ww]alkthrough|[Ii]nterview', na=False, regex=True)])

    for person, pkey in [("Rajeswari Menon", "Raje"), ("Sneha S", "Sneha")]:
        sub = df[df['Employee name'] == person]
        add_entries(f"{pkey}_blogs", sub[sub['Activity'].str.contains(r'[Bb]log', na=False, regex=True)])
        add_entries(f"{pkey}_email", sub[sub['Activity'].str.contains(r'[Ee]mail', na=False, regex=True)])
        add_entries(f"{pkey}_video", sub[sub['Activity'].str.contains(r'[Vv]ideo [Ss]cript', na=False, regex=True)])
        add_entries(f"{pkey}_pages", sub[sub['Activity'].str.contains(r'[Ll]anding [Pp]age|[Pp]age.*revamp|[Pp]arking.*page|[Ii]ndustry [Pp]age|[Bb]usiness.*page', na=False, regex=True)])
        add_entries(f"{pkey}_strategic", sub[sub['Activity'].str.contains(r'[Ww]hite [Pp]aper|[Cc]ase [Ss]tudy', na=False, regex=True)])
        add_entries(f"{pkey}_adcopy", sub[sub['Activity'].str.contains(r'[Aa]d [Cc]opy|[Bb]anner [Cc]opy|[Ff]lyer', na=False, regex=True)])
        add_entries(f"{pkey}_hubspot", sub[sub['Activity'].str.contains(r'[Hh]ub[Ss]pot|GA4.*monitor', na=False, regex=True)])
        add_entries(f"{pkey}_recruit", sub[sub['Activity'].str.contains(r'[Ii]nterview|[Hh]iring|[Cc]andidates', na=False, regex=True)])
        add_entries(f"{pkey}_meetings", sub[sub['Activity'].str.contains(r'[Mm]eeting|[Ss]ync|1:1', na=False, regex=True)])

    sub = df[df['Employee name'] == 'Fanny Dorris']
    add_entries("Fanny_ap_email", sub[sub['Activity'].str.contains(r'[Aa]irport.*[Ee]mail', na=False, regex=True)])
    add_entries("Fanny_omni_email", sub[sub['Activity'].str.contains(r'[Oo]mni[Cc]hannel [Ee]mail', na=False, regex=True)])
    add_entries("Fanny_way_pages", sub[sub['Activity'].str.contains(r'[Ww]ay.*[Pp]age|[Ww]ay.*[Ss]tatic', na=False, regex=True)])
    add_entries("Fanny_gap_pages", sub[sub['Activity'].str.contains(r'[Gg][Aa][Pp].*[Ss]tatic|[Gg][Aa][Pp].*[Pp]age', na=False, regex=True)])
    add_entries("Fanny_blogs", sub[sub['Activity'].str.contains(r'[Bb]2[Bb].*[Bb]log|[Bb]log.*revamp', na=False, regex=True)])
    add_entries("Fanny_cruise", sub[sub['Activity'].str.contains(r'[Cc]ruise', na=False, regex=True)])
    add_entries("Fanny_video", sub[sub['Activity'].str.contains(r'[Vv]ideo [Ss]cript', na=False, regex=True)])
    add_entries("Fanny_meetings", sub[sub['Activity'].str.contains(r'[Mm]eeting', na=False, regex=True)])

    # ── SEO / EMAIL / SOCIAL / B2B DIGITAL — single-assignment classifier ──
    def classify(person, activity, deliverable):
        t = (str(activity) + " " + str(deliverable)).lower()

        if person == 'Bajan':
            if re.search(r'keyword|ranking', t): return 'keyword'
            if re.search(r'ticket|prioritization', t): return 'ticket'
            if re.search(r'dashboard|deck|report', t): return 'dashboard'
            if re.search(r'standup|meeting|connect', t): return 'meetings'
            return 'audit'
        if person == 'Haripriya L':
            if re.search(r'content review|content', t): return 'content'
            if re.search(r'keyword|ranking|baseline', t): return 'keyword'
            if re.search(r'landing page|fifa', t): return 'landing'
            if re.search(r'deck|dashboard|performance', t): return 'dashboard'
            if re.search(r'standup|meeting|review|walkthrough|sync', t): return 'meetings'
            return 'audit'
        if person == 'Naveen PC':
            if re.search(r'crawl|screaming frog', t): return 'crawl'
            if re.search(r'backlink|disavow', t): return 'backlink'
            if re.search(r'ticket|follow-up', t): return 'ticket'
            if re.search(r'standup|meeting|stakeholder', t): return 'meetings'
            return 'audit'
        if person == 'Arun Nath J':
            if re.search(r'keyword|ranking', t): return 'keyword'
            if re.search(r'audit|website|url', t): return 'audit'
            if re.search(r'research|plan', t): return 'research'
            if re.search(r'jira|ticket|meeting', t): return 'meetings'
            return 'analysis'
        if person == 'Balavignesh P':
            if re.search(r'email id check|lead feeder', t): return 'emailops'
            if re.search(r'template|newsletter', t): return 'template'
            if re.search(r'task scheduling|naitik', t): return 'delegation'
            if re.search(r'ga tracking|hubspot|jira', t): return 'tracking'
            if re.search(r'time sheet|meeting|sync', t): return 'meetings'
            return 'tracking'
        if person == 'Kulwinder Singh':
            if re.search(r'template', t): return 'template'
            if re.search(r'qa|klaviyo event', t): return 'qa'
            if re.search(r'klaviyo|mailchimp|netcore|sendgrid', t): return 'platform'
            if re.search(r'review call|meeting|report', t): return 'meetings'
            return 'platform'
        if person == 'Ajay Singh':
            if re.search(r'campaign|klaviyo', t): return 'campaign'
            if re.search(r'sendgrid|tableau|report|analysis', t): return 'reports'
            if re.search(r'qa|approval', t): return 'qa'
            if re.search(r'meet|weekly', t): return 'meetings'
            return 'schedule'
        if person == 'Savitha Vasanthan':
            return 'template'
        if person == 'Priya Kumari':
            if re.search(r'social media|cadence', t): return 'social'
            if re.search(r'proposal|pitch|deck', t): return 'proposal'
            if re.search(r'landing page|campaign', t): return 'landing'
            if re.search(r'meeting|review', t): return 'meetings'
            return 'audit'
        if person == 'Devika Sheeja':
            if re.search(r'story', t): return 'story'
            if re.search(r'video|reel', t): return 'video'
            if re.search(r'template|tv updation|bday', t): return 'template'
            if re.search(r'calendar|schedule|campaign', t): return 'calendar'
            if re.search(r'meeting', t): return 'meetings'
            return 'posts'
        if person == 'Jofia Joseph':
            if re.search(r'video|reel', t): return 'video'
            if re.search(r'metric|comment|report', t): return 'metrics'
            if re.search(r'schedule|calendar', t): return 'schedule'
            if re.search(r'meeting|research', t): return 'meetings'
            return 'posts'
        if person == 'Seethal vargheese':
            if re.search(r'lead|partner|hubspot', t): return 'leadmgmt'
            if re.search(r'linkedin|facebook|content calendar', t): return 'social'
            if re.search(r'design|wireframe|banner|qr code', t): return 'design'
            if re.search(r'utm|campaign|newsletter|video', t): return 'campaign'
            if re.search(r'meeting|follow-up|alignment', t): return 'meetings'
            return 'campaign'
        return 'other'

    pkey_map = {'Bajan': 'Bajan', 'Haripriya L': 'Haripriya', 'Naveen PC': 'Naveen',
                'Arun Nath J': 'ArunNath', 'Balavignesh P': 'Bala', 'Kulwinder Singh': 'Kulwinder',
                'Ajay Singh': 'Ajay', 'Savitha Vasanthan': 'Savitha', 'Priya Kumari': 'Priya',
                'Devika Sheeja': 'Devika', 'Jofia Joseph': 'Jofia', 'Seethal vargheese': 'Seethal'}

    from collections import defaultdict
    buckets = defaultdict(list)
    for person, pk in pkey_map.items():
        sub = df[df['Employee name'] == person]
        for _, r in sub.iterrows():
            bucket = classify(person, r.get('Activity', ''), r.get('Work Output / Deliverable', ''))
            key = f"{pk}_{bucket}"
            cnt = extract_count(str(r.get('Work Output / Deliverable', '')), str(r.get('Activity', '')))
            mins = int(r['Time Spent (in Mins)']) if pd.notna(r['Time Spent (in Mins)']) else 0
            buckets[key].append({
                "date": fmt_date(r['Date']),
                "iso": iso_date(r['Date']),
                "activity": _clean(str(r.get('Activity', '')))[:200],
                "deliverable": _clean(str(r.get('Work Output / Deliverable', '')))[:300],
                "mins": mins,
                "freq": str(r.get('Frequency', '—')).strip() if pd.notna(r.get('Frequency')) else "—",
                "impact": str(r.get('Impact Bucket', '—')).strip() if pd.notna(r.get('Impact Bucket')) else "—",
                "del_count": cnt,
                "mins_per_del": round(mins / cnt) if cnt and cnt > 0 and mins > 0 else None,
            })

    all_data.update(dict(buckets))
    return all_data


# ════════════════════════════════════════════════════════════════
# STEP 3 — UPDATE HTML
# ════════════════════════════════════════════════════════════════

def update_html(html_path: str, new_data: dict) -> bool:
    """Returns True if the file content actually changed."""
    path = Path(html_path)
    if not path.exists():
        raise FileNotFoundError(f"{html_path} not found in repo root.")

    with open(path, encoding="utf-8") as f:
        html = f.read()

    new_json = json.dumps(new_data, ensure_ascii=False)
    new_const = f"const T = {new_json};"

    pattern = r'const T = \{.*?\};'
    if not re.search(pattern, html, flags=re.DOTALL):
        raise ValueError("Could not find 'const T = {...};' in dashboard.html")

    html_updated = re.sub(pattern, new_const, html, flags=re.DOTALL)

    ts = datetime.utcnow().strftime("%d %b %Y, %I:%M %p UTC")
    if "Last updated:" in html_updated:
        html_updated = re.sub(r'Last updated:.*?(?=<)', f'Last updated: {ts} ', html_updated)
    else:
        # Insert a timestamp near the rh-date badge if not present
        html_updated = html_updated.replace(
            '<div class="rh-date">',
            f'<div class="rh-date">Last updated: {ts} · ',
            1
        )

    changed = html_updated != html
    if changed:
        with open(path, "w", encoding="utf-8") as f:
            f.write(html_updated)
    return changed


# ════════════════════════════════════════════════════════════════
# MAIN
# ════════════════════════════════════════════════════════════════

# ════════════════════════════════════════════════════════════════
# MISSING-UPDATES CHECK
# ════════════════════════════════════════════════════════════════
# The full expected roster — everyone who should be logging activity.
# When a new joiner is added to the dashboard, add their exact
# SharePoint name here too so the compliance check includes them.
ROSTER = [
    "Archa Ullas", "Anna Mary", "Gautham S", "Gokul Nath", "Sreejith SL",
    "Shilpa Sara", "Arun Mahadev", "Rajeswari Menon", "Sneha S", "Fanny Dorris",
    "Bajan", "Haripriya L", "Naveen PC", "Arun Nath J", "Balavignesh P",
    "Kulwinder Singh", "Ajay Singh", "Savitha Vasanthan", "Priya Kumari",
    "Devika Sheeja", "Jofia Joseph", "Seethal vargheese",
]


def check_missing_updates(df_all: pd.DataFrame) -> None:
    """Check who did NOT log activity for the previous WORKING day.

    Runs once daily (the workflow schedules the evening-IST run). Skips
    weekends: if the previous calendar day was a weekend, no check is done.
    Writes missing_yesterday.json for downstream email delivery.
    """
    from datetime import timedelta, timezone

    # "Now" in IST (UTC+5:30). We check the day that just ended.
    ist = timezone(timedelta(hours=5, minutes=30))
    now_ist = datetime.now(ist)
    target = (now_ist - timedelta(days=1)).date()   # yesterday, IST

    # Skip if the day we'd be checking is a weekend (Sat=5, Sun=6)
    if target.weekday() >= 5:
        log.info(f"Missing-check skipped — {target} was a weekend.")
        return

    target_iso = target.strftime("%Y-%m-%d")

    # Who logged at least one row dated the target day?
    logged = set(
        df_all.loc[df_all['Date'].dt.strftime("%Y-%m-%d") == target_iso, 'Employee name']
    )
    missing = [name for name in ROSTER if name not in logged]

    payload = {
        "checked_on": now_ist.strftime("%Y-%m-%d %H:%M IST"),
        "for_date": target.strftime("%a, %d %b %Y"),
        "for_date_iso": target_iso,
        "total_roster": len(ROSTER),
        "missing_count": len(missing),
        "missing_names": missing,
    }
    with open("missing_yesterday.json", "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)

    if missing:
        log.info(f"MISSING activity for {payload['for_date']} "
                 f"({len(missing)}/{len(ROSTER)}): {', '.join(missing)}")
    else:
        log.info(f"All {len(ROSTER)} logged activity for {payload['for_date']} ✓")


def main():
    start = time.time()
    log.info("=" * 60)
    log.info(f"Dashboard update started — {datetime.utcnow().isoformat()}")

    tmp_excel = None
    try:
        tmp_excel = download_excel()

        log.info(f"Reading sheet: {SHEET_NAME}")
        df = pd.read_excel(tmp_excel, sheet_name=SHEET_NAME, header=1)
        df['Employee name'] = df['Employee name'].astype(str).str.strip()
        df['Date'] = pd.to_datetime(df['Date'], errors='coerce')

        # Missing-updates compliance check. This is cheap, so we always run
        # it and write missing_yesterday.json. Whether an EMAIL is sent is
        # controlled separately by the workflow (only the 5:15 PM IST run and
        # manual runs send mail). Weekends are skipped inside the function.
        check_missing_updates(df.copy())

        before = len(df)
        if DATE_END:
            df = df[(df['Date'] >= DATE_START) & (df['Date'] <= DATE_END)]
        else:
            df = df[df['Date'] >= DATE_START]
        log.info(f"Date filter from {DATE_START} (no upper cap): {before} → {len(df)} rows")

        df = df.drop_duplicates(
            subset=['Employee name', 'Date', 'Activity',
                    'Work Output / Deliverable', 'Time Spent (in Mins)']
        )
        log.info(f"After dedup: {len(df)} rows, {df['Employee name'].nunique()} people")

        task_data = build_task_data(df)
        total_tasks = sum(len(v) for v in task_data.values())
        log.info(f"Built {len(task_data)} KPI keys, {total_tasks} task entries")

        changed = update_html(DASHBOARD_HTML, task_data)
        if changed:
            log.info("dashboard.html updated — changes will be committed")
        else:
            log.info("No data changes detected — nothing to commit")

        log.info(f"Done in {round(time.time()-start,1)}s ✓")

    except Exception as e:
        log.error(f"UPDATE FAILED: {e}")
        raise
    finally:
        if tmp_excel and tmp_excel.exists():
            tmp_excel.unlink()


if __name__ == "__main__":
    main()
