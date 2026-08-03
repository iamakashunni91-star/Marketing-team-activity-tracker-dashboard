"""
Marketing Team Activity Dashboard — Data Builder
=================================================
Runs hourly via GitHub Actions. Downloads the latest tracker from
SharePoint, classifies every activity into per-person KPI buckets,
and updates index.html so the dashboard reflects live data.

Structure (from director's new format):
  1. Organic SEO       — SEO + SEO Content
  2. Omnichannel       — Omnichannel email/content + Social
  3. Local Marketing   — Priya
  4. B2B               — B2B email + content + digital
  5. Cross Functional  — Fanny, Akash, Kiran

Every person has an "Other" KPI catch-all so every logged minute
is preserved. Dashboard per-person totals always match Excel totals.
"""

import json
import logging
import os
import re
import sys
import tempfile
import time
from datetime import datetime, timedelta, timezone
from pathlib import Path

import pandas as pd
import requests

logging.basicConfig(level=logging.INFO, format="%(asctime)s  %(levelname)s  %(message)s")
log = logging.getLogger(__name__)

# ═══════════════════════════════════════════════════════════════════════
# CONFIG
# ═══════════════════════════════════════════════════════════════════════
DASHBOARD_HTML = "dashboard_plain.html"
SHEET_NAME = "Master activity matrix - June"
DATE_START = "2026-06-01"
DATE_END = None   # None = no upper limit (include future dates)

# Stored as a GitHub Secret named SHAREPOINT_URL
SHAREPOINT_URL = os.environ.get("SHAREPOINT_URL", "")


# ═══════════════════════════════════════════════════════════════════════
# ROSTER — team & tab mapping
# ═══════════════════════════════════════════════════════════════════════
# Each entry: display name -> {tab, prefix (for HTML data-keys)}
# "Bajan BJ" rows are normalized to "Bajan" before classification.
ROSTER = {
    # ═══ Organic SEO tab ═══
    # SEO Technical
    "Bajan":             {"tab": "seo",   "subteam": "SEO Technical",     "prefix": "Bajan"},
    "Haripriya L":       {"tab": "seo",   "subteam": "SEO Technical",     "prefix": "Haripriya"},
    "Naveen PC":         {"tab": "seo",   "subteam": "SEO Technical",     "prefix": "Naveen"},
    "Arun Nath J":       {"tab": "seo",   "subteam": "SEO Technical",     "prefix": "ArunNath"},
    # SEO Content
    "Gautham S":         {"tab": "seo",   "subteam": "SEO Content",       "prefix": "Gautham"},
    "Gokul Nath":        {"tab": "seo",   "subteam": "SEO Content",       "prefix": "Gokul"},
    "Sreejith SL":       {"tab": "seo",   "subteam": "SEO Content",       "prefix": "Sreejith"},
    "Megha B":           {"tab": "seo",   "subteam": "SEO Content",       "prefix": "Megha"},
    # UX Copy
    "Shilpa Sara":       {"tab": "seo",   "subteam": "UX Copy",           "prefix": "Shilpa"},
    "Arun Mahadev":      {"tab": "seo",   "subteam": "UX Copy",           "prefix": "ArunM"},

    # ═══ Omnichannel Marketing tab ═══
    # Omnichannel Email
    "Ajay Singh":        {"tab": "omni",  "subteam": "Omnichannel Email", "prefix": "Ajay"},
    "Kulwinder Singh":   {"tab": "omni",  "subteam": "Omnichannel Email", "prefix": "Kulwinder"},
    "Savitha Vasanthan": {"tab": "omni",  "subteam": "Omnichannel Email", "prefix": "Savitha"},
    "Shivakumar patil":  {"tab": "omni",  "subteam": "Omnichannel Email", "prefix": "Shivakumar"},
    # Omnichannel Content
    "Anna Mary":         {"tab": "omni",  "subteam": "Omnichannel Content","prefix": "Anna"},
    "Archa Ullas":       {"tab": "omni",  "subteam": "Omnichannel Content","prefix": "Archa"},
    # Omnichannel Social
    "Devika Sheeja":     {"tab": "omni",  "subteam": "Omnichannel Social","prefix": "Devika"},
    "Jofia Joseph":      {"tab": "omni",  "subteam": "Omnichannel Social","prefix": "Jofia"},

    # ═══ Local Marketing tab ═══
    "Priya Kumari":      {"tab": "local", "subteam": "Local Marketing",   "prefix": "Priya"},

    # ═══ B2B tab ═══
    # B2B Email
    "Balavignesh P":     {"tab": "b2b",   "subteam": "B2B Email",         "prefix": "Bala"},
    "Adersh Raj":        {"tab": "b2b",   "subteam": "B2B Email",         "prefix": "Adersh"},
    # B2B Content
    "Rajeswari Menon":   {"tab": "b2b",   "subteam": "B2B Content",       "prefix": "Raje"},
    "Sneha S":           {"tab": "b2b",   "subteam": "B2B Content",       "prefix": "Sneha"},
    # B2B Digital
    "Seethal vargheese": {"tab": "b2b",   "subteam": "B2B Digital",       "prefix": "Seethal"},

    # ═══ Cross Functional tab ═══
    # QA
    "Fanny Dorris":      {"tab": "cross", "subteam": "QA",                "prefix": "Fanny"},
    # Mgmt Trainee
    "Akash R S":         {"tab": "cross", "subteam": "Mgmt Trainee",      "prefix": "Akash"},
    "Kiran Mathew":      {"tab": "cross", "subteam": "Mgmt Trainee",      "prefix": "Kiran"},
}

# Names in the tracker to exclude from the dashboard entirely.
EXCLUDED_FROM_DASHBOARD = {"Naitik vyas"}

# Roster used ONLY for the daily 12 PM missing-updates email.
# Deliberately excludes Akash & Kiran per requirements.
EMAIL_ROSTER = [
    "Archa Ullas", "Anna Mary", "Gautham S", "Gokul Nath", "Sreejith SL",
    "Shilpa Sara", "Arun Mahadev", "Rajeswari Menon", "Sneha S", "Fanny Dorris",
    "Bajan", "Haripriya L", "Naveen PC", "Arun Nath J", "Balavignesh P",
    "Kulwinder Singh", "Ajay Singh", "Savitha Vasanthan", "Priya Kumari",
    "Devika Sheeja", "Jofia Joseph", "Seethal vargheese",
]


# ═══════════════════════════════════════════════════════════════════════
# KPI DEFINITIONS per person
# ═══════════════════════════════════════════════════════════════════════
# Each person has an ordered list of KPI buckets. Each bucket has a
# key, display label and keyword patterns. Order matters — first match
# wins. Every person automatically gets an "Other" catch-all at the end.
#
# HTML data-keys: f"{prefix}_{bucket_key}"
KPI_BUCKETS = {
    # ─── ORGANIC SEO ─────────────────────────────────────────────────
    "Bajan": [
        ("audit",      "Technical Audits",
         ["audit", "screaming frog", "seo4ajax", "schema", "meta brand",
          "source code", "crawl", "sitemap"]),
        ("ticket",     "Ticket Management",
         ["ticket", "ipull", "jira", "prioritization"]),
        ("analysis",   "Lot/City/Page Analysis",
         ["lot level", "city", "page level", "page wise", "traffic",
          "revenue review", "rate field", "keyword", "url", "linking",
          "projection", "monthly parking", "gsc", "vendor",
          "content optimization", "fifa"]),
        ("reporting",  "Reporting",
         ["weekly report", "weekly reporting", "dashboard", "deck", "ppt",
          "leadership", "anand", "binu", "status"]),
        ("meetings",   "Meetings & Standups",
         ["standup", "meeting", "sync", "syncup", "all hands", "call",
          "alignment"]),
    ],
    "Haripriya L": [
        ("audit",      "Source Code Audits",
         ["source code audit", "audit", "schema", "seo qa", "qa validation",
          "product schema", "pre-launch", "post-launch", "post-deployment"]),
        ("content_review", "SEO Content Reviews",
         ["content review", "seo content review", "content deployment",
          "content team discussion", "port of", "internal linking",
          "interlinks", "content qa"]),
        ("keyword",    "Keyword Research",
         ["keyword", "meta title", "meta detail", "meta tag",
          "keyword ranking", "keyword position", "ranking analysis"]),
        ("reporting",  "Projections & Reports",
         ["projection", "forecast", "deck", "dashboard", "weekly review",
          "weekly report", "reporting", "insights", "growth projection",
          "performance monitoring", "tracker", "documentation",
          "improvement analysis", "ctr improvement", "recommendations",
          "sitemap update", "crawlability", "ticket created"]),
        ("meetings",   "Meetings",
         ["meeting", "standup", "stand-up", "stakeholder", "review meeting",
          "team coordination", "sync", "alignment", "walkthrough",
          "follow-up", "followup"]),
    ],
    "Naveen PC": [
        ("crawl",      "Crawl & Screaming Frog",
         ["screaming frog", "crawl", "seo4ajax", "cloudflare", "nginx",
          "crawlability"]),
        ("tickets",    "Ticket Follow-ups",
         ["ticket", "implementation", "coordinated seo ticket",
          "follow-up", "follow up"]),
        ("audit",      "Backlink & Audit Reviews",
         ["backlink", "disavow", "ahrefs", "audit", "validation",
          "qa validation", "recommendations", "keyword performance",
          "ranking movement", "monitor implementation"]),
        ("meetings",   "Standups & Meetings",
         ["standup", "meeting", "review", "sync", "stakeholder",
          "discussion", "coordination meeting"]),
    ],
    "Arun Nath J": [
        ("keyword",    "Keyword Research",
         ["keyword research", "keyword ranking", "keyword optimization"]),
        ("audit",      "Site Audits",
         ["audit", "source code", "gmb", "schema", "sitemap", "llms.txt",
          "index status", "ai visibility", "url structure"]),
        ("optimization","URL & Content Optimization",
         ["url optimization", "content optimization", "meta tag",
          "traffic projection"]),
        ("analytics",  "Analytics & Projections",
         ["analytics", "forecast", "projection", "organic session",
          "gsc analysis", "competitive analysis"]),
        ("tickets",    "Tickets & Reviews",
         ["ticket review", "jira"]),
    ],
    "Gautham S": [
        ("gap",        "GAP Pages",
         ["gap static", "gap page", "gap review"]),
        ("way",        "Way Pages",
         ["way static", "way airport"]),
        ("cruise",     "Cruise Pages",
         ["cruise parking"]),
        ("airport",    "Airport Content",
         ["airport parking content"]),
        ("meetings",   "Meetings & Interviews",
         ["meeting", "interview", "all hands", "sync", "csr",
          "content writer", "content team", "team meeting"]),
    ],
    "Gokul Nath": [
        ("gap",        "GAP Pages",
         ["gap static"]),
        ("way",        "Way Pages",
         ["way static", "way airport"]),
        ("cruise",     "Cruise/Airport Pages",
         ["cruise parking", "airport"]),
        ("reddit",     "Reddit Posts",
         ["reddit"]),
        ("meetings",   "Meetings",
         ["meeting", "all hands", "seo content team", "sync"]),
    ],
    "Sreejith SL": [
        ("gap",        "GAP Pages",
         ["gap static"]),
        ("way",        "Way Pages",
         ["way static"]),
        ("cruise",     "Cruise/Airport Pages",
         ["cruise parking", "airport parking"]),
        ("reddit",     "Reddit Posts",
         ["reddit"]),
        ("meetings",   "Meetings",
         ["meeting", "all hands", "seo content team", "report preparation"]),
    ],
    "Shilpa Sara": [
        ("uiux",       "UI/UX Copy",
         ["ui/ux copy", "ui/ux", "ux copy", "ux content"]),
        ("product",    "Product-specific Copy",
         ["way+", "way +", "r&m", "car wash", "carwash", "hyundai",
          "pleos", "way repair tech", "retention"]),
        ("meetings",   "Walkthroughs & Meetings",
         ["walkthrough", "meeting", "all hands", "interview",
          "design connect", "induction"]),
    ],
    "Arun Mahadev": [
        ("uiux",       "UI/UX Copy",
         ["ui/ux copy", "ui/ux"]),
        ("uxcontent",  "UX Content",
         ["ux copy", "ux content"]),
    ],

    # ─── OMNICHANNEL MARKETING ───────────────────────────────────────
    "Ajay Singh": [
        ("klaviyo",    "Klaviyo Campaigns",
         ["klaviyo", "flow"]),
        ("mailchimp",  "Mailchimp Campaigns",
         ["mailchimp", " mc "]),
        ("fifa_event", "FIFA & Event Campaigns",
         ["fifa", "event", "monthly parking", "summer theme",
          "independence day", "father", "jazz"]),
        ("parking",    "Parking Email Campaigns",
         ["city parking", "airport parking", "all parking", "all vertical",
          "parking email", "parking campaign", "lot pricing",
          "new segment", "new york"]),
        ("reports",    "Reports & Analysis",
         ["sendgrid", "tableau", "analysis", "report", "revenue analysis",
          "roadmap", "dashboard", "deck", "comparison", "data"]),
        ("meetings",   "Meetings",
         ["meeting", "meet", "all hands", "adhoc", "weekly", "sync",
          "review meet", "wayfarer", "training"]),
    ],
    "Kulwinder Singh": [
        ("klaviyo_qa", "Klaviyo QA & Attribution",
         ["klaviyo", "attribution"]),
        ("segmentation", "Segmentation",
         ["segment", "audience", "engagement", "unknown audience"]),
        ("connects",   "Vendor/Team Connects",
         ["connect", "netcore", "mailchimp team", "freelancer"]),
        ("reviews",    "Reviews & Templates",
         ["template", "review call", "cfo review", "review report",
          "content strategy", "ideation", "flow trigger", "flow ideation"]),
    ],
    "Savitha Vasanthan": [
        ("templates",  "Template Creation & Scheduling",
         ["template", "schedule", "gap", "airport", "cruise", "cap",
          "find", "alax"]),
    ],
    "Anna Mary": [
        ("omni_email", "Omnichannel Email Copy",
         ["omnichannel", "omni channel", "omni email",
          "writing email copy (omnichannel)"]),
        ("microsite",  "Microsite Email Copy",
         ["microsite", "microsites"]),
        ("sem_copy",   "SEM Ad Copy",
         ["sem", "google search ad", "meta ad", "ad copy", "video script"]),
        ("meetings",   "Interviews & Meetings",
         ["interview", "invigilation", "evaluation", "meeting", "all hands",
          "miscellaneous", "intake form"]),
    ],
    "Archa Ullas": [
        ("omni_email", "Omnichannel Email Copy",
         ["omni channel", "omnichannel", "cmnichannel"]),
        ("microsite",  "Microsite Email Copy",
         ["microsite"]),
        ("sem_copy",   "SEM Ad Copy",
         ["sem", "google search ad", "meta ad", "ad copy"]),
        ("push",       "Push Copy",
         ["push notification"]),
        ("meetings",   "Meetings",
         ["meeting", "email team weekly", "omnichannel weekly", "all hands",
          "leadership tracker", "intake form", "meta ad workflow"]),
    ],
    "Devika Sheeja": [
        ("posts",      "Post Creation",
         ["post", "carousel", "carousal", "static", "gap post", "cruise post",
          "fifa post", "fifa static", "nba", "diy", "tier list", "rage list",
          "one app", "grandma", "sign", "budgeting", "no convience",
          "shuttle", "ratings parking", "gas"]),
        ("video",      "Video Editing",
         ["video", "reel", "shoot", "footage", "shot life", "life at way",
          "father's day", "fathers day", "fifa video", "youtube", "yt "]),
        ("stories",    "Stories & TV",
         ["story", "tv updation", "life at way story", "cruise creative",
          "tow truck", "concert"]),
        ("templates",  "Templates & Updates",
         ["birthday", "anniversary", "bday", "template", "spotlight",
          "jira updation"]),
        ("meetings",   "Meetings & Connects",
         ["meeting", "connect with", "meet with", "discussion", "sync",
          "all hands", "wayfarer", "survey", "seat"]),
    ],
    "Jofia Joseph": [
        ("video",      "Video/Reel Production",
         ["video", "reel", "shoot", "footage", "shot life", "life at way",
          "father's day", "fathers day", "father"]),
        ("posts",      "Post Creation & Editing",
         ["post creation", "post", "carousel", "carousal", "static",
          "way post", "gap ", "cruise", "fifa", "budget", "diy",
          "shuttle", "take 5", "star review", "wayfarer", "auto repair"]),
        ("planning",   "Planning & Ideation",
         ["ideation", "ideated", "3 month plan", "one month plan",
          "planning", "researching", "researched", "recaliberation",
          "extension", "unachievable", "briefing", "brief", "catchup",
          "devika"]),
        ("calendar",   "Content Calendar",
         ["content calendar", "calendar", "utm"]),
        ("metrics",    "Metrics & Analytics",
         ["metric", "mom/yoy", "mom", "yoy", "comment", "analysis",
          "check comment", "dec-june", "previous post"]),
        ("meetings",   "Meetings",
         ["meeting", "meet with", "connect with", "sync", "all hands",
          "wayfarer", "timesheet", "time sheet", "tracker", "jira",
          "schedule", "posted", "discussion"]),
    ],

    # ─── LOCAL MARKETING ─────────────────────────────────────────────
    "Priya Kumari": [
        ("proposals",  "Client Proposals & Decks",
         ["proposal", "pitch deck", "deck", "customer deck", "pricing",
          "documentation", "business case", "research", "audit",
          "analysis", "orca", "nakayama", "delivery model",
          "calibrating"]),
        ("website",    "Website & Landing Pages",
         ["website", "landing page", "figma", "prototype", "template",
          "sk auto", "family autocare", "diablo", "autotronics", "eddie",
          "custom alignment", "champs ad"]),
        ("social",     "Social Media Strategy",
         ["social media", "social", "orm"]),
        ("csproduct",  "CS/Product Meetings",
         ["cs collaboration", "cs marketing", "cs+marketing", "cs -",
          "cs +", "product team", "engineering", "growth marketing",
          "stand-up", "standup", "cadence", "daily sync"]),
        ("campaign",   "Campaign Testing",
         ["campaign", "testing", "ads", "prepay"]),
        ("meetings",   "Meetings & Reviews",
         ["meeting", "candidate interview", "interview", "connect", "review",
          "feedback", "debrief", "sync", "hiring", "stakeholder"]),
    ],

    # ─── B2B ─────────────────────────────────────────────────────────
    "Balavignesh P": [
        ("certification","HubSpot Certifications",
         ["hubspot certification", "certification"]),
        ("workflows",  "HubSpot Workflows & Automation",
         ["hubspot", "workflow", "automation", "nurture series", "r&m",
          "carwash", "car wash", "ins ", "insurance", "email template",
          "smartlead", "neverbounce", "web analytics", "marketing insights",
          "segment", "utm", "leadfeeder", "warmly", "clay"]),
        ("newsletter", "Newsletter Design/Build",
         ["newsletter", "keith", "naga"]),
        ("naitik_review","Reviews with Naitik",
         ["naitik", "task scheduling", "task brief"]),
        ("meetings",   "Meetings",
         ["meeting", "meet", "call", "sync", "checkin", "interview",
          "all hands", "wayfarer", "b2b sync", "aileen", "review",
          "jira", "task update", "time sheet", "notes preparation",
          "naming convention"]),
    ],
    "Rajeswari Menon": [
        ("blog",       "Blog & Thought Leadership",
         ["blog", "thought leadership", "white paper"]),
        ("landing",    "Landing Page Copy",
         ["landing page", "page revamp", "page rewrite", "ndpp",
          "way+", "parking page", "solution", "industry page",
          "roadside", "b2b2c", "banner", "roadside assistance",
          "repair tech external"]),
        ("casestudy",  "Case Studies",
         ["case study", "case study development", "star park", "bridger"]),
        ("email_ad",   "Email & Ad Copy",
         ["email copy", "ad copy", "shopowner", "pc&d", "vendor payment",
          "ux copy", "shuttle tracking", "one-pager", "flyer copy",
          "shardas", "video script", "video sript", "script rewrite",
          "script", "battle card", "b2b email catch"]),
        ("meetings",   "Meetings & Reviews",
         ["meeting", "review", "sync", "aileen", "interview",
          "hiring", "town hall", "meet", "all hands", "publish",
          "seo audit", "seo optimization", "content tasks",
          "authors", "reorganization", "hubspot", "leadfeeder",
          "ga4", "monitoring", "press release", "newsroom",
          "linkedin", "ai marketing"]),
    ],
    "Sneha S": [
        ("blog",       "Blog Writing",
         ["blog", "blog revamp", "blog revision"]),
        ("blog_pub",   "Blog Publishing",
         ["blog published", "blog publishing", "completed and published",
          "published blog"]),
        ("script",     "Video Scripts",
         ["product marketing video script", "video script"]),
        ("meetings",   "Meetings",
         ["meeting", "b2b marketing", "all hands", "jira walkthrough",
          "software installation"]),
    ],
    "Seethal vargheese": [
        ("linkedin",   "LinkedIn Content Calendar",
         ["linkedin", "content calendar", "facebook", "organic post",
          "paid post", "harbour", "ipmi", "way+ roadside"]),
        ("leads",      "Warmly/Dealfront Lead Tracking",
         ["warmly", "dealfront", "warm lead", "visitor", "hubspot verification",
          "lead tracker", "partner lead", "partner data", "partner sheet",
          "partner request", "shared visitor", "shared the visitor"]),
        ("design",     "Design Coordination",
         ["design", "qr code", "flyer", "banner", "thumbnail",
          "marketplace", "amplitude", "champs", "video 1", "video 2",
          "ipmi video", "adobe", "figma", "resize", "backdrop"]),
        ("jira_pmo",   "Jira/PMO",
         ["jira", "pmo", "sprint", "scrum", "epic", "task",
          "time log", "time sheet", "tracker", "report", "b2b sync",
          "sync", "sprint planning"]),
        ("meetings",   "Meetings",
         ["meeting", "meet", "eow", "all hands", "aileen", "sync",
          "walkthrough", "call", "hr", "performance review",
          "checkin", "b2b marketing"]),
    ],

    # ─── CROSS FUNCTIONAL ────────────────────────────────────────────
    "Fanny Dorris": [
        ("gap_qa",     "GAP Pages QA",
         ["gap static"]),
        ("airport_qa", "Airport Emails QA",
         ["airport parking email"]),
        ("omni_qa",    "Omnichannel Emails QA",
         ["omnichannel email"]),
        ("blog_qa",    "B2B Blog QA",
         ["b2b blog", "blog revamp"]),
        ("way_qa",     "Way Static Pages QA",
         ["way page", "way pages", "way static",
          "way airport parking static", "city parking page"]),
        ("cruise_qa",  "Cruise QA",
         ["cruise parking"]),
        ("video_qa",   "Video Scripts QA",
         ["video script"]),
        ("meetings",   "Meetings",
         ["meeting", "seo meeting", "all hands", "timesheet"]),
    ],
    "Akash R S": [
        ("local_reports", "Local Marketing Reports",
         ["local marketing report", "local markting", "local marketing"]),
        ("weekly_reports", "Weekly Marketing Reports",
         ["weekly marketing", "weekly data", "weekly governance",
          "weekly report", "marketing weekly", "channel wise"]),
        ("sales_decks", "Sales Decks",
         ["sales pitch deck", "sales deck"]),
        ("dashboards",  "Dashboards & Automation",
         ["dashboard", "claude", "marketing activity tracker",
          "airport parking seo", "gtm", "google analytics", "ga4",
          "gsc", "search console", "semrush", "keyword analysis"]),
        ("meetings",   "Meetings & Discussions",
         ["discussion", "meeting", "sync", "sync up", "all hands",
          "syncup", "meet", "1:1", "tracker", "kt "]),
    ],
    "Kiran Mathew": [
        ("perf_track", "Marketing Performance Tracking",
         ["marketing performance", "performance tracking",
          "performance numbers", "marketing tracker", "sla tracker",
          "marketing tracking", "tracking and review"]),
        ("weekly_ppt", "Weekly Review Decks",
         ["weekly marketing review", "weekly marketing reveiw",
          "review ppt", "weekly review meeting",
          "weekly tracker", "marketing tracker dashboard"]),
        ("affiliate",  "Affiliate Marketing",
         ["affiliate", "advertise purple", "content partnerships",
          "leakage"]),
        ("dashboards", "Email/SEO Dashboards",
         ["dashboard", "email marketing", "klaviyo", "seo tracking",
          "seo tracked", "airport parking dashboard",
          "city parking dashboard", "segment", "customer segment"]),
        ("pm",         "Program Management",
         ["pm sync", "pmo", "program management", "program manager",
          "mileage tracker", "way+", "car wash", "insurance reporting",
          "ticket tracking", "gmb drop"]),
        ("meetings",   "Meetings",
         ["meeting", "meet", "sync", "all hands", "wayfarer",
          "kickoff", "co-ordinating", "coordinating", "followed up",
          "follow up", "followup", "roadmap", "attribution"]),
    ],
    # ─── NEW JOINERS ───
    "Megha B": [
        ("landing", "Airport Landing Pages",
         ["landing page","clt","abq","hnl","yyc","anc","pit","iah","cheap charlotte",
          "charlotte airport","pillar page","market page"]),
        ("supporting", "Airport Supporting Pages",
         ["terminal","airport map","tsa","flight status","accesibility","accessibility",
          "long term parking","long-term parking","supporting page"]),
        ("qa", "QA & Reviews",
         ["qa","comparitive","comparative","rework","reworked","redited","gaps"]),
        ("meetings", "Meetings & Induction",
         ["meeting","induction","team meet","briefing"]),
    ],
    "Adersh Raj": [
        ("enrichment", "Lead Enrichment & Data",
         ["enrichment","enriched","scraping","scraped","clay","data cleanup","cleaned",
          "deduplication","data quality","data validation","data matching","contact data",
          "prospect","extraction","dedup","zip code","domain enrichment","icp"]),
        ("hubspot", "HubSpot & Lead Scoring",
         ["hubspot","lead scoring","lead score","scoring model","workflow","segment",
          "segmentation","crm","form submission","enrollment"]),
        ("campaigns", "Email Campaigns & Deliverability",
         ["campaign","smartlead","neverbounce","email validation","deliverability",
          "email account","outreach","sequence","email design","email template",
          "cname","utm","domain health","arnold","nick","eric","email verification"]),
        ("dashboards", "Dashboards & Analysis",
         ["dashboard","analysis","health analysis","metrics"]),
        ("meetings", "Meetings",
         ["meeting","discussion","review","feedback","preparation"]),
    ],
    "Shivakumar patil": [
        ("klaviyo", "Klaviyo Campaigns", ["klaviyo","flow"]),
        ("mailchimp", "Mailchimp Campaigns", ["mailchimp"]),
        ("reports", "Reports & Analysis", ["report","analysis","dashboard"]),
        ("meetings", "Meetings", ["meeting","sync","all hands"]),
    ],
}


# ═══════════════════════════════════════════════════════════════════════
# HELPERS
# ═══════════════════════════════════════════════════════════════════════
def _clean(s):
    if not isinstance(s, str):
        return s
    return "".join(c if c >= " " else " " for c in s).strip()


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


def extract_count(deliverable: str, activity: str) -> int:
    """Extract deliverable count. Fallback to 1."""
    d = str(deliverable) if deliverable else ""
    a = str(activity) if activity else ""
    combined = d + " " + a
    if '"' in d or "'" in d:
        return 1
    m = re.search(
        r'\b(\d+)\s*(?:emails?|screens?|pages?|blogs?|posts?|templates?|'
        r'campaigns?|reports?|decks?|slides?|documents?|threads?|comments?|'
        r'videos?|scripts?|flows?|articles?|shots?|carousels?|carousals?|'
        r'reels?|stories?|drafts?|tickets?|tasks?|workflows?|leads?|'
        r'automations?|newsletters?|variations?|versions?|banners?|'
        r'thumbnails?|images?|copies?)\b',
        combined, re.IGNORECASE,
    )
    if m:
        return int(m.group(1))
    word_map = {"one": 1, "two": 2, "three": 3, "four": 4, "five": 5,
                "six": 6, "seven": 7, "eight": 8, "nine": 9, "ten": 10}
    for w, n in word_map.items():
        if re.search(r'\b' + w + r'\b', combined, re.IGNORECASE):
            return n
    return 1


def normalize_name(n: str) -> str:
    """Normalize name variants. Bajan BJ -> Bajan."""
    n = str(n).strip()
    if n == "Bajan BJ":
        return "Bajan"
    return n


# ═══════════════════════════════════════════════════════════════════════
# CLASSIFIER
# ═══════════════════════════════════════════════════════════════════════
def classify(name: str, activity: str, deliverable: str) -> str:
    """Return the bucket_key for this row. Unmatched -> 'other'."""
    buckets = KPI_BUCKETS.get(name, [])
    haystack = (str(activity) + " " + str(deliverable)).lower()
    for bucket_key, _label, patterns in buckets:
        for pat in patterns:
            if pat.lower() in haystack:
                return bucket_key
    return "other"


# ═══════════════════════════════════════════════════════════════════════
# TASK DATA BUILDER
# ═══════════════════════════════════════════════════════════════════════
def build_task_data(df: pd.DataFrame) -> dict:
    """Build {person_bucket_key: [tasks...]} for embedding into HTML."""
    task_data = {}

    # Pre-seed every KPI key so HTML cells always find their key.
    for name, meta in ROSTER.items():
        prefix = meta["prefix"]
        for bucket_key, _label, _pats in KPI_BUCKETS.get(name, []):
            task_data[f"{prefix}_{bucket_key}"] = []
        task_data[f"{prefix}_other"] = []

    for _, r in df.iterrows():
        name = normalize_name(r["Employee name"])
        if name not in ROSTER:
            continue
        prefix = ROSTER[name]["prefix"]

        activity = _clean(str(r.get("Activity", "")))[:200]
        deliverable = _clean(str(r.get("Work Output / Deliverable", "")))[:400]

        mins_raw = pd.to_numeric(r.get("Time Spent (in Mins)", 0), errors="coerce")
        if pd.isna(mins_raw):
            continue
        mins = int(mins_raw)
        if mins <= 0:
            continue

        freq = _clean(str(r.get("Frequency", "")))[:40] or "—"
        impact = _clean(str(r.get("Impact Bucket", "")))[:80] or "—"
        del_count = extract_count(deliverable, activity)
        mins_per_del = round(mins / del_count) if del_count else mins

        bucket = classify(name, activity, deliverable)
        key = f"{prefix}_{bucket}"

        task_data.setdefault(key, []).append({
            "date": fmt_date(r["Date"]),
            "iso": iso_date(r["Date"]),
            "activity": activity,
            "deliverable": deliverable,
            "mins": mins,
            "freq": freq,
            "impact": impact,
            "del_count": del_count,
            "mins_per_del": mins_per_del,
        })

    return task_data


# ═══════════════════════════════════════════════════════════════════════
# SHAREPOINT DOWNLOAD
# ═══════════════════════════════════════════════════════════════════════
def convert_to_download_url(share_url: str) -> str:
    """SharePoint 'anyone with the link' URLs need '?download=1' to
    force a file download instead of the web viewer."""
    if "download=1" in share_url:
        return share_url
    sep = "&" if "?" in share_url else "?"
    return f"{share_url}{sep}download=1"


def download_excel() -> Path:
    if not SHAREPOINT_URL:
        raise RuntimeError(
            "SHAREPOINT_URL secret is not set. "
            "Add it in Settings → Secrets and variables → Actions."
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
    ct = resp.headers.get("Content-Type", "")
    if "html" in ct.lower() and len(resp.content) < 50_000:
        raise RuntimeError(
            "Got an HTML page instead of the Excel file. "
            "Check the SharePoint link is 'Anyone with the link can view'."
        )
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
    tmp.write(resp.content)
    tmp.close()
    log.info(f"Downloaded {len(resp.content)/1024:.1f} KB")
    return Path(tmp.name)


# ═══════════════════════════════════════════════════════════════════════
# HTML UPDATE
# ═══════════════════════════════════════════════════════════════════════
def update_html(html_path: str, new_data: dict, raw: list = None, snapmeta: dict = None) -> bool:
    """Rewrite the const T, const RAW and const SNAPMETA blocks."""
    path = Path(html_path)
    if not path.exists():
        raise FileNotFoundError(f"{html_path} not found in repo root.")

    with open(path, encoding="utf-8") as f:
        html = f.read()

    new_json = json.dumps(new_data, ensure_ascii=False)
    new_const = f"const T = {new_json};"

    pattern = r'const T = \{.*?\};'
    if not re.search(pattern, html, flags=re.DOTALL):
        raise ValueError("Could not find 'const T = {...};' in index.html")

    html_updated = re.sub(pattern, new_const, html, flags=re.DOTALL)

    # Embed the raw rows (const RAW = [...];) — snapshot is recomputed in-browser
    if raw is not None:
        raw_const = "const RAW = " + json.dumps(raw, ensure_ascii=False) + ";"
        if re.search(r'const RAW = \[.*?\];', html_updated, flags=re.DOTALL):
            html_updated = re.sub(r'const RAW = \[.*?\];', lambda m: raw_const,
                                  html_updated, count=1, flags=re.DOTALL)

    # Embed filter-independent snapshot metadata (const SNAPMETA = {...};)
    if snapmeta is not None:
        meta_const = "const SNAPMETA = " + json.dumps(snapmeta, ensure_ascii=False) + ";"
        if re.search(r'const SNAPMETA = \{.*?\};', html_updated, flags=re.DOTALL):
            html_updated = re.sub(r'const SNAPMETA = \{.*?\};', lambda m: meta_const,
                                  html_updated, count=1, flags=re.DOTALL)

    ts = datetime.utcnow().strftime("%d %b %Y, %I:%M %p UTC")
    if "Last updated:" in html_updated:
        html_updated = re.sub(r'Last updated:.*?(?=<)',
                              f'Last updated: {ts} ', html_updated)

    changed = html_updated != html
    if changed:
        with open(path, "w", encoding="utf-8") as f:
            f.write(html_updated)
    return changed


# ═══════════════════════════════════════════════════════════════════════
# MISSING-UPDATES CHECK (used by the 12 PM IST email)
# ═══════════════════════════════════════════════════════════════════════
def check_missing_updates(df_all: pd.DataFrame) -> None:
    """Who did NOT log activity for the previous WORKING day?
    Writes missing_yesterday.json for the email step. Weekend-skip."""
    ist = timezone(timedelta(hours=5, minutes=30))
    now_ist = datetime.now(ist)
    target = (now_ist - timedelta(days=1)).date()

    if target.weekday() >= 5:
        log.info(f"Missing-check skipped — {target} was a weekend.")
        return

    target_iso = target.strftime("%Y-%m-%d")
    logged = set(
        df_all.loc[df_all['Date'].dt.strftime("%Y-%m-%d") == target_iso,
                   'Employee name'].map(normalize_name)
    )
    missing = [n for n in EMAIL_ROSTER if n not in logged]

    payload = {
        "checked_on": now_ist.strftime("%Y-%m-%d %H:%M IST"),
        "for_date": target.strftime("%a, %d %b %Y"),
        "for_date_iso": target_iso,
        "total_roster": len(EMAIL_ROSTER),
        "missing_count": len(missing),
        "missing_names": missing,
    }
    with open("missing_yesterday.json", "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)

    if missing:
        log.info(f"MISSING activity for {payload['for_date']} "
                 f"({len(missing)}/{len(EMAIL_ROSTER)}): {', '.join(missing)}")
    else:
        log.info(f"All {len(EMAIL_ROSTER)} logged activity for {payload['for_date']} ✓")


# ═══════════════════════════════════════════════════════════════════════
# MAIN
# ═══════════════════════════════════════════════════════════════════════

# ═══════════════════════════════════════════════════════════════════════
# EXECUTIVE SNAPSHOT — data extraction (verticals, AOP, initiatives)
# ═══════════════════════════════════════════════════════════════════════
REVENUE_SHAREPOINT_URL = os.environ.get("REVENUE_SHAREPOINT_URL", "")

# The 15 fixed verticals shown as their own rows; everything else -> Internal & Operational
FIXED_VERTICALS = ['Airport Parking','City Parking','Insurance','Auto Re-Finance','Way+',
    'Car Wash','Mileage Tracker','R&M','R&M SaaS','Car Wash SaaS','EV','Gas','B2B',
    'Branding','Governance']

TEAM_NORM = {'SEO':'SEO','B2B Marketing':'B2B','Omnichannel Marketing':'Omnichannel',
             'Local Marketing':'Local','Cross Functional':'Cross Functional'}

INIT_EXCLUDE = ['meeting','standup','stand up','stand-up','sync','syncup','sync up',
    'timesheet','time sheet','catch up','catchup','1:1','all hands','wayfarer',
    'weekly team','daily','induction','feedback session','360 feedback','team meet',
    'discussion','townhall','town hall','kt ','handover','leave','week off','weekend meeting']


def norm_vertical(v):
    if not v: return "Unspecified"
    s = str(v).strip().lower()
    if 'r&m' in s and 'saas' in s: return 'R&M SaaS'
    if s.startswith('r&m'): return 'R&M'
    if 'car wash saas' in s or 'carwash saas' in s: return 'Car Wash SaaS'
    if 'car wash' in s or 'carwash' in s: return 'Car Wash'
    m = {'airport parking':'Airport Parking','city parking':'City Parking','parking':'City Parking',
         'b2b':'B2B','governance':'Governance','meetings':'Meetings','internal meeting':'Meetings',
         'cross-functional meeting':'Meetings','internal':'Meetings',
         'internal activities or syncups':'Internal activities or syncups','mileage tracker':'Mileage Tracker',
         'branding':'Branding','insurance':'Insurance','way.com':'Way.com','way+':'Way+',
         'auto re-finance':'Auto Re-Finance','auto repair diy':'Auto Repair','auto repair':'Auto Repair',
         'business plan/sop':'Business Plan/SOP','ev':'EV','gas':'Gas','interviews':'Interviews'}
    return m.get(s, str(v).strip())


def _impact_weight(impact):
    if not impact: return 1.0
    s = str(impact).strip().lower()
    if s.startswith('rev') or s.startswith('revn'): return 3.0
    if s in ('retention','traffic growth'): return 2.0
    if s.startswith('non'): return 1.0
    return 1.5


def _norm_revenue(v):
    if not v: return None
    s = str(v).strip().lower()
    if s.startswith('rev') or s.startswith('revn'): return 'Revenue'
    if s.startswith('non'): return 'Non-Revenue'
    return 'Other'


def _norm_auto(v):
    if not v: return None
    s = str(v).strip().lower()
    if s == 'high': return 'High'
    if s in ('partial','partially','moderate','medium'): return 'Partial'
    if s in ('minimal','low'): return 'Minimal'
    return None


def _num(x):
    if x is None: return None
    if isinstance(x,(int,float)): return round(float(x),1)
    try: return round(float(str(x).replace('$','').replace(',','').strip()),1)
    except: return None


def download_revenue_excel():
    """Download the revenue/AOP file from SharePoint. Returns Path or None."""
    if not REVENUE_SHAREPOINT_URL:
        log.warning("REVENUE_SHAREPOINT_URL not set — AOP charts will be empty.")
        return None
    url = convert_to_download_url(REVENUE_SHAREPOINT_URL)
    log.info("Downloading Revenue Excel from SharePoint...")
    headers = {"User-Agent": "Mozilla/5.0"}
    resp = requests.get(url, headers=headers, timeout=60, allow_redirects=True)
    if resp.status_code != 200:
        log.warning(f"Revenue download failed HTTP {resp.status_code}")
        return None
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
    tmp.write(resp.content); tmp.close()
    log.info(f"Revenue file downloaded {len(resp.content)/1024:.1f} KB")
    return Path(tmp.name)


def extract_aop(revenue_path):
    """Extract the 4 AOP chart series from the revenue file."""
    if not revenue_path:
        return {}
    import openpyxl
    wb = openpyxl.load_workbook(revenue_path, read_only=True, data_only=True)
    def extract(sheet):
        ws = wb[sheet]
        rows = list(ws.iter_rows(min_row=4, max_row=15, values_only=True))
        d = {'seo_aop':[],'seo_ach':[],'seo_2025':[],'email_aop':[],'email_ach':[],'email_2025':[]}
        for r in rows:
            r = list(r) + [None]*(15-len(r)) if len(r) < 15 else r
            d['seo_aop'].append(_num(r[2]))
            d['seo_ach'].append(_num(r[4]) if _num(r[4]) is not None else _num(r[3]))
            d['seo_2025'].append(_num(r[6]))
            d['email_aop'].append(_num(r[10]))
            d['email_ach'].append(_num(r[12]) if _num(r[12]) is not None else _num(r[11]))
            d['email_2025'].append(_num(r[14]))
        return d
    try:
        ap = extract('AP - Summary ')
        cp = extract('CP-Summary')
    except Exception as e:
        log.warning(f"AOP extract failed: {e}")
        wb.close(); return {}
    wb.close()
    def series(d, kind):
        return list(d[f'{kind}_2025']) + list(d[f'{kind}_ach']), [None]*12 + list(d[f'{kind}_aop'])
    charts = {}
    for key, d, kind, title in [
        ('ap_seo', ap, 'seo', "AOP vs Achieved — airport SEO revenue ($K)"),
        ('ap_email', ap, 'email', "AOP vs Achieved — airport Email revenue ($K)"),
        ('cp_seo', cp, 'seo', "AOP vs Achieved — city SEO revenue ($K)"),
        ('cp_email', cp, 'email', "AOP vs Achieved — city Email revenue ($K)")]:
        a, t = series(d, kind)
        charts[key] = {'title':title,'actual':a,'target':t}
    return charts


def _mint_ramp(n):
    dark=(0x08,0x3d,0x30); light=(0xa9,0xf7,0xe4); cols=[]
    for i in range(n):
        t=i/max(1,n-1)
        cols.append("#%02x%02x%02x"%(round(dark[0]+(light[0]-dark[0])*t),
            round(dark[1]+(light[1]-dark[1])*t),round(dark[2]+(light[2]-dark[2])*t)))
    return cols


def build_raw(df):
    """Flatten the activity df into compact per-row records for the dashboard.

    The Executive Snapshot is now recomputed in the browser for ANY selected
    date range (MTD / month / custom), so every row it needs is embedded here
    with team/vertical already normalized and impact/automation pre-classified.
    Keys are short to keep the embedded payload small:
      d  iso date            n  employee name        tm team (normalized)
      v  vertical            m  work minutes         o  time-off minutes
      rv revenue class       w  impact weight        au automation class
      rt routine flag(0/1)   a  activity (<=120)     dl deliverable (<=200)
      bi business impact (<=120)
    """
    df = df.copy()
    df = df[~df['Employee name'].isin(EXCLUDED_FROM_DASHBOARD)]
    df = df.dropna(subset=['Date'])

    out = []
    for _, r in df.iterrows():
        mins_raw = pd.to_numeric(r.get('Time Spent (in Mins)', 0), errors='coerce')
        mins = int(mins_raw) if pd.notna(mins_raw) else 0
        toff = pd.to_numeric(r.get('Time off in Min ( Leaves taken)', 0), errors='coerce')
        off = int(toff) if pd.notna(toff) else 0
        if mins <= 0 and off <= 0:
            continue

        name = normalize_name(r['Employee name'])
        team = TEAM_NORM.get(str(r.get('Employee team', '')).strip(),
                             str(r.get('Employee team', '')).strip())
        vertical = norm_vertical(r.get('Vertical '))
        activity = _clean(str(r.get('Activity', '')))[:120]
        deliverable = _clean(str(r.get('Work Output / Deliverable', '')))[:200]
        bimpact = _clean(str(r.get('Business Outcome / Impact', '')))[:120]
        ib = r.get('Impact Bucket')

        out.append({
            'd': iso_date(r['Date']),
            'n': name,
            'tm': team,
            'v': vertical,
            'm': mins,
            'o': off,
            'rv': _norm_revenue(ib) or '',
            'w': round(_impact_weight(ib), 2),
            'au': _norm_auto(r.get('AI/Automation Potential')) or '',
            'rt': 1 if any(p in activity.lower() for p in INIT_EXCLUDE) else 0,
            'a': activity,
            'dl': deliverable,
            'bi': bimpact,
        })
    return out


def build_snapmeta(revenue_path):
    """Filter-independent snapshot metadata embedded once.

    Holds the 24-month AOP revenue charts (refreshed from the revenue file on
    every builder run — so weekly revenue-file updates flow straight through),
    plus the static maps the browser needs to render the recomputed snapshot.
    """
    charts = extract_aop(revenue_path)
    months24 = [f"{m} '25" for m in ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec']] + \
               [f"{m} '26" for m in ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec']]
    return {
        'charts': charts,
        'months24': months24,
        'vert_charts': {'Airport Parking': ['ap_seo', 'ap_email'],
                        'City Parking': ['cp_seo', 'cp_email']},
        'fixed_verticals': FIXED_VERTICALS,
        'team_colors': {'SEO': '#0BEFBA', 'B2B': '#F4C542', 'Omnichannel': '#14b8a6',
                        'Cross Functional': '#8FA89A', 'Local': '#e6a817', 'Unknown': '#6B8577'},
        'team_label': {'SEO': 'Organic SEO', 'Omnichannel': 'Omnichannel',
                       'Local': 'Local Marketing', 'B2B': 'B2B', 'Cross Functional': 'Cross Functional'},
    }


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

        # Compliance check runs every time; email is sent only on the
        # scheduled 12 PM IST run (workflow-gated).
        check_missing_updates(df.copy())

        before = len(df)
        if DATE_END:
            df = df[(df['Date'] >= DATE_START) & (df['Date'] <= DATE_END)]
        else:
            df = df[df['Date'] >= DATE_START]
        # Also drop excluded people (e.g. Naitik) so their rows don't
        # affect dedup or per-person aggregates.
        df = df[~df['Employee name'].isin(EXCLUDED_FROM_DASHBOARD)]
        log.info(f"Date filter from {DATE_START} (no upper cap): "
                 f"{before} → {len(df)} rows")

        df = df.drop_duplicates(
            subset=['Employee name', 'Date', 'Activity',
                    'Work Output / Deliverable', 'Time Spent (in Mins)']
        )
        log.info(f"After dedup: {len(df)} rows, "
                 f"{df['Employee name'].nunique()} people")

        task_data = build_task_data(df)

        # Executive Snapshot data — raw rows (recomputed in-browser for any
        # date range) + filter-independent metadata (AOP charts, maps).
        raw = build_raw(df)
        rev_path = None
        snapmeta = {}
        try:
            rev_path = download_revenue_excel()
            snapmeta = build_snapmeta(rev_path)
            log.info(f"Snapshot data: {len(raw)} raw rows, "
                     f"{len(snapmeta.get('charts',{}))} AOP charts")
        except Exception as e:
            log.warning(f"Snapmeta build failed (dashboard still updates): {e}")
        finally:
            if rev_path and rev_path.exists():
                rev_path.unlink()

        total_tasks = sum(len(v) for v in task_data.values())
        total_mins = sum(sum(r['mins'] for r in rows)
                         for rows in task_data.values())
        log.info(f"Built {len(task_data)} KPI keys, {total_tasks} task entries, "
                 f"{total_mins:,} total mins")

        changed = update_html(DASHBOARD_HTML, task_data, raw, snapmeta)
        if changed:
            log.info("index.html updated — changes will be committed")
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
    if len(sys.argv) > 1 and sys.argv[1] == "--local":
        # Local test mode: python dashboard.py --local path/to/excel.xlsx
        path = sys.argv[2]
        df = pd.read_excel(path, sheet_name=SHEET_NAME, header=1)
        data = build_task_data(df)
        print(f"KPI keys: {len(data)}")
        print(f"Total tasks: {sum(len(v) for v in data.values())}")
        print(f"Total mins: {sum(sum(r['mins'] for r in rows) for rows in data.values()):,}")
    else:
        main()
