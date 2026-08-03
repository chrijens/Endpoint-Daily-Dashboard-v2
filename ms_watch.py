#!/usr/bin/env python3
"""
Microsoft Watch - a self-updating dashboard for the Microsoft 365 roadmap
and general Microsoft announcements.

Fetches public feeds, diffs them against the previous run to flag new items,
and writes a single self-contained HTML file you can open in any browser.

Standard library only. No pip install required.

Usage:
    python3 ms_watch.py                      # fetch once, write dashboard.html
    python3 ms_watch.py --watch 60           # refresh every 60 minutes
    python3 ms_watch.py --output ~/dash.html # choose where the file lands
    python3 ms_watch.py --demo               # render with bundled sample data
"""

import argparse
import html
import json
import os
import re
import sys
import time
import urllib.error
import urllib.request
import xml.etree.ElementTree as ET
from datetime import datetime, timezone
from email.utils import parsedate_to_datetime

# ---------------------------------------------------------------------------
# Configuration - edit these to change what the dashboard tracks
# ---------------------------------------------------------------------------

ROADMAP_API = "https://www.microsoft.com/releasecommunications/api/v1/m365"

# Roadmap product tags to keep. Empty list = show everything.
# These strings must match Microsoft's own tag names exactly. The ones below
# are taken from the roadmap site's Product filter list.
PRODUCT_FILTER = [
    "Microsoft Intune",
    "Microsoft Entra",
    "Azure Active Directory",
    "Microsoft Defender for Endpoint",
    "Microsoft 365 admin center",
]

# Feeds, grouped by the column they appear in.
#
# CONFIDENCE column meanings:
#   verified  - the URL pattern is documented or was confirmed working
#   candidate - the pattern is right but the exact id/path is a best guess
#
# Run `python3 ms_watch.py --check-feeds` to test every URL and prune the
# dead ones. Anything that fails is skipped at build time, never fatal.
FEEDS = [
    # --- Microsoft: endpoint and identity -------------------------------
    ("Windows Insider", "https://blogs.windows.com/windows-insider/feed/",
     "Endpoint & identity"),
    ("Microsoft Security", "https://www.microsoft.com/en-us/security/blog/feed/",
     "Endpoint & identity"),
    # Microsoft Learn docs-change feeds. The search API builds an RSS feed from
    # any query - the only real option for products with no blog, like ConfigMgr.
    ("Intune docs", "https://learn.microsoft.com/api/search/rss"
                    "?search=Intune+what%27s+new&locale=en-us", "Endpoint & identity"),
    ("Entra docs", "https://learn.microsoft.com/api/search/rss"
                   "?search=Microsoft+Entra+what%27s+new&locale=en-us", "Endpoint & identity"),
    ("ConfigMgr docs", "https://learn.microsoft.com/api/search/rss"
                       "?search=Configuration+Manager+what%27s+new&locale=en-us",
     "Endpoint & identity"),
    ("Windows release health", "https://learn.microsoft.com/api/search/rss"
                               "?search=Windows+release+health&locale=en-us",
     "Endpoint & identity"),

    # --- Third-party service status --------------------------------------
    # NOTE: these are Statuspage incident histories, not release notes. They
    # tell you when Zoom or TeamViewer is broken, not what shipped in 15.80.
    ("Zoom status", "https://www.zoomstatus.com/history.rss", "Third-party status"),
    ("TeamViewer status", "https://status.teamviewer.com/history.rss",
     "Third-party status"),

    # --- Community and company news --------------------------------------
    ("4sysops", "https://4sysops.com/feed/", "Community & news"),
    ("Official Microsoft Blog", "https://blogs.microsoft.com/feed/", "Community & news"),
]

# Column order in the right-hand rail.
CATEGORIES = ["Endpoint & identity", "Third-party status", "Community & news"]

MAX_ROADMAP_ITEMS = 60
MAX_NEWS_ITEMS = 20          # per category
STATE_FILE = os.path.expanduser("~/.ms_watch_state.json")
USER_AGENT = "Mozilla/5.0 (compatible; MSWatchDashboard/1.0)"
TIMEOUT = 30

STATUS_ORDER = ["Rolling out", "In development", "Launched"]


# ---------------------------------------------------------------------------
# Fetch helpers
# ---------------------------------------------------------------------------

def fetch(url):
    req = urllib.request.Request(url, headers={
        "User-Agent": USER_AGENT,
        "Accept": "application/json, application/rss+xml, application/xml, text/xml, */*",
    })
    with urllib.request.urlopen(req, timeout=TIMEOUT) as resp:
        return resp.read()


def ci_get(d, *names, default=None):
    """Case-insensitive dict lookup. The roadmap API has shifted casing before."""
    if not isinstance(d, dict):
        return default
    lowered = {k.lower(): v for k, v in d.items()}
    for name in names:
        if name.lower() in lowered:
            val = lowered[name.lower()]
            if val is not None:
                return val
    return default


def strip_html(text):
    if not text:
        return ""
    text = re.sub(r"<br\s*/?>", " ", text, flags=re.I)
    text = re.sub(r"<[^>]+>", "", text)
    text = html.unescape(text)
    return re.sub(r"\s+", " ", text).strip()


def truncate(text, limit=280):
    if len(text) <= limit:
        return text
    return text[:limit].rsplit(" ", 1)[0] + "…"


def parse_date(value):
    """Handle both ISO 8601 and RFC 822 date strings."""
    if not value:
        return None
    value = value.strip()
    try:
        return parsedate_to_datetime(value)
    except (TypeError, ValueError, IndexError):
        pass
    cleaned = value.replace("Z", "+00:00")
    for candidate in (cleaned, cleaned.split(".")[0]):
        try:
            dt = datetime.fromisoformat(candidate)
            return dt if dt.tzinfo else dt.replace(tzinfo=timezone.utc)
        except ValueError:
            continue
    return None


def matches_filter(products):
    """Loose, case-insensitive match. Microsoft renames tags often - 'Microsoft
    Entra' became 'Microsoft Entra ID' - so substring matching in either
    direction avoids silently dropping everything after a rename."""
    for product in products:
        p = product.lower()
        for wanted in PRODUCT_FILTER:
            w = wanted.lower()
            if w in p or p in w:
                return True
    return False


def normalize_status(raw):
    s = (raw or "").strip().lower()
    if "rolling" in s:
        return "Rolling out"
    if "launch" in s or "released" in s:
        return "Launched"
    if "develop" in s:
        return "In development"
    return raw.strip() if raw else "Unknown"


# ---------------------------------------------------------------------------
# Sources
# ---------------------------------------------------------------------------

def fetch_roadmap():
    """Pull the Microsoft 365 public roadmap JSON API."""
    raw = fetch(ROADMAP_API)
    data = json.loads(raw.decode("utf-8", errors="replace"))
    if isinstance(data, dict):
        data = ci_get(data, "features", "value", "items", default=[])

    items = []
    for entry in data:
        tags = ci_get(entry, "tagsContainer", "tags_container", default={}) or {}
        products = [ci_get(p, "tagName", "name", default="")
                    for p in (ci_get(tags, "products", default=[]) or [])]
        platforms = [ci_get(p, "tagName", "name", default="")
                     for p in (ci_get(tags, "platforms", default=[]) or [])]
        products = [p for p in products if p]
        platforms = [p for p in platforms if p]

        if PRODUCT_FILTER and not matches_filter(products):
            continue

        feature_id = str(ci_get(entry, "id", "featureId", default="")).strip()
        modified = parse_date(ci_get(entry, "modified", "modifiedDate", "created"))

        items.append({
            "id": feature_id,
            "title": strip_html(ci_get(entry, "title", default="Untitled")),
            "description": truncate(strip_html(ci_get(entry, "description", default=""))),
            "status": normalize_status(ci_get(entry, "status", default="")),
            "products": products,
            "platforms": platforms,
            "ga_date": strip_html(str(ci_get(
                entry, "publicDisclosureAvailabilityDate", "gaDate", default="") or "")),
            "modified": modified,
            "link": f"https://www.microsoft.com/microsoft-365/roadmap?id={feature_id}",
        })

    items.sort(key=lambda i: i["modified"] or datetime.min.replace(tzinfo=timezone.utc),
               reverse=True)
    return items[:MAX_ROADMAP_ITEMS]


def fetch_news_feed(source_name, url):
    """Parse an RSS 2.0 or Atom feed into a common shape."""
    raw = fetch(url)
    root = ET.fromstring(raw)
    ns = {"atom": "http://www.w3.org/2005/Atom"}
    entries = []

    nodes = root.findall(".//item")
    is_atom = False
    if not nodes:
        nodes = root.findall(".//atom:entry", ns)
        is_atom = True

    for node in nodes:
        if is_atom:
            title = node.findtext("atom:title", default="", namespaces=ns)
            link_el = node.find("atom:link", ns)
            link = link_el.get("href") if link_el is not None else ""
            summary = node.findtext("atom:summary", default="", namespaces=ns) or \
                node.findtext("atom:content", default="", namespaces=ns)
            date_raw = node.findtext("atom:updated", default="", namespaces=ns) or \
                node.findtext("atom:published", default="", namespaces=ns)
        else:
            title = node.findtext("title", default="")
            link = node.findtext("link", default="")
            summary = node.findtext("description", default="")
            date_raw = node.findtext("pubDate", default="")

        if not title:
            continue
        entries.append({
            "id": link or title,
            "title": strip_html(title),
            "summary": truncate(strip_html(summary), 200),
            "link": link.strip(),
            "source": source_name,
            "published": parse_date(date_raw),
        })
    return entries


def gather():
    """Collect everything, recording any source that failed."""
    problems = []

    try:
        roadmap = fetch_roadmap()
    except (urllib.error.URLError, json.JSONDecodeError, ValueError, OSError) as exc:
        roadmap = []
        problems.append(("Microsoft 365 roadmap", str(exc)))

    by_category = {c: [] for c in CATEGORIES}
    for name, url, category in FEEDS:
        try:
            entries = fetch_news_feed(name, url)
        except (urllib.error.URLError, ET.ParseError, ValueError, OSError) as exc:
            problems.append((name, str(exc)))
            continue
        by_category.setdefault(category, []).extend(entries)

    for category, entries in by_category.items():
        entries.sort(key=lambda i: i["published"] or datetime.min.replace(tzinfo=timezone.utc),
                     reverse=True)
        by_category[category] = entries[:MAX_NEWS_ITEMS]

    return roadmap, by_category, problems


# ---------------------------------------------------------------------------
# State - so the dashboard can flag what changed since last run
# ---------------------------------------------------------------------------

def load_state():
    try:
        with open(STATE_FILE, "r", encoding="utf-8") as fh:
            return json.load(fh)
    except (OSError, json.JSONDecodeError):
        return {"roadmap": [], "news": []}


def save_state(roadmap, news, previous=None):
    """Persist seen ids. If a source came back empty (a failed fetch), keep the
    old ids rather than wiping them - otherwise the next good run flags
    everything as new."""
    previous = previous or {}
    state = {
        "roadmap": [i["id"] for i in roadmap] or previous.get("roadmap", []),
        "news": [i["id"] for i in news] or previous.get("news", []),
        "saved": datetime.now(timezone.utc).isoformat(),
    }
    try:
        with open(STATE_FILE, "w", encoding="utf-8") as fh:
            json.dump(state, fh)
    except OSError:
        pass


def mark_new(roadmap, news, state):
    first_run = not state.get("roadmap") and not state.get("news")
    seen_roadmap = set(state.get("roadmap", []))
    seen_news = set(state.get("news", []))
    for item in roadmap:
        item["is_new"] = not first_run and item["id"] not in seen_roadmap
    for item in news:
        item["is_new"] = not first_run and item["id"] not in seen_news
    return roadmap, news


# ---------------------------------------------------------------------------
# Rendering
# ---------------------------------------------------------------------------

def e(text):
    return html.escape(str(text or ""))


def fmt_date(dt):
    if not dt:
        return "—"
    return dt.strftime("%d %b %Y")


def relative(dt):
    if not dt:
        return ""
    delta = datetime.now(timezone.utc) - dt.astimezone(timezone.utc)
    days = delta.days
    if days <= 0:
        hours = max(delta.seconds // 3600, 0)
        return "today" if hours < 1 else f"{hours}h ago"
    if days == 1:
        return "yesterday"
    if days < 30:
        return f"{days}d ago"
    return f"{days // 30}mo ago"


CSS = """
:root {
  --ink: #14161d;
  --ink-soft: #5a6172;
  --paper: #e9edf2;
  --surface: #ffffff;
  --line: #d2d9e3;
  --dev: #6f5cf0;
  --rolling: #0d8a6a;
  --launched: #5b6478;
  --news: #b45309;
  --new: #d0245c;
}
* { box-sizing: border-box; }
body {
  margin: 0;
  background: var(--paper);
  color: var(--ink);
  font-family: "Public Sans", -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
  font-size: 15px;
  line-height: 1.5;
  -webkit-font-smoothing: antialiased;
}
.wrap { max-width: 1180px; margin: 0 auto; padding: 32px 20px 64px; }

.masthead {
  display: flex; flex-wrap: wrap; align-items: baseline;
  justify-content: space-between; gap: 12px;
  border-bottom: 2px solid var(--ink); padding-bottom: 14px;
}
.brand {
  font-family: "Bricolage Grotesque", "Public Sans", sans-serif;
  font-size: clamp(30px, 5vw, 46px);
  font-weight: 800; letter-spacing: -0.03em; line-height: 1;
}
.brand span { color: var(--ink-soft); font-weight: 500; }
.stamp {
  font-family: "DM Mono", ui-monospace, monospace;
  font-size: 12px; color: var(--ink-soft); text-align: right;
  text-transform: uppercase; letter-spacing: 0.06em;
}

.pipeline { margin: 26px 0 34px; }
.pipeline-bar {
  display: flex; height: 12px; border-radius: 2px;
  overflow: hidden; background: var(--line);
}
.seg { height: 100%; }
.seg-rolling { background: var(--rolling); }
.seg-dev { background: var(--dev); }
.seg-launched { background: var(--launched); }
.pipeline-key {
  display: flex; flex-wrap: wrap; gap: 22px; margin-top: 12px;
  font-family: "DM Mono", ui-monospace, monospace; font-size: 12px;
  text-transform: uppercase; letter-spacing: 0.06em; color: var(--ink-soft);
}
.pipeline-key b { color: var(--ink); font-size: 15px; margin-right: 6px; }
.dot { display: inline-block; width: 8px; height: 8px; border-radius: 50%; margin-right: 7px; }

.layout { display: grid; grid-template-columns: 1fr 336px; gap: 40px; align-items: start; }

.section-head {
  display: flex; align-items: baseline; justify-content: space-between;
  gap: 10px; margin-bottom: 14px;
}
.section-head h2 {
  font-family: "Bricolage Grotesque", sans-serif;
  font-size: 15px; font-weight: 700; margin: 0;
  text-transform: uppercase; letter-spacing: 0.1em;
}
.count { font-family: "DM Mono", monospace; font-size: 12px; color: var(--ink-soft); }

.controls { display: flex; flex-wrap: wrap; gap: 8px; margin-bottom: 18px; }
.filter {
  font: inherit; font-size: 13px; padding: 6px 13px;
  border: 1px solid var(--line); background: var(--surface);
  border-radius: 100px; cursor: pointer; color: var(--ink-soft);
}
.filter:hover { border-color: var(--ink-soft); }
.filter[aria-pressed="true"] { background: var(--ink); border-color: var(--ink); color: #fff; }
.search {
  font: inherit; font-size: 13px; padding: 6px 13px; flex: 1; min-width: 150px;
  border: 1px solid var(--line); border-radius: 100px; background: var(--surface);
}
.filter:focus-visible, .search:focus-visible, a:focus-visible {
  outline: 2px solid var(--dev); outline-offset: 2px;
}

.items { list-style: none; margin: 0; padding: 0; }
.item {
  position: relative; background: var(--surface); border: 1px solid var(--line);
  border-left: 3px solid var(--launched); border-radius: 3px;
  padding: 16px 18px; margin-bottom: 10px;
}
.item[data-status="In development"] { border-left-color: var(--dev); }
.item[data-status="Rolling out"] { border-left-color: var(--rolling); }
.item.hidden { display: none; }
.item-top {
  display: flex; flex-wrap: wrap; gap: 10px; align-items: center;
  font-family: "DM Mono", monospace; font-size: 11px;
  text-transform: uppercase; letter-spacing: 0.07em; color: var(--ink-soft);
  margin-bottom: 8px;
}
.status { font-weight: 500; }
.item[data-status="In development"] .status { color: var(--dev); }
.item[data-status="Rolling out"] .status { color: var(--rolling); }
.badge-new {
  background: var(--new); color: #fff; padding: 2px 7px;
  border-radius: 2px; letter-spacing: 0.1em;
}
.item h3 { margin: 0 0 6px; font-size: 17px; font-weight: 650; line-height: 1.3; }
.item h3 a { color: inherit; text-decoration: none; }
.item h3 a:hover { text-decoration: underline; }
.item p { margin: 0 0 10px; color: var(--ink-soft); font-size: 14px; }
.tags { display: flex; flex-wrap: wrap; gap: 6px; }
.tag {
  font-size: 11.5px; padding: 2px 8px; border-radius: 2px;
  background: var(--paper); color: var(--ink-soft);
}

.rail-head {
  display: flex; align-items: baseline; justify-content: space-between;
  font-family: "Bricolage Grotesque", sans-serif; font-size: 12px;
  text-transform: uppercase; letter-spacing: 0.1em; font-weight: 700;
  margin: 26px 0 4px; padding-bottom: 6px; border-bottom: 2px solid var(--ink);
}
.rail-head:first-of-type { margin-top: 0; }
.rail-count { font-family: "DM Mono", monospace; font-weight: 400; color: var(--ink-soft); }

.news-item { padding: 14px 0; border-bottom: 1px solid var(--line); }
.news-item:first-child { padding-top: 0; }
.news-item .src {
  font-family: "DM Mono", monospace; font-size: 11px; color: var(--news);
  text-transform: uppercase; letter-spacing: 0.07em;
}
.news-item h4 { margin: 5px 0 4px; font-size: 15px; font-weight: 600; line-height: 1.35; }
.news-item h4 a { color: inherit; text-decoration: none; }
.news-item h4 a:hover { text-decoration: underline; }
.news-item .when { font-size: 12px; color: var(--ink-soft); }

.notice {
  background: #fff6e8; border: 1px solid #f0d5a8; border-radius: 3px;
  padding: 12px 15px; margin-bottom: 22px; font-size: 13.5px;
}
.notice strong { display: block; margin-bottom: 3px; }
.empty { color: var(--ink-soft); font-size: 14px; padding: 20px 0; }

footer {
  margin-top: 44px; padding-top: 16px; border-top: 1px solid var(--line);
  font-family: "DM Mono", monospace; font-size: 11.5px; color: var(--ink-soft);
  text-transform: uppercase; letter-spacing: 0.06em;
}

@media (max-width: 900px) {
  .layout { grid-template-columns: 1fr; gap: 34px; }
}
@media (prefers-reduced-motion: no-preference) {
  .item, .news-item { animation: rise .35s ease both; }
  @keyframes rise { from { opacity: 0; transform: translateY(5px); } }
}
"""

JS = """
const items = Array.from(document.querySelectorAll('.item'));
const buttons = Array.from(document.querySelectorAll('.filter'));
const search = document.getElementById('search');
let active = 'all';

function apply() {
  const q = (search.value || '').toLowerCase().trim();
  let shown = 0;
  items.forEach(el => {
    const okStatus = active === 'all' || el.dataset.status === active;
    const okText = !q || el.dataset.haystack.includes(q);
    const show = okStatus && okText;
    el.classList.toggle('hidden', !show);
    if (show) shown++;
  });
  document.getElementById('shown').textContent = shown;
}

buttons.forEach(btn => btn.addEventListener('click', () => {
  active = btn.dataset.status;
  buttons.forEach(b => b.setAttribute('aria-pressed', String(b === btn)));
  apply();
}));
search.addEventListener('input', apply);
"""


def render(roadmap, news_by_category, problems, refresh_minutes=None):
    counts = {s: sum(1 for i in roadmap if i["status"] == s) for s in STATUS_ORDER}
    total = sum(counts.values()) or 1
    flat_news = [i for c in CATEGORIES for i in news_by_category.get(c, [])]
    new_count = sum(1 for i in roadmap if i.get("is_new")) + \
        sum(1 for i in flat_news if i.get("is_new"))

    seg_class = {"Rolling out": "seg-rolling", "In development": "seg-dev",
                 "Launched": "seg-launched"}
    dot_color = {"Rolling out": "var(--rolling)", "In development": "var(--dev)",
                 "Launched": "var(--launched)"}

    segments = "".join(
        f'<div class="seg {seg_class[s]}" style="width:{counts[s] / total * 100:.2f}%"></div>'
        for s in STATUS_ORDER if counts[s]
    )
    key = "".join(
        f'<span><i class="dot" style="background:{dot_color[s]}"></i>'
        f'<b>{counts[s]}</b>{e(s)}</span>'
        for s in STATUS_ORDER
    )

    filters = ['<button class="filter" data-status="all" aria-pressed="true">All</button>']
    filters += [f'<button class="filter" data-status="{e(s)}" aria-pressed="false">{e(s)}</button>'
                for s in STATUS_ORDER if counts[s]]
    filters.append('<input id="search" class="search" type="search" '
                   'placeholder="Filter by product or keyword" aria-label="Filter roadmap">')

    roadmap_html = []
    for item in roadmap:
        haystack = " ".join([item["title"], item["description"],
                             " ".join(item["products"]),
                             " ".join(item["platforms"])]).lower()
        meta = [f'<span class="status">{e(item["status"])}</span>']
        if item.get("is_new"):
            meta.append('<span class="badge-new">New</span>')
        if item["id"]:
            meta.append(f'<span>ID {e(item["id"])}</span>')
        if item["ga_date"]:
            meta.append(f'<span>GA {e(item["ga_date"])}</span>')
        if item["modified"]:
            meta.append(f'<span>Updated {e(relative(item["modified"]))}</span>')

        tags = "".join(f'<span class="tag">{e(t)}</span>'
                       for t in (item["products"] + item["platforms"])[:7])

        roadmap_html.append(f"""    <li class="item" data-status="{e(item['status'])}"
        data-haystack="{e(haystack)}">
      <div class="item-top">{''.join(meta)}</div>
      <h3><a href="{e(item['link'])}" target="_blank" rel="noopener">{e(item['title'])}</a></h3>
      <p>{e(item['description'])}</p>
      <div class="tags">{tags}</div>
    </li>""")

    if not roadmap_html:
        roadmap_html.append('<p class="empty">No roadmap items loaded. '
                            'Check the connection or the product filter in the script.</p>')

    news_html = []
    for category in CATEGORIES:
        entries = news_by_category.get(category, [])
        news_html.append(f'    <h3 class="rail-head">{e(category)}'
                         f'<span class="rail-count">{len(entries)}</span></h3>')
        if not entries:
            news_html.append('    <p class="empty">Nothing loaded for this group.</p>')
            continue
        for item in entries:
            badge = ' <span class="badge-new">New</span>' if item.get("is_new") else ""
            news_html.append(f"""    <article class="news-item">
      <div class="src">{e(item['source'])}{badge}</div>
      <h4><a href="{e(item['link'])}" target="_blank" rel="noopener">{e(item['title'])}</a></h4>
      <div class="when">{e(fmt_date(item['published']))} · {e(relative(item['published']))}</div>
    </article>""")

    notice = ""
    if problems:
        rows = "<br>".join(f"{e(name)} — {e(truncate(msg, 90))}" for name, msg in problems)
        notice = (f'<div class="notice"><strong>Some sources did not load</strong>'
                  f'{rows}</div>')

    now = datetime.now().astimezone()
    refresh_note = (f" · auto-refresh every {refresh_minutes} min"
                    if refresh_minutes else " · run the script again to refresh")
    meta_refresh = (f'<meta http-equiv="refresh" content="{refresh_minutes * 60}">'
                    if refresh_minutes else "")

    return f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
{meta_refresh}
<title>Microsoft Watch</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Bricolage+Grotesque:opsz,wght@12..96,500;12..96,700;12..96,800&family=Public+Sans:wght@400;600;650&family=DM+Mono:wght@400;500&display=swap" rel="stylesheet">
<style>{CSS}</style>
</head>
<body>
<div class="wrap">

  <header class="masthead">
    <div class="brand">Microsoft <span>Watch</span></div>
    <div class="stamp">
      Updated {e(now.strftime('%d %b %Y, %H:%M %Z'))}<br>
      {new_count} new since last run{e(refresh_note)}
    </div>
  </header>

  <section class="pipeline" aria-label="Roadmap pipeline by status">
    <div class="pipeline-bar">{segments}</div>
    <div class="pipeline-key">{key}</div>
  </section>

  {notice}

  <div class="layout">
    <section>
      <div class="section-head">
        <h2>Microsoft 365 roadmap</h2>
        <span class="count"><span id="shown">{len(roadmap)}</span> of {len(roadmap)} shown</span>
      </div>
      <div class="controls">{''.join(filters)}</div>
      <ol class="items">
{chr(10).join(roadmap_html)}
      </ol>
    </section>

    <aside>
      <div class="section-head"><h2>Release feeds</h2></div>
{chr(10).join(news_html)}
    </aside>
  </div>

  <footer>Microsoft 365 roadmap API · {len(FEEDS)} release feeds ·
  filtered to {e(', '.join(PRODUCT_FILTER)) if PRODUCT_FILTER else 'all products'}</footer>
</div>
<script>{JS}</script>
</body>
</html>"""


# ---------------------------------------------------------------------------
# Demo fixture
# ---------------------------------------------------------------------------

def demo_data():
    now = datetime.now(timezone.utc)
    from datetime import timedelta
    roadmap = [
        {"id": "567316", "title": "Outlook: Reply position notification",
         "description": "When replying to an email conversation and the message being replied "
                        "to is not the latest in the conversation, Outlook for Windows and web "
                        "will display a notification indicating this.",
         "status": "In development", "products": ["Outlook"], "platforms": ["Web"],
         "ga_date": "August CY2026", "modified": now - timedelta(hours=6),
         "link": "https://www.microsoft.com/microsoft-365/roadmap?id=567316", "is_new": True},
        {"id": "512429", "title": "Microsoft Copilot: Seamless search and chat integration",
         "description": "Microsoft 365 Chat brings conversational AI into Copilot Search, "
                        "letting users move from finding information to acting on it.",
         "status": "Rolling out",
         "products": ["Microsoft Copilot (Microsoft 365)", "Microsoft 365 app"],
         "platforms": ["Desktop", "Web"], "ga_date": "June CY2026",
         "modified": now - timedelta(days=2),
         "link": "https://www.microsoft.com/microsoft-365/roadmap?id=512429", "is_new": True},
        {"id": "492622", "title": "The next generation of file and folder sharing",
         "description": "A third generation of the Microsoft 365 sharing experience built "
                        "around the hero link, a single link that controls access to your files.",
         "status": "In development",
         "products": ["OneDrive", "SharePoint", "Word", "Excel", "PowerPoint"],
         "platforms": ["Desktop", "Web", "iOS", "Android"], "ga_date": "",
         "modified": now - timedelta(days=5),
         "link": "https://www.microsoft.com/microsoft-365/roadmap?id=492622", "is_new": False},
        {"id": "512431", "title": "Planner: Custom templates",
         "description": "Custom templates let you create reusable, pre-designed layouts "
                        "tailored to your organization's needs.",
         "status": "In development", "products": ["Planner"],
         "platforms": ["Web", "Desktop", "Mac"], "ga_date": "",
         "modified": now - timedelta(days=11),
         "link": "https://www.microsoft.com/microsoft-365/roadmap?id=512431", "is_new": False},
        {"id": "478611", "title": "Microsoft Teams: Facilitator agent",
         "description": "The Facilitator agent works alongside your team to help manage "
                        "meetings, take notes and track follow-ups.",
         "status": "Launched", "products": ["Microsoft Teams"],
         "platforms": ["Desktop", "Web", "Mac"], "ga_date": "",
         "modified": now - timedelta(days=40),
         "link": "https://www.microsoft.com/microsoft-365/roadmap?id=478611", "is_new": False},
        {"id": "561200", "title": "Teams: Asynchronous large file uploads",
         "description": "Users can keep sending messages while a file uploads in the "
                        "background, reducing perceived latency in collaboration.",
         "status": "Rolling out", "products": ["Microsoft Teams"],
         "platforms": ["Desktop", "Web"], "ga_date": "August CY2026",
         "modified": now - timedelta(days=1),
         "link": "https://www.microsoft.com/microsoft-365/roadmap?id=561200", "is_new": False},
    ]
    news = {
        "Endpoint & identity": [
            {"id": "e1", "title": "Sample: What's new in Microsoft Intune - July",
             "summary": "", "link": "https://techcommunity.microsoft.com/",
             "source": "Intune Blog", "published": now - timedelta(hours=9), "is_new": True},
            {"id": "e2", "title": "Sample: Windows 11 servicing update and KMS attestation",
             "summary": "", "link": "https://techcommunity.microsoft.com/",
             "source": "Windows IT Pro", "published": now - timedelta(days=2), "is_new": True},
            {"id": "e3", "title": "Sample: Entra conditional access policy changes",
             "summary": "", "link": "https://techcommunity.microsoft.com/",
             "source": "Entra / Identity", "published": now - timedelta(days=4), "is_new": False},
            {"id": "e4", "title": "Sample: Configuration Manager 2603 documentation update",
             "summary": "", "link": "https://learn.microsoft.com/",
             "source": "ConfigMgr docs", "published": now - timedelta(days=8), "is_new": False},
        ],
        "Third-party status": [
            {"id": "t1", "title": "Sample: Zoom - degraded performance for meeting recordings",
             "summary": "", "link": "https://www.zoomstatus.com/",
             "source": "Zoom status", "published": now - timedelta(days=1), "is_new": True},
            {"id": "t2", "title": "Sample: TeamViewer - scheduled maintenance on EU routers",
             "summary": "", "link": "https://status.teamviewer.com/",
             "source": "TeamViewer status", "published": now - timedelta(days=5),
             "is_new": False},
        ],
        "Community & news": [
            {"id": "c1", "title": "Sample: Restrict Active Directory account logon hours",
             "summary": "", "link": "https://4sysops.com/",
             "source": "4sysops", "published": now - timedelta(hours=14), "is_new": True},
            {"id": "n1", "title": "Sample: quarterly earnings and cloud growth",
             "summary": "", "link": "https://blogs.microsoft.com/",
             "source": "Official Microsoft Blog", "published": now - timedelta(hours=20),
             "is_new": False},
        ],
    }
    return roadmap, news, []


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

def build(output, refresh_minutes=None, demo=False):
    if demo:
        roadmap, news_by_category, problems = demo_data()
    else:
        roadmap, news_by_category, problems = gather()
        previous = load_state()
        # mark_new mutates the item dicts, which the category lists still hold.
        flat = [i for c in CATEGORIES for i in news_by_category.get(c, [])]
        roadmap, flat = mark_new(roadmap, flat, previous)
        save_state(roadmap, flat, previous)

    page = render(roadmap, news_by_category, problems, refresh_minutes)
    with open(output, "w", encoding="utf-8") as fh:
        fh.write(page)

    news_total = sum(len(v) for v in news_by_category.values())
    stamp = datetime.now().strftime("%H:%M:%S")
    print(f"[{stamp}] wrote {output} — "
          f"{len(roadmap)} roadmap items, {news_total} feed items"
          + (f", {len(problems)} source(s) failed" if problems else ""))
    for name, msg in problems:
        print(f"    ! {name}: {truncate(msg, 120)}", file=sys.stderr)

    return bool(roadmap or news_total)


def check_feeds():
    """Test every configured feed and report which ones actually work.

    Several vendors here publish no official RSS, so the shipped URLs are
    educated guesses. This tells you which to keep in about ten seconds.
    """
    print(f"Testing {len(FEEDS) + 1} sources\n")
    width = max([len(n) for n, _, _ in FEEDS] + [len("M365 roadmap")]) + 2
    ok = dead = 0

    try:
        items = fetch_roadmap()
        print(f"  PASS  {'M365 roadmap'.ljust(width)} {len(items)} items after product filter")
        if not items:
            print(f"        {''.ljust(width)} check PRODUCT_FILTER - tag names may have changed")
        ok += 1
    except Exception as exc:
        print(f"  FAIL  {'M365 roadmap'.ljust(width)} {truncate(str(exc), 60)}")
        dead += 1

    for name, url, category in FEEDS:
        try:
            entries = fetch_news_feed(name, url)
            if entries:
                newest = fmt_date(entries[0]["published"])
                print(f"  PASS  {name.ljust(width)} {len(entries)} items, newest {newest}")
                ok += 1
            else:
                print(f"  EMPTY {name.ljust(width)} parsed but returned nothing")
                dead += 1
        except Exception as exc:
            print(f"  FAIL  {name.ljust(width)} {truncate(str(exc), 60)}")
            dead += 1

    print(f"\n{ok} working, {dead} to fix or remove.")
    if dead:
        print("Delete the failing lines from FEEDS, or correct the URL.")
    return dead == 0


def main():
    global STATE_FILE

    ap = argparse.ArgumentParser(description="Build a Microsoft announcements dashboard.")
    ap.add_argument("--output", default="dashboard.html", help="output HTML path")
    ap.add_argument("--state", default=None,
                    help="where to keep the seen-items file, needed for the New badge "
                         "(default: ~/.ms_watch_state.json)")
    ap.add_argument("--refresh", type=int, metavar="MINUTES",
                    help="add a meta-refresh so an open tab reloads itself")
    ap.add_argument("--watch", type=int, metavar="MINUTES",
                    help="keep running and rebuild every N minutes")
    ap.add_argument("--demo", action="store_true",
                    help="render with bundled sample data, no network needed")
    ap.add_argument("--check-feeds", action="store_true",
                    help="test every configured source and report which work")
    ap.add_argument("--strict", action="store_true",
                    help="exit non-zero if every source failed (useful in CI)")
    args = ap.parse_args()

    if args.check_feeds:
        sys.exit(0 if check_feeds() else 1)

    if args.state:
        STATE_FILE = os.path.expanduser(args.state)
    parent = os.path.dirname(STATE_FILE)
    if parent:
        os.makedirs(parent, exist_ok=True)

    output = os.path.expanduser(args.output)
    out_dir = os.path.dirname(output)
    if out_dir:
        os.makedirs(out_dir, exist_ok=True)

    if args.watch:
        refresh = args.refresh or args.watch
        print(f"Watching. Rebuilding {output} every {args.watch} min. Ctrl-C to stop.")
        while True:
            try:
                build(output, refresh, args.demo)
            except Exception as exc:                      # keep the loop alive
                print(f"  build failed: {exc}", file=sys.stderr)
            time.sleep(args.watch * 60)
    else:
        ok = build(output, args.refresh, args.demo)
        if args.strict and not ok:
            print("Every source failed. Exiting non-zero.", file=sys.stderr)
            sys.exit(1)


if __name__ == "__main__":
    main()
