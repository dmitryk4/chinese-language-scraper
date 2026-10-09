#!/usr/bin/env python3
"""
Jereh Group (SZSE: 002353) - full document scraper
==================================================

Downloads everything public about Jereh into one folder:

  1. Exchange filings: every announcement since listing, plus investor
     relations meeting records, from cninfo.com.cn (China's official
     disclosure site). If cninfo fails for a year, that year's filings are
     pulled from szse.cn (the Shenzhen Stock Exchange's own site) instead.
  2. Websites: every page (saved as text) and every document (PDF, Word,
     Excel, PowerPoint, ZIP) on jereh.com in Chinese and English, plus the
     Jereh group sites (Jereh Power, Jereh Gas, Jereh Services, ...).
  3. Text and keywords: extracts the text of every PDF and Office file and
     flags documents that mention J&F, FTAI, gas turbines, data centers, etc.

Setup (once):
    pip install requests beautifulsoup4 pypdf

Run:
    python jereh_scraper.py --check      # 1-minute test of every source
    python jereh_scraper.py              # scrape everything
    python jereh_scraper.py --scan-only  # rebuild the reports only

Options:
    --source filings|website|all   what to scrape (default: all)
    --since 2025-01-01             filings from this date only (default: all)
    --main-site-only               skip the Jereh group sites
    --max-pages 500                stop the website crawl after N pages
    --delay 0.5                    seconds between requests (default: 1.0)
    --keywords "Kawasaki,LM2500"   extra keywords to flag
    --out D:/jereh                 output folder (default: ./jereh_docs)
    --filings-from szse            use szse.cn instead of cninfo
    --save-html                    also keep the raw HTML of every page
    --fresh                        restart the website crawl from scratch

A full run takes several hours: thousands of filings and pages, fetched
politely one at a time. You can stop it at any time with Ctrl+C and run the
same command again. Finished downloads are skipped, failed ones are retried,
and the website crawl picks up where it stopped.

Output, inside the output folder:
    filings/cninfo/<year>/<date>_<id>_<title>.pdf  filings (+ .txt with text)
    filings/szse/<year>/...                         only if szse.cn was used
    website/<site>/pages/*.txt                      text of every web page
    website/<site>/files/*                          documents from the sites
    index.csv          every item: source, date, title, link, file, keywords
    keyword_hits.csv   one row per document per keyword, with snippets
    scrape_log.txt     full log of every run
    manifest.jsonl     record of everything downloaded (used to resume)
"""

import argparse
import csv
import datetime as dt
import gzip
import hashlib
import html
import io
import json
import logging
import os
import random
import re
import sys
import time
import warnings
import zipfile
from collections import deque
from pathlib import Path
from urllib import robotparser
from urllib.parse import parse_qsl, unquote, urlencode, urljoin, urlparse, urlunparse

try:
    import requests
    from bs4 import BeautifulSoup
except ImportError:
    sys.exit("Missing packages. Run this first:\n    pip install requests beautifulsoup4 pypdf")

try:
    from pypdf import PdfReader
except ImportError:  # text extraction from PDFs is skipped without it
    PdfReader = None

logging.getLogger("pypdf").setLevel(logging.ERROR)
warnings.filterwarnings("ignore", module="bs4")
warnings.filterwarnings("ignore", module="pypdf")

# ===================================================================== settings
STOCK_CODE = "002353"

CNINFO_BASE = "https://www.cninfo.com.cn"
CNINFO_STATIC = "https://static.cninfo.com.cn/"
SZSE_BASE = "https://www.szse.cn"
SZSE_STATIC = "https://disc.static.szse.cn"

# Where the website crawl starts. Every page and document linked from these
# pages on the same sites is followed.
MAIN_SITE_ROOTS = [
    "https://www.jereh.com/cn/",
    "https://jereh.com/en/",
    "https://www.jereh.com/cn/site/sitemap",
    "https://jereh.com/en/site/sitemap",
    "https://www.jereh.com/cn/investor/",
    "https://jereh.com/en/investor/",
    "https://jereh.com/en/news/press-release",
]
# Jereh group sites listed in the jereh.com footer ("Website Group").
GROUP_SITE_ROOTS = [
    "https://en.jereh-power.com/",
    "https://www.jereh-power.com/",
    "https://www.jereh-pe.com/",
    "https://offshore.jereh-pe.com/en",
    "https://www.jereh-ne.com/en/",
    "https://www.jereh-gas.com/",
    "https://www.jereh-engineering.com/",
    "https://www.jereh-env.com/en/",
    "https://www.jereh-services.com/en/",
    "https://www.jereh-eet.com/",
    "https://www.jerehglobal.com/",
    "https://www.jereh-powertech.com/",
    "https://www.jereh-gk.com/",
    "https://en.jereh-hengri.com/",
    "https://www.jasonenergy.com/en/",
    "https://www.grifco-oiltools.com/en/",
]

# Documents mentioning any of these are listed in keyword_hits.csv.
# Matching ignores spaces, line breaks and upper/lower case.
KEYWORDS = [
    # the J&F joint venture and FTAI
    "J&F", "J&F Power", "FTAI", "Mod-1", "CFM56", "aeroderivative", "航改",
    # gas turbines, power and data centers
    "燃气轮机", "燃机", "发电机组", "数据中心", "云服务", "算力", "北美",
    # ownership and structure
    "控股子公司", "合资", "持股比例", "出资",
    # other turbine suppliers and competitors
    "Siemens", "西门子", "Baker Hughes", "贝克休斯", "Kawasaki", "川崎",
    "Everllence", "GE Vernova", "LM2500", "LM6000",
    # regulation
    "14420号", "Executive Order 14420", "CFIUS", "关税",
]

DOC_EXTENSIONS = (".pdf", ".doc", ".docx", ".xls", ".xlsx", ".ppt", ".pptx",
                  ".wps", ".csv", ".zip", ".rar", ".7z")
SKIP_EXTENSIONS = (".jpg", ".jpeg", ".png", ".gif", ".bmp", ".webp", ".svg", ".ico",
                   ".tif", ".tiff", ".psd", ".mp4", ".m4v", ".mov", ".avi", ".wmv",
                   ".flv", ".mkv", ".webm", ".mp3", ".wav", ".m4a", ".css", ".js",
                   ".json", ".xml", ".woff", ".woff2", ".ttf", ".otf", ".eot", ".exe",
                   ".msi", ".apk", ".dmg", ".iso", ".swf")
DOC_CONTENT_TYPES = {
    "application/pdf": ".pdf",
    "application/x-pdf": ".pdf",
    "application/msword": ".doc",
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document": ".docx",
    "application/vnd.ms-excel": ".xls",
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": ".xlsx",
    "application/vnd.ms-powerpoint": ".ppt",
    "application/vnd.openxmlformats-officedocument.presentationml.presentation": ".pptx",
    "application/zip": ".zip",
    "application/x-zip-compressed": ".zip",
    "application/x-rar-compressed": ".rar",
    "application/vnd.rar": ".rar",
    "application/x-7z-compressed": ".7z",
    "text/csv": ".csv",
}
HTML_TYPES = ("text/html", "application/xhtml+xml", "")
TEXT_EXTRACTABLE = (".pdf", ".docx", ".xlsx", ".pptx", ".html", ".htm", ".csv")

USER_AGENT = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
              "(KHTML, like Gecko) Chrome/126.0 Safari/537.36")
ROBOTS_AGENT = "Mozilla"
BROWSER_HEADERS = {
    "User-Agent": USER_AGENT,
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8",
}

API_DELAY_FACTOR = 2.5         # search APIs get a longer pause than page loads
RETRY_WAITS = (5, 20, 60)      # seconds to wait before each retry
MAX_FILE_MB = 300              # skip website files bigger than this
MAX_PAGE_MB = 15
MAX_QUERY_VARIANTS = 2000      # e.g. ?nowPage=1..2000 for one list page
HOST_FAILURE_LIMIT = 8         # pause a site after this many failures in a row
STATE_SAVE_SECONDS = 30
BEIJING = dt.timezone(dt.timedelta(hours=8))
DEFAULT_ORG_SOURCE = "standard Shenzhen pattern"

RETRY_STATUSES = (429, 500, 502, 503, 504)
API_RETRY_STATUSES = (403, 429, 500, 502, 503, 504)
NO_RETRY_HINTS = ("SSLError", "CERTIFICATE_VERIFY_FAILED", "Failed to resolve",
                  "NameResolutionError", "getaddrinfo failed", "Name or service not known",
                  "nodename nor servname", "No address associated", "InvalidURL",
                  "MissingSchema", "InvalidSchema", "TooManyRedirects")
BOT_CHALLENGE_HINTS = ("acw_sc__v2", "jsl_clearance", "yunsuo_session_verify",
                       "_guard/auto.js", "challenge-platform", "wzws_cid", "captcha")

# ====================================================================== helpers
_log_file = None


def log(msg=""):
    print(msg, flush=True)
    if _log_file is not None:
        _log_file.write(msg + "\n")
        _log_file.flush()


class FetchError(Exception):
    pass


def short_error(exc):
    text = str(exc)
    name = exc.__class__.__name__
    caused = re.search(r"Caused by (\w+)\(", text)
    if caused:
        name = caused.group(1)
    return f"{name}: {text[:180]}"


FRIENDLY_ERRORS = (
    ("ProxyError", "could not connect through the network proxy"),
    ("SSLError", "the secure connection (SSL) failed"),
    ("CERTIFICATE_VERIFY_FAILED", "the secure connection (SSL) failed"),
    ("NameResolutionError", "the site's address could not be found (DNS)"),
    ("Failed to resolve", "the site's address could not be found (DNS)"),
    ("getaddrinfo failed", "the site's address could not be found (DNS)"),
    ("Name or service not known", "the site's address could not be found (DNS)"),
    ("nodename nor servname", "the site's address could not be found (DNS)"),
    ("ConnectTimeout", "the connection timed out"),
    ("ReadTimeout", "the site stopped responding"),
    ("Connection refused", "the connection was refused"),
    ("ConnectionResetError", "the site reset the connection (it may be limiting requests)"),
    ("RemoteDisconnected", "the site closed the connection"),
    ("ChunkedEncodingError", "the download was cut off"),
)


def friendly_error(exc):
    raw = f"{exc.__class__.__name__}: {exc}"
    for hint, text in FRIENDLY_ERRORS:
        if hint in raw:
            return text
    return short_error(exc)


def plain(exc):
    """An error message without the URL that FetchError appends."""
    return re.sub(r"\s\[[^\]]*\]$", "", str(exc))


def truncate_bytes(text, max_bytes):
    raw = text.encode("utf-8")
    if len(raw) <= max_bytes:
        return text
    return raw[:max_bytes].decode("utf-8", "ignore")


_BAD_CHARS = re.compile(r'[\\/:*?"<>|\x00-\x1f]+')


def safe_name(text, max_bytes=150):
    """A file-name-safe version of text, short enough for Windows and macOS."""
    text = _BAD_CHARS.sub("_", text or "")
    text = re.sub(r"\s+", "_", text).strip("._ ")
    text = truncate_bytes(text, max_bytes).rstrip("._ ")
    return text or "untitled"


def clean_title(text):
    text = re.sub(r"<[^>]+>", "", text or "")
    return " ".join(html.unescape(text).split())


def norm_title(text):
    """Title reduced to letters and digits, for spotting the same filing twice."""
    text = re.sub(r"^\s*杰瑞股份\s*[:：]\s*", "", clean_title(text))
    return re.sub(r"[\W_]+", "", text).lower()


def write_atomic(path, data):
    """Write via a temporary file so an interrupted run never leaves half a file."""
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_name(path.name + ".part")
    with open(tmp, "wb") as fh:
        fh.write(data)
    os.replace(tmp, path)


def year_ranges(since, until):
    first, last = dt.date.fromisoformat(since), dt.date.fromisoformat(until)
    for year in range(first.year, last.year + 1):
        start = max(first, dt.date(year, 1, 1))
        end = min(last, dt.date(year, 12, 31))
        yield start.isoformat(), end.isoformat()


def read_capped(resp, limit_mb):
    limit = limit_mb * 1024 * 1024
    length = resp.headers.get("Content-Length", "")
    if length.isdigit() and int(length) > limit:
        raise FetchError(f"larger than {limit_mb} MB, skipped")
    buf = io.BytesIO()
    try:
        for chunk in resp.iter_content(65536):
            buf.write(chunk)
            if buf.tell() > limit:
                raise FetchError(f"larger than {limit_mb} MB, skipped")
    except requests.RequestException as e:
        raise FetchError(f"download interrupted ({friendly_error(e)})")
    return buf.getvalue()


# ================================================================ HTTP client
class Fetcher:
    """requests.Session with polite pacing per site and automatic retries."""

    def __init__(self, delay):
        self.session = requests.Session()
        self.session.headers.update(BROWSER_HEADERS)
        self.web_gap = max(0.0, delay)
        self.file_gap = max(0.0, delay)
        self.api_gap = max(0.0, delay * API_DELAY_FACTOR)
        self.max_retries = 3
        self._last = {}

    def _pace(self, url, gap):
        if gap <= 0:
            return
        host = urlparse(url).netloc
        wait = gap * random.uniform(0.7, 1.3) - (time.time() - self._last.get(host, 0.0))
        if wait > 0:
            time.sleep(wait)

    def request(self, method, url, gap=None, retries=3, retry_statuses=RETRY_STATUSES, **kw):
        gap = self.web_gap if gap is None else gap
        retries = min(retries, self.max_retries)
        kw.setdefault("timeout", (20, 120))
        error, retry_after = "unknown error", None
        for attempt in range(retries + 1):
            self._pace(url, gap)
            try:
                resp = self.session.request(method, url, **kw)
            except requests.RequestException as e:
                error, retry_after = friendly_error(e), None
                self._last[urlparse(url).netloc] = time.time()
                if any(hint in f"{e.__class__.__name__}: {e}" for hint in NO_RETRY_HINTS):
                    break
            else:
                self._last[urlparse(url).netloc] = time.time()
                if resp.status_code not in retry_statuses:
                    return resp
                error, retry_after = f"HTTP {resp.status_code}", resp.headers.get("Retry-After")
                resp.close()
            if attempt < retries and RETRY_WAITS:
                wait = RETRY_WAITS[min(attempt, len(RETRY_WAITS) - 1)]
                if retry_after and retry_after.isdigit():
                    wait = max(wait, min(int(retry_after), 300))
                wait *= random.uniform(0.8, 1.2)
                log(f"      {error} - retrying in {wait:.0f}s")
                time.sleep(wait)
        raise FetchError(f"{error} [{url}]")

    def json(self, method, url, what, **kw):
        """Request JSON, retrying if a blocking page comes back instead."""
        kw.setdefault("retry_statuses", API_RETRY_STATUSES)
        problem = ""
        for attempt in range(3):
            resp = self.request(method, url, **kw)
            if resp.status_code != 200:
                raise FetchError(f"{what}: HTTP {resp.status_code}")
            try:
                return resp.json()
            except ValueError:
                problem = (f"{what}: the reply was not data (the site may be blocking "
                           "automated requests; try again later or use a larger --delay)")
            if attempt < 2 and RETRY_WAITS:
                time.sleep(RETRY_WAITS[min(attempt, len(RETRY_WAITS) - 1)])
        raise FetchError(problem)


# =================================================================== manifest
class Manifest:
    """Append-only log of every item, so runs can be stopped and resumed."""

    def __init__(self, out):
        self.path = out / "manifest.jsonl"
        self.items = {}
        if self.path.exists():
            with open(self.path, encoding="utf-8") as fh:
                for line in fh:
                    try:
                        rec = json.loads(line)
                        self.items[rec["key"]] = rec
                    except (ValueError, KeyError, TypeError):
                        pass  # a line cut short by an interrupted run
        self._fh = open(self.path, "a", encoding="utf-8")

    def add(self, rec):
        if self.items.get(rec["key"]) == rec:
            return
        self.items[rec["key"]] = rec
        self._fh.write(json.dumps(rec, ensure_ascii=False) + "\n")
        self._fh.flush()

    def close(self):
        self._fh.close()


def is_ok(rec):
    return str(rec.get("status", "")).startswith("ok")


# ============================================================ exchange filings
def cninfo_headers():
    return {"Accept": "application/json, text/javascript, */*; q=0.01",
            "X-Requested-With": "XMLHttpRequest",
            "Origin": CNINFO_BASE,
            "Referer": CNINFO_BASE + "/new/commonUrl/pageOfSearch?url=disclosure/list/search"}


def szse_headers():
    return {"Accept": "application/json, text/javascript, */*; q=0.01",
            "Content-Type": "application/json",
            "X-Requested-With": "XMLHttpRequest",
            "X-Request-Type": "ajax",
            "Origin": SZSE_BASE,
            "Referer": SZSE_BASE + "/disclosure/listed/notice/index.html"}


def warm_up(f, url):
    """Open the site's search page first, as a browser would (sets cookies)."""
    try:
        f.request("GET", url, gap=f.api_gap, retries=1).close()
    except FetchError:
        pass


def cninfo_org_id(f):
    """cninfo identifies companies by an internal orgId as well as the stock code."""
    try:
        data = f.json("GET", CNINFO_BASE + "/new/data/szse_stock.json", "cninfo company list",
                      gap=f.api_gap)
        for row in (data or {}).get("stockList") or []:
            if str(row.get("code")) == STOCK_CODE and row.get("orgId"):
                return row["orgId"], "cninfo company list"
    except (FetchError, AttributeError) as e:
        log(f"  cninfo company list lookup failed: {plain(e)}")
    try:
        data = f.json("POST", CNINFO_BASE + "/new/information/topSearch/query", "cninfo search",
                      data={"keyWord": STOCK_CODE, "maxNum": 10}, headers=cninfo_headers(),
                      gap=f.api_gap)
        for row in data if isinstance(data, list) else []:
            if str(row.get("code")) == STOCK_CODE and row.get("orgId"):
                return row["orgId"], "cninfo search"
    except (FetchError, AttributeError) as e:
        log(f"  cninfo search lookup failed: {plain(e)}")
    return "gssz0" + STOCK_CODE, DEFAULT_ORG_SOURCE


CNINFO_PAGE_SIZE = 30


def cninfo_list(f, org_id, tab, start, end):
    """All cninfo announcements for the stock in a date range. tab: fulltext or relation."""
    found, foreign, page, ids = [], 0, 1, set()
    while True:
        form = {"pageNum": page, "pageSize": CNINFO_PAGE_SIZE, "column": "szse",
                "tabName": tab, "plate": "", "stock": f"{STOCK_CODE},{org_id}",
                "searchkey": "", "secid": "", "category": "", "trade": "",
                "seDate": f"{start}~{end}", "sortName": "", "sortType": "",
                "isHLtitle": "true"}
        data = f.json("POST", CNINFO_BASE + "/new/hisAnnouncement/query",
                      "cninfo filing search", data=form, headers=cninfo_headers(),
                      gap=f.api_gap)
        if not isinstance(data, dict):
            raise FetchError("cninfo filing search: unexpected reply")
        batch = data.get("announcements") or []
        new = {str(item.get("announcementId")) for item in batch} - ids
        ids |= new
        for item in batch:
            if str(item.get("secCode") or STOCK_CODE) == STOCK_CODE:
                found.append(item)
            else:
                foreign += 1
        if foreign and not found:
            raise FetchError("cninfo returned other companies' filings "
                             "(its stock filter did not work)")
        has_more = data.get("hasMore")
        if has_more is None:
            has_more = len(batch) >= CNINFO_PAGE_SIZE
        if not batch or not new or not has_more or page >= 1000:
            return found  # (not new: the site repeated a page, so stop paging)
        page += 1


def cninfo_record(item, tab, org_id):
    aid = str(item.get("announcementId") or "").strip() or "noid"
    adjunct = str(item.get("adjunctUrl") or "").strip().lstrip("/")
    title = clean_title(item.get("announcementTitle")) or "untitled"
    try:
        stamp = float(item.get("announcementTime"))
        date = dt.datetime.fromtimestamp(stamp / 1000, BEIJING).date().isoformat()
    except (TypeError, ValueError, OverflowError, OSError):
        found = re.search(r"\d{4}-\d{2}-\d{2}", adjunct)
        date = found.group(0) if found else "unknown-date"
    ext = (os.path.splitext(adjunct)[1] or ".pdf").lower()
    kind = "ir_record" if (tab == "relation" or "投资者关系活动记录" in title) else "filing"
    name = f"{date}_{safe_name(aid, 30)}_{safe_name(title, 140)}{ext}"
    return {
        "key": "cninfo:" + aid, "source": "cninfo", "kind": kind, "date": date,
        "title": title,
        "url": CNINFO_STATIC + adjunct if adjunct else "",
        "page_url": (f"{CNINFO_BASE}/new/disclosure/detail?stockCode={STOCK_CODE}"
                     f"&announcementId={aid}&orgId={org_id}&announcementTime={date}"),
        "file": Path("filings", "cninfo", date[:4], name).as_posix(),
    }


def cninfo_records(f, org_id, start, end, tabs):
    """Filings (tab fulltext) plus investor-relations records (tab relation)."""
    found = {}
    for tab in list(tabs):
        try:
            items = cninfo_list(f, org_id, tab, start, end)
        except FetchError as e:
            if tab == "fulltext":
                raise
            tabs.remove(tab)
            log(f"  The investor-relations records search failed ({plain(e)}); skipping it from "
                "now on. Most of these records also appear among the regular filings.")
            continue
        for item in items:
            rec = cninfo_record(item, tab, org_id)
            found.setdefault(rec["key"], rec)
    return list(found.values())


SZSE_PAGE_SIZE = 50


def szse_list(f, start, end):
    found, page, ids = [], 1, set()
    while True:
        payload = {"seDate": [start, end], "stock": [STOCK_CODE],
                   "channelCode": ["listedNotice_disc"],
                   "pageSize": SZSE_PAGE_SIZE, "pageNum": page}
        url = f"{SZSE_BASE}/api/disc/announcement/annList?random={random.random()}"
        data = f.json("POST", url, "szse.cn filing search", json=payload,
                      headers=szse_headers(), gap=f.api_gap)
        if not isinstance(data, dict):
            raise FetchError("szse.cn filing search: unexpected reply")
        batch = data.get("data") or []
        total = data.get("announceCount")
        new = {str(item.get("id")) for item in batch} - ids
        ids |= new
        for item in batch:
            codes = item.get("secCode") or [STOCK_CODE]
            codes = [codes] if isinstance(codes, str) else codes
            if STOCK_CODE in [str(c) for c in codes] and str(item.get("id")) in new:
                found.append(item)
        done = isinstance(total, int) and page * SZSE_PAGE_SIZE >= total
        if not batch or not new or len(batch) < SZSE_PAGE_SIZE or done or page >= 1000:
            return found
        page += 1


def szse_record(item):
    sid = str(item.get("id") or item.get("annId") or "").strip() or "noid"
    title = clean_title(item.get("title")) or "untitled"
    date = str(item.get("publishTime") or "")[:10] or "unknown-date"
    path = str(item.get("attachPath") or "").strip()
    if path and not path.startswith("/"):
        path = "/" + path
    fmt = str(item.get("attachFormat") or "").strip().lower()
    ext = (os.path.splitext(path)[1] or ("." + fmt if fmt else ".pdf")).lower()
    name = f"{date}_{safe_name(sid, 40)}_{safe_name(title, 120)}{ext}"
    return {
        "key": "szse:" + sid, "source": "szse", "kind": "filing", "date": date,
        "title": title, "url": SZSE_STATIC + path if path else "", "page_url": "",
        "file": Path("filings", "szse", date[:4], name).as_posix(),
    }


def szse_records(f, start, end):
    return [szse_record(item) for item in szse_list(f, start, end)]


def download_filing(f, man, out, rec):
    dest = out / rec["file"]
    if dest.exists() and dest.stat().st_size > 0:
        man.add(dict(rec, status="ok", size=dest.stat().st_size))
        return "existing"
    if not rec["url"]:
        man.add(dict(rec, status="failed: no file attached to this filing", size=0))
        return "failed"
    try:
        resp = f.request("GET", rec["url"], gap=f.file_gap)
        if resp.status_code != 200:
            raise FetchError(f"HTTP {resp.status_code}")
        data = resp.content
        if dest.suffix.lower() == ".pdf" and b"%PDF" not in data[:1024]:
            raise FetchError("the server sent an error page instead of the PDF")
    except FetchError as e:
        man.add(dict(rec, status=f"failed: {plain(e)}", size=0))
        return "failed"
    write_atomic(dest, data)
    man.add(dict(rec, status="ok", size=len(data)))
    return "downloaded"


def save_filings(f, man, out, recs, seen):
    counts = {"downloaded": 0, "existing": 0, "failed": 0, "duplicate": 0, "missing": 0}
    for rec in recs:
        same = (rec["date"], norm_title(rec["title"]))
        if rec["source"] == "szse" and same in seen:
            counts["duplicate"] += 1  # already saved from cninfo
            continue
        seen.add(same)
        before = man.items.get(rec["key"], {})
        if re.match(r"failed: HTTP (404|410)", str(before.get("status", ""))):
            counts["missing"] += 1  # gone from the server; not retried
            continue
        result = download_filing(f, man, out, rec)
        counts[result] += 1
        if result == "downloaded":
            log(f"    {rec['date']}  {rec['title'][:80]}")
        elif result == "failed":
            log(f"    FAILED {rec['date']}  {rec['title'][:60]}  ({man.items[rec['key']]['status']})")
    return counts


def scrape_filings(f, man, out, args):
    log("")
    log("EXCHANGE FILINGS")
    mode = args.filings_from
    org_id = None
    if mode in ("auto", "cninfo", "both"):
        warm_up(f, CNINFO_BASE + "/new/commonUrl/pageOfSearch?url=disclosure/list/search")
        org_id, how = cninfo_org_id(f)
        log(f"  cninfo company id: {org_id} (from {how})")
    szse_ready = False
    strikes, szse_strikes = 0, 0
    tabs = ["fulltext", "relation"]
    seen = set()
    for start, end in year_ranges(args.since, args.until):
        year = start[:4]
        cninfo_worked = False
        if org_id:
            try:
                recs = cninfo_records(f, org_id, start, end, tabs)
                cninfo_worked, strikes = True, 0
                if recs:
                    log(f"  {year}: {len(recs)} filings on cninfo")
                    c = save_filings(f, man, out, recs, seen)
                    log(f"  {year}: {c['downloaded']} new, {c['existing']} already saved, "
                        f"{c['failed']} failed")
            except FetchError as e:
                strikes += 1
                log(f"  {year}: cninfo failed - {plain(e)}")
                if mode == "auto" and strikes >= 2:
                    log("  cninfo keeps failing; using szse.cn for the remaining years")
                    org_id = None
        if szse_strikes < 3 and (mode in ("szse", "both")
                                 or (mode == "auto" and not cninfo_worked)):
            if not szse_ready:
                warm_up(f, SZSE_BASE + "/disclosure/listed/notice/index.html")
                szse_ready = True
            try:
                recs = szse_records(f, start, end)
                szse_strikes = 0
                if recs:
                    log(f"  {year}: {len(recs)} filings on szse.cn")
                    c = save_filings(f, man, out, recs, seen)
                    log(f"  {year}: {c['downloaded']} new, {c['existing']} already saved, "
                        f"{c['duplicate']} duplicates of cninfo filings skipped, "
                        f"{c['failed']} failed")
            except FetchError as e:
                szse_strikes += 1
                log(f"  {year}: szse.cn failed - {plain(e)}")
                if szse_strikes >= 3:
                    log("  szse.cn keeps failing too; it is skipped for the remaining years")
        if not org_id and szse_strikes >= 3:
            log("  Both filing sources are failing - stopping the filings step. "
                "Run the same command again later; finished downloads are kept.")
            break


# ==================================================================== websites
def host_of(url):
    return (urlparse(url).hostname or "").lower()


def base_domain(host):
    host = host.lower().rstrip(".")
    if host == "localhost" or re.fullmatch(r"[\d.]+", host):
        return host
    parts = host.split(".")
    if len(parts) >= 3 and parts[-2] in ("com", "net", "org", "gov", "edu", "co", "ac"):
        return ".".join(parts[-3:])
    return ".".join(parts[-2:])


def tidy_url(url):
    p = urlparse(url)
    path = re.sub(r"/{2,}", "/", p.path or "/")
    return urlunparse((p.scheme.lower(), p.netloc, path, p.params, p.query, ""))


def canonical(url):
    """One key per page, so /en//news and www./non-www. copies are fetched once."""
    p = urlparse(url)
    host = (p.hostname or "").lower()
    if host.startswith("www."):
        host = host[4:]
    try:
        port = f":{p.port}" if p.port and p.port not in (80, 443) else ""
    except ValueError:
        port = ""
    path = re.sub(r"/{2,}", "/", unquote(p.path or "/"))
    if len(path) > 1:
        path = path.rstrip("/")
    query = urlencode(sorted((k, v) for k, v in parse_qsl(p.query, keep_blank_values=True)
                             if not k.lower().startswith("utm_")))
    return host + port + path + ("?" + query if query else "")


def site_folder(url):
    p = urlparse(url)
    host = (p.hostname or "site").lower()
    if host.startswith("www."):
        host = host[4:]
    try:
        if p.port and p.port not in (80, 443):
            host += f"_{p.port}"
    except ValueError:
        pass
    return safe_name(host, 80)


def url_slug(url):
    p = urlparse(url)
    slug = unquote(p.path).strip("/").replace("/", "_") or "index"
    if p.query:
        slug += "__" + unquote(p.query)
    return safe_name(slug, 100)


def url_hash(url):
    return hashlib.sha1(canonical(url).encode("utf-8")).hexdigest()[:8]


def declared_charset(resp):
    found = re.search(r"charset=([\w-]+)", resp.headers.get("Content-Type", ""), re.I)
    return found.group(1) if found else None


SEARCH_URL = re.compile(r"/search|[?&](q|wd|kw|key|keyword|keywords|search)=", re.I)
DOC_IN_HTML = re.compile(
    r"""["'(=]\s*([^"'()<>\s]+?\.(?:pdf|docx?|xlsx?|pptx?|wps|zip|rar|7z))"""
    r"""(?:[?#][^"'()<>\s]*)?\s*["')]""", re.I)
URL_IN_SCRIPT = re.compile(
    r"""['"]([^'"\s]+\.(?:html?|aspx?|php|jsp|pdf|docx?|xlsx?|pptx?)(?:\?[^'"\s]*)?)['"]""", re.I)


def disposition_name(headers):
    disposition = headers.get("Content-Disposition") or ""
    named = re.search(r"""filename\*?=(?:UTF-8'')?["']?([^"';]+)""", disposition, re.I)
    return unquote(named.group(1)).strip() if named else ""


def doc_extension(url, content_type, headers):
    """The document type of a response, or None if it is a web page or something else."""
    for name in (disposition_name(headers), urlparse(url).path):
        ext = os.path.splitext(name)[1].lower()
        if ext in DOC_EXTENSIONS:
            return ext
    return DOC_CONTENT_TYPES.get(content_type)


SNIFF_TYPES = ("", "application/octet-stream", "binary/octet-stream",
               "application/force-download", "application/x-download", "application/download")


def sniff(data):
    """Guess the type of a download whose server did not say what it is."""
    if b"%PDF" in data[:1024]:
        return ".pdf"
    if data[:4] == b"PK\x03\x04":
        try:
            names = zipfile.ZipFile(io.BytesIO(data)).namelist()
        except zipfile.BadZipFile:
            return None
        for prefix, ext in (("word/", ".docx"), ("xl/", ".xlsx"), ("ppt/", ".pptx")):
            if any(n.startswith(prefix) for n in names):
                return ext
        return ".zip"
    head = data[:1024].lstrip().lower()
    if head.startswith(b"<!doctype html") or b"<html" in head:
        return "html"
    return None


def extract_links(soup, raw_html, page_url):
    base = page_url
    base_tag = soup.find("base", href=True)
    if base_tag:
        base = urljoin(page_url, base_tag["href"])
    found = []
    for tag, attr in (("a", "href"), ("area", "href"), ("iframe", "src"),
                      ("frame", "src"), ("embed", "src"), ("object", "data")):
        for node in soup.find_all(tag):
            value = node.get(attr)
            if value:
                label = node.get_text(" ", strip=True) if tag == "a" else ""
                found.append((value, label or node.get("title") or ""))
    for attr in ("data-href", "data-url", "data-link"):
        for node in soup.find_all(attrs={attr: True}):
            found.append((node[attr], node.get_text(" ", strip=True)))
    for node in soup.find_all(onclick=True):
        for value in URL_IN_SCRIPT.findall(node["onclick"]):
            found.append((value, node.get_text(" ", strip=True)))
    for value in DOC_IN_HTML.findall(raw_html):
        found.append((value, ""))
    links = []
    for value, label in found:
        value = html.unescape(str(value)).strip()
        if not value or value.startswith(("#", "javascript:", "mailto:", "tel:", "data:")):
            continue
        links.append((tidy_url(urljoin(base, value)), " ".join(label.split())[:200]))
    return links


class Crawler:
    def __init__(self, fetcher, manifest, out, roots, max_pages=0, save_html=False,
                 fresh=False):
        self.f = fetcher
        self.man = manifest
        self.out = out
        self.roots = [tidy_url(u) for u in roots]
        self.domains = sorted({base_domain(host_of(u)) for u in self.roots})
        self.max_pages = max_pages
        self.save_html = save_html
        self.state_path = out / "website" / "crawl_state.json"
        self.queue = deque()
        self.deferred = []          # pages on sites that are down right now
        self.seen = set()
        self.variants = {}
        self.labels = {}            # document key -> [link text, page it was linked from]
        self.robots = {}
        self.robot_sitemaps = {}
        self.host_gap = {}
        self.failures_in_row = {}
        self.paused_hosts = set()
        self.http_only = set()
        self.warned_challenge = set()
        self.stats = {"pages": 0, "files": 0, "errors": 0, "robots": 0}
        self.resumed = False
        if not fresh:
            self._load_state()

    # ---------------------------------------------------------- crawl state
    def _load_state(self):
        try:
            state = json.loads(self.state_path.read_text(encoding="utf-8"))
        except (OSError, ValueError):
            return
        if state.get("queue") and state.get("domains") == self.domains:
            self.queue = deque(state["queue"])
            self.seen = set(state.get("seen") or [])
            self.variants = state.get("variants") or {}
            self.labels = state.get("labels") or {}
            self.resumed = True

    def _save_state(self):
        state = {"domains": self.domains, "queue": list(self.queue) + self.deferred,
                 "seen": sorted(self.seen), "variants": self.variants, "labels": self.labels}
        data = json.dumps(state, ensure_ascii=False).encode("utf-8")
        write_atomic(self.state_path, data)

    # ---------------------------------------------------------------- scope
    def in_scope(self, url):
        host = host_of(url)
        return any(host == d or host.endswith("." + d) for d in self.domains)

    def enqueue(self, url, label="", referrer=""):
        p = urlparse(url)
        if p.scheme not in ("http", "https") or not self.in_scope(url):
            return False
        ext = os.path.splitext(p.path.lower())[1]
        if ext in SKIP_EXTENSIONS or SEARCH_URL.search(url):
            return False
        key = canonical(url)
        if key in self.seen:
            return False
        if p.query:
            stem = key.split("?", 1)[0]
            count = self.variants.get(stem, 0)
            if count >= MAX_QUERY_VARIANTS:
                return False
            self.variants[stem] = count + 1
        self.seen.add(key)
        if ext in DOC_EXTENSIONS:
            self.labels[key] = [label, referrer]
        self.queue.append(url)
        return True

    # --------------------------------------------------- robots and sitemaps
    def robots_for(self, origin):
        if origin in self.robots:
            return self.robots[origin]
        parser = None
        try:
            resp = self.f.request("GET", origin + "/robots.txt", retries=1)
            ctype = resp.headers.get("Content-Type", "").lower()
            text = resp.text if resp.status_code == 200 and "html" not in ctype else ""
            resp.close()
            if text.strip():
                parser = robotparser.RobotFileParser()
                parser.parse(text.splitlines())
                self.robot_sitemaps[origin] = re.findall(r"(?im)^\s*sitemap:\s*(\S+)", text)
                delay = parser.crawl_delay(ROBOTS_AGENT)
                if delay:
                    self.host_gap[origin] = max(self.f.web_gap, min(float(delay), 30.0))
        except FetchError:
            pass
        self.robots[origin] = parser
        return parser

    def sitemap_urls(self, origin):
        self.robots_for(origin)
        todo = deque(self.robot_sitemaps.get(origin) or [])
        todo.append(origin + "/sitemap.xml")
        checked, urls = set(), []
        while todo and len(checked) < 200:
            sitemap = todo.popleft()
            if sitemap in checked:
                continue
            checked.add(sitemap)
            try:
                resp = self.f.request("GET", sitemap, retries=1)
            except FetchError:
                continue
            body, status = resp.content, resp.status_code
            resp.close()
            if status != 200:
                continue
            if body[:2] == b"\x1f\x8b":
                try:
                    body = gzip.decompress(body)
                except OSError:
                    continue
            text = body.decode("utf-8", "ignore")
            locs = [html.unescape(x).strip()
                    for x in re.findall(r"<loc>\s*(.*?)\s*</loc>", text, re.S | re.I)]
            if re.search(r"<sitemapindex", text, re.I):
                todo.extend(locs)
            elif re.search(r"<urlset", text, re.I):
                urls.extend(locs)
        if urls:
            log(f"  {origin}: sitemap lists {len(urls)} pages")
        return urls

    # ---------------------------------------------------------------- main
    def run(self):
        log("")
        log("WEBSITES")
        log("  Sites: " + ", ".join(self.domains))
        if self.resumed:
            log(f"  Resuming the earlier crawl: {len(self.queue)} pages still to visit")
        else:
            for root in self.roots:
                self.enqueue(root)
            origins = []
            for root in self.roots:
                p = urlparse(root)
                origin = f"{p.scheme}://{p.netloc}"
                if origin not in origins:
                    origins.append(origin)
            for origin in origins:
                for url in self.sitemap_urls(origin):
                    self.enqueue(tidy_url(url))
        fetched, last_save = 0, time.time()
        try:
            while self.queue:
                if self.max_pages and fetched >= self.max_pages:
                    log(f"  Stopped after --max-pages {self.max_pages}; "
                        "run again to continue the crawl.")
                    break
                url = self.queue.popleft()
                try:
                    if self.process(url):
                        fetched += 1
                except KeyboardInterrupt:
                    self.queue.appendleft(url)
                    raise
                if time.time() - last_save > STATE_SAVE_SECONDS:
                    self._save_state()
                    last_save = time.time()
        finally:
            self._save_state()
        left = len(self.queue) + len(self.deferred)
        status = "complete" if not left else f"{left} pages saved for the next run"
        log(f"  Website crawl: {self.stats['pages']} pages, {self.stats['files']} documents, "
            f"{self.stats['errors']} errors ({status})")

    def page_rel(self, url):
        return f"website/{site_folder(url)}/pages/{url_slug(url)}_{url_hash(url)}.txt"

    def file_rel(self, url, ext):
        stem = url_slug(url)
        if stem.lower().endswith(ext):
            stem = stem[: -len(ext)]
        return f"website/{site_folder(url)}/files/{stem}_{url_hash(url)}{ext}"

    def _failure(self, host, url, reason, serious=True):
        self.stats["errors"] += 1
        log(f"    failed: {url} ({reason})")
        if not serious:
            return
        count = self.failures_in_row.get(host, 0) + 1
        self.failures_in_row[host] = count
        if count >= HOST_FAILURE_LIMIT and host not in self.paused_hosts:
            self.paused_hosts.add(host)
            log(f"  {host}: {count} failures in a row - skipping this site until the next run")

    def process(self, url):
        """Fetch one queued URL. Returns True if the network was used."""
        p = urlparse(url)
        host = (p.hostname or "").lower()
        if host in self.paused_hosts:
            self.deferred.append(url)
            return False
        if host in self.http_only and p.scheme == "https":
            url = urlunparse(("http",) + tuple(p[1:]))
            p = urlparse(url)
        origin = f"{p.scheme}://{p.netloc}"
        robots = self.robots_for(origin)
        if robots is not None and not robots.can_fetch(ROBOTS_AGENT, url):
            self.stats["robots"] += 1
            return False
        key = canonical(url)
        path_ext = os.path.splitext(p.path.lower())[1]
        if path_ext in DOC_EXTENSIONS:
            rec = self.file_record(url, key, path_ext)
            dest = self.out / rec["file"]
            if dest.exists() and dest.stat().st_size > 0:
                self.man.add(dict(rec, status="ok", size=dest.stat().st_size))
                return False  # downloaded on an earlier run
        current = url
        for _hop in range(6):  # follow redirects here, so each page is fetched once
            cp = urlparse(current)
            corigin = f"{cp.scheme}://{cp.netloc}"
            try:
                resp = self.f.request("GET", current, gap=self.host_gap.get(corigin),
                                      retries=1, stream=True, allow_redirects=False)
            except FetchError as e:
                if "SSL" in str(e) and cp.scheme == "https" and host not in self.http_only:
                    self.http_only.add(host)
                    log(f"  {host}: HTTPS certificate problem, switching to plain HTTP")
                    self.queue.appendleft(url)
                    return True
                self._failure(host, url, plain(e))
                return True
            location = resp.headers.get("Location")
            if resp.status_code not in (301, 302, 303, 307, 308) or not location:
                break
            resp.close()
            target = tidy_url(urljoin(current, location))
            if urlparse(target).scheme not in ("http", "https") or not self.in_scope(target):
                return True  # redirected to another website
            target_key = canonical(target)
            if target_key != canonical(current):
                if target_key in self.seen:
                    return True  # redirected to a page that is fetched separately
                self.seen.add(target_key)
            tp = urlparse(target)
            target_robots = self.robots_for(f"{tp.scheme}://{tp.netloc}")
            if target_robots is not None and not target_robots.can_fetch(ROBOTS_AGENT, target):
                self.stats["robots"] += 1
                return True
            current = target
        else:
            self._failure(host, url, "too many redirects", serious=False)
            return True
        try:
            final = current
            if resp.status_code != 200:
                serious = resp.status_code in (401, 403) or resp.status_code >= 500
                self._failure(host, url, f"HTTP {resp.status_code}", serious)
                return True
            self.failures_in_row[host] = 0
            ctype = (resp.headers.get("Content-Type") or "").split(";")[0].strip().lower()
            ext = doc_extension(final, ctype, resp.headers)
            try:
                if ext:
                    self.store_file(url, key, read_capped(resp, MAX_FILE_MB), ext,
                                    resp.headers)
                elif ctype in HTML_TYPES[:2]:
                    self.save_page(url, key, resp, read_capped(resp, MAX_PAGE_MB))
                elif ctype in SNIFF_TYPES:
                    data = read_capped(resp, MAX_FILE_MB)
                    kind = sniff(data)
                    if kind == "html":
                        self.save_page(url, key, resp, data)
                    elif kind:
                        self.store_file(url, key, data, kind, resp.headers)
                # anything else (images, video, scripts) is ignored
            except FetchError as e:
                if ext:
                    self.record_failed_file(url, key, ext, str(e))
                self._failure(host, url, str(e), serious=False)
        finally:
            resp.close()
        return True

    def save_page(self, url, key, resp, body):
        soup = BeautifulSoup(body, "html.parser", from_encoding=declared_charset(resp))
        try:
            raw = body.decode(soup.original_encoding or "utf-8", "ignore")
        except LookupError:
            raw = body.decode("utf-8", "ignore")
        title = clean_title(soup.title.get_text()) if soup.title else ""
        links = extract_links(soup, raw, url)
        for node in soup(["script", "style", "noscript", "svg", "template"]):
            node.decompose()
        lines = (" ".join(line.split()) for line in soup.get_text("\n").splitlines())
        text = "\n".join(line for line in lines if line)
        rel = self.page_rel(url)
        header = (f"URL: {url}\nTitle: {title}\n"
                  f"Saved: {dt.datetime.now().isoformat(timespec='seconds')}\n\n")
        write_atomic(self.out / rel, (header + text).encode("utf-8"))
        if self.save_html:
            write_atomic(self.out / (rel[:-4] + ".html"), body)
        self.man.add({"key": "web:" + key, "source": "website", "kind": "page", "date": "",
                      "title": title or url, "url": url, "page_url": "", "file": rel,
                      "status": "ok", "size": len(text)})
        self.stats["pages"] += 1
        log(f"  page {self.stats['pages']:>5}  {url}")
        for link, label in links:
            self.enqueue(link, label, url)
        host = host_of(url)
        if not links and host not in self.warned_challenge and any(
                hint in raw.lower() for hint in BOT_CHALLENGE_HINTS):
            self.warned_challenge.add(host)
            log(f"  WARNING: {host} returned a bot-check page instead of content; "
                "pages from this site may be missing")

    def file_record(self, url, key, ext, headers=None):
        label, referrer = self.labels.get(key, ["", ""])
        title = (label or disposition_name(headers or {})
                 or unquote(os.path.basename(urlparse(url).path)) or url)
        return {"key": "web:" + key, "source": "website", "kind": "file", "date": "",
                "title": title, "url": url, "page_url": referrer,
                "file": self.file_rel(url, ext)}

    def record_failed_file(self, url, key, ext, reason):
        self.man.add(dict(self.file_record(url, key, ext), status=f"failed: {reason}", size=0))

    def store_file(self, url, key, data, ext, headers):
        rec = self.file_record(url, key, ext, headers)
        dest = self.out / rec["file"]
        if dest.exists() and dest.stat().st_size > 0:
            self.man.add(dict(rec, status="ok", size=dest.stat().st_size))
            return
        if ext == ".pdf" and b"%PDF" not in data[:1024]:
            raise FetchError("not a valid PDF (the server sent something else)")
        write_atomic(dest, data)
        self.man.add(dict(rec, status="ok", size=len(data)))
        self.stats["files"] += 1
        log(f"  file {self.stats['files']:>5}  {rec['title'][:60]}  {url}")


# ============================================================ text & keywords
def ooxml_text(path):
    parts = []
    with zipfile.ZipFile(path) as archive:
        for name in sorted(archive.namelist()):
            if re.match(r"(word/(document|header\d*|footer\d*|footnotes)|xl/sharedStrings"
                        r"|ppt/slides/slide\d+)\.xml$", name):
                xml = archive.read(name).decode("utf-8", "ignore")
                xml = re.sub(r"</(w:p|a:p|si)>", "\n", xml)
                parts.append(html.unescape(re.sub(r"<[^>]+>", "", xml)))
    return "\n".join(parts)


def extract_text(path):
    ext = path.suffix.lower()
    try:
        if ext == ".pdf":
            if PdfReader is None:
                return ""
            reader = PdfReader(str(path))
            if reader.is_encrypted:
                reader.decrypt("")
            pages = []
            for page in reader.pages:
                try:
                    pages.append(page.extract_text() or "")
                except Exception:
                    pages.append("")
            return "\n".join(pages)
        if ext in (".docx", ".xlsx", ".pptx"):
            return ooxml_text(path)
        if ext in (".html", ".htm"):
            return BeautifulSoup(path.read_bytes(), "html.parser").get_text("\n")
        if ext == ".csv":
            return path.read_text(encoding="utf-8", errors="ignore")
    except Exception as e:
        log(f"    could not read the text of {path.name}: {short_error(e)}")
    return ""


def sidecar(path):
    return path.with_name(path.name + ".txt")


def needs_extraction(out, rec):
    if rec.get("kind") == "page" or not is_ok(rec) or not rec.get("file"):
        return False
    path = out / rec["file"]
    return (path.suffix.lower() in TEXT_EXTRACTABLE and path.exists()
            and not sidecar(path).exists())


def item_text(out, rec, extract):
    if not is_ok(rec) or not rec.get("file"):
        return ""
    path = out / rec["file"]
    if not path.exists():
        return ""
    if rec.get("kind") == "page":
        return path.read_text(encoding="utf-8", errors="ignore")
    side = sidecar(path)
    if side.exists():
        return side.read_text(encoding="utf-8", errors="ignore")
    if not extract or path.suffix.lower() not in TEXT_EXTRACTABLE:
        return ""
    text = extract_text(path)
    if text.strip():
        side.write_text(text, encoding="utf-8")
    return text


def find_keywords(text, keywords):
    """[(keyword, count, snippets)], ignoring spaces, line breaks and case."""
    kept, where = [], []
    for i, ch in enumerate(text):
        if not ch.isspace():
            kept.append(ch.lower())
            where.append(i)
    squeezed = "".join(kept)
    results = []
    for keyword in keywords:
        needle = "".join(keyword.split()).lower()
        if not needle:
            continue
        hits, start = [], 0
        while True:
            pos = squeezed.find(needle, start)
            if pos < 0:
                break
            hits.append(pos)
            start = pos + len(needle)
        if not hits:
            continue
        snippets = []
        for pos in hits[:3]:
            a, b = where[pos], where[pos + len(needle) - 1] + 1
            snippets.append("…" + " ".join(text[max(0, a - 80): b + 80].split()) + "…")
        results.append((keyword, len(hits), snippets))
    return results


INDEX_FIELDS = ["source", "kind", "date", "title", "keywords", "url", "page_url", "file",
                "status", "size_kb", "text_chars"]
HIT_FIELDS = ["keyword", "count", "source", "kind", "date", "title", "snippets", "url",
              "file"]


def write_csv(path, fields, rows):
    try:
        fh = open(path, "w", newline="", encoding="utf-8-sig")
    except PermissionError:  # e.g. the file is open in Excel
        path = path.with_name(f"{path.stem}_{dt.datetime.now():%Y%m%d_%H%M%S}{path.suffix}")
        log(f"  The report was open in another program; wrote {path.name} instead")
        fh = open(path, "w", newline="", encoding="utf-8-sig")
    with fh:
        writer = csv.DictWriter(fh, fieldnames=fields, extrasaction="ignore")
        writer.writeheader()
        writer.writerows(rows)
    return path


def build_reports(out, man, keywords, extract=True):
    log("")
    log("REPORTS")
    recs = list(man.items.values())
    filings = sorted((r for r in recs if r.get("source") != "website"),
                     key=lambda r: (r.get("date") or "", r.get("file") or ""), reverse=True)
    web = sorted((r for r in recs if r.get("source") == "website"),
                 key=lambda r: (r.get("kind") or "", r.get("file") or r.get("url") or ""))
    todo = {r["key"] for r in filings + web if extract and needs_extraction(out, r)}
    if todo:
        if PdfReader is None:
            log("  pypdf is not installed, so PDF text is skipped "
                "(pip install pypdf, then run with --scan-only)")
        log(f"  Extracting text from {len(todo)} documents - this can take a while")
    rows, hits, done = [], [], 0
    for rec in filings + web:
        text = item_text(out, rec, extract)
        if rec["key"] in todo:
            done += 1
            if done % 25 == 0 or done == len(todo):
                log(f"    text extracted: {done}/{len(todo)}")
        found = find_keywords(text + "\n" + rec.get("title", ""), keywords)
        rows.append(dict(rec, keywords="; ".join(f"{k} ({n})" for k, n, _ in found),
                         size_kb=round((rec.get("size") or 0) / 1024, 1) if rec.get(
                             "kind") != "page" else "",
                         text_chars=len(text)))
        for keyword, count, snippets in found:
            hits.append(dict(rec, keyword=keyword, count=count, snippets=" | ".join(snippets)))
    index_path = write_csv(out / "index.csv", INDEX_FIELDS, rows)
    hits_path = write_csv(out / "keyword_hits.csv", HIT_FIELDS, hits)
    matched = len({h["key"] for h in hits})
    log(f"  {index_path.name}: {len(rows)} items")
    log(f"  {hits_path.name}: {len(hits)} keyword matches in {matched} items")


def summarize(out, man):
    recs = list(man.items.values())
    filings = [r for r in recs if r.get("source") != "website"]
    pages = [r for r in recs if r.get("kind") == "page"]
    files = [r for r in recs if r.get("kind") == "file"]
    failed = [r for r in recs if not is_ok(r)]
    log("")
    log("SUMMARY")
    log(f"  Filings:        {sum(map(is_ok, filings))} saved, "
        f"{len(filings) - sum(map(is_ok, filings))} failed")
    log(f"  Website pages:  {len(pages)}")
    log(f"  Website files:  {sum(map(is_ok, files))} saved, "
        f"{len(files) - sum(map(is_ok, files))} failed")
    log(f"  Output folder:  {out}")
    log("  Start with keyword_hits.csv: every document that mentions a keyword.")
    if failed:
        log("  Failed items are retried automatically when you run the same command again.")


# ================================================================ quick check
def run_check(f, args):
    log("Testing each source with one or two requests...")
    f.max_retries = 1  # a failing source should fail fast here
    levels = []

    def report(level, name, detail):
        levels.append(level)
        log(f"  [{level}] {name}: {plain(detail)}")

    today = dt.date.today()
    start, end = dt.date(today.year - 1, 1, 1).isoformat(), today.isoformat()
    filings_ok = False

    warm_up(f, CNINFO_BASE + "/new/commonUrl/pageOfSearch?url=disclosure/list/search")
    org_id, how = cninfo_org_id(f)
    report("PASS" if how != DEFAULT_ORG_SOURCE else "WARN", "cninfo company lookup",
           f"id {org_id} (from {how})")
    try:
        recs = [cninfo_record(i, "fulltext", org_id)
                for i in cninfo_list(f, org_id, "fulltext", start, end)]
        if recs:
            recs.sort(key=lambda r: r["date"], reverse=True)
            report("PASS", "cninfo filing search",
                   f"{len(recs)} filings since {start}; newest: {recs[0]['date']} "
                   f"{recs[0]['title'][:50]}")
            sample = next((r for r in recs if r["url"]), None)
            if sample:
                resp = f.request("GET", sample["url"], gap=f.file_gap)
                good = resp.status_code == 200 and b"%PDF" in resp.content[:1024]
                report("PASS" if good else "FAIL", "cninfo PDF download",
                       f"{len(resp.content) // 1024} KB" if good else f"HTTP {resp.status_code}")
                filings_ok = filings_ok or good
        else:
            report("FAIL", "cninfo filing search", "no filings found")
    except FetchError as e:
        report("FAIL", "cninfo filing search", str(e))
    try:
        records = cninfo_list(f, org_id, "relation", start, end)
        report("PASS", "cninfo investor-relations records", f"{len(records)} since {start}")
    except FetchError as e:
        report("WARN", "cninfo investor-relations records", str(e))

    warm_up(f, SZSE_BASE + "/disclosure/listed/notice/index.html")
    try:
        recs = szse_records(f, start, end)
        if recs:
            sample = next((r for r in recs if r["url"]), recs[0])
            resp = f.request("GET", sample["url"], gap=f.file_gap) if sample["url"] else None
            good = resp is not None and resp.status_code == 200 and b"%PDF" in resp.content[:1024]
            report("PASS" if good else "WARN", "szse.cn backup source",
                   f"{len(recs)} filings since {start}; PDF download "
                   + ("works" if good else "failed"))
            filings_ok = filings_ok or good
        else:
            report("WARN", "szse.cn backup source", "no filings found")
    except FetchError as e:
        report("WARN", "szse.cn backup source", str(e))

    site_ok = False
    crawler = Crawler(f, None, Path("."), MAIN_SITE_ROOTS, fresh=True)
    for root in MAIN_SITE_ROOTS[:2]:
        p = urlparse(root)
        robots = crawler.robots_for(f"{p.scheme}://{p.netloc}")
        if robots is not None and not robots.can_fetch(ROBOTS_AGENT, root):
            report("FAIL", root, "robots.txt does not allow crawling this page")
            continue
        try:
            resp = f.request("GET", root, retries=1)
            body = resp.content
            soup = BeautifulSoup(body, "html.parser", from_encoding=declared_charset(resp))
            raw = body.decode(soup.original_encoding or "utf-8", "ignore")
            links = {u for u, _ in extract_links(soup, raw, resp.url) if crawler.in_scope(u)}
            if resp.status_code == 200 and links:
                report("PASS", root, f"page loads, {len(links)} links to follow")
                site_ok = True
            else:
                hint = (" (the site returned a bot-check page)"
                        if any(h in raw.lower() for h in BOT_CHALLENGE_HINTS) else "")
                report("FAIL", root, f"HTTP {resp.status_code}, no links found{hint}")
        except (FetchError, LookupError) as e:
            report("FAIL", root, str(e))

    if not args.main_site_only:
        reachable = 0
        for root in GROUP_SITE_ROOTS:
            try:
                resp = f.request("GET", root, retries=0, timeout=(10, 30))
                reachable += resp.status_code == 200
                resp.close()
            except FetchError:
                pass
        report("PASS" if reachable else "WARN", "Jereh group sites",
               f"{reachable} of {len(GROUP_SITE_ROOTS)} reachable")

    log("")
    if filings_ok and site_ok:
        log("Ready. Start the full scrape with:  python jereh_scraper.py")
        return 0
    if filings_ok:
        log("Filings work but the website check failed; you can still run "
            "python jereh_scraper.py --source filings")
    elif site_ok:
        log("The website works but filings failed; you can still run "
            "python jereh_scraper.py --source website")
    else:
        log("Neither filings nor the website could be reached from this computer. "
            "Check your internet connection or proxy and try again.")
    return 1


# ======================================================================= main
def parse_args(argv):
    p = argparse.ArgumentParser(
        description="Download every public Jereh (002353) filing, web page and document.")
    p.add_argument("--check", action="store_true", help="quickly test every source and exit")
    p.add_argument("--source", choices=["all", "filings", "website"], default="all")
    p.add_argument("--filings-from", choices=["auto", "cninfo", "szse", "both"],
                   default="auto",
                   help="auto: cninfo, with szse.cn as backup for any year cninfo fails")
    p.add_argument("--since", default="2009-01-01",
                   help="first filing date, YYYY-MM-DD (default covers Jereh's 2010 IPO)")
    p.add_argument("--until", default=dt.date.today().isoformat(),
                   help="last filing date, YYYY-MM-DD")
    p.add_argument("--main-site-only", action="store_true",
                   help="crawl jereh.com only, not the group sites")
    p.add_argument("--max-pages", type=int, default=0,
                   help="stop the website crawl after this many pages (0 = no limit)")
    p.add_argument("--delay", type=float, default=1.0,
                   help="seconds between requests to the same site (default 1.0)")
    p.add_argument("--keywords", default="", help="extra keywords, comma-separated")
    p.add_argument("--out", default="jereh_docs", help="output folder")
    p.add_argument("--scan-only", action="store_true",
                   help="rebuild text and reports from what is already downloaded")
    p.add_argument("--no-text", action="store_true",
                   help="skip text extraction (reports use titles and page text only)")
    p.add_argument("--save-html", action="store_true", help="also save raw HTML of each page")
    p.add_argument("--fresh", action="store_true", help="restart the website crawl from scratch")
    args = p.parse_args(argv)
    for name in ("since", "until"):
        try:
            dt.date.fromisoformat(getattr(args, name))
        except ValueError:
            p.error(f"--{name} must look like 2025-01-31")
    if args.since > args.until:
        p.error("--since must be on or before --until")
    args.max_pages = max(0, args.max_pages)
    return args


def main(argv=None):
    global _log_file
    args = parse_args(argv)
    try:
        sys.stdout.reconfigure(errors="replace")
    except (AttributeError, ValueError):
        pass
    out = Path(args.out).expanduser().resolve()
    out.mkdir(parents=True, exist_ok=True)
    _log_file = open(out / "scrape_log.txt", "a", encoding="utf-8")
    log("")
    log(f"=== Jereh scraper started {dt.datetime.now():%Y-%m-%d %H:%M} "
        f"({' '.join(argv if argv is not None else sys.argv[1:]) or 'no options'}) ===")
    fetcher = Fetcher(args.delay)
    try:
        if args.check:
            return run_check(fetcher, args)
        keywords = KEYWORDS + [k.strip() for k in args.keywords.split(",") if k.strip()]
        man = Manifest(out)
        try:
            if not args.scan_only:
                if args.source in ("all", "filings"):
                    scrape_filings(fetcher, man, out, args)
                if args.source in ("all", "website"):
                    roots = MAIN_SITE_ROOTS + ([] if args.main_site_only else GROUP_SITE_ROOTS)
                    Crawler(fetcher, man, out, roots, args.max_pages, args.save_html,
                            args.fresh).run()
            build_reports(out, man, keywords, extract=not args.no_text)
            summarize(out, man)
        except KeyboardInterrupt:
            log("")
            log("Stopped. Progress is saved - run the same command again to continue.")
            return 130
        finally:
            man.close()
        return 0
    finally:
        _log_file.close()
        _log_file = None


if __name__ == "__main__":
    sys.exit(main())
