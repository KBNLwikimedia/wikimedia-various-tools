"""
Wikimedia Commons File Downloader (categories & list modes)

*A robust, Windows-safe downloader for Wikimedia Commons files.*

This script downloads media files from Wikimedia Commons in two modes and keeps an
incremental Excel log of what was fetched (or would be fetched in dry-run).

Modes
-----
1) MODE="categories"
   • Harvests files from one or more root Commons categories, traversing subcategories
     up to DEPTH (BFS).
   • Computes two sets of counts for a clear plan/preview:
       - RAW candidates  : Commons’ *direct* membership counts per bucket (no de-dup).
       - UNIQUE eligible : files actually eligible after extension filtering and
         de-duplication across the whole traversal (what the script will download).
   • Prints a nested, alphabetically sorted tree with per-bucket:
       direct/total RAW vs. direct/total UNIQUE, plus how many were skipped here as
       duplicates found elsewhere in the tree.
   • Supports a single *global* 1-based slice applied AFTER building the UNIQUE pool
     across all selected roots (CATEGORIES_GLOBAL_RANGE_START/END). The order is
     deterministic and configurable (by root-then-title, or by title globally).
   • Downloads to: LOCAL_BASE_FOLDER / CATEGORIES_DOWNLOAD_SUBFOLDER / <Root> /
     (<flattened path>), where
       - If FLATTEN_CATEGORIES_PATHS=True: only the <Root> folder is used on disk to
         avoid very deep Windows paths; the full category chain is still recorded in the log.

2) MODE="list"
   • Reads a separate list file (Excel/CSV/TSV/TXT) that contains exactly ONE column
     named LIST_SINGLE_COLUMN_NAME (default: "CommonsInput"). An optional
     LIST_SOURCE_COLUMN may be present.
   • Each value may be any of:
        - File:Title
        - https://commons.wikimedia.org/wiki/File:Title
        - MediaInfo ID like M12345
        - Commons concept/entity URL (…/entity/M12345 or …/wiki/Special:EntityData/M12345)
        - Direct upload URL (https://upload.wikimedia.org/… including thumb URLs)
   • Every row is resolved to a canonical "File:Title" (+ MID and upload URL) and
     downloaded to: LOCAL_BASE_FOLDER / LIST_DOWNLOAD_SUBFOLDER.

Shared behavior & guarantees
----------------------------
• Windows-safe filenames:
  - Illegal characters replaced; trailing spaces/dots stripped; device names avoided.
  - Local filename ALWAYS ends with a suffix before the extension:
        "--<MID>"   (e.g., "Example--M12345.jpg"), or
        "--NO-MID-<short-hash>"
• Path length safety:
  - Enforces FULL_PATH_BUDGET on the absolute path; trims the stem to fit while preserving
    suffix + extension. Short root/subfolder names are encouraged.
• Extension filtering:
  - FILE_EXTS controls eligibility. Set to tuple() to disable filtering and include all.
• Networking:
  - A single requests.Session with Wikimedia-friendly User-Agent, retry/backoff and timeouts.
• DRY_RUN:
  - When True, nothing is written to disk (no folders, no images, no Excel); the plan is
    still built and printed, and upload URLs are resolved via API where possible.
• Incremental Excel logging (per mode sheet):
  - Log written to EXCEL_LOG_DIR / EXCEL_LOG_NAME (SHEET_NAME_CATEGORIES or SHEET_NAME_LIST).
  - Rows appended in chunks every LOG_FLUSH_ROWS successful downloads.
  - De-duplicates on ("CommonsFileURL", "MediaInfoID") and replaces the sheet in place.
  - Column order:
      CommonsFileName | CommonsFileURL | MediaInfoID | CommonsConceptURI | CommonsImageURL
      | SourceCategory | CommonsCategoryPath | LocalBaseFolder | LocalSubFolder | LocalFilename
• Overwrite policy:
  - If OVERWRITE_EXISTING=False and the exact local filename is present, skip writing the
    file and still log/resolve the canonical upload URL.

Typical workflow
----------------
1) Edit the CONFIG block (MODE, inputs, depth/slice, paths, filters).
2) Run the script; review the printed plan summary (especially in categories mode).
3) Confirm when prompted (unless you disabled CONFIRM_BEFORE_DOWNLOAD).
4) Files are saved and the Excel log is updated incrementally; safe to re-run — the log
   is de-duplicated and existing files can be skipped or overwritten per config.

Counts explained
----------------
• RAW = Commons’ direct membership counts by bucket; matches what you see on the category
  page (per bucket), and totals include descendants in the printed tree.
• UNIQUE = the script’s download set after extension filtering and first-hit de-dup across
  the traversal; this is what will actually be downloaded (and can be globally sliced).

License & contact
=================
Author: ChatGPT, prompted by Olaf Janssen, KB (Koninklijke Bibliotheek)
Latest update: 2025-10-24
License: CC0
"""

#===============================
# 0) Imports & typing
# Collect all stdlib/third-party imports and typing in one place. Keeps dependencies visible and
# avoids circular imports. If you add new libs (e.g., pathlib, time), add them here.
#================================

import os
import re
import hashlib
from urllib.parse import quote, urlparse, unquote
import requests
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry
import pandas as pd
from collections import Counter, deque
from typing import Dict, List, Tuple, Optional


# =========================
# 1) Configuration parameters
# Central place for knobs: MODE, paths, Excel sheet names, retry/timeouts, depth, list input options, DRY_RUN.
# Readers can understand behavior without diving into code.
# Treat as the only place users should edit.
# =========================

# --- Download mode selection ---
# "categories" → harvest/download via Commons categories with DEPTH (subcats)
# "list"       → read a separate file (txt, Excel, csv) containing mixed Commons file identifiers
MODE = "categories"   # Configure your category specific settings below
#MODE = "list"          # Configure your list specific settings below

# --- Shared settings ---
# Allow only files with certain formats/extensions to be downloaded
#FILE_EXTS  = (".jpg", ".jpeg", ".png", ".gif", ".tif", ".tiff", ".webp", ".svg", ".pdf")
FILE_EXTS  = tuple()  # or None # No filtering on file format to be downloaded
UA = "Wikimedia Commons File Downloader by User:OlafJanssen (olaf.janssen@kb.nl)"

# Where downloads are stored on disk. Top-level folder for all downloads (both modes).
# Name kept short on purpose to avoid exhaustion of the FULL_PATH_BUDGET = 250
LOCAL_BASE_FOLDER = "dwnlds"

# Download log Excel file
EXCEL_LOG_NAME = "downloads_log.xlsx"      # Excel admin file name
# Set the directory where EXCEL_LOG_NAME is written:
#   - Absolute path example (Windows): r"D:\KB-OPEN\logs"
#   - Absolute path example (POSIX):   "/var/tmp/kb-logs"
#   - Relative path: "logs" (resolved relative to the *script folder*)
#   - "./" → the *script folder* itself (recommended if you want it next to the script)
#   - None → default: write next to downloads (i.e., inside LOCAL_BASE_FOLDER)
EXCEL_LOG_DIR: str | None = "./"

# Output Excel log file sheet names (per mode)
SHEET_NAME_CATEGORIES = "CategoriesDownloads"   # used when MODE="categories"
SHEET_NAME_LIST       = "ListDownloads"         # used when MODE="list"

# Runtime behavior
OVERWRITE_EXISTING = False       # if False, skip download when the destination file already exists
CONFIRM_BEFORE_DOWNLOAD = True   # prompt after preflight count
LOG_FLUSH_ROWS = 10              # after every N successful downloads, append to Excel log

# Dry-run (no files/logs written). When True:
#   - The script scans and plans (counts images, resolves titles/MIDs/URLs where possible)
#   - It prints what WOULD happen
#   - It does NOT create folders, write image files, or update the Excel log
DRY_RUN = False

# Conservative Windows full-path limit. This is the max number of characters that the local full file path on Windows
# can take, including the local filename of the downloaded image. When that file name is too long, it will be truncated
# dynamically to fit within the overall FULL_PATH_BUDGET number of characters.
FULL_PATH_BUDGET = 250

# Networking / API
COMMONS_API = "https://commons.wikimedia.org/w/api.php"
CATEGORY_PAGE_LIMIT = 500                 # per-page API limit (continuation will fetch the rest)
TIMEOUT_SECS = 20
RETRIES_TOTAL = 5
RETRIES_BACKOFF = 0.6

# ----------------------------
# MODE = "categories" specific settings
# ----------------------------
# Root Wikimedia Commons categories to harvest (WITHOUT 'Category:'; spaces will be converted to underscores)
#CATEGORIES = ["A.W. Nieuwenhuis","Ad Grimmon"] # Multiple simple small categories, formatted as a list []
#CATEGORIES = ["Ad Grimmon"] # Even for single root category, this MUST be formatted as a list []
# -- KB collections categories
#CATEGORIES = ["Magazines from Koninklijke Bibliotheek"] # Simple small nested category tree
#CATEGORIES = ["Collections from Koninklijke Bibliotheek"] # larger nested category tree, to experiment with DEPTH setting
CATEGORIES = ["Media contributed by Koninklijke Bibliotheek"] # Flat dump of all KB contributed media files, DEPTH=0

DEPTH = 0   # 0 = download only images from the main category; 1 = include immediate subcats; 2 = deeper, etc.

# Put all category downloads under a fixed subfolder and FLATTEN subcats on Windows file system
CATEGORIES_DOWNLOAD_SUBFOLDER = "cats" # Name kept short on purpose to avoid exhaustion of the FULL_PATH_BUDGET = 250
#  Ignore subcats in filesystem, otherwise the sub(sub(sub))cat tree on Windows might grow too large for
#  the file system to handle
FLATTEN_CATEGORIES_PATHS = True

# For very large lists of root categories, process only a subset/slice of all categories (1-based, inclusive)
# - What it slices: the list of root categories in CATEGORIES[].
# - When it applies: before any harvesting—i.e., it decides which roots you’ll even visit for harvesting files from.
# - Use it when: you have a long CATEGORIES list and only want, say, roots 2–8 this run.
# - No effect if you have only one root in CATEGORIES[]
CATEGORIES_RANGE_START: int | None = None  # e.g., None = 1, or  None = None for all root categories
CATEGORIES_RANGE_END:   int | None = None  # e.g., None = 4 for first 4 root cats in CATEGORIES[], or  None = None for all root categories

# Global file subset/slice across ALL roots (1-based, inclusive) AFTER unique/eligible pool is built
# - What it slices: the single, combined pool of unique & eligible files built after harvesting all selected roots (deduped, extension-filtered).
# - When it applies: after harvesting and deduping—i.e., it decides which files from the grand pool you’ll actually download.
# - Ordering: controlled by CATEGORIES_GLOBAL_SLICE_ORDER ("root_then_title" or "title").
CATEGORIES_GLOBAL_RANGE_START: int | None = 1  # e.g. None = 1, or  None = None for all files
CATEGORIES_GLOBAL_RANGE_END:   int | None = 5  # e.g. None = 5 to download first 5 files , or  None = None for all files

# Order used when building the global pool for slicing:
#   "root_then_title" → preserve root order from CATEGORIES; sort titles A→Z within each root (deterministic, readable)
#   "title"           → ignore root; sort all titles A→Z globally
CATEGORIES_GLOBAL_SLICE_ORDER = "root_then_title"  # or "title"

# ---------------------
# MODE = "list" specific settings
# ---------------------
## TXT input file
# LIST_INPUT_PATH = "list-of-tobe-downloaded-files.txt"
# LIST_INPUT_FORMAT = "txt"

## Excel input file
LIST_INPUT_PATH = "list-of-tobe-downloaded-files.xlsx"     # can be .xlsx, .csv, .tsv, or .txt
LIST_INPUT_FORMAT = "excel"                 # "auto" | "excel" | "csv" | "tsv" | "txt"
LIST_EXCEL_SHEET = "FilesList"            # name of input Excel sheet - only used when LIST_INPUT_FORMAT="excel"
LIST_TEXT_DELIM: str | None = None        # for csv/tsv/txt; None = auto-sniff (csv), "\t" for TSV

# Single-column input for LIST mode
LIST_SINGLE_COLUMN_NAME = "CommonsInput"
# The list file should have ONE column with this header; each cell may contain:
#   - File:Title
#   - https://commons.wikimedia.org/wiki/File:Title
#   - M123456
#   - https://commons.wikimedia.org/entity/M123456  (or /wiki/Special:EntityData/M123456)
#   - https://upload.wikimedia.org/wikipedia/commons/.../Filename.ext

# Downloads destination for LIST mode (as a subfolder of LOCAL_BASE_FOLDER)
LIST_DOWNLOAD_SUBFOLDER = "list"      # files go under LOCAL_BASE_FOLDER / LIST_DOWNLOAD_SUBFOLDER
# Name kept short om purpose to avoid exhaustion of the FULL_PATH_BUDGET = 250

# Optionally limit how many rows from the list to process (1-based, inclusive)
LIST_RANGE_START: int | None = 1  # e.g., None = 1, or None = None for all files
LIST_RANGE_END:   int | None = 5 # e.g., None = 5 to process first files, or  None = None for all files

# Optional: if your list file includes a source tag column, name it here
LIST_SOURCE_COLUMN = "SourceCategory"  # leave as-is if you use that header name
# When the list file lacks a 'SourceCategory' column, use this tag in the log file
LIST_SOURCE_TAG = "Manual list"

# =========================
# End of configuration
# =========================

#===============================
# 2) Script/FS utilities (very generic helpers)
# Tiny helpers that don’t know about Commons or networking: e.g., locating the script folder, resolving the log
# directory (handling ./, absolute/relative paths). Keeps path logic consistent across OSes.
#================================

def script_dir() -> str:
    """Absolute path to this script's directory (fallback to CWD if __file__ unavailable)."""
    try:
        return os.path.dirname(os.path.abspath(__file__))
    except NameError:
        return os.getcwd()

def resolve_log_dir(base_folder: str) -> str:
    """
    Resolve EXCEL_LOG_DIR:
      - None  → base_folder
      - "./"  → script folder
      - other → absolute if given, else relative to the script folder
    """
    if EXCEL_LOG_DIR is None:
        return os.path.abspath(base_folder)
    if EXCEL_LOG_DIR.strip() == "./":
        return script_dir()
    path = EXCEL_LOG_DIR
    if not os.path.isabs(path):
        path = os.path.join(script_dir(), path)
    return os.path.abspath(os.path.expanduser(path))

def _short_hash(text: str, n: int = 8) -> str:
    return hashlib.blake2s(text.encode("utf-8"), digest_size=8).hexdigest()[:n]

#===============================
# 3) Path & filename safety
# Everything about Windows-safe names and path-length budgets: sanitizing folder segments, building
# safe filenames with --MID/--NO-MID-<hash>, trimming to fit FULL_PATH_BUDGET. Prevents runtime errors on long
# or illegal paths.
#================================

def sanitize_filename(
    filename: str,
    folder: Optional[str] = None,
    max_length: int = 120,
    full_path_budget: Optional[int] = FULL_PATH_BUDGET,
    mid: Optional[str] = None,
) -> str:
    """
    Sanitize a filename for Windows and ALWAYS append a suffix before the extension:
      - '--<MID>' when a MediaInfo ID is available (e.g., '--M12345')
      - '--NO-MID-<short-hash>' when no MID is available

    If 'folder' is provided, ensure the final ABSOLUTE path length stays within 'full_path_budget'
    by truncating the stem as needed (suffix + extension are preserved).
    """
    original = filename

    # 1) Replace Windows-illegal chars & strip trailing dots/spaces
    invalid = ['"', '#', '%', '&', '{', '}', '\\', '<', '>', '|', ':', '*', '?', '/']
    for ch in invalid:
        filename = filename.replace(ch, '_')
    filename = filename.rstrip(' .')

    # 2) Split into stem/ext
    stem, ext = os.path.splitext(filename)

    # 3) Per-name cap FIRST (without suffix yet) to keep things reasonable
    if len(filename) > max_length:
        room_for_stem = max(1, max_length - len(ext))
        stem = stem[:room_for_stem]

    # 4) Avoid reserved device names (on the bare stem)
    reserved = {
        "con","prn","aux","nul",
        *{f"com{i}" for i in range(1,10)},
        *{f"lpt{i}" for i in range(1,10)}
    }
    if stem.strip('. ').lower() in reserved:
        stem = stem + "_"

    # 5) Build the suffix we ALWAYS append
    suffix = f"--{mid.strip()}" if (mid and str(mid).strip()) else f"--NO-MID-{_short_hash(original)}"

    # 6) How many characters we can use for the WHOLE name (stem + suffix + ext)
    if folder:
        try:
            abs_dir = os.path.abspath(folder)
        except Exception:
            abs_dir = folder
        available_for_name = (full_path_budget or FULL_PATH_BUDGET) - len(abs_dir) - 1  # + path sep
        if available_for_name < len(ext) + len(suffix) + 1:
            available_for_name = len(ext) + len(suffix) + 1
    else:
        available_for_name = max_length

    # 7) Trim stem if needed to fit [stem + suffix + ext] within the budget
    name_len = len(stem) + len(suffix) + len(ext)
    if name_len > available_for_name:
        room_for_stem = max(1, available_for_name - len(suffix) - len(ext))
        stem = stem[:room_for_stem]

    # 8) Return final
    return stem + suffix + ext

def sanitize_folder_component(name: str, max_len: int = 80) -> str:
    """Sanitize a single folder segment for Windows; shorten long names with a hash."""
    invalid = ['"', '#', '%', '&', '{', '}', '\\', '<', '>', '|', ':', '*', '?', '/']
    for ch in invalid:
        name = name.replace(ch, '_')
    name = name.rstrip(' .')
    if len(name) > max_len:
        name = name[:max_len - 10] + "__" + _short_hash(name)
    return name or "NA"


#===============================
# 4) HTTP session (transport layer)
# Create a configured requests.Session with UserAgent, retries, backoff, and timeouts. All network calls use this
# session for consistent behavior and Wikimedia-friendly etiquette.
#================================

def build_session() -> requests.Session:
    """Create a requests Session with Wikimedia-friendly UA, retries and backoff."""
    s = requests.Session()
    s.headers.update({
        "User-Agent": UA,
        "Accept": "application/json",
    })
    retry = Retry(
        total=RETRIES_TOTAL,
        connect=RETRIES_TOTAL,
        read=RETRIES_TOTAL,
        status=RETRIES_TOTAL,
        backoff_factor=RETRIES_BACKOFF,
        status_forcelist=(429, 500, 502, 503, 504),
        allowed_methods=frozenset({"GET"}),
        raise_on_status=False,
    )
    s.mount("https://", HTTPAdapter(max_retries=retry))
    return s


#===============================
# 5) Input parsing/normalization (no network)
# String parsers only: normalize File: titles, detect if a URL is /wiki/File:… vs /entity/M…
# vs direct upload.wikimedia.org, extract M… from any string, canonicalize thumb URLs to originals.
# Pure functions you can unit test easily.
#================================

def normalize_title(s: str) -> str:
    s = (s or "").strip()
    if not s:
        return ""
    if s.lower().startswith("file:"):
        return "File:" + s.split(":", 1)[1].strip()
    return "File:" + s

def parse_title_from_commons_url(url: str) -> str:
    try:
        p = urlparse(url)
        if "commons.wikimedia.org" not in p.netloc.lower():
            return ""
        parts = p.path.split("/")
        # Typical form: /wiki/File:Example.jpg
        if len(parts) >= 3 and parts[-2].lower() == "wiki":
            last = unquote(parts[-1])
            return last if last.startswith("File:") else ""
        # Fallback: find '/wiki/<something>'
        if "wiki" in [seg.lower() for seg in parts]:
            idx = [seg.lower() for seg in parts].index("wiki")
            if idx + 1 < len(parts):
                last = unquote(parts[idx + 1])
                return last if last.startswith("File:") else ""
    except Exception:
        pass
    return ""


def extract_mid_from_uri(s: str) -> str:
    """
    From strings like 'M123', 'https://commons.wikimedia.org/entity/M123',
    or 'https://commons.wikimedia.org/wiki/Special:EntityData/M123'
    extract 'M123'
    """
    if not s:
        return ""
    m = re.search(r"(?:^|/|=)(M\d+)(?:$|[^0-9])", s)
    return m.group(1) if m else ""

def is_upload_url(url: str) -> bool:
    try:
        return "upload.wikimedia.org" in urlparse(url).netloc.lower()
    except Exception:
        return False

def canonicalize_upload_url(url: str) -> str:
    """
    If given a thumbnail URL (.../thumb/.../<size>px-Filename.ext), convert to the original image URL.
    Otherwise return as-is.
    """
    try:
        p = urlparse(url)
        if "/thumb/" not in p.path:
            return url
        parts = p.path.split("/")
        # Remove '/thumb' segment
        thumb_idx = parts.index("thumb")
        base_parts = parts[:thumb_idx] + parts[thumb_idx + 1:]
        # Remove the last size-part (e.g., '1234px-Filename.ext'); keep the preceding path to original file
        if base_parts:
            base_parts = base_parts[:-1]
        new_path = "/".join(base_parts)
        return f"{p.scheme}://{p.netloc}{new_path}"
    except Exception:
        return url

#===============================
# 6) Commons API wrappers
# Small, composable calls that do one thing: fetch category members (with continuation), resolve upload URL
# from a title, resolve (title, MID, URL) from a title/MID/upload-URL. Keeps API details in one place and
# simplifies higher layers.
#================================

def fetch_category_members(
    session: requests.Session,
    category_title: str,
    cmtype: str,  # "file" or "subcat"
) -> List[dict]:
    """
    Fetch all members of a Commons category for a given type, following continuation.
    Returns a list of dicts (formatversion=2) with at least 'title' and (for files) 'pageid'.
    """
    params = {
        "action": "query",
        "format": "json",
        "formatversion": "2",
        "list": "categorymembers",
        "cmtitle": f"Category:{category_title}",
        "cmtype": cmtype,
        "cmprop": "ids|title",
        "cmlimit": str(CATEGORY_PAGE_LIMIT),
    }

    out: List[dict] = []
    cont: Optional[dict] = None
    while True:
        if cont:
            params.update(cont)
        resp = session.get(COMMONS_API, params=params, timeout=TIMEOUT_SECS)
        resp.raise_for_status()
        try:
            data = resp.json()
        except ValueError:
            snippet = resp.text[:200].replace("\n", " ")
            ctype = resp.headers.get("content-type", "")
            raise RuntimeError(
                f"Non-JSON response from Commons API (status {resp.status_code}, content-type {ctype}): {snippet}"
            )
        members = (data.get("query") or {}).get("categorymembers") or []
        out.extend(members)
        cont = data.get("continue")
        if not cont:
            break
    return out

def resolve_upload_url_via_api(session: requests.Session, title: str) -> str:
    """
    Return the direct upload URL (https://upload.wikimedia.org/...) for a Commons file.
    'title' must be like 'File:Example.jpg'. Follows redirects.
    """
    params = {
        "action": "query",
        "format": "json",
        "formatversion": "2",
        "prop": "imageinfo",
        "titles": title,
        "iiprop": "url",
        "redirects": "1",
    }
    resp = session.get(COMMONS_API, params=params, timeout=TIMEOUT_SECS)
    resp.raise_for_status()
    try:
        data = resp.json()
    except ValueError:
        return ""
    pages = (data.get("query") or {}).get("pages") or []
    if pages and pages[0].get("imageinfo"):
        return pages[0]["imageinfo"][0].get("url") or ""
    return ""

def title_mid_url_from_title(session: requests.Session, title: str) -> Tuple[str, str, str]:
    """
    Given a 'File:...' title, return (normalized_title, mid, upload_url).
    """
    params = {
        "action": "query",
        "format": "json",
        "formatversion": "2",
        "prop": "imageinfo|info",
        "titles": title,
        "iiprop": "url",
        "redirects": "1",
    }
    r = session.get(COMMONS_API, params=params, timeout=TIMEOUT_SECS)
    r.raise_for_status()
    data = r.json()
    pages = (data.get("query") or {}).get("pages") or []
    if not pages:
        return title, "", ""
    page = pages[0]
    page_title = page.get("title") or title
    pageid = page.get("pageid")
    mid = f"M{pageid}" if pageid else ""
    upload_url = ""
    if page.get("imageinfo"):
        upload_url = page["imageinfo"][0].get("url", "") or ""
    return page_title, mid, upload_url

def title_mid_url_from_mid(session: requests.Session, mid: str) -> Tuple[str, str, str]:
    """
    Given a MediaInfo ID 'M123', return (title, mid, upload_url).
    """
    m = re.search(r"M(\d+)", str(mid))
    if not m:
        return "", "", ""
    pid = m.group(1)
    params = {
        "action": "query",
        "format": "json",
        "formatversion": "2",
        "prop": "imageinfo|info",
        "pageids": pid,
        "iiprop": "url",
        "redirects": "1",
    }
    r = session.get(COMMONS_API, params=params, timeout=TIMEOUT_SECS)
    r.raise_for_status()
    data = r.json()
    pages = (data.get("query") or {}).get("pages") or []
    if not pages:
        return "", "", f"M{pid}"
    page = pages[0]
    title = page.get("title") or ""
    upload_url = ""
    if page.get("imageinfo"):
        upload_url = page["imageinfo"][0].get("url", "") or ""
    return title, f"M{pid}", upload_url


def title_mid_url_from_upload_url(session: requests.Session, upload_url: str) -> Tuple[str, str, str]:
    """
    From an upload URL (possibly thumb), derive (title, mid, canonical_upload_url).
    """
    orig = canonicalize_upload_url(upload_url)
    # Extract filename from last path segment
    fn = unquote(urlparse(orig).path.split("/")[-1])
    title = f"File:{fn}"
    t2, mid, _ = title_mid_url_from_title(session, title)
    return t2 or title, mid, orig


#===============================
# 7) List-mode I/O (file loading only)
# Read Excel/CSV/TSV/TXT into a DataFrame with your single configured column (e.g., CommonsInput) and
# optional SourceCategory. Coalesce legacy headers if present. No network here—just parsing the user’s file.
#================================

def read_list_input() -> pd.DataFrame:
    """
    Load the list file (Excel/CSV/TSV/TXT) into a DataFrame with ONE column:
      <LIST_SINGLE_COLUMN_NAME>
    The cell may contain any of: File:Title | Commons File URL | MID | Concept URI | Upload URL.
    Backward-compat: if legacy columns exist, they are merged into the primary column.
    """
    path = LIST_INPUT_PATH
    fmt = (LIST_INPUT_FORMAT or "auto").lower()
    _, ext = os.path.splitext(path.lower())

    if fmt == "auto":
        if ext in (".xlsx", ".xlsm", ".xls"):
            fmt = "excel"
        elif ext == ".csv":
            fmt = "csv"
        elif ext == ".tsv":
            fmt = "tsv"
        else:
            fmt = "txt"

    if fmt == "excel":
        df = pd.read_excel(path, sheet_name=LIST_EXCEL_SHEET, dtype="string")
    elif fmt == "csv":
        df = pd.read_csv(path, dtype="string", sep=LIST_TEXT_DELIM or ",", engine="python")
    elif fmt == "tsv":
        df = pd.read_csv(path, dtype="string", sep=LIST_TEXT_DELIM or "\t", engine="python")
    else:
        # TXT: default to tab to avoid comma-splitting filenames with commas
        sep = LIST_TEXT_DELIM if LIST_TEXT_DELIM is not None else "\t"
        # Detect header: if first non-empty line matches the primary column name
        with open(path, "r", encoding="utf-8") as f:
            first = (f.readline() or "").strip()
        has_header = first.strip().lower() == LIST_SINGLE_COLUMN_NAME.lower()
        df = pd.read_csv(
            path,
            dtype="string",
            sep=sep,
            engine="python",
            encoding="utf-8",
            header=0 if has_header else None,
        )
        if not has_header:
            # One-column TXT without header → assign the primary column name
            if df.shape[1] == 1:
                df.columns = [LIST_SINGLE_COLUMN_NAME]

    # Normalize headers (trim)
    df = df.rename(columns={c: c.strip() for c in df.columns})

    # If the primary column exists, keep only it
    if LIST_SINGLE_COLUMN_NAME in df.columns:
        df = df[[LIST_SINGLE_COLUMN_NAME]]
        return df

    # Backward-compat: merge legacy columns into the single column
    legacy_cols = [
        "CommonsFileName", "CommonsFileURL", "MediaInfoID",
        "CommonsConceptURI", "CommonsImageURL"
    ]
    present = [c for c in legacy_cols if c in df.columns]

    if present:
        # Merge by priority: MID/Concept → FileName/FileURL → ImageURL
        merged = (
            df.get("MediaInfoID")
              .fillna(df.get("CommonsConceptURI"))
              .fillna(df.get("CommonsFileName"))
              .fillna(df.get("CommonsFileURL"))
              .fillna(df.get("CommonsImageURL"))
        )
        df = pd.DataFrame({LIST_SINGLE_COLUMN_NAME: merged.astype("string")})
        return df

    # If there's exactly one unnamed column, adopt it
    if df.shape[1] == 1:
        df.columns = [LIST_SINGLE_COLUMN_NAME]
        return df

    raise ValueError(
        f"Could not find '{LIST_SINGLE_COLUMN_NAME}' and no legacy columns present; "
        f"found columns: {list(df.columns)}"
    )

#===============================
# 8) Planning (turn inputs → canonical file items)
# -- plan_from_list: For each row, resolve “whatever the user typed” (File title, MID, concept URL, upload URL)
#    into a canonical File: title (+ MID, upload URL). Prints explicit mapping to stdout.
# -- harvest_files_with_depth: BFS a category tree up to DEPTH, dedupe files, capture their category path.
# -- fetch_and_plan_categories: Thin wrapper that logs counts per root category.
# -- print_plan_summary_categories: Pretty-print a per-root breakdown before confirmation:
#     - DEPTH, flatten setting, and extension filter
#     - Total files per root
#     - Count of files at root level vs. subcategory buckets
#     - Top-N subcategory paths with file counts (within the chosen DEPTH)
# -- build_global_selection: Flatten the unique eligible files from all roots into one ordered list,
#     slice it by 1-based [start..end], then regroup per root.
# Outputs are plans (dicts) and never write to disk.
#================================

def plan_from_list(session: requests.Session) -> Dict[str, dict]:
    """
    Build a dict keyed by canonical 'File:...' title from a single-column list file.
    The single column (LIST_SINGLE_COLUMN_NAME) may contain:
      - File:Title
      - Commons file URL (/wiki/File:Title)
      - MediaInfoID (M12345)
      - Commons concept/entity URL (/entity/M12345 or /wiki/Special:EntityData/M12345)
      - Direct upload URL (upload.wikimedia.org, incl. /thumb/...)

    Optional second column LIST_SOURCE_COLUMN is carried to 'SourceCategory'.
    Prints explicit resolution mappings.
    """
    df = read_list_input()

    # Optional range
    if LIST_RANGE_START is not None or LIST_RANGE_END is not None:
        start = (LIST_RANGE_START or 1) - 1
        end = LIST_RANGE_END
        df = df.iloc[start:end]

    files: Dict[str, dict] = {}
    total = len(df)

    for idx, (_, row) in enumerate(df.iterrows(), start=1):
        raw = (row.get(LIST_SINGLE_COLUMN_NAME) or "").strip()
        if not raw:
            print(f"[list {idx}/{total}] ⚠️ Empty row; skipping.")
            continue

        source_cat = (row.get(LIST_SOURCE_COLUMN) or "").strip() if LIST_SOURCE_COLUMN in row.index else ""
        if not source_cat:
            source_cat = LIST_SOURCE_TAG

        title: str = ""
        mid: str = ""
        upload_url: str = ""

        # A) Try MID extraction from the raw string (works for 'M123', entity URLs, etc.)
        mid = extract_mid_from_uri(raw)
        if mid:
            t, m2, u = title_mid_url_from_mid(session, mid)
            title = t or ""
            mid = m2 or mid
            upload_url = u or upload_url
            print(f"[list {idx}/{total}] {raw} → MID {mid} → {title or '(unresolved title)'}")
        else:
            # B) If it's a URL, decide what kind
            if raw.lower().startswith("http"):
                if is_upload_url(raw):
                    t, m2, u = title_mid_url_from_upload_url(session, raw)
                    title = t or title
                    mid = m2 or mid
                    upload_url = u or upload_url
                    print(f"[list {idx}/{total}] ImageURL {raw} → {title or '(unresolved title)'}")
                else:
                    # Try as /wiki/File:… first
                    g = parse_title_from_commons_url(raw)
                    if g:
                        t, m2, u = title_mid_url_from_title(session, g)
                        title = t or g
                        mid = m2 or mid
                        upload_url = u or upload_url
                        print(f"[list {idx}/{total}] FileURL {raw} → {title}")
                    else:
                        # Maybe it's a concept/entity URL after all
                        mid2 = extract_mid_from_uri(raw)
                        if mid2:
                            t, m2, u = title_mid_url_from_mid(session, mid2)
                            title = t or title
                            mid = m2 or mid2
                            upload_url = u or upload_url
                            print(f"[list {idx}/{total}] ConceptURI {raw} → MID {mid} → {title or '(unresolved title)'}")
            else:
                # C) Treat as a File name; add 'File:' if missing
                guess = normalize_title(raw)
                t, m2, u = title_mid_url_from_title(session, guess)
                title = t or guess
                mid = m2 or mid
                upload_url = u or upload_url
                print(f"[list {idx}/{total}] FileName {raw} → {title}")

        if not title:
            print(f"[list {idx}/{total}] ⚠️ Could not resolve row; skipping.")
            continue

        # Deduplicate by title
        if title not in files:
            files[title] = {
                "title": title,
                "mid": mid,
                "upload_url": upload_url,
                "source_category": source_cat,
                "path_segments": [],
            }

    return files


def harvest_files_with_depth(
    session: requests.Session,
    root_category: str,
    max_depth: int,
) -> Tuple[Dict[str, dict], Counter]:
    """
    Traverse a category tree up to 'max_depth'.

    Returns:
      - files_unique: { 'File:…': {
            'title': 'File:…',
            'pageid': 12345,
            'mid': 'M12345',
            'source_category': 'Category:…',   # where first seen
            'path_segments': ['Sub1','Sub2',…] # path from root where first seen
        }, …}
      - raw_membership: Counter keyed by tuple(path_segments) → raw direct count
        e.g. () for root-level, ('Sub1',) for first-level, etc.
    """
    seen_categories: set[str] = set()
    files_unique: Dict[str, dict] = {}
    raw_membership: Counter = Counter()

    q: deque[Tuple[str, int, List[str]]] = deque()
    q.append((root_category, 0, []))
    seen_categories.add(f"Category:{root_category}")

    while q:
        cat, d, path = q.popleft()

        # Count raw direct members for this bucket (no dedupe)
        items = fetch_category_members(session, cat, "file")
        raw_membership[tuple(path)] += len(items)

        # Unique assignment (keep first hit only)
        for item in items:
            title = (item.get("title") or "")
            if not title.startswith("File:"):
                continue
            filename = title.replace("File:", "")
            if FILE_EXTS :
                if not filename.lower().endswith(FILE_EXTS ):
                    continue
            if title not in files_unique:
                pageid = item.get("pageid")
                mid = f"M{pageid}" if pageid and str(pageid).isdigit() else ""
                files_unique[title] = {
                    "title": title,
                    "pageid": pageid,
                    "mid": mid,
                    "source_category": "Category:" + cat,
                    "path_segments": path.copy(),
                }

        # Enqueue subcats if depth allows
        if d < max_depth:
            for item in fetch_category_members(session, cat, "subcat"):
                sub_title = item.get("title") or ""  # "Category:Something"
                if not sub_title.startswith("Category:"):
                    continue
                sub_name = sub_title.replace("Category:", "")
                norm = f"Category:{sub_name}"
                if norm in seen_categories:
                    continue
                seen_categories.add(norm)
                q.append((sub_name, d + 1, path + [sub_name]))

    return files_unique, raw_membership


def fetch_and_plan_categories(
    session: requests.Session,
    root: str,
    max_depth: int,
) -> Tuple[Dict[str, dict], Counter]:
    """Thin wrapper with progress prints; returns (unique_files, raw_membership_counter)."""
    print(f"Scanning Category:{root} (DEPTH={max_depth})… for large category trees this might take some time…")
    files, raw = harvest_files_with_depth(session, root, max_depth)
    print(f"  → {len(files)} unique file(s) eligible for downloading from https://commons.wikimedia.org/wiki/Category:{root}")
    return files, raw


def print_plan_summary_categories(
    grand_plan: List[Tuple[str, Dict[str, dict], Counter]],
    depth: int,
    flatten_paths: bool,
    file_exts : Tuple[str, ...] | tuple = (),
) -> None:
    """
    Nested, alphabetically sorted summary per root:
      - Raw (Commons) counts: direct & totals (no dedupe)
      - Unique-assigned counts: direct & totals (what the script will download)
    """

    def build_tree_from_counter(counter: Counter) -> dict:
        node = {"__direct__": counter.get((), 0), "children": {}}
        for segs, cnt in counter.items():
            if not segs:
                continue
            cur = node
            for seg in segs:
                cur = cur["children"].setdefault(seg, {"__direct__": 0, "children": {}})
            cur["__direct__"] += cnt
        return node

    def add_unique_directs(node: dict, unique_counter: Counter, path: Tuple[str, ...] = ()) -> None:
        # Set unique direct for this node (default 0)
        node.setdefault("__unique_direct__", 0)
        node["__unique_direct__"] += unique_counter.get(path, 0)
        # Recurse into children
        for name, child in node["children"].items():
            add_unique_directs(child, unique_counter, path + (name,))

    def annotate_totals(node: dict) -> Tuple[int, int]:
        """Return (raw_total, unique_total) and store as __raw_total__/__unique_total__."""
        raw_total = node.get("__direct__") or 0
        uniq_total = node.get("__unique_direct__") or 0
        for child in node["children"].values():
            cr, cu = annotate_totals(child)
            raw_total += cr
            uniq_total += cu
        node["__raw_total__"] = raw_total
        node["__unique_total__"] = uniq_total
        return raw_total, uniq_total

    def print_children(node: dict, path_parts: List[str], level: int) -> None:
        for name in sorted(node["children"].keys(), key=lambda s: s.casefold()):
            child = node["children"][name]
            full_path = "/".join(path_parts + [name])
            rd = child.get("__direct__", 0)
            ud = child.get("__unique_direct__", 0)
            rt = child.get("__raw_total__", 0)
            ut = child.get("__unique_total__", 0)
            skipped_here = max(0, rd - ud)
            indent = "  " * level
            print(f"{indent}- {full_path} → Candidates: direct={rd}, total={rt} | "
                  f"Eligible: direct={ud}, total={ut} | Skipped files already present elsewhere in the category tree={skipped_here} | This category is on depth level {level}")
            print_children(child, path_parts + [name], level + 1)

    ext_desc = "ALL" if not file_exts  else ", ".join(file_exts )
    print("\nPLAN SUMMARY")
    print(f"  MODE=categories | DEPTH={depth} | flatten_paths={flatten_paths} | extensions filter={ext_desc}\n")

    grand_total_raw = 0
    grand_total_unique = 0

    for root, files_unique, raw_counter in grand_plan:
        # Unique direct counts by assigned bucket
        unique_bucket = Counter(tuple(m.get("path_segments", [])) for m in files_unique.values())

        # Build tree from raw (Commons) counts, then layer in unique counts
        tree = build_tree_from_counter(raw_counter)
        add_unique_directs(tree, unique_bucket)
        annotate_totals(tree)

        # Totals for this root
        root_raw_direct = tree.get("__direct__", 0)
        root_unique_direct = tree.get("__unique_direct__", 0)
        root_raw_total = tree.get("__raw_total__", 0)
        root_unique_total = tree.get("__unique_total__", 0)
        grand_total_raw += root_raw_total
        grand_total_unique += root_unique_total

        # Header per root
        non_root_buckets = len([k for k in raw_counter.keys() if k])
        print(
            f"• Category:{root} → Raw non-unique candidates={root_raw_total} | Unique & eligible files={root_unique_total} "
            f"(root-level: candidates={root_raw_direct}, eligible={root_unique_direct}; "
            f"buckets={non_root_buckets})"
        )

        # Nested tree under this root
        print_children(tree, [root], level=1)

    print(f"\nGRAND TOTAL across {len(grand_plan)} root(s): "
          f"Raw candidates={grand_total_raw} file(s) | Eligible & unique for downloading={grand_total_unique} file(s).")

def build_global_selection(
    grand_plan: List[Tuple[str, Dict[str, dict], Counter]],
    normalized_roots: List[str],
    start_1based: Optional[int],
    end_1based: Optional[int],
    order: str = "root_then_title",
) -> Tuple[Dict[str, Dict[str, dict]], int, int]:
    """
    Flatten the unique eligible files from all roots into one ordered list,
    slice it by 1-based [start..end], then regroup per root.

    Returns:
      selected_by_root: { root: { 'File:Title': meta, ... }, ... }
      total_pool: total unique files across all roots (before slicing)
      selected_count: number selected by slice
    """
    # Build pool [(key, root, title, meta), ...] with deterministic ordering
    pool: List[Tuple[str, str, str, dict]] = []

    if order == "title":
        # All titles across roots, A→Z (case-insensitive)
        for root, files_unique, _raw in grand_plan:
            for t, m in files_unique.items():
                pool.append((t.casefold(), root, t, m))
        pool.sort(key=lambda x: x[0])
    else:
        # "root_then_title": iterate roots in configured order; A→Z within each root
        root_map = {r: fu for (r, fu, _raw) in grand_plan}
        for root in normalized_roots:
            fu = root_map.get(root, {})
            for t in sorted(fu.keys(), key=str.casefold):
                pool.append(("", root, t, fu[t]))  # key not used for sorting here

    total_pool = len(pool)

    # Compute 0-based slice
    if (start_1based is not None) or (end_1based is not None):
        start = max(0, (start_1based or 1) - 1)
        end = end_1based  # None means to the end
        pool = pool[start:end]

    # Group back by root
    selected_by_root: Dict[str, Dict[str, dict]] = {}
    for _k, root, title, meta in pool:
        selected_by_root.setdefault(root, {})[title] = meta

    return selected_by_root, total_pool, sum(len(d) for d in selected_by_root.values())


#===============================
# 9) Filesystem target resolution
# Given a base, a root segment, and subcategory path segments, build (and optionally create) the target directory.
# All path composition funnels through here so your folder structure is predictable.
#================================

def ensure_folder(base: str, root_segment: str, path_segments: List[str], create: bool = True) -> str:
    """
    Build a nested folder path:
      LOCAL_BASE_FOLDER / <root_segment> / <sanitized path_segments...>
    If create=True, create directories.
    """
    parts = [base, sanitize_folder_component(root_segment)] if root_segment else [base]
    parts += [sanitize_folder_component(p) for p in path_segments]
    folder = os.path.join(*parts)
    if create and not DRY_RUN:
        os.makedirs(folder, exist_ok=True)
    return folder

#===============================
# 10) Download + logging
# -- download_file: Follow Special:FilePath redirects, save with safe name, always return the final
#   upload.wikimedia.org URL; supports OVERWRITE_EXISTING and DRY_RUN.
# -- append_rows_to_excel_log: Incrementally write rows to the Excel log (chunked), dedupe by
#   (CommonsFileURL, MediaInfoID), enforce column order, and replace the sheet in place.
#================================

def download_file(
    session: requests.Session,
    url: str,
    folder: str,
    filename: str,
    mid: Optional[str] = None,
    title: Optional[str] = None,   # 'File:...' for API fallback
    dry_run: bool = DRY_RUN,
) -> tuple[str, str]:
    """
    Download via Special:FilePath into 'folder' using a sanitized, path-length-safe name.
    Returns (final_stored_filename, resolved_upload_url).
    The resolved URL is always filled (via API if we didn't download).
    In DRY_RUN, nothing is written; we still resolve upload URL via API.
    """
    safe_name = sanitize_filename(filename, folder=folder, mid=mid)
    out_path = os.path.join(folder, safe_name)

    # DRY RUN: resolve URL via API and return
    if dry_run:
        resolved_url = resolve_upload_url_via_api(session, title or f"File:{filename}")
        return safe_name, resolved_url

    headers = {"User-Agent": UA, "Accept": "*/*"}

    # If we're not overwriting and file exists, just resolve upload URL via API and return.
    if not OVERWRITE_EXISTING and os.path.exists(out_path):
        resolved_url = resolve_upload_url_via_api(session, title or f"File:{filename}")
        return safe_name, resolved_url

    # Fetch via Special:FilePath (follows redirects to upload.wikimedia.org)
    with session.get(url, headers=headers, stream=True, timeout=60) as r:
        r.raise_for_status()
        resolved_url = r.url  # should be https://upload.wikimedia.org/...
        if not DRY_RUN:
            with open(out_path, "wb") as f:
                for chunk in r.iter_content(1024 * 128):
                    if chunk:
                        f.write(chunk)

    # Safety: if for some reason it's not the upload host, ask the API
    if not resolved_url.startswith("https://upload.wikimedia.org/"):
        api_url = resolve_upload_url_via_api(session, title or f"File:{filename}")
        if api_url:
            resolved_url = api_url

    return safe_name, resolved_url

def append_rows_to_excel_log(
    base_folder: str,
    rows: List[dict],
    sheet_name: str,
    dedupe_keys: Tuple[str, str] = ("CommonsFileURL", "MediaInfoID"),
) -> str:
    """
    Append rows to the Excel log (create if missing), drop duplicates by 'dedupe_keys',
    and replace the sheet with the deduped result.

    Log file location:
      - If EXCEL_LOG_DIR is set, the log is written there (relative to script folder if not absolute).
      - Otherwise it's written inside 'base_folder' (DEFAULT).

    Columns (left-most first):
      - CommonsFileName
      - CommonsFileURL
      - MediaInfoID
      - CommonsConceptURI
      - CommonsImageURL
      - SourceCategory
      - LocalBaseFolder
      - CommonsCategoryPath
      - LocalSubFolder
      - LocalFilename
    """
    # Respect DRY_RUN: do not write any logs
    if DRY_RUN:
        return os.path.join(resolve_log_dir(base_folder), EXCEL_LOG_NAME)

    if not rows:
        log_dir = resolve_log_dir(base_folder)
        return os.path.join(log_dir, EXCEL_LOG_NAME)

    log_dir = resolve_log_dir(base_folder)
    os.makedirs(log_dir, exist_ok=True)
    log_path = os.path.join(log_dir, EXCEL_LOG_NAME)

    desired_order = [
        "CommonsFileName",
        "CommonsFileURL",
        "MediaInfoID",
        "CommonsConceptURI",
        "CommonsImageURL",
        "SourceCategory",
        "CommonsCategoryPath",
        "LocalBaseFolder",
        "LocalSubFolder",
        "LocalFilename",
    ]

    new_df = pd.DataFrame(rows)

    # First write (no file yet)
    if not os.path.exists(log_path):
        cols = [c for c in desired_order if c in new_df.columns] + \
               [c for c in new_df.columns if c not in desired_order]
        new_df = new_df.reindex(columns=cols)
        with pd.ExcelWriter(log_path, engine="openpyxl") as w:
            new_df.to_excel(w, sheet_name=sheet_name, index=False)
        return log_path

    # Append + de-dup
    try:
        existing = pd.read_excel(log_path, sheet_name=sheet_name, dtype="string")
    except ValueError:
        existing = pd.DataFrame(columns=desired_order)

    all_cols = list(dict.fromkeys(list(existing.columns) + list(new_df.columns)))
    existing = existing.reindex(columns=all_cols)
    new_df = new_df.reindex(columns=all_cols)
    combined = pd.concat([existing, new_df], ignore_index=True)

    present_keys = [k for k in dedupe_keys if k in combined.columns]
    if present_keys:
        combined = combined.drop_duplicates(subset=present_keys, keep="first", ignore_index=True)

    cols = [c for c in desired_order if c in combined.columns] + \
           [c for c in combined.columns if c not in desired_order]
    combined = combined.reindex(columns=cols)

    with pd.ExcelWriter(log_path, engine="openpyxl", mode="a", if_sheet_exists="replace") as w:
        combined.to_excel(w, sheet_name=sheet_name, index=False)

    return log_path


#===============================
# 11) Execution units (do work on a batch)
# perform_downloads: The per-batch worker. Iterates the planned items, downloads each file, prints progress [i/n],
# and flushes rows to the log every LOG_FLUSH_ROWS. Fills admin columns, including
# LocalBaseFolder, LocalSubFolder, and LocalFilename.
#================================

def perform_downloads(
    session: requests.Session,
    root_segment: str,                   # e.g., "from-categories" or "from-list"
    files: Dict[str, dict],
    base_folder: str,
    sheet_name: str,
    log_flush_rows: int = LOG_FLUSH_ROWS,
    flatten_paths: bool = False,         # <— NEW
) -> int:
    """
    Download all files for one root segment and append rows to Excel in small chunks.
    Returns the number of successfully processed files.
    """
    rows_buffer: List[dict] = []
    total = len(files)
    written = 0

    for idx, (title, meta) in enumerate(files.items(), 1):
        filename = title.replace("File:", "")

        # Full, original subpath (for logging):
        orig_segments = list(meta.get("path_segments", []))
        commons_category_path = os.path.join(*orig_segments) if orig_segments else ""

        # Decide what subpath to use on disk:
        # - categories mode: path_segments are [RootCategory, Sub1, Sub2, ...]
        # - flatten => keep only [RootCategory]; non-flatten => keep all
        if flatten_paths and orig_segments:
            local_segments = [orig_segments[0]]     # only RootCategory
        else:
            local_segments = orig_segments

        # Build target folder (LOCAL_BASE_FOLDER / root_segment / local_segments...)
        folder = ensure_folder(base_folder, root_segment, local_segments, create=not DRY_RUN)

        # Relative subfolder path under the LocalBaseFolder (what exists on disk)
        subfolder_rel = os.path.relpath(folder, start=base_folder)
        if subfolder_rel == ".":
            subfolder_rel = ""

        # Build Special:FilePath URL
        file_url = f"https://commons.wikimedia.org/wiki/Special:FilePath/{filename}"
        escaped_url = quote(file_url, safe='/:')
        try:
            stored_name, upload_url = download_file(
                session=session,
                url=escaped_url,
                folder=folder,
                filename=filename,
                mid=meta.get("mid"),
                title=meta.get("title") or title,   # 'File:...'
                dry_run=DRY_RUN,
            )
        except Exception as e:
            print(f"[{idx}/{total}] {title} → ERROR: {e}")
            continue

        row = {
            "CommonsFileName": title,
            "CommonsFileURL": f"https://commons.wikimedia.org/wiki/{title}",
            "MediaInfoID": meta.get("mid", ""),
            "CommonsConceptURI": (f"https://commons.wikimedia.org/entity/{meta['mid']}" if meta.get("mid") else ""),
            "CommonsImageURL": upload_url,
            "SourceCategory": meta.get("source_category", ""),
            "LocalBaseFolder": base_folder,
            "CommonsCategoryPath": commons_category_path,   # <— NEW: full original chain
            "LocalSubFolder": subfolder_rel,           # what’s actually on disk
            "LocalFilename": stored_name,
        }
        rows_buffer.append(row)
        written += 1

        verb = "would save as" if DRY_RUN else "saved as"
        print(f"[{idx}/{total}] {title} → {verb} {stored_name}")

        # Chunked flush to Excel log
        if not DRY_RUN and log_flush_rows and (len(rows_buffer) >= log_flush_rows):
            path = append_rows_to_excel_log(base_folder, rows_buffer, sheet_name=sheet_name)
            print(f"   ↳ flushed {len(rows_buffer)} row(s) to Excel log: {path}")
            rows_buffer.clear()

    # Final flush
    if not DRY_RUN and rows_buffer:
        path = append_rows_to_excel_log(base_folder, rows_buffer, sheet_name=sheet_name)
        print(f"   ↳ final flush: {len(rows_buffer)} row(s) to Excel log: {path}")
        rows_buffer.clear()

    return written



#===============================
# 12) Entry point
# main: Wires everything together depending on MODE:
# -- categories → harvest plans (under from-categories/<Root>/<Subcats…>), optional confirmation, download with
#    chunked logging to SHEET_NAME_CATEGORIES.
# -- list → read/resolve single-column list (under from-list/), optional confirmation, download to SHEET_NAME_LIST.
# Respects DRY_RUN throughout, and prints a clear end-of-run summary.
#================================

def main():
    # Create base folder (unless DRY_RUN)
    if not DRY_RUN:
        os.makedirs(LOCAL_BASE_FOLDER, exist_ok=True)

    session = build_session()

    if MODE.lower() == "categories":
        # Normalize category names (spaces → underscores) and slice which ROOTS to include (if requested)
        normalized_roots = [c.replace(" ", "_") for c in CATEGORIES]
        if CATEGORIES_RANGE_START is not None or CATEGORIES_RANGE_END is not None:
            start = (CATEGORIES_RANGE_START or 1) - 1
            end = CATEGORIES_RANGE_END
            normalized_roots = normalized_roots[start:end]

        # 1) Preflight: collect unique files AND raw membership per root
        grand_plan: List[Tuple[str, Dict[str, dict], Counter]] = []
        grand_total_unique_full = 0
        for root in normalized_roots:
            files_unique, raw_counter = fetch_and_plan_categories(session, root, DEPTH)
            grand_plan.append((root, files_unique, raw_counter))
            grand_total_unique_full += len(files_unique)

        if grand_total_unique_full == 0:
            print("No images found with the current settings (depth and filters). Nothing to do.")
            return

        # 1b) Show nested, alphabetically sorted plan with RAW vs UNIQUE counts (full, before slicing)
        print_plan_summary_categories(
            grand_plan=grand_plan,
            depth=DEPTH,
            flatten_paths=FLATTEN_CATEGORIES_PATHS,
            file_exts =FILE_EXTS ,
        )

        # 1c) Build a GLOBAL slice across all roots (after unique/eligible pool is built)
        selected_by_root, total_pool, selected_count = build_global_selection(
            grand_plan=grand_plan,
            normalized_roots=normalized_roots,
            start_1based=CATEGORIES_GLOBAL_RANGE_START,
            end_1based=CATEGORIES_GLOBAL_RANGE_END,
            order=CATEGORIES_GLOBAL_SLICE_ORDER,
        )

        # Small post-slice summary per root (concise)
        if (CATEGORIES_GLOBAL_RANGE_START is not None) or (CATEGORIES_GLOBAL_RANGE_END is not None):
            s_from = CATEGORIES_GLOBAL_RANGE_START or 1
            s_to = CATEGORIES_GLOBAL_RANGE_END or total_pool
            print(
                f"\nGLOBAL SLICE: selecting items {s_from}–{s_to} of {total_pool} UNIQUE eligible files "
                f"(order={CATEGORIES_GLOBAL_SLICE_ORDER})."
            )
            for root in normalized_roots:
                n = len(selected_by_root.get(root, {}))
                print(f"  - Category:{root} → {n} file(s) selected")

        # 2) Confirmation prompt (explicit about DEPTH/flatten/settings and slice)
        if CONFIRM_BEFORE_DOWNLOAD:
            slice_note = ""
            if (CATEGORIES_GLOBAL_RANGE_START is not None) or (CATEGORIES_GLOBAL_RANGE_END is not None):
                s_from = CATEGORIES_GLOBAL_RANGE_START or 1
                s_to = CATEGORIES_GLOBAL_RANGE_END or total_pool
                slice_note = f" | global slice {s_from}–{s_to} of {total_pool}"
            print(
                f"\nThis will {'(dry-run) ' if DRY_RUN else ''}"
                f"download {selected_count} UNIQUE file(s) across {len(normalized_roots)} root category/ies "
                f"at DEPTH={DEPTH} (flatten_paths={FLATTEN_CATEGORIES_PATHS}){slice_note}."
            )
            answer = input("Proceed with download? [y/N]: ").strip().lower()
            if answer not in {"y", "yes"}:
                print("Aborted. Adjust DEPTH or slice settings, then run again.")
                return

        # 3) Download + incremental Excel logging (only the selected subset)
        sheet_name = SHEET_NAME_CATEGORIES
        total_written = 0
        for root in normalized_roots:
            files_for_root = selected_by_root.get(root, {})
            if not files_for_root:
                continue
            # Prefix the root into the path segments so we can flatten to <from-categories>/<root>/...
            adjusted = {
                t: {**m, "path_segments": [root] + m.get("path_segments", [])}
                for t, m in files_for_root.items()
            }
            total_written += perform_downloads(
                session,
                root_segment=CATEGORIES_DOWNLOAD_SUBFOLDER,
                files=adjusted,
                base_folder=LOCAL_BASE_FOLDER,
                sheet_name=sheet_name,
                flatten_paths=FLATTEN_CATEGORIES_PATHS,
            )

        if total_written:
            print(f"\n✅ Finished. {'(dry-run) ' if DRY_RUN else ''}Processed {total_written} file(s).")
            if not DRY_RUN:
                print(
                    f"   Excel log updated incrementally: "
                    f"{os.path.join(resolve_log_dir(LOCAL_BASE_FOLDER), EXCEL_LOG_NAME)} (sheet: {sheet_name})"
                )
        else:
            print("\nNo files were processed; no Excel log written.")

    elif MODE.lower() == "list":
        # 1) Preflight: build plan from list
        print(f"Loading list from {LIST_INPUT_PATH} …")
        files = plan_from_list(session)
        titles = list(files.keys())
        print(f"  → resolved {len(titles)} unique image(s) from list input.")

        if len(titles) == 0:
            print("Nothing to do (no valid rows).")
            return

        # 2) Confirmation
        if CONFIRM_BEFORE_DOWNLOAD:
            print(
                f"\nThis will {'(dry-run) ' if DRY_RUN else ''}"
                f"process a TOTAL of {len(titles)} file(s) from the list."
            )
            answer = input("Proceed? [y/N]: ").strip().lower()
            if answer not in {"y", "yes"}:
                print("Aborted. Adjust the list or range settings, then run again.")
                return

        # For list mode, use a single root segment under LOCAL_BASE_FOLDER
        root_segment = LIST_DOWNLOAD_SUBFOLDER
        sheet_name = SHEET_NAME_LIST

        total_written = perform_downloads(
            session, root_segment, files, LOCAL_BASE_FOLDER, sheet_name=sheet_name
        )

        if total_written:
            print(f"\n✅ Finished. {'(dry-run) ' if DRY_RUN else ''}Processed {total_written} file(s).")
            if not DRY_RUN:
                print(
                    f"   Excel log updated incrementally: "
                    f"{os.path.join(resolve_log_dir(LOCAL_BASE_FOLDER), EXCEL_LOG_NAME)} (sheet: {sheet_name})"
                )
        else:
            print("\nNo files were processed; no Excel log written.")

    else:
        print(f"Unknown MODE='{MODE}'. Use 'categories' or 'list'.")

if __name__ == "__main__":
    main()