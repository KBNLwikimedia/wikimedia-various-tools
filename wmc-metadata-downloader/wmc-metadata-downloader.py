"""
Wikimedia Commons File Metadata Downloader & Excel Writer

Overview
========
This script builds a reliable, resumable pipeline for collecting **per-file metadata**
from Wikimedia Commons files or categories and writing it into an Excel workbook. It supports two
complementary modes:

1) **manual-list** — read a list of Commons file titles from an input sheet.
2) **category** — harvest all (or a selected **range**) of file titles from a Commons
   category with API continuation, write them into an input sheet, and then process
   them the same way as the manual list.

For **every file**, the script:
- Fetches full JSON metadata via the Commons API (`prop=imageinfo`, including
  `extmetadata`, `url`, `size`, `sha1`, `mime`, `mediatype`, `timestamp`, `user`).
- Derives the **MediaInfo ID (MID)** from the MediaWiki pageid (e.g., `"M12345"`),
  and constructs a human URL to the entity page.
- Saves the **exact JSON response** to disk in `downloaded_metadata/` using a
  **Windows-safe** filename of the form:
  `<CommonsFileName>__<MID or NOID>.json`. If the full path would exceed a
  conservative length budget, the filename is truncated and a short hash is added.
  **If the same filename is produced again, the file is overwritten**.
- Flattens the JSON into **dotted columns** and appends rows to an **output sheet**
  in the same workbook, **in chunks** (batches) so that progress is durable for
  very large lists/categories.

Workbook & Sheets
=================
- **Workbook** (both modes): ``wmc-inputfiles.xlsx``

- **Manual-list mode**
  - **Input sheet:** ``Files-Manual``
    Columns:
    - ``CommonsFileName`` (required; e.g., ``File:Example.jpg`` — the script will
      normalize missing ``File:`` prefixes automatically)
    - ``SourceCategory`` (optional; carried through to outputs)
  - **Output sheet:** ``FilesMetadata-Manual``

- **Category mode**
  - **Harvest sheet:** ``Files-Category`` (built/updated by the script)
    Columns:
    - ``CommonsFileName`` (harvested from the category)
    - ``SourceCategory`` (the category you harvested)
  - **Output sheet:** ``FilesMetadata-Category``

Output Columns (both modes)
===========================
Each row written to the output sheet contains **front/base columns** followed by
**all flattened JSON fields**:

- ``Input_CommonsFileName`` — original title from the input sheet
- ``SourceCategory`` — carried through from input/harvest
- ``Requested_API_URL`` — the fully prepared request URL used
- ``Local_JSON_File`` — path to the saved JSON file
- ``Computed_MediaID`` — e.g., ``M12345`` (empty if not found)
- ``Computed_MediaID_URL`` — link to the entity page (empty if no MID)
- ``BatchIndex`` — 1-based index of the chunk that wrote this row
- ``… (flattened JSON columns)`` — e.g., ``query.pages.0.title``,
  ``query.pages.0.imageinfo.0.url``,
  ``query.pages.0.imageinfo.0.extmetadata.Artist.value``, etc.

Modes & Flow
============
**Manual-list mode**
1. Read ``Files-Manual`` (Columns: ``CommonsFileName``, optional ``SourceCategory``).
2. Normalize titles to ``File:…``.
3. Process in **chunks** (see CHUNK_SIZE): fetch JSON → save JSON → flatten →
   append rows to ``FilesMetadata-Manual`` (expanding columns when new keys appear).

**Category mode**
1. **Harvest** from ``CATEGORY_TITLE`` using `list=categorymembers`:
   - Handles API **continuation** (`cmcontinue`) until the category is exhausted,
     or until an optional **1-based inclusive range** is satisfied
     (``CATEGORY_RANGE_START``, ``CATEGORY_RANGE_END``; set in code).
   - Writes harvested titles in **batches** to ``Files-Category`` (see
     HARVEST_FLUSH_ROWS).
2. **Process** the harvested sheet in **chunks** (see CHUNK_SIZE), appending to
   ``FilesMetadata-Category`` exactly as in manual-list mode.

Chunking & Flush Cadence
========================
- ``HARVEST_FLUSH_ROWS`` (category mode only): how many harvested file titles to
  buffer **before writing them to** ``Files-Category`` during the *harvest phase*.
  Smaller = more frequent Excel writes; larger = fewer writes but more memory.
- ``CHUNK_SIZE`` (both modes): how many input rows to process per batch during the
  *processing phase*. Each batch is **appended** to the output sheet; if a batch
  introduces new JSON columns, the sheet is widened and replaced once.

Network Etiquette & Robustness
==============================
- Requests are sent with a polite **User-Agent** that includes contact info:
  ``KB WMC metadata fetcher - User:OlafJanssen - Contact: olaf.janssen@kb.nl)``.
- Automatic **retry/backoff** on transient errors (HTTP 429/5xx).
- The API call follows Commons **redirects** (e.g., for renamed files).
- Per-file progress is printed:
  ``[123/8120] Fetching File:Example.jpg … done (MID=M123456)``

JSON Files on Disk
==================
- Directory: ``downloaded_metadata/`` (created if missing).
- Naming: ``<CommonsFileName>__<MID or NOID>.json``
  If too long for Windows, the script truncates the base name and adds an
  8-character hash; as a last resort the filename collapses to
  ``<MID or NOID>__<hash>.json`` or just ``<hash>.json``.
- **Overwrite policy:** if a run generates the **same** filename as a previous
  run, it **overwrites** that JSON file.

Configuration (edit in-code)
============================
The configuration block is included **verbatim** near the top of the script. Key settings:

- **Mode selection**
  - ``MODE``: ``"manual-list"`` or ``"category"``
  - ``CATEGORY_TITLE``: required for category mode
- **Workbook & sheets**
  - ``XLSX_PATH = "wmc-inputfiles.xlsx"``
  - Manual input/output: ``Files-Manual`` → ``FilesMetadata-Manual``
  - Category input/output: ``Files-Category`` → ``FilesMetadata-Category``
- **Harvesting**
  - ``CATEGORY_PAGE_LIMIT`` (usually 500 for non-bot)
  - ``HARVEST_FLUSH_ROWS`` (category mode only)
  - Optional: ``CATEGORY_RANGE_START`` / ``CATEGORY_RANGE_END`` (1-based inclusive)
- **Processing**
  - ``CHUNK_SIZE`` (both modes)
- **API & files**
  - ``EXTMETA_LANG`` (e.g., ``"en"`` or ``"nl"``), ``USER_AGENT``, ``DOWNLOAD_DIR``
  - ``FULL_PATH_BUDGET`` (conservative Windows full-path length limit)

Installation & Running
======================
- Requires Python 3.9+ recommended. Install dependencies:

  ``pip install -r requirements.txt``

  with:

  ``pandas>=1.5.0``, ``openpyxl>=3.1.0``, ``requests>=2.31.0``, ``urllib3>=2.0.0``

- Ensure the Excel workbook is **closed** before running.
- Edit the config block (MODE, CATEGORY_TITLE, etc.) and execute:

  ``python wmc_metadata.py``

Error Handling & Limits
=======================
- Input reading errors (missing workbook/sheet/column) raise clear exceptions.
- Network failures are captured per file; the row will include an ``error.message``
  column and processing will continue.
- JSON file write failures are logged; the row still lands in Excel with an empty
  ``Local_JSON_File`` path.
- For extremely large categories, expect a long harvest with many continuation
  steps—intermediate writes keep progress durable.

Notes
=====
- ``SourceCategory`` is a free-form input in manual-list mode; in category mode it
  is set to ``CATEGORY_TITLE`` for each harvested row.
- The output schema widens dynamically as new JSON keys appear in later chunks.
- If you enable long paths on Windows, you can increase ``FULL_PATH_BUDGET``.

Author: ChatGPT, prompted by Olaf Janssen, KB (Koninklijke Bibliotheek)
Latest update: 2025-10-22
License: CC0
"""

from __future__ import annotations
import hashlib
import json
import math
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple
import pandas as pd
import requests
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry


# =========================
# YOUR CONFIG
# =========================
# Comment out the mode you do NOT want to use:
#MODE = "manual-list" # Comment out if MODE = category
MODE = "category"  # In this mode, you must configure the CATEGORY_TITLE below.
CATEGORY_TITLE = "Category:Catchpenny prints from Koninklijke Bibliotheek"
# Optional range for CATEGORY mode (1-based inclusive). First 10 files in this category:
CATEGORY_RANGE_START: Optional[int] = 1  # Set to None to process all
CATEGORY_RANGE_END:   Optional[int] = 10  # None = process all

# Adapt the User-Agent for with your own details
USER_AGENT = "Wikimedia Commons File Metadata Downloader - User:OlafJanssen - Contact: olaf.janssen@kb.nl)"

# Input workbook - required in both modes
XLSX_PATH = "wmc-inputfiles.xlsx"

# Input sheets
INPUT_SHEET_MANUAL = "Files-Manual"
INPUT_SHEET_CATEGORY = "Files-Category"
# Output sheets
OUTPUT_SHEET_MANUAL = "FilesMetadata-Manual"
OUTPUT_SHEET_CATEGORY = "FilesMetadata-Category"

# Category harvesting
CATEGORY_PAGE_LIMIT = 500  # API max for non-bot

 # HARVEST_FLUSH_ROWS : Category-mode only. How many harvested filenames (from the Commons category) we
# accumulate before writing them into the input sheet WMCFiles-Category.
HARVEST_FLUSH_ROWS = 100

# CHUNK_SIZE → Both modes. How many files we actually process per batch (fetch JSON, save per-file JSON, flatten,
# and append rows to the output sheet).
CHUNK_SIZE = 100  # per your spec

# Where to drop per-file JSON
DOWNLOAD_DIR = Path("downloaded_metadata")

# API & etiquette
COMMONS_API = "https://commons.wikimedia.org/w/api.php"
EXTMETA_LANG = "en"  # or "nl"

# Requests
TIMEOUT_SECS = 20
RETRIES_TOTAL = 5
RETRIES_BACKOFF = 0.6

# Windows path safety
FULL_PATH_BUDGET = 240  # conservative full-path length budget

# =========================

# ---------- HTTP session ----------

def build_session() -> requests.Session:
    """
    Create a requests.Session configured with retry/backoff and a polite User-Agent.

    Returns:
        requests.Session: session preconfigured for Commons requests.

    Notes:
        - Retries on 429/5xx with exponential backoff.
        - UA contains contact info per Wikimedia API etiquette.
    """
    s = requests.Session()
    retry = Retry(
        total=RETRIES_TOTAL,
        connect=RETRIES_TOTAL,
        read=RETRIES_TOTAL,
        status=RETRIES_TOTAL,
        status_forcelist=(429, 500, 502, 503, 504),
        allowed_methods=frozenset({"GET"}),
        backoff_factor=RETRIES_BACKOFF,
        raise_on_status=False,
    )
    s.mount("https://", HTTPAdapter(max_retries=retry))
    s.headers.update({"User-Agent": USER_AGENT, "Accept": "application/json"})
    return s


# ---------- Utilities ----------

def norm_file_title(name: str) -> str:
    """
    Ensure a Commons title is prefixed with 'File:'.

    Args:
        name: raw title or filename.

    Returns:
        str: normalized title or empty string if input was empty.
    """
    name = (name or "").strip()
    if not name:
        return ""
    return name if name.lower().startswith("file:") else f"File:{name}"


def compute_mid(pageid: Optional[str]) -> str:
    """
    Convert a MediaWiki pageid to a MediaInfo ID (MID).

    Args:
        pageid: page id from the API.

    Returns:
        str: 'M{pageid}' or '' if not numeric/empty.
    """
    return f"M{pageid}" if pageid and str(pageid).isdigit() else ""


def mid_url(mid: str) -> str:
    """
    Build a human-readable MediaInfo entity URL.

    Args:
        mid: MediaInfo id (e.g., 'M12345').

    Returns:
        str: URL or empty string if mid is empty.
    """
    return f"https://commons.wikimedia.org/wiki/Special:EntityPage/{mid}" if mid else ""


def flatten_json(obj: Any, parent_key: str = "", sep: str = ".") -> Dict[str, Any]:
    """
    Flatten a nested dict/list/scalar JSON-like object to a single dict with dotted keys.

    Args:
        obj: dict/list/scalar to flatten.
        parent_key: prefix for recursive calls.
        sep: separator for keys.

    Returns:
        Dict[str, Any]: flattened mapping.

    Safety:
        - Never raises on unexpected shapes; treats unknowns as scalars.
    """
    items: List[Tuple[str, Any]] = []
    if isinstance(obj, dict):
        for k, v in obj.items():
            new_key = f"{parent_key}{sep}{k}" if parent_key else str(k)
            items.extend(flatten_json(v, new_key, sep=sep).items())
    elif isinstance(obj, list):
        for i, v in enumerate(obj):
            new_key = f"{parent_key}{sep}{i}" if parent_key else str(i)
            items.extend(flatten_json(v, new_key, sep=sep).items())
    else:
        items.append((parent_key, obj))
    return dict(items)


def extract_pageid_title(data: Dict[str, Any]) -> Tuple[str, str]:
    """
    Safely read pageid and title from a Commons API 'formatversion=2' response.

    Args:
        data: parsed JSON from API.

    Returns:
        (pageid, title): both '' if missing/missing page.
    """
    try:
        pages = (data.get("query") or {}).get("pages") or []
        if not pages:
            return "", ""
        page = pages[0] or {}
        if page.get("missing"):
            return "", ""
        pid = str(page.get("pageid") or "")
        ttl = str(page.get("title") or "")
        return pid, ttl
    except Exception:
        return "", ""


def safe_component(value: str) -> str:
    """
    Sanitize text for filesystem compatibility (keep alnum, '-', '_', '.').

    Args:
        value: input string.

    Returns:
        str: sanitized or 'NA' if empty.
    """
    value = (value or "").strip()
    if not value:
        return "NA"
    return "".join(c if c.isalnum() or c in ("-", "_", ".") else "_" for c in value)


def short_hash(text: str, n: int = 8) -> str:
    """
    Stable short hash for filename disambiguation.

    Args:
        text: source text to hash.
        n: length of hex digest to return.

    Returns:
        str: first n hex chars of BLAKE2s digest.
    """
    return hashlib.blake2s(text.encode("utf-8"), digest_size=8).hexdigest()[:n]


def build_safe_json_path(
    dir_path: Path,
    input_name: str,
    mid: str,
    budget_full_path: int = FULL_PATH_BUDGET,
    ext: str = ".json",
) -> Path:
    """
    Build a Windows-safe JSON file path derived from input name + MID, keeping
    the FULL PATH below a conservative length budget.

    Behavior:
        - Uses '<input_basename>__<MID or NOID>.json' when possible.
        - If too long, truncates base and adds an 8-char hash.
        - Final fallback: '<MID or NOID>__<hash>.json' or just '<hash>.json'.
        - Deterministic; if the same path is returned twice, the caller will
          OVERWRITE it by opening with mode='w'.

    Args:
        dir_path: directory for the JSON file.
        input_name: original Commons file name (with or without 'File:').
        mid: MediaInfo ID or '' if not found.
        budget_full_path: max allowed length for full path string.
        ext: file extension.

    Returns:
        Path: filesystem path (not created).
    """
    base_raw = (input_name or "").strip()
    base_core = base_raw[5:] if base_raw.lower().startswith("file:") else base_raw
    base = safe_component(base_core) or "NA"
    mid_tag = safe_component(mid or "NOID")

    candidate = f"{base}__{mid_tag}{ext}"

    def within_budget(fname: str) -> bool:
        try:
            return len(str((dir_path / fname).resolve())) <= budget_full_path
        except Exception:
            # If resolve() fails on some FS, fall back to len(dirname)+len(fname)
            return len(str(dir_path)) + 1 + len(fname) <= budget_full_path

    # 1) Try full base
    if within_budget(candidate):
        return dir_path / candidate

    # 2) Add short hash and truncate base to fit
    h = short_hash(base_core)
    suffix = f"__{h}__{mid_tag}{ext}"

    try:
        dir_abs = str(dir_path.resolve())
    except Exception:
        dir_abs = str(dir_path)
    avail_for_fname = max(16, budget_full_path - len(dir_abs) - 1)  # minus path sep
    room_for_base = max(0, avail_for_fname - len(suffix))

    if room_for_base > 0:
        base_trunc = base[:room_for_base]
        candidate2 = f"{base_trunc}{suffix}"
        if within_budget(candidate2):
            return dir_path / candidate2

    # 3) Final fallbacks
    fallback = f"{mid_tag}__{h}{ext}"
    if within_budget(fallback):
        return dir_path / fallback

    tiny = f"{h}{ext}"
    return dir_path / tiny


# ---------- Excel helpers ----------

def sheet_exists(xlsx_path: str, sheet_name: str) -> bool:
    """
    Check if a sheet exists in the workbook.

    Args:
        xlsx_path: path to workbook.
        sheet_name: sheet to check.

    Returns:
        bool: True if sheet is present, False otherwise.
    """
    try:
        with pd.ExcelFile(xlsx_path) as xf:
            return sheet_name in xf.sheet_names
    except FileNotFoundError:
        return False
    except Exception as e:
        print(f"⚠️  Could not read workbook '{xlsx_path}': {e}")
        return False


def read_sheet_df(xlsx_path: str, sheet_name: str) -> pd.DataFrame:
    """
    Read a sheet as DataFrame with dtype=object.

    Raises:
        FileNotFoundError: if workbook is missing.
        ValueError: if sheet missing.
    """
    return pd.read_excel(xlsx_path, sheet_name=sheet_name, dtype="object")


def write_new_sheet(xlsx_path: str, sheet_name: str, df: pd.DataFrame) -> None:
    """
    Create a new sheet (or new workbook if needed) with the given DataFrame.

    Notes:
        - Opens workbook in append mode if file exists.
    """
    try:
        mode = "a" if Path(xlsx_path).exists() else None
        if mode:
            with pd.ExcelWriter(xlsx_path, engine="openpyxl", mode="a") as w:
                df.to_excel(w, sheet_name=sheet_name, index=False)
        else:
            with pd.ExcelWriter(xlsx_path, engine="openpyxl") as w:
                df.to_excel(w, sheet_name=sheet_name, index=False)
    except Exception as e:
        print(f"❌ Failed to write new sheet '{sheet_name}': {e}")
        raise


def replace_sheet(xlsx_path: str, sheet_name: str, df: pd.DataFrame) -> None:
    """
    Replace (or create) a sheet with the provided DataFrame while keeping other sheets intact.

    Compatibility:
        - Uses if_sheet_exists='replace' when available.
        - Falls back to openpyxl removal for older pandas.
    """
    try:
        with pd.ExcelWriter(xlsx_path, engine="openpyxl", mode="a", if_sheet_exists="replace") as w:
            df.to_excel(w, sheet_name=sheet_name, index=False)
    except TypeError:
        try:
            from openpyxl import load_workbook
            wb = load_workbook(xlsx_path)
            if sheet_name in wb.sheetnames:
                ws = wb[sheet_name]
                wb.remove(ws)
                wb.save(xlsx_path)
            with pd.ExcelWriter(xlsx_path, engine="openpyxl", mode="a") as w:
                df.to_excel(w, sheet_name=sheet_name, index=False)
        except Exception as e:
            print(f"❌ Failed to replace sheet '{sheet_name}': {e}")
            raise
    except FileNotFoundError:
        # Workbook doesn't exist yet; create it
        try:
            with pd.ExcelWriter(xlsx_path, engine="openpyxl") as w:
                df.to_excel(w, sheet_name=sheet_name, index=False)
        except Exception as e:
            print(f"❌ Failed to create workbook '{xlsx_path}': {e}")
            raise
    except Exception as e:
        print(f"❌ Unexpected error replacing sheet '{sheet_name}': {e}")
        raise


def append_chunk_to_sheet(xlsx_path: str, sheet_name: str, chunk: pd.DataFrame) -> None:
    """
    Append a chunk to an existing sheet, widening columns if needed.
    If the sheet doesn't exist, create it with this chunk.

    Strategy:
        - Read existing sheet.
        - Union columns (preserve order).
        - Concat existing+chunk.
        - Replace the sheet atomically.

    Raises:
        Any exceptions from read/write are surfaced after logging.
    """
    try:
        if not sheet_exists(xlsx_path, sheet_name):
            write_new_sheet(xlsx_path, sheet_name, chunk)
            return

        existing = read_sheet_df(xlsx_path, sheet_name)
        all_cols = list(dict.fromkeys(list(existing.columns) + list(chunk.columns)))
        existing = existing.reindex(columns=all_cols)
        chunk = chunk.reindex(columns=all_cols)
        combined = pd.concat([existing, chunk], ignore_index=True)
        replace_sheet(xlsx_path, sheet_name, combined)
    except Exception as e:
        print(f"❌ Failed to append chunk to '{sheet_name}': {e}")
        raise


# ---------- API calls ----------

def commons_params(file_title: str) -> Dict[str, str]:
    """
    Build query parameters for a Commons file metadata request.

    Args:
        file_title: normalized 'File:...' title.

    Returns:
        dict: query parameters for action=query request.
    """
    return {
        "action": "query",
        "format": "json",
        "formatversion": "2",
        "titles": file_title,
        "redirects": "1",
        "prop": "imageinfo",
        "iiprop": "extmetadata|url|size|sha1|mime|mediatype|timestamp|user",
        "iiextmetadatalanguage": EXTMETA_LANG,
    }


def fetch_one(session: requests.Session, file_title: str) -> Tuple[Dict[str, Any], str]:
    """
    Fetch Commons metadata JSON for a single file.

    Args:
        session: configured HTTP session.
        file_title: normalized 'File:...' title.

    Returns:
        (data, effective_url): parsed JSON dict and the fully prepared request URL.

    Raises:
        requests.RequestException: for network/HTTP errors (with context).
    """
    try:
        req = requests.Request("GET", COMMONS_API, params=commons_params(file_title))
        prepped = session.prepare_request(req)
        resp = session.send(prepped, timeout=TIMEOUT_SECS)
        resp.raise_for_status()
        return resp.json(), (prepped.url or "")
    except requests.RequestException as e:
        # Attach title for context
        e.args = (f"{e.args[0] if e.args else e} [title={file_title}]",)
        raise


# ---------- Category harvest (supports optional RANGE) ----------

def harvest_category_to_sheet(
    session: requests.Session,
    category_title: str,
    xlsx_path: str,
    sheet_name: str,
    flush_rows: int = HARVEST_FLUSH_ROWS,
    index_start: Optional[int] = None,  # 1-based inclusive
    index_end: Optional[int] = None,    # 1-based inclusive
) -> None:
    """
    Harvest files from a Commons category (with continuation) and write/append to
    the input sheet ('CommonsFileName', 'SourceCategory'). Supports selecting a
    1-based inclusive RANGE within the category (e.g., 20..40).

    Args:
        session: configured HTTP session.
        category_title: 'Category:...' page to harvest.
        xlsx_path: workbook path.
        sheet_name: sheet for harvested items.
        flush_rows: write to Excel after this many rows buffered.
        index_start: first item index to include (1-based, inclusive).
        index_end: last item index to include (1-based, inclusive).

    Behavior:
        - Replaces the sheet header at the start to ensure a clean slate.
        - Stops early once the requested range is fully written.
        - Prints progress and totals.
    """
    if index_start is not None and index_end is not None and index_end < index_start:
        print(f"⚠️  Empty range ({index_start}..{index_end}); nothing to harvest.")
        replace_sheet(xlsx_path, sheet_name, pd.DataFrame(columns=["CommonsFileName", "SourceCategory"]))
        return

    print(f"Harvesting category: {category_title}"
          + (f" [range {index_start}..{index_end}]" if index_start or index_end else ""))

    # Start fresh sheet for this harvest
    replace_sheet(xlsx_path, sheet_name, pd.DataFrame(columns=["CommonsFileName", "SourceCategory"]))

    params = {
        "action": "query",
        "format": "json",
        "list": "categorymembers",
        "cmtitle": category_title,
        "cmtype": "file",
        "cmprop": "title",
        "cmlimit": str(CATEGORY_PAGE_LIMIT),
    }

    harvested_rows: List[Dict[str, Any]] = []
    total_seen = 0
    total_written = 0
    cont: Optional[Dict[str, str]] = None
    want_start = index_start or 1
    want_end = index_end or float("inf")

    try:
        while True:
            if cont:
                params.update(cont)
            resp = session.get(COMMONS_API, params=params, timeout=TIMEOUT_SECS)
            resp.raise_for_status()
            data = resp.json()
            members = (data.get("query") or {}).get("categorymembers") or []
            if not members:
                break

            page_first = total_seen + 1
            page_last = total_seen + len(members)

            # Overlap with requested range
            take_from = max(want_start, page_first)
            take_to = min(want_end, page_last)

            if take_from <= take_to:
                slice_start = take_from - page_first
                slice_end = take_to - page_first + 1  # inclusive → exclusive
                subset = members[slice_start:slice_end]
                for m in subset:
                    title = str(m.get("title") or "")
                    if not title:
                        continue
                    harvested_rows.append({"CommonsFileName": title, "SourceCategory": category_title})
                    total_written += 1

            total_seen = page_last

            # Flush periodically
            if harvested_rows and (len(harvested_rows) >= flush_rows):
                df_flush = pd.DataFrame(harvested_rows, columns=["CommonsFileName", "SourceCategory"])
                append_chunk_to_sheet(xlsx_path, sheet_name, df_flush)
                harvested_rows.clear()
                print(f"  • Harvested {total_written} within range; scanned {total_seen} items…")

            # Stop early when end reached
            if total_written >= (want_end - want_start + 1 if want_end != float('inf') else total_written + 1):
                break

            cont = data.get("continue")
            if not cont:
                break
    except requests.RequestException as e:
        print(f"❌ Category harvest failed: {e}")
        # Keep anything harvested so far; caller can inspect the partial sheet.
    except Exception as e:
        print(f"❌ Unexpected error during harvest: {e}")
        # Keep partial results.

    # Final flush
    try:
        if harvested_rows:
            df_flush = pd.DataFrame(harvested_rows, columns=["CommonsFileName", "SourceCategory"])
            append_chunk_to_sheet(xlsx_path, sheet_name, df_flush)
            harvested_rows.clear()
    except Exception as e:
        print(f"❌ Failed to finalize harvest writes: {e}")

    print(f"✅ Harvest complete: wrote {total_written} row(s) to '{sheet_name}'. (Category items scanned: {total_seen})")


# ---------- Processing (chunked) ----------

def process_input_sheet_chunked(
    session: requests.Session,
    xlsx_path: str,
    input_sheet: str,
    output_sheet: str,
    chunk_size: int,
) -> None:
    """
    Process rows from an input sheet in chunks and append results to an output sheet.

    Steps per file:
        - Normalize title to 'File:'.
        - Fetch Commons JSON (with redirects).
        - Derive MID and MID URL.
        - Save full JSON to disk (safe name, overwrite if identical).
        - Flatten JSON and build an output row with base columns + BatchIndex.
        - Append each chunk to the output sheet (widen columns as needed).

    Args:
        session: HTTP session.
        xlsx_path: workbook path.
        input_sheet: the sheet containing 'CommonsFileName' and optional 'SourceCategory'.
        output_sheet: the destination sheet for flattened metadata.
        chunk_size: number of rows per batch.

    Raises:
        FileNotFoundError / ValueError for input read failures.
        Other write errors will be logged and re-raised by helpers.
    """
    try:
        df_in = pd.read_excel(xlsx_path, sheet_name=input_sheet, dtype="string")
    except FileNotFoundError:
        raise FileNotFoundError(f"Input workbook not found: {xlsx_path}")
    except ValueError as e:
        raise ValueError(f"Input sheet '{input_sheet}' not found in {xlsx_path}: {e}")

    if "CommonsFileName" not in df_in.columns:
        raise KeyError(f"Input sheet '{input_sheet}' must contain 'CommonsFileName'.")
    if "SourceCategory" not in df_in.columns:
        df_in["SourceCategory"] = ""

    df_in["CommonsFileName"] = df_in["CommonsFileName"].fillna("").astype(str).str.strip()
    df_in["SourceCategory"] = df_in["SourceCategory"].fillna("").astype(str).str.strip()

    total = len(df_in)
    if total == 0:
        print(f"Nothing to process in '{input_sheet}'.")
        return

    # Ensure JSON dir exists
    try:
        DOWNLOAD_DIR.mkdir(parents=True, exist_ok=True)
    except Exception as e:
        print(f"❌ Cannot create JSON output directory '{DOWNLOAD_DIR}': {e}")
        raise

    num_batches = math.ceil(total / chunk_size)
    print(f"Processing {total} rows from '{input_sheet}' in {num_batches} batch(es) of {chunk_size}…")

    processed_so_far = 0
    for batch_index in range(num_batches):
        start = batch_index * chunk_size
        end = min(start + chunk_size, total)
        batch = df_in.iloc[start:end].copy()

        rows_out: List[Dict[str, Any]] = []
        for i, row in batch.iterrows():
            input_name = (row.get("CommonsFileName") or "").strip()
            source_cat = (row.get("SourceCategory") or "").strip()
            file_title = norm_file_title(input_name)

            print(f"[{i + 1}/{total}] Fetching {file_title or '<EMPTY>'} … ", end="", flush=True)

            if not file_title:
                rows_out.append({
                    "Input_CommonsFileName": input_name,
                    "SourceCategory": source_cat,
                    "Requested_API_URL": "",
                    "Local_JSON_File": "",
                    "Computed_MediaID": "",
                    "Computed_MediaID_URL": "",
                    "BatchIndex": batch_index + 1,
                })
                print("skipped (empty filename).")
                continue

            try:
                data, req_url = fetch_one(session, file_title)
            except requests.RequestException as e:
                rows_out.append({
                    "Input_CommonsFileName": input_name,
                    "SourceCategory": source_cat,
                    "Requested_API_URL": "",
                    "Local_JSON_File": "",
                    "Computed_MediaID": "",
                    "Computed_MediaID_URL": "",
                    "BatchIndex": batch_index + 1,
                    "error.message": str(e),
                })
                print(f"error: {e}")
                continue

            pageid, api_title = extract_pageid_title(data)
            mid = compute_mid(pageid)
            midlink = mid_url(mid)

            # Build JSON path and write (overwrite if same filename)
            try:
                json_path = build_safe_json_path(DOWNLOAD_DIR, input_name or api_title or file_title, mid)
                with json_path.open("w", encoding="utf-8") as f:
                    json.dump(data, f, ensure_ascii=False, indent=4, sort_keys=True)
            except Exception as e:
                # Keep going; record the error in output row
                print(f"⚠️  JSON write failed: {e}")
                json_path = Path("")

            flat = flatten_json(data)
            row_dict: Dict[str, Any] = {
                "Input_CommonsFileName": input_name,
                "SourceCategory": source_cat,
                "Requested_API_URL": req_url,
                "Local_JSON_File": str(json_path),
                "Computed_MediaID": mid,
                "Computed_MediaID_URL": midlink,
                "BatchIndex": batch_index + 1,
            }
            row_dict.update(flat)
            rows_out.append(row_dict)

            print(f"done ({'MID=' + mid if mid else 'MID=NOT FOUND'}).")

        # Build chunk DF with base columns first
        chunk_df = pd.DataFrame(rows_out)
        front_cols = [
            "Input_CommonsFileName",
            "SourceCategory",
            "Requested_API_URL",
            "Local_JSON_File",
            "Computed_MediaID",
            "Computed_MediaID_URL",
            "BatchIndex",
        ]
        cols = [c for c in front_cols if c in chunk_df.columns] + [c for c in chunk_df.columns if c not in front_cols]
        chunk_df = chunk_df.reindex(columns=cols)

        # Append/widen/replace output sheet
        append_chunk_to_sheet(xlsx_path, output_sheet, chunk_df)

        processed_so_far += len(batch)
        print(f"[Batch {batch_index + 1}/{num_batches}] Wrote {len(batch)} rows → '{output_sheet}' (total {processed_so_far}/{total}).")

    print(f"✅ Processing complete → '{output_sheet}'.")


# ---------- Orchestration ----------

def run_manual_list() -> None:
    """
    Execute the manual-list mode: read 'Files-Manual' and write metadata to 'FilesMetadata-Manual'.
    """
    session = build_session()
    process_input_sheet_chunked(
        session=session,
        xlsx_path=XLSX_PATH,
        input_sheet=INPUT_SHEET_MANUAL,
        output_sheet=OUTPUT_SHEET_MANUAL,
        chunk_size=CHUNK_SIZE,
    )


def run_category() -> None:
    """
    Execute the category mode:
      - Harvest 'CATEGORY_TITLE' to 'Files-Category' (optionally a 1-based range).
      - Process that sheet in chunks into 'FilesMetadata-Category'.
    """
    session = build_session()
    # Phase A: harvest category → Files-Category (incremental), honoring optional range
    harvest_category_to_sheet(
        session=session,
        category_title=CATEGORY_TITLE,
        xlsx_path=XLSX_PATH,
        sheet_name=INPUT_SHEET_CATEGORY,
        flush_rows=HARVEST_FLUSH_ROWS,
        index_start=CATEGORY_RANGE_START,
        index_end=CATEGORY_RANGE_END,
    )
    # Phase B: process harvested list in chunks → FilesMetadata-Category
    process_input_sheet_chunked(
        session=session,
        xlsx_path=XLSX_PATH,
        input_sheet=INPUT_SHEET_CATEGORY,
        output_sheet=OUTPUT_SHEET_CATEGORY,
        chunk_size=CHUNK_SIZE,
    )


def main() -> None:
    """
    Entry point. Reads MODE and runs the corresponding orchestration.
    """
    if MODE == "manual-list":
        print("Mode: manual-list")
        run_manual_list()
    elif MODE == "category":
        print("Mode: category")
        run_category()
    else:
        raise ValueError(f"Unknown MODE: {MODE!r}. Use 'manual-list' or 'category'.")


if __name__ == "__main__":
    # Ensure the workbook is closed before running.
    main()
