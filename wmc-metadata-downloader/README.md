# Wikimedia Commons File Metadata Downloader

*A practical tool to collect per-file metadata from Wikimedia Commons and write it into an Excel workbook — safely, in chunks, and with per-file JSON snapshots.*

It has two operation modes:
* **Manual list mode**: process a list of Commons file titles you provide.
* **Category mode**: harvest files from a Commons category (even huge ones), optionally only a **range** (e.g., items 20–40), then process them the same way.

The script saves the **exact JSON response** for each file into `downloaded_metadata/` (Windows-safe filenames, overwrite if identical) and appends **flattened JSON columns** to your Excel workbook as it progresses.

---

## What you’ll get

* **Input workbook**: `wmc-inputfiles.xlsx`
* **Manual list**

  * Input sheet: `Files-Manual`
  * Output sheet: `FilesMetadata-Manual`
* **Category**

  * Harvest sheet: `Files-Category` (built/updated by the script)
  * Output sheet: `FilesMetadata-Category`
* **Per-file JSON**: saved to `downloaded_metadata/<CommonsFileName>__<MID-or-NOID>.json`

> The output sheets are updated **incrementally** in batches (chunks), so you can stop and resume without losing progress.

---

## Prerequisites

* Python **3.9+** (3.10/3.11 recommended)
* An Excel workbook named **`wmc-inputfiles.xlsx`** (created in this repo folder)
* The packages listed in `requirements.txt`

### Install dependencies

Using your existing `requirements.txt` (make sure it includes at least these):

```
pandas>=1.5.0
openpyxl>=3.1.0
requests>=2.31.0
urllib3>=2.0.0
```

Create a virtual environment and install:

**macOS / Linux**

```bash
python3 -m venv .venv
source .venv/bin/activate
python -m pip install -r requirements.txt
```

**Windows (PowerShell)**

```powershell
py -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
```

---

## Configuration

Open the script (e.g., `wmc_metadata.py`) and edit the **configuration block** near the top. This block is included **verbatim** in the script:

```python
# Comment out the mode you do NOT want to use:
#MODE = "manual-list" # Comment out if MODE = category
MODE = "category"  # In this mode, you must configure the CATEGORY_TITLE below.
CATEGORY_TITLE = "Category:Catchpenny prints from Koninklijke Bibliotheek"

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
USER_AGENT = "KB WMC metadata fetcher - User:OlafJanssen - Contact: olaf.janssen@kb.nl)"

# Requests
TIMEOUT_SECS = 20
RETRIES_TOTAL = 5
RETRIES_BACKOFF = 0.6

# Windows path safety
FULL_PATH_BUDGET = 240  # conservative full-path length budget
```

### Optional: Category range (slice)

Just below the config block you can set a **1-based inclusive** slice for category mode:

```python
CATEGORY_RANGE_START = 20   # e.g., start at the 20th file
CATEGORY_RANGE_END   = 40   # …end at the 40th (inclusive)
```

Leave them as `None` to harvest the full category.

---

## Preparing the workbook

Create `wmc-inputfiles.xlsx` in the same folder as the script.

### Manual-list mode

Create a sheet named **`Files-Manual`** with these columns:

| CommonsFileName  | SourceCategory              |
| ---------------- | --------------------------- |
| File:Example.jpg | Category:My Project Images  |
| Example2.png     |                             |
| File:Another.svg | Category:Another Collection |

* `CommonsFileName` is **required**. If you forget the `File:` prefix, the script adds it.
* `SourceCategory` is optional (it’s copied into the output).

### Category mode

No prep needed for input: the script will **create/update** the **`Files-Category`** sheet with harvested items from `CATEGORY_TITLE`.

---

## Running

> Make sure `wmc-inputfiles.xlsx` is **closed** in Excel before running.

From your virtual environment:

```bash
python wmc_metadata.py
```

* The script prints progress, e.g.:

  * Harvest: `Harvested 500 within range; scanned 1500 items…`
  * Processing: `[123/8120] Fetching File:Example.jpg … done (MID=M123456)`
  * Batch writes: `[Batch 7/41] Wrote 100 rows → 'FilesMetadata-Category' (total 700/4100).`

---

## What the script does (both modes)

1. **Normalize titles** to `File:…`.
2. **Fetch metadata** from Commons (`prop=imageinfo`, `extmetadata`, `url`, `sha1`, `mime`, `timestamp`, etc.; with redirects handled).
3. **Compute MediaInfo ID (MID)** from pageid, and **MID URL**.
4. **Save the full JSON** to `downloaded_metadata/` using a Windows-safe name:
   `<CommonsFileName>__<MID or NOID>.json`

   * If too long: the filename is truncated and a short hash is added.
   * If the same filename is produced again, it is **overwritten**.
5. **Flatten JSON** into dotted columns and **append** to the output sheet in **chunks**.

   * If a chunk introduces new JSON keys, the output sheet’s columns are widened and replaced once, then appending continues.

---

## Understanding the two “batch sizes”

* **`HARVEST_FLUSH_ROWS`** *(category mode only)* – how many harvested file titles to buffer before appending them to **`Files-Category`** during **harvest**.
* **`CHUNK_SIZE`** *(both modes)* – how many files to **process per batch** when fetching JSON and appending rows to **`FilesMetadata-…`** during **processing**.

Smaller values = more frequent writes, faster visible progress, more I/O.
Larger values = fewer writes, more memory per batch.

---

## Outputs

### Output sheets

* **Manual-list** → `FilesMetadata-Manual`
* **Category** → `FilesMetadata-Category`

Each row includes:

* `Input_CommonsFileName`
* `SourceCategory`
* `Requested_API_URL`
* `Local_JSON_File` (path to the saved JSON)
* `Computed_MediaID` (e.g., `M12345`, can be empty)
* `Computed_MediaID_URL`
* `BatchIndex` (1-based)
* **All flattened JSON fields** (e.g., `query.pages.0.title`, `query.pages.0.imageinfo.0.url`, `query.pages.0.imageinfo.0.extmetadata.Artist.value`, …)

### JSON files

* Directory: `downloaded_metadata/`
* Name: `<CommonsFileName>__<MID or NOID>.json`
  (with truncation+hash if needed)
* Overwrite: **Yes** (same name → overwritten)

---

## Troubleshooting

* **`Input workbook not found`**
  Ensure `wmc-inputfiles.xlsx` exists in the repo folder.

* **`Input sheet 'Files-Manual' not found`**
  Check sheet names in Excel match the config (`Files-Manual` / `Files-Category`).

* **`Input sheet must contain 'CommonsFileName'`**
  Add this column header to your input sheet.

* **HTTP 429 or 5xx**
  The script retries with backoff. If it persists, lower `CHUNK_SIZE` or add short sleeps (we can add this if needed).

* **Excel is locked**
  Close the workbook in Excel before running (the script writes to it).

* **Paths too long on Windows**
  The script shortens filenames automatically; if you still hit limits, lower `FULL_PATH_BUDGET` or run from a shorter folder path (e.g., `C:\w\`).

---

## Tips & etiquette

* The `USER_AGENT` includes a real contact (recommended for Wikimedia API use).
* For Dutch metadata, set `EXTMETA_LANG = "nl"`.
* If your category has very many, that’s fine — the script harvests with continuation and writes **incrementally**.
* Consider **git-ignoring** `downloaded_metadata/` if it grows large.

---

## Licensing

<img src="../media/icon_cc0.png" width="100" style="4px 10px 0px 20px;" align="right"/>

Released into the public domain under [CC0 1.0 public domain dedication](LICENSE). Feel free to reuse and adapt. Attribution *(KB, National Library of the Netherlands)* is appreciated but not required.

## Contact & Credits

<img src="../media/icon_kb2.png" width="200" style="margin:4px 10px 0px 20px;" align="right"/>

* Author: Olaf Janssen, Wikimedia coordinator [@ KB, National Library of the Netherlands](https://www.kb.nl)
* Contact via [KB expert page](https://www.kb.nl/over-ons/experts/olaf-janssen) or [Wikimedia user page](https://commons.wikimedia.org/wiki/User:OlafJanssen).
* User-Agent: `"Wikimedia Commons File Metadata Downloader - User:OlafJanssen - Contact: olaf.janssen@kb.nl)"`


