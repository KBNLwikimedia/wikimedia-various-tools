# Wikimedia Commons File Metadata Downloader

*A tool to collect metadata from Wikimedia Commons files or categories and write them into an Excel sheet — safely, in chunks, and with per-file JSON snapshots.*

This tool has two operation modes:
* **Manual list mode**: process a list of Commons file titles you manually provide.
* **Category mode**: harvest files from a Commons category (even huge ones), optionally only a range (e.g., items 20–40), then process them the same way.

The script saves the exact API response for each file as JSON into `downloaded_metadata/` (Windows-safe filenames, overwrite if identical) and appends flattened JSON columns to the Excel workbook as it progresses.

---

## What you’ll get

* **Input Excel file**: `wmc-inputfiles.xlsx`
* **Manual list mode**
  * Input sheet: `Files-Manual` (manually provided by the user)
  * Output sheet: `FilesMetadata-Manual`
  
* **Category mode**
  * Harvest sheet: `Files-Category` (built/updated by the script)
  * Output sheet: `FilesMetadata-Category`
* **Per-file JSON**: saved as JSON files to folder `downloaded_metadata/` with filename syntax `<CommonsFileName>__<MID-or-NOID>.json`

> The output sheets are updated incrementally in configurable batches (chunks). This will lower the risk of intermediate data losses.

---

## Prerequisites

* Python **3.9+** (3.10/3.11 recommended)
* An Excel workbook named `wmc-inputfiles.xlsx` (an example is provided in this repo)
* The packages listed in `requirements.txt`

### Install dependencies

Using your existing `requirements.txt` (make sure it includes at least these):

```
pandas>=1.5.0
openpyxl>=3.1.0
requests>=2.31.0
urllib3>=2.0.0
```

If not already done, you can create a virtual environment and install:

**macOS / Linux**
```bash
python3 -m venv .venv  # Create virtual environment in the .venv folder (optional, of not already done)
source .venv/bin/activate # Activate the venv 
python -m pip install -r requirements.txt # Install the Python packages in the venv
```

**Windows (PowerShell)**
```powershell
py -m venv .venv  # Create virtual environment in the .venv folder (optional, of not already done)
.\.venv\Scripts\Activate.ps1 # Activate the venv 
python -m pip install -r requirements.txt # Install the Python packages in the venv
```

---

## Configuration

Open the script `wmc_metadata.py` and edit the configuration block near the top. 

```python
# Comment out the mode you do NOT want to use:
#MODE = "manual-list" # Comment out if MODE = category
MODE = "category"  # In this mode, you must configure the CATEGORY_TITLE below.
CATEGORY_TITLE = "Category:Catchpenny prints from Koninklijke Bibliotheek" # Change to your own Commons category
# Optional range for CATEGORY mode (1-based inclusive). First 10 files in this category:
CATEGORY_RANGE_START: Optional[int] = 1  # Set to =None to process all
CATEGORY_RANGE_END:   Optional[int] = 10  # or =None to process all

USER_AGENT = "Wikimedia Commons File Metadata Downloader - User:YourWikiUserName - Contact:you@email.tld)" # Please set your own contact info here.

# Input Excel file - required in both modes
XLSX_PATH = "wmc-inputfiles.xlsx"
```
---

## Preparing the workbook

Create `wmc-inputfiles.xlsx` in the same folder as the script.

### 1) Manual-list mode

Create a sheet named `Files-Manual` with these columns:

| CommonsFileName  | SourceCategory *(optional)* |
| ---------------- |-----------------------------|
| File:Example.jpg | Category:My Project Images  |
| Example2.png     |                             |
| File:Another.svg | Category:Another Collection |

* `CommonsFileName` is required. If you forget the `File:` prefix, the script adds it.
* `SourceCategory` is optional (it’s copied into the output).

### 2) Category mode

No preparations needed for input: the script will create/update the `Files-Category` sheet with harvested items from `CATEGORY_TITLE`.

---

## Running

> Make sure `wmc-inputfiles.xlsx` is closed in Excel before running.

From your virtual environment (if used), run:

```bash
python wmc_metadata.py
```

Or run from IDEs like PyCharm, VSCode, etc.

The script prints progress, e.g.:
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
   * If the same filename is produced again, it is overwritten.
5. **Flatten JSON** into dotted columns and append to the output sheet in chunks.
   * If a chunk introduces new JSON keys, the output sheet’s columns are widened and replaced once, then appending continues.

---

## Understanding the two “batch sizes”

* `HARVEST_FLUSH_ROWS` *(category mode only)* – how many harvested file titles to buffer before appending them to the **`Files-Category`** input sheet during harvest.
* `CHUNK_SIZE` *(both modes)* – how many files to process per batch when fetching JSON and appending rows to the two `FilesMetadata-…` output sheets during processing.

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
* Overwrite: Yes (same name → overwritten)

---

## Troubleshooting

* `Input workbook not found`: Ensure `wmc-inputfiles.xlsx` exists in the repo folder.

* `Input sheet 'Files-Manual' not found`: Check sheet names in Excel match the config (`Files-Manual` / `Files-Category`).

* `Input sheet must contain 'CommonsFileName'`: Add this column header to your input sheet.

* *HTTP 429 or 5xx*: The script retries with backoff. If it persists, lower `CHUNK_SIZE` or add short sleeps (we can add this if needed).

* *Excel is locked*: Close the workbook in Excel before running (the script writes to it).

* *Paths too long on Windows*: The script shortens filenames automatically; if you still hit limits, lower `FULL_PATH_BUDGET` or run from a shorter folder path (e.g., `C:\w\`).

---

## Tips & etiquette

* The `USER_AGENT` includes a real contact (recommended for Wikimedia API use).
* For Dutch metadata, set `EXTMETA_LANG = "nl"`.
* If your category contains very many files, that’s fine — the script harvests with continuation and writes incrementally.
* Consider git-ignoring `downloaded_metadata/` if it grows large.

---

## Licensing

<img src="../media/icon_cc0.png" width="100" style="4px 10px 0px 20px;" align="right"/>

Released into the public domain under [CC0 1.0 public domain dedication](LICENSE). Feel free to reuse and adapt. Attribution *(KB, National Library of the Netherlands)* is appreciated but not required.

## Contact & Credits

<img src="../media/icon_kb2.png" width="200" style="margin:4px 10px 0px 20px;" align="right"/>

* Author: Olaf Janssen, Wikimedia coordinator [@ KB, National Library of the Netherlands](https://www.kb.nl)
* Contact via [KB expert page](https://www.kb.nl/over-ons/experts/olaf-janssen) or [Wikimedia user page](https://commons.wikimedia.org/wiki/User:OlafJanssen).
* User-Agent: `"Wikimedia Commons File Metadata Downloader - User:OlafJanssen - Contact: olaf.janssen@kb.nl)"`


