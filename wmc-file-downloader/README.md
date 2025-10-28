# Wikimedia Commons File Downloader

*A robust, Windows-safe downloader for Wikimedia Commons files*

Features include:

- **Two modes**: download entire Commons categories up to a chosen `DEPTH`, or read file(URLs) from a  a single-column list.
- **File summary preview**: nested, alphabetically sorted category tree showing potential, non-deduplicated Commons candidates vs. eligible, unique files to be downloaded.
- **Slicing**: after building the full unique file pool across all selected Commons categories, download only a subset.
- **Safe filenames**: always suffix `--MID` (or `--NO-MID-<hash>`), with Windows file system path-length control to avoid too long file paths.
- **Incremental downloads log** in Excel, chunked, deduped, safe to re-run.

*Latest update*: 28 October 2025

---

## Requirements

* Python **3.9+** (3.10/3.11 recommended)
* The packages listed in `requirements.txt`

### Install dependencies

Use the list of packages in `requirements.txt` (make sure it includes at least these):

```
pandas>=1.5.0
openpyxl>=3.1.0
requests>=2.31.0
urllib3>=2.0.0
```

If not already done, you can create a virtual environment and install the required packages:

**macOS / Linux**
```bash
python3 -m venv .venv  # Create virtual environment in the .venv folder (optional, of not already done)
source .venv/bin/activate # Activate the venv 
python -m pip install -r requirements.txt # Install the Python packages in the venv
```

**Windows (PowerShell)**
```powershell
py -m venv .venv  # Create virtual environment in a .venv folder (optional, of not already done)
.venv\Scripts\Activate.ps1 # Activate the venv 
python -m pip install -r requirements.txt # Install the Python packages in the venv
```

---

## Configuration (edit in the script)

All knobs live near the top in the ```Configuration parameters``` block. Key settings are shown here with defaults:

```python
# --- Mode selection ---
MODE = "categories"                  # or "list"

# --- Shared settings ---
FILE_EXTS  = tuple()                 # () = no filtering; or e.g. (".jpg",".png",".tiff")
UA = "Wikimedia Commons File Downloader by User:YourWikiUserName (yourname@email.tld)"
LOCAL_BASE_FOLDER = "dwnlds"         # short string to help avoid Windows path length issues

EXCEL_LOG_NAME = "downloads_log.xlsx" # admin logs for all downloaded files
EXCEL_LOG_DIR  = "./"                # "./" = next to python script; None = inside LOCAL_BASE_FOLDER
SHEET_NAME_CATEGORIES = "CategoriesDownloads" # Output Excel log file sheet name for MODE = "categories"
SHEET_NAME_LIST       = "ListDownloads" # sheet name for MODE = "list"

# --- MODE=categories ---
CATEGORIES = ["Media contributed by Koninklijke Bibliotheek"] # or ["cat1", "cat2", "cat3"]
DEPTH = 0
CATEGORIES_DOWNLOAD_SUBFOLDER = "cats"
FLATTEN_CATEGORIES_PATHS = True # To avoid deeply nested and thus potentially too long sub(sub(sub(sub))etc) folder trees in local file system

# Slice WHICH ROOT CATEGORIES you harvest (1-based) — affects the CATEGORIES list itself
CATEGORIES_RANGE_START = None # or =2 to work with cats 2-5 in CATEGORIES[]
CATEGORIES_RANGE_END   = None # or =5 to work with cats 2-5 in CATEGORIES[]

# Slice WHICH FILES you download (1-based), after building the global UNIQUE pool. 
# First 5 files from global file pool: 
CATEGORIES_GLOBAL_RANGE_START = 1 # set to None for all files
CATEGORIES_GLOBAL_RANGE_END   = 5 # set to None for all files
CATEGORIES_GLOBAL_SLICE_ORDER = "root_then_title"   # or "title"

# --- MODE=list ---
LIST_INPUT_PATH   = "list-of-tobe-downloaded-files.xlsx"  # .xlsx/.csv/.tsv/.txt
LIST_INPUT_FORMAT = "excel"                                # "auto"|"excel"|"csv"|"tsv"|"txt"
LIST_EXCEL_SHEET  = "FilesList" # name of input Excel sheet - only used when LIST_INPUT_FORMAT="excel"
LIST_TEXT_DELIM   = None

LIST_DOWNLOAD_SUBFOLDER = "list"
# Download first 5 files from list:
LIST_RANGE_START        = 1 # set to None for all files
LIST_RANGE_END          = 5 # set to None for all files
```

> **Tip:** If you want to include specific formats only (e.g., just images), set:
> `FILE_EXTS = (".jpg",".jpeg",".png",".gif",".tif",".tiff",".webp",".svg")`.
> To include PDFs: add `".pdf"`. To include *everything*: keep `FILE_EXTS = tuple()`.

---

## How to decide what to download

### MODE="categories"

1. Harvest each root category in `CATEGORIES[]` and its subcats to level `DEPTH` using Breadth-First Search (BFS, a way to traverse a tree/graph level by level).
2. Build **RAW** counts (Commons *direct* membership per category; totals include non-deduplicated descendants/children. Non-deduplicated = one single file can be in more than one Commons category.
3. Build **UNIQUE** pool (post-`FILE_EXTS` filtering and de-duplicated across the file tree; the actual set of unique files to be downloaded).
4. Optionally apply a **global slice** across this UNIQUE pool, choose a subset from all files in all root categories combined.
5. Download & log incrementally.

> Folder layout on disk:
> ```
> LOCAL_BASE_FOLDER /                # 'dwnlds'
>   CATEGORIES_DOWNLOAD_SUBFOLDER /  # 'cats'
>     <RootCategory> /               # 'Magazines_from_Koninklijke_Bibliotheek'
>       (possibly flattened subcats per FLATTEN_CATEGORIES_PATHS)
>         LocalFilename.ext
> ```

### MODE="list"

1. Load a single-column list file (column name = `LIST_SINGLE_COLUMN_NAME`).
2. Each cell may be one of:
   * Commons `File:Title`
   * Commons File URL (`https://commons.wikimedia.org/wiki/File:Title`)
   * Commons MediaInfo ID (`M12345`)
   * Commons concept/entity URL (`https://commons.wikimedia.org/entity/M12345` or `https://commons.wikimedia.org/wiki/Special:EntityData/M12345`)
   * Commons direct upload URL (`https://upload.wikimedia.org/...`)
3. All these variations resolve to `File:Title` (+ MID and upload URL), de-dup by title, then download.

> Folder layout on disk:
> ```
> LOCAL_BASE_FOLDER /          # 'dwnlds'
>   LIST_DOWNLOAD_SUBFOLDER /  # 'list'
>     LocalFilename.ext
> ```

---

## Running

> Close the Excel log file `EXCEL_LOG_NAME` if it’s open in Excel before running.

Then run

```bash
python wmc-file-downloader.py
```

(or run it from with your IDE, such as PyCharm)

* In **categories mode**, a **Download plan summary** prints a nested, A→Z tree of (numbers of) files to be downloaded:

  ```
  • Category:Root → Raw non-unique candidates=…, Unique & eligible files=…
    - Root/Sub1 → Candidates: direct=.., total=.. | Eligible: direct=.., total=.. | Skipped files already present elsewhere in the category tree=.. | This category is on depth level 1
      - Root/Sub1/Sub2 → …
  ```
* If you configured a **global slice**, you’ll see:

  ```
  GLOBAL SLICE: selecting items X–Y of Z UNIQUE eligible files (order=root_then_title).
    - Category:RootA → n file(s) selected
    - Category:RootB → m file(s) selected
  ```
* Confirm to proceed (unless you disabled `CONFIRM_BEFORE_DOWNLOAD`).

---

## Excel download log

* **Location**: `EXCEL_LOG_DIR` / `EXCEL_LOG_NAME` (default: `downloads_log.xlsx`)
* **Sheets**: `SHEET_NAME_CATEGORIES` or `SHEET_NAME_LIST`
* **Dedup key**: (`CommonsFileURL`, `MediaInfoID`)
* **Columns (left → right)**
  - `CommonsFileName` — canonical `File:Title`
  - `CommonsFileURL` — `https://commons.wikimedia.org/wiki/File:Title`
  -  `MediaInfoID` — e.g., `M12345`
  -  `CommonsConceptURI` — `https://commons.wikimedia.org/entity/M12345` (or empty)
  -  `CommonsImageURL` — final `https://upload.wikimedia.org/...` URL
  -  `SourceCategory` 
     - For list mode: as provided via `LIST_SOURCE_COLUMN` and `LIST_SOURCE_TAG` 
     - For categories mode: as provided by top level Commons category in `CATEGORIES[]` 
  -  `CommonsCategoryPath` — full original Commons category/subcategory chain (for categories mode)
  -  `LocalBaseFolder` — as set by `LOCAL_BASE_FOLDER`
  -  `LocalSubFolder` — relative subfolder on disk (may be flattened to avoid too long/deep local file trees)
  -  `LocalFilename` — local stored filename, always suffixed with `--MID` or `--NO-MID-<hash>`. Length depending on remaining number of characters within `FULL_PATH_BUDGET`  

The log is written incrementally in batches of `LOG_FLUSH_ROWS` rows, so progress is mostly saved even if a run is interrupted. Set small batches for frequent saving.

---

## Filename & path safety

* Illegal characters (on Windows filesystem, such as `['"', '#', '%', '&', '{', '}', '\\', '<', '>', '|', ':', '*', '?', '/']`) that might occur in Commons file names are replaced by `'_'` and trailing spaces/dots stripped.
* The locally downloaded filename always ends with `--MID` or `--NO-MID-<hash>` so truncated names remain unique.
* The absolute path (tested on Window file systems) respects `FULL_PATH_BUDGET`. To avoid too long paths:

  * Keep `LOCAL_BASE_FOLDER` (default = `"dwnlds"`) and `CATEGORIES_DOWNLOAD_SUBFOLDER` (default = `"cats"`) and root names short. 
  * Set `FLATTEN_CATEGORIES_PATHS=True` (default) to avoid too deep subfolder trees.

---

## Troubleshooting

You might encounter one or more of the following issues: 

* **File counts reported in the output of the tool don’t seem to match with those reported in Commons category pages**:
  * Category pages on Wikimedia Commons always show RAW direct membership. One single file can be in more than one Commons category.
  * The script’s plan shows both RAW and UNIQUE; UNIQUE is post-filter + de-dup (what you’ll download).
* **“Could not resolve row; skipping.” (list mode)**:
  * Ensure your input has ONE column named exactly `LIST_SINGLE_COLUMN_NAME` (default `CommonsInput`).
  * You can paste File titles, MIDs, concept URLs, or upload URLs — any mix is fine.
* **`CommonsImageURL` column in downloads Excel is empty**:
  * If `Special:FilePath` didn’t yield the final upload host in the HTTP response, the script queries the API to resolve it. Check connectivity 
   and that the title exists.
* **Excel file seems to append duplicates**:
  * The log de-duplicates on (`CommonsFileURL`, `MediaInfoID`). If both are empty for some reason, rows can’t be deduped; verify your inputs resolve correctly.
* **Windows full file path too long**:
  * Lower `FULL_PATH_BUDGET`, 
  * Keep `LOCAL_BASE_FOLDER` (default = `"dwnlds"`) and `CATEGORIES_DOWNLOAD_SUBFOLDER` (default = `"cats"`) short. 
  * Set `FLATTEN_CATEGORIES_PATHS=True` (default) to avoid too deep subfolder trees.

---

## Licensing

<img src="../media/icon_cc0.png" width="100" style="4px 10px 0px 20px;" align="right"/>

Released into the public domain under [CC0 1.0 public domain dedication](LICENSE). Feel free to reuse and adapt. Attribution *(KB, National Library of the Netherlands)* is appreciated but not required.

## Contact & Credits

<img src="../media/icon_kb2.png" width="200" style="margin:4px 10px 0px 20px;" align="right"/>

* Author: Olaf Janssen, Wikimedia coordinator [@ KB, National Library of the Netherlands](https://www.kb.nl)
* Contact via [KB expert page](https://www.kb.nl/over-ons/experts/olaf-janssen) or [Wikimedia user page](https://commons.wikimedia.org/wiki/User:OlafJanssen).
* User-Agent: `"Wikimedia Commons File Downloader by User:OlafJanssen (olaf.janssen@kb.nl)"`