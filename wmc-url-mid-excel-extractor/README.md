# Wikimedia Commons URL M-ID Excel Extractor

Reads a Wikimedia Commons **FileURL** column from an Excel sheet, looks up the corresponding **MediaInfo entity IDs (M-IDs)**, and writes the results back **into the same workbook**, adding/updating two columns next to the FileURL column:

* `FileMid` — the MediaInfo identifier in the form `M{pageid}`
* `FileMidURL` — the human-readable entity page URL, e.g. `https://commons.wikimedia.org/wiki/Special:EntityPage/M12345`

The script preserves all other sheets in the workbook and replaces only the target sheet’s contents.

*Latest update*: 17 October 2025

--------------

## Features

* Robust URL parsing for multiple Commons URL shapes:

  * `/wiki/File:…`
  * `?title=File:…`
  * `/wiki/Special:FilePath/…`
  * `/wiki/Special:Redirect/file/…`
* Batched API requests (≤ 50 titles/request) with retries & exponential backoff
* Redirect + normalization handling, so titles resolve to the correct page
* One output per input row; unresolved lookups yield `NOT FOUND` in both columns
* Detailed error log written to a CSV (`errors.csv`)

---

## How it works (high level)

1. Reads the Excel workbook and sheet containing a `FileURL` column.
2. Extracts a normalized `File:Title.ext` from each URL.
3. Queries the Commons API (`prop=info`) in **batches** to get the page ID.
4. Converts `pageid` → `M{pageid}` and builds the entity URL.
5. Inserts/updates `FileMid` and `FileMidURL` **right after** `FileURL`.
6. Writes back **in place** to the same workbook (only the specified sheet is replaced).
7. Writes an `errors.csv` with any failures (including “not found”).

---

## Requirements

* Python **3.9+**
* Packages: `pandas`, `openpyxl`, `requests`, `urllib3`

Install:

```bash
pip install -U pandas openpyxl requests urllib3
# or
pip install -r requirements.txt
```

`requirements.txt` example:

```
pandas>=1.5
openpyxl>=3.1
requests>=2.31
urllib3>=2.0
```

---

## Configuration

Edit the constants at the top of the script:

```python
XLSX_PATH = "testfile.xlsx"   # same file used for reading and writing
SHEET_NAME = "FileURLs"        # sheet with the FileURL column
URL_COLUMN = "FileURL"         # input column name

MID_COLUMN = "FileMid"         # output column 1
MID_URL_COLUMN = "FileMidURL"  # output column 2

ERRORS_CSV = "errors.csv"      # error log path
BATCH_SIZE = 50                # <= 50 for non-bot requests

USER_AGENT = "WikiCommons-MID-Extractor/1.0 (contact: KB, national library of the Netherlands - olaf.janssen@kb.nl)"
```

> **API etiquette:** Put a real contact in `USER_AGENT` (email or page) per Wikimedia guidelines.

---

## Usage

1. **Close the Excel file** (it cannot be open while writing).
2. Run:

```bash
python mid_extractor.py
```

3. Reopen the workbook. On `SHEET_NAME`, you’ll see `FileMid` and `FileMidURL` inserted immediately after `FileURL`.

---

## Input & Output

**Input sheet (minimal):**

| FileURL                                                                                                                              |
| ------------------------------------------------------------------------------------------------------------------------------------ |
| [https://commons.wikimedia.org/wiki/File:Example.jpg](https://commons.wikimedia.org/wiki/File:Example.jpg)                           |
| [https://commons.wikimedia.org/w/index.php?title=File:Another.jpg](https://commons.wikimedia.org/w/index.php?title=File:Another.jpg) |
| [https://commons.wikimedia.org/wiki/Special:FilePath/Third.png](https://commons.wikimedia.org/wiki/Special:FilePath/Third.png)       |

**Output sheet (same workbook, same sheet name):**

| FileURL                      | FileMid   | FileMidURL                                                                                                                     |
| ---------------------------- | --------- | ------------------------------------------------------------------------------------------------------------------------------ |
| …/File:Example.jpg           | M123456   | [https://commons.wikimedia.org/wiki/Special:EntityPage/M123456](https://commons.wikimedia.org/wiki/Special:EntityPage/M123456) |
| …title=File:Another.jpg      | NOT FOUND | NOT FOUND                                                                                                                      |
| …/Special:FilePath/Third.png | M987654   | [https://commons.wikimedia.org/wiki/Special:EntityPage/M987654](https://commons.wikimedia.org/wiki/Special:EntityPage/M987654) |

**Error log (`errors.csv`):**

| FileURL                 | Error                         |
| ----------------------- | ----------------------------- |
| …title=File:Another.jpg | Page missing or lookup failed |

---

## Troubleshooting

* **“Permission denied” / file locked**: Make sure the Excel file is closed.
* **`URL_COLUMN` not found**: Check the exact column name and sheet name.
* **Lots of `NOT FOUND`**: Verify the URLs point to **file pages** on Commons and parse into `File:…`.
* **HTTP 429/5xx**: The script retries with backoff; if it persists, slow down or split the input.

---

## Notes & Limitations

* The M-ID is derived from the MediaWiki **pageid** for the file page (`prop=info`), then formatted as `M{pageid}`.
* `FileMidURL` points to the human-readable entity page. If you prefer machine-readable JSON, switch to `https://commons.wikimedia.org/wiki/Special:EntityData/{mid}.json`.
* Batching is capped at 50 titles/request for non-bot clients (per MediaWiki limits).

---

## Contact & Credits

<img src="../media/icon_kb2.png" width="200" style="margin:4px 10px 0px 20px;" align="right"/>

* Author: Olaf Janssen, Wikimedia coordinator [@ KB, National Library of the Netherlands](https://www.kb.nl)
* Contact via [KB expert page](https://www.kb.nl/over-ons/experts/olaf-janssen) or [Wikimedia user page](https://commons.wikimedia.org/wiki/User:OlafJanssen).

---

## Licensing

<img src="../media/icon_cc0.png" width="100" style="4px 10px 0px 20px;" align="right"/>

Released into the public domain under [CC0 1.0 public domain dedication](LICENSE). Feel free to reuse and adapt. Attribution *(KB, National Library of the Netherlands)* is appreciated but not required.

