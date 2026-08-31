# pdf_search.py

A command-line tool for bulk-scanning a folder of PDFs to extract embedded
links, URLs found in visible text, keyword matches, and (optionally) Bates
numbers. Results are aggregated and written to a formatted Excel workbook,
with large reference lists offloaded to companion JSON files.

Useful for document review, e-discovery, and auditing large sets of PDFs for
specific links or terms without opening each file by hand.

## Capabilities

- **Single-file or folder scanning**: by default, processes one PDF file.
  Pass `--folder` to instead recursively scan a folder (and all subfolders)
  for `.pdf` files.
- **Link annotation extraction** (`--link-annotations`): pulls embedded
  hyperlinks (URI annotations) from each PDF.
- **Text URL extraction** (`--text-urls`): scans visible page text for
  URL-like strings using a regex, even if they aren't embedded as clickable
  links.
- **Keyword search** (`--keywords`, `--keywords-file`): case-insensitive
  search for one or more literal terms across the visible text of every page.
- **Bates numbering — footer** (`--bates-footer-prefix`): detects Bates
  numbers in the bottom-right corner of each page matching a given prefix
  (e.g. `MyCompany_000123`), and attaches the detected ID to every match
  found on that page.
- **Bates numbering — body** (`--bates-body-prefix`): detects Bates numbers
  anywhere in a page's visible text/content (not just the footer), useful
  when stamps aren't confined to the bottom-right corner. Can be used
  together with `--bates-footer-prefix` to capture both independently.
- **Context extraction**: each match includes surrounding text
  (`--context-window` characters before/after) to help you see it in
  context without opening the source PDF.
- **De-duplication & grouping**: matches are grouped by exact URL/keyword
  across the whole document set, with a reference count and a list of every
  file/page/context where it occurred.
- **Excel output**: results are written to an auto-formatted `.xlsx` file
  (bold headers, wrapped text, auto-sized columns), with a `URLs` sheet and/or
  `Keywords` sheet depending on which options were used.
- **Overflow handling**: if a match's reference list is too large to fit
  cleanly in a cell, it's saved to a separate JSON file in a
  `references_json/` subfolder, and the cell instead shows a pointer like
  `[See somefile.json]`.
- **Non-destructive output**: each run gets its own uniquely-numbered folder
  under `output/` (e.g. `output/run`, `output/run_1`, `output/run_2`, ...), so
  repeated runs never overwrite previous results.

## Installation

Requires Python 3.8+.

```bash
python3 -m venv venv
source venv/bin/activate   # on Windows: venv\Scripts\activate
pip install -r requirements.txt
```

Dependencies (see [requirements.txt](requirements.txt)):

- [PyMuPDF](https://pymupdf.readthedocs.io/) (`fitz`) — PDF parsing
- [pandas](https://pandas.pydata.org/) — data aggregation
- [tqdm](https://github.com/tqdm/tqdm) — progress bar
- [openpyxl](https://openpyxl.readthedocs.io/) — Excel output/formatting

## Usage

```bash
python3 pdf_search.py <path> [options]
```

`<path>` is the only required argument. By default it's treated as a single
PDF file. Pass `--folder` to instead treat it as a folder, which will be
scanned recursively for `.pdf` files. You must also pass at least one of
`--link-annotations`, `--text-urls`, or `--keywords`/`--keywords-file` to get
any output; the script does nothing on its own.

### Options

| Flag | Description |
|---|---|
| `path` | (positional, required) Path to a single PDF file (default), or a folder when `--folder` is set. |
| `--folder` | Treat `path` as a folder and recursively scan it for PDF files, instead of a single PDF file. |
| `--output PATH` | Output Excel filename. Default: `pdf_extraction_output.xlsx`. The file is placed inside an auto-generated run folder under `output/`. |
| `--link-annotations` | Extract embedded link annotations (clickable URIs) from each PDF. |
| `--text-urls` | Extract URL-like strings found in the visible page text. |
| `--keywords WORD [WORD ...]` | One or more keywords/phrases to search for (case-insensitive, literal match). Can be combined with `--keywords-file`. |
| `--keywords-file PATH` | Path to a text file with one keyword per line; appended to `--keywords`. |
| `--bates-footer-prefix PREFIX` | Prefix used to identify Bates numbers in the bottom-right footer of each page (e.g. `MyCompany` matches `MyCompany_000123`). Adds a `Bates ID (Footer)` column to results. |
| `--bates-body-prefix PREFIX` | Prefix used to identify Bates numbers anywhere in a page's visible text/content, not just the footer. Adds a `Bates ID (Body)` column to results. Can be combined with `--bates-footer-prefix`. |
| `--context-window N` | Number of characters of surrounding text to include before/after each match in the `Context` column. Default: `100`. |

### Examples

Extract every embedded link and text-based URL from a single PDF:

```bash
python3 pdf_search.py document.pdf --link-annotations --text-urls
```

Scan an entire folder of PDFs for specific keywords, with footer Bates number tagging:

```bash
python3 pdf_search.py ./docs --folder --keywords "confidential" "settlement" --bates-footer-prefix ACME
```

Search using a keyword list file, with a wider context window:

```bash
python3 pdf_search.py ./docs --folder --keywords-file terms.txt --context-window 200
```

Tag matches with both footer and body Bates numbers (useful when stamps aren't
confined to the bottom-right corner):

```bash
python3 pdf_search.py ./docs --folder --keywords-file terms.txt \
  --bates-footer-prefix ACME \
  --bates-body-prefix ACME
```

Do everything at once, over a folder, with a custom output filename:

```bash
python3 pdf_search.py ./docs --folder \
  --link-annotations \
  --text-urls \
  --keywords-file terms.txt \
  --bates-footer-prefix ACME \
  --output review_results.xlsx
```

## Output

Each run creates a new folder under `output/` (e.g. `output/review_results`,
incrementing with `_1`, `_2`, ... if the name is already taken), containing:

- **`<output>.xlsx`** — the results workbook:
  - **`URLs` sheet** (if `--link-annotations` and/or `--text-urls` used):
    one row per unique URL, with `Base URL` (domain), `Match Type`
    (`link` or `text_url`), `Reference Count`, and `References` (list of
    every filename/page/context where it appeared, plus `Bates ID (Footer)`
    and/or `Bates ID (Body)` if those options were used).
  - **`Keywords` sheet** (if keywords used): one row per unique matched
    keyword, with `Match Type`, `Reference Count`, and `References` (same
    structure as above).
- **`references_json/`** — JSON files for any match whose full reference
  list was too large to embed directly in the spreadsheet cell; the cell
  will contain a `[See <file>.json]` pointer instead.

If no links or keywords are found, no Excel file is written and the script
prints a message to that effect.
