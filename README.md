# Automation Tool

Upload a spreadsheet of UK companies and get back, for each one, its turnover (or revenue),
prior-year turnover, cash and debtors from its latest annual accounts at Companies House, its
persons with significant control and a ready-made LinkedIn search, as an Excel file. It can also
search the same accounts for a list of keywords (foreign exchange, hedging, exports...) and count
how often each appears. Nothing to install for the people using it, and no API key needed.

## Using it

1. Open the app link.
2. Drop in your spreadsheet (`.xlsx`, `.xls`, `.xlsm` or `.csv`). It needs a column of company
   names, company numbers, or both. Numbers give exact matches. Not sure? Download the sample
   file from the app.
   Next to it is the keyword box, filled with a default list (foreign exchange, hedging,
   derivatives, overseas operations and so on). Edit it, paste your own (one per line), or clear
   it to look up turnover only.
3. Check that the app picked the right columns (it shows a preview), then press **Find turnover**.
4. Download the results: one row per company, with a **Summary** sheet.

## The results file

| Column | What it holds |
|---|---|
| Company Name, Company Number | As registered at Companies House (or as typed, if the company wasn't found) |
| Turnover Status | See below |
| Turnover or "Revenue" | This year's figure, whichever word the accounts use ("sales" counts too) |
| Prior Year "Turnover" or "Revenue" | Last year's figure, when the accounts make it unambiguous |
| Cash | Cash at bank and in hand (or cash and cash equivalents) at the balance sheet date |
| Debtors | Total debtors (or trade and other receivables) at the balance sheet date |
| Persons with significant control | The active PSCs, separated by semicolons; "none registered" if there are none |
| linkedin search | A link to a Google search for the company, its first individual PSC and "LinkedIn" |
| Keyword Count, Keywords Found | How many different keywords the accounts mention, and each one with its count |
| Notes | Anything worth knowing about the row (see below) |
| Latest Accounts | The accounts the figures came from, e.g. "Group of companies' accounts made up to 31 March 2025" |
| Accounts PDF, Companies House Page | Links |

Money columns use a £ accounting format. Small companies that don't file their turnover still
usually show cash and debtors on their balance sheet, so those columns are filled more often
than turnover.

## What the results mean

| Status | Meaning |
|---|---|
| FOUND | Read exactly from the filed accounts. |
| VERIFY | Read from a scanned PDF with OCR. The page is given; check the figure against the PDF. |
| NOT DISCLOSED | The company didn't file a turnover figure. Notes say why: micro-entity, small company that left out its profit and loss, dormant, medium company, or no accounts yet. |
| CHECK PDF | A figure couldn't be read automatically. Open the linked PDF. |
| NOT FOUND | The company couldn't be matched on the register. Add or check the company number. |
| ERROR | Companies House couldn't be reached for that row. Run it again later. |

The Notes column also flags: dissolved companies or ones in liquidation, accounts overdue at
Companies House, accounts more than 2 years old, accounts covering more or less than a year,
accounts in a currency other than sterling, figures read by OCR (with the page), names that
matched several companies, and a company number that belongs to a different name.

**Expect a lot of NOT DISCLOSED.** UK small and micro companies are allowed to leave the profit
and loss account out of the accounts they file, and most do. In one day of real filings, only
about 1 in 100 electronically filed accounts included turnover. Large companies always disclose
it, but they usually file scanned PDFs, which is why those are read with OCR and marked VERIFY.

## Keyword search

Each company's latest accounts are searched for every keyword. Two columns go next to the
turnover:

- **Keyword Count**: how many different keywords were found (sortable).
- **Keywords Found**: each keyword found with how many times it appears, most mentioned first,
  e.g. "hedging (16), foreign exchange (10), derivatives (7)".

"none found" means the accounts were searched and no keyword appeared; a blank means there was
no document to search (for example a new company with no accounts yet).

- Case and punctuation don't matter: "Foreign-Exchange" matches "foreign exchange", including
  when it's split across two lines.
- Only whole words and phrases count, so "FX" doesn't match inside "fixtures", "import"
  doesn't match "important", and "derivative" doesn't count "derivatives" (the default list has
  both forms as separate keywords).
- In scanned PDFs, OCR sometimes runs words together ("Foreignexchangerisk"). For that text
  only, phrases and longer words are also matched with the spaces ignored.
- Electronic accounts and text PDFs are searched in full at no extra cost. Scanned PDFs have to
  be read page by page with OCR (about 3 seconds a page), which is what makes big annual reports
  slow. Every page is read so the counts are complete, unless a page or time limit is set. The
  sidebar switch **Search scanned PDFs for keywords** turns this off; those companies are then
  noted as not searched.
- The default list lives in `default_keywords.txt`, one keyword per line, so it can be changed
  without touching the code.

## How it works

For each company it looks up the company page and its persons with significant control, finds
the latest annual accounts in the filing history (skipping interim accounts), and reads them,
best source first:

1. **Electronic accounts (iXBRL):** turnover, cash and debtors are tagged, so they're exact.
   For group accounts it takes the consolidated figures.
2. **Untagged figures** in electronic accounts, read as text from the income statement and
   balance sheet. This includes filings converted from PDF, where every word is placed by its
   position on the page; those rows are rebuilt line by line.
3. **The PDF:** text PDFs are read directly; scanned PDFs are read with OCR. It finds the
   income statement and balance sheet pages (skipping notes to the accounts), works out the
   units (£, £000, £m) from the column headings, and spots non-sterling accounts.

Every step has a fallback. If the electronic version won't download it reads the PDF, if a
full set of accounts has no tagged turnover it reads the PDF, network errors and Companies House
server errors are retried, and any problem with one company only affects that row.

## Speed

Companies House allows about one request a second. Measured on a laptop, October 2026:

| Kind of company | Time each |
|---|---|
| Small company, new company, not found (with or without keywords) | 1-5 seconds |
| Large company with a scanned annual report (100-280 pages) | 1-4 minutes, sometimes up to 6 |
| Same, also searching it for keywords | about 3 seconds a page: Whitbread (110 pages) 5.6 minutes, Tesco (all 236 pages) 12.7 minutes |

The 10-company sample file takes under a minute, keywords included. 100 typical small companies take about 8
minutes. On free hosting, scanned reports take roughly twice as long.

By default there are no limits: every page of every PDF is read. Sidebar settings to go faster:

- **Time limit per company (minutes)**, 0 = no limit. Stops reading a PDF when the time is up.
- **Max pages per PDF**, 0 = no limit. Reads at most that many pages (starting a quarter of the
  way in, where the financial statements usually are).
- **Search scanned PDFs for keywords** off: scanned reports are only read until the turnover is
  found, as without keywords.
- **Read scanned PDFs** off: scanned reports aren't read at all (marked CHECK PDF).

Whenever a limit stops the reading early, the row says how many pages were read, and a turnover
that wasn't reached in time is marked CHECK PDF rather than NOT DISCLOSED.

While a run is going, a **Download the N done so far** button appears (updated every 5
companies). Clicking it downloads the finished rows without stopping the run.

Keep the browser tab open while it runs: switching tabs or windows is fine, but closing the tab,
the laptop sleeping or losing the connection ends the run. If a run is interrupted by changing a
setting, the results finished so far can still be downloaded.

## Hosting it (free)

1. Put this folder in a GitHub repository (it can be private).
2. Go to [share.streamlit.io](https://share.streamlit.io), sign in with GitHub, click
   **Create app**, pick the repository, and set the main file to `app.py`.
3. Under **Advanced settings**, choose **Python 3.12**. The scanned-PDF reader needs 3.12 or
   earlier. On a newer Python the app still works, but scanned PDFs are marked CHECK PDF and the
   sidebar says so.
4. Deploy and share the link. Each person who opens it gets their own private session: uploads
   and results are never stored or shared.

Free apps go to sleep after a few days without visitors; the next visitor clicks "wake up" and
waits about 30 seconds.

## Running it on your own computer

```bash
pip install -r requirements.txt
streamlit run app.py
```

Or from the command line, without the web page:

```bash
python turnover.py clients.xlsx --name-col "Company Name" --number-col "Company Number"
```

Options: `--header-row 3` if the headings aren't in row 1, `--sheet "Sheet2"`, `--no-ocr` to skip
scanned PDFs, `--max-pages 100` and `--time-limit 10` (minutes per company; both 0 = no limit),
`--keywords-file default_keywords.txt` or
`--keywords "FX, hedging"` to search for keywords, `--no-keyword-ocr` to search only electronic
accounts and text PDFs.

## Tests

```bash
pip install -r requirements.txt -r requirements-dev.txt
python -m pytest                          # 224 offline tests, about 25 seconds
RUN_LIVE=1 python -m pytest tests/test_live.py -v    # 10 checks against the real site, about 8 minutes
```

The offline tests use saved Companies House pages, real electronic filings and generated PDFs
(`tests/fixtures`), and cover every route: each document type, every fallback, network failures,
odd spreadsheets (title rows, formulas, `.xls`, semicolon CSVs, numbers Excel has mangled), and
the web app itself.

## Files

| File | What it is |
|---|---|
| `app.py` | The web page |
| `turnover.py` | Finding companies, filings and turnover; writing the Excel results |
| `inputs.py` | Reading uploaded spreadsheets and guessing the columns |
| `default_keywords.txt` | The keyword list the app starts with, one per line |
| `sample_companies.xlsx` | A sample list (company name and number): 10 real companies, mostly importers, exporters and travel firms with foreign exchange exposure, plus two domestic businesses and a micro company for contrast |
| `tests/` | The test suite and its fixtures |

