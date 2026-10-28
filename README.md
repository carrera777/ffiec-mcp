# FFIEC CDR for Claude (Call Reports and UBPR)

A Claude Desktop extension that pulls bank regulatory data from the FFIEC Central Data Repository (CDR) Public Web Service: Call Reports (FFIEC 031, 041, 051) and Uniform Bank Performance Reports (UBPR). Ask Claude about any U.S. bank in plain English and it will look the bank up, pull its filings, and export formatted Excel workbooks.

It is read-only. It never files or submits anything.

## What you can ask for

- "Find First Horizon Bank's RSSD ID."
- "Summarize JPMorgan Chase Bank's latest Call Report."
- "Show net income and total deposits for RSSD 852218 for the last 8 quarters."
- "Export Wells Fargo and U.S. Bank Call Reports for 2024Q1 to 2026Q2 to Excel, with Schedule RC-C detail."
- "What is the UBPR code for net interest margin, and what was it for this bank last quarter?"
- "Which banks filed or amended Q2 2026 Call Reports since September 1?"

## Tools

| Tool | What it does |
|---|---|
| `ffiec_find_bank` | Search every bank that files a Call Report by name, RSSD ID, FDIC cert, ABA routing number, state, or city |
| `ffiec_reporting_periods` | List available Call Report or UBPR periods |
| `ffiec_recent_filers` | Banks that filed or amended since a given date, newest first |
| `ffiec_call_report` | A bank's Call Report line items, filtered by item code, schedule, or keyword |
| `ffiec_bank_summary` | Headline balance sheet, income, and capital figures for up to 25 banks |
| `ffiec_bank_trend` | Multi-quarter time series for one bank |
| `ffiec_download_call_report` | Save the filing as PDF, XBRL, or SDF |
| `ffiec_ubpr` | UBPR data points with plain-English names |
| `ffiec_lookup_items` | Look up item names by code, or codes by name (understands ROA, ROE, NIM, NCO, NPL, CET1) |
| `ffiec_export_excel` | Formatted Excel workbook with year-to-date and discrete-quarter income statements |
| `ffiec_import_bulk` | Load every bank's Call Report for a quarter from an FFIEC bulk-download ZIP, with no API calls |
| `ffiec_cache_status` | What is stored locally, what is stale, and when each quarter was last checked for amendments |

## Federal Reserve holding company filings (FR Y-9C and Y-9SP)

Version 0.3 reads the Federal Reserve's holding company financial data from the FFIEC National Information Center (NIC). Banks file Call Reports with their regulator; their holding companies file separately with the Federal Reserve:

- **FR Y-9C**: consolidated, quarterly, for larger holding companies (about $3 billion of assets and up). Full balance sheet, income statement, capital ratios and asset quality.
- **FR Y-9SP**: parent-only, semiannual (June and December), for smaller holding companies. It also reports the holding company's consolidated assets.
- **FR Y-9LP**: parent-only, for holding companies that file the Y-9C.

Tools: `fry9_status`, `fry9_find`, `fry9_summary`, `fry9_trend`, `fry9_items`, `fry9_import_file`. Income items are year to date as filed; `fry9_trend` also derives single-quarter (Y-9C) or single-half (Y-9SP) values by differencing within the year when the prior filing is loaded. For small holding companies the consolidated picture is in the subsidiary bank's Call Report, so pair `fry9_find` with `ffiec_find_bank`.

**Getting the data.** The NIC download page is behind a browser check that blocks scripts, so the quarterly files are added by hand:

1. Open https://www.ffiec.gov/npw/FinancialReport/FinancialDataDownload in a browser, pick the year, and click `BHCFYYYYMMDD.ZIP` (each is 0.6 to 1.5 MB).
2. Ask Claude to run `fry9_import_file` with the downloaded path, or drop the zip into the holding company data folder (default: `FRY9` inside the download folder).
3. Files appear once institutions file (deadline is 45 days after quarter end); `fry9_status` says when a newer quarter should exist. Download both March/September (Y-9C only) and June/December (Y-9C plus Y-9SP) files; a semiannual filer shows up only in June and December.

Capital ratios are taken as filed (BHCA7204, BHCA7205, BHCA7206, BHCAP793). Other ratios (efficiency, ROA, nonperforming loans, net charge-offs) are computed from filed items and labeled; Y-9C net interest income and provision can differ slightly from commercial data providers because of definitions (for example, provision for unfunded commitments).

## Before you install: get FFIEC credentials

Each person needs their own free FFIEC account. Do not share credentials.

1. Go to https://cdr.ffiec.gov/public/PWS/ and register for a Public Web Service (PWS) account.
2. Log in and generate a **REST API security token**. It is a long string (several hundred characters) that starts with `ey`. Copy all of it.
3. Note your PWS **username** exactly as you log in with it (it is not your email address).

Tokens expire every **90 days**; generate a new one in the same portal. FFIEC deletes accounts after a year of inactivity.

## Install

You need Claude Desktop (Windows or macOS). Nothing else: Claude Desktop installs Python and the libraries automatically on first launch.

1. Download `ffiec.mcpb` from the **Releases** page of this repository (right side of the repo home page, latest release, under Assets).
2. Install it in either of these ways:
   - Double-click `ffiec.mcpb`, or
   - In Claude Desktop, open **Settings > Extensions**, then drag the file into the window (or use the install button there).
3. When prompted, enter:
   - **CDR PWS Username**: your FFIEC PWS username
   - **CDR PWS Security Token**: the full token
   - **Download Folder**: where downloaded filings and Excel exports are saved. Leave the default to use `Documents\FFIEC`.
4. **Fully quit Claude Desktop** (Windows: right-click the Claude icon in the system tray and choose Quit; macOS: Claude > Quit) and reopen it. Closing the window is not enough.
5. Test it: ask Claude "List the latest FFIEC Call Report periods."

The first launch takes a few extra seconds while dependencies install.

## Updating

- **New token (every 90 days):** Settings > Extensions > FFIEC CDR > settings, paste the new token, then fully quit and reopen Claude Desktop.
- **New version of the extension:** download the new `ffiec.mcpb` from Releases, remove the old extension in Settings > Extensions, install the new one, re-enter your credentials, then fully quit and reopen.

## Important: income statement items are year-to-date

Call Report income statement items (Schedule RI) are cumulative from January 1: Q1 covers 3 months, Q2 covers 6, Q3 covers 9, Q4 covers the full year. Never add quarterly YTD figures together.

The Excel export handles this for you. Each bank sheet shows:

- **Balance sheet**, period-end values.
- **Income statement, year-to-date as filed** (blue, hardcoded from the filing).
- **Income statement, discrete quarter** (black, live Excel formulas): current YTD less prior-quarter YTD of the same year; Q1 equals Q1 YTD. The prior quarter needed for the first calculation is fetched automatically and shown as a grey helper column marked `*`.
- **Key ratios and regulatory capital.** Annualized ratios scale YTD income by 12 / months elapsed.

Optional extras: per-bank Detail sheets for any schedule (each row flagged YTD or period end using the filing's own XBRL), a Compare sheet when you export several banks, a long-format Data sheet for pivot tables, and a UBPR sheet.

## Units

- Call Report figures (Call report, summary, trend, Excel): **$ thousands**.
- XBRL downloads and UBPR dollar items from `ffiec_ubpr`: **whole dollars**. (The Excel UBPR sheet converts them to thousands.)
- UBPR ratios are in percent units as published (1.05 means 1.05%) and are already YTD and annualized.
- Exception: UBPR stores three capital ratios carried over from the Call Report, `UBPR7204`, `UBPR7205`, and `UBPR7206`, as decimal fractions rounded to two places (0.11 means 11%). The tools flag them; for precise figures use `UBPRD486` / `UBPRD487` or the Call Report's own capital ratios.

## Troubleshooting

| Symptom | Fix |
|---|---|
| `Error Code 5002: Security Token is invalid` | Generate a new REST API token, paste all of it, confirm the username, then fully quit and reopen Claude Desktop. |
| Still 5002 right after updating the token | Claude Desktop is still running the old server process. Quit from the system tray, not the window. |
| `HTTP 403 ... quota` | FFIEC allows about 2,500 calls per hour per account. Wait for the next hour. |
| `No content: facsimile not found` | The bank did not file for that period (new, merged, or not yet filed). |
| An export or trend returns `status: in_progress` | Normal for big requests. Claude calls the tool again automatically; downloads are cached, so each call picks up where the last stopped. If Claude stops, say "continue the export". |
| "database is locked" | The download folder is probably synced by OneDrive (or similar). Set Download Folder to a folder outside OneDrive, for example `C:\FFIEC`. |
| Extension fails to start | Check the log: Settings > Developer (or on Windows `%LOCALAPPDATA%\Claude\Logs\mcp-server-FFIEC CDR (Call Reports & UBPR).log`) and send it to the maintainer. |

Large exports download in stages of about 40 seconds so Claude Desktop does not time out; each stage resumes from the cache. A 3-bank, 4-quarter export with detail and UBPR is about 30 filings. The extension tracks FFIEC's hourly limit (about 2,500 calls) and pauses before hitting it. Downloaded filings are cached in the download folder, so repeat requests are fast and free; amended filings are detected and downloaded again (see "Local cache and amended filings").

## Local cache and amended filings

Everything the extension downloads stays on your machine: filings in the download folder, plus a small database at `<download folder>/.ffiec_cache/cache.sqlite` that records what was fetched and when. Nothing is shared or uploaded. The database also keeps the bank list and reporting periods, so restarting Claude Desktop does not fetch them again.

Banks amend Call Reports after filing. Once a day per quarter, the first time you use cached data for that quarter, the extension makes one API call asking FFIEC which banks filed or amended since your copy was downloaded. Any older copy is marked stale and downloaded again the next time you ask for it. If that check fails (network, token, hourly limit), cached data is still served and the check is retried an hour later. To change the interval, set the environment variable `FFIEC_CHECK_HOURS` (default 24).

Ask "show the FFIEC cache status" to see what is stored.

## Loading a whole quarter at once (bulk download)

For peer work across many banks, load a whole quarter from FFIEC's bulk file instead of calling the API bank by bank:

1. Go to https://cdr.ffiec.gov/public/PWS/DownloadBulkData.aspx.
2. Choose **Call Reports -- Single Period**, the quarter, and **Tab Delimited**, then download the ZIP.
3. Ask Claude: "Import the FFIEC bulk file at C:\Users\me\Downloads\FFIEC CDR Call Bulk All Schedules 06302025.zip".

A quarter (about 4,300 banks) typically loads in one tool call of 10 to 30 seconds, and adds about 35 MB to the database; on a slow machine it may take a second call. After that, Call Report tools, summaries, trends, and Excel exports read those banks locally. The amendment check above covers bulk-loaded banks too. Line numbers are blank for bulk-loaded items, and Detail sheets still download one XBRL file per bank for the YTD/period-end flags. UBPR data is not included in the bulk import; it still comes from the API.

## Data sources

- FFIEC CDR Public Web Service, SIS611 specification: https://cdr.ffiec.gov/public/Files/SIS611_-_Retrieve_Public_Data_via_Web_Service.pdf
- Item names: Federal Reserve Micro Data Reference Manual (MDRM), https://www.federalreserve.gov/apps/mdrm/ . A snapshot is bundled; ask Claude to "refresh the FFIEC item dictionary" to pull the latest.

## Credits

Created and maintained by [carrera777](https://github.com/carrera777/ffiec-mcp). If you fork or redistribute this, please keep the credit and link back. Not affiliated with the FFIEC or the Federal Reserve.

---

## For maintainers

```
manifest.json          Extension manifest (version, settings, tool list)
src/server.py          MCP server: FFIEC API client and tools
src/excel_export.py    Excel workbook builder
src/store.py           Local SQLite cache: filing catalog, amendment checks, persisted lookups, bulk import
src/fry9.py            Holding company (FR Y-9) reader for NIC BHCF files
src/mdrm_labels.json.gz  Bundled MDRM item names
src/bhc_labels.json.gz   Bundled holding company item names
tests/test_offline.py  Offline tests with a mocked FFIEC API
tests/test_fry9.py     Holding company tests
build.py               Packages dist/ffiec.mcpb
```

Run the tests (requires Python 3.10+ with `mcp<2`, `httpx`, `openpyxl`):

```bash
python tests/test_offline.py
python tests/test_fry9.py
```

Build the bundle:

```bash
python build.py
```

To release: bump `version` in `manifest.json`, run the tests, run `python build.py`, then create a GitHub release tagged `v<version>` and attach `dist/ffiec.mcpb`.

Notes:

- Dependencies are declared inline at the top of `src/server.py` (PEP 723) and installed by Claude Desktop's bundled `uv`. Keep `mcp` pinned below 2: version 2 removed `mcp.server.fastmcp`.
- The FFIEC API expects the header `Authentication: Bearer <token>` (not `Authorization`) plus `UserID`. A bad token returns HTTP 500 with error code 5002.
- The SDF format lists each item under only one schedule (total assets appears under RC-R, not RC), so the Excel statements use explicit item-code templates rather than schedule filters.
