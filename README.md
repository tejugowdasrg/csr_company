# CSR Company Scraper — Pune & Mumbai
### Live Internet Data Extractor | Technical Documentation |

---

`scrape_csr_live.py` is a Python web scraper that automatically collects **Corporate Social Responsibility (CSR)** data for companies headquartered in **Pune and Mumbai** by browsing multiple public websites.

It extracts:
- Company name, city, sector, company type
- CSR spend amount (₹ Cr)
- CSR Foundation / Page name
- CSR Contact Person & Contact Email
- Key CSR Focus Areas (Education, Health, Environment, etc.)
- Approx. Net Profit (FY 2023-24)
- Listed / Unlisted status

The script then writes everything into a formatted Excel file that **matches the `CSR_Contact_Person_Final_2__1_.xlsx` template exactly** — one row per focus area, with alternating colour rows, yellow focus area column, and all 15 columns intact.

---

### Data Sources

| # | Source | URL | Data Extracted |
|---|--------|-----|----------------|
| 1 | The CSR Journal | `thecsrjournal.in` | CSR spend, focus areas, foundation name, sector |
| 2 | CSRBOX | `csrbox.org` | Company names, project descriptions, CSR amounts |
| 3 | CSR Portal (Official) | `csr.gov.in` *(via Playwright browser)* | Official CSR spend, Schedule VII sector, company list |
| 4 | Wikipedia | `en.wikipedia.org` | Industry sector, company type, net profit |

---

## 2. Output File

The script produces a single Excel file:

```
CSR_Companies_Pune_Mumbai_Live.xlsx
```

It contains **2 sheets**:

### Sheet 1 — CSR Focus Areas – Expanded
One row per focus area per company. Contains all 15 columns:

| Column | Description |
|--------|-------------|
| Sr. | Serial number |
| Company Name | Full company name |
| CSR Spend (₹ Cr) | Amount spent on CSR in FY 2023-24 |
| City | Mumbai or Pune |
| Sector / Industry | e.g. Banking, IT / Technology, Pharmaceuticals |
| Company Type | Public Ltd / Private Ltd / PSU |
| Listed / Unlisted | Listed (BSE/NSE) or Unlisted |
| Mandatory CSR | Yes / No |
| Approx. Net Profit (₹ Cr, FY24) | Estimated net profit |
| CSR Foundation / Page | Name of CSR arm or foundation |
| CSR Contact Person | Name and designation of CSR head |
| Contact Email | Official CSR contact email |
| Key CSR Focus Area | ONE focus area per row (Yellow column) |
| Focus # | Sequence number within the company |
| Total Focus Areas | Total number of focus areas for that company |

### Sheet 2 — Focus Area Stats
Summary view — one row per company:

| Column | Description |
|--------|-------------|
| Company Name | Full company name |
| City | Mumbai or Pune |
| Sector | Industry sector |
| No. of Focus Areas | Total CSR focus areas |

---

## 3. Prerequisites

Before running the scraper, make sure you have:

- **Python 3.10 or higher** installed
- **pip** (comes bundled with Python)
- **Active internet connection** (the script fetches live data)
- **Windows, macOS, or Linux** operating system

Check your Python version:

```bash
python --version
```

You should see `Python 3.10.x` or higher.

---

## 4. Installation — Step by Step

### Step A — Create a Virtual Environment

A virtual environment keeps project dependencies isolated from your system Python.

**Windows:**
```bash
cd E:\internship\priorup
python -m venv venv
venv\Scripts\activate
```

**Mac / Linux:**
```bash
cd ~/your-project-folder
python3 -m venv venv
source venv/bin/activate
```

> After activation, your terminal prompt will show `(venv)` at the start.

---

### Step B — Install Python Dependencies

Install all required packages using the `requirements.txt` file:

```bash
pip install -r requirements.txt
```

The `requirements.txt` contains:

| Package | Version | Purpose |
|---------|---------|---------|
| `requests` | 2.33.1 | HTTP requests — fetches web pages |
| `beautifulsoup4` | 4.14.3 | HTML parsing — extracts data from web pages |
| `lxml` | 6.1.0 | Fast HTML/XML parser (used by BeautifulSoup) |
| `openpyxl` | 3.1.5 | Creates and formats the Excel output file |
| `playwright` | 1.59.0 | Headless browser — scrapes JavaScript-rendered pages (csr.gov.in) |

---

### Step C — Install Playwright Browser

Playwright needs a Chromium browser to render JavaScript-heavy pages like `csr.gov.in`. Install it once:

```bash
playwright install chromium
```

> This downloads ~300 MB of browser binaries. It only needs to be done **once**.

---

## 5. Running the Scraper

Make sure your virtual environment is active (you see `(venv)` in the terminal), then run:

```bash
python scrape_csr_live.py
```

The scraper will automatically run through these steps:

| Step | Source | What Happens |
|------|--------|--------------|
| 1 | `thecsrjournal.in` | Fetches verified CSR report pages for 35 companies |
| 2 | `csrbox.org` | Scrapes Maharashtra & Mumbai CSR project articles |
| 3 | `csr.gov.in` (Pune) | Playwright opens browser, navigates to Pune district filter |
| 4 | `csr.gov.in` (Mumbai) | Playwright opens browser, navigates to Mumbai district filter |
| 5 | Merge & Deduplicate | All records merged, duplicates removed, focus areas unioned |
| 6 | Wikipedia | Enriches sector, company type, net profit from Wikipedia infoboxes |
| 7 | Excel Writer | Writes `CSR_Companies_Pune_Mumbai_Live.xlsx` with 2 formatted sheets |

>  **Expected runtime: 10–20 minutes** depending on your internet speed and how many pages are accessible.

---

## 6. Output File Location

The Excel file is saved in the **same folder where you run the script**.

| If you run from... | Output file is saved at... |
|--------------------|---------------------------|
| `E:\internship\priorup\` | `E:\internship\priorup\CSR_Companies_Pune_Mumbai_Live.xlsx` |
| `C:\Users\YourName\Desktop\` | `C:\Users\YourName\Desktop\CSR_Companies_Pune_Mumbai_Live.xlsx` |
| `~/projects/csr/` | `~/projects/csr/CSR_Companies_Pune_Mumbai_Live.xlsx` |

To confirm where you are before running the script:

```bash
# Windows
cd

# Mac / Linux
pwd
```

To open the file after it's generated:

```bash
# Windows
start CSR_Companies_Pune_Mumbai_Live.xlsx

# Mac
open CSR_Companies_Pune_Mumbai_Live.xlsx

# Linux
xdg-open CSR_Companies_Pune_Mumbai_Live.xlsx
```

---

## 7. Project Files Summary

| File | Type | Description |
|------|------|-------------|
| `scrape_csr_live.py` | Python script | Main scraper — run this to collect CSR data |
| `requirements.txt` | Text file | Lists all Python packages needed |
| `CSR_Companies_Pune_Mumbai_Live.xlsx` | Excel output | Generated after running the script |
| `CSR_Contact_Person_Final_2__1_.xlsx` | Reference template | Original template — output matches this format exactly |
| `CSR_Scraper_Documentation.md` | Documentation | This file |

---

## 8. Troubleshooting

| Error / Issue | Solution |
|---------------|----------|
| `[403]` or `[404]` warnings in terminal | **Normal** — some pages block automated access. The script skips them and continues. Output is still generated. |
| `ModuleNotFoundError: No module named 'requests'` | Run: `pip install -r requirements.txt` — make sure `(venv)` is active first |
| `playwright: command not found` | Run: `pip install playwright` then `playwright install chromium` |
| SSL Certificate errors | The script automatically disables SSL verification for Indian government sites. No action needed. |
| Output file has 0 companies | Check your internet connection. All source sites may be temporarily blocking the scraper. |
| `PermissionError` on Excel file | Close `CSR_Companies_Pune_Mumbai_Live.xlsx` in Excel before running the script again. |
| `python` not recognized | Use `python3` instead of `python` on Mac/Linux |

---

## 9. Quick Reference — Full Setup Commands

Copy and paste these commands in order to set up and run from scratch:

**Windows:**
```bash
cd E:\internship\priorup
python -m venv venv
venv\Scripts\activate
pip install -r requirements.txt
playwright install chromium
python scrape_csr_live.py
```

**Mac / Linux:**
```bash
cd ~/your-project-folder
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt
playwright install chromium
python scrape_csr_live.py
```

---

## 10. Notes

- **403 / 404 errors are expected** — Indian government and CSR portal sites often block automated requests. The scraper handles this gracefully and still produces output.
- **Contact person emails** are sourced from official company annual reports (FY 2023-24) and are stored in the script's `KNOWN_CONTACTS` dictionary, since no public website lists them in a structured way.
- **Re-running the script** overwrites the previous output file. Rename or move the old file if you want to keep it.
- **csr.gov.in scraping** requires Playwright (the real Chromium browser) because the site uses JavaScript to render data — a standard `requests` call returns a 403 error.

---

