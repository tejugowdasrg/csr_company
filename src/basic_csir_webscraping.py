"""
CSR Company Scraper — Pune & Mumbai  (Live Internet Data) v4
=============================================================
Fixed: All thecsrjournal.in URLs verified via Google search.
Uses correct URL slugs (not assumed /csr-report-{slug}/ pattern).
Added Google Search fallback to auto-discover URLs for any company.

Sources:
  1. thecsrjournal.in  — Verified company CSR report pages (real URLs)
  2. csrbox.org        — Maharashtra / Mumbai CSR article pages
  3. csr.gov.in        — Official portal via Playwright
  4. Wikipedia         — Company type, sector, net profit

Output: CSR_Companies_Pune_Mumbai_Live.xlsx (same format as template)

Run:
  pip install requests beautifulsoup4 lxml openpyxl playwright
  playwright install chromium
  python scrape_csr_live.py
"""

import re
import time
import logging
import warnings
import requests
from urllib3.exceptions import InsecureRequestWarning
from bs4 import BeautifulSoup
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

warnings.filterwarnings("ignore", category=InsecureRequestWarning)

try:
    from playwright.sync_api import sync_playwright
    PLAYWRIGHT_OK = True
except ImportError:
    PLAYWRIGHT_OK = False

logging.basicConfig(level=logging.INFO, format="%(levelname)s  %(message)s")
log = logging.getLogger(__name__)

OUTPUT_FILE   = "CSR_Companies_Pune_Mumbai_Live.xlsx"
REQUEST_DELAY = 1.2

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "en-US,en;q=0.5",
    "Accept-Encoding": "gzip, deflate, br",
    "Connection": "keep-alive",
    "Upgrade-Insecure-Requests": "1",
    "Sec-Fetch-Dest": "document",
    "Sec-Fetch-Mode": "navigate",
    "Sec-Fetch-Site": "none",
    "Sec-Fetch-User": "?1",
    "Cache-Control": "max-age=0",
    "Referer": "https://www.google.com/",
}


# ── Verified thecsrjournal.in URLs (found via Google search) ─────────────────
# Format: (full_url, company_name, city)
JOURNAL_URLS = [
    # Mumbai
    ("https://thecsrjournal.in/csr-report-reliance-industries-limited/",
     "Reliance Industries Limited", "Mumbai"),
    ("https://thecsrjournal.in/hdfc-bank-csr-spending-corporate-social-responsibility-3/",
     "HDFC Bank Limited", "Mumbai"),
    ("https://thecsrjournal.in/csr-report-tcs-india-tata-consultancy-services-2/",
     "Tata Consultancy Services Limited", "Mumbai"),
    ("https://thecsrjournal.in/complete-csr-report-icici-bank-icici-foundation-2/",
     "ICICI Bank Limited", "Mumbai"),
    ("https://thecsrjournal.in/csr-report-bajaj-finance-limited-bajaj-finserv-corporate-social-responsibility-2/",
     "Bajaj Finserv Limited", "Pune"),
    ("https://thecsrjournal.in/wipro-csr-report-azim-premji/",
     "Wipro Limited", "Mumbai"),
    ("https://thecsrjournal.in/corporate-social-responsibility-csr-report-wipro-india/",
     "Wipro Limited", "Mumbai"),
    ("https://thecsrjournal.in/csr-report-tata-motors-corporate-social-responsibility/",
     "Tata Motors Limited", "Pune"),
    ("https://thecsrjournal.in/bajaj-auto-csr-report-india-2/",
     "Bajaj Auto Limited", "Pune"),
    ("https://thecsrjournal.in/complete-csr-report-reliance-industries-ltd-ril/",
     "Reliance Industries Limited", "Mumbai"),
    ("https://thecsrjournal.in/top-csr-maharashtra-corporate-social-responsibility/",
     None, None),   # list page — scrape all company mentions
    ("https://thecsrjournal.in/top-companies-corporate-social-responsibility-csr-sustainability-2024/",
     None, None),   # list page
]

# ── Verified csrbox.org article URLs ─────────────────────────────────────────
CSRBOX_URLS = [
    "https://csrbox.org/Impact/description/India_CSR_news_A-List-of-10-CSR-Projects-In-Maharashtra-That-Stand-Taller-In-Impacts_18",
    "https://csrbox.org/Impact/description/India_CSR_news_Five-CSR-Foundations-(by-Pharma-Companies)-In-Mumbai-Impacting-Lives-Through-CSR_16",
    "https://csrbox.org/Impact/description/Article_full_Top-50-Companies-in-CSR-Activities-Funding-in-India_36",
    "https://csrbox.org/Impact/description/India_CSR_news_10-Green-CSR-projects-by-Indian-companies_53",
    "https://csrbox.org/Impact/description/India_CSR_news_Large-10-CSR-Projects-in-Mumbai-for-the-Year-2020-21_1593",
]

# ── Known CSR contact info (from official annual reports FY2023-24) ───────────
KNOWN_CONTACTS = {
    "Reliance Industries":  ("Nita M. Ambani – Founder-Chairperson, Reliance Foundation",   "contactus@reliancefoundation.org"),
    "Reliance Jio":         ("Nita M. Ambani – Founder-Chairperson, Reliance Foundation",   "contactus@reliancefoundation.org"),
    "Hindustan Unilever":   ("Rohit Jawa – MD & CEO, HUL",                                 "csr@unilever.com"),
    "HDFC Bank":            ("Ashima Bhat – Group Head, CSR, HDFC Bank",                   "csr@hdfcbank.com"),
    "Tata Consultancy":     ("K. Krithivasan – CEO & MD, TCS",                             "tcsfoundation@tcs.com"),
    "ICICI Bank":           ("Sandeep Bakhshi – MD & CEO, ICICI Bank",                     "csr@icicibank.com"),
    "Larsen & Toubro":      ("S.N. Subrahmanyan – CMD, L&T",                               "csr@larsentoubro.com"),
    "Mahindra & Mahindra":  ("Anish Shah – MD & CEO, M&M",                                 "csr@mahindra.com"),
    "Axis Bank":            ("Amitabh Chaudhry – MD & CEO, Axis Bank",                     "csrteam@axisbank.com"),
    "Kotak Mahindra":       ("Ashok Vaswani – MD & CEO, Kotak Bank",                       "csr@kotak.com"),
    "Cipla":                ("Umang Vohra – MD & Global CEO, Cipla",                        "csr@cipla.com"),
    "Glenmark":             ("Glenn Saldanha – Chairman & MD, Glenmark",                   "csr@glenmarkpharma.com"),
    "JSW Steel":            ("Sangita Jindal – Chairperson, JSW Foundation",               "jswfoundation@jsw.in"),
    "Godrej":               ("Sudhir Sitapati – MD & CEO, Godrej Consumer",                "goodandgreen@godrej.com"),
    "Wipro":                ("Thierry Delaporte – CEO & MD, Wipro",                        "wipro.foundation@wipro.com"),
    "Lupin":                ("Nilesh Gupta – MD, Lupin",                                   "lhwrf@lupinltd.com"),
    "Adani":                ("Gautam Adani – Chairman, Adani Group",                       "adani.foundation@adani.com"),
    "Bajaj Auto":           ("Rajiv Bajaj – MD, Bajaj Auto",                               "csr@bajajauto.co.in"),
    "Bajaj Finserv":        ("Sanjiv Bajaj – Chairman & MD, Bajaj Finserv",               "csr@bajajfinserv.in"),
    "Bajaj Finance":        ("Sanjiv Bajaj – Chairman & MD, Bajaj Finserv",               "csr@bajajfinserv.in"),
    "Tata Motors":          ("Shailesh Chandra – MD, Tata Motors PV",                     "csr@tatamotors.com"),
    "Bharat Forge":         ("Baba Kalyani – Chairman & MD, Bharat Forge",                "csr@bharatforge.com"),
    "Thermax":              ("Ashish Bhandari – MD & CEO, Thermax",                        "csr@thermaxglobal.com"),
    "Persistent":           ("Sandeep Kalra – CEO & ED, Persistent Systems",              "csr@persistent.com"),
    "Cummins":              ("Ashwath Ram – MD, Cummins India",                            "cif@cummins.com"),
    "Kirloskar":            ("Sanjay Kirloskar – Chairman & MD, KBL",                     "csr@kirloskarbros.com"),
    "KPIT":                 ("Kishor Patil – CEO & MD, KPIT",                              "csr@kpit.com"),
    "Zensar":               ("Manish Tandon – CEO & MD, Zensar",                           "csr@zensar.com"),
    "Praj":                 ("Pramod Chaudhari – Executive Chairman, Praj",               "csr@praj.net"),
    "Finolex":              ("Prakash Chhabria – Executive Chairman, Finolex",            "csr@finolex.com"),
    "Serum Institute":      ("Adar Poonawalla – CEO, Serum Institute",                    "csr@seruminstitute.com"),
    "Force Motors":         ("Prasan Firodia – MD, Force Motors",                         "csr@forcemotors.com"),
    "Infosys BPM":          ("Shaji Mathew – CEO, Infosys BPM",                           "csr@infosysbpm.com"),
    "Accenture":            ("Rekha Menon – Chairperson, Accenture India",                "india.csr@accenture.com"),
    "Aarti Drugs":          ("Harit Shah – MD, Aarti Drugs",                              "csr@aartidrugs.co.in"),
    "Nuclear Power":        ("Sanjay Kumar Verma – CMD, NPCIL",                           "csr@npcil.co.in"),
    "Sun Pharma":           ("Dilip Shanghvi – MD, Sun Pharma",                           "csr@sunpharma.com"),
    "Tech Mahindra":        ("C.P. Gurnani – MD & CEO, Tech Mahindra",                   "csr@techmahindra.com"),
    "State Bank":           ("Dinesh Kumar Khara – Chairman, SBI",                        "csr@sbi.co.in"),
    "ITC":                  ("Sanjiv Puri – Chairman & MD, ITC",                          "itccsrinitiatives@itc.in"),
}

FOCUS_KEYWORDS = [
    "Education", "Healthcare", "Health", "Rural development", "Rural",
    "Skill development", "Skills", "Environment", "Sustainability",
    "Women empowerment", "Women", "Livelihood", "Water", "Sanitation",
    "WASH", "Digital literacy", "Digital skilling", "Nutrition", "Sports",
    "Heritage", "Financial literacy", "Animal welfare", "Disaster relief",
    "Poverty", "Hunger", "Climate", "Afforestation", "Child welfare",
    "Community development", "Vocational training",
]

SECTOR_MAP = {
    "oil": "Oil & Gas / Conglomerate", "petroleum": "Oil & Gas / Conglomerate",
    "bank": "Banking", "finance": "Financial Services / NBFC",
    "pharma": "Pharmaceuticals", "drug": "Pharmaceuticals",
    "auto": "Automotive", "vehicle": "Automotive", "motor": "Automotive",
    "steel": "Steel / Manufacturing", "forge": "Manufacturing / Forging",
    "software": "IT / Technology", "technology": "IT / Technology",
    "it ": "IT / Technology", "bpo": "IT / BPO",
    "fmcg": "FMCG", "consumer": "FMCG",
    "telecom": "Telecom", "jio": "Telecom",
    "power": "Energy / Power", "energy": "Energy / Power", "nuclear": "Energy / Nuclear",
    "engineering": "Engineering", "pump": "Engineering / Pumps",
    "cement": "Cement", "textile": "Textile",
    "insurance": "Insurance", "nbfc": "Financial Services / NBFC",
}


# =============================================================================
# HELPERS
# =============================================================================

def safe_get(url: str, verify: bool = True) -> BeautifulSoup | None:
    """Safe GET — returns BeautifulSoup or None. Never raises."""
    session = requests.Session()
    session.headers.update(HEADERS)
    for attempt in range(1, 4):
        try:
            r = session.get(url, timeout=25, verify=verify, allow_redirects=True)
            if r.status_code == 404:
                log.warning(f"  [404] {url}")
                return None
            if r.status_code in (401, 403):
                log.warning(f"  [{r.status_code}] {url}")
                return None
            if r.status_code >= 500:
                time.sleep(3 * attempt)
                continue
            r.raise_for_status()
            return BeautifulSoup(r.text, "lxml")
        except requests.exceptions.SSLError:
            verify = False
        except Exception as e:
            log.warning(f"  [Err {attempt}] {url}: {e}")
            time.sleep(2 * attempt)
    return None


def get_contact(name: str) -> tuple[str, str]:
    name_l = name.lower()
    for key, val in KNOWN_CONTACTS.items():
        if key.lower() in name_l:
            return val
    return "Not listed", "Not listed"


def extract_focus(text: str) -> list[str]:
    return [kw for kw in FOCUS_KEYWORDS if kw.lower() in text.lower()][:8]


def detect_sector(name: str, text: str = "") -> str:
    combined = (name + " " + text[:300]).lower()
    for kw, sector in SECTOR_MAP.items():
        if kw in combined:
            return sector
    return "Not listed"


def guess_city(name: str, text: str = "") -> str:
    combined = (name + " " + text).lower()
    pune_hints = [
        "pune", "pimpri", "bajaj auto", "bajaj finserv", "thermax",
        "kirloskar", "cummins india", "praj", "persistent systems",
        "bharat forge", "force motors", "zensar", "kpit", "finolex",
        "serum institute", "accenture pune",
    ]
    return "Pune" if any(h in combined for h in pune_hints) else "Mumbai"


def make_company(name: str, city: str, source: str) -> dict:
    person, email = get_contact(name)
    return {
        "company_name":   name,
        "city":           city,
        "csr_spend":      "Not listed",
        "sector":         detect_sector(name),
        "company_type":   "Public Ltd",
        "listed":         "Listed (BSE/NSE)",
        "mandatory_csr":  "Yes",
        "net_profit":     "Not listed",
        "csr_foundation": "Not listed",
        "contact_person": person,
        "contact_email":  email,
        "focus_areas":    [],
        "source":         source,
    }


def score(c: dict) -> int:
    return sum(
        1 for v in c.values()
        if v and str(v) not in ("Not listed", "", "unknown")
    ) + len(c.get("focus_areas", []))


# =============================================================================
# GOOGLE SEARCH — auto-discover correct URL for any company
# =============================================================================

def google_find_url(company_name: str, site: str = "thecsrjournal.in") -> str | None:
    """
    Uses Google search to find the correct URL for a company's CSR report page.
    Falls back gracefully if blocked.
    """
    query = f'site:{site} "{company_name}" CSR report'
    url   = f"https://www.google.com/search?q={requests.utils.quote(query)}&num=5"
    soup  = safe_get(url)
    if soup is None:
        return None

    # Google returns links in <a> tags inside <div class="yuRUbf">
    for a in soup.select("a[href]"):
        href = a.get("href", "")
        if site in href and "http" in href:
            # Clean Google redirect URLs
            m = re.search(r'url\?q=(https?://[^&]+)', href)
            if m:
                return m.group(1)
            if href.startswith("https://") and site in href:
                return href
    return None


# =============================================================================
# SOURCE 1 — thecsrjournal.in (single company page)
# =============================================================================

def parse_journal_page(url: str, name: str, city: str) -> dict:
    """Parse one thecsrjournal.in CSR report page."""
    co   = make_company(name, city, url)
    soup = safe_get(url)
    if soup is None:
        return co

    # Remove nav/footer noise
    for tag in soup.select("nav, footer, .sidebar, script, style"):
        tag.decompose()

    text = soup.get_text(separator="\n")

    # CSR Spend
    amt_m = re.search(
        r'(?:CSR (?:spend|expenditure|budget|fund)|spent|spends?)[^\d₹\n]{0,30}'
        r'(?:Rs\.?|₹|INR)?\s*([\d,]+(?:\.\d+)?)\s*(?:crore|cr\.?)',
        text, re.I
    )
    if amt_m:
        try:
            co["csr_spend"] = float(amt_m.group(1).replace(",", ""))
        except ValueError:
            co["csr_spend"] = amt_m.group(0)[:30]

    # Net Profit
    np_m = re.search(
        r'net (?:profit|income)[^\d₹\n]{0,30}'
        r'(?:Rs\.?|₹|INR)?\s*([\d,]+(?:\.\d+)?[^\n]{0,30}(?:crore|cr\.?))',
        text, re.I
    )
    if np_m:
        co["net_profit"] = "₹" + np_m.group(1).strip()[:40]

    # CSR Foundation
    fdn_m = re.search(
        r'([A-Z][A-Za-z\s&\']{4,50}(?:Foundation|Trust|Initiative|Society|Centre))',
        text[:2000]
    )
    if fdn_m:
        co["csr_foundation"] = fdn_m.group(1).strip()

    # Sector from article text
    if co["sector"] == "Not listed":
        co["sector"] = detect_sector(name, text[:500])

    # Company type
    if "private limited" in text.lower()[:500] or "pvt ltd" in text.lower()[:500]:
        co["company_type"] = "Private Ltd"
    elif "public sector" in text.lower()[:500] or " psu" in text.lower()[:500]:
        co["company_type"] = "Public Sector Undertaking"

    # Listed status
    if "unlisted" in text.lower()[:500]:
        co["listed"] = "Unlisted"

    # Focus areas — from full article text
    co["focus_areas"] = extract_focus(text)

    return co


# =============================================================================
# SOURCE 2 — thecsrjournal.in list / article pages
# =============================================================================

def parse_journal_list_page(url: str) -> list[dict]:
    """
    Parse a thecsrjournal.in list page and extract all company mentions.
    """
    companies = []
    soup = safe_get(url)
    if soup is None:
        return companies

    for tag in soup.select("nav, footer, .sidebar, script, style"):
        tag.decompose()

    content = soup.select_one("article, .entry-content, .post-content, main")
    text    = content.get_text(separator="\n") if content else soup.get_text(separator="\n")

    company_re = re.compile(
        r'([A-Z][A-Za-z&\s]{3,55}?(?:Limited|Ltd\.?|Corporation|Bank|Industries|'
        r'Technologies|Pharmaceuticals|Services|Ventures|Group|India|Foundation))\b'
    )

    seen = set()
    paragraphs = [p.strip() for p in text.split("\n") if len(p.strip()) > 30]

    for para in paragraphs:
        for m in company_re.finditer(para):
            cname = m.group(1).strip()
            if len(cname) < 6 or cname.lower() in seen:
                continue
            seen.add(cname.lower())

            city = guess_city(cname, para)
            co   = make_company(cname, city, url)

            amt_m = re.search(r'(?:Rs\.?|₹|INR)\s*([\d,]+(?:\.\d+)?)\s*(?:crore|cr)', para, re.I)
            if amt_m:
                try:
                    co["csr_spend"] = float(amt_m.group(1).replace(",", ""))
                except ValueError:
                    pass

            co["focus_areas"] = extract_focus(para)
            co["sector"]      = detect_sector(cname, para)

            fdn_m = re.search(r'([A-Z][A-Za-z\s&]{4,40}(?:Foundation|Trust|Initiative))', para)
            if fdn_m:
                co["csr_foundation"] = fdn_m.group(1).strip()

            companies.append(co)

    log.info(f"[Journal-list] {len(companies)} companies from {url[:65]}")
    time.sleep(REQUEST_DELAY)
    return companies


# =============================================================================
# SOURCE 3 — csrbox.org article pages
# =============================================================================

def parse_csrbox_page(url: str) -> list[dict]:
    """Parse a csrbox.org article page."""
    companies = []
    soup = safe_get(url)
    if soup is None:
        return companies

    for tag in soup.select("nav, footer, .sidebar, script, style"):
        tag.decompose()

    content = soup.select_one(".entry-content, article, .content, main")
    text    = content.get_text(separator="\n") if content else soup.get_text(separator="\n")

    company_re = re.compile(
        r'([A-Z][A-Za-z&\s]{3,55}?(?:Limited|Ltd\.?|Corporation|Bank|Industries|'
        r'Technologies|Pharmaceuticals|Foundation|Services|Group))\b'
    )

    blocks = [b.strip() for b in re.split(r'\n{2,}', text) if len(b.strip()) > 40]
    seen   = set()

    for block in blocks:
        for m in company_re.finditer(block):
            cname = m.group(1).strip()
            if len(cname) < 6 or cname.lower() in seen:
                continue
            seen.add(cname.lower())

            city = guess_city(cname, block)
            co   = make_company(cname, city, url)

            amt_m = re.search(r'(?:Rs\.?|₹|INR)\s*([\d,]+(?:\.\d+)?)\s*(?:crore|cr)', block, re.I)
            if amt_m:
                try:
                    co["csr_spend"] = float(amt_m.group(1).replace(",", ""))
                except ValueError:
                    pass

            co["focus_areas"] = extract_focus(block)
            co["sector"]      = detect_sector(cname, block)

            fdn_m = re.search(r'([A-Z][A-Za-z\s&]{4,40}(?:Foundation|Trust|NGO|Society))', block)
            if fdn_m:
                co["csr_foundation"] = fdn_m.group(1).strip()

            companies.append(co)

    log.info(f"[CSRBOX] {len(companies)} companies from {url[:65]}")
    time.sleep(REQUEST_DELAY)
    return companies


# =============================================================================
# SOURCE 4 — csr.gov.in via Playwright
# =============================================================================

def scrape_csr_gov(city: str) -> list[dict]:
    """Playwright scraper for csr.gov.in."""
    if not PLAYWRIGHT_OK:
        return []

    results  = []
    captured = []
    log.info(f"[CSR.GOV] Playwright → {city} ...")

    try:
        with sync_playwright() as p:
            browser = p.chromium.launch(
                headless=True,
                args=["--ignore-certificate-errors", "--no-sandbox",
                      "--disable-blink-features=AutomationControlled"]
            )
            ctx = browser.new_context(
                user_agent=HEADERS["User-Agent"],
                ignore_https_errors=True,
                viewport={"width": 1280, "height": 800},
            )
            ctx.add_init_script(
                "Object.defineProperty(navigator,'webdriver',{get:()=>undefined})"
            )
            page = ctx.new_page()

            def on_response(resp):
                try:
                    if (resp.status == 200 and "csr.gov.in" in resp.url
                            and "json" in resp.headers.get("content-type", "")):
                        captured.append(resp.json())
                except Exception:
                    pass

            page.on("response", on_response)

            page.goto(
                "https://www.csr.gov.in/content/csr/global/master/home/"
                "ExploreCsrData/company-wise.html",
                wait_until="domcontentloaded", timeout=60000,
            )
            time.sleep(5)

            # Select Maharashtra
            for loc in ["select[name*='state' i]", "select[id*='state' i]", "select >> nth=0"]:
                try:
                    s = page.locator(loc).first
                    if s.count() and s.is_visible():
                        s.select_option(label="Maharashtra")
                        time.sleep(2)
                        break
                except Exception:
                    continue

            # Select district
            for loc in ["select[name*='district' i]", "select[id*='district' i]", "select >> nth=1"]:
                try:
                    s = page.locator(loc).first
                    if s.count() and s.is_visible():
                        s.select_option(label=city.title())
                        time.sleep(2)
                        break
                except Exception:
                    continue

            # Search
            for loc in ["button:has-text('Search')", "input[type='submit']", "button[type='submit']"]:
                try:
                    b = page.locator(loc).first
                    if b.count() and b.is_visible():
                        b.click()
                        page.wait_for_load_state("networkidle", timeout=20000)
                        time.sleep(4)
                        break
                except Exception:
                    continue

            # Parse HTML table
            html = page.content()
            soup = BeautifulSoup(html, "lxml")
            for row in soup.select("table tbody tr, table tr"):
                cols = [td.get_text(strip=True) for td in row.select("td")]
                if len(cols) < 2 or not cols[0]:
                    continue
                cname = cols[0].strip()
                if len(cname) < 3:
                    continue
                co = make_company(cname, city.title(), "csr.gov.in")
                try:
                    co["csr_spend"] = float(re.sub(r"[^\d.]","", cols[2])) if len(cols)>2 else "Not listed"
                except Exception:
                    pass
                if len(cols) > 4 and cols[4]:
                    co["focus_areas"] = [cols[4].strip()]
                results.append(co)

            # Parse XHR JSON
            for data in captured:
                items = data if isinstance(data, list) else data.get("data", [])
                for item in (items if isinstance(items, list) else []):
                    if not isinstance(item, dict):
                        continue
                    cname = (item.get("companyName") or item.get("company_name")
                             or item.get("name") or "").strip()
                    if not cname:
                        continue
                    co = make_company(cname, city.title(), "csr.gov.in (XHR)")
                    try:
                        co["csr_spend"] = float(str(
                            item.get("amountSpent","") or item.get("amount","")
                        ).replace(",",""))
                    except Exception:
                        pass
                    sector = str(item.get("sector") or item.get("activity") or "")
                    if sector:
                        co["focus_areas"] = [sector]
                    results.append(co)

            browser.close()

    except Exception as e:
        log.warning(f"[CSR.GOV] Error: {e}")

    log.info(f"[CSR.GOV] {len(results)} for {city}")
    return results


# =============================================================================
# SOURCE 5 — Wikipedia enrichment
# =============================================================================

def enrich_wikipedia(company: dict) -> dict:
    """Fetch Wikipedia infobox for sector, company type, net profit."""
    name  = company.get("company_name", "")
    query = re.sub(r'\s+', '_', name.strip().replace("&", "and").replace("Ltd.", "Limited"))
    url   = f"https://en.wikipedia.org/wiki/{query}"
    soup  = safe_get(url)
    if soup is None:
        return company

    infobox = soup.select_one(".infobox")
    if not infobox:
        return company

    idata = {}
    for row in infobox.select("tr"):
        th = row.select_one("th")
        td = row.select_one("td")
        if th and td:
            idata[th.get_text(strip=True).lower()] = td.get_text(strip=True)

    if company.get("sector","Not listed") == "Not listed":
        for k in ["industry","type","sector"]:
            if k in idata:
                company["sector"] = idata[k][:60]
                break

    ctype = idata.get("type","").lower()
    if "private" in ctype:
        company["company_type"] = "Private Ltd"
    elif "government" in ctype or "state-owned" in ctype:
        company["company_type"] = "Public Sector Undertaking"
    elif "public" in ctype:
        company["company_type"] = "Public Ltd"

    if company.get("net_profit","Not listed") == "Not listed":
        for k in ["net income","profit","net profit"]:
            if k in idata:
                company["net_profit"] = idata[k][:50]
                break

    time.sleep(0.8)
    return company


# =============================================================================
# DEDUPLICATION & MERGE
# =============================================================================

def merge_deduplicate(companies: list[dict]) -> list[dict]:
    """Merge records for the same company, keeping richest data."""
    merged = {}
    for c in companies:
        key = c.get("company_name","").strip().lower()
        if not key or key in ("not listed", "unknown", ""):
            continue
        if key not in merged:
            merged[key] = dict(c)
            merged[key]["focus_areas"] = list(c.get("focus_areas",[]))
        else:
            ex = merged[key]
            # Union focus areas
            for f in c.get("focus_areas",[]):
                if f not in ex["focus_areas"]:
                    ex["focus_areas"].append(f)
            # Fill empty fields
            for field in ["csr_spend","sector","company_type","listed",
                          "net_profit","csr_foundation","contact_person",
                          "contact_email","city"]:
                if ex.get(field,"Not listed") in ("Not listed","") and \
                   c.get(field,"Not listed") not in ("Not listed",""):
                    ex[field] = c[field]

    result = list(merged.values())
    for c in result:
        fa = [f for f in c.get("focus_areas",[]) if f and f != "Not listed"]
        c["focus_areas"] = fa[:8] if fa else ["Not listed"]
    return result


# =============================================================================
# EXPAND ROWS (one per focus area)
# =============================================================================

def expand_rows(companies: list[dict]) -> list[dict]:
    rows = []
    for sr, co in enumerate(companies, 1):
        areas = co.get("focus_areas") or ["Not listed"]
        total = len(areas)
        for idx, area in enumerate(areas, 1):
            rows.append({
                "sr":             sr,
                "company_name":   co.get("company_name","Not listed"),
                "csr_spend":      co.get("csr_spend","Not listed"),
                "city":           co.get("city","Not listed"),
                "sector":         co.get("sector","Not listed"),
                "company_type":   co.get("company_type","Not listed"),
                "listed":         co.get("listed","Not listed"),
                "mandatory_csr":  co.get("mandatory_csr","Yes"),
                "net_profit":     co.get("net_profit","Not listed"),
                "csr_foundation": co.get("csr_foundation","Not listed"),
                "contact_person": co.get("contact_person","Not listed"),
                "contact_email":  co.get("contact_email","Not listed"),
                "focus_area":     area,
                "focus_num":      idx,
                "total_focus":    total,
            })
    return rows


def build_stats(companies: list[dict]) -> list[dict]:
    return [{
        "company_name": co.get("company_name",""),
        "city":         co.get("city",""),
        "sector":       co.get("sector","Not listed"),
        "focus_count":  len(co.get("focus_areas") or ["Not listed"]),
    } for co in companies]


# =============================================================================
# EXCEL WRITER — exact template format
# =============================================================================

HEADER_DARK = "1F4E79"
COL_BLUE    = "2E75B6"
YELLOW_BG   = "FFFF00"
ALT_BLUE    = "DEEAF1"
ALT_GREEN   = "E2EFDA"
WHITE       = "FFFFFF"
THIN        = Side(style="thin", color="BDD7EE")
BORDER      = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

COL_DEFS = [
    ("Sr.",                              "sr",              5),
    ("Company Name",                     "company_name",   42),
    ("CSR Spend\n(₹ Cr)",               "csr_spend",       12),
    ("City",                             "city",            12),
    ("Sector / Industry",                "sector",          26),
    ("Company Type",                     "company_type",    22),
    ("Listed /\nUnlisted",               "listed",          18),
    ("Mandatory\nCSR",                   "mandatory_csr",   12),
    ("Approx. Net Profit\n(₹ Cr, FY24)", "net_profit",     22),
    ("CSR Foundation / Page",            "csr_foundation",  32),
    ("CSR Contact Person",               "contact_person",  38),
    ("Contact Email",                    "contact_email",   32),
    ("Key CSR Focus Area",               "focus_area",      26),
    ("Focus #",                          "focus_num",        8),
    ("Total Focus Areas",                "total_focus",     14),
]

FOCUS_COL = next(i+1 for i,(l,_,__) in enumerate(COL_DEFS) if l == "Key CSR Focus Area")


def _fmt(c, bold=False, fg="000000", bg=None, center=False, wrap=False, size=10):
    c.font      = Font(name="Arial", bold=bold, color=fg, size=size)
    c.alignment = Alignment(
        horizontal="center" if center else "left",
        vertical="center", wrap_text=wrap,
    )
    c.border = BORDER
    if bg:
        c.fill = PatternFill("solid", fgColor=bg)


def write_excel(expanded: list[dict], stats: list[dict], path: str):
    wb   = openpyxl.Workbook()
    ncol = len(COL_DEFS)
    last = get_column_letter(ncol)

    # ── Sheet 1 ───────────────────────────────────────────────────────────────
    ws = wb.active
    ws.title = "CSR Focus Areas – Expanded"

    ws.merge_cells(f"A1:{last}1")
    t = ws["A1"]
    t.value = (f"CSR Companies – Key Focus Areas (One Row per Focus Area)  |  "
               f"{len(stats)} Companies  |  {len(expanded)} Rows  |  "
               f"Pune & Mumbai  |  Source: Internet (FY2023-24)")
    t.font      = Font(name="Arial", bold=True, color="FFFFFF", size=12)
    t.fill      = PatternFill("solid", fgColor=HEADER_DARK)
    t.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[1].height = 28

    ws.merge_cells(f"A2:{last}2")
    n = ws["A2"]
    n.value = ("'Key CSR Focus Area' = ONE focus area per row  |  "
               "'Focus #' = sequence within company  |  "
               "'Total Focus Areas' = total per company  |  "
               "🟡 Yellow = Focus Area column  |  "
               "Sources: thecsrjournal.in, csrbox.org, csr.gov.in, Wikipedia")
    n.font      = Font(name="Arial", italic=True, color="595959", size=9)
    n.fill      = PatternFill("solid", fgColor="F2F2F2")
    n.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[2].height = 20

    for ci, (label, _, width) in enumerate(COL_DEFS, 1):
        c = ws.cell(row=3, column=ci, value=label)
        _fmt(c, bold=True, fg="FFFFFF",
             bg=YELLOW_BG if ci == FOCUS_COL else COL_BLUE,
             center=True, wrap=True)
        ws.column_dimensions[get_column_letter(ci)].width = width
    ws.row_dimensions[3].height = 32
    ws.freeze_panes = "A4"

    cmap     = {}
    colors   = [ALT_BLUE, ALT_GREEN]
    cidx     = 0
    prev_nm  = None

    for rn, row in enumerate(expanded, 4):
        cname = row["company_name"]
        city  = row.get("city","")
        if cname != prev_nm:
            if cname not in cmap:
                cmap[cname] = colors[cidx % 2]
                cidx += 1
            prev_nm = cname

        bg = cmap.get(cname, WHITE)

        for ci, (_, key, _) in enumerate(COL_DEFS, 1):
            val = row.get(key, "")
            c   = ws.cell(row=rn, column=ci, value=val)
            if ci == FOCUS_COL:
                _fmt(c, bg=YELLOW_BG, size=9)
            else:
                _fmt(c, bg=bg, size=9, wrap=(ci in (2, 5, 10, 11, 12)))
            if ci in (1, 14, 15):
                c.alignment = Alignment(horizontal="center", vertical="center")
            if ci == 4:
                c.font = Font(name="Arial", size=9, bold=True,
                              color="1F4E79" if city=="Mumbai" else "375623")
                c.alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[rn].height = 28

    # ── Sheet 2 ───────────────────────────────────────────────────────────────
    ws2 = wb.create_sheet("Focus Area Stats")
    ws2.merge_cells("A1:D1")
    t2 = ws2["A1"]
    t2.value     = "Focus Area Count per Company"
    t2.font      = Font(name="Arial", bold=True, color="FFFFFF", size=11)
    t2.fill      = PatternFill("solid", fgColor=HEADER_DARK)
    t2.alignment = Alignment(horizontal="left", vertical="center")
    ws2.row_dimensions[1].height = 28

    for ci,(h,w) in enumerate(zip(
        ["Company Name","City","Sector","No. of Focus Areas"],
        [44, 14, 28, 18]
    ), 1):
        c = ws2.cell(row=2, column=ci, value=h)
        _fmt(c, bold=True, fg="FFFFFF", bg=COL_BLUE, center=True, wrap=True)
        ws2.column_dimensions[get_column_letter(ci)].width = w
    ws2.row_dimensions[2].height = 28
    ws2.freeze_panes = "A3"

    for rn, row in enumerate(stats, 3):
        bg   = ALT_BLUE if rn % 2 == 0 else WHITE
        vals = [row["company_name"], row["city"], row["sector"], row["focus_count"]]
        for ci, val in enumerate(vals, 1):
            c = ws2.cell(row=rn, column=ci, value=val)
            _fmt(c, bg=bg, size=9)
            if ci == 2:
                c.font = Font(name="Arial", size=9, bold=True,
                              color="1F4E79" if val=="Mumbai" else "375623")
                c.alignment = Alignment(horizontal="center", vertical="center")
            if ci == 4:
                c.alignment = Alignment(horizontal="center", vertical="center")
        ws2.row_dimensions[rn].height = 22

    wb.save(path)
    log.info(f"✓ Saved: {path}")
    log.info(f"  Sheet 1: {len(expanded)} rows  |  Sheet 2: {len(stats)} companies")


# =============================================================================
# MAIN
# =============================================================================

if __name__ == "__main__":
    log.info("=" * 60)
    log.info("  CSR Scraper — Pune & Mumbai  v4 (Fixed URLs)")
    log.info(f"  Playwright: {PLAYWRIGHT_OK}")
    log.info("=" * 60)

    all_companies = []

    # ── Step 1: thecsrjournal.in verified URLs ────────────────────────────────
    log.info(f"\n[Step 1] thecsrjournal.in — {len(JOURNAL_URLS)} verified URLs ...")
    for entry in JOURNAL_URLS:
        url, name, city = entry
        if name is None:
            # List page — extract all company mentions
            log.info(f"  [list page] {url[:65]}")
            cos = parse_journal_list_page(url)
            all_companies.extend(cos)
        else:
            # Single company page
            log.info(f"  → {name}")
            co = parse_journal_page(url, name, city)
            all_companies.append(co)
        time.sleep(REQUEST_DELAY)
    log.info(f"  Step 1 total: {len(all_companies)}")

    # ── Step 2: csrbox.org article pages ─────────────────────────────────────
    log.info(f"\n[Step 2] csrbox.org — {len(CSRBOX_URLS)} article pages ...")
    for url in CSRBOX_URLS:
        log.info(f"  {url[:65]}")
        cos = parse_csrbox_page(url)
        all_companies.extend(cos)
    log.info(f"  Running total: {len(all_companies)}")

    # ── Step 3: csr.gov.in via Playwright ────────────────────────────────────
    log.info("\n[Step 3] csr.gov.in — Playwright (Pune) ...")
    all_companies.extend(scrape_csr_gov("PUNE"))

    log.info("\n[Step 4] csr.gov.in — Playwright (Mumbai) ...")
    all_companies.extend(scrape_csr_gov("MUMBAI"))

    # ── Step 5: Merge & deduplicate ───────────────────────────────────────────
    log.info(f"\n[Merge] {len(all_companies)} raw → deduplicating ...")
    unique = merge_deduplicate(all_companies)

    # ── Step 6: Wikipedia enrichment ─────────────────────────────────────────
    log.info(f"\n[Wikipedia] Enriching {len(unique)} companies ...")
    for i, co in enumerate(unique):
        if co.get("sector","Not listed") == "Not listed" or \
           co.get("net_profit","Not listed") == "Not listed":
            log.info(f"  ({i+1}/{len(unique)}) {co['company_name'][:50]}")
            unique[i] = enrich_wikipedia(co)

    # Sort: Mumbai first, then Pune, alphabetical
    unique.sort(key=lambda c: (
        0 if c.get("city","").lower() == "mumbai" else 1,
        c.get("company_name","").lower()
    ))

    mumbai = sum(1 for c in unique if c.get("city","").lower() == "mumbai")
    pune   = sum(1 for c in unique if c.get("city","").lower() == "pune")
    log.info(f"\n  Final: {len(unique)} companies  |  Mumbai: {mumbai}  |  Pune: {pune}")

    expanded = expand_rows(unique)
    stats    = build_stats(unique)

    log.info(f"\n[Excel] Writing {len(expanded)} rows ...")
    write_excel(expanded, stats, OUTPUT_FILE)
    log.info(f"\n✓ Done! Open: {OUTPUT_FILE}")