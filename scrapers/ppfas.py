"""PPFAS AMC scraper — downloads per-fund XLSX files and parses holdings."""
import re
import io
import requests
import openpyxl
from typing import Optional
from .base import parse_date_from_text, clean_str, safe_float

DISCLOSURE_URL = "https://amc.ppfas.com/downloads/portfolio-disclosure/"
BASE_URL = "https://amc.ppfas.com"

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "en-US,en;q=0.5",
    "Referer": "https://amc.ppfas.com/",
}

# sheet_code prefix → scheme_code (from MF API)
SHEET_TO_SCHEME = {
    "PPFCF":  "122639",  # Parag Parikh Flexi Cap Fund
    "PPTSF":  "147481",  # Parag Parikh Tax Saver Fund (ELSS)
    "PPLCF":  "148742",  # Parag Parikh Large Cap Fund
    "PPCHF":  "150299",  # Parag Parikh Conservative Hybrid Fund
    "PPAF":   "151014",  # Parag Parikh Arbitrage Fund
    "PPLF":   "151869",  # Parag Parikh Liquid Fund
    "PPDAAF": "152500",  # Parag Parikh Dynamic Asset Allocation Fund
}


def _find_xlsx_urls(html: str) -> list[tuple[str, str]]:
    """
    Return list of (sheet_code, full_url) for the latest month's individual XLSX files.
    URLs look like: /downloads/portfolio-disclosure/2026/PPFCF_PPFAS_Monthly_...xlsx?...
    We want per-fund files (prefixed with sheet code), not the combined .xls file.
    """
    results = []
    seen_codes = set()

    # Find all .xlsx links (per-fund files have a sheet_code prefix before _PPFAS_)
    pattern = r'href=["\']([^"\']*portfolio-disclosure/(\d{4})/([A-Z]+)_PPFAS_[^"\']*\.xlsx[^"\']*)["\']'
    for year in ["2026", "2025", "2024"]:
        matches = re.findall(pattern, html, re.IGNORECASE)
        for path, yr, code in matches:
            if yr != year:
                continue
            if code in seen_codes:
                continue
            seen_codes.add(code)
            url = BASE_URL + path if path.startswith("/") else path
            results.append((code, url))

        if results:
            break

    return results


def _parse_sheet(ws) -> dict:
    """Parse one worksheet and return {date, holdings: []}."""
    date_str = None
    holdings = []
    header_found = False

    for row in ws.iter_rows(values_only=True):
        cells = [clean_str(c) for c in row]
        text = " ".join(cells)

        # Extract date from header rows
        if not date_str and "as on" in text.lower():
            date_str = parse_date_from_text(text)

        # Detect column header row
        if not header_found:
            if any("name of the instrument" in c.lower() for c in cells):
                header_found = True
            continue

        # Skip blank rows
        if not any(cells):
            continue

        # Columns (0-indexed): 0=code, 1=name, 2=ISIN, 3=sector/rating, 4=quantity, 5=market_value, 6=pct_nav
        name   = cells[1] if len(cells) > 1 else ""
        isin   = cells[2] if len(cells) > 2 else ""
        sector = cells[3] if len(cells) > 3 else ""

        if not name:
            continue

        # Stop at Sub Total / Total rows
        if re.match(r'^(sub\s*total|total|grand\s*total)$', name.strip(), re.IGNORECASE):
            continue

        # Skip section headers (name is a label, ISIN is empty)
        if not isin or re.match(
            r'^(\(a\)|\(b\)|\(c\)|listed|unlisted|equity|debt|others|money market)',
            name.strip(), re.IGNORECASE
        ):
            continue

        quantity = safe_float(cells[4]) if len(cells) > 4 else None
        mkt_val  = safe_float(cells[5]) if len(cells) > 5 else None   # Rs. in Lakhs
        pct_nav  = safe_float(cells[6]) if len(cells) > 6 else None

        # pct_nav stored as decimal (0.0793 → 7.93%)
        if pct_nav is not None and pct_nav < 1.0:
            pct_nav = round(pct_nav * 100, 4)

        holdings.append({
            "isin":              isin,
            "name":              name,
            "sector":            sector,
            "quantity":          quantity,
            "market_value_lakh": mkt_val,
            "pct_nav":           pct_nav,
        })

    return {"date": date_str, "holdings": holdings}


def fetch() -> dict:
    """Fetch latest PPFAS portfolio disclosures and return structured JSON."""
    session = requests.Session()
    session.headers.update(HEADERS)

    # Step 1: get disclosure page to find per-fund XLSX links
    resp = session.get(DISCLOSURE_URL, timeout=30)
    resp.raise_for_status()
    xlsx_urls = _find_xlsx_urls(resp.text)

    if not xlsx_urls:
        raise RuntimeError("Could not find any PPFAS per-fund XLSX links on disclosure page")

    print(f"  PPFAS: found {len(xlsx_urls)} fund files")

    schemes = []
    global_date = None

    for sheet_code, url in xlsx_urls:
        print(f"  PPFAS: downloading {sheet_code} from {url.split('?')[0]}")
        try:
            dl = session.get(url, timeout=60)
            dl.raise_for_status()
        except Exception as e:
            print(f"  PPFAS: WARNING — failed to download {sheet_code}: {e}")
            continue

        wb = openpyxl.load_workbook(io.BytesIO(dl.content), data_only=True)
        ws = wb.active
        result = _parse_sheet(ws)

        if result["date"] and not global_date:
            global_date = result["date"]

        if not result["holdings"]:
            print(f"  PPFAS: WARNING — no holdings parsed for {sheet_code}")
            continue

        schemes.append({
            "sheet_code":  sheet_code,
            "scheme_code": SHEET_TO_SCHEME.get(sheet_code),
            "holdings":    result["holdings"],
        })

    return {
        "amc":     "ppfas",
        "date":    global_date,
        "source":  DISCLOSURE_URL,
        "schemes": schemes,
    }
