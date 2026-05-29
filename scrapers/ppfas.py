"""PPFAS AMC scraper — downloads monthly portfolio XLSX and parses holdings."""
import re
import io
import requests
import openpyxl
from typing import Optional
from .base import parse_date_from_text, clean_str, safe_float

DISCLOSURE_URL = "https://amc.ppfas.com/downloads/portfolio-disclosure/"

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

# Known sheet-code → scheme_code mapping (from MF API)
SHEET_TO_SCHEME = {
    "PPFCF":  "122639",  # Parag Parikh Flexi Cap Fund
    "PPTSF":  "147481",  # Parag Parikh Tax Saver Fund (ELSS)
    "PPLCF":  "148742",  # Parag Parikh Large Cap Fund
    "PPCF":   "150299",  # Parag Parikh Conservative Hybrid Fund
    "PPABF":  "151014",  # Parag Parikh Arbitrage Fund
    "PPLF":   "151869",  # Parag Parikh Liquid Fund
    "PPBF":   "152500",  # Parag Parikh Balanced Advantage Fund
}


def _find_latest_xls_url(html: str) -> Optional[str]:
    """Return the first .xls/.xlsx link found under the 2026 (or latest year) folder."""
    # Prefer current year, fall back to any year
    for year in ["2026", "2025", "2024"]:
        pattern = rf'href="(/downloads/portfolio-disclosure/{year}/[^"]+\.xls[x]?)"'
        m = re.search(pattern, html, re.IGNORECASE)
        if m:
            return "https://amc.ppfas.com" + m.group(1)
    return None


def _parse_sheet(ws) -> dict:
    """Parse one sheet and return {date, holdings: []}."""
    date_str = None
    holdings = []
    header_found = False

    for row in ws.iter_rows(values_only=True):
        cells = [clean_str(c) for c in row]
        text = " ".join(cells)

        # Extract date from first few rows
        if not date_str and "as on" in text.lower():
            date_str = parse_date_from_text(text)

        # Detect column header row
        if not header_found:
            lowered = [c.lower() for c in cells]
            if any("name of the instrument" in c for c in lowered):
                header_found = True
            continue

        # Stop at totals / blank rows
        first = cells[0] if cells else ""
        if not any(cells):
            continue
        if re.match(r'^(sub\s*total|total|grand\s*total)$', first, re.IGNORECASE):
            break

        # Skip section headers (ISIN column empty, row is just a label)
        isin = cells[2] if len(cells) > 2 else ""
        if not isin or re.match(r'^(listed|unlisted|equities|debt|others|money market)', first, re.IGNORECASE):
            continue

        name      = cells[1] if len(cells) > 1 else ""
        sector    = cells[3] if len(cells) > 3 else ""
        quantity  = safe_float(cells[4]) if len(cells) > 4 else None
        mkt_val   = safe_float(cells[5]) if len(cells) > 5 else None  # in lakhs
        pct_nav   = safe_float(cells[6]) if len(cells) > 6 else None

        if not name:
            continue

        # pct_nav stored as decimal (0.0793 → 7.93%)
        if pct_nav is not None and pct_nav < 1.0:
            pct_nav = round(pct_nav * 100, 4)

        holdings.append({
            "isin":     isin,
            "name":     name,
            "sector":   sector,
            "quantity": quantity,
            "market_value_lakh": mkt_val,
            "pct_nav":  pct_nav,
        })

    return {"date": date_str, "holdings": holdings}


def fetch() -> dict:
    """Fetch latest PPFAS portfolio disclosure and return structured JSON."""
    session = requests.Session()
    session.headers.update(HEADERS)

    # Step 1: get disclosure page to find latest XLS link
    resp = session.get(DISCLOSURE_URL, timeout=30)
    resp.raise_for_status()
    xls_url = _find_latest_xls_url(resp.text)
    if not xls_url:
        raise RuntimeError("Could not find PPFAS portfolio XLS link on disclosure page")

    print(f"  PPFAS: downloading {xls_url}")

    # Step 2: download XLS (actually XLSX despite extension)
    dl = session.get(xls_url, timeout=60)
    dl.raise_for_status()

    wb = openpyxl.load_workbook(io.BytesIO(dl.content), data_only=True)

    schemes = []
    global_date = None

    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        result = _parse_sheet(ws)

        if result["date"] and not global_date:
            global_date = result["date"]

        if not result["holdings"]:
            continue

        schemes.append({
            "sheet_code":  sheet_name,
            "scheme_code": SHEET_TO_SCHEME.get(sheet_name),
            "holdings":    result["holdings"],
        })

    return {
        "amc":      "ppfas",
        "date":     global_date,
        "source":   xls_url,
        "schemes":  schemes,
    }
