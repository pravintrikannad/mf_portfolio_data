"""Canara Robeco AMC scraper — downloads per-fund XLSX portfolio disclosures."""
import re
import io
import requests
import openpyxl
from typing import Optional
from .base import parse_date_from_text, clean_str, safe_float

DISCLOSURE_URL = "https://www.canararobeco.com/portfolio-disclosure"
BASE_URL = "https://www.canararobeco.com"

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Referer": "https://www.canararobeco.com/",
}

# Canara Robeco scheme codes
FUND_NAME_TO_SCHEME = {
    "flexicap":         "118275",  # Canara Robeco Flexi Cap Fund
    "bluechip":         "118274",  # Canara Robeco Blue Chip Equity Fund
    "emerging equities":"118276",  # Canara Robeco Emerging Equities Fund
    "small cap":        "118277",  # Canara Robeco Small Cap Fund
}


def _parse_sheet(ws, scheme_code: Optional[str] = None) -> dict:
    """Parse one worksheet; columns: 0=code, 1=name, 2=ISIN, 3=sector, 4=qty, 5=mkt_val, 6=pct_nav."""
    date_str = None
    holdings = []
    header_found = False

    for row in ws.iter_rows(values_only=True):
        cells = [clean_str(c) for c in row]
        text = " ".join(cells)

        if not date_str and "as on" in text.lower():
            date_str = parse_date_from_text(text)

        if not header_found:
            if any("name of the instrument" in c.lower() for c in cells):
                header_found = True
            continue

        if not any(cells):
            continue

        name   = cells[1] if len(cells) > 1 else ""
        isin   = cells[2] if len(cells) > 2 else ""
        sector = cells[3] if len(cells) > 3 else ""

        if not name:
            continue

        # Stop at footer rows
        if re.match(r'^(sub\s*total|total|grand\s*total|net receivable|margin on|treps)', name.strip(), re.IGNORECASE):
            continue

        # Skip section headers (no ISIN)
        if not isin or re.match(
            r'^(\(a\)|\(b\)|\(c\)|listed|unlisted|equity|debt|others|money market)',
            name.strip(), re.IGNORECASE
        ):
            continue

        quantity = safe_float(cells[4]) if len(cells) > 4 else None
        mkt_val  = safe_float(cells[5]) if len(cells) > 5 else None
        pct_nav  = safe_float(cells[6]) if len(cells) > 6 else None

        # Already in percentage form (6.89, not 0.0689)
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
    """Fetch latest Canara Robeco portfolio disclosures."""
    session = requests.Session()
    session.headers.update(HEADERS)

    resp = session.get(DISCLOSURE_URL, timeout=30)
    resp.raise_for_status()

    # Find XLSX links — Canara uses pattern like DV-*-April-2026.xlsx
    pattern = r'href=["\']([^"\']*\.xlsx[^"\']*)["\']'
    links = re.findall(pattern, resp.text, re.IGNORECASE)

    # Filter to current year disclosures
    xlsx_links = [l for l in links if "2026" in l or "2025" in l]
    if not xlsx_links:
        raise RuntimeError("Could not find Canara Robeco XLSX links on disclosure page")

    print(f"  Canara Robeco: found {len(xlsx_links)} files")

    schemes = []
    global_date = None

    for path in xlsx_links:
        url = (BASE_URL + path) if path.startswith("/") else path
        filename = url.split("/")[-1].split("?")[0].lower()

        # Map filename to scheme_code
        scheme_code = None
        for kw, code in FUND_NAME_TO_SCHEME.items():
            if kw in filename:
                scheme_code = code
                break

        print(f"  Canara Robeco: downloading {filename[:60]}")
        try:
            dl = session.get(url, timeout=60)
            dl.raise_for_status()
        except Exception as e:
            print(f"  Canara Robeco: WARNING — failed: {e}")
            continue

        wb = openpyxl.load_workbook(io.BytesIO(dl.content), data_only=True)
        ws = wb.active
        result = _parse_sheet(ws, scheme_code)

        if result["date"] and not global_date:
            global_date = result["date"]

        if not result["holdings"]:
            print(f"  Canara Robeco: WARNING — no holdings in {filename[:40]}")
            continue

        sheet_name = wb.sheetnames[0]
        schemes.append({
            "sheet_code":  sheet_name,
            "scheme_code": scheme_code,
            "filename":    filename,
            "holdings":    result["holdings"],
        })

    return {
        "amc":     "canara_robeco",
        "date":    global_date,
        "source":  DISCLOSURE_URL,
        "schemes": schemes,
    }
