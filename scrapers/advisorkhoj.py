"""
Advisorkhoj-based AMC scraper.
Fetches the latest monthly portfolio file URL for each AMC from advisorkhoj,
downloads the file (XLS/XLSX/ZIP), and parses all sheets.
"""
import re
import io
import zipfile
import datetime
import requests
import openpyxl
try:
    import xlrd
    _HAS_XLRD = True
except ImportError:
    _HAS_XLRD = False
from typing import Optional
from .base import parse_date_from_text, clean_str, safe_float


class _XlrdSheetWrapper:
    """Wraps an xlrd Sheet to mimic openpyxl's iter_rows(values_only=True)."""
    def __init__(self, sheet):
        self._sheet = sheet
        self.title = sheet.name

    def iter_rows(self, values_only=True):
        for i in range(self._sheet.nrows):
            row = self._sheet.row(i)
            yield tuple(self._cell_value(c) for c in row)

    def _cell_value(self, cell):
        import xlrd
        if cell.ctype == xlrd.XL_CELL_DATE:
            try:
                t = xlrd.xldate_as_tuple(cell.value, 0)
                return datetime.datetime(*t) if t[3:] != (0, 0, 0) else datetime.date(*t[:3])
            except Exception:
                return cell.value
        return cell.value if cell.ctype != xlrd.XL_CELL_EMPTY else None


class _XlrdWorkbookWrapper:
    """Wraps an xlrd Book to mimic openpyxl Workbook."""
    def __init__(self, book):
        self._book = book
        self.sheetnames = [book.sheet_by_index(i).name for i in range(book.nsheets)]

    def __getitem__(self, name):
        return _XlrdSheetWrapper(self._book.sheet_by_name(name))

ADVISORKHOJ_BASE = "https://www.advisorkhoj.com/form-download-centre/Mutual/{amc_slug}/Monthly-Portfolio-Disclosures"

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36",
    "Referer": "https://www.advisorkhoj.com/",
}

# AMC registry: amc_key → (advisorkhoj_slug, {sheet_code: scheme_code})
# Sheet codes come from the Index sheet of each consolidated file.
# scheme_codes verified against the DB.
AMC_REGISTRY = {
    "absl": {
        "slug": "Aditya-Birla-Sun-Life-Mutual-Fund",
        "layout": "std",   # col1=name, col2=ISIN, col3=sector, col4=qty, col5=mktval, col6=pct
    },
    "sbi": {
        "slug": "SBI-Mutual-Fund",
        "layout": "sbi",   # col2=name, col3=ISIN, col4=sector, col5=qty, col6=mktval, col7=pct
    },
    "nippon": {
        "slug": "Nippon-India-Mutual-Fund",
        "layout": "nippon",
    },
    "icici_pru": {
        "slug": "ICICI-Prudential-Mutual-Fund",
        "layout": "std",
    },
    "uti": {
        "slug": "UTI-Mutual-Fund",
        "layout": "uti",
        "filename_filter": "sebi exposure",  # pick the right file from ZIP
    },
    "dsp": {
        "slug": "DSP-Mutual-Fund",
        "layout": "std",
    },
    "franklin": {
        "slug": "Franklin-Templeton-Mutual-Fund",
        "layout": "franklin",
    },
    "tata": {
        "slug": "Tata-Mutual-Fund",
        "layout": "tata",
    },
    "motilal": {
        "slug": "Motilal-Oswal-Mutual-Fund",
        "layout": "std",
    },
    "quant": {
        "slug": "Quant-Mutual-Fund",
        "layout": "quant",
    },
    # "edelweiss": blocked — site returns 403 on direct downloads
    # "edelweiss": {
    #     "slug": "Edelweiss-Mutual-Fund",
    #     "layout": "std",
    # },
    "bajaj": {
        "slug": "Bajaj-Finserv-Mutual-Fund",
        "layout": "std",
    },
    "nj": {
        "slug": "NJ-Mutual-Fund",
        "layout": "std",
    },
    "sundaram": {
        "slug": "Sundaram-Mutual-Fund",
        "layout": "nippon",   # col0=serial, col1=ISIN, col2=name, col3=sector, col4=qty, col5=mktval, col6=pct
    },
    "trust": {
        "slug": "Trust-Mutual-Fund",
        "layout": "std",
    },
    "360one": {
        "slug": "360-ONE-Mutual-Fund",
        "layout": "nippon",   # col0=serial, col1=ISIN, col2=name — same as Sundaram/Nippon
    },
}


# ── File URL discovery ────────────────────────────────────────────────────────

def _get_latest_file_url(slug: str) -> Optional[str]:
    """Scrape advisorkhoj page and return the latest portfolio file URL."""
    url = ADVISORKHOJ_BASE.format(amc_slug=slug)
    resp = requests.get(url, headers=HEADERS, timeout=20)
    resp.raise_for_status()

    links = re.findall(r'https?://[^\s"\'<>]+(?:xls|xlsx|zip)[^\s"\'<>]*', resp.text, re.IGNORECASE)

    # Prefer most recent April 2026 link, then any 2026, then newest available
    for priority in [
        lambda l: ("april" in l.lower() or "apr" in l.lower()) and ("2026" in l or "-26" in l.lower()),
        lambda l: "2026" in l or "-26" in l.lower(),
        lambda l: True,
    ]:
        matches = [l for l in links if priority(l)]
        if matches:
            return matches[0]
    return None


# ── File download + unpack ────────────────────────────────────────────────────

def _open_workbook(data: bytes, filename: str):
    """Try openpyxl (xlsx) then xlrd (xls). Returns workbook-like object or None."""
    if filename.lower().endswith('.xlsx') or data[:2] == b'PK':
        try:
            return openpyxl.load_workbook(io.BytesIO(data), data_only=True)
        except Exception:
            pass
    if _HAS_XLRD:
        try:
            book = xlrd.open_workbook(file_contents=data)
            return _XlrdWorkbookWrapper(book)
        except Exception:
            pass
    # Fallback: try openpyxl regardless
    try:
        return openpyxl.load_workbook(io.BytesIO(data), data_only=True)
    except Exception:
        return None


def _download_workbook(url: str) -> list[openpyxl.Workbook]:
    """Download URL and return list of openpyxl Workbooks (ZIP may contain multiple)."""
    try:
        resp = requests.get(url, headers=HEADERS, timeout=120)
    except Exception:
        # Retry without SSL verification (some AMC sites have cert issues)
        import warnings
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            resp = requests.get(url, headers=HEADERS, timeout=120, verify=False)
    resp.raise_for_status()
    content = resp.content

    fname = url.split('/')[-1].split('?')[0].lower()

    # Try direct workbook first (XLSX files also have PK header — must check before ZIP)
    if fname.endswith('.xlsx') or fname.endswith('.xls'):
        wb = _open_workbook(content, fname)
        if wb:
            return [(fname, wb)]

    # ZIP file containing one or more workbooks
    if content[:2] == b'PK' and content[2:4] in (b'\x03\x04', b'\x05\x06', b'\x07\x08'):
        try:
            z = zipfile.ZipFile(io.BytesIO(content))
            workbooks = []
            for name in z.namelist():
                if name.lower().endswith(('.xlsx', '.xls')) and not name.startswith('__'):
                    data = z.read(name)
                    wb = _open_workbook(data, name)
                    if wb:
                        workbooks.append((name, wb))
                    else:
                        print(f"    Skipping {name}: unsupported format")
            if workbooks:
                return workbooks
        except zipfile.BadZipFile:
            pass

    # Last resort: try as workbook regardless of extension
    wb = _open_workbook(content, fname)
    if wb:
        return [(fname, wb)]
    raise RuntimeError(f"Could not open file from {url}")


# ── Sheet parsers ─────────────────────────────────────────────────────────────

def _parse_std(ws) -> tuple[Optional[str], list]:
    """Standard layout: col1=name, col2=ISIN, col3=sector, col4=qty, col5=mktval, col6=pct."""
    date_str = None
    holdings = []
    header_found = False

    for row in ws.iter_rows(values_only=True):
        cells = [clean_str(c) for c in row]
        text = " ".join(cells)

        if not date_str and "as on" in text.lower():
            date_str = parse_date_from_text(text)

        if not header_found:
            if any(
                ("name of" in c.lower() and ("instrument" in c.lower() or "issuer" in c.lower()))
                or ("company" in c.lower() and ("issuer" in c.lower() or "instrument" in c.lower()))
                or c.lower() in ("name of the instrument", "company/issuer/instrument name",
                                 "name of instruments", "name of issuer")
                for c in cells
            ):
                header_found = True
            continue

        if not any(cells):
            continue

        name   = cells[1] if len(cells) > 1 else ""
        isin   = cells[2] if len(cells) > 2 else ""
        sector = cells[3] if len(cells) > 3 else ""

        if not name or not isin:
            continue
        if re.match(r"^(sub\s*total|total|grand\s*total|net receivable|margin|treps)", name.strip(), re.IGNORECASE):
            continue
        if re.match(r"^(\(a\)|\(b\)|\(c\)|listed|unlisted|equity|debt|others|money market|[ivxlcdm]+\))", name.strip(), re.IGNORECASE):
            continue

        pct = safe_float(cells[6]) if len(cells) > 6 else None
        holdings.append({
            "isin": isin, "name": name, "sector": sector,
            "quantity": safe_float(cells[4]) if len(cells) > 4 else None,
            "market_value_lakh": safe_float(cells[5]) if len(cells) > 5 else None,
            "pct_nav": pct, "_r": pct,
        })

    _fix_decimal_pct(holdings)
    return date_str, holdings


def _parse_sbi(ws) -> tuple[Optional[str], list]:
    """SBI layout: col2=name, col3=ISIN, col4=sector, col5=qty, col6=mktval, col7=pct."""
    date_str = None
    holdings = []

    for row in ws.iter_rows(values_only=True):
        cells = [clean_str(c) for c in row]
        text = " ".join(cells)

        if not date_str:
            # SBI stores date as datetime object in col3
            for c in row:
                if isinstance(c, datetime.datetime):
                    date_str = c.strftime("%Y-%m-%d")
                    break
            if not date_str and "as on" in text.lower():
                date_str = parse_date_from_text(text)

        if len(cells) < 8:
            continue

        name   = cells[2]
        isin   = cells[3]
        sector = cells[4]

        if not name or not isin or not isin.startswith("IN"):
            continue
        if re.match(r"^(sub\s*total|total|grand\s*total|net receivable|margin|treps)", name.strip(), re.IGNORECASE):
            continue
        if re.match(r"^([a-z]\)|\(a\)|\(b\)|listed|unlisted|equity|debt|others)", name.strip(), re.IGNORECASE):
            continue

        pct = safe_float(cells[7])
        holdings.append({
            "isin": isin, "name": name, "sector": sector,
            "quantity": safe_float(cells[5]),
            "market_value_lakh": safe_float(cells[6]),
            "pct_nav": pct, "_r": pct,
        })

    _fix_decimal_pct(holdings)
    return date_str, holdings


def _fix_decimal_pct(holdings: list):
    """Auto-detect if pct_nav is decimal (0.0639) vs plain % (6.39) and fix."""
    raw = [h["_r"] for h in holdings if h["_r"] is not None]
    if raw and all(p < 1.5 for p in raw):
        for h in holdings:
            if h["pct_nav"] is not None:
                h["pct_nav"] = round(h["pct_nav"] * 100, 4)
    for h in holdings:
        h.pop("_r", None)


def _parse_tata(ws) -> tuple[Optional[str], list]:
    """Tata layout: col1=name, col2=yield, col3=sector, col4=ISIN, col5=qty, col6=mktval, col7=pct."""
    date_str = None
    holdings = []
    header_found = False

    for row in ws.iter_rows(values_only=True):
        cells = [clean_str(c) for c in row]
        text = " ".join(cells)

        if not date_str and "as on" in text.lower():
            date_str = parse_date_from_text(text)

        if not header_found:
            if (len(cells) > 4 and
                    any("name of" in c.lower() and "instrument" in c.lower() for c in cells) and
                    any("isin" in c.lower() for c in cells) and
                    any("quantity" in c.lower() for c in cells)):
                header_found = True
            continue

        if not any(cells):
            continue

        name   = cells[1] if len(cells) > 1 else ""
        isin   = cells[4] if len(cells) > 4 else ""
        sector = cells[3] if len(cells) > 3 else ""

        if not name or not isin or not isin.startswith("IN"):
            continue
        if re.match(r"^(sub\s*total|total|grand\s*total|net receivable|margin|treps)", name.strip(), re.IGNORECASE):
            continue
        if re.match(r"^([a-z]\)|\(a\)|\(b\)|listed|unlisted|equity|debt|others)", name.strip(), re.IGNORECASE):
            continue

        pct = safe_float(cells[7]) if len(cells) > 7 else None
        holdings.append({
            "isin": isin, "name": name, "sector": sector,
            "quantity": safe_float(cells[5]) if len(cells) > 5 else None,
            "market_value_lakh": safe_float(cells[6]) if len(cells) > 6 else None,
            "pct_nav": pct, "_r": pct,
        })

    _fix_decimal_pct(holdings)
    return date_str, holdings


def _parse_quant(ws) -> tuple[Optional[str], list]:
    """Quant layout: col0=serial, col1=ISIN, col2=name, col3=rating, col4=sector, col5=qty, col6=mktval, col7=pct."""
    date_str = None
    holdings = []
    header_found = False

    for row in ws.iter_rows(values_only=True):
        cells = [clean_str(c) for c in row]
        text = " ".join(cells)

        if not date_str and "as on" in text.lower():
            date_str = parse_date_from_text(text)

        if not header_found:
            if len(cells) > 2 and cells[1] and cells[1].lower().strip() in ("isin", "isin no", "isin number", "isin code", "isin number/code"):
                header_found = True
            continue

        if not any(cells):
            continue

        isin   = cells[1] if len(cells) > 1 else ""
        name   = cells[2] if len(cells) > 2 else ""
        sector = cells[4] if len(cells) > 4 else ""

        if not name or not isin or not isin.startswith("IN"):
            continue
        if re.match(r"^(sub\s*total|total|grand\s*total|net receivable|margin|treps)", name.strip(), re.IGNORECASE):
            continue
        if re.match(r"^(\(a\)|\(b\)|\(c\)|listed|unlisted|equity|debt|others|money market)", name.strip(), re.IGNORECASE):
            continue

        pct = safe_float(cells[7]) if len(cells) > 7 else None
        holdings.append({
            "isin": isin, "name": name, "sector": sector,
            "quantity": safe_float(cells[5]) if len(cells) > 5 else None,
            "market_value_lakh": safe_float(cells[6]) if len(cells) > 6 else None,
            "pct_nav": pct, "_r": pct,
        })

    _fix_decimal_pct(holdings)
    return date_str, holdings


def _parse_franklin(ws) -> tuple[Optional[str], list]:
    """Franklin layout: col0=ISIN, col1=name, col2=rating/sector, col3=qty, col4=mktval, col5=pct."""
    date_str = None
    holdings = []
    header_found = False

    for row in ws.iter_rows(values_only=True):
        cells = [clean_str(c) for c in row]
        text = " ".join(cells)

        if not date_str and "as on" in text.lower():
            date_str = parse_date_from_text(text)

        if not header_found:
            if cells[0].lower() in ("isin number", "isin", "isin no"):
                header_found = True
            continue

        if not any(cells):
            continue

        isin   = cells[0] if len(cells) > 0 else ""
        name   = cells[1] if len(cells) > 1 else ""
        sector = cells[2] if len(cells) > 2 else ""

        if not name or not isin or not isin.startswith("IN"):
            continue
        if re.match(r"^(sub\s*total|total|grand\s*total|net receivable|margin|treps)", name.strip(), re.IGNORECASE):
            continue

        pct = safe_float(cells[5]) if len(cells) > 5 else None
        holdings.append({
            "isin": isin, "name": name, "sector": sector,
            "quantity": safe_float(cells[3]) if len(cells) > 3 else None,
            "market_value_lakh": safe_float(cells[4]) if len(cells) > 4 else None,
            "pct_nav": pct, "_r": pct,
        })

    _fix_decimal_pct(holdings)
    return date_str, holdings


def _parse_nippon(ws) -> tuple[Optional[str], list]:
    """Nippon layout: col0=marker, col1=ISIN, col2=name, col3=sector, col4=qty, col5=mktval, col6=pct."""
    date_str = None
    holdings = []
    header_found = False

    for row in ws.iter_rows(values_only=True):
        cells = [clean_str(c) for c in row]
        text = " ".join(cells)

        if not date_str and "as on" in text.lower():
            date_str = parse_date_from_text(text)

        if not header_found:
            if len(cells) > 2 and cells[1] and cells[1].lower() in ("isin", "isin no", "isin number", "isin code", "isin number/code"):
                header_found = True
            continue

        if not any(cells):
            continue

        isin   = cells[1] if len(cells) > 1 else ""
        name   = cells[2] if len(cells) > 2 else ""
        sector = cells[3] if len(cells) > 3 else ""

        if not name or not isin or not isin.startswith("IN"):
            continue
        if re.match(r"^(sub\s*total|total|grand\s*total|net receivable|margin|treps)", name.strip(), re.IGNORECASE):
            continue
        if re.match(r"^(\(a\)|\(b\)|\(c\)|listed|unlisted|equity|debt|others|money market|[ivxlcdm]+\))", name.strip(), re.IGNORECASE):
            continue

        pct = safe_float(cells[6]) if len(cells) > 6 else None
        holdings.append({
            "isin": isin, "name": name, "sector": sector,
            "quantity": safe_float(cells[4]) if len(cells) > 4 else None,
            "market_value_lakh": safe_float(cells[5]) if len(cells) > 5 else None,
            "pct_nav": pct, "_r": pct,
        })

    _fix_decimal_pct(holdings)
    return date_str, holdings


def _parse_uti_sheet(ws) -> list[dict]:
    """
    UTI puts all schemes in one sheet separated by 'SCHEME CODE002STARTS' markers.
    Layout: col0=name, col1=sector, col2=qty, col3=mktval, col4=pct, col7=ISIN
    Returns list of {scheme_name, date, holdings} dicts.
    """
    schemes = []
    cur_scheme = None
    cur_date = None
    cur_holdings = []
    in_holdings = False

    for row in ws.iter_rows(values_only=True):
        cells = [clean_str(c) for c in row]
        if not any(cells):
            continue
        text = " ".join(cells)

        c0 = cells[0].upper()
        if re.match(r"SCHEME\s*CODE\d*STARTS", c0):
            # Save previous scheme
            if cur_scheme and cur_holdings:
                _fix_decimal_pct(cur_holdings)
                schemes.append({"scheme_name": cur_scheme, "date": cur_date, "holdings": cur_holdings})
            cur_scheme = None
            cur_date = None
            cur_holdings = []
            in_holdings = False
            continue
        if re.match(r"SCHEME\s*CODE\d*ENDS", c0):
            continue

        if "SCHEME:" in cells[0].upper():
            cur_scheme = cells[0].replace("SCHEME:", "").strip()
            continue

        if not cur_date and "as of" in text.lower():
            m = re.search(r'as of\s+(\d{2}/\d{2}/\d{4})', text, re.IGNORECASE)
            if m:
                try:
                    cur_date = datetime.datetime.strptime(m.group(1), "%d/%m/%Y").strftime("%Y-%m-%d")
                except ValueError:
                    pass

        if cells[0].upper() == "NAME OF THE INSTRUMENT":
            in_holdings = True
            continue

        if not in_holdings:
            continue

        isin   = cells[7] if len(cells) > 7 else ""
        name   = cells[0]
        sector = cells[1] if len(cells) > 1 else ""

        if not isin or not isin.startswith("IN") or not name:
            continue
        # Strip "EQ - " prefix on names
        name = re.sub(r'^EQ\s*-\s*', '', name).strip()
        if re.match(r"^(sub\s*total|total|grand\s*total|net receivable|treps)", name, re.IGNORECASE):
            continue

        pct = safe_float(cells[4]) if len(cells) > 4 else None
        cur_holdings.append({
            "isin": isin, "name": name, "sector": sector,
            "quantity": safe_float(cells[2]) if len(cells) > 2 else None,
            "market_value_lakh": safe_float(cells[3]) if len(cells) > 3 else None,
            "pct_nav": pct, "_r": pct,
        })

    # Don't forget the last scheme
    if cur_scheme and cur_holdings:
        _fix_decimal_pct(cur_holdings)
        schemes.append({"scheme_name": cur_scheme, "date": cur_date, "holdings": cur_holdings})

    return schemes


LAYOUT_PARSERS = {
    "std": _parse_std,
    "sbi": _parse_sbi,
    "nippon": _parse_nippon,
    "franklin": _parse_franklin,
    "tata": _parse_tata,
    "quant": _parse_quant,
    "uti": None,  # handled specially in fetch_amc
}


# ── Index sheet reader ────────────────────────────────────────────────────────

def _extract_sheet_scheme_name(ws) -> str:
    """Try to extract scheme name from top rows of a sheet (before the holdings header)."""
    for row in ws.iter_rows(values_only=True):
        cells = [clean_str(c) for c in row]
        if not any(cells):
            continue
        text = " ".join(cells).strip()
        # Skip obvious non-names
        if len(text) < 5:
            continue
        if re.match(r"^(monthly|portfolio|statement|as on|name of|isin|industry|quantity|market)", text, re.IGNORECASE):
            break  # hit the header, stop looking
        # Skip single codes (like "BFARB")
        if re.match(r"^[A-Z0-9]{2,10}$", text):
            continue
        # Find first non-empty cell that looks like a fund name (not just AMC name)
        for c in cells:
            if (c and len(c) > 10 and not c.startswith("(")
                    and re.search(r"fund|scheme|etf|fof", c, re.IGNORECASE)
                    and not re.search(r"^.{0,30}\s+mutual\s+fund\s*$", c, re.IGNORECASE)):
                return c
    return ""


def _read_index(wb) -> dict[str, str]:
    """Read Index sheet → {sheet_code: scheme_name}.

    Handles multiple formats:
    - ABSL/Tata/SBI:  cells[1]=code, cells[2]=name
    - Nippon:         cells[0]=code, cells[1]=name (cells[2] empty)
    - Motilal:        cells[2]=serial, cells[3]=name, cells[4]=code
    """
    sheet = None
    for n in wb.sheetnames:
        if n.lower() == "index":
            sheet = wb[n]
            break
    if not sheet:
        return {}
    index = {}
    for row in sheet.iter_rows(values_only=True):
        cells = [clean_str(c) for c in row]
        skip = {"Scheme Short code", "INDEX", "CLASSIFICATION", "Sr No.", ""}

        # Standard: cells[1]=code, cells[2]=name
        if len(cells) >= 3 and cells[1] and cells[2] and cells[1] not in skip:
            index[cells[1]] = cells[2]

        # Motilal: cells[0]=None, cells[1]=None, cells[2]=serial, cells[3]=name, cells[4]=code
        elif len(cells) >= 5 and not cells[0] and not cells[1] and cells[3] and cells[4] and cells[4] not in skip:
            index[cells[4]] = cells[3]

        # Nippon: cells[0]=code, cells[1]=name (cells[2] empty)
        elif len(cells) >= 2 and cells[0] and cells[1] and cells[0] not in skip:
            if cells[0] not in index:
                index[cells[0]] = cells[1]

    return index


# ── Main fetch function ───────────────────────────────────────────────────────

def fetch_amc(amc_key: str) -> Optional[dict]:
    config = AMC_REGISTRY.get(amc_key)
    if not config:
        raise ValueError(f"Unknown AMC key: {amc_key}")

    slug            = config["slug"]
    layout          = config.get("layout", "std")
    filename_filter = config.get("filename_filter", "").lower()

    print(f"  [{amc_key}] Finding file URL via advisorkhoj...")
    file_url = _get_latest_file_url(slug)
    if not file_url:
        raise RuntimeError(f"[{amc_key}] Could not find portfolio file URL on advisorkhoj")

    print(f"  [{amc_key}] Downloading: {file_url.split('?')[0][-70:]}")
    workbooks = _download_workbook(file_url)
    if filename_filter:
        workbooks = [(fn, wb) for fn, wb in workbooks if filename_filter in fn.lower()]
    print(f"  [{amc_key}] Opened {len(workbooks)} workbook(s)")

    all_schemes = []
    global_date = None

    # ── UTI: special multi-scheme single-sheet format ──
    if layout == "uti":
        for filename, wb in workbooks:
            for sheet_name in wb.sheetnames:
                ws = wb[sheet_name]
                schemes = _parse_uti_sheet(ws)
                for s in schemes:
                    if not s["holdings"]:
                        continue
                    if s["date"] and not global_date:
                        global_date = s["date"]
                    all_schemes.append({
                        "sheet_code":  sheet_name,
                        "scheme_code": None,
                        "scheme_name": s["scheme_name"],
                        "holdings":    s["holdings"],
                    })
        print(f"  [{amc_key}] Parsed {len(all_schemes)} schemes, date={global_date}")
        return {"amc": amc_key, "date": global_date, "source": file_url, "schemes": all_schemes}

    parser = LAYOUT_PARSERS[layout]

    for filename, wb in workbooks:
        index = _read_index(wb)
        # For per-fund ZIP files (e.g. ICICI Pru), the fund name is in the filename
        fname_scheme = re.sub(r'\.xlsx?$', '', filename, flags=re.IGNORECASE).strip()

        for sheet_name in wb.sheetnames:
            if sheet_name.lower() in ("index", "notes", "disclaimer", "common notes"):
                continue

            ws = wb[sheet_name]
            date_str, holdings = parser(ws)

            if not holdings:
                continue
            if date_str and not global_date:
                global_date = date_str

            # scheme_name: prefer Index sheet → per-fund filename → extract from sheet content
            scheme_name = (index.get(sheet_name)
                           or (fname_scheme if len(workbooks) > 1 else "")
                           or _extract_sheet_scheme_name(ws))
            all_schemes.append({
                "sheet_code":   sheet_name,
                "scheme_code":  None,          # resolved server-side via scheme name match
                "scheme_name":  scheme_name,
                "holdings":     holdings,
            })

    print(f"  [{amc_key}] Parsed {len(all_schemes)} schemes, date={global_date}")
    return {
        "amc":     amc_key,
        "date":    global_date,
        "source":  file_url,
        "schemes": all_schemes,
    }
