"""
Manual file parser — called by watch_downloads.py.
Detects file layout and routes to the right parsing logic.
"""
import re
import io
import datetime
from pathlib import Path
import openpyxl
from scrapers.base import parse_date_from_text, clean_str, safe_float


# ── Layout detectors ──────────────────────────────────────────────────────────

def _peek_rows(ws, n=15):
    rows = []
    for i, row in enumerate(ws.iter_rows(values_only=True)):
        rows.append([clean_str(c) for c in row])
        if i >= n:
            break
    return rows


def _detect_layout(ws) -> str:
    """
    Returns one of:
      'std'    — col1=name, col2=ISIN, col3=sector, col4=qty, col5=mktval, col6=pct
      'hdfc'   — col0=marker, col1=ISIN, col3=name, col4=sector, col5=qty, col6=mktval, col7=pct
      'kotak'  — col0=section, col1=space, col2=name, col3=ISIN, col4=sector, col6=qty, col7=mktval, col8=pct
    """
    rows = _peek_rows(ws)
    for cells in rows:
        # HDFC: ISIN in col1, name in col3
        if len(cells) > 3 and cells[1].startswith("IN") and cells[3] and not cells[0].startswith("IN"):
            return "hdfc"
        # Kotak: ISIN in col3, name in col2, pct in col8
        if len(cells) > 8 and cells[3].startswith("IN") and cells[2] and cells[8]:
            try:
                float(str(cells[8]).replace(",", ""))
                return "kotak"
            except (ValueError, TypeError):
                pass
        # Std: ISIN in col2, name in col1
        if len(cells) > 2 and cells[2].startswith("IN") and cells[1]:
            return "std"
    return "std"


# ── Per-layout parsers ────────────────────────────────────────────────────────

def _parse_std(ws):
    """Standard layout: PPFAS, Canara Robeco, PGIM."""
    date_str = None
    holdings = []
    header_found = False

    for row in ws.iter_rows(values_only=True):
        cells = [clean_str(c) for c in row]
        text  = " ".join(cells)

        if not date_str and "as on" in text.lower():
            date_str = parse_date_from_text(text)

        if not header_found:
            if any("name of" in c.lower() and "instrument" in c.lower() for c in cells):
                header_found = True
            continue

        if not any(cells):
            continue

        name   = cells[1] if len(cells) > 1 else ""
        isin   = cells[2] if len(cells) > 2 else ""
        sector = cells[3] if len(cells) > 3 else ""

        if not name or not isin:
            continue
        if re.match(r"^(sub\s*total|total|grand\s*total|net receivable|margin on|treps)", name.strip(), re.IGNORECASE):
            continue
        if re.match(r"^(\(a\)|\(b\)|\(c\)|listed|unlisted|equity|debt|others|money market)", name.strip(), re.IGNORECASE):
            continue

        pct_nav = safe_float(cells[6]) if len(cells) > 6 else None
        # PPFAS uses decimal (0.0793); everyone else uses plain % (6.39)
        # Heuristic: if ALL non-None pct values < 1.0, it's decimal format
        holdings.append({
            "isin":              isin,
            "name":              name,
            "sector":            sector,
            "quantity":          safe_float(cells[4]) if len(cells) > 4 else None,
            "market_value_lakh": safe_float(cells[5]) if len(cells) > 5 else None,
            "pct_nav":           pct_nav,
            "_raw_pct":          pct_nav,
        })

    # Auto-detect decimal vs percent
    raw_pcts = [h["_raw_pct"] for h in holdings if h["_raw_pct"] is not None]
    if raw_pcts and all(p < 1.5 for p in raw_pcts):
        for h in holdings:
            if h["pct_nav"] is not None:
                h["pct_nav"] = round(h["pct_nav"] * 100, 4)
    for h in holdings:
        h.pop("_raw_pct", None)

    return date_str, holdings


def _parse_hdfc(ws):
    """HDFC layout: ISIN col1, name col3, pct col7."""
    date_str = None
    holdings = []

    for row in ws.iter_rows(values_only=True):
        cells = [clean_str(c) for c in row]
        text  = " ".join(cells)

        if not date_str:
            m = re.search(r"portfolio as on\s+(\d{1,2}-\w{3}-\d{4})", text, re.IGNORECASE)
            if m:
                try:
                    date_str = datetime.datetime.strptime(m.group(1), "%d-%b-%Y").strftime("%Y-%m-%d")
                except ValueError:
                    pass
            if not date_str and "as on" in text.lower():
                date_str = parse_date_from_text(text)

        if len(cells) < 8:
            continue

        isin   = cells[1]
        name   = cells[3]
        sector = cells[4]

        if not isin or not isin.startswith("IN") or not name:
            continue
        if re.match(r"^(sub\s*total|total|grand\s*total)", name.strip(), re.IGNORECASE):
            continue

        holdings.append({
            "isin":              isin,
            "name":              name,
            "sector":            sector,
            "quantity":          safe_float(cells[5]),
            "market_value_lakh": safe_float(cells[6]),
            "pct_nav":           safe_float(cells[7]),
        })

    return date_str, holdings


def _parse_kotak_sheet(ws):
    """Kotak layout: name col2, ISIN col3, pct col8."""
    date_str = None
    holdings = []

    for row in ws.iter_rows(values_only=True):
        cells = [clean_str(c) for c in row]
        text  = " ".join(cells)

        if not date_str and "as on" in text.lower():
            date_str = parse_date_from_text(text)

        if len(cells) < 9:
            continue

        name   = cells[2]
        isin   = cells[3]
        sector = cells[4]

        if not isin or not isin.startswith("IN") or not name:
            continue
        if re.match(r"^(sub\s*total|total|grand\s*total)", name.strip(), re.IGNORECASE):
            continue

        holdings.append({
            "isin":              isin,
            "name":              name,
            "sector":            sector,
            "quantity":          safe_float(cells[6]),
            "market_value_lakh": safe_float(cells[7]),
            "pct_nav":           safe_float(cells[8]),
        })

    return date_str, holdings


# ── Scheme code lookup ────────────────────────────────────────────────────────

# Kotak sheet → scheme_code
KOTAK_SHEET_TO_SCHEME = {
    "SEF": "120166", "NEF": "119775", "KBC": "120152",
    "MSC": "120164", "BAL": "133035", "NIF": "148978",
    "MCF": "149185", "MID": "120158", "ELS": "119773",
}

# Canara Robeco filename keyword → scheme_code
CANARA_SCHEME = {
    "flexicap": "118275", "flexi cap": "118275", "flexi-cap": "118275",
    "bluechip": "118274", "blue chip": "118274",
    "small cap": "118277", "smallcap": "118277",
    "emerging":  "118276",
}

PGIM_SCHEME = {
    "flexi cap": "133839", "flexi-cap": "133839",
    "midcap":    "118668",
    "large cap": "118669",
}

HDFC_SCHEME = {
    "flexi cap":                        "118955",
    "small cap":                        "130503",
    "mid cap":                          "118989",
    "large cap":                        "119018",
    "focused":                          "118950",
    "multi cap":                        "149368",
    "elss tax saver":                   "119060",
    "elss tax":                         "119060",
    "hybrid equity":                    "119062",
    "hybrid debt":                      "119118",
    "multi-asset":                      "119131",
    "arbitrage":                        "118931",
    "corporate bond":                   "118987",
    "credit risk":                      "128051",
    "short term  debt":                 "119016",
    "short term debt":                  "119016",
    "medium term":                      "119081",
    "dynamic debt":                     "119075",
    "income fund":                      "119069",
    "gilt fund":                        "119116",
    "liquid fund":                      "119091",
    "money market":                     "119092",
    "overnight":                        "119110",
    "ultra short term":                 "145034",
    "long duration":                    "151313",
    "bse 500 index":                    "151728",
    "bse india sector leaders":         "153959",
    "crisil-ibx financial services 3-6":"153517",
    "crisil-ibx financial services 9-12":"154308",
    "nifty 100 index":                  "149868",
    "nifty 100 equal weight":           "149870",
    "nifty50 equal weight":             "149107",
    "nifty next 50":                    "149288",
    "nifty midcap 150":                 "151724",
    "nifty smallcap 250":               "151727",
    "nifty largemidcap 250":            "152889",
    "nifty g-sec jun 2027":             "151181",
    "nifty g-sec dec 2026":             "150845",
    "nifty g-sec apr 2029":             "151495",
    "nifty g-sec jul 2031":             "150847",
    "nifty g-sec sep 2032":             "151183",
    "nifty g-sec jun 2036":             "151489",
    "nifty sdl oct 2026":               "151456",
    "nifty sdl plus g-sec jun 2027":    "151570",
    "nifty india consumption":          "154179",
    "nifty india digital":              "153097",
    "developed world":                  "149180",
}


def _parse_edel(ws):
    """Edelweiss layout: col0=name, col1=ISIN, col2=sector, col3=qty, col4=mktval, col5=pct."""
    date_str = None
    holdings = []
    header_found = False

    for row in ws.iter_rows(values_only=True):
        cells = [clean_str(c) for c in row]
        text  = " ".join(cells)

        if not date_str and "as on" in text.lower():
            date_str = parse_date_from_text(text)

        if not header_found:
            if cells[0] and "name of" in cells[0].lower() and (
                    "instrument" in cells[0].lower() or "issuer" in cells[0].lower()):
                header_found = True
            continue

        if not any(cells):
            continue

        name   = cells[0] if len(cells) > 0 else ""
        isin   = cells[1] if len(cells) > 1 else ""
        sector = cells[2] if len(cells) > 2 else ""

        if not name or not isin:
            continue
        if not isin.startswith("IN"):
            continue
        if re.match(r"^(sub\s*total|total|grand\s*total|net receivable|margin|treps)", name.strip(), re.IGNORECASE):
            continue
        if re.match(r"^(\(a\)|\(b\)|\(c\)|listed|unlisted|equity|debt|others|money market)", name.strip(), re.IGNORECASE):
            continue

        pct_nav = safe_float(cells[5]) if len(cells) > 5 else None
        holdings.append({
            "isin":              isin,
            "name":              name,
            "sector":            sector,
            "quantity":          safe_float(cells[3]) if len(cells) > 3 else None,
            "market_value_lakh": safe_float(cells[4]) if len(cells) > 4 else None,
            "pct_nav":           pct_nav,
            "_raw_pct":          pct_nav,
        })

    raw_pcts = [h["_raw_pct"] for h in holdings if h["_raw_pct"] is not None]
    if raw_pcts and all(p < 1.5 for p in raw_pcts):
        for h in holdings:
            if h["pct_nav"] is not None:
                h["pct_nav"] = round(h["pct_nav"] * 100, 4)
    for h in holdings:
        h.pop("_raw_pct", None)

    return date_str, holdings


def _read_edel_index(wb) -> dict[str, str]:
    """Read Edelweiss Index sheet → {sheet_code: scheme_name}. col0=Fund Id, col1=Fund Desc."""
    if "Index" not in wb.sheetnames:
        return {}
    index = {}
    for row in wb["Index"].iter_rows(values_only=True):
        cells = [clean_str(c) for c in row]
        if len(cells) < 2 or not cells[0] or not cells[1]:
            continue
        if cells[0].lower() in ("fund id", "index", ""):
            continue
        if re.match(r"^(edelweiss|portfolio|mutual fund)\b", cells[0], re.IGNORECASE):
            continue
        index[cells[0]] = cells[1]
    return index


def _match_scheme(name_lower: str, mapping: dict) -> str | None:
    for kw, code in mapping.items():
        if kw in name_lower:
            return code
    return None


# ── Main entry point ──────────────────────────────────────────────────────────

def parse_file(filepath: Path, amc_key: str) -> dict | None:
    wb = openpyxl.load_workbook(str(filepath), data_only=True)
    name_lower = filepath.name.lower()

    if amc_key == "kotak":
        schemes = []
        global_date = None
        for sheet_code, scheme_code in KOTAK_SHEET_TO_SCHEME.items():
            if sheet_code not in wb.sheetnames:
                continue
            date_str, holdings = _parse_kotak_sheet(wb[sheet_code])
            if date_str and not global_date:
                global_date = date_str
            if holdings:
                schemes.append({"sheet_code": sheet_code, "scheme_code": scheme_code, "holdings": holdings})
        if not schemes:
            return None
        return {"amc": "kotak", "date": global_date, "source": str(filepath.name), "schemes": schemes}

    if amc_key == "hdfc":
        ws = next((wb[s] for s in wb.sheetnames if "derivative" not in s.lower()), wb.active)
        date_str, holdings = _parse_hdfc(ws)
        scheme_code = _match_scheme(name_lower, HDFC_SCHEME)
        return {"amc": "hdfc", "date": date_str, "source": str(filepath.name),
                "schemes": [{"sheet_code": wb.sheetnames[0], "scheme_code": scheme_code, "holdings": holdings}]}

    if amc_key == "canara_robeco":
        ws = wb.active
        date_str, holdings = _parse_std(ws)
        scheme_code = _match_scheme(name_lower, CANARA_SCHEME)
        return {"amc": "canara_robeco", "date": date_str, "source": str(filepath.name),
                "schemes": [{"sheet_code": wb.sheetnames[0], "scheme_code": scheme_code, "holdings": holdings}]}

    if amc_key == "pgim":
        ws = wb.active
        date_str, holdings = _parse_std(ws)
        scheme_code = _match_scheme(name_lower, PGIM_SCHEME) or "133839"
        return {"amc": "pgim", "date": date_str, "source": str(filepath.name),
                "schemes": [{"sheet_code": wb.sheetnames[0], "scheme_code": scheme_code, "holdings": holdings}]}

    if amc_key == "ppfas_manual":
        # Single-sheet PPFAS file — find matching scheme from sheet name
        from scrapers.ppfas import SHEET_TO_SCHEME, _parse_sheet
        schemes = []
        global_date = None
        for sheet_name in wb.sheetnames:
            result = _parse_sheet(wb[sheet_name])
            if result["date"] and not global_date:
                global_date = result["date"]
            if result["holdings"]:
                schemes.append({
                    "sheet_code":  sheet_name,
                    "scheme_code": SHEET_TO_SCHEME.get(sheet_name),
                    "holdings":    result["holdings"],
                })
        return {"amc": "ppfas", "date": global_date, "source": str(filepath.name), "schemes": schemes}

    if amc_key == "edelweiss":
        index = _read_edel_index(wb)
        schemes = []
        global_date = None
        for sheet_name in wb.sheetnames:
            if sheet_name.lower() in ("index", "notes", "disclaimer"):
                continue
            date_str, holdings = _parse_edel(wb[sheet_name])
            if not holdings:
                continue
            if date_str and not global_date:
                global_date = date_str
            scheme_name = index.get(sheet_name, "")
            schemes.append({
                "sheet_code":  sheet_name,
                "scheme_code": None,
                "scheme_name": scheme_name,
                "holdings":    holdings,
            })
        if not schemes:
            return None
        return {"amc": "edelweiss", "date": global_date, "source": str(filepath.name), "schemes": schemes}

    return None
