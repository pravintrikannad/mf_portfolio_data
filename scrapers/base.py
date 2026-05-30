"""Shared utilities for AMC scrapers."""
import re
import datetime
from typing import Optional


def parse_date_from_text(text: str) -> Optional[str]:
    """Extract date from strings like 'Monthly Portfolio Statement as on April 30, 2026'."""
    # Try "as on April 30, 2026" / "as on April 30,2026"
    m = re.search(r'as on\s+(\w+ \d{1,2},?\s*\d{4})', text, re.IGNORECASE)
    if m:
        raw = m.group(1).strip()
        for fmt in ("%B %d, %Y", "%B %d,%Y", "%B %d %Y", "%b %d, %Y", "%b %d,%Y", "%b %d %Y"):
            try:
                dt = datetime.datetime.strptime(raw, fmt)
                return dt.strftime("%Y-%m-%d")
            except ValueError:
                continue
    # Try "as on 30-Apr-2026" (Kotak style)
    m = re.search(r'as on\s+(\d{1,2}-\w{3}-\d{4})', text, re.IGNORECASE)
    if m:
        try:
            dt = datetime.datetime.strptime(m.group(1).strip(), "%d-%b-%Y")
            return dt.strftime("%Y-%m-%d")
        except ValueError:
            pass
    # fallback: YYYY-MM-DD anywhere in text
    m2 = re.search(r'(\d{4}-\d{2}-\d{2})', text)
    if m2:
        return m2.group(1)
    return None


def clean_str(val) -> str:
    if val is None:
        return ""
    return str(val).strip()


def safe_float(val) -> Optional[float]:
    if val is None or clean_str(val) == "":
        return None
    try:
        return float(str(val).replace(",", "").strip())
    except (ValueError, TypeError):
        return None
