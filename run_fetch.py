"""
Orchestrator for AMC holdings fetchers.
Run by GitHub Actions; outputs JSON files to data/<amc>.json.

Two types of fetchers:
  1. Direct scrapers (ppfas, kotak, hdfc, etc.) — module/fetch_fn style
  2. Advisorkhoj-based (absl, sbi, nippon, etc.) — use advisorkhoj.fetch_amc
"""
import json
import sys
import traceback
from pathlib import Path

DATA_DIR = Path(__file__).parent / "data"
DATA_DIR.mkdir(exist_ok=True)

# AMC registry
AMC_REGISTRY = {
    # ── Direct scrapers ─────────────────────────────────────────────
    "ppfas": {
        "module": "scrapers.ppfas",
        "fetch_fn": "fetch",
        "description": "Parag Parikh AMC",
    },
    "canara_robeco": {
        "module": "scrapers.canara_robeco",
        "fetch_fn": "fetch",
        "description": "Canara Robeco AMC",
    },
    "pgim": {
        "module": "scrapers.pgim",
        "fetch_fn": "fetch",
        "description": "PGIM India AMC",
    },
    "kotak": {
        "module": "scrapers.kotak",
        "fetch_fn": "fetch",
        "description": "Kotak AMC",
    },
    "hdfc": {
        "module": "scrapers.hdfc",
        "fetch_fn": "fetch",
        "description": "HDFC AMC",
    },
    # ── Advisorkhoj-based scrapers ──────────────────────────────────
    "absl":      {"advisorkhoj": True, "description": "Aditya Birla Sun Life AMC"},
    "sbi":       {"advisorkhoj": True, "description": "SBI Mutual Fund"},
    "nippon":    {"advisorkhoj": True, "description": "Nippon India AMC"},
    "icici_pru": {"advisorkhoj": True, "description": "ICICI Prudential AMC"},
    "uti":       {"advisorkhoj": True, "description": "UTI AMC"},
    "dsp":       {"advisorkhoj": True, "description": "DSP AMC"},
    "franklin":  {"advisorkhoj": True, "description": "Franklin Templeton AMC"},
    "tata":      {"advisorkhoj": True, "description": "Tata AMC"},
    "motilal":   {"advisorkhoj": True, "description": "Motilal Oswal AMC"},
    "quant":     {"advisorkhoj": True, "description": "Quant AMC"},
    "bajaj":     {"advisorkhoj": True, "description": "Bajaj Finserv AMC"},
    "nj":        {"advisorkhoj": True, "description": "NJ AMC"},
}


def run_amc(amc_key: str, config: dict) -> bool:
    desc = config["description"]

    print(f"\n{'='*50}")
    print(f"Fetching: {desc} ({amc_key})")
    print(f"{'='*50}")

    try:
        if config.get("advisorkhoj"):
            from scrapers.advisorkhoj import fetch_amc
            data = fetch_amc(amc_key)
        else:
            import importlib
            mod = importlib.import_module(config["module"])
            fn  = getattr(mod, config["fetch_fn"])
            data = fn()

        if not data or not data.get("schemes"):
            print(f"  WARNING: No schemes returned for {amc_key}")
            return False

        total_holdings = sum(len(s["holdings"]) for s in data["schemes"])
        print(f"  OK: {len(data['schemes'])} schemes, {total_holdings} total holdings, date={data.get('date')}")

        out_path = DATA_DIR / f"{amc_key}.json"
        out_path.write_text(json.dumps(data, indent=2, ensure_ascii=False))
        print(f"  Saved: {out_path}")
        return True

    except Exception:
        print(f"  ERROR fetching {amc_key}:")
        traceback.print_exc()
        return False


def main():
    # Allow running a subset: python run_fetch.py ppfas hdfc
    targets = sys.argv[1:] if len(sys.argv) > 1 else list(AMC_REGISTRY.keys())

    results = {}
    for key in targets:
        if key not in AMC_REGISTRY:
            print(f"Unknown AMC key: {key}. Available: {list(AMC_REGISTRY.keys())}")
            results[key] = False
            continue
        results[key] = run_amc(key, AMC_REGISTRY[key])

    print("\n" + "="*50)
    print("SUMMARY")
    print("="*50)
    for key, ok in results.items():
        status = "OK" if ok else "FAILED"
        print(f"  {key:20s} {status}")

    failed = [k for k, ok in results.items() if not ok]
    if failed:
        print(f"\nFailed: {failed}")
        sys.exit(1)
    else:
        print("\nAll done.")


if __name__ == "__main__":
    main()
