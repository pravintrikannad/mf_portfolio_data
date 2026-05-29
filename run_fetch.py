"""
Orchestrator for AMC holdings fetchers.
Run by GitHub Actions; outputs JSON files to data/<amc>.json.
"""
import json
import sys
import traceback
from pathlib import Path

DATA_DIR = Path(__file__).parent / "data"
DATA_DIR.mkdir(exist_ok=True)

# AMC registry — add entries here for Phase 2
AMC_REGISTRY = {
    "ppfas": {
        "module": "scrapers.ppfas",
        "fetch_fn": "fetch",
        "description": "Parag Parikh AMC",
    },
    # Phase 2 additions (uncomment + implement scraper):
    # "hdfc":    {"module": "scrapers.hdfc",    "fetch_fn": "fetch", "description": "HDFC AMC"},
    # "axis":    {"module": "scrapers.axis",    "fetch_fn": "fetch", "description": "Axis AMC"},
    # "nippon":  {"module": "scrapers.nippon",  "fetch_fn": "fetch", "description": "Nippon AMC"},
    # "uti":     {"module": "scrapers.uti",     "fetch_fn": "fetch", "description": "UTI AMC"},
}


def run_amc(amc_key: str, config: dict) -> bool:
    mod_name  = config["module"]
    fn_name   = config["fetch_fn"]
    desc      = config["description"]

    print(f"\n{'='*50}")
    print(f"Fetching: {desc} ({amc_key})")
    print(f"{'='*50}")

    try:
        import importlib
        mod = importlib.import_module(mod_name)
        fn  = getattr(mod, fn_name)
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
