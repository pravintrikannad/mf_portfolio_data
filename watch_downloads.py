"""
Watches ~/Downloads for new MF portfolio XLSX files.
On detection: auto-parses, updates data/<amc>.json, commits and pushes to GitHub.

Run once manually or set up as a LaunchAgent (see README or setup_launchagent.sh).
"""
import re
import time
import subprocess
import logging
from pathlib import Path
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler

DOWNLOADS = Path.home() / "Downloads"
REPO_DIR  = Path(__file__).parent.resolve()

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(message)s",
    datefmt="%H:%M:%S",
)
log = logging.getLogger(__name__)


# ── AMC detection rules ───────────────────────────────────────────────────────
# Each entry: (regex_pattern_on_filename, amc_key)
# Checked in order — first match wins.
AMC_PATTERNS = [
    (r"pgim.*flexi",          "pgim"),
    (r"pgim",                 "pgim"),
    (r"canara.*robeco",       "canara_robeco"),
    (r"canara",               "canara_robeco"),
    (r"consolidated.*sebi",   "kotak"),
    (r"kotak",                "kotak"),
    (r"hdfc.*flexi",          "hdfc"),
    (r"hdfc",                 "hdfc"),
    (r"ppfas.*portfolio",     "ppfas_manual"),
    (r"ppfas",                "ppfas_manual"),
    (r"mirae",                "mirae"),
    (r"axis.*flexi",          "axis"),
    (r"axis",                 "axis"),
    (r"sbi.*flexi",           "sbi"),
    (r"nippon",               "nippon"),
    (r"dsp",                  "dsp"),
    (r"franklin",             "franklin"),
]


def detect_amc(filename: str) -> str | None:
    name = filename.lower()
    for pattern, amc_key in AMC_PATTERNS:
        if re.search(pattern, name):
            return amc_key
    return None


def parse_and_commit(filepath: Path, amc_key: str):
    import sys
    sys.path.insert(0, str(REPO_DIR))
    from parse_manual import parse_file

    log.info(f"Parsing {filepath.name} as [{amc_key}]...")
    try:
        result = parse_file(filepath, amc_key)
        if not result:
            log.warning(f"Parser returned no data for {filepath.name}")
            return
    except Exception as e:
        log.error(f"Parse error: {e}")
        return

    # Drop schemes with no scheme_code (fund not in our universe)
    result["schemes"] = [s for s in result.get("schemes", []) if s.get("scheme_code")]
    if not result["schemes"]:
        log.info(f"Skipped {filepath.name} — no schemes in universe (scheme_code not mapped).")
        return

    import json
    out_path = REPO_DIR / "data" / f"{amc_key}.json"
    out_path.write_text(json.dumps(result, indent=2, ensure_ascii=False))
    log.info(f"Saved {out_path.name}: {len(result.get('schemes', []))} schemes")

    # Git commit + push
    try:
        subprocess.run(["git", "-C", str(REPO_DIR), "add", f"data/{amc_key}.json"], check=True)
        msg = f"chore: {amc_key} holdings {result.get('date', 'unknown')} (auto-import)"
        subprocess.run(["git", "-C", str(REPO_DIR), "commit", "-m", msg], check=True)
        subprocess.run(["git", "-C", str(REPO_DIR), "push"], check=True)
        log.info(f"Pushed {amc_key}.json to GitHub.")
    except subprocess.CalledProcessError as e:
        log.error(f"Git error: {e}")


class DownloadHandler(FileSystemEventHandler):
    def __init__(self):
        self._seen = set()

    def on_created(self, event):
        if event.is_directory:
            return
        path = Path(event.src_path)
        if path.suffix.lower() not in (".xlsx", ".xls"):
            return
        if path in self._seen:
            return
        self._seen.add(path)

        # Wait for file to finish writing (Safari/Chrome write in chunks)
        self._wait_stable(path)

        amc_key = detect_amc(path.name)
        if not amc_key:
            log.info(f"Ignored (no AMC match): {path.name}")
            return

        log.info(f"Detected: {path.name}  →  [{amc_key}]")
        parse_and_commit(path, amc_key)

    def _wait_stable(self, path: Path, timeout: int = 30):
        """Wait until file size stops changing (download complete)."""
        prev = -1
        deadline = time.time() + timeout
        while time.time() < deadline:
            try:
                size = path.stat().st_size
            except FileNotFoundError:
                time.sleep(0.5)
                continue
            if size == prev and size > 0:
                return
            prev = size
            time.sleep(1)


def main():
    log.info(f"Watching {DOWNLOADS} for MF portfolio files...")
    log.info("Download any AMC portfolio XLSX — it will auto-parse and push.")
    log.info("Press Ctrl+C to stop.\n")

    handler  = DownloadHandler()
    observer = Observer()
    observer.schedule(handler, str(DOWNLOADS), recursive=False)
    observer.start()
    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        observer.stop()
    observer.join()


if __name__ == "__main__":
    main()
