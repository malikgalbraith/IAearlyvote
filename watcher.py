#!/usr/bin/env python3
"""
Iowa Absentee Ballot Watcher
Monitors the script folder for new/modified Absentee County*.pdf files
and automatically re-runs the breakdown analysis.

Usage:  python3 watcher.py   (leave running in a terminal tab)
"""

import subprocess
import sys
import time
from datetime import datetime
from pathlib import Path

try:
    from watchdog.events import FileSystemEventHandler
    from watchdog.observers import Observer
except ImportError:
    subprocess.check_call([sys.executable, "-m", "pip", "install", "watchdog"])
    from watchdog.events import FileSystemEventHandler
    from watchdog.observers import Observer

try:
    import plotly  # noqa: F401
except ImportError:
    subprocess.check_call([sys.executable, "-m", "pip", "install", "plotly"])

WATCH_DIR = Path(__file__).parent / "data"
DEBOUNCE_SECONDS = 5

_last_run: dict[str, float] = {}


def _run_breakdown(pdf_path: Path) -> None:
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    print(f"\n[{ts}] Detected: {pdf_path.name}")
    print("Running breakdown...")
    try:
        import breakdown
        breakdown.main(county_pdf_override=pdf_path)
    except Exception as exc:
        print(f"  ERROR: {exc}")


class AbsenteeHandler(FileSystemEventHandler):
    def _should_process(self, src_path: str) -> bool:
        p = Path(src_path)
        if not p.name.startswith("Absentee County"):
            return False
        if p.suffix.lower() != ".pdf":
            return False
        now = time.monotonic()
        if now - _last_run.get(src_path, 0) < DEBOUNCE_SECONDS:
            return False
        _last_run[src_path] = now
        return True

    def on_created(self, event):
        if not event.is_directory and self._should_process(event.src_path):
            _run_breakdown(Path(event.src_path))

    def on_modified(self, event):
        if not event.is_directory and self._should_process(event.src_path):
            _run_breakdown(Path(event.src_path))


if __name__ == "__main__":
    print(f"Watching {WATCH_DIR} for Absentee County*.pdf …")
    print("Press Ctrl+C to stop.\n")

    observer = Observer()
    observer.schedule(AbsenteeHandler(), str(WATCH_DIR), recursive=False)
    observer.start()
    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        observer.stop()
    observer.join()
