"""
Auto-push watcher for Pipeline repo.
Run once: python auto_push.py
Any change to .py or .xlsx files in this folder will be automatically
committed and pushed to GitHub, triggering a Streamlit Cloud redeploy.
"""

import time
import subprocess
import sys
from pathlib import Path
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler

REPO_DIR   = Path(__file__).parent
WATCH_EXTS = {".py", ".xlsx", ".txt"}
DEBOUNCE   = 3   # seconds to wait after last change before pushing

class AutoPushHandler(FileSystemEventHandler):
    def __init__(self):
        self._pending = False
        self._last_event = 0

    def on_modified(self, event):
        if event.is_directory:
            return
        p = Path(event.src_path)
        if p.suffix in WATCH_EXTS and ".git" not in str(p):
            self._pending = True
            self._last_event = time.time()

    on_created = on_modified

    def flush_if_ready(self):
        if self._pending and (time.time() - self._last_event) >= DEBOUNCE:
            self._pending = False
            self._push()

    def _push(self):
        try:
            result = subprocess.run(
                ["git", "status", "--porcelain"],
                cwd=REPO_DIR, capture_output=True, text=True
            )
            if not result.stdout.strip():
                return   # nothing changed

            subprocess.run(["git", "add", "-A"], cwd=REPO_DIR, check=True)

            msg = f"Auto update {time.strftime('%Y-%m-%d %H:%M:%S')}"
            subprocess.run(["git", "commit", "-m", msg], cwd=REPO_DIR, check=True)

            push = subprocess.run(
                ["git", "push", "origin", "main"],
                cwd=REPO_DIR, capture_output=True, text=True
            )
            if push.returncode == 0:
                print(f"[{time.strftime('%H:%M:%S')}] Pushed to GitHub ✓")
            else:
                print(f"[{time.strftime('%H:%M:%S')}] Push failed: {push.stderr.strip()}")

        except subprocess.CalledProcessError as e:
            print(f"[{time.strftime('%H:%M:%S')}] Git error: {e}")


def main():
    print(f"Watching: {REPO_DIR}")
    print("Any saved change will auto-push to GitHub → Streamlit redeploys.")
    print("Press Ctrl+C to stop.\n")

    handler  = AutoPushHandler()
    observer = Observer()
    observer.schedule(handler, str(REPO_DIR), recursive=False)
    observer.start()

    try:
        while True:
            handler.flush_if_ready()
            time.sleep(1)
    except KeyboardInterrupt:
        observer.stop()
        print("\nWatcher stopped.")
    observer.join()


if __name__ == "__main__":
    main()
