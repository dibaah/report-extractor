from __future__ import annotations

import os
import shutil
import subprocess
import sys
import tempfile
import threading
import time
from pathlib import Path

import uvicorn


def find_edge() -> str | None:
    candidates = [
        shutil.which("msedge"),
        os.path.join(os.environ.get("PROGRAMFILES(X86)", ""), "Microsoft", "Edge", "Application", "msedge.exe"),
        os.path.join(os.environ.get("PROGRAMFILES", ""), "Microsoft", "Edge", "Application", "msedge.exe"),
        os.path.join(os.environ.get("LOCALAPPDATA", ""), "Microsoft", "Edge", "Application", "msedge.exe"),
    ]
    return next((p for p in candidates if p and os.path.isfile(p)), None)


def main():
    config = uvicorn.Config("app:app", host="127.0.0.1", port=8765, log_level="warning")
    server = uvicorn.Server(config)
    thread = threading.Thread(target=server.run, name="ctep-api", daemon=True)
    thread.start()

    for _ in range(100):
        if server.started:
            break
        if not thread.is_alive():
            raise RuntimeError("Dashboard server failed to start.")
        time.sleep(0.1)

    edge = find_edge()
    if not edge:
        server.should_exit = True
        thread.join(timeout=5)
        raise RuntimeError("Microsoft Edge was not found. Please install Edge to launch the dashboard window.")

    profile = Path(tempfile.mkdtemp(prefix="CTEP-Shift-Report-Archive-Edge-"))
    try:
        browser = subprocess.Popen([
            edge, "--app=http://127.0.0.1:8765",
            f"--user-data-dir={profile}", "--no-first-run", "--no-default-browser-check"
        ])
        browser.wait()
    finally:
        # Closing the app window ends this launcher and its local API server.
        server.should_exit = True
        thread.join(timeout=10)
        if thread.is_alive():
            server.force_exit = True
            thread.join(timeout=3)
        shutil.rmtree(profile, ignore_errors=True)


if __name__ == "__main__":
    main()
