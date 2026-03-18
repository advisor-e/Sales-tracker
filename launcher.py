from __future__ import annotations

import os
import subprocess
import sys
from pathlib import Path


APP_DIR = Path(__file__).resolve().parent
APP_FILE = APP_DIR / "app.py"


def main() -> int:
    env = os.environ.copy()
    env.setdefault("STREAMLIT_SERVER_HEADLESS", "false")
    env.setdefault("STREAMLIT_BROWSER_GATHER_USAGE_STATS", "false")

    command = [
        sys.executable,
        "-m",
        "streamlit",
        "run",
        str(APP_FILE),
        "--server.port",
        "8501",
        "--server.address",
        "127.0.0.1",
    ]

    try:
        return subprocess.call(command, cwd=str(APP_DIR), env=env)
    except KeyboardInterrupt:
        return 0


if __name__ == "__main__":
    raise SystemExit(main())
