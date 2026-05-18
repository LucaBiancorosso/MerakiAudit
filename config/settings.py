from __future__ import annotations

import os
from pathlib import Path
from dotenv import load_dotenv

load_dotenv()


BASE_DIR = Path(__file__).resolve().parent.parent
OUTPUT_DIR = BASE_DIR / "output"
OUTPUT_DIR.mkdir(exist_ok=True)

MERAKI_DASHBOARD_API_KEY = os.getenv("MERAKI_DASHBOARD_API_KEY", "")
MERAKI_LOG_PATH = os.getenv("MERAKI_LOG_PATH", "")
MERAKI_PRINT_CONSOLE = os.getenv("MERAKI_PRINT_CONSOLE", "false").lower() == "true"