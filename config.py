from __future__ import annotations

import os
import shutil
from pathlib import Path
from dotenv import load_dotenv

BASE_DIR = Path(__file__).resolve().parent
load_dotenv(BASE_DIR / ".env")
DATA_DIR_VALUE = os.getenv("DATA_DIR", "").strip()
DATA_DIR = Path(DATA_DIR_VALUE).expanduser() if DATA_DIR_VALUE else BASE_DIR / "data"
INCOMING_DIR = DATA_DIR / "incoming"
DATA_DIR.mkdir(parents=True, exist_ok=True)
INCOMING_DIR.mkdir(parents=True, exist_ok=True)

MAX_BOT_TOKEN = os.getenv("MAX_BOT_TOKEN", "").strip()
MAX_ADMIN_IDS = {
    int(part.strip())
    for part in os.getenv("MAX_ADMIN_IDS", "").split(",")
    if part.strip().isdigit()
}
MAX_SOURCE_CHAT_ID = os.getenv("MAX_SOURCE_CHAT_ID", "").strip()
MAX_POLL_TIMEOUT = max(1, min(int(os.getenv("MAX_POLL_TIMEOUT", "60")), 90))
MAX_CA_BUNDLE = os.getenv("MAX_CA_BUNDLE", "").strip()
MAX_WEBHOOK_URL = os.getenv("MAX_WEBHOOK_URL", "").strip()
WEBHOOK_SECRET = os.getenv("WEBHOOK_SECRET", "").strip()
WEBHOOK_PATH = os.getenv("WEBHOOK_PATH", "/webhook").strip() or "/webhook"
if not WEBHOOK_PATH.startswith("/"):
    WEBHOOK_PATH = "/" + WEBHOOK_PATH
PORT = max(1, min(int(os.getenv("PORT", "8080")), 65535))
MAX_STATE_PATH = DATA_DIR / "max_state.json"
MAX_USER_DB = DATA_DIR / "max_users.sqlite3"
SCHEDULE_XLSX = DATA_DIR / "schedule.xlsx"
BUNDLED_SCHEDULE_XLSX = BASE_DIR / "assets" / "schedule_seed.xlsx"
if not SCHEDULE_XLSX.exists() and BUNDLED_SCHEDULE_XLSX.exists():
    shutil.copy2(BUNDLED_SCHEDULE_XLSX, SCHEDULE_XLSX)
SCHEDULE_JSON = DATA_DIR / "schedule.json"
BELL_JSON = DATA_DIR / "bell_schedule.json"
MAX_UPLOAD_BYTES = max(1, int(os.getenv("MAX_UPLOAD_MB", "20"))) * 1024 * 1024
TIMEZONE = os.getenv("TIMEZONE", "Europe/Warsaw").strip() or "Europe/Warsaw"
NOTIFY_USERS_ON_UPDATE = os.getenv("NOTIFY_USERS_ON_UPDATE", "true").lower() in {"1", "true", "yes", "да"}
