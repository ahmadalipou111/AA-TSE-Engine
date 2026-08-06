from pathlib import Path
from dotenv import load_dotenv
import os

# Root folder of project
BASE_DIR = Path(__file__).resolve().parent

# Load .env
load_dotenv(BASE_DIR / ".env")

# API KEY
BRS_API_KEY = os.getenv("BRS_API_KEY")

# Folders
LOG_DIR = BASE_DIR / "logs"
BACKUP_DIR = BASE_DIR / "backup"
OUTPUT_DIR = BASE_DIR / "output"

# Create folders if missing
for folder in [LOG_DIR, BACKUP_DIR, OUTPUT_DIR]:
    folder.mkdir(exist_ok=True)