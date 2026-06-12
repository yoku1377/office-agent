"""Shared filesystem paths for the office-agent service."""
from pathlib import Path


ROOT_DIR = Path(__file__).resolve().parents[1]
ASSETS_DIR = ROOT_DIR / "assets"
STORAGE_DIR = ROOT_DIR / "storage"
UPLOAD_DIR = STORAGE_DIR / "uploads"
OUTPUT_DIR = STORAGE_DIR / "outputs"
TASK_DIR = STORAGE_DIR / "tasks"


def ensure_storage_dirs() -> None:
    for path in (UPLOAD_DIR, OUTPUT_DIR, TASK_DIR):
        path.mkdir(parents=True, exist_ok=True)
