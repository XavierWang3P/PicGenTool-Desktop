from __future__ import annotations

import os
import subprocess
import sys
from datetime import date
from pathlib import Path


def project_root() -> Path:
    return Path(__file__).resolve().parent.parent


def ensure_directory(path: Path) -> Path:
    path.mkdir(parents=True, exist_ok=True)
    return path


def sanitize_filename(value: str) -> str:
    cleaned = "".join(ch for ch in value if ch not in '<>:"/\\|?*').strip()
    return cleaned or "未命名活动"


def build_default_output_path(first_image: Path | None, title: str, activity_date: date) -> Path:
    date_str = f"{str(activity_date.year)[2:]}.{activity_date.month:02}.{activity_date.day:02}"
    file_name = f"{date_str}-{sanitize_filename(title)}照片.docx"
    if first_image:
        return first_image.resolve().parent / file_name
    return project_root() / file_name


def open_with_default_app(path: Path) -> None:
    if sys.platform == "darwin":
        subprocess.run(["open", str(path)], check=False)
        return
    if sys.platform == "win32":
        os.startfile(str(path))  # type: ignore[attr-defined]
        return
    subprocess.run(["xdg-open", str(path)], check=False)

