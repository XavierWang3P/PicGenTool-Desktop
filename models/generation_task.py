from __future__ import annotations

from dataclasses import dataclass
from datetime import date
from pathlib import Path


@dataclass(slots=True)
class GenerationTask:
    title: str
    location: str
    activity_date: date
    image_paths: list[Path]
    save_path: Path
    open_file: bool = True
