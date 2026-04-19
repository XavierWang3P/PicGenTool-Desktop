from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path


@dataclass(slots=True)
class WordExportPayload:
    template_path: Path
    text_context: dict[str, str]
    image_paths: list[Path]
    output_path: Path
