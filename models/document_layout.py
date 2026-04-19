from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Any


@dataclass(slots=True)
class WordExportPayload:
    template_path: Path
    context: dict[str, Any]
    output_path: Path

