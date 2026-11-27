from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Any


@dataclass(slots=True)
class PreparedWordExport:
    creator_thread_id: int
    document: Any
    text_context: dict[str, str]
    output_path: Path
