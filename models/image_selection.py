from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path


@dataclass(frozen=True, slots=True)
class DeduplicationFailure:
    filename: str
    reason: str


@dataclass(frozen=True, slots=True)
class DeduplicationResult:
    paths: list[Path]
    duplicate_count: int
    failures: tuple[DeduplicationFailure, ...] = ()
