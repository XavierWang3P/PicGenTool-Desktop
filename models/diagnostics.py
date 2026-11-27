from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

from models.template_validation import TemplateValidationResult
from services.explorer_integration import ExplorerRegistrationStatus


@dataclass(frozen=True, slots=True)
class DiagnosticSnapshot:
    app_version: str
    python_version: str
    platform: str
    packaged: bool
    heif_available: bool
    heif_version: str | None
    template_result: TemplateValidationResult
    log_path: Path | None
    explorer_status: ExplorerRegistrationStatus
