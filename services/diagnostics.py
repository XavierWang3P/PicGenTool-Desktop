from __future__ import annotations

import platform
import sys
from pathlib import Path

from config import APP_VERSION, MAX_IMAGES
from models.diagnostics import DiagnosticSnapshot
from services.explorer_integration import registration_status
from services.template_service import TemplateService
from utils.logging_utils import current_log_path


class DiagnosticsService:
    def __init__(self, template_service: TemplateService | None = None) -> None:
        self.template_service = template_service or TemplateService()

    def collect(self, deep_templates: bool = True) -> DiagnosticSnapshot:
        try:
            import pi_heif

            heif_available = True
            heif_version = pi_heif.__version__
        except (ImportError, AttributeError):
            heif_available = False
            heif_version = None

        return DiagnosticSnapshot(
            app_version=APP_VERSION,
            python_version=platform.python_version(),
            platform=f"{platform.system()} {platform.release()} {platform.machine()}".strip(),
            packaged=bool(getattr(sys, "frozen", False)),
            heif_available=heif_available,
            heif_version=heif_version,
            template_result=self.template_service.validate_all(deep=deep_templates),
            log_path=current_log_path(),
            explorer_status=registration_status(Path(sys.executable)),
        )

    @staticmethod
    def format_for_copy(snapshot: DiagnosticSnapshot) -> str:
        template = snapshot.template_result
        if template.valid:
            template_status = f"valid ({template.checked_count}/{MAX_IMAGES})"
        else:
            template_status = (
                f"invalid ({template.checked_count}/{MAX_IMAGES}, "
                f"{len(template.issues)} issues)"
            )
        heif_status = "unavailable"
        if snapshot.heif_available:
            heif_status = f"available ({snapshot.heif_version or 'unknown'})"
        log_status = "unavailable"
        if snapshot.log_path is not None:
            log_status = f"available ({snapshot.log_path.name})"

        return "\n".join(
            [
                f"PicGenTool: {snapshot.app_version}",
                f"Python: {snapshot.python_version}",
                f"Platform: {snapshot.platform}",
                f"Packaged: {'yes' if snapshot.packaged else 'no'}",
                f"HEIF: {heif_status}",
                f"Templates: {template_status}",
                *[
                    f"Template-{issue.template_number}: {issue.code}: {issue.message}"
                    for issue in template.issues
                ],
                f"Log: {log_status}",
                f"Explorer menu: {snapshot.explorer_status.value}",
            ]
        )
