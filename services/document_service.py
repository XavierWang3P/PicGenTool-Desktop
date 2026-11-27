from __future__ import annotations

import logging
from typing import TYPE_CHECKING, Callable

from pathlib import Path

from models.generation_task import GenerationTask
from models.word_export_session import PreparedWordExport
from services.layout_service import LayoutService
from services.word_exporter import WordExporter
if TYPE_CHECKING:
    from services.cancellation import CancellationToken

ProgressCallback = Callable[[str, int], None]


class DocumentService:
    def __init__(
        self,
        layout_service: LayoutService | None = None,
        word_exporter: WordExporter | None = None,
    ) -> None:
        self.layout_service = layout_service or LayoutService()
        self.word_exporter = word_exporter or WordExporter()
        self.logger = logging.getLogger("pgt.document")

    def generate(
        self,
        task: GenerationTask,
        prepared_images: list[Path],
        progress: ProgressCallback | None = None,
        cancellation: CancellationToken | None = None,
    ) -> Path:
        from utils.performance import measure

        if progress:
            progress("正在整理文档布局", 65)
        prepared = self.prepare(task, cancellation=cancellation)
        return self.export_prepared(
            prepared,
            prepared_images,
            progress=progress,
            cancellation=cancellation,
        )

    def prepare(
        self,
        task: GenerationTask,
        cancellation: CancellationToken | None = None,
    ) -> PreparedWordExport:
        from utils.performance import measure

        if cancellation:
            cancellation.raise_if_cancelled()
        with measure(self.logger, "layout", image_count=len(task.image_paths)):
            payload = self.layout_service.build_word_payload(task, [])
        if cancellation:
            cancellation.raise_if_cancelled()
        return self.word_exporter.prepare(payload, cancellation=cancellation)

    def export_prepared(
        self,
        prepared: PreparedWordExport,
        image_paths: list[Path],
        progress: ProgressCallback | None = None,
        cancellation: CancellationToken | None = None,
    ) -> Path:
        if cancellation:
            cancellation.raise_if_cancelled()
        if progress:
            progress("正在导出 Word 文档", 82)
        return self.word_exporter.export_prepared(
            prepared,
            image_paths,
            cancellation=cancellation,
        )
