from __future__ import annotations

from typing import Callable

from pathlib import Path

from models.generation_task import GenerationTask
from services.layout_service import LayoutService
from services.word_exporter import WordExporter

ProgressCallback = Callable[[str, int], None]


class DocumentService:
    def __init__(
        self,
        layout_service: LayoutService | None = None,
        word_exporter: WordExporter | None = None,
    ) -> None:
        self.layout_service = layout_service or LayoutService()
        self.word_exporter = word_exporter or WordExporter()

    def generate(
        self,
        task: GenerationTask,
        prepared_images: list[Path],
        progress: ProgressCallback | None = None,
    ) -> Path:
        if progress:
            progress("正在整理文档布局", 65)

        payload = self.layout_service.build_word_payload(task, prepared_images)

        if progress:
            progress("正在导出 Word 文档", 82)

        return self.word_exporter.export(payload)
