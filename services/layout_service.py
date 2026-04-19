from __future__ import annotations

from pathlib import Path

from models.document_layout import WordExportPayload
from models.generation_task import GenerationTask
from services.template_service import TemplateService


class LayoutService:
    def __init__(self, template_service: TemplateService | None = None) -> None:
        self.template_service = template_service or TemplateService()

    def build_word_payload(
        self,
        task: GenerationTask,
        prepared_images: list[Path],
    ) -> WordExportPayload:
        template_path = self.template_service.get_template_path(len(task.image_paths))
        text_context = {
            "title": task.title,
            "location": task.location,
            "dateForDoc": (
                f"{task.activity_date.year}年"
                f"{task.activity_date.month}月"
                f"{task.activity_date.day}日"
            ),
        }

        return WordExportPayload(
            template_path=template_path,
            text_context=text_context,
            image_paths=prepared_images,
            output_path=task.save_path,
        )
