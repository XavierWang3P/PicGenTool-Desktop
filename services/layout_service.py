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
        from docx.shared import Mm
        from docxtpl import DocxTemplate, InlineImage

        template_path = self.template_service.get_template_path(len(task.image_paths))
        doc = DocxTemplate(str(template_path))

        context = {
            "title": task.title,
            "location": task.location,
            "dateForDoc": (
                f"{task.activity_date.year}年"
                f"{task.activity_date.month}月"
                f"{task.activity_date.day}日"
            ),
        }
        for index, path in enumerate(prepared_images, start=1):
            context[f"image{index}"] = InlineImage(doc, str(path), width=Mm(72))

        return WordExportPayload(
            template_path=template_path,
            context=context,
            output_path=task.save_path,
        )
