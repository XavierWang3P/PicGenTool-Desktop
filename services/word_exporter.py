from __future__ import annotations

from pathlib import Path

from models.document_layout import WordExportPayload


class WordExporter:
    def export(self, payload: WordExportPayload) -> Path:
        from docx.shared import Mm
        from docxtpl import DocxTemplate, InlineImage

        try:
            doc = DocxTemplate(str(payload.template_path))
        except Exception as exc:
            raise IOError(f"模板文件打开失败：{exc}") from exc

        try:
            context = dict(payload.text_context)
            for index, path in enumerate(payload.image_paths, start=1):
                context[f"image{index}"] = InlineImage(doc, str(path), width=Mm(72))

            doc.render(context)
            doc.save(str(payload.output_path))
        except PermissionError as exc:
            raise IOError("无法保存文档，请检查文件权限或文件是否被占用") from exc
        except Exception as exc:
            raise IOError(f"文档生成失败：{exc}") from exc

        return payload.output_path
