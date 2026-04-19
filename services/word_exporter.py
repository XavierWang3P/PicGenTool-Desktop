from __future__ import annotations

from pathlib import Path

from models.document_layout import WordExportPayload


class WordExporter:
    def export(self, payload: WordExportPayload) -> Path:
        from docxtpl import DocxTemplate

        try:
            doc = DocxTemplate(str(payload.template_path))
        except Exception as exc:
            raise IOError(f"模板文件打开失败：{exc}") from exc

        try:
            doc.render(payload.context)
            doc.save(str(payload.output_path))
        except PermissionError as exc:
            raise IOError("无法保存文档，请检查文件权限或文件是否被占用") from exc
        except Exception as exc:
            raise IOError(f"文档生成失败：{exc}") from exc

        return payload.output_path
