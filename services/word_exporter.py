from __future__ import annotations

import logging
from pathlib import Path
from tempfile import NamedTemporaryFile
import threading
from typing import TYPE_CHECKING

from models.document_layout import WordExportPayload
from models.word_export_session import PreparedWordExport
if TYPE_CHECKING:
    from services.cancellation import CancellationToken


class WordExporter:
    def prepare(
        self,
        payload: WordExportPayload,
        cancellation: CancellationToken | None = None,
    ) -> PreparedWordExport:
        from docxtpl import DocxTemplate
        from utils.performance import measure

        if cancellation:
            cancellation.raise_if_cancelled()
        logger = logging.getLogger("pgt.word")
        try:
            with measure(logger, "word_open", image_count=len(payload.image_paths)):
                document = DocxTemplate(str(payload.template_path))
        except Exception as exc:
            raise IOError("模板文件打开失败，请检查模板完整性") from exc
        if cancellation:
            cancellation.raise_if_cancelled()
        return PreparedWordExport(
            creator_thread_id=threading.get_ident(),
            document=document,
            text_context=dict(payload.text_context),
            output_path=payload.output_path,
        )

    def export(
        self,
        payload: WordExportPayload,
        cancellation: CancellationToken | None = None,
    ) -> Path:
        prepared = self.prepare(payload, cancellation=cancellation)
        return self.export_prepared(
            prepared,
            payload.image_paths,
            cancellation=cancellation,
        )

    def export_prepared(
        self,
        prepared: PreparedWordExport,
        image_paths: list[Path],
        cancellation: CancellationToken | None = None,
    ) -> Path:
        from docx.shared import Mm
        from docxtpl import InlineImage
        from services.cancellation import GenerationCancelled
        from utils.performance import measure

        if threading.get_ident() != prepared.creator_thread_id:
            raise RuntimeError("Word 导出必须在创建线程中完成")
        logger = logging.getLogger("pgt.word")
        if cancellation:
            cancellation.raise_if_cancelled()
        doc = prepared.document

        temporary_path: Path | None = None
        try:
            context = dict(prepared.text_context)
            for index, path in enumerate(image_paths, start=1):
                context[f"image{index}"] = InlineImage(doc, str(path), width=Mm(72))

            with measure(logger, "word_render", image_count=len(image_paths)):
                doc.render(context)
            if cancellation:
                cancellation.raise_if_cancelled()

            with NamedTemporaryFile(
                dir=prepared.output_path.parent,
                suffix=".docx",
                delete=False,
            ) as temporary_file:
                temporary_path = Path(temporary_file.name)

            with measure(logger, "word_save", image_count=len(image_paths)):
                doc.save(str(temporary_path))
            if cancellation:
                cancellation.raise_if_cancelled()

            if cancellation:
                cancellation.commit(lambda: temporary_path.replace(prepared.output_path))
            else:
                temporary_path.replace(prepared.output_path)
            temporary_path = None
        except GenerationCancelled:
            raise
        except PermissionError as exc:
            raise IOError("无法保存文档，请检查文件权限或文件是否被占用") from exc
        except Exception as exc:
            raise IOError("文档生成失败，请检查模板、图片和保存位置") from exc
        finally:
            if temporary_path is not None:
                try:
                    temporary_path.unlink(missing_ok=True)
                except OSError as exc:
                    logger.warning(
                        "无法清理 Word 临时文件: error_type=%s",
                        type(exc).__name__,
                    )

        return prepared.output_path
