from __future__ import annotations

import logging
from concurrent.futures import Future, ThreadPoolExecutor
from pathlib import Path
from typing import TYPE_CHECKING, Callable

from models.generation_task import GenerationTask
from services.document_service import DocumentService
from services.image_service import ImageService
from services.output_claims import claim_output

if TYPE_CHECKING:
    from services.cancellation import CancellationToken

ProgressCallback = Callable[[str, int], None]


class GenerationService:
    """Run the existing image and Word pipeline without any UI dependency."""

    def __init__(
        self,
        image_service: ImageService | None = None,
        document_service: DocumentService | None = None,
    ) -> None:
        self.image_service = image_service or ImageService()
        self.document_service = document_service or DocumentService()
        self.logger = logging.getLogger("pgt.generation")

    def run(
        self,
        task: GenerationTask,
        progress: ProgressCallback | None = None,
        cancellation: CancellationToken | None = None,
        *,
        output_policy: str = "overwrite",
    ) -> Path | None:
        """Return the document path, or None when an existing output is skipped."""
        from services.cancellation import CancellationToken, GenerationCancelled

        cancellation = cancellation or CancellationToken()
        if output_policy not in {"fail", "skip", "overwrite"}:
            raise ValueError("无效的输出策略")
        cancellation.raise_if_cancelled()
        prepared = None
        image_future: Future | None = None
        executor = ThreadPoolExecutor(max_workers=1, thread_name_prefix="pgt-image")
        with claim_output(task.save_path, output_policy) as skipped:
            if skipped:
                return None
            try:
                image_future = executor.submit(
                    self.image_service.prepare_images,
                    task.image_paths,
                    progress,
                    cancellation=cancellation,
                )
                if progress:
                    progress("正在准备 Word 模板", 60)
                try:
                    prepared_word = self.document_service.prepare(
                        task,
                        cancellation=cancellation,
                    )
                except BaseException:
                    cancellation.cancel()
                    try:
                        prepared = image_future.result()
                    except BaseException as image_error:
                        if not isinstance(image_error, GenerationCancelled):
                            raise image_error
                    raise

                prepared = image_future.result()
                docx_path = self.document_service.export_prepared(
                    prepared_word,
                    prepared.image_paths,
                    progress,
                    cancellation=cancellation,
                )
                cancellation.raise_if_cancelled()
                if progress:
                    progress("生成完成", 100)
                return docx_path
            except BaseException:
                cancellation.cancel()
                raise
            finally:
                if image_future is not None:
                    try:
                        prepared = image_future.result()
                    except BaseException:
                        pass
                executor.shutdown(wait=True, cancel_futures=False)
                if prepared:
                    try:
                        prepared.cleanup()
                    except OSError as exc:
                        self.logger.warning(
                            "清理图片临时文件失败: error_type=%s",
                            type(exc).__name__,
                        )
