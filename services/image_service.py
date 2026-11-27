from __future__ import annotations

import logging
import os
import hashlib
import threading
from dataclasses import dataclass
from pathlib import Path
from tempfile import TemporaryDirectory
from typing import TYPE_CHECKING, Callable

from config import (
    DPI,
    IMAGE_OPTIMIZE,
    IMAGE_QUALITY,
    MAX_IMAGE_WORKERS,
    SUPPORTED_FORMATS,
    TARGET_HEIGHT_CM,
    TARGET_WIDTH_CM,
)
from models.image_selection import DeduplicationFailure, DeduplicationResult
if TYPE_CHECKING:
    from services.cancellation import CancellationToken


ProgressCallback = Callable[[str, int], None]
LOGGER = logging.getLogger("pgt.image")
_HEIF_EXTENSIONS = {".heic", ".heif"}
_HEIF_LOCK = threading.Lock()
_HEIF_REGISTERED = False
_HASH_CHUNK_SIZE = 1024 * 1024


@dataclass(slots=True)
class PreparedImages:
    directory: TemporaryDirectory[str]
    image_paths: list[Path]

    def cleanup(self) -> None:
        self.directory.cleanup()


class ImageService:
    def is_supported_image(self, path: str | Path) -> bool:
        return Path(path).suffix.lower() in SUPPORTED_FORMATS

    def filter_supported(self, files: list[str]) -> list[Path]:
        return [Path(file) for file in files if self.is_supported_image(file)]

    def deduplicate(
        self,
        files: list[str | Path],
        existing: list[Path] | None = None,
    ) -> DeduplicationResult:
        selected: list[Path] = []
        failures: list[DeduplicationFailure] = []
        duplicate_count = 0
        seen_paths: set[str] = set()
        seen_digests: dict[int, set[str]] = {}

        for path in existing or []:
            seen_paths.add(self._path_key(path))
            try:
                size = path.stat().st_size
                digest = self._content_digest(path)
            except OSError:
                continue
            seen_digests.setdefault(size, set()).add(digest)

        for value in files:
            path = Path(value)
            if not self.is_supported_image(path):
                continue

            path_key = self._path_key(path)
            if path_key in seen_paths:
                duplicate_count += 1
                continue

            try:
                size = path.stat().st_size
                digest = self._content_digest(path)
            except OSError as exc:
                failures.append(
                    DeduplicationFailure(
                        filename=path.name,
                        reason=f"无法读取文件（{type(exc).__name__}）",
                    )
                )
                continue

            digests = seen_digests.setdefault(size, set())
            if digest in digests:
                duplicate_count += 1
                seen_paths.add(path_key)
                continue

            selected.append(path)
            seen_paths.add(path_key)
            digests.add(digest)

        return DeduplicationResult(
            paths=selected,
            duplicate_count=duplicate_count,
            failures=tuple(failures),
        )

    def _content_digest(self, path: Path) -> str:
        digest = hashlib.sha256()
        with path.open("rb") as source:
            while chunk := source.read(_HASH_CHUNK_SIZE):
                digest.update(chunk)
        return digest.hexdigest()

    def _path_key(self, path: Path) -> str:
        return os.path.normcase(str(path.resolve(strict=False)))

    def _ensure_heif_support(self) -> None:
        global _HEIF_REGISTERED

        if _HEIF_REGISTERED:
            return
        with _HEIF_LOCK:
            if _HEIF_REGISTERED:
                return
            from pi_heif import register_heif_opener

            register_heif_opener()
            _HEIF_REGISTERED = True

    def prepare_images(
        self,
        image_paths: list[Path],
        progress: ProgressCallback | None = None,
        cancellation: CancellationToken | None = None,
    ) -> PreparedImages:
        from concurrent.futures import ThreadPoolExecutor, as_completed

        from services.cancellation import CancellationToken, GenerationCancelled
        from utils.performance import measure

        if not image_paths:
            raise ValueError("请选择至少一张图片")
        cancellation = cancellation or CancellationToken()
        cancellation.raise_if_cancelled()

        temp_dir = TemporaryDirectory(prefix="pgt-images-")
        output_dir = Path(temp_dir.name)
        pixels_per_cm = DPI / 2.54
        target_width = int(TARGET_WIDTH_CM * pixels_per_cm)
        target_height = int(TARGET_HEIGHT_CM * pixels_per_cm)
        target_ratio = target_width / target_height
        total = len(image_paths)
        prepared_paths: list[Path | None] = [None] * total
        executor = ThreadPoolExecutor(
            max_workers=self._worker_count(total),
            thread_name_prefix="pgt-image",
        )
        futures: dict[object, tuple[int, Path]] = {}

        if progress:
            progress(f"正在处理图片 0/{total}", 0)

        try:
            with measure(LOGGER, "image_total", image_count=total):
                for index, image_path in enumerate(image_paths):
                    output_path = output_dir / f"image_{index + 1:02d}.jpg"
                    future = executor.submit(
                        self._prepare_one_timed,
                        image_path,
                        output_path,
                        (target_width, target_height),
                        target_ratio,
                        index + 1,
                        cancellation,
                    )
                    futures[future] = (index, image_path)

                completed = 0
                for future in as_completed(futures):
                    index, image_path = futures[future]
                    try:
                        prepared_paths[index] = future.result()
                        if cancellation:
                            cancellation.raise_if_cancelled()
                    except GenerationCancelled:
                        for pending in futures:
                            pending.cancel()
                        raise
                    except Exception as exc:
                        cancellation.cancel()
                        for pending in futures:
                            pending.cancel()
                        raise ValueError(
                            f"处理图片失败：{image_path.name}，"
                            f"原因：{self._safe_error_reason(exc)}"
                        ) from None

                    completed += 1
                    if progress:
                        progress(
                            f"已完成图片 {completed}/{total}",
                            self._progress(completed, total),
                        )
        except Exception:
            executor.shutdown(wait=True, cancel_futures=True)
            temp_dir.cleanup()
            raise
        else:
            executor.shutdown(wait=True)

        return PreparedImages(
            directory=temp_dir,
            image_paths=[path for path in prepared_paths if path is not None],
        )

    def _worker_count(self, image_count: int) -> int:
        return min(MAX_IMAGE_WORKERS, os.cpu_count() or 1, max(1, image_count))

    def _safe_error_reason(self, error: Exception) -> str:
        error_name = type(error).__name__
        if error_name == "UnidentifiedImageError":
            return "无法识别图片格式"
        if isinstance(error, OSError):
            return "图片解码或编码失败"
        return f"图片处理异常（{error_name}）"

    def _prepare_one_timed(
        self,
        image_path: Path,
        output_path: Path,
        target_size: tuple[int, int],
        target_ratio: float,
        image_index: int,
        cancellation: CancellationToken | None,
    ) -> Path:
        from utils.performance import measure

        if cancellation:
            cancellation.raise_if_cancelled()
        with measure(LOGGER, "image_item", image_index=image_index):
            result = self._prepare_one(image_path, output_path, target_size, target_ratio)
        if cancellation:
            cancellation.raise_if_cancelled()
        return result

    def _prepare_one(
        self,
        image_path: Path,
        output_path: Path,
        target_size: tuple[int, int],
        target_ratio: float,
    ) -> Path:
        from PIL import Image, ImageOps

        if image_path.suffix.lower() in _HEIF_EXTENSIONS:
            self._ensure_heif_support()

        intermediates: list[Image.Image] = []
        with Image.open(image_path) as original:
            if original.format == "JPEG":
                original.draft("RGB", target_size)

            try:
                oriented = ImageOps.exif_transpose(original)
                intermediates.append(oriented)

                normalized = self._normalize_mode(oriented)
                if normalized is not oriented:
                    intermediates.append(normalized)

                cropped = self._crop_to_ratio(normalized, target_ratio)
                if cropped is not normalized:
                    intermediates.append(cropped)

                resized = cropped.resize(target_size, Image.Resampling.LANCZOS)
                intermediates.append(resized)
                resized.save(
                    output_path,
                    format="JPEG",
                    quality=IMAGE_QUALITY,
                    optimize=IMAGE_OPTIMIZE,
                    dpi=(DPI, DPI),
                )
            finally:
                closed: set[int] = set()
                for image in reversed(intermediates):
                    if id(image) not in closed:
                        image.close()
                        closed.add(id(image))

        return output_path

    def _normalize_mode(self, image):
        if image.mode in {"RGBA", "P", "LA"}:
            return image.convert("RGB")
        if image.mode != "RGB":
            return image.convert("RGB")
        return image

    def _crop_to_ratio(self, image, target_ratio: float):
        width, height = image.size
        if width <= 0 or height <= 0:
            raise ValueError("图片尺寸无效")

        current_ratio = width / height
        if current_ratio > target_ratio:
            new_width = int(height * target_ratio)
            left = (width - new_width) // 2
            return image.crop((left, 0, left + new_width, height))
        if current_ratio < target_ratio:
            new_height = int(width / target_ratio)
            top = (height - new_height) // 2
            return image.crop((0, top, width, top + new_height))
        return image

    def _progress(self, completed: int, total: int) -> int:
        if total <= 0:
            return 0
        return int(completed / total * 60)
