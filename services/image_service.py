from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from tempfile import TemporaryDirectory
from typing import Callable

from config import (
    DPI,
    IMAGE_QUALITY,
    SUPPORTED_FORMATS,
    TARGET_HEIGHT_CM,
    TARGET_WIDTH_CM,
)


ProgressCallback = Callable[[str, int], None]


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

    def prepare_images(
        self, image_paths: list[Path], progress: ProgressCallback | None = None
    ) -> PreparedImages:
        if not image_paths:
            raise ValueError("请选择至少一张图片")

        from PIL import Image

        temp_dir = TemporaryDirectory(prefix="pgt-images-")
        output_dir = Path(temp_dir.name)
        prepared_paths: list[Path] = []

        pixels_per_cm = DPI / 2.54
        target_width = int(TARGET_WIDTH_CM * pixels_per_cm)
        target_height = int(TARGET_HEIGHT_CM * pixels_per_cm)
        target_ratio = target_width / target_height
        total = len(image_paths)

        for index, image_path in enumerate(image_paths, start=1):
            if progress:
                progress(f"正在处理图片 {index}/{total}", self._progress(index - 1, total))

            try:
                with Image.open(image_path) as original:
                    image = original.copy()
            except Exception as exc:
                raise ValueError(f"无法打开图片文件：{image_path.name}，原因：{exc}") from exc

            try:
                image = self._normalize_mode(image)
                image = self._crop_to_ratio(image, target_ratio)
                image = image.resize((target_width, target_height), Image.Resampling.LANCZOS)

                output_path = output_dir / f"image_{index:02d}.jpg"
                image.save(
                    output_path,
                    format="JPEG",
                    quality=IMAGE_QUALITY,
                    optimize=True,
                    dpi=(DPI, DPI),
                )
                prepared_paths.append(output_path)
            except Exception as exc:
                temp_dir.cleanup()
                raise ValueError(f"处理图片失败：{image_path.name}，原因：{exc}") from exc
            finally:
                image.close()

            if progress:
                progress(f"已完成图片 {index}/{total}", self._progress(index, total))

        return PreparedImages(directory=temp_dir, image_paths=prepared_paths)

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

