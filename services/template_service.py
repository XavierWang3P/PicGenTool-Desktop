from __future__ import annotations

from pathlib import Path

from config import MAX_IMAGES
from utils.path_utils import project_root


class TemplateService:
    def __init__(self, template_dir: Path | None = None) -> None:
        self.template_dir = template_dir or project_root() / "templates"

    def get_template_path(self, image_count: int) -> Path:
        if image_count < 1:
            raise ValueError("请选择至少一张图片")
        if image_count > MAX_IMAGES:
            raise ValueError(f"图片数量不能超过 {MAX_IMAGES} 张")

        template_path = self.template_dir / f"Template-{image_count}.docx"
        if not template_path.exists():
            raise FileNotFoundError(
                f"未找到对应 {image_count} 张图片的模板文件，请检查 templates 目录"
            )
        return template_path

