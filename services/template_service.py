from __future__ import annotations

from pathlib import Path
import re
from xml.etree import ElementTree
from zipfile import BadZipFile, ZipFile

from config import MAX_IMAGES
from models.template_validation import TemplateIssue, TemplateValidationResult
from utils.path_utils import project_root


_PLACEHOLDER_PATTERN = re.compile(r"\{\{\s*([^{}]+?)\s*\}\}")
_COMMON_PLACEHOLDERS = {"title", "location", "dateForDoc"}


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

    def validate_all(self, deep: bool = False) -> TemplateValidationResult:
        issues: list[TemplateIssue] = []
        checked_count = 0

        for image_count in range(1, MAX_IMAGES + 1):
            template_path = self.template_dir / f"Template-{image_count}.docx"
            if not template_path.is_file():
                issues.append(
                    TemplateIssue(image_count, "missing_file", "模板文件不存在")
                )
                continue

            checked_count += 1
            if deep:
                issues.extend(self._validate_template(template_path, image_count))

        issues.sort(key=lambda issue: (issue.template_number, issue.code, issue.message))
        return TemplateValidationResult(
            valid=not issues,
            checked_count=checked_count,
            issues=tuple(issues),
        )

    def _validate_template(
        self,
        template_path: Path,
        image_count: int,
    ) -> list[TemplateIssue]:
        try:
            with ZipFile(template_path) as archive:
                try:
                    document_xml = archive.read("word/document.xml")
                except KeyError:
                    return [
                        TemplateIssue(
                            image_count,
                            "missing_document_xml",
                            "缺少 word/document.xml",
                        )
                    ]
        except (BadZipFile, OSError):
            return [TemplateIssue(image_count, "invalid_docx", "DOCX 文件损坏或不可读")]

        try:
            root = ElementTree.fromstring(document_xml)
        except ElementTree.ParseError:
            return [TemplateIssue(image_count, "invalid_document_xml", "文档 XML 无效")]

        text = "".join(root.itertext())
        actual = {match.strip() for match in _PLACEHOLDER_PATTERN.findall(text)}
        expected = _COMMON_PLACEHOLDERS | {
            f"image{index}" for index in range(1, image_count + 1)
        }
        template_issues: list[TemplateIssue] = []
        missing = sorted(expected - actual)
        unexpected = sorted(actual - expected)
        if missing:
            template_issues.append(
                TemplateIssue(
                    image_count,
                    "missing_placeholders",
                    f"缺少占位符：{'、'.join(missing)}",
                )
            )
        if unexpected:
            template_issues.append(
                TemplateIssue(
                    image_count,
                    "unexpected_placeholders",
                    f"多余占位符：{'、'.join(unexpected)}",
                )
            )
        return template_issues
