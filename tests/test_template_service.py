from __future__ import annotations

from pathlib import Path
from zipfile import ZipFile

from config import MAX_IMAGES
from services.template_service import TemplateService


COMMON_PLACEHOLDERS = {"title", "location", "dateForDoc"}


def _expected_placeholders(image_count: int) -> set[str]:
    return COMMON_PLACEHOLDERS | {f"image{index}" for index in range(1, image_count + 1)}


def _write_template(
    directory: Path,
    image_count: int,
    placeholders: set[str] | None = None,
    *,
    include_document_xml: bool = True,
) -> None:
    path = directory / f"Template-{image_count}.docx"
    with ZipFile(path, "w") as archive:
        archive.writestr("[Content_Types].xml", "<Types />")
        if include_document_xml:
            values = placeholders if placeholders is not None else _expected_placeholders(image_count)
            body = "".join(f"<w:t>{{{{{value}}}}}</w:t>" for value in sorted(values))
            archive.writestr(
                "word/document.xml",
                '<w:document xmlns:w="urn:test"><w:body>' + body + "</w:body></w:document>",
            )


def _write_valid_set(directory: Path) -> None:
    directory.mkdir(parents=True, exist_ok=True)
    for image_count in range(1, MAX_IMAGES + 1):
        _write_template(directory, image_count)


def test_real_templates_pass_deep_validation() -> None:
    result = TemplateService().validate_all(deep=True)

    assert result.valid
    assert result.checked_count == MAX_IMAGES
    assert result.issues == ()


def test_validate_all_reports_missing_template_in_light_mode(tmp_path: Path) -> None:
    _write_valid_set(tmp_path)
    (tmp_path / "Template-7.docx").unlink()

    result = TemplateService(tmp_path).validate_all(deep=False)

    assert not result.valid
    assert result.checked_count == MAX_IMAGES - 1
    assert [(issue.template_number, issue.code) for issue in result.issues] == [
        (7, "missing_file")
    ]


def test_validate_all_reports_corrupt_docx(tmp_path: Path) -> None:
    _write_valid_set(tmp_path)
    (tmp_path / "Template-4.docx").write_bytes(b"not-a-zip")

    result = TemplateService(tmp_path).validate_all(deep=True)

    assert (4, "invalid_docx") in [
        (issue.template_number, issue.code) for issue in result.issues
    ]


def test_validate_all_reports_missing_document_xml(tmp_path: Path) -> None:
    _write_valid_set(tmp_path)
    _write_template(tmp_path, 8, include_document_xml=False)

    result = TemplateService(tmp_path).validate_all(deep=True)

    assert (8, "missing_document_xml") in [
        (issue.template_number, issue.code) for issue in result.issues
    ]


def test_validate_all_reports_missing_and_unexpected_placeholders(tmp_path: Path) -> None:
    _write_valid_set(tmp_path)
    placeholders = _expected_placeholders(3)
    placeholders.remove("location")
    placeholders.remove("image2")
    placeholders.add("image4")
    _write_template(tmp_path, 3, placeholders)

    result = TemplateService(tmp_path).validate_all(deep=True)

    issues = {(issue.code, issue.message) for issue in result.issues if issue.template_number == 3}
    assert ("missing_placeholders", "缺少占位符：image2、location") in issues
    assert ("unexpected_placeholders", "多余占位符：image4") in issues
