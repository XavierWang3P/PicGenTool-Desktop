from __future__ import annotations

from datetime import date
from pathlib import Path
from zipfile import ZipFile

import pytest
from PIL import Image

from models.generation_task import GenerationTask
from services.document_service import DocumentService
from services.image_service import ImageService


@pytest.mark.parametrize("image_count", [1, 10, 23, 30])
def test_real_templates_generate_compatible_docx(tmp_path: Path, image_count: int) -> None:
    source_paths: list[Path] = []
    for index in range(image_count):
        path = tmp_path / f"source-{index + 1:02d}.jpg"
        Image.new(
            "RGB",
            (1000, 750),
            ((index * 37) % 256, (index * 73) % 256, (index * 109) % 256),
        ).save(path, format="JPEG", quality=90)
        source_paths.append(path)

    output_path = tmp_path / f"result-{image_count}.docx"
    task = GenerationTask(
        title="性能验收活动",
        location="测试地点",
        activity_date=date(2026, 11, 25),
        image_paths=source_paths,
        save_path=output_path,
        open_file=False,
    )
    prepared = ImageService().prepare_images(source_paths)
    try:
        result = DocumentService().generate(task, prepared.image_paths)
    finally:
        prepared.cleanup()

    assert result == output_path
    assert output_path.exists()
    with ZipFile(output_path) as archive:
        document_xml = archive.read("word/document.xml").decode("utf-8")
        media_files = [name for name in archive.namelist() if name.startswith("word/media/")]

    assert "性能验收活动" in document_xml
    assert "测试地点" in document_xml
    assert "2026年11月25日" in document_xml
    assert len(media_files) == image_count
