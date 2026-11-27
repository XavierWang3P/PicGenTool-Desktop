from __future__ import annotations

from pathlib import Path
import shutil
import time

import pytest
from PIL import Image, ImageDraw

from config import DPI, TARGET_HEIGHT_CM, TARGET_WIDTH_CM
from services.image_service import ImageService


def test_prepare_images_decodes_heic_and_heif_to_rgb_jpeg(tmp_path: Path) -> None:
    fixture_dir = Path(__file__).resolve().parent / "fixtures"
    expected_digest = "003c2834ea3609a1f314fd01feb48046aa1816209c0d7317bf347585bc6a56b4"
    sources: list[Path] = []
    for suffix in (".heic", ".heif"):
        fixture = fixture_dir / f"sample{suffix}"
        assert fixture.is_file()
        assert ImageService()._content_digest(fixture) == expected_digest
        shutil.copyfile(fixture, tmp_path / f"source{suffix}")
        sources.append(tmp_path / f"source{suffix}")

    prepared = ImageService().prepare_images(sources)
    try:
        assert len(prepared.image_paths) == 2
        for output_path in prepared.image_paths:
            with Image.open(output_path) as output:
                assert output.format == "JPEG"
                assert output.mode == "RGB"
    finally:
        prepared.cleanup()


def test_deduplicate_removes_same_path_and_identical_copy_preserving_first_order(
    tmp_path: Path,
) -> None:
    first = tmp_path / "first.jpg"
    duplicate_copy = tmp_path / "copy.jpg"
    distinct = tmp_path / "distinct.jpg"
    first.write_bytes(b"same-image-content")
    shutil.copyfile(first, duplicate_copy)
    distinct.write_bytes(b"different-content")

    result = ImageService().deduplicate(
        [first, first, duplicate_copy, distinct],
    )

    assert result.paths == [first, distinct]
    assert result.duplicate_count == 2
    assert result.failures == ()


def test_deduplicate_keeps_same_size_different_content(tmp_path: Path) -> None:
    first = tmp_path / "first.jpg"
    second = tmp_path / "second.jpg"
    first.write_bytes(b"abc")
    second.write_bytes(b"xyz")

    result = ImageService().deduplicate([first, second])

    assert result.paths == [first, second]
    assert result.duplicate_count == 0


def test_deduplicate_compares_new_files_with_existing_selection(tmp_path: Path) -> None:
    existing = tmp_path / "existing.jpg"
    duplicate = tmp_path / "duplicate.jpg"
    new_image = tmp_path / "new.jpg"
    existing.write_bytes(b"existing-content")
    duplicate.write_bytes(b"existing-content")
    new_image.write_bytes(b"new-content")

    result = ImageService().deduplicate([duplicate, new_image], existing=[existing])

    assert result.paths == [new_image]
    assert result.duplicate_count == 1


def test_deduplicate_unreadable_file_reports_safe_failure_and_keeps_readable_files(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    private_directory = tmp_path / "private-customer-album"
    private_directory.mkdir()
    broken = private_directory / "broken.jpg"
    readable = private_directory / "readable.jpg"
    broken.write_bytes(b"broken")
    readable.write_bytes(b"readable")
    service = ImageService()
    original_digest = service._content_digest

    def fail_broken(path: Path) -> str:
        if path == broken:
            raise PermissionError(f"cannot read {private_directory}")
        return original_digest(path)

    monkeypatch.setattr(service, "_content_digest", fail_broken)

    result = service.deduplicate([broken, readable])

    assert result.paths == [readable]
    assert len(result.failures) == 1
    assert result.failures[0].filename == "broken.jpg"
    assert str(private_directory) not in result.failures[0].reason


def test_prepare_images_preserves_count_size_mode_and_order(tmp_path: Path) -> None:
    colors = [(240, 20, 20), (20, 240, 20), (20, 20, 240)]
    source_paths: list[Path] = []
    for index, color in enumerate(colors, start=1):
        source_path = tmp_path / f"source-{index}.png"
        Image.new("RGB", (900, 600), color).save(source_path)
        source_paths.append(source_path)

    prepared = ImageService().prepare_images(source_paths)
    try:
        assert len(prepared.image_paths) == 3
        assert [path.name for path in prepared.image_paths] == [
            "image_01.jpg",
            "image_02.jpg",
            "image_03.jpg",
        ]

        expected_size = (
            int(TARGET_WIDTH_CM * DPI / 2.54),
            int(TARGET_HEIGHT_CM * DPI / 2.54),
        )
        observed_colors: list[tuple[int, int, int]] = []
        for output_path in prepared.image_paths:
            with Image.open(output_path) as image:
                assert image.size == expected_size
                assert image.mode == "RGB"
                observed_colors.append(image.getpixel((0, 0)))

        for observed, expected in zip(observed_colors, colors, strict=True):
            assert all(abs(actual - wanted) <= 5 for actual, wanted in zip(observed, expected))
    finally:
        prepared.cleanup()


def test_prepare_one_applies_exif_orientation_before_crop(tmp_path: Path) -> None:
    source_path = tmp_path / "oriented.jpg"
    source = Image.new("RGB", (120, 80))
    draw = ImageDraw.Draw(source)
    draw.rectangle((0, 0, 59, 39), fill=(240, 20, 20))
    draw.rectangle((60, 0, 119, 39), fill=(20, 240, 20))
    draw.rectangle((0, 40, 59, 79), fill=(20, 20, 240))
    draw.rectangle((60, 40, 119, 79), fill=(240, 240, 20))
    exif = Image.Exif()
    exif[274] = 6
    source.save(source_path, format="JPEG", quality=95, exif=exif)
    source.close()

    output_path = tmp_path / "prepared.jpg"
    service = ImageService()
    target_size = (80, 80)
    service._prepare_one(source_path, output_path, target_size, target_size[0] / target_size[1])

    with Image.open(output_path) as actual:
        expected_corners = [(20, 20, 240), (240, 20, 20), (240, 240, 20), (20, 240, 20)]
        observed_corners = [
            actual.getpixel((5, 5)),
            actual.getpixel((75, 5)),
            actual.getpixel((5, 75)),
            actual.getpixel((75, 75)),
        ]
    for observed, expected in zip(observed_corners, expected_corners, strict=True):
        assert all(abs(actual - wanted) <= 35 for actual, wanted in zip(observed, expected))


@pytest.mark.parametrize(
    ("orientation", "expected_corners"),
    [
        (6, [(20, 20, 240), (240, 20, 20), (240, 240, 20), (20, 240, 20)]),
        (8, [(20, 240, 20), (240, 240, 20), (240, 20, 20), (20, 20, 240)]),
    ],
)
def test_prepare_one_handles_exif_6_and_8_visual_orientation(
    tmp_path: Path,
    orientation: int,
    expected_corners: list[tuple[int, int, int]],
) -> None:
    source_path = tmp_path / f"oriented-{orientation}.jpg"
    source = Image.new("RGB", (120, 80))
    draw = ImageDraw.Draw(source)
    draw.rectangle((0, 0, 59, 39), fill=(240, 20, 20))
    draw.rectangle((60, 0, 119, 39), fill=(20, 240, 20))
    draw.rectangle((0, 40, 59, 79), fill=(20, 20, 240))
    draw.rectangle((60, 40, 119, 79), fill=(240, 240, 20))
    exif = Image.Exif()
    exif[274] = orientation
    source.save(source_path, format="JPEG", quality=95, exif=exif)
    source.close()

    output_path = tmp_path / f"prepared-{orientation}.jpg"
    ImageService()._prepare_one(source_path, output_path, (80, 80), 1.0)

    with Image.open(output_path) as output:
        observed_corners = [
            output.getpixel((5, 5)),
            output.getpixel((75, 5)),
            output.getpixel((5, 75)),
            output.getpixel((75, 75)),
        ]
    for observed, expected in zip(observed_corners, expected_corners, strict=True):
        assert all(abs(actual - wanted) <= 35 for actual, wanted in zip(observed, expected))


@pytest.mark.parametrize("mode", ["RGBA", "L"])
def test_prepare_one_converts_rgba_and_grayscale_to_rgb(tmp_path: Path, mode: str) -> None:
    source_path = tmp_path / f"source-{mode}.png"
    color = (10, 20, 30, 128) if mode == "RGBA" else 120
    Image.new(mode, (100, 80), color).save(source_path)
    output_path = tmp_path / f"output-{mode}.jpg"

    ImageService()._prepare_one(source_path, output_path, (80, 60), 4 / 3)

    with Image.open(output_path) as output:
        assert output.mode == "RGB"


def test_prepare_one_uses_configured_quality_without_optimize(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    source_path = tmp_path / "source.jpg"
    Image.new("RGB", (100, 80), (20, 40, 60)).save(source_path)
    captured: dict[str, object] = {}

    def capture_save(_image, output_path, **kwargs):
        captured["output_path"] = output_path
        captured.update(kwargs)

    monkeypatch.setattr(Image.Image, "save", capture_save)
    output_path = tmp_path / "output.jpg"

    ImageService()._prepare_one(source_path, output_path, (80, 60), 4 / 3)

    assert captured == {
        "output_path": output_path,
        "format": "JPEG",
        "quality": 60,
        "optimize": False,
        "dpi": (240, 240),
    }


def test_worker_count_is_bounded_by_cpu_config_and_image_count(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    import services.image_service as image_service_module

    service = ImageService()
    monkeypatch.setattr(image_service_module, "MAX_IMAGE_WORKERS", 4)
    monkeypatch.setattr(image_service_module.os, "cpu_count", lambda: 8)
    assert service._worker_count(2) == 2
    assert service._worker_count(30) == 4

    monkeypatch.setattr(image_service_module.os, "cpu_count", lambda: 2)
    assert service._worker_count(30) == 2

    monkeypatch.setattr(image_service_module.os, "cpu_count", lambda: None)
    assert service._worker_count(30) == 1


def test_parallel_completion_preserves_input_order(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    sources = [tmp_path / f"source-{index}.jpg" for index in range(3)]
    completion_order: list[int] = []

    def prepare_out_of_order(source, output, _target_size, _target_ratio):
        index = sources.index(source)
        time.sleep((2 - index) * 0.02)
        output.write_text(source.name, encoding="utf-8")
        completion_order.append(index)
        return output

    service = ImageService()
    monkeypatch.setattr(service, "_prepare_one", prepare_out_of_order)
    prepared = service.prepare_images(sources)
    try:
        assert [path.name for path in prepared.image_paths] == [
            "image_01.jpg",
            "image_02.jpg",
            "image_03.jpg",
        ]
        assert [path.read_text(encoding="utf-8") for path in prepared.image_paths] == [
            path.name for path in sources
        ]
        assert completion_order != [0, 1, 2]
    finally:
        prepared.cleanup()


def test_progress_is_monotonic_and_reaches_60(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    sources = [tmp_path / f"source-{index}.jpg" for index in range(4)]
    values: list[int] = []

    def prepare(source, output, _target_size, _target_ratio):
        output.write_text(source.name, encoding="utf-8")
        return output

    service = ImageService()
    monkeypatch.setattr(service, "_prepare_one", prepare)
    prepared = service.prepare_images(sources, lambda _message, value: values.append(value))
    try:
        assert values == sorted(values)
        assert values[-1] == 60
    finally:
        prepared.cleanup()


def test_worker_failure_cleans_temp_directory_and_names_source_file(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    sources = [tmp_path / "good.jpg", tmp_path / "broken.jpg", tmp_path / "later.jpg"]
    output_directories: list[Path] = []

    def prepare(source, output, _target_size, _target_ratio):
        output_directories.append(output.parent)
        if source.name == "broken.jpg":
            raise OSError("invalid image")
        time.sleep(0.01)
        output.write_text(source.name, encoding="utf-8")
        return output

    service = ImageService()
    monkeypatch.setattr(service, "_prepare_one", prepare)

    with pytest.raises(ValueError, match="broken.jpg"):
        service.prepare_images(sources)

    assert output_directories
    assert all(not directory.exists() for directory in output_directories)


def test_corrupt_image_reports_source_name(tmp_path: Path) -> None:
    private_directory = tmp_path / "private-customer-album"
    private_directory.mkdir()
    corrupt_path = private_directory / "corrupt.jpg"
    corrupt_path.write_bytes(b"not a valid image")

    with pytest.raises(ValueError, match="corrupt.jpg") as error:
        ImageService().prepare_images([corrupt_path])

    assert str(private_directory) not in str(error.value)


def test_prepare_images_does_not_expose_path_from_underlying_exception(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    private_directory = tmp_path / "private-customer-album"
    source_path = private_directory / "photo.jpg"
    service = ImageService()

    def fail(*_args):
        raise OSError(f"decoder failed for {source_path}")

    monkeypatch.setattr(service, "_prepare_one", fail)

    with pytest.raises(ValueError, match="photo.jpg") as error:
        service.prepare_images([source_path])

    assert str(private_directory) not in str(error.value)
