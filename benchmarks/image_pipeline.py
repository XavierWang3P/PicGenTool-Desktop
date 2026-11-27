from __future__ import annotations

import argparse
import json
import statistics
import sys
import tempfile
import time
from pathlib import Path

from PIL import Image, ImageDraw


PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from services.image_service import ImageService


BENCHMARK_COUNTS = (1, 10, 30)
SOURCE_SIZE = (4032, 3024)


def _create_sources(directory: Path, count: int) -> list[Path]:
    paths: list[Path] = []
    for index in range(count):
        base = ((index * 47) % 256, (index * 83) % 256, (index * 131) % 256)
        image = Image.new("RGB", SOURCE_SIZE, base)
        draw = ImageDraw.Draw(image)
        band_height = SOURCE_SIZE[1] // 32
        for band in range(32):
            color = tuple((channel + band * 7) % 256 for channel in base)
            top = band * band_height
            draw.rectangle((0, top, SOURCE_SIZE[0], top + band_height), fill=color)
        path = directory / f"sample-{index + 1:02d}.jpg"
        image.save(path, format="JPEG", quality=92, exif=Image.Exif())
        image.close()
        paths.append(path)
    return paths


def _sources_from(directory: Path, count: int) -> list[Path]:
    supported = ImageService().filter_supported(
        [str(path) for path in sorted(directory.iterdir()) if path.is_file()]
    )
    if len(supported) < count:
        raise ValueError(f"源目录至少需要 {count} 张受支持的图片")
    return supported[:count]


def run_benchmark(
    count: int,
    runs: int,
    source_dir: Path | None = None,
) -> dict[str, object]:
    if count not in BENCHMARK_COUNTS:
        raise ValueError(f"count 必须是 {BENCHMARK_COUNTS} 之一")
    if runs < 1:
        raise ValueError("runs 必须大于 0")

    service = ImageService()
    with tempfile.TemporaryDirectory(prefix="pgt-benchmark-") as temp_name:
        sources = (
            _sources_from(source_dir, count)
            if source_dir is not None
            else _create_sources(Path(temp_name), count)
        )
        runs_ms: list[float] = []
        for _ in range(runs):
            started_at = time.perf_counter()
            prepared = service.prepare_images(sources)
            try:
                runs_ms.append(round((time.perf_counter() - started_at) * 1000, 3))
            finally:
                prepared.cleanup()

    return {
        "count": count,
        "runs_ms": runs_ms,
        "median_ms": round(statistics.median(runs_ms), 3),
    }


def main() -> int:
    parser = argparse.ArgumentParser(description="PicGenTool 图片流水线性能基准")
    parser.add_argument("--count", type=int, choices=BENCHMARK_COUNTS, required=True)
    parser.add_argument("--runs", type=int, default=3)
    parser.add_argument("--source-dir", type=Path)
    parser.add_argument("--output", type=Path, required=True)
    args = parser.parse_args()

    result = run_benchmark(args.count, args.runs, args.source_dir)
    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(
        json.dumps(result, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )
    print(json.dumps(result, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
