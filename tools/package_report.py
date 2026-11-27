from __future__ import annotations

import argparse
import json
from dataclasses import dataclass
from pathlib import Path


@dataclass(frozen=True, slots=True)
class PackageReport:
    root: Path
    file_count: int
    total_bytes: int
    groups: dict[str, int]
    largest_files: list[tuple[str, int]]
    max_bytes: int | None = None

    @property
    def total_mib(self) -> float:
        return self.total_bytes / (1024 * 1024)

    @property
    def within_budget(self) -> bool:
        return self.max_bytes is None or self.total_bytes <= self.max_bytes


def scan_distribution(root: Path, max_bytes: int | None = None) -> PackageReport:
    root = root.resolve()
    files = sorted((path for path in root.rglob("*") if path.is_file()), key=lambda path: path.relative_to(root).as_posix())
    largest_files = sorted(
        ((path.relative_to(root).as_posix(), path.stat().st_size) for path in files),
        key=lambda item: (-item[1], item[0]),
    )
    groups: dict[str, int] = {}
    for relative_path, size in largest_files:
        group = relative_path.split("/", 1)[0] if "/" in relative_path else "(root)"
        groups[group] = groups.get(group, 0) + size
    return PackageReport(
        root=root,
        file_count=len(files),
        total_bytes=sum(size for _, size in largest_files),
        groups=dict(sorted(groups.items(), key=lambda item: (-item[1], item[0]))),
        largest_files=largest_files[:20],
        max_bytes=max_bytes,
    )


def build_report(root: Path, max_bytes: int | None = None) -> PackageReport:
    return scan_distribution(root, max_bytes=max_bytes)


def report_to_dict(report: PackageReport) -> dict[str, object]:
    return {
        "root": str(report.root),
        "file_count": report.file_count,
        "total_bytes": report.total_bytes,
        "total_mib": report.total_mib,
        "max_bytes": report.max_bytes,
        "within_budget": report.within_budget,
        "groups": report.groups,
        "largest_files": [
            {"path": path, "bytes": size}
            for path, size in report.largest_files
        ],
    }


def write_report(report: PackageReport, destination: Path) -> None:
    destination.write_text(
        json.dumps(report_to_dict(report), ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )


def main() -> int:
    parser = argparse.ArgumentParser(description="报告 PyInstaller 发布目录体积")
    parser.add_argument("distribution", type=Path)
    parser.add_argument("--max-mib", type=float)
    parser.add_argument("--json", type=Path)
    args = parser.parse_args()

    if not args.distribution.is_dir():
        parser.error(f"distribution directory does not exist: {args.distribution}")
    max_bytes = None if args.max_mib is None else int(args.max_mib * 1024 * 1024)
    report = build_report(args.distribution, max_bytes=max_bytes)
    if args.json:
        write_report(report, args.json)

    print(f"Files: {report.file_count}")
    print(f"Total: {report.total_mib:.2f} MiB ({report.total_bytes} bytes)")
    if report.max_bytes is not None:
        print(f"Budget: {args.max_mib:.2f} MiB")
    print("Largest files:")
    for path, size in report.largest_files:
        print(f"  {size / 1024 / 1024:8.2f} MiB  {path}")
    print("Groups:")
    for group, size in report.groups.items():
        print(f"  {size / 1024 / 1024:8.2f} MiB  {group}")

    if not report.within_budget:
        print("Package size budget exceeded")
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
