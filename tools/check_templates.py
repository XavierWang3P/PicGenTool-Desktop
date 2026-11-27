from __future__ import annotations

import argparse
import sys
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from config import MAX_IMAGES  # noqa: E402
from services.template_service import TemplateService  # noqa: E402


def main() -> int:
    parser = argparse.ArgumentParser(description="检查 PicGenTool Word 模板完整性")
    parser.add_argument(
        "--template-dir",
        type=Path,
        default=PROJECT_ROOT / "templates",
        help="模板目录",
    )
    args = parser.parse_args()

    result = TemplateService(args.template_dir).validate_all(deep=True)
    if result.valid:
        print(f"{result.checked_count}/{MAX_IMAGES} templates valid")
        return 0

    print(
        f"Template validation failed: checked {result.checked_count}/{MAX_IMAGES}",
        file=sys.stderr,
    )
    for issue in result.issues:
        print(
            f"Template-{issue.template_number}: {issue.code}: {issue.message}",
            file=sys.stderr,
        )
    return 1


if __name__ == "__main__":
    raise SystemExit(main())
