"""Register or remove the Windows Explorer context menu entry.

Per-user registration under HKCU, no administrator rights required.
Run after moving or rebuilding the application to refresh the command path.
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from services.explorer_integration import _verb_key, register, unregister  # noqa: E402,F401


def main() -> int:
    parser = argparse.ArgumentParser(description="注册图片右键菜单项")
    parser.add_argument(
        "--exe",
        type=Path,
        default=PROJECT_ROOT / "dist" / "PGT" / "PGT.exe",
        help="PGT.exe 路径",
    )
    parser.add_argument("--unregister", action="store_true", help="移除右键菜单项")
    args = parser.parse_args()

    if sys.platform != "win32":
        print("仅支持 Windows")
        return 1

    if args.unregister:
        unregister()
        print("已移除右键菜单项")
        return 0

    executable = args.exe.resolve()
    if not executable.is_file():
        print(f"找不到可执行文件：{executable}")
        return 1
    register(executable)
    print(f"已注册右键菜单项，命令目标：{executable}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
