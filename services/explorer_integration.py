from __future__ import annotations

import sys
from enum import Enum
from pathlib import Path

from config import SUPPORTED_FORMATS


EXTENSIONS = tuple(sorted(SUPPORTED_FORMATS))
VERB_KEY = "PicGenToolAddImages"
MENU_LABEL = "一键添加进PGT工具"


class ExplorerRegistrationStatus(Enum):
    UNSUPPORTED = "unsupported"
    NOT_REGISTERED = "not_registered"
    CURRENT = "current"
    STALE = "stale"
    PARTIAL = "partial"


def _load_winreg():
    import winreg

    return winreg


def _verb_key(extension: str) -> str:
    return rf"Software\Classes\SystemFileAssociations\{extension}\shell\{VERB_KEY}"


def build_command(executable: Path) -> str:
    return f'"{executable}" --add-images "%1"'


def registration_status(executable: Path) -> ExplorerRegistrationStatus:
    if sys.platform != "win32":
        return ExplorerRegistrationStatus.UNSUPPORTED

    winreg = _load_winreg()
    expected_command = build_command(executable)
    expected_icon = f'"{executable}",0'
    states: list[str] = []

    for extension in EXTENSIONS:
        base = _verb_key(extension)
        command_key = base + r"\command"
        base_values = _read_values(winreg, base, ("", "Icon", "MultiSelectModel"))
        command_values = _read_values(winreg, command_key, ("",))
        if base_values is None and command_values is None:
            states.append("absent")
            continue
        if base_values is None or command_values is None:
            states.append("partial")
            continue
        if None in (*base_values, *command_values):
            states.append("partial")
            continue
        current = base_values == (MENU_LABEL, expected_icon, "Player") and command_values == (
            expected_command,
        )
        states.append("current" if current else "stale")

    if all(state == "absent" for state in states):
        return ExplorerRegistrationStatus.NOT_REGISTERED
    if any(state in {"absent", "partial"} for state in states):
        return ExplorerRegistrationStatus.PARTIAL
    if all(state == "current" for state in states):
        return ExplorerRegistrationStatus.CURRENT
    return ExplorerRegistrationStatus.STALE


def _read_values(winreg, path: str, names: tuple[str, ...]) -> tuple[str | None, ...] | None:
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, path) as key:
            values: list[str | None] = []
            for name in names:
                try:
                    values.append(winreg.QueryValueEx(key, name)[0])
                except FileNotFoundError:
                    values.append(None)
            return tuple(values)
    except FileNotFoundError:
        return None
    except PermissionError:
        return tuple(None for _ in names)


def register(executable: Path) -> None:
    if sys.platform != "win32":
        raise OSError("Explorer 右键菜单仅支持 Windows")
    winreg = _load_winreg()
    command = build_command(executable)
    for extension in EXTENSIONS:
        base = _verb_key(extension)
        with winreg.CreateKey(winreg.HKEY_CURRENT_USER, base) as key:
            winreg.SetValueEx(key, "", 0, winreg.REG_SZ, MENU_LABEL)
            winreg.SetValueEx(key, "Icon", 0, winreg.REG_SZ, f'"{executable}",0')
            winreg.SetValueEx(key, "MultiSelectModel", 0, winreg.REG_SZ, "Player")
        with winreg.CreateKey(winreg.HKEY_CURRENT_USER, base + r"\command") as key:
            winreg.SetValueEx(key, "", 0, winreg.REG_SZ, command)


def unregister() -> None:
    if sys.platform != "win32":
        return
    winreg = _load_winreg()
    for extension in EXTENSIONS:
        base = _verb_key(extension)
        for path in (base + r"\command", base):
            try:
                winreg.DeleteKey(winreg.HKEY_CURRENT_USER, path)
            except FileNotFoundError:
                pass


def repair_if_registered(executable: Path, packaged: bool) -> bool:
    if not packaged:
        return False
    status = registration_status(executable)
    if status not in {
        ExplorerRegistrationStatus.STALE,
        ExplorerRegistrationStatus.PARTIAL,
    }:
        return False
    register(executable)
    return True
