from __future__ import annotations

from pathlib import Path

import pytest

from services import explorer_integration
from services.explorer_integration import ExplorerRegistrationStatus


class _Key:
    def __init__(self, path: str) -> None:
        self.path = path

    def __enter__(self):
        return self

    def __exit__(self, *_args) -> None:
        return None


class FakeWinReg:
    HKEY_CURRENT_USER = object()
    REG_SZ = 1

    def __init__(self) -> None:
        self.values: dict[str, dict[str, str]] = {}

    def CreateKey(self, _root, path: str):
        self.values.setdefault(path, {})
        return _Key(path)

    def OpenKey(self, _root, path: str):
        if path not in self.values:
            raise FileNotFoundError(path)
        return _Key(path)

    def SetValueEx(self, key: _Key, name: str, _reserved, _kind, value: str) -> None:
        self.values[key.path][name] = value

    def QueryValueEx(self, key: _Key, name: str):
        try:
            return self.values[key.path][name], self.REG_SZ
        except KeyError as exc:
            raise FileNotFoundError(name) from exc

    def DeleteKey(self, _root, path: str) -> None:
        if path not in self.values:
            raise FileNotFoundError(path)
        del self.values[path]


@pytest.fixture
def registry(monkeypatch: pytest.MonkeyPatch) -> FakeWinReg:
    fake = FakeWinReg()
    monkeypatch.setattr(explorer_integration.sys, "platform", "win32")
    monkeypatch.setattr(explorer_integration, "_load_winreg", lambda: fake)
    return fake


def test_build_command_quotes_executable_and_selected_file() -> None:
    executable = Path(r"C:\Program Files\PGT\PGT.exe")

    assert explorer_integration.build_command(executable) == (
        r'"C:\Program Files\PGT\PGT.exe" --add-images "%1"'
    )


def test_supported_registry_extensions_include_heic_and_heif() -> None:
    assert {".heic", ".heif"}.issubset(explorer_integration.EXTENSIONS)


def test_registration_status_covers_current_partial_stale_and_missing(
    registry: FakeWinReg,
) -> None:
    current = Path(r"C:\Apps\PGT\PGT.exe")
    old = Path(r"D:\Old\PGT.exe")
    assert explorer_integration.registration_status(current) is ExplorerRegistrationStatus.NOT_REGISTERED

    explorer_integration.register(current)
    assert explorer_integration.registration_status(current) is ExplorerRegistrationStatus.CURRENT

    missing_command = explorer_integration._verb_key(".heic") + r"\command"
    del registry.values[missing_command]
    assert explorer_integration.registration_status(current) is ExplorerRegistrationStatus.PARTIAL

    registry.values.clear()
    explorer_integration.register(old)
    assert explorer_integration.registration_status(current) is ExplorerRegistrationStatus.STALE


def test_non_windows_is_unsupported(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr(explorer_integration.sys, "platform", "linux")

    assert (
        explorer_integration.registration_status(Path("/opt/pgt"))
        is ExplorerRegistrationStatus.UNSUPPORTED
    )


def test_repair_is_conservative_and_only_runs_for_packaged_existing_registration(
    registry: FakeWinReg,
) -> None:
    old = Path(r"D:\Old\PGT.exe")
    current = Path(r"C:\Apps\PGT\PGT.exe")
    explorer_integration.register(old)

    assert not explorer_integration.repair_if_registered(current, packaged=False)
    assert explorer_integration.registration_status(current) is ExplorerRegistrationStatus.STALE
    assert explorer_integration.repair_if_registered(current, packaged=True)
    assert explorer_integration.registration_status(current) is ExplorerRegistrationStatus.CURRENT

    registry.values.clear()
    assert not explorer_integration.repair_if_registered(current, packaged=True)
    assert registry.values == {}
