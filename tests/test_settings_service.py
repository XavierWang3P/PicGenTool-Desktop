from __future__ import annotations

from datetime import date
from pathlib import Path

from PyQt6.QtCore import QSettings

from config import DEFAULT_ACTIVITY_LOCATION
from services.settings_service import SettingsService, UserPreferences


def _service(path: Path) -> SettingsService:
    return SettingsService(QSettings(str(path), QSettings.Format.IniFormat))


def test_preferences_round_trip_without_private_paths_or_image_list(tmp_path: Path) -> None:
    ini_path = tmp_path / "settings.ini"
    recent = tmp_path / "photos"
    recent.mkdir()
    preferences = UserPreferences(
        title="Activity",
        location="Hall",
        activity_date=date(2026, 11, 26),
        open_file=False,
        always_on_top=False,
        recent_image_dir=recent,
    )

    _service(ini_path).save(preferences)
    loaded = _service(ini_path).load()
    raw = ini_path.read_text(encoding="utf-8")

    assert loaded == preferences
    assert "image_paths" not in raw
    assert "output" not in raw.lower()


def test_invalid_date_and_missing_recent_directory_fall_back_safely(tmp_path: Path) -> None:
    ini_path = tmp_path / "settings.ini"
    settings = QSettings(str(ini_path), QSettings.Format.IniFormat)
    settings.setValue("activity_date", "not-a-date")
    settings.setValue("recent_image_dir", str(tmp_path / "missing"))
    settings.sync()

    loaded = SettingsService(settings).load()

    assert loaded.activity_date == date.today()
    assert loaded.recent_image_dir is None


def test_clear_title_preserves_other_preferences(tmp_path: Path) -> None:
    service = _service(tmp_path / "settings.ini")
    service.save(
        UserPreferences(
            title="Activity",
            location="Hall",
            activity_date=date(2026, 11, 26),
            open_file=True,
            always_on_top=True,
            recent_image_dir=None,
        )
    )

    service.clear_title()
    loaded = service.load()

    assert loaded.title == ""
    assert loaded.location == "Hall"
    assert loaded.activity_date == date(2026, 11, 26)


def test_defaults_are_stable(tmp_path: Path) -> None:
    loaded = _service(tmp_path / "settings.ini").load()

    assert loaded.title == ""
    assert loaded.location == DEFAULT_ACTIVITY_LOCATION
    assert loaded.activity_date == date.today()
    assert loaded.open_file
    assert loaded.always_on_top
