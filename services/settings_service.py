from __future__ import annotations

from dataclasses import dataclass
from datetime import date
from pathlib import Path

from PyQt6.QtCore import QSettings

from config import DEFAULT_ACTIVITY_LOCATION


@dataclass(frozen=True, slots=True)
class UserPreferences:
    title: str
    location: str
    activity_date: date
    open_file: bool
    always_on_top: bool
    recent_image_dir: Path | None


class SettingsService:
    def __init__(self, settings: QSettings | None = None) -> None:
        self.settings = settings or QSettings("PicGenTool", "PicGenTool")

    def load(self) -> UserPreferences:
        activity_date = self._load_date(self.settings.value("activity_date", ""))
        recent_image_dir = self._load_directory(
            self.settings.value("recent_image_dir", "")
        )
        return UserPreferences(
            title=str(self.settings.value("title", "")),
            location=str(
                self.settings.value("location", DEFAULT_ACTIVITY_LOCATION)
            ),
            activity_date=activity_date,
            open_file=self._as_bool(self.settings.value("open_file", True), True),
            always_on_top=self._as_bool(
                self.settings.value("always_on_top", True),
                True,
            ),
            recent_image_dir=recent_image_dir,
        )

    def save(self, preferences: UserPreferences) -> None:
        self.settings.setValue("title", preferences.title)
        self.settings.setValue("location", preferences.location)
        self.settings.setValue("activity_date", preferences.activity_date.isoformat())
        self.settings.setValue("open_file", preferences.open_file)
        self.settings.setValue("always_on_top", preferences.always_on_top)
        if preferences.recent_image_dir is None:
            self.settings.remove("recent_image_dir")
        else:
            self.settings.setValue(
                "recent_image_dir",
                str(preferences.recent_image_dir),
            )
        self.settings.sync()

    def clear_title(self) -> None:
        self.settings.setValue("title", "")
        self.settings.sync()

    @staticmethod
    def _load_date(value) -> date:
        try:
            return date.fromisoformat(str(value))
        except (TypeError, ValueError):
            return date.today()

    @staticmethod
    def _load_directory(value) -> Path | None:
        if not value:
            return None
        path = Path(str(value))
        return path if path.is_dir() else None

    @staticmethod
    def _as_bool(value, default: bool) -> bool:
        if isinstance(value, bool):
            return value
        if isinstance(value, str):
            normalized = value.strip().lower()
            if normalized in {"1", "true", "yes", "on"}:
                return True
            if normalized in {"0", "false", "no", "off"}:
                return False
        return default
