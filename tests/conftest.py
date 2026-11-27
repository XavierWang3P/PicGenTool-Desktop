from __future__ import annotations

import os
import sys
from pathlib import Path

import pytest
from PyQt6.QtCore import QSettings


PROJECT_ROOT = Path(__file__).resolve().parents[1]
os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from services.settings_service import SettingsService


@pytest.fixture
def window_settings(tmp_path: Path) -> SettingsService:
    return SettingsService(
        QSettings(str(tmp_path / "window-settings.ini"), QSettings.Format.IniFormat)
    )
