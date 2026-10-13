from __future__ import annotations

from shared import settings

APP_TITLE = "Finance Automation Center"
APP_THEME = "flatly"
APP_SIZE = (1500, 920)
APP_MIN_SIZE = (1200, 760)

TEST_MODE = bool(settings.get("test_mode"))
