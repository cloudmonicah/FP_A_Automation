import json

import pytest

from shared import settings


@pytest.fixture
def isolated(tmp_path, monkeypatch):
    app = tmp_path / "app"
    app.mkdir()
    monkeypatch.setattr(settings, "app_dir", lambda: app)
    monkeypatch.setenv("APPDATA", str(tmp_path / "appdata"))
    settings.load(refresh=True)
    yield app
    settings.load(refresh=True)


def test_defaults(isolated):
    assert settings.get("workbooks_root") == settings.DEFAULTS["workbooks_root"]
    assert settings.get_path("validation_output_root") is None


def test_layering_and_save(isolated):
    (isolated / "settings.json").write_text(
        json.dumps({"workbooks_root": "/shared/wb", "unknown": 1}), encoding="utf-8"
    )
    settings.load(refresh=True)
    assert settings.get("workbooks_root") == "/shared/wb"
    assert "unknown" not in settings.load()

    path = settings.save_user_settings({"test_mode": True, "workbooks_root": "/shared/wb"})
    saved = json.loads(path.read_text(encoding="utf-8"))
    assert saved == {"test_mode": True}  # values equal to the shared file are not stored
    assert settings.get("test_mode") is True
    assert settings.workbook_path("OPEX", "Legal.twb").replace("\\", "/") == "/shared/wb/OPEX/Legal.twb"


def test_bad_json_is_ignored(isolated):
    (isolated / "settings.json").write_text("{not json", encoding="utf-8")
    assert settings.load(refresh=True)["log_level"] == "INFO"


def test_environment_variables_expand(isolated, monkeypatch):
    monkeypatch.setenv("FPA_TEST_HOME", "/users/alex")
    settings.save_user_settings({"workbooks_root": "$FPA_TEST_HOME/Workbooks"})
    assert settings.workbook_path("OPEX").replace("\\", "/") == "/users/alex/Workbooks/OPEX"
