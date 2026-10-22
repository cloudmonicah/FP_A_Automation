"""All output goes to the Teams "Validation Results" folder in the person's own OneDrive."""

import pytest

from shared import output_root, settings
from shared.friendly_errors import PlainLanguageError

NAME = output_root.SHARED_FOLDER_NAME


@pytest.fixture
def onedrive(tmp_path, monkeypatch):
    home = tmp_path / "Users" / "jsmith"
    drive = home / "OneDrive - Aliv"
    drive.mkdir(parents=True)
    monkeypatch.setenv("USERPROFILE", str(home))
    monkeypatch.setenv("OneDriveCommercial", str(drive))
    monkeypatch.delenv("OneDrive", raising=False)
    monkeypatch.setenv("LOCALAPPDATA", str(home / "AppData" / "Local"))
    real = settings.get_path
    monkeypatch.setattr(settings, "get_path", lambda key: None if key.endswith("_root") else real(key))
    return drive


def test_finds_the_teams_shortcut_in_this_persons_onedrive(onedrive):
    folder = onedrive / "Shortcuts" / NAME
    folder.mkdir(parents=True)
    assert output_root.shared_output_root() == folder


def test_shortcut_at_the_top_of_onedrive_is_found_too(onedrive):
    (onedrive / NAME).mkdir()
    assert output_root.shared_output_root() == onedrive / NAME


def test_never_falls_back_to_someone_elses_folder(onedrive):
    (onedrive / "Shortcuts" / "Stacy Brown-Johnson's files - Actuals_Validation_Results").mkdir(parents=True)
    with pytest.raises(PlainLanguageError) as error:
        output_root.shared_output_root()
    assert "Add shortcut to My files" in str(error.value) and "Validation Results" in str(error.value)


def test_settings_can_point_elsewhere(onedrive, tmp_path, monkeypatch):
    other = tmp_path / "Server" / "FP&A"
    other.mkdir(parents=True)
    monkeypatch.setattr(settings, "get_path", lambda key: other if key == "shared_output_root" else None)
    assert output_root.shared_output_root() == other


def test_forecast_and_actuals_share_one_tree(onedrive):
    from automations.actuals.validation_output_folders import (
        create_validation_output_folders,
        find_actuals_validation_results_folder,
    )
    from shared.forecast_output_paths import get_forecast_validation_output_folder

    root = onedrive / "Shortcuts" / NAME
    root.mkdir(parents=True)
    forecast = get_forecast_validation_output_folder("FY27", 10, 2026, create=True)
    actuals = create_validation_output_folders(find_actuals_validation_results_folder(), "Oct-26")
    assert forecast == root / "FY27" / "2026-10 - October" / "Forecast Validation"
    assert actuals["month_folder"] == root / "FY27" / "2026-10 - October"
    assert actuals["forecast_validation"] == forecast


def test_tableau_temp_copies_stay_on_this_pc(onedrive, tmp_path):
    from shared.tableau_temp_copies import tableau_temp_folder

    folder = tableau_temp_folder()
    assert folder == tmp_path / "Users" / "jsmith" / "AppData" / "Local" / "FPA Automation" / "Tableau Temp Copies"
    assert onedrive not in folder.parents


def test_missing_shared_folder_stops_forecast_with_a_plain_message(onedrive):
    from shared.forecast_output_paths import get_forecast_validation_output_folder

    with pytest.raises(PlainLanguageError, match="not on this computer"):
        get_forecast_validation_output_folder("FY27", 10, 2026, create=True)
