# Rolling out FP&A Automation to colleagues

## What each colleague needs

| Requirement | Why |
| --- | --- |
| Windows 10 or 11 | The Tableau steps drive Tableau Desktop through Windows UI Automation. |
| Tableau Desktop installed, licensed and signed in, the same version as yours | The app opens `.twb` files with whatever program Windows associates with them, then reads Tableau's windows. A different version can change window and menu names. |
| `.twb` files open in Tableau Desktop when double-clicked | Same reason. |
| Read access to the shared Tableau workbooks folder | Set as `workbooks_root` in Settings. |
| Read access to the financial statements | They pick these in the app each month. |
| Write access to the validation output folder | Reports are saved there. |
| Excel | Only to open the reports. Close a report in Excel before re-running that validation, or Windows will block the save. |

Colleagues do **not** need Python.

### While a Tableau step is running
- Don't use the mouse or keyboard, and don't lock the screen. The app is clicking through Tableau.
- Run the app normally, not "as administrator". Windows blocks UI automation between programs running at different privilege levels.
- Tableau automation can be sensitive to display scaling. If Tableau steps fail on one PC only, compare Settings > Display > Scale with a PC where they work.

## Building the app (you, on a Windows PC)

1. Install Python 3.13 from python.org (tick "Add to PATH").
2. In the project folder, copy `settings.example.json` to `settings.json` and fill in the shared-drive paths. Prefer UNC paths (`\\server\share\...`) over drive letters, because colleagues may map drives differently. Commit this file, so every build ships the same defaults.
3. Double-click `build_exe.bat`. It creates a clean build environment, runs the tests, builds the app and copies `settings.json` into the build. The build stops if any test fails.
4. The result is the folder `dist\FPA Automation\`. Run `FPA Automation.exe` inside it once yourself.

## Using Microsoft 365 (SharePoint / Teams) instead of a network drive

Keep three things apart:

| What | Where | Who |
| --- | --- | --- |
| Source code | GitHub, cloned to a local, non-OneDrive folder such as `C:\Dev\FPA_Automation` | Only whoever maintains the app |
| The built app | A SharePoint/Teams library, e.g. *FP&A team > Documents > Apps* | Everyone, read-only |
| Workbooks, statements, reports | SharePoint/Teams libraries, synced to each PC | Everyone |

The app reads and writes ordinary Windows folders, so SharePoint libraries must be **synced** to each PC: open the library in the browser and click **Sync** (or **Add shortcut to My files**). Then:

- In File Explorer, right-click the synced Workbooks folder and choose **Always keep on this device**. Tableau and the app can't open files that are still "online-only".
- Synced folders live under each person's profile, so their paths differ per person. Use `%USERPROFILE%` in `settings.json` so one file works for everyone, e.g.
  `"workbooks_root": "%USERPROFILE%\\Aliv\\FP&A - Workbooks"`.
  Check the exact folder name on one PC: it is usually `C:\Users\<name>\<Company>\<Site> - <Library>`.
- Have one person run a validation at a time. SharePoint sync can create "conflict" copies if two people save the same report simultaneously.

For releasing the app itself, either colleagues copy the app folder from the Apps library to `C:\Apps\FPA Automation`, or ask IT to package it in **Intune / Company Portal**, which installs and updates it on every PC for you.

## Releasing a version

1. Copy the whole `dist\FPA Automation` folder to the shared drive, named with its version, e.g. `\\server\FP&A\Apps\FPA Automation 1.0`.
2. Colleagues copy that folder to their own PC (e.g. `C:\Apps\FPA Automation`) and pin `FPA Automation.exe` to the taskbar. A local copy starts much faster than running it from the network.
3. For an update, release a new numbered folder and ask colleagues to replace their copy. Keep the previous version folder so you can roll back.

Windows SmartScreen or antivirus may warn about a new, unsigned .exe. Ask IT to allow the shared app folder, or to sign the .exe with the company's code-signing certificate.

## Settings

Settings are layered, later ones win:

1. Built-in defaults (`shared/settings.py`).
2. `settings.json` next to `FPA Automation.exe`: the team defaults you ship.
3. `%APPDATA%\FPA Automation\settings.json`: one person's own changes, made on the app's **Settings** page.

| Setting | Meaning |
| --- | --- |
| `workbooks_root` | Folder holding the `OPEX`, `SAC`, `Revenue` and `Direct Cost` workbook folders. |
| `validation_output_root` | Where actuals validation reports go. Blank = find the synced `Finance_Validation_Results` OneDrive folder. |
| `forecast_validation_output_root` | Where forecast validation reports go. Blank = the synced OneDrive shortcut. |
| `budget_source_folder`, `budget_output_folder` | Defaults for the Initial Budget Setup page. |
| `test_mode` | Shows the TEST MODE badge. |
| `log_level` | `INFO` normally; `DEBUG` when chasing a problem. |

Folder changes take effect after restarting the app.

## When something goes wrong

Every run writes to `%APPDATA%\FPA Automation\logs\fpa_automation.log`. The Settings page has an **Open log folder** button. Ask colleagues to send you that file along with a description of what they clicked.

## Before the first release

- [ ] Run every automation end to end on **your** PC with a month whose results you already know, and compare the reports.
- [ ] Repeat on a **colleague's** PC, using only the built .exe. This catches missing access, a different Tableau version or different drive mappings.
- [ ] Try a failure: disconnect the shared drive, pick a wrong file. Check that the message is understandable and the log has the details.
- [ ] Open each report (OPEX, Direct Cost, Revenue) in Excel and print-preview it.
