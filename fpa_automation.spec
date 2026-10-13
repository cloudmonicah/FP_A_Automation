# PyInstaller build file. Build with:  build_exe.bat   (or: pyinstaller fpa_automation.spec --noconfirm)
# Output: dist\FPA Automation\FPA Automation.exe  (a folder build: faster start-up
# and fewer antivirus false alarms than a single-file .exe)

from PyInstaller.utils.hooks import collect_data_files, collect_submodules

hiddenimports = (
    collect_submodules("automations")
    + collect_submodules("shared")
    + collect_submodules("ui")
    + collect_submodules("pywinauto")
    + collect_submodules("comtypes")
)

a = Analysis(
    ["main.py"],
    pathex=["."],
    datas=collect_data_files("ttkbootstrap") + [("settings.example.json", ".")],
    hiddenimports=hiddenimports,
    excludes=["pytest"],
    noarchive=False,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name="FPA Automation",
    console=False,  # no black console window; output goes to the log file
    upx=False,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    name="FPA Automation",
    upx=False,
)
