# PyInstaller recipe for the desktop app. Build from the project folder:
#     pyinstaller desktop/automation_tool.spec --noconfirm
# Output: dist/Automation Tool/ (Windows: run "Automation Tool.exe") or dist/Automation Tool.app (Mac).
# Build on each system you want it for (a Windows build needs Windows); GitHub Actions does both.

import sys
from pathlib import Path
from PyInstaller.utils.hooks import collect_all, copy_metadata

ROOT = Path(SPECPATH).parent
NAME = "Automation Tool"

datas = [(str(ROOT / f), ".") for f in
         ("app.py", "turnover.py", "inputs.py", "run_queue.py", "default_keywords.txt", "sample_companies.xlsx")]
binaries, hiddenimports = [], ["turnover", "inputs", "run_queue"]
for pkg in ("streamlit", "rapidocr_onnxruntime", "onnxruntime", "pypdfium2", "pypdfium2_raw", "rapidfuzz"):
    d, b, h = collect_all(pkg)
    datas += d
    binaries += b
    hiddenimports += h
for pkg in ("streamlit", "pandas", "numpy", "pyarrow", "openpyxl", "altair"):
    datas += copy_metadata(pkg)

a = Analysis([str(ROOT / "desktop" / "launcher.py")], pathex=[str(ROOT)], binaries=binaries, datas=datas,
             hiddenimports=hiddenimports, excludes=["pytest", "IPython", "matplotlib"])
pyz = PYZ(a.pure)
exe = EXE(pyz, a.scripts, [], exclude_binaries=True, name=NAME, console=False,
          icon=str(ROOT / "desktop" / ("icon.ico" if sys.platform == "win32" else "icon.icns"))
          if (ROOT / "desktop" / ("icon.ico" if sys.platform == "win32" else "icon.icns")).exists() else None)
coll = COLLECT(exe, a.binaries, a.datas, name=NAME)
if sys.platform == "darwin":
    app = BUNDLE(coll, name=f"{NAME}.app", bundle_identifier="uk.automationtool.app",
                 info_plist={"NSHighResolutionCapable": True})
