# Runtime hook for PyInstaller / Windows: tell Playwright where to find browsers
import os, sys, pathlib
base = pathlib.Path(getattr(sys, "_MEIPASS", pathlib.Path(".")))
cand = base / "ms-playwright"
if cand.exists():
    os.environ["PLAYWRIGHT_BROWSERS_PATH"] = str(cand)

