from __future__ import annotations

import os
import sys
from pathlib import Path


PROJECT_DIR = Path(__file__).resolve().parent
sys.path.insert(0, str(PROJECT_DIR / "src"))


def _configure_frozen_dll_search_path() -> tuple[object, ...]:
    """Keep bundled Qt/shiboken directories ahead of system PATH on Windows."""
    if not getattr(sys, "frozen", False) or not hasattr(os, "add_dll_directory"):
        return ()
    bundle_dir = Path(getattr(sys, "_MEIPASS", PROJECT_DIR))
    candidates = (
        bundle_dir,
        bundle_dir / "PySide6",
        bundle_dir / "shiboken6",
    )
    return tuple(
        os.add_dll_directory(str(directory))
        for directory in candidates
        if directory.is_dir()
    )


_DLL_DIRECTORY_HANDLES = _configure_frozen_dll_search_path()

from honyu_app.main import main


if __name__ == "__main__":
    raise SystemExit(main())
