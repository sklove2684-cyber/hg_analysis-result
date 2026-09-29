from __future__ import annotations

import os
import sys
from pathlib import Path


def _configure_frozen_dll_search_path() -> tuple[object, ...]:
    if not getattr(sys, "frozen", False) or not hasattr(os, "add_dll_directory"):
        return ()
    bundle_dir = Path(getattr(sys, "_MEIPASS", Path(sys.executable).parent))
    return tuple(
        os.add_dll_directory(str(directory))
        for directory in (
            bundle_dir,
            bundle_dir / "PySide6",
            bundle_dir / "shiboken6",
        )
        if directory.is_dir()
    )


_DLL_DIRECTORY_HANDLES = _configure_frozen_dll_search_path()

from PySide6.QtCore import QTimer
from PySide6.QtWidgets import QApplication, QMessageBox


def main() -> int:
    app = QApplication(sys.argv)
    box = QMessageBox(QMessageBox.Icon.Information, "Qt smoke test", "QtCore / QtWidgets 정상")
    QTimer.singleShot(1500, box.accept)
    box.show()
    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())
