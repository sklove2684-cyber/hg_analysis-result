import sys

from honyu_app.startup_timing import mark

mark("process 시작")

from PySide6.QtCore import QTimer
from PySide6.QtWidgets import QApplication

from honyu_app.application.shared_folder import SharedFolderController
from honyu_app.config.paths import local_app_data_dir
from honyu_app.config.dependencies import build_services
from honyu_app.config.settings import load_settings
from honyu_app.infrastructure.filesystem.shared_folder_service import (
    WindowsSharedFolderService,
)
from honyu_app.infrastructure.storage.local_settings import LocalSettingsStore
from honyu_app.infrastructure.pdf.labsolutions_parser import LabSolutionsParser
from honyu_app.ui.main_window import MainWindow

mark("Python/앱 import 완료")


def main() -> int:
    app = QApplication(sys.argv)
    settings = load_settings()
    mark("설정 로드 완료")
    mark("DB 초기화 시작")
    services = build_services(settings)
    connection = services.database.check_connection()
    if not connection.connected:
        raise RuntimeError(connection.message)
    mark("DB 초기화 완료")
    shared_folder_service = WindowsSharedFolderService(
        settings.unc_base_path, settings.z_fallback_path, settings.local_export_path
    )
    settings_store = LocalSettingsStore(local_app_data_dir() / "settings.json")
    controller = SharedFolderController(shared_folder_service, settings_store)
    mark("MainWindow 생성 시작")
    window = MainWindow(controller, LabSolutionsParser(), services.database)
    mark("MainWindow 생성 완료")
    window.show()
    app.processEvents()
    mark("GUI visible")
    QTimer.singleShot(0, window.start_background_initialization)
    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())
