import sys

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


def main() -> int:
    app = QApplication(sys.argv)
    settings = load_settings()
    services = build_services(settings)
    connection = services.database.check_connection()
    if not connection.connected:
        raise RuntimeError(connection.message)
    shared_folder_service = WindowsSharedFolderService(
        settings.unc_base_path, settings.z_fallback_path
    )
    settings_store = LocalSettingsStore(local_app_data_dir() / "settings.json")
    controller = SharedFolderController(shared_folder_service, settings_store)
    window = MainWindow(controller, LabSolutionsParser(), services.database)
    window.show()
    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())
