from pathlib import Path
import ctypes
import os


def local_app_data_dir() -> Path:
    base = os.getenv("LOCALAPPDATA")
    if not base:
        raise RuntimeError("LOCALAPPDATA 환경 변수를 찾을 수 없습니다.")
    return Path(base) / "HonyuAutomation"


def mock_db_dir() -> Path:
    return local_app_data_dir() / "mock_db"


def pending_dir() -> Path:
    return local_app_data_dir() / "pending"


def excel_work_dir() -> Path:
    return pending_dir() / "excel_work"


def failed_exports_dir() -> Path:
    return pending_dir() / "failed_exports"


def windows_desktop_dir() -> Path:
    if os.name == "nt":
        buffer = ctypes.create_unicode_buffer(260)
        # CSIDL_DESKTOPDIRECTORY follows the user's configured/redirected Desktop.
        result = ctypes.windll.shell32.SHGetFolderPathW(None, 0x10, None, 0, buffer)
        if result == 0 and buffer.value:
            return Path(buffer.value)
    return Path.home() / "Desktop"


def default_local_export_dir() -> Path:
    return windows_desktop_dir() / "분석프로그램"
