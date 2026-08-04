from pathlib import Path
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

