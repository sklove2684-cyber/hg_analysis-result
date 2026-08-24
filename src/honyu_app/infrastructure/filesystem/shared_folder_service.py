from __future__ import annotations

import os
from pathlib import Path
import re

from honyu_app.domain.enums import HalfYear
from honyu_app.domain.errors import SharedFolderUnavailableError
from honyu_app.domain.results import FolderValidationResult, SharedFolderConnectionStatus
from honyu_app.config.paths import default_local_export_dir


def _natural_key(value: str) -> list[tuple[int, object]]:
    parts = re.split(r"(\d+)", value.casefold())
    return [(0, int(part)) if part.isdigit() else (1, part) for part in parts]


class WindowsSharedFolderService:
    """Read-only discovery of the company shared-folder hierarchy."""

    def __init__(
        self,
        unc_base_path: str,
        z_fallback_path: str,
        local_export_path: str | None = None,
    ) -> None:
        self._unc_base_path = Path(unc_base_path)
        self._z_fallback_path = Path(z_fallback_path)
        self._local_export_path = Path(local_export_path or default_local_export_dir())
        self._active_base_path: Path | None = None

    @property
    def active_base_path(self) -> Path | None:
        return self._active_base_path

    def check_connection(self) -> SharedFolderConnectionStatus:
        attempted: list[str] = []
        for used_fallback, candidate in (
            (False, self._unc_base_path),
            (True, self._z_fallback_path),
        ):
            attempted.append(str(candidate))
            try:
                if candidate.is_dir():
                    self._active_base_path = candidate
                    return SharedFolderConnectionStatus(
                        connected=True,
                        active_base_path=str(candidate),
                        used_fallback=used_fallback,
                        attempted_paths=tuple(attempted),
                        message=(
                            "회사 공유폴더 저장 · Z: 보조 경로"
                            if used_fallback
                            else "회사 공유폴더 저장 · UNC 연결됨"
                        ),
                        storage_mode="company",
                    )
            except OSError:
                continue
        try:
            self._local_export_path.mkdir(parents=True, exist_ok=True)
        except OSError as exc:
            self._active_base_path = None
            return SharedFolderConnectionStatus(
                connected=False,
                active_base_path=None,
                used_fallback=False,
                attempted_paths=tuple(attempted),
                message=f"로컬 저장 폴더를 준비할 수 없습니다: {exc}",
                storage_mode="unavailable",
            )
        self._active_base_path = self._local_export_path
        return SharedFolderConnectionStatus(
            connected=False,
            active_base_path=str(self._local_export_path),
            used_fallback=False,
            attempted_paths=tuple(attempted),
            message="로컬 저장 모드 · 회사 공유폴더에 연결할 수 없어 로컬에 저장합니다.",
            storage_mode="local",
        )

    def _require_connection(self) -> Path:
        if self._active_base_path is None:
            status = self.check_connection()
            if not status.connected:
                raise SharedFolderUnavailableError(
                    f"공유폴더 연결 실패: {', '.join(status.attempted_paths)}"
                )
        assert self._active_base_path is not None
        return self._active_base_path

    def list_workplaces(self) -> list[str]:
        base = self._require_connection()
        try:
            names = [entry.name for entry in os.scandir(base) if entry.is_dir() and not entry.name.startswith(".")]
        except OSError as exc:
            self._active_base_path = None
            raise SharedFolderUnavailableError(f"작업장 목록을 읽을 수 없습니다: {base}") from exc
        return sorted(names, key=_natural_key)

    def build_period_path(self, workplace: str, year: int, half: HalfYear) -> Path:
        base = self._require_connection()
        self._validate_workplace_name(workplace)
        self._validate_year(year)
        return base / workplace / f"{year % 100:02d}{half.folder_suffix}"

    def validate_period_path(
        self, workplace: str, year: int, half: HalfYear
    ) -> FolderValidationResult:
        path = self.build_period_path(workplace, year, half)
        if not path.is_dir():
            return FolderValidationResult(
                valid=False,
                path=str(path),
                message="기간 폴더가 없습니다. 자동으로 생성하지 않습니다.",
            )
        return FolderValidationResult(True, str(path), "기간 폴더를 확인했습니다.")

    def validate_final_folder(
        self, folder: Path, *, workplace: str, year: int, half: HalfYear
    ) -> FolderValidationResult:
        period = self.build_period_path(workplace, year, half)
        if not folder.is_dir():
            return FolderValidationResult(False, str(folder), "선택한 최종 폴더가 없습니다.")
        try:
            folder.resolve().relative_to(period.resolve())
        except ValueError:
            return FolderValidationResult(
                False,
                str(folder),
                f"최종 폴더는 기간 폴더 아래에서 선택해야 합니다: {period}",
            )
        return FolderValidationResult(True, str(folder), "최종 저장 폴더를 확인했습니다.")

    @staticmethod
    def _validate_workplace_name(workplace: str) -> None:
        if not workplace.strip() or workplace in {".", ".."}:
            raise ValueError("작업장을 선택해야 합니다.")
        if any(char in workplace for char in "\\/:"):
            raise ValueError("작업장 이름에 경로 구분자를 사용할 수 없습니다.")

    @staticmethod
    def _validate_year(year: int) -> None:
        if not 2000 <= year <= 2099:
            raise ValueError("연도는 2000~2099 범위여야 합니다.")
