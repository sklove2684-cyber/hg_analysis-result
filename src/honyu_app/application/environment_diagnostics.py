from __future__ import annotations

from dataclasses import dataclass
from hashlib import sha256
from importlib import metadata
import json
import os
from pathlib import Path
import platform
import subprocess
import sys
from typing import Callable

from honyu_app.application.preview_excel_export import (
    IPA_AREA_PROFILE,
    IPA_PROFILE,
    PreviewExcelExportService,
)
from honyu_app.config.analysis_types import infer_analysis_type
from honyu_app.domain.models import ExcelPreviewResult
from honyu_app.domain.queries import BatchSearchQuery
from honyu_app.infrastructure.excel.workbook_inspector import XlsxTemplateInspector
from honyu_app.services.database_service import DatabaseService


FAILURE = "확인 실패"
NOT_SELECTED = "선택되지 않음"


@dataclass(frozen=True, slots=True)
class EnvironmentDiagnosticReport:
    text: str
    local_sha: str
    origin_sha: str
    git_matches: bool | None
    worktree_clean: bool | None
    pdf_path: str
    pdf_sha256: str
    excel_path: str
    excel_sha256: str
    excel_profile_key: str
    excel_profile_name: str


def file_sha256(path: Path, *, chunk_size: int = 1024 * 1024) -> str:
    digest = sha256()
    with Path(path).open("rb") as source:
        while chunk := source.read(chunk_size):
            digest.update(chunk)
    return digest.hexdigest().upper()


def _safe_value(action: Callable[[], str], failure: str = FAILURE) -> str:
    try:
        return action()
    except Exception:
        return failure


class EnvironmentDiagnosticService:
    def __init__(
        self,
        database: DatabaseService,
        *,
        project_directory: Path | None = None,
        inspector: XlsxTemplateInspector | None = None,
        git_timeout_seconds: float = 3.0,
    ) -> None:
        self._database = database
        self._project_directory = (
            Path(project_directory)
            if project_directory is not None
            else Path(__file__).resolve().parents[3]
        )
        self._inspector = inspector or XlsxTemplateInspector()
        self._git_timeout_seconds = git_timeout_seconds

    def _git(self, *arguments: str) -> str:
        completed = subprocess.run(
            ["git", *arguments],
            cwd=self._project_directory,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=self._git_timeout_seconds,
            check=False,
            shell=False,
            creationflags=(
                subprocess.CREATE_NO_WINDOW
                if os.name == "nt" and hasattr(subprocess, "CREATE_NO_WINDOW")
                else 0
            ),
        )
        if completed.returncode != 0:
            raise RuntimeError(completed.stderr.strip() or "git command failed")
        return completed.stdout.strip()

    @staticmethod
    def _package_version() -> str:
        return _safe_value(lambda: metadata.version("honyu-automation"))

    @staticmethod
    def _editable_install() -> str:
        def inspect_install() -> str:
            distribution = metadata.distribution("honyu-automation")
            payload = distribution.read_text("direct_url.json")
            if not payload:
                return "아니오"
            value = json.loads(payload)
            return "예" if value.get("dir_info", {}).get("editable") else "아니오"

        return _safe_value(inspect_install)

    def _database_lines(self) -> list[str]:
        path_value = getattr(self._database, "database_file", None)
        path = Path(path_value).resolve() if path_value is not None else None
        connection = None
        try:
            connection = self._database.check_connection()
        except Exception:
            pass
        kind = getattr(connection, "mode", None) or type(self._database).__name__
        lines = [
            f"DB 경로: {path if path is not None else FAILURE}",
            f"DB 종류: {kind}",
        ]
        if path is None:
            lines.extend((f"DB 크기: {FAILURE}", f"DB SHA-256: {FAILURE}"))
        else:
            lines.append(f"DB 크기: {_safe_value(lambda: f'{path.stat().st_size} bytes')}")
            lines.append(f"DB SHA-256: {_safe_value(lambda: file_sha256(path), '접근 실패')}")
        try:
            latest = self._database.search_batches(BatchSearchQuery())[:1]
            if latest:
                batch = latest[0]
                lines.extend(
                    (
                        f"최신 batch ID: {batch.batch_id}",
                        f"최신 batch code: {batch.batch_code}",
                        f"최신 analysis_type: {batch.analysis_type}",
                        f"최신 source PDF: {batch.pdf_filename}",
                    )
                )
            else:
                lines.append("최신 batch: 없음")
        except Exception:
            lines.append(f"최신 batch: {FAILURE}")
        return lines

    @staticmethod
    def _file_lines(label: str, path_text: str) -> tuple[list[str], str]:
        if not path_text.strip():
            return [f"{label}: {NOT_SELECTED}"], NOT_SELECTED
        path = Path(path_text)
        lines = [f"전체 경로: {path_text}", f"파일명: {path.name}"]
        lines.append(f"파일 크기: {_safe_value(lambda: f'{path.stat().st_size} bytes', '접근 실패')}")
        digest = _safe_value(lambda: file_sha256(path), "접근 실패")
        lines.append(f"SHA-256: {digest}")
        return lines, digest

    def _pdf_lines(self, path_text: str) -> tuple[list[str], str]:
        lines, digest = self._file_lines("현재 선택 PDF", path_text)
        if path_text.strip():
            detected = _safe_value(
                lambda: infer_analysis_type(Path(path_text).name) or "판별되지 않음"
            )
            lines.append(f"자동 판별 분석종류: {detected}")
        return lines, digest

    @staticmethod
    def _profile_display(profile) -> str:
        if profile is IPA_PROFILE:
            return "IPA - LOD(area입력)형"
        if profile is IPA_AREA_PROFILE:
            return "IPA - area형"
        return profile.name

    @staticmethod
    def _evidence_addresses(area_sheet: str) -> tuple[str, ...]:
        if area_sheet == "결과입력(area입력)":
            return ("F2",)
        if area_sheet == "LOD(area입력)":
            return ("D2", "E2", "F3", "I3", "J3")
        if area_sheet == "area입력":
            return ("F3", "I3", "L3", "O3", "R3")
        if area_sheet == "area":
            return ("F3", "I3", "J3", "L3", "O3", "R3")
        return ()

    def _excel_lines(self, path_text: str) -> tuple[list[str], str, str, str]:
        lines, digest = self._file_lines("현재 선택 Excel", path_text)
        if not path_text.strip():
            return lines, digest, NOT_SELECTED, NOT_SELECTED
        try:
            snapshot = self._inspector.inspect(Path(path_text))
            preview = ExcelPreviewResult(Path(path_text), "A")
            profile = PreviewExcelExportService._template_profile(snapshot, preview)
            lines.append(f"실제 시트 목록: {', '.join(snapshot.sheet_names)}")
            if profile is None:
                profile_key = FAILURE
                profile_name = FAILURE
                detail = preview.issues[0].message if preview.issues else FAILURE
                lines.append(f"Excel profile: {detail}")
            else:
                profile_key = profile.key
                profile_name = self._profile_display(profile)
                lines.append(f"[판별] {profile_key} / {profile_name}")
                for address in self._evidence_addresses(profile.area_sheet):
                    cell = snapshot.cell(profile.area_sheet, address)
                    if cell.value not in (None, ""):
                        lines.append(f"핵심 셀: {profile.area_sheet}!{address} = {cell.value}")
        except Exception:
            lines.extend(
                (
                    "실제 시트 목록: 접근 실패",
                    "Excel profile key: 확인 실패",
                    "Excel 양식 이름: 확인 실패",
                )
            )
            profile_key = FAILURE
            profile_name = FAILURE
        return lines, digest, profile_key, profile_name

    def collect(
        self, pdf_path: str = "", excel_path: str = ""
    ) -> EnvironmentDiagnosticReport:
        branch = _safe_value(lambda: self._git("branch", "--show-current"))
        local_sha = _safe_value(lambda: self._git("rev-parse", "HEAD"))
        title = _safe_value(lambda: self._git("log", "-1", "--format=%s"))
        origin_sha = _safe_value(
            lambda: self._git("rev-parse", "refs/remotes/origin/main")
        )
        dirty_text = _safe_value(lambda: self._git("status", "--porcelain"))
        git_matches = (
            local_sha == origin_sha
            if FAILURE not in (local_sha, origin_sha)
            else None
        )
        worktree_clean = dirty_text == "" if dirty_text != FAILURE else None
        git_status = (
            "[정상] Local == origin/main"
            if git_matches is True
            else "[주의] Local != origin/main"
            if git_matches is False
            else "[주의] Git 비교 확인 실패"
        )

        pdf_lines, pdf_digest = self._pdf_lines(pdf_path)
        excel_lines, excel_digest, profile_key, profile_name = self._excel_lines(
            excel_path
        )
        sections = [
            "[소스/버전]",
            f"Git branch: {branch}",
            f"Git commit: {local_sha}",
            f"Commit 제목: {title}",
            f"origin/main SHA: {origin_sha}",
            git_status,
            f"작업 트리: {'clean' if worktree_clean is True else 'dirty' if worktree_clean is False else FAILURE}",
            f"프로그램 패키지 버전: {self._package_version()}",
            "",
            "[실행환경]",
            f"프로젝트/실행 기준 경로: {self._project_directory.resolve()}",
            f"실제 실행 entry 경로: {_safe_value(lambda: str(Path(sys.argv[0]).resolve()))}",
            f"Python executable: {sys.executable}",
            f"Python 버전: {platform.python_version()}",
            f"PID: {os.getpid()}",
            f"Editable 설치: {self._editable_install()}",
            "",
            "[DB]",
            *self._database_lines(),
            "",
            "[현재 선택 PDF]",
            *pdf_lines,
            "",
            "[현재 선택 Excel]",
            *excel_lines,
        ]
        return EnvironmentDiagnosticReport(
            text="\n".join(sections),
            local_sha=local_sha,
            origin_sha=origin_sha,
            git_matches=git_matches,
            worktree_clean=worktree_clean,
            pdf_path=pdf_path or NOT_SELECTED,
            pdf_sha256=pdf_digest,
            excel_path=excel_path or NOT_SELECTED,
            excel_sha256=excel_digest,
            excel_profile_key=profile_key,
            excel_profile_name=profile_name,
        )
