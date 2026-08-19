from __future__ import annotations

import base64
from pathlib import Path
import shutil
import subprocess
from uuid import uuid4
from xml.etree import ElementTree as ET

from honyu_app.domain.errors import ExcelRecalculationError


class ExcelComRecalculator:
    def __init__(self, timeout_seconds: int = 180) -> None:
        self._timeout_seconds = timeout_seconds

    def recalculate(self, workbook_path: Path) -> None:
        workbook_path = Path(workbook_path).resolve()
        recalculated_path = workbook_path.with_name(
            f"honyu_recalculated_{uuid4().hex}.xlsx"
        )
        shutil.copy2(workbook_path, recalculated_path)
        path = str(recalculated_path).replace("'", "''")
        calculation_timeout = max(30, self._timeout_seconds - 30)
        script = rf"""
$ErrorActionPreference = 'Stop'
if (Get-Process -Name EXCEL -ErrorAction SilentlyContinue) {{
    throw 'EXCEL_ALREADY_RUNNING'
}}
$excelPath = (Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\excel.exe').'(default)'
if (-not $excelPath) {{ throw 'EXCEL_NOT_FOUND' }}
$process = Start-Process -FilePath $excelPath -ArgumentList '/x', ('"' + '{path}' + '"') -PassThru
$excel = $null
$book = $null
try {{
    $watch = [System.Diagnostics.Stopwatch]::StartNew()
    while ($excel -eq $null) {{
        $process.Refresh()
        if ($process.MainWindowTitle -match '제품 인증 실패|Product Activation Failed') {{
            throw 'OFFICE_NOT_ACTIVATED'
        }}
        try {{ $excel = [Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application') }} catch {{}}
        if ($watch.Elapsed.TotalSeconds -gt 30) {{ throw 'EXCEL_CONNECTION_TIMEOUT' }}
        Start-Sleep -Milliseconds 300
    }}
    $excel.DisplayAlerts = $false
    $excel.AskToUpdateLinks = $false
    $excel.EnableEvents = $false
    while ($book -eq $null) {{
        try {{ $book = $excel.ActiveWorkbook }} catch {{}}
        if ($watch.Elapsed.TotalSeconds -gt 30) {{ throw 'WORKBOOK_OPEN_TIMEOUT' }}
        Start-Sleep -Milliseconds 300
    }}
    while (-not $excel.Ready) {{
        if ($watch.Elapsed.TotalSeconds -gt 60) {{ throw 'EXCEL_READY_TIMEOUT' }}
        Start-Sleep -Milliseconds 300
    }}
    $excel.CalculateFullRebuild()
    while ($excel.CalculationState -ne 0) {{
        if ($watch.Elapsed.TotalSeconds -gt {calculation_timeout}) {{ throw 'EXCEL_CALCULATION_TIMEOUT' }}
        Start-Sleep -Milliseconds 300
    }}
    $book.Save()
    $book.Close($false)
    $book = $null
    $excel.Quit()
    $excel = $null
}} finally {{
    if ($book -ne $null) {{ try {{ $book.Close($false) }} catch {{}} }}
    if ($excel -ne $null) {{ try {{ $excel.Quit() }} catch {{}} }}
    if ($process -ne $null -and -not $process.HasExited) {{
        Stop-Process -Id $process.Id -Force -ErrorAction SilentlyContinue
    }}
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}}
"""
        encoded = base64.b64encode(script.encode("utf-16le")).decode("ascii")
        try:
            completed = subprocess.run(
                [
                    "powershell.exe", "-NoProfile", "-NonInteractive",
                    "-ExecutionPolicy", "Bypass", "-EncodedCommand", encoded,
                ],
                capture_output=True,
                text=True,
                encoding="mbcs",
                errors="replace",
                timeout=self._timeout_seconds,
                check=False,
            )
        except subprocess.TimeoutExpired as exc:
            recalculated_path.unlink(missing_ok=True)
            raise ExcelRecalculationError(
                f"Excel 전체 재계산이 {self._timeout_seconds}초 안에 끝나지 않았습니다."
            ) from exc
        except OSError as exc:
            recalculated_path.unlink(missing_ok=True)
            raise ExcelRecalculationError(f"Excel 실행을 시작할 수 없습니다: {exc}") from exc
        if completed.returncode != 0:
            recalculated_path.unlink(missing_ok=True)
            detail = self._readable_error(
                completed.stderr or completed.stdout or "알 수 없는 COM 오류"
            )
            if "EXCEL_ALREADY_RUNNING" in detail:
                raise ExcelRecalculationError(
                    "열려 있는 Microsoft Excel 창을 모두 닫고 다시 시도하세요."
                )
            if "EXCEL_NOT_FOUND" in detail:
                raise ExcelRecalculationError(
                    "이 PC에서 Microsoft Excel 실행 파일을 찾을 수 없습니다."
                )
            if "OFFICE_NOT_ACTIVATED" in detail:
                raise ExcelRecalculationError(
                    "Microsoft Office 제품 인증이 필요합니다. "
                    "Office 정품 인증을 완료한 뒤 다시 시도하세요.",
                    code="OFFICE_NOT_ACTIVATED",
                )
            raise ExcelRecalculationError(f"Excel 전체 재계산에 실패했습니다: {detail}")
        if not recalculated_path.is_file():
            raise ExcelRecalculationError("Excel 재계산 복사본을 찾을 수 없습니다.")
        recalculated_path.replace(workbook_path)

    @staticmethod
    def _readable_error(raw: str) -> str:
        raw = raw.strip()
        marker = raw.find("<Objs")
        if marker >= 0:
            try:
                root = ET.fromstring(raw[marker:])
                messages = [
                    (item.text or "").replace("_x000D__x000A_", "\n").strip()
                    for item in root.iter()
                    if item.tag.endswith("}S") and item.attrib.get("S") == "Error"
                ]
                readable = " ".join(message for message in messages if message)
                if readable:
                    return " ".join(readable.split())
            except ET.ParseError:
                pass
        return " ".join(raw.replace("_x000D__x000A_", " ").split())
