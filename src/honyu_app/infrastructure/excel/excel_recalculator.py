from __future__ import annotations

import base64
from pathlib import Path
import subprocess
from xml.etree import ElementTree as ET

from honyu_app.domain.errors import ExcelRecalculationError


class ExcelComRecalculator:
    def __init__(self, timeout_seconds: int = 180) -> None:
        self._timeout_seconds = timeout_seconds

    def recalculate(self, workbook_path: Path) -> None:
        path = str(Path(workbook_path).resolve()).replace("'", "''")
        calculation_timeout = max(30, self._timeout_seconds - 30)
        script = f"""
$ErrorActionPreference = 'Stop'
$excel = $null
$book = $null
try {{
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AskToUpdateLinks = $false
    $book = $excel.Workbooks.Open('{path}', 0, $false)
    $excel.CalculateFullRebuild()
    $watch = [System.Diagnostics.Stopwatch]::StartNew()
    while ($excel.CalculationState -ne 0) {{
        if ($watch.Elapsed.TotalSeconds -gt {calculation_timeout}) {{
            throw 'Excel 전체 재계산 시간이 초과되었습니다.'
        }}
        Start-Sleep -Milliseconds 200
    }}
    $book.Save()
    $book.Close($true)
    $book = $null
    $excel.Quit()
    $excel = $null
}} finally {{
    if ($book -ne $null) {{
        try {{ $book.Close($false) }} catch {{}}
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($book)
    }}
    if ($excel -ne $null) {{
        try {{ $excel.Quit() }} catch {{}}
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)
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
                timeout=self._timeout_seconds,
                check=False,
            )
        except subprocess.TimeoutExpired as exc:
            raise ExcelRecalculationError(
                f"Excel 전체 재계산이 {self._timeout_seconds}초 안에 끝나지 않았습니다."
            ) from exc
        except OSError as exc:
            raise ExcelRecalculationError(f"Excel 실행을 시작할 수 없습니다: {exc}") from exc
        if completed.returncode != 0:
            detail = self._readable_error(
                completed.stderr or completed.stdout or "알 수 없는 COM 오류"
            )
            raise ExcelRecalculationError(f"Excel 전체 재계산에 실패했습니다: {detail}")

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
