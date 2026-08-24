from __future__ import annotations

import base64
from pathlib import Path
import shutil
import subprocess
from uuid import uuid4
from xml.etree import ElementTree as ET
from zipfile import BadZipFile, ZipFile

from honyu_app.domain.errors import ExcelRecalculationError


class ExcelComRecalculator:
    def __init__(self, timeout_seconds: int = 180) -> None:
        self._timeout_seconds = timeout_seconds

    def recalculate(self, workbook_path: Path) -> None:
        workbook_path = Path(workbook_path).resolve()
        recalculated_path = workbook_path.with_name(
            f"honyu_recalculated_{uuid4().hex}.xlsx"
        )
        owner_pid_path = recalculated_path.with_suffix(".excel.pid")
        diagnostic_path = workbook_path.with_suffix(
            workbook_path.suffix + ".recalculation.log"
        )
        shutil.copy2(workbook_path, recalculated_path)
        path = str(recalculated_path).replace("'", "''")
        pid_path = str(owner_pid_path).replace("'", "''")
        log_path = str(diagnostic_path).replace("'", "''")
        script = rf"""
$ErrorActionPreference = 'Stop'
function Write-Diagnostic([string]$eventName, [string]$details) {{
    $line = ('{{0:o}} event={{1}} {{2}}' -f [DateTime]::Now, $eventName, $details)
    Add-Content -LiteralPath '{log_path}' -Value $line -Encoding UTF8
}}
function Write-WorkbookState([string]$eventName) {{
    $fullName = '<not-open>'
    $readOnly = '<unknown>'
    $saved = '<unknown>'
    $linkCount = 0
    $connectionCount = 0
    $queryCount = 0
    $circularReference = '<none>'
    if ($book -ne $null) {{
        try {{ $fullName = $book.FullName }} catch {{ $fullName = '<error>' }}
        try {{ $readOnly = $book.ReadOnly }} catch {{}}
        try {{ $saved = $book.Saved }} catch {{}}
        try {{
            $links = @($book.LinkSources(1))
            if ($links.Count -eq 1 -and $null -eq $links[0]) {{ $linkCount = 0 }}
            else {{ $linkCount = $links.Count }}
        }} catch {{ $linkCount = -1 }}
        try {{ $connectionCount = $book.Connections.Count }} catch {{ $connectionCount = -1 }}
        try {{
            foreach ($sheet in @($book.Worksheets)) {{
                $queryCount += $sheet.QueryTables.Count
                foreach ($table in @($sheet.ListObjects)) {{
                    if ($table.SourceType -ne 1) {{ $queryCount += 1 }}
                }}
            }}
        }} catch {{ $queryCount = -1 }}
        try {{
            if ($excel.CircularReference -ne $null) {{
                $circularReference = $excel.CircularReference.Address($true, $true, 1, $true)
            }}
        }} catch {{ $circularReference = '<error>' }}
    }}
    $detail = 'excelPid={{0}} workbook="{{1}}" calculation={{2}} calculationState={{3}} ready={{4}} readOnly={{5}} saved={{6}} externalLinks={{7}} connections={{8}} queries={{9}} circularReference="{{10}}" calculateBeforeSave={{11}} iteration={{12}} enableEvents={{13}} displayAlerts={{14}} screenUpdating={{15}}' -f $ownedExcelPid, $fullName, $excel.Calculation, $excel.CalculationState, $excel.Ready, $readOnly, $saved, $linkCount, $connectionCount, $queryCount, $circularReference, $excel.CalculateBeforeSave, $excel.Iteration, $excel.EnableEvents, $excel.DisplayAlerts, $excel.ScreenUpdating
    Write-Diagnostic $eventName $detail
}}
function Wait-CalculationDone([string]$methodName, [double]$stageSeconds) {{
    $stageWatch = [System.Diagnostics.Stopwatch]::StartNew()
    while ($excel.CalculationState -ne 0) {{
        if ($calculationWatch.Elapsed.TotalSeconds -gt {self._timeout_seconds}) {{
            Write-WorkbookState ('calculation_total_timeout method=' + $methodName)
            throw 'EXCEL_CALCULATION_TIMEOUT'
        }}
        if ($stageWatch.Elapsed.TotalSeconds -ge $stageSeconds) {{
            Write-WorkbookState ('calculation_stage_pending method=' + $methodName)
            return $false
        }}
        Start-Sleep -Milliseconds 200
    }}
    Write-WorkbookState ('calculation_method_done method=' + $methodName)
    return $true
}}
$existingExcelPids = @(
    Get-Process -Name EXCEL -ErrorAction SilentlyContinue |
        ForEach-Object {{ $_.Id }}
)
Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
public static class ExcelWindowProcess {{
    [DllImport("user32.dll")]
    public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);
}}
'@
$excel = $null
$book = $null
$bootstrapBook = $null
$ownsExcel = $false
$ownedExcelPid = 0
try {{
    $excelPath = (Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\excel.exe' -ErrorAction SilentlyContinue).'(default)'
    if (-not $excelPath) {{ throw 'EXCEL_NOT_FOUND' }}
    try {{
        $excel = New-Object -ComObject Excel.Application
    }} catch {{
        throw ('EXCEL_COM_CREATE_FAILED: ' + $_.Exception.Message)
    }}

    [uint32]$excelProcessId = 0
    [ExcelWindowProcess]::GetWindowThreadProcessId(
        [IntPtr]$excel.Hwnd, [ref]$excelProcessId
    ) | Out-Null
    if ($excelProcessId -eq 0 -or $existingExcelPids -contains [int]$excelProcessId) {{
        throw 'EXCEL_DEDICATED_INSTANCE_NOT_CREATED'
    }}
    $ownedExcelPid = [int]$excelProcessId
    $ownsExcel = $true
    Set-Content -LiteralPath '{pid_path}' -Value $ownedExcelPid -Encoding ascii
    Write-Diagnostic 'dedicated_instance_created' ('excelPid={{0}} hwnd={{1}}' -f $ownedExcelPid, $excel.Hwnd)
    $excelProcess = Get-Process -Id $ownedExcelPid -ErrorAction Stop
    if ($excelProcess.MainWindowTitle -match '제품 인증 실패|Product Activation Failed') {{
        throw 'OFFICE_NOT_ACTIVATED'
    }}

    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AskToUpdateLinks = $false
    $excel.EnableEvents = $false
    $excel.ScreenUpdating = $false
    $excel.AutomationSecurity = 3
    $bootstrapBook = $excel.Workbooks.Add()
    $excel.Calculation = -4135

    $book = $excel.Workbooks.Open('{path}', 0, $false)
    if ($book -eq $null) {{ throw 'WORKBOOK_OPEN_FAILED' }}
    # Keep the bootstrap workbook open until the target opens. Otherwise Excel
    # adopts the target workbook's saved automatic mode and starts calculating
    # before the program can select the intended calculation scope.
    $excel.Calculation = -4135
    $bootstrapBook.Close($false)
    [Runtime.InteropServices.Marshal]::FinalReleaseComObject($bootstrapBook) | Out-Null
    $bootstrapBook = $null
    if ($excel.Workbooks.Count -ne 1) {{
        throw ('EXCEL_UNEXPECTED_WORKBOOK_COUNT: ' + $excel.Workbooks.Count)
    }}
    Write-WorkbookState 'workbook_opened'

    # The XML writer changes input values, not formulas. Rebuilding Excel's entire
    # dependency tree is unnecessary and can stall on legacy/broken references.
    $book.ForceFullCalculation = $false
    $originalCalculationMode = $excel.Calculation
    Write-Diagnostic 'calculation_mode_before' ('excelPid={{0}} calculation={{1}}' -f $ownedExcelPid, $originalCalculationMode)
    $calculationWatch = [System.Diagnostics.Stopwatch]::StartNew()
    $excel.Calculation = -4105
    Write-WorkbookState 'automatic_calculation_enabled'
    Write-Diagnostic 'calculation_method_selected' 'method=excel.Calculate()'
    Write-WorkbookState 'calculation_before method=excel.Calculate()'
    $excel.Calculate()
    Write-WorkbookState 'calculation_after method=excel.Calculate()'
    $calculationDone = Wait-CalculationDone 'excel.Calculate()' 8

    if (-not $calculationDone) {{
        Write-Diagnostic 'calculation_method_selected' 'method=Worksheet.Calculate() order=reverse formulaSheetsOnly'
        for ($sheetIndex = $book.Worksheets.Count; $sheetIndex -ge 1; $sheetIndex--) {{
            $worksheet = $null
            $formulaCells = $null
            try {{
                $worksheet = $book.Worksheets.Item($sheetIndex)
                try {{ $formulaCells = $worksheet.UsedRange.SpecialCells(-4123) }} catch {{}}
                if ($formulaCells -eq $null) {{
                    Write-Diagnostic 'calculation_method_skipped' ('method=Worksheet.Calculate() sheet="{{0}}" reason=noFormulas' -f $worksheet.Name)
                    continue
                }}
                Write-WorkbookState ('calculation_before method=Worksheet.Calculate() sheet=' + $worksheet.Name)
                $worksheet.Calculate()
                Write-WorkbookState ('calculation_after method=Worksheet.Calculate() sheet=' + $worksheet.Name)
            }} catch {{
                Write-Diagnostic 'calculation_method_error' ('method=Worksheet.Calculate() sheetIndex={{0}} error="{{1}}"' -f $sheetIndex, $_.Exception.Message)
            }} finally {{
                if ($formulaCells -ne $null) {{ try {{ [Runtime.InteropServices.Marshal]::FinalReleaseComObject($formulaCells) | Out-Null }} catch {{}} }}
                if ($worksheet -ne $null) {{ try {{ [Runtime.InteropServices.Marshal]::FinalReleaseComObject($worksheet) | Out-Null }} catch {{}} }}
            }}
        }}
        $calculationDone = Wait-CalculationDone 'Worksheet.Calculate()' 8
    }}

    if (-not $calculationDone) {{
        Write-Diagnostic 'calculation_method_selected' 'method=Range.Calculate() order=reverse formulaRangesOnly'
        for ($sheetIndex = $book.Worksheets.Count; $sheetIndex -ge 1; $sheetIndex--) {{
            $worksheet = $null
            $formulaCells = $null
            try {{
                $worksheet = $book.Worksheets.Item($sheetIndex)
                try {{ $formulaCells = $worksheet.UsedRange.SpecialCells(-4123) }} catch {{}}
                if ($formulaCells -eq $null) {{ continue }}
                Write-WorkbookState ('calculation_before method=Range.Calculate() sheet=' + $worksheet.Name)
                $formulaCells.Calculate()
                Write-WorkbookState ('calculation_after method=Range.Calculate() sheet=' + $worksheet.Name)
            }} catch {{
                Write-Diagnostic 'calculation_method_error' ('method=Range.Calculate() sheetIndex={{0}} error="{{1}}"' -f $sheetIndex, $_.Exception.Message)
            }} finally {{
                if ($formulaCells -ne $null) {{ try {{ [Runtime.InteropServices.Marshal]::FinalReleaseComObject($formulaCells) | Out-Null }} catch {{}} }}
                if ($worksheet -ne $null) {{ try {{ [Runtime.InteropServices.Marshal]::FinalReleaseComObject($worksheet) | Out-Null }} catch {{}} }}
            }}
        }}
        $calculationDone = Wait-CalculationDone 'Range.Calculate()' 8
    }}

    if (-not $calculationDone) {{
        Write-Diagnostic 'calculation_method_selected' 'method=excel.CalculateFull()'
        Write-WorkbookState 'calculation_before method=excel.CalculateFull()'
        $excel.CalculateFull()
        Write-WorkbookState 'calculation_after method=excel.CalculateFull()'
        $remainingSeconds = {self._timeout_seconds} - $calculationWatch.Elapsed.TotalSeconds
        if ($remainingSeconds -le 0) {{ throw 'EXCEL_CALCULATION_TIMEOUT' }}
        $calculationDone = Wait-CalculationDone 'excel.CalculateFull()' $remainingSeconds
    }}

    if (-not $calculationDone -or $excel.CalculationState -ne 0) {{
        Write-WorkbookState 'workbook_calculation_timeout'
        throw 'EXCEL_CALCULATION_TIMEOUT'
    }}
    Write-WorkbookState 'workbook_calculation_completed'
    Write-Diagnostic 'calculation_mode_retained' ('excelPid={{0}} calculation={{1}} reason=dedicated_instance' -f $ownedExcelPid, $excel.Calculation)

    $book.Save()
    Write-WorkbookState 'workbook_saved'
    $book.Close($false)
    $book = $null
    $excel.Quit()
    $excel = $null
}} finally {{
    if ($bootstrapBook -ne $null) {{ try {{ $bootstrapBook.Close($false) }} catch {{}} }}
    if ($book -ne $null) {{ try {{ $book.Close($false) }} catch {{}} }}
    if ($excel -ne $null -and $ownsExcel) {{ try {{ $excel.Quit() }} catch {{}} }}
    if ($book -ne $null) {{
        try {{ [Runtime.InteropServices.Marshal]::FinalReleaseComObject($book) | Out-Null }} catch {{}}
    }}
    if ($bootstrapBook -ne $null) {{
        try {{ [Runtime.InteropServices.Marshal]::FinalReleaseComObject($bootstrapBook) | Out-Null }} catch {{}}
    }}
    if ($excel -ne $null) {{
        try {{ [Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) | Out-Null }} catch {{}}
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
                timeout=self._timeout_seconds + 45,
                check=False,
            )
        except subprocess.TimeoutExpired as exc:
            self._stop_owned_excel(owner_pid_path)
            recalculated_path.unlink(missing_ok=True)
            detail = self._diagnostic_detail(diagnostic_path)
            raise ExcelRecalculationError(
                f"Excel 통합문서 재계산이 {self._timeout_seconds}초 안에 끝나지 않았습니다."
                f"{detail}"
            ) from exc
        except OSError as exc:
            recalculated_path.unlink(missing_ok=True)
            raise ExcelRecalculationError(f"Excel 실행을 시작할 수 없습니다: {exc}") from exc
        finally:
            owner_pid_path.unlink(missing_ok=True)
        if completed.returncode != 0:
            recalculated_path.unlink(missing_ok=True)
            detail = self._readable_error(
                completed.stderr or completed.stdout or "알 수 없는 COM 오류"
            )
            diagnostic = self._diagnostic_detail(diagnostic_path)
            if "EXCEL_DEDICATED_INSTANCE_NOT_CREATED" in detail:
                raise ExcelRecalculationError(
                    "사용자 Excel과 분리된 전용 Excel 인스턴스를 만들지 못했습니다."
                )
            if "EXCEL_NOT_FOUND" in detail:
                raise ExcelRecalculationError(
                    "이 PC에서 Microsoft Excel 실행 파일을 찾을 수 없습니다."
                )
            if "EXCEL_COM_CREATE_FAILED" in detail:
                raise ExcelRecalculationError(
                    "Microsoft Excel 전용 COM 인스턴스를 만들 수 없습니다: "
                    f"{detail}"
                )
            if "OFFICE_NOT_ACTIVATED" in detail:
                raise ExcelRecalculationError(
                    "Microsoft Office 제품 인증이 필요합니다. "
                    "Office 정품 인증을 완료한 뒤 다시 시도하세요.",
                    code="OFFICE_NOT_ACTIVATED",
                )
            raise ExcelRecalculationError(
                f"Excel 통합문서 재계산에 실패했습니다: {detail}{diagnostic}"
            )
        if not recalculated_path.is_file():
            raise ExcelRecalculationError("Excel 재계산 복사본을 찾을 수 없습니다.")
        before_formula_cache = self._formula_cache_values(workbook_path)
        after_formula_cache = self._formula_cache_values(recalculated_path)
        changed_formula_count = sum(
            before_formula_cache.get(key) != value
            for key, value in after_formula_cache.items()
        )
        with diagnostic_path.open("a", encoding="utf-8") as log:
            log.write(
                "formula_cache_verification "
                f"formulaCount={len(after_formula_cache)} "
                f"changedFormulaResults={changed_formula_count}\n"
            )
        if before_formula_cache and changed_formula_count == 0:
            recalculated_path.unlink(missing_ok=True)
            raise ExcelRecalculationError(
                "Excel 계산은 완료됐지만 변경된 수식 결과를 확인하지 못했습니다."
                f"{self._diagnostic_detail(diagnostic_path)}"
            )
        recalculated_path.replace(workbook_path)
        diagnostic_path.unlink(missing_ok=True)

    @staticmethod
    def _formula_cache_values(workbook_path: Path) -> dict[str, tuple[str, str, str]]:
        values: dict[str, tuple[str, str, str]] = {}
        try:
            with ZipFile(workbook_path) as archive:
                sheet_names = sorted(
                    name for name in archive.namelist()
                    if name.startswith("xl/worksheets/sheet") and name.endswith(".xml")
                )
                namespace = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
                for sheet_name in sheet_names:
                    root = ET.fromstring(archive.read(sheet_name))
                    for cell in root.findall(".//m:c", namespace):
                        formula = cell.find("m:f", namespace)
                        if formula is None:
                            continue
                        cached = cell.find("m:v", namespace)
                        values[f"{sheet_name}!{cell.attrib.get('r', '')}"] = (
                            formula.text or "",
                            cached.text if cached is not None and cached.text is not None else "",
                            cell.attrib.get("t", ""),
                        )
        except (OSError, BadZipFile, ET.ParseError, KeyError):
            return {}
        return values

    @staticmethod
    def _diagnostic_detail(diagnostic_path: Path) -> str:
        try:
            lines = diagnostic_path.read_text(encoding="utf-8-sig").strip().splitlines()
        except OSError:
            return "\n진단 로그를 읽을 수 없습니다."
        if not lines:
            return "\n진단 로그가 비어 있습니다."
        return f"\n진단 로그: {diagnostic_path}\n" + "\n".join(lines[-8:])

    @staticmethod
    def _stop_owned_excel(owner_pid_path: Path) -> None:
        try:
            pid = int(owner_pid_path.read_text(encoding="ascii").strip())
        except (OSError, ValueError):
            return
        subprocess.run(
            [
                "powershell.exe",
                "-NoProfile",
                "-NonInteractive",
                "-Command",
                (
                    f"$p = Get-Process -Id {pid} -ErrorAction SilentlyContinue; "
                    "if ($p -and $p.ProcessName -eq 'EXCEL') { "
                    "Stop-Process -Id $p.Id -Force -ErrorAction SilentlyContinue }"
                ),
            ],
            capture_output=True,
            check=False,
        )

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
