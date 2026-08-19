from __future__ import annotations

from datetime import datetime
from pathlib import Path
import platform
from uuid import UUID

from PySide6.QtCore import Qt
from PySide6.QtGui import QColor
from PySide6.QtWidgets import (
    QComboBox,
    QFileDialog,
    QFrame,
    QGridLayout,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPushButton,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from honyu_app.application.create_excel_export import CreateExcelExportService
from honyu_app.application.preview_excel_export import PreviewExcelExportService
from honyu_app.domain.enums import ExcelPreviewStatus, StdMethod
from honyu_app.domain.models import AnalysisBatch, ExcelPreviewResult
from honyu_app.domain.queries import BatchSearchQuery
from honyu_app.services.database_service import DatabaseService
from honyu_app.ui.theme import Card, field_label, set_status_tone


COLUMNS = (
    "Sample", "구분", "물질", "Peak", "RT", "원본 Area", "적용 Area",
    "DIBK 순위", "시트", "셀", "기존 유형", "수식", "상태", "제외/오류 사유",
)


def _stat_box(label: str) -> tuple[QFrame, QLabel]:
    frame = QFrame()
    frame.setProperty("uiCard", True)
    layout = QVBoxLayout(frame)
    layout.setContentsMargins(16, 10, 16, 10)
    layout.setSpacing(0)
    value = QLabel("0")
    value.setProperty("statValue", True)
    caption = QLabel(label)
    caption.setProperty("statLabel", True)
    layout.addWidget(value)
    layout.addWidget(caption)
    return frame, value


class ExcelExportPage(QWidget):
    def __init__(
        self,
        database: DatabaseService,
        preview_service: PreviewExcelExportService,
        create_service: CreateExcelExportService,
    ) -> None:
        super().__init__()
        self.setObjectName("pageBody")
        self._database = database
        self._preview_service = preview_service
        self._create_service = create_service
        self._result: ExcelPreviewResult | None = None

        layout = QVBoxLayout(self)
        layout.setContentsMargins(28, 20, 28, 24)
        layout.setSpacing(12)

        setup = Card("Excel 생성 설정", "DB 배치와 Excel 양식을 선택하고 결과 파일 위치를 지정하세요.")
        grid = QGridLayout()
        grid.setHorizontalSpacing(10)
        grid.setVerticalSpacing(8)
        self.batch_combo = QComboBox()
        refresh = QPushButton("새로고침")
        refresh.clicked.connect(self.refresh_batches)
        self.std_method = QComboBox()
        self.std_method.addItem("방식 A  ·  STD1~5", StdMethod.A.value)
        self.std_method.addItem("방식 B  ·  STD1~4 + STD6", StdMethod.B.value)
        grid.addWidget(field_label("DB 분석 배치"), 0, 0)
        grid.addWidget(field_label("STD 방식"), 0, 2)
        grid.addWidget(self.batch_combo, 1, 0)
        grid.addWidget(refresh, 1, 1)
        grid.addWidget(self.std_method, 1, 2)

        self.template_path = QLineEdit()
        self.template_path.setReadOnly(True)
        self.template_path.setPlaceholderText("Excel 양식 파일을 선택하세요")
        template_button = QPushButton("원본 선택")
        template_button.clicked.connect(self.choose_template)
        grid.addWidget(field_label("Excel 양식"), 2, 0, 1, 3)
        grid.addWidget(self.template_path, 3, 0, 1, 2)
        grid.addWidget(template_button, 3, 2)

        self.output_path = QLineEdit()
        self.output_path.setReadOnly(True)
        self.output_path.setPlaceholderText("결과 Excel 저장 위치를 선택하세요")
        output_button = QPushButton("저장 위치")
        output_button.clicked.connect(self.choose_output)
        grid.addWidget(field_label("결과 Excel"), 4, 0, 1, 3)
        grid.addWidget(self.output_path, 5, 0, 1, 2)
        grid.addWidget(output_button, 5, 2)
        grid.setColumnStretch(0, 1)
        grid.setColumnStretch(2, 0)
        setup.body.addLayout(grid)

        controls = QHBoxLayout()
        self.status = QLabel("배치와 Excel 양식을 선택하세요.")
        self.status.setProperty("statusTone", "neutral")
        self.status.setWordWrap(True)
        self.preview_button = QPushButton("입력 미리보기")
        self.preview_button.clicked.connect(self.preview)
        self.create_button = QPushButton("검증 후 Excel 생성")
        self.create_button.setProperty("kind", "primary")
        self.create_button.setEnabled(False)
        self.create_button.clicked.connect(self.create_excel)
        controls.addWidget(self.status, 1)
        controls.addWidget(self.preview_button)
        controls.addWidget(self.create_button)
        setup.body.addLayout(controls)
        layout.addWidget(setup)

        summary = QHBoxLayout()
        summary.setSpacing(10)
        boxes = (_stat_box("Excel 입력"), _stat_box("입력 제외"), _stat_box("검증 오류"))
        self.mapped_count = boxes[0][1]
        self.excluded_count = boxes[1][1]
        self.error_count = boxes[2][1]
        for frame, _ in boxes:
            summary.addWidget(frame, 1)
        layout.addLayout(summary)

        self.table = QTableWidget(0, len(COLUMNS))
        self.table.setHorizontalHeaderLabels(COLUMNS)
        self.table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self.table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.table.setSelectionMode(QTableWidget.SelectionMode.SingleSelection)
        self.table.setAlternatingRowColors(True)
        self.table.verticalHeader().setVisible(False)
        self.table.verticalHeader().setDefaultSectionSize(34)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.ResizeToContents)
        self.table.horizontalHeader().setStretchLastSection(True)
        layout.addWidget(self.table, 1)

        self.issue_label = QLabel("미리보기를 실행하면 입력 제외 및 오류 사유가 표시됩니다.")
        self.issue_label.setProperty("statusTone", "neutral")
        self.issue_label.setWordWrap(True)
        layout.addWidget(self.issue_label)
        self.batch_combo.currentIndexChanged.connect(self._preview_input_changed)
        self.std_method.currentIndexChanged.connect(self._preview_input_changed)
        self.output_path.textChanged.connect(self._update_create_button)
        self.refresh_batches()

    def refresh_batches(self) -> None:
        selected = self.batch_combo.currentData()
        self.batch_combo.clear()
        for summary in self._database.search_batches(BatchSearchQuery()):
            label = (
                f"{summary.batch_code}  ·  {summary.pdf_filename}  ·  "
                f"{summary.analysis_no_start}-{summary.analysis_no_end}"
            )
            self.batch_combo.addItem(label, summary.batch_id)
        if selected is not None:
            index = self.batch_combo.findData(selected)
            if index >= 0:
                self.batch_combo.setCurrentIndex(index)

    def load_batch(self, batch: AnalysisBatch) -> None:
        self.refresh_batches()
        index = self.batch_combo.findData(batch.batch_id)
        if index >= 0:
            self.batch_combo.setCurrentIndex(index)

    def choose_template(self) -> None:
        selected, _ = QFileDialog.getOpenFileName(
            self, "Excel 양식 선택", self.template_path.text(), "Excel (*.xlsx)"
        )
        if selected:
            self._result = None
            self.template_path.setText(selected)
            template = Path(selected)
            if not self.output_path.text().strip():
                self.output_path.setText(str(template.with_name(f"{template.stem}_결과.xlsx")))
            self.create_button.setEnabled(False)
            self.status.setText("Excel 양식을 선택했습니다. 입력 미리보기를 실행하세요.")
            set_status_tone(self.status, "neutral")

    def choose_output(self) -> None:
        initial = self.output_path.text().strip() or self.template_path.text().strip()
        selected, _ = QFileDialog.getSaveFileName(
            self, "결과 Excel 저장", initial, "Excel (*.xlsx)"
        )
        if selected:
            if not selected.lower().endswith(".xlsx"):
                selected += ".xlsx"
            self.output_path.setText(selected)

    def _preview_input_changed(self, *_args) -> None:
        if self._result is None:
            return
        self._result = None
        self.create_button.setEnabled(False)
        self.status.setText("선택 내용이 변경되었습니다. 입력 미리보기를 다시 실행하세요.")
        set_status_tone(self.status, "neutral")

    def _update_create_button(self, *_args) -> None:
        output = Path(self.output_path.text().strip())
        self.create_button.setEnabled(
            bool(
                self._result
                and self._result.can_generate
                and output.parent.is_dir()
                and not output.exists()
            )
        )

    def _ensure_available_output_path(self) -> Path | None:
        raw = self.output_path.text().strip()
        if not raw:
            return None
        output = Path(raw)
        if not output.exists():
            return None
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        candidate = output.with_name(f"{output.stem}_{stamp}{output.suffix}")
        counter = 2
        while candidate.exists():
            candidate = output.with_name(
                f"{output.stem}_{stamp}_{counter}{output.suffix}"
            )
            counter += 1
        self.output_path.setText(str(candidate))
        return candidate

    def preview(self) -> None:
        batch_id = self.batch_combo.currentData()
        template = Path(self.template_path.text().strip())
        if not isinstance(batch_id, UUID):
            QMessageBox.warning(self, "배치 선택 필요", "DB 분석 배치를 선택하세요.")
            return
        if not template.is_file():
            QMessageBox.warning(self, "Excel 양식 필요", "Excel 양식 파일을 선택하세요.")
            return
        self._result = None
        self.create_button.setEnabled(False)
        self.status.setText("Excel 입력 위치를 검증하고 있습니다...")
        set_status_tone(self.status, "neutral")
        try:
            self._result = self._preview_service.preview(
                batch_id, template, self.std_method.currentData()
            )
        except Exception as exc:
            self.status.setText("입력 미리보기에 실패했습니다. 설정을 확인한 뒤 다시 시도하세요.")
            set_status_tone(self.status, "error")
            QMessageBox.critical(self, "미리보기 실패", str(exc))
            return
        renamed_output = self._ensure_available_output_path()
        self._show_result(self._result)
        if renamed_output is not None and not self._result.issues:
            self.issue_label.setText(
                "검증 오류가 없습니다. 기존 결과 파일을 보존하기 위해 "
                f"새 파일명으로 저장합니다:\n{renamed_output.name}"
            )
            set_status_tone(self.issue_label, "success")

    def create_excel(self) -> None:
        batch_id = self.batch_combo.currentData()
        template = Path(self.template_path.text().strip())
        output = Path(self.output_path.text().strip())
        if not isinstance(batch_id, UUID) or self._result is None or not self._result.can_generate:
            QMessageBox.warning(self, "미리보기 필요", "오류 없는 입력 미리보기를 먼저 실행하세요.")
            return
        if not output.parent.is_dir() or output.exists():
            QMessageBox.warning(
                self, "저장 위치 확인", "존재하는 폴더 안에 아직 없는 결과 파일명을 선택하세요."
            )
            return
        answer = QMessageBox.question(
            self,
            "Excel 생성 확인",
            f"원본은 변경하지 않고 아래 결과 파일을 생성합니다.\n\n{output}\n\n계속할까요?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
        )
        if answer != QMessageBox.StandardButton.Yes:
            return
        self.create_button.setEnabled(False)
        self.status.setText("숫자 입력 · 구조 검증 · Excel 전체 재계산 중...")
        set_status_tone(self.status, "neutral")
        try:
            result = self._create_service.create(
                batch_id, template, output, self.std_method.currentData(), platform.node()
            )
        except Exception as exc:
            self.status.setText(
                "Excel 생성에 실패했습니다. 미리보기는 유지되며 저장 위치를 바꿔 다시 시도할 수 있습니다."
            )
            set_status_tone(self.status, "error")
            QMessageBox.critical(self, "Excel 생성 실패", str(exc))
            self._update_create_button()
            return
        if result.recalculated:
            self.status.setText(
                f"생성 완료  ·  입력 {result.mapped_cell_count}개  ·  재계산 및 구조검증 완료"
            )
            message = f"결과 파일:\n{result.output_path}"
        else:
            self.status.setText(
                f"생성 완료  ·  입력 {result.mapped_cell_count}개  ·  수식 재계산 대기"
            )
            message = (
                f"결과 파일:\n{result.output_path}\n\n"
                "Office 미인증으로 수식 재계산은 생략했습니다. "
                "입력값·수식·서식은 보존되어 있으며, 인증된 Excel에서 파일을 열면 "
                "전체 수식이 자동 재계산됩니다."
            )
        set_status_tone(self.status, "success")
        QMessageBox.information(self, "Excel 생성 완료", message)

    def _show_result(self, result: ExcelPreviewResult) -> None:
        self.table.setRowCount(0)
        for preview in result.rows:
            row = self.table.rowCount()
            self.table.insertRow(row)
            values = (
                preview.sample_name, preview.sample_type.value, preview.material or "",
                preview.peak_no, str(preview.retention_time), preview.area_raw,
                preview.applied_area, preview.dibk_area_rank or "",
                preview.target_sheet or "", preview.target_cell or "",
                preview.existing_value_type or "",
                "예" if preview.existing_has_formula else "아니요",
                preview.status.value, preview.exclude_reason or preview.message or "",
            )
            for column, value in enumerate(values):
                item = QTableWidgetItem(str(value))
                if column in {3, 4, 5, 6, 7}:
                    item.setTextAlignment(
                        Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter
                    )
                if preview.status is ExcelPreviewStatus.ERROR:
                    item.setBackground(QColor("#ffeded"))
                elif preview.status is ExcelPreviewStatus.EXCLUDED:
                    item.setForeground(QColor("#7a8495"))
                self.table.setItem(row, column, item)

        self.mapped_count.setText(f"{result.mapped_count:,}")
        self.excluded_count.setText(f"{result.excluded_count:,}")
        self.error_count.setText(f"{result.error_count:,}")
        state = "생성 가능" if result.can_generate else "오류 수정 필요"
        self.status.setText(
            f"{state}  ·  입력 {result.mapped_count}개  ·  제외 {result.excluded_count}개  ·  "
            f"오류 {result.error_count}개"
        )
        set_status_tone(self.status, "success" if result.can_generate else "error")
        self._update_create_button()
        if result.issues:
            self.issue_label.setText(
                "\n".join(
                    f"[{issue.severity.value}] {issue.code}  ·  {issue.message}"
                    for issue in result.issues[:8]
                )
            )
            set_status_tone(self.issue_label, "error" if result.error_count else "warning")
        else:
            self.issue_label.setText("검증 오류가 없습니다. Excel을 생성할 수 있습니다.")
            set_status_tone(self.issue_label, "success")
