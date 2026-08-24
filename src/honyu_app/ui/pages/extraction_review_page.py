from __future__ import annotations

from pathlib import Path
import socket
from uuid import UUID

from PySide6.QtCore import Qt, Signal
from PySide6.QtGui import QColor
from PySide6.QtWidgets import (
    QComboBox,
    QFileDialog,
    QFrame,
    QHBoxLayout,
    QHeaderView,
    QInputDialog,
    QLabel,
    QMessageBox,
    QPushButton,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from honyu_app.application.review_extraction import ReviewExtractionService
from honyu_app.domain.enums import ExcludeReason, ReviewStatus, SampleType
from honyu_app.domain.models import AnalysisBatch, Peak, Sample
from honyu_app.ui.theme import Card, field_label, set_status_tone


COLUMNS = (
    "페이지", "Sample", "구분", "반복", "Peak", "RT", "원본 Area",
    "수정 Area", "적용 Area", "원문 물질", "표준 물질", "DIBK", "Excel", "제외 사유",
)


def _stat_box(label: str) -> tuple[QFrame, QLabel]:
    frame = QFrame()
    frame.setProperty("uiCard", True)
    layout = QVBoxLayout(frame)
    layout.setContentsMargins(16, 12, 16, 12)
    layout.setSpacing(1)
    value = QLabel("0")
    value.setProperty("statValue", True)
    caption = QLabel(label)
    caption.setProperty("statLabel", True)
    layout.addWidget(value)
    layout.addWidget(caption)
    return frame, value


class ExtractionReviewPage(QWidget):
    batch_saved = Signal(object)
    excel_requested = Signal(object)

    def __init__(self, service: ReviewExtractionService) -> None:
        super().__init__()
        self.setObjectName("pageBody")
        self._service = service
        self._batch: AnalysisBatch | None = None
        self._saved_batch_id: UUID | None = None
        self._row_peaks: list[tuple[Sample, Peak]] = []
        self._latest_areas: dict[UUID, int] = {}

        layout = QVBoxLayout(self)
        layout.setContentsMargins(28, 20, 28, 24)
        layout.setSpacing(12)

        summary = QHBoxLayout()
        summary.setSpacing(10)
        boxes = (
            _stat_box("Sample"),
            _stat_box("전체 Peak"),
            _stat_box("Excel 입력"),
            _stat_box("제외 Peak"),
        )
        self.sample_count = boxes[0][1]
        self.peak_count = boxes[1][1]
        self.include_count = boxes[2][1]
        self.exclude_count = boxes[3][1]
        for frame, _ in boxes:
            summary.addWidget(frame, 1)
        layout.addLayout(summary)

        toolbar = Card("검토 작업", "행을 선택해 물질명, 포함 여부 또는 Area를 수정하세요.")
        saved_row = QHBoxLayout()
        saved_row.setSpacing(8)
        self.saved_batches = QComboBox()
        self.saved_batches.setMinimumWidth(420)
        refresh_saved = QPushButton("목록 새로고침")
        refresh_saved.clicked.connect(self.refresh_saved_batches)
        open_saved = QPushButton("저장된 배치 열기")
        open_saved.setProperty("kind", "primary")
        open_saved.clicked.connect(self.open_saved_batch)
        saved_row.addWidget(field_label("저장된 DB 배치"))
        saved_row.addWidget(self.saved_batches, 1)
        saved_row.addWidget(refresh_saved)
        saved_row.addWidget(open_saved)
        toolbar.body.addLayout(saved_row)
        self.status = QLabel("PDF 추출 결과를 기다리고 있습니다.")
        self.status.setProperty("statusTone", "neutral")
        self.status.setWordWrap(True)
        toolbar.body.addWidget(self.status)
        toolbar_row = QHBoxLayout()
        toolbar_row.setSpacing(8)
        self.filter = QComboBox()
        self.filter.addItems(
            ["전체", "입력 대상", "제외", "STD", "회수율", "Blank", "미등록 물질", "DIBK"]
        )
        self.filter.currentTextChanged.connect(self.apply_filter)
        toolbar_row.addWidget(field_label("표시 필터"))
        toolbar_row.addWidget(self.filter)
        toolbar_row.addSpacing(10)
        edit_actions = (
            ("material_button", "물질명 수정", self.edit_material),
            ("toggle_button", "포함/제외 전환", self.toggle_peak),
            ("area_button", "Area 수정", self.correct_area),
            ("history_button", "수정 이력", self.show_area_history),
        )
        for attribute, text, slot in edit_actions:
            button = QPushButton(text)
            button.clicked.connect(slot)
            button.setEnabled(False)
            setattr(self, attribute, button)
            toolbar_row.addWidget(button)
        toolbar_row.addStretch(1)
        toolbar.body.addLayout(toolbar_row)
        layout.addWidget(toolbar)

        self.table = QTableWidget(0, len(COLUMNS))
        self.table.setHorizontalHeaderLabels(COLUMNS)
        self.table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.table.setSelectionMode(QTableWidget.SelectionMode.SingleSelection)
        self.table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self.table.setAlternatingRowColors(True)
        self.table.setSortingEnabled(False)
        self.table.verticalHeader().setVisible(False)
        self.table.verticalHeader().setDefaultSectionSize(34)
        self.table.horizontalHeader().setStretchLastSection(True)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.ResizeToContents)
        layout.addWidget(self.table, 1)

        actions = QHBoxLayout()
        export = QPushButton("CSV 내보내기")
        export.clicked.connect(self.save_csv)
        self.complete_button = QPushButton("검토 완료")
        self.complete_button.setEnabled(False)
        self.complete_button.clicked.connect(self.complete_review)
        self.save_button = QPushButton("DB에 저장")
        self.save_button.setProperty("kind", "primary")
        self.save_button.setEnabled(False)
        self.save_button.clicked.connect(self.save_batch)
        self.excel_button = QPushButton("Excel 생성으로 이동")
        self.excel_button.setProperty("kind", "primary")
        self.excel_button.setEnabled(False)
        self.excel_button.clicked.connect(self.open_excel_export)
        actions.addWidget(export)
        actions.addStretch(1)
        actions.addWidget(self.complete_button)
        actions.addWidget(self.save_button)
        actions.addWidget(self.excel_button)
        layout.addLayout(actions)
        self.refresh_saved_batches()

    def refresh_saved_batches(self) -> None:
        selected = self._saved_batch_id
        self.saved_batches.clear()
        self.saved_batches.addItem("저장된 배치를 선택하세요", None)
        for summary in self._service.list_saved_batches():
            self.saved_batches.addItem(
                f"{summary.batch_code}  ·  {summary.pdf_filename}  ·  "
                f"{summary.analysis_no_start}-{summary.analysis_no_end}",
                summary.batch_id,
            )
        self.saved_batches.setCurrentIndex(0)
        if selected is not None:
            index = self.saved_batches.findData(selected)
            if index >= 0:
                self.saved_batches.setCurrentIndex(index)

    def open_saved_batch(self) -> None:
        batch_id = self.saved_batches.currentData()
        if not isinstance(batch_id, UUID):
            QMessageBox.warning(self, "저장 배치 없음", "먼저 DB에 저장된 배치가 없습니다.")
            return
        try:
            self._saved_batch_id = batch_id
            self.load_batch(self._service.load_saved_batch(batch_id))
        except Exception as exc:
            QMessageBox.critical(self, "배치 열기 실패", str(exc))

    def load_batch(self, batch: AnalysisBatch) -> None:
        self._batch = batch
        if batch.review_status is not ReviewStatus.SAVED:
            self._saved_batch_id = None
            self.saved_batches.setCurrentIndex(0)
        self._latest_areas.clear()
        if batch.review_status is ReviewStatus.SAVED:
            for sample in batch.samples:
                for peak in sample.peaks:
                    history = self._service.list_area_corrections(peak.peak_id)
                    if history:
                        self._latest_areas[peak.peak_id] = history[-1].area_after
        self.refresh_table()
        saved = batch.review_status is ReviewStatus.SAVED
        prefix = "기존 DB 배치" if saved else "신규 추출 결과"
        self.status.setText(
            f"{prefix}  ·  {batch.source_file.original_name}  ·  "
            f"Sample {len(batch.samples)}개  ·  검토 경고 {batch.warning_count}개"
            + ("  ·  DB 저장 완료  ·  Excel 생성 가능" if saved else "")
        )
        set_status_tone(self.status, "success" if saved or not batch.warning_count else "warning")
        self._sync_actions()

    def reset_for_new_work(self, *_args) -> None:
        self._batch = None
        self._saved_batch_id = None
        self._latest_areas.clear()
        self._row_peaks.clear()
        self.saved_batches.setCurrentIndex(0)
        self.table.setRowCount(0)
        self.filter.setCurrentIndex(0)
        self.status.setText("새 PDF가 선택되었습니다. 새 추출 결과를 기다리고 있습니다.")
        set_status_tone(self.status, "neutral")
        self._update_summary()
        self._sync_actions()

    def _sync_actions(self) -> None:
        status = self._batch.review_status if self._batch else None
        saved = status is ReviewStatus.SAVED
        reviewed = status is ReviewStatus.REVIEWED
        has_batch = self._batch is not None
        self.complete_button.setEnabled(has_batch and not saved and not reviewed)
        self.save_button.setEnabled(reviewed)
        self.excel_button.setEnabled(saved)
        self.complete_button.setText("검토 완료됨" if saved or reviewed else "검토 완료")
        self.save_button.setText("DB 저장됨" if saved else "DB에 저장")
        self.material_button.setEnabled(has_batch and not saved)
        self.toggle_button.setEnabled(has_batch and not saved)
        self.area_button.setEnabled(saved)
        self.history_button.setEnabled(saved)

    def open_excel_export(self) -> None:
        if self._batch is None or self._batch.review_status is not ReviewStatus.SAVED:
            QMessageBox.warning(self, "DB 저장 필요", "DB에 저장된 배치만 Excel 생성에 사용할 수 있습니다.")
            return
        self.excel_requested.emit(self._batch)

    def refresh_table(self) -> None:
        self.table.setRowCount(0)
        self._row_peaks.clear()
        if not self._batch:
            self._update_summary()
            return
        for sample in self._batch.samples:
            for peak in sample.peaks:
                row = self.table.rowCount()
                self.table.insertRow(row)
                self._row_peaks.append((sample, peak))
                corrected = self._latest_areas.get(peak.peak_id)
                values = (
                    sample.page_no, sample.sample_name_raw, sample.sample_type.value,
                    sample.replicate_no or "", peak.peak_no, str(peak.retention_time),
                    peak.area_raw, corrected if corrected is not None else "",
                    corrected if corrected is not None else peak.area_raw,
                    peak.material_raw or "", peak.material_standard or "",
                    peak.peak_group_no or "", "예" if peak.include_for_excel else "아니오",
                    peak.exclude_reason.value if peak.exclude_reason else "",
                )
                for column, value in enumerate(values):
                    item = QTableWidgetItem(str(value))
                    if column in {0, 3, 4, 5, 6, 7, 8, 11}:
                        item.setTextAlignment(
                            Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter
                        )
                    if peak.exclude_reason is ExcludeReason.UNKNOWN_MATERIAL:
                        item.setBackground(QColor("#ffeded"))
                    elif not peak.include_for_excel:
                        item.setForeground(QColor("#7a8495"))
                    elif corrected is not None and column in {7, 8}:
                        item.setBackground(QColor("#fff6df"))
                    self.table.setItem(row, column, item)
        self._update_summary()
        self.apply_filter()

    def _update_summary(self) -> None:
        samples = len(self._batch.samples) if self._batch else 0
        peaks = len(self._row_peaks)
        included = sum(peak.include_for_excel for _, peak in self._row_peaks)
        self.sample_count.setText(f"{samples:,}")
        self.peak_count.setText(f"{peaks:,}")
        self.include_count.setText(f"{included:,}")
        self.exclude_count.setText(f"{peaks - included:,}")

    def apply_filter(self) -> None:
        selected = self.filter.currentText()
        for row, (sample, peak) in enumerate(self._row_peaks):
            visible = (
                selected == "전체"
                or (selected == "입력 대상" and peak.include_for_excel)
                or (selected == "제외" and not peak.include_for_excel)
                or (selected == "STD" and sample.sample_type is SampleType.STD)
                or (selected == "회수율" and sample.sample_type is SampleType.RECOVERY)
                or (selected == "Blank" and sample.is_blank)
                or (selected == "미등록 물질" and peak.exclude_reason is ExcludeReason.UNKNOWN_MATERIAL)
                or (selected == "DIBK" and peak.material_standard == "DIBK")
            )
            self.table.setRowHidden(row, not visible)

    def _selected(self) -> tuple[Sample, Peak] | None:
        row = self.table.currentRow()
        if row < 0 or row >= len(self._row_peaks):
            QMessageBox.warning(self, "선택 필요", "Peak 행을 먼저 선택하세요.")
            return None
        return self._row_peaks[row]

    def edit_material(self) -> None:
        selected = self._selected()
        if not selected or not self._batch:
            return
        _, peak = selected
        value, accepted = QInputDialog.getItem(
            self, "표준 물질명 선택", "표준 물질명", self._service.standard_names, 0, False
        )
        if accepted:
            try:
                self._service.set_material_mapping(self._batch, peak.peak_id, value)
                self.refresh_table()
            except Exception as exc:
                QMessageBox.critical(self, "매핑 오류", str(exc))

    def toggle_peak(self) -> None:
        selected = self._selected()
        if not selected or not self._batch:
            return
        _, peak = selected
        try:
            self._service.set_peak_included(
                self._batch, peak.peak_id, not peak.include_for_excel
            )
            self.refresh_table()
        except Exception as exc:
            QMessageBox.warning(self, "전환 불가", str(exc))

    def complete_review(self) -> None:
        if not self._batch:
            return
        try:
            self._service.complete_review(self._batch)
            self.status.setText("검토가 완료되었습니다. DB에 저장할 수 있습니다.")
            set_status_tone(self.status, "success")
            self._sync_actions()
        except Exception as exc:
            QMessageBox.warning(self, "검토 미완료", str(exc))

    def save_batch(self) -> None:
        if not self._batch:
            return
        try:
            result = self._service.save_batch(self._batch)
            self._saved_batch_id = self._batch.batch_id
            self.batch_saved.emit(self._batch)
            self.status.setText(f"DB 저장 완료  ·  {result.batch_code}")
            set_status_tone(self.status, "success")
            self._sync_actions()
            self.refresh_saved_batches()
        except Exception as exc:
            QMessageBox.critical(self, "DB 저장 실패", str(exc))

    def correct_area(self) -> None:
        selected = self._selected()
        if not selected or not self._batch:
            return
        if self._batch.review_status is not ReviewStatus.SAVED:
            QMessageBox.warning(self, "DB 저장 필요", "DB 저장 후 Area를 수정할 수 있습니다.")
            return
        _, peak = selected
        current = self._latest_areas.get(peak.peak_id, peak.area_raw)
        area, accepted = QInputDialog.getInt(
            self, "Area 수정", f"원본 Area: {peak.area_raw}\n수정 Area", current, 0, 2_147_483_647
        )
        if not accepted:
            return
        reason, accepted = QInputDialog.getText(self, "수정 사유", "수정 사유")
        if not accepted:
            return
        try:
            correction = self._service.add_area_correction(
                peak.peak_id, area, reason, socket.gethostname()
            )
            self._latest_areas[peak.peak_id] = correction.area_after
            self.refresh_table()
            self.status.setText(
                f"Area 수정 완료  ·  revision {correction.revision_no}  ·  {correction.device_id}"
            )
            set_status_tone(self.status, "success")
        except Exception as exc:
            QMessageBox.critical(self, "Area 수정 실패", str(exc))

    def save_csv(self) -> None:
        if not self._batch:
            return
        selected, _ = QFileDialog.getSaveFileName(
            self, "추출 결과 CSV 저장", f"{self._batch.batch_code}.csv", "CSV (*.csv)"
        )
        if not selected:
            return
        try:
            self._service.export_csv(self._batch, Path(selected))
            self.status.setText(f"CSV 저장 완료  ·  {selected}")
            set_status_tone(self.status, "success")
        except Exception as exc:
            QMessageBox.critical(self, "CSV 저장 실패", str(exc))

    def show_area_history(self) -> None:
        selected = self._selected()
        if not selected:
            return
        _, peak = selected
        try:
            history = self._service.list_area_corrections(peak.peak_id)
        except Exception as exc:
            QMessageBox.critical(self, "이력 조회 실패", str(exc))
            return
        if not history:
            QMessageBox.information(
                self, "Area 수정 이력", f"원본 Area: {peak.area_raw}\n수정 이력이 없습니다."
            )
            return
        lines = [f"원본 Area: {peak.area_raw}"]
        for value in history:
            lines.append(
                f"revision {value.revision_no}: {value.area_before} → {value.area_after}\n"
                f"사유: {value.reason}\n일시: {value.corrected_at.isoformat()}\n"
                f"PC: {value.device_id}"
            )
        QMessageBox.information(self, "Area 수정 이력", "\n\n".join(lines))
