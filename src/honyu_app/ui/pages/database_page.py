from __future__ import annotations

from uuid import UUID

from PySide6.QtCore import Qt, Signal
from PySide6.QtWidgets import (
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QMessageBox,
    QPushButton,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from honyu_app.domain.enums import ReviewStatus
from honyu_app.domain.queries import BatchSearchQuery
from honyu_app.services.database_service import DatabaseService
from honyu_app.ui.theme import Card, set_status_tone


COLUMNS = (
    "배치 코드", "PDF 파일", "분석 종류", "분석번호", "상태", "Area 수정", "Excel 생성",
)


class DatabasePage(QWidget):
    review_requested = Signal(object)
    excel_requested = Signal(object)

    def __init__(self, database: DatabaseService) -> None:
        super().__init__()
        self.setObjectName("pageBody")
        self._database = database
        self._batch_ids: list[UUID] = []

        layout = QVBoxLayout(self)
        layout.setContentsMargins(28, 20, 28, 24)
        layout.setSpacing(12)

        header = Card("저장된 분석 배치", "DB에 저장된 결과를 다시 열거나 Excel 생성 단계로 보냅니다.")
        actions = QHBoxLayout()
        self.status = QLabel("저장 배치를 불러오는 중입니다.")
        self.status.setProperty("statusTone", "neutral")
        refresh = QPushButton("목록 새로고침")
        refresh.clicked.connect(self.refresh_batches)
        review = QPushButton("검토 화면에서 열기")
        review.clicked.connect(self.open_review)
        excel = QPushButton("Excel 생성으로 이동")
        excel.setProperty("kind", "primary")
        excel.clicked.connect(self.open_excel)
        actions.addWidget(self.status, 1)
        actions.addWidget(refresh)
        actions.addWidget(review)
        actions.addWidget(excel)
        header.body.addLayout(actions)
        layout.addWidget(header)

        self.table = QTableWidget(0, len(COLUMNS))
        self.table.setHorizontalHeaderLabels(COLUMNS)
        self.table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self.table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.table.setSelectionMode(QTableWidget.SelectionMode.SingleSelection)
        self.table.setAlternatingRowColors(True)
        self.table.verticalHeader().setVisible(False)
        self.table.verticalHeader().setDefaultSectionSize(36)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.table.itemDoubleClicked.connect(lambda _: self.open_review())
        layout.addWidget(self.table, 1)
        self.refresh_batches()

    def refresh_batches(self) -> None:
        self.table.setRowCount(0)
        self._batch_ids.clear()
        try:
            summaries = self._database.search_batches(BatchSearchQuery())
        except Exception as exc:
            self.status.setText(f"DB 조회 실패: {exc}")
            set_status_tone(self.status, "error")
            return
        for summary in summaries:
            row = self.table.rowCount()
            self.table.insertRow(row)
            self._batch_ids.append(summary.batch_id)
            values = (
                summary.batch_code,
                summary.pdf_filename,
                summary.analysis_type,
                f"{summary.analysis_no_start}-{summary.analysis_no_end}",
                "저장 완료" if summary.review_status == ReviewStatus.SAVED.value else summary.review_status,
                summary.correction_count,
                summary.export_count,
            )
            for column, value in enumerate(values):
                item = QTableWidgetItem(str(value))
                if column in {3, 5, 6}:
                    item.setTextAlignment(
                        Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter
                    )
                self.table.setItem(row, column, item)
        if summaries:
            self.table.selectRow(0)
            self.status.setText(f"저장된 배치 {len(summaries)}건")
            set_status_tone(self.status, "success")
        else:
            self.status.setText("DB에 저장된 분석 배치가 없습니다.")
            set_status_tone(self.status, "neutral")

    def _selected_batch(self):
        row = self.table.currentRow()
        if row < 0 or row >= len(self._batch_ids):
            QMessageBox.warning(self, "배치 선택 필요", "저장된 분석 배치를 선택하세요.")
            return None
        try:
            batch = self._database.get_batch_detail(self._batch_ids[row])
            batch.review_status = ReviewStatus.SAVED
            return batch
        except Exception as exc:
            QMessageBox.critical(self, "배치 열기 실패", str(exc))
            return None

    def open_review(self) -> None:
        batch = self._selected_batch()
        if batch is not None:
            self.review_requested.emit(batch)

    def open_excel(self) -> None:
        batch = self._selected_batch()
        if batch is not None:
            self.excel_requested.emit(batch)
