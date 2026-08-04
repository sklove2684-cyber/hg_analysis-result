from datetime import date
from pathlib import Path
from threading import Event

from PySide6.QtCore import QObject, QThread, Signal, Slot
from PySide6.QtWidgets import (
    QComboBox,
    QFileDialog,
    QGridLayout,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QProgressBar,
    QPushButton,
    QSpinBox,
    QVBoxLayout,
    QWidget,
)

from honyu_app.application.shared_folder import SharedFolderController
from honyu_app.domain.enums import HalfYear
from honyu_app.infrastructure.pdf.labsolutions_parser import LabSolutionsParser
from honyu_app.services.database_service import DatabaseService
from honyu_app.ui.theme import Card, field_label, make_path_label, set_status_tone


class PdfExtractionWorker(QObject):
    completed = Signal(object)
    failed = Signal(str)
    progress = Signal(int, int)
    finished = Signal()

    def __init__(
        self,
        parser: LabSolutionsParser,
        pdf_path: Path,
        analysis_type: str,
        start_no: int,
        end_no: int,
    ) -> None:
        super().__init__()
        self._parser = parser
        self._pdf_path = pdf_path
        self._analysis_type = analysis_type
        self._start_no = start_no
        self._end_no = end_no
        self._cancelled = Event()

    @Slot()
    def run(self) -> None:
        try:
            batch = self._parser.parse(
                self._pdf_path,
                analysis_type=self._analysis_type,
                analysis_no_start=self._start_no,
                analysis_no_end=self._end_no,
                progress_callback=self.progress.emit,
                cancel_check=self._cancelled.is_set,
            )
            self.completed.emit(batch)
        except Exception as exc:
            self.failed.emit(str(exc))
        finally:
            self.finished.emit()

    def cancel(self) -> None:
        self._cancelled.set()


class PdfRegistrationPage(QWidget):
    extraction_ready = Signal(object)

    def __init__(
        self,
        controller: SharedFolderController,
        parser: LabSolutionsParser,
        database: DatabaseService,
    ) -> None:
        super().__init__()
        self.setObjectName("pageBody")
        self._controller = controller
        self._parser = parser
        self._database = database
        self._final_folder: Path | None = None
        self._thread: QThread | None = None
        self._worker: PdfExtractionWorker | None = None

        layout = QVBoxLayout(self)
        layout.setContentsMargins(28, 24, 28, 28)
        layout.setSpacing(16)

        self.connection_status = QLabel("공유폴더 연결 확인 중...")
        self.connection_status.setProperty("statusTone", "neutral")
        self.connection_status.setWordWrap(True)
        layout.addWidget(self.connection_status)

        cards = QHBoxLayout()
        cards.setSpacing(16)
        location = Card("1. 작업 위치", "결과 파일을 저장할 작업장과 기간 폴더를 지정하세요.")
        location_grid = QGridLayout()
        location_grid.setHorizontalSpacing(12)
        location_grid.setVerticalSpacing(8)
        self.workplace = QComboBox()
        self.year = QSpinBox()
        self.year.setRange(2000, 2099)
        self.year.setValue(date.today().year)
        self.half = QComboBox()
        self.half.addItems([HalfYear.FIRST.value, HalfYear.SECOND.value])
        location_grid.addWidget(field_label("작업장"), 0, 0)
        location_grid.addWidget(self.workplace, 1, 0, 1, 2)
        location_grid.addWidget(field_label("연도"), 2, 0)
        location_grid.addWidget(field_label("주기"), 2, 1)
        location_grid.addWidget(self.year, 3, 0)
        location_grid.addWidget(self.half, 3, 1)
        location.body.addLayout(location_grid)
        location.body.addWidget(field_label("기간 경로"))
        self.period_path = make_path_label("작업장을 선택하세요.")
        location.body.addWidget(self.period_path)
        location.body.addWidget(field_label("최종 저장 폴더"))
        self.final_path = make_path_label("선택되지 않음")
        location.body.addWidget(self.final_path)
        location_actions = QHBoxLayout()
        refresh = QPushButton("작업장 새로고침")
        refresh.clicked.connect(self.refresh_workplaces)
        choose_folder = QPushButton("저장 폴더 선택")
        choose_folder.setProperty("kind", "primary")
        choose_folder.clicked.connect(self.choose_final_folder)
        location_actions.addWidget(refresh)
        location_actions.addStretch(1)
        location_actions.addWidget(choose_folder)
        location.body.addLayout(location_actions)
        cards.addWidget(location, 1)

        source = Card("2. 분석 PDF", "LabSolutions PDF를 선택하면 파일명에서 분석번호를 자동으로 찾습니다.")
        source_grid = QGridLayout()
        source_grid.setHorizontalSpacing(12)
        source_grid.setVerticalSpacing(8)
        self.analysis_type = QComboBox()
        self.analysis_type.addItem("혼유")
        source_grid.addWidget(field_label("분석 종류"), 0, 0, 1, 2)
        source_grid.addWidget(self.analysis_type, 1, 0, 1, 2)
        source_grid.addWidget(field_label("PDF 파일"), 2, 0, 1, 2)
        self.pdf_path = QLineEdit()
        self.pdf_path.setReadOnly(True)
        self.pdf_path.setPlaceholderText("PDF 파일을 선택하세요")
        choose_pdf = QPushButton("파일 선택")
        choose_pdf.clicked.connect(self.choose_pdf)
        source_grid.addWidget(self.pdf_path, 3, 0)
        source_grid.addWidget(choose_pdf, 3, 1)
        source_grid.addWidget(field_label("분석번호 시작"), 4, 0)
        source_grid.addWidget(field_label("분석번호 종료"), 4, 1)
        self.start_no = QSpinBox()
        self.end_no = QSpinBox()
        for box in (self.start_no, self.end_no):
            box.setRange(1, 999999)
        source_grid.addWidget(self.start_no, 5, 0)
        source_grid.addWidget(self.end_no, 5, 1)
        source.body.addLayout(source_grid)
        guide = QLabel("처리 순서  ·  PDF 선택 → 자동 추출 → Peak 검토 → DB 저장")
        guide.setProperty("statusTone", "neutral")
        guide.setWordWrap(True)
        source.body.addWidget(guide)
        source.body.addStretch(1)
        cards.addWidget(source, 1)
        layout.addLayout(cards, 1)

        action = Card("3. PDF 추출", "선택한 PDF를 읽어 Peak Table과 Sample 정보를 추출합니다.")
        self.extraction_status = QLabel("PDF를 선택하면 추출을 시작할 수 있습니다.")
        self.extraction_status.setProperty("statusTone", "neutral")
        self.extraction_status.setWordWrap(True)
        action.body.addWidget(self.extraction_status)
        self.progress = QProgressBar()
        self.progress.setRange(0, 1)
        self.progress.setValue(0)
        self.progress.setTextVisible(False)
        action.body.addWidget(self.progress)
        action_buttons = QHBoxLayout()
        action_buttons.addStretch(1)
        self.cancel_button = QPushButton("추출 취소")
        self.cancel_button.setProperty("kind", "danger")
        self.cancel_button.setEnabled(False)
        self.cancel_button.clicked.connect(self.cancel_extraction)
        self.extract_button = QPushButton("PDF 추출 시작")
        self.extract_button.setProperty("kind", "primary")
        self.extract_button.clicked.connect(self.start_extraction)
        action_buttons.addWidget(self.cancel_button)
        action_buttons.addWidget(self.extract_button)
        action.body.addLayout(action_buttons)
        layout.addWidget(action)

        self.workplace.currentTextChanged.connect(self.update_period_path)
        self.year.valueChanged.connect(self.update_period_path)
        self.half.currentTextChanged.connect(self.update_period_path)
        self.refresh_workplaces()

    def selected_half(self) -> HalfYear:
        return HalfYear(self.half.currentText())

    def refresh_workplaces(self) -> None:
        try:
            state = self._controller.refresh()
        except Exception as exc:
            self.connection_status.setText(f"공유폴더 확인 오류: {exc}")
            set_status_tone(self.connection_status, "error")
            self.workplace.clear()
            return
        self.connection_status.setText(state.connection.message)
        set_status_tone(
            self.connection_status,
            "success" if state.connection.connected else "warning",
        )
        self.workplace.clear()
        self.workplace.addItems(state.workplaces)
        if state.recent.workplace in state.workplaces:
            self.workplace.setCurrentText(state.recent.workplace)
        if state.recent.year:
            self.year.setValue(state.recent.year)
        if state.recent.half in {HalfYear.FIRST.value, HalfYear.SECOND.value}:
            self.half.setCurrentText(state.recent.half)
        if state.recent.final_folder and Path(state.recent.final_folder).is_dir():
            self._final_folder = Path(state.recent.final_folder)
            self.final_path.setText(state.recent.final_folder)
        else:
            self._final_folder = None
            self.final_path.setText("선택되지 않음")
        self.update_period_path()

    def update_period_path(self) -> None:
        workplace = self.workplace.currentText()
        if not workplace:
            self.period_path.setText("작업장을 선택하세요.")
            return
        try:
            result = self._controller.validate_period(
                workplace, self.year.value(), self.selected_half()
            )
            self.period_path.setText(f"{result.path}\n{result.message}")
        except Exception as exc:
            self.period_path.setText(str(exc))

    def choose_final_folder(self) -> None:
        workplace = self.workplace.currentText()
        if not workplace:
            self.final_path.setText("작업장을 먼저 선택하세요.")
            return
        period = self._controller.validate_period(
            workplace, self.year.value(), self.selected_half()
        )
        if not period.valid:
            self.final_path.setText(period.message)
            return
        selected = QFileDialog.getExistingDirectory(self, "최종 저장 폴더 선택", period.path)
        if not selected:
            return
        result = self._controller.select_final_folder(
            Path(selected), workplace=workplace, year=self.year.value(),
            half=self.selected_half(),
        )
        self.final_path.setText(result.path if result.valid else result.message)
        self._final_folder = Path(result.path) if result.valid else None

    def choose_pdf(self) -> None:
        selected, _ = QFileDialog.getOpenFileName(
            self, "LabSolutions PDF 선택", "", "PDF 파일 (*.pdf)"
        )
        if not selected:
            return
        self.pdf_path.setText(selected)
        parsed_range = self._parser.extract_analysis_range(Path(selected).name)
        if parsed_range:
            self.start_no.setValue(parsed_range[0])
            self.end_no.setValue(parsed_range[1])
            self.extraction_status.setText("분석번호를 자동 확인했습니다. PDF 추출을 시작하세요.")
            set_status_tone(self.extraction_status, "success")
        else:
            self.extraction_status.setText("분석번호 범위를 직접 입력하고 확인하세요.")
            set_status_tone(self.extraction_status, "warning")

    def start_extraction(self) -> None:
        path = Path(self.pdf_path.text())
        if not path.is_file():
            self.extraction_status.setText("PDF 파일을 선택하세요.")
            set_status_tone(self.extraction_status, "warning")
            return
        if self.start_no.value() > self.end_no.value():
            self.extraction_status.setText("분석번호 시작값이 종료값보다 큽니다.")
            set_status_tone(self.extraction_status, "error")
            return
        self._thread = QThread(self)
        self._worker = PdfExtractionWorker(
            self._parser, path, self.analysis_type.currentText(),
            self.start_no.value(), self.end_no.value(),
        )
        self._worker.moveToThread(self._thread)
        self._thread.started.connect(self._worker.run)
        self._worker.progress.connect(self._on_progress)
        self._worker.completed.connect(self._on_extraction_completed)
        self._worker.failed.connect(self._on_extraction_failed)
        self._worker.finished.connect(self._thread.quit)
        self._worker.finished.connect(self._worker.deleteLater)
        self._thread.finished.connect(self._thread.deleteLater)
        self._thread.finished.connect(self._on_thread_finished)
        self.extract_button.setEnabled(False)
        self.cancel_button.setEnabled(True)
        self.extraction_status.setText("PDF 내용을 분석하고 있습니다...")
        set_status_tone(self.extraction_status, "neutral")
        self.progress.setRange(0, 0)
        self._thread.start()

    @Slot(int, int)
    def _on_progress(self, current: int, total: int) -> None:
        self.progress.setRange(0, max(total, 1))
        self.progress.setValue(current)
        self.extraction_status.setText(f"PDF 추출 중  ·  {current}/{total} 페이지")

    @Slot(str)
    def _on_extraction_failed(self, message: str) -> None:
        self.extraction_status.setText(message)
        set_status_tone(self.extraction_status, "error")

    @Slot(object)
    def _on_extraction_completed(self, batch) -> None:
        duplicate = self._database.check_duplicate(batch.source_file.file_hash)
        if duplicate.is_duplicate:
            self.extraction_status.setText(
                f"동일 PDF가 이미 저장되어 있습니다: {duplicate.existing_batch_code}"
            )
            set_status_tone(self.extraction_status, "warning")
            return
        self.extraction_status.setText(
            f"추출 완료  ·  Sample {len(batch.samples)}개  ·  경고 {batch.warning_count}개"
        )
        set_status_tone(self.extraction_status, "success")
        self.extraction_ready.emit(batch)

    def cancel_extraction(self) -> None:
        if self._worker:
            self._worker.cancel()
            self.extraction_status.setText("추출 취소 요청 중...")
            set_status_tone(self.extraction_status, "warning")

    def _on_thread_finished(self) -> None:
        self.extract_button.setEnabled(True)
        self.cancel_button.setEnabled(False)
        if self.progress.maximum() == 0:
            self.progress.setRange(0, 1)
        self._thread = None
        self._worker = None
