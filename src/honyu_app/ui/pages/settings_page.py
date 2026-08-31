from __future__ import annotations

from PySide6.QtCore import QObject, QThread, Signal, Slot
from PySide6.QtWidgets import (
    QApplication,
    QHBoxLayout,
    QLabel,
    QPlainTextEdit,
    QPushButton,
    QVBoxLayout,
    QWidget,
)

from honyu_app.application.environment_diagnostics import (
    EnvironmentDiagnosticReport,
    EnvironmentDiagnosticService,
)
from honyu_app.application.shared_folder import SharedFolderController
from honyu_app.ui.theme import Card, make_path_label, set_status_tone


class DiagnosticWorker(QObject):
    completed = Signal(object)
    failed = Signal(str)
    finished = Signal()

    def __init__(
        self,
        service: EnvironmentDiagnosticService,
        pdf_path: str,
        excel_path: str,
    ) -> None:
        super().__init__()
        self._service = service
        self._pdf_path = pdf_path
        self._excel_path = excel_path

    @Slot()
    def run(self) -> None:
        try:
            self.completed.emit(
                self._service.collect(self._pdf_path, self._excel_path)
            )
        except Exception as exc:
            self.failed.emit(str(exc))
        finally:
            self.finished.emit()


class SettingsPage(QWidget):
    storage_refresh_requested = Signal()

    def __init__(
        self,
        controller: SharedFolderController,
        diagnostic_service: EnvironmentDiagnosticService,
    ) -> None:
        super().__init__()
        self.setObjectName("pageBody")
        self._controller = controller
        self._diagnostic_service = diagnostic_service
        self._pdf_path = ""
        self._excel_path = ""
        self._diagnostic_thread: QThread | None = None
        self._diagnostic_worker: DiagnosticWorker | None = None
        self._diagnostic_refresh_pending = False
        self._last_report: EnvironmentDiagnosticReport | None = None

        layout = QVBoxLayout(self)
        layout.setContentsMargins(28, 24, 28, 28)
        layout.setSpacing(14)
        card = Card("저장 경로", "회사 공유폴더 연결 상태와 현재 실제 저장 경로입니다.")
        self.mode = QLabel()
        self.mode.setWordWrap(True)
        self.path = make_path_label("")
        self.attempted = QLabel()
        self.attempted.setWordWrap(True)
        refresh = QPushButton("연결 상태 다시 확인")
        refresh.clicked.connect(self.storage_refresh_requested.emit)
        card.body.addWidget(self.mode)
        card.body.addWidget(self.path)
        card.body.addWidget(self.attempted)
        card.body.addWidget(refresh)
        layout.addWidget(card)

        diagnostic = Card(
            "환경 진단 정보",
            "현재 실행본과 실제 선택한 PDF/Excel 경로를 기준으로 계산합니다.",
        )
        self.diagnostic_status = QLabel("진단 정보 확인 대기 중...")
        self.diagnostic_status.setWordWrap(True)
        set_status_tone(self.diagnostic_status, "neutral")
        self.diagnostic_text = QPlainTextEdit()
        self.diagnostic_text.setReadOnly(True)
        self.diagnostic_text.setLineWrapMode(QPlainTextEdit.LineWrapMode.NoWrap)
        self.diagnostic_text.setPlaceholderText("진단 정보를 수집하고 있습니다.")
        actions = QHBoxLayout()
        refresh_diagnostic = QPushButton("진단정보 새로고침")
        refresh_diagnostic.clicked.connect(self.request_diagnostic_refresh)
        self.copy_diagnostic_button = QPushButton("진단정보 복사")
        self.copy_diagnostic_button.setProperty("kind", "primary")
        self.copy_diagnostic_button.setEnabled(False)
        self.copy_diagnostic_button.clicked.connect(self.copy_diagnostic)
        actions.addStretch(1)
        actions.addWidget(refresh_diagnostic)
        actions.addWidget(self.copy_diagnostic_button)
        diagnostic.body.addWidget(self.diagnostic_status)
        diagnostic.body.addWidget(self.diagnostic_text, 1)
        diagnostic.body.addLayout(actions)
        layout.addWidget(diagnostic, 1)

        self.mode.setText("저장 경로 확인 대기 중...")
        self.path.setText("GUI 표시 후 공유폴더 연결을 확인합니다.")
        self.attempted.setText("")

    def refresh_status(self) -> None:
        directory, status = self._controller.export_directory()
        self.apply_storage_status(directory, status)

    def apply_storage_status(self, directory, status) -> None:
        company = status.storage_mode == "company"
        local = status.storage_mode == "local"
        self.mode.setText(
            "회사 공유폴더 저장" if company
            else "로컬 저장 모드" if local
            else "저장 경로 사용 불가"
        )
        set_status_tone(self.mode, "success" if company else "warning" if local else "error")
        self.path.setText(f"현재 저장 경로\n{directory or '사용 가능한 저장 경로 없음'}")
        self.attempted.setText(
            "공유폴더 확인 경로\n" + "\n".join(status.attempted_paths)
        )

    @Slot(str)
    def set_selected_pdf(self, path: str) -> None:
        if path == self._pdf_path:
            return
        self._pdf_path = path
        if self.isVisible():
            self.request_diagnostic_refresh()

    @Slot(str)
    def set_selected_excel(self, path: str) -> None:
        if path == self._excel_path:
            return
        self._excel_path = path
        if self.isVisible():
            self.request_diagnostic_refresh()

    @Slot()
    def request_diagnostic_refresh(self) -> None:
        if self._diagnostic_thread is not None:
            self._diagnostic_refresh_pending = True
            return
        self.diagnostic_status.setText("현재 경로의 진단 정보를 확인하고 있습니다...")
        set_status_tone(self.diagnostic_status, "neutral")
        thread = QThread(self)
        worker = DiagnosticWorker(
            self._diagnostic_service, self._pdf_path, self._excel_path
        )
        worker.moveToThread(thread)
        thread.started.connect(worker.run)
        worker.completed.connect(self._apply_diagnostic_report)
        worker.failed.connect(self._diagnostic_failed)
        worker.finished.connect(thread.quit)
        worker.finished.connect(worker.deleteLater)
        thread.finished.connect(thread.deleteLater)
        thread.finished.connect(self._diagnostic_finished)
        self._diagnostic_thread = thread
        self._diagnostic_worker = worker
        thread.start()

    @Slot(object)
    def _apply_diagnostic_report(self, report: EnvironmentDiagnosticReport) -> None:
        self._last_report = report
        self.diagnostic_text.setPlainText(report.text)
        self.copy_diagnostic_button.setEnabled(True)
        if report.git_matches is True:
            self.diagnostic_status.setText("[정상] Local == origin/main")
            set_status_tone(self.diagnostic_status, "success")
        elif report.git_matches is False:
            self.diagnostic_status.setText("[주의] Local != origin/main")
            set_status_tone(self.diagnostic_status, "warning")
        else:
            self.diagnostic_status.setText("[주의] Git 비교 확인 실패")
            set_status_tone(self.diagnostic_status, "warning")

    @Slot(str)
    def _diagnostic_failed(self, detail: str) -> None:
        self.diagnostic_status.setText(f"진단 정보 일부를 확인하지 못했습니다: {detail}")
        set_status_tone(self.diagnostic_status, "warning")

    @Slot()
    def _diagnostic_finished(self) -> None:
        self._diagnostic_worker = None
        self._diagnostic_thread = None
        if self._diagnostic_refresh_pending:
            self._diagnostic_refresh_pending = False
            self.request_diagnostic_refresh()

    @Slot()
    def copy_diagnostic(self) -> None:
        text = self.diagnostic_text.toPlainText()
        if not text:
            return
        QApplication.clipboard().setText(text)
        self.diagnostic_status.setText("진단정보를 클립보드에 복사했습니다.")
        set_status_tone(self.diagnostic_status, "success")
