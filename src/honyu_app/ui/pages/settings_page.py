from __future__ import annotations

from PySide6.QtCore import Signal
from PySide6.QtWidgets import QLabel, QPushButton, QVBoxLayout, QWidget

from honyu_app.application.shared_folder import SharedFolderController
from honyu_app.ui.theme import Card, make_path_label, set_status_tone


class SettingsPage(QWidget):
    storage_refresh_requested = Signal()

    def __init__(self, controller: SharedFolderController) -> None:
        super().__init__()
        self.setObjectName("pageBody")
        self._controller = controller
        layout = QVBoxLayout(self)
        layout.setContentsMargins(28, 24, 28, 28)
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
        layout.addStretch(1)
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
