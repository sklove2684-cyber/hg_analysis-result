from __future__ import annotations

from PySide6.QtCore import Qt
from PySide6.QtWidgets import QFrame, QLabel, QVBoxLayout, QWidget


APP_STYLESHEET = """
* {
    font-family: "Segoe UI Variable", "Segoe UI", "Malgun Gothic";
    font-size: 13px;
    color: #172033;
}
QMainWindow, QWidget#appRoot, QStackedWidget#pageStack {
    background: #f4f6fa;
}
QWidget#sidebar {
    background: #172033;
    border: none;
}
QLabel#brandMark {
    color: #ffffff;
    background: #2f6fed;
    border-radius: 10px;
    font-size: 16px;
    font-weight: 700;
    min-width: 40px;
    max-width: 40px;
    min-height: 40px;
    max-height: 40px;
}
QLabel#brandTitle {
    color: #ffffff;
    font-size: 15px;
    font-weight: 700;
}
QLabel#brandSubtitle, QLabel#sidebarMeta {
    color: #9aa6bb;
    font-size: 11px;
}
QListWidget#sidebarNavigation {
    background: transparent;
    border: none;
    outline: none;
    color: #cbd3e1;
    padding: 0px;
}
QListWidget#sidebarNavigation::item {
    border-radius: 8px;
    margin: 3px 0px;
    padding: 12px 14px;
    min-height: 22px;
}
QListWidget#sidebarNavigation::item:hover {
    background: #222d43;
    color: #ffffff;
}
QListWidget#sidebarNavigation::item:selected {
    background: #2f6fed;
    color: #ffffff;
    font-weight: 600;
}
QFrame#topBar {
    background: #ffffff;
    border-bottom: 1px solid #e4e8f0;
}
QLabel#pageTitle {
    font-size: 22px;
    font-weight: 700;
    color: #111827;
}
QLabel#pageSubtitle {
    color: #667085;
    font-size: 12px;
}
QWidget#pageBody {
    background: #f4f6fa;
}
QFrame[uiCard="true"] {
    background: #ffffff;
    border: 1px solid #e2e7f0;
    border-radius: 10px;
}
QLabel[sectionTitle="true"] {
    color: #111827;
    font-size: 15px;
    font-weight: 700;
}
QLabel[sectionHint="true"] {
    color: #667085;
    font-size: 12px;
}
QLabel[fieldLabel="true"] {
    color: #475467;
    font-size: 12px;
    font-weight: 600;
}
QLabel[statValue="true"] {
    color: #111827;
    font-size: 22px;
    font-weight: 700;
}
QLabel[statLabel="true"] {
    color: #667085;
    font-size: 11px;
}
QLabel[pathValue="true"] {
    background: #f8fafc;
    border: 1px solid #e5e9f0;
    border-radius: 7px;
    padding: 9px 11px;
    color: #475467;
}
QLabel[statusTone="neutral"] {
    background: #eef2f7;
    color: #475467;
    border-radius: 7px;
    padding: 8px 11px;
}
QLabel[statusTone="success"] {
    background: #e9f8f0;
    color: #16643b;
    border-radius: 7px;
    padding: 8px 11px;
}
QLabel[statusTone="warning"] {
    background: #fff6df;
    color: #8a5700;
    border-radius: 7px;
    padding: 8px 11px;
}
QLabel[statusTone="error"] {
    background: #ffeded;
    color: #a32626;
    border-radius: 7px;
    padding: 8px 11px;
}
QLineEdit, QComboBox, QSpinBox {
    background: #ffffff;
    border: 1px solid #cfd6e2;
    border-radius: 7px;
    padding: 7px 10px;
    min-height: 20px;
    selection-background-color: #2f6fed;
}
QLineEdit:focus, QComboBox:focus, QSpinBox:focus {
    border: 2px solid #2f6fed;
    padding: 6px 9px;
}
QLineEdit:read-only {
    background: #f8fafc;
    color: #475467;
}
QComboBox::drop-down, QSpinBox::up-button, QSpinBox::down-button {
    border: none;
    width: 24px;
}
QPushButton {
    background: #ffffff;
    border: 1px solid #cfd6e2;
    border-radius: 7px;
    padding: 8px 14px;
    min-height: 20px;
    font-weight: 600;
}
QPushButton:hover {
    background: #f5f7fb;
    border-color: #aeb8c8;
}
QPushButton:pressed {
    background: #e9edf4;
}
QPushButton[kind="primary"] {
    background: #2f6fed;
    color: #ffffff;
    border-color: #2f6fed;
}
QPushButton[kind="primary"]:hover {
    background: #255ed0;
    border-color: #255ed0;
}
QPushButton[kind="danger"] {
    background: #ffffff;
    color: #b42318;
    border-color: #f0b7b2;
}
QPushButton:disabled {
    background: #edf0f4;
    color: #98a2b3;
    border-color: #e2e7ed;
}
QTableWidget {
    background: #ffffff;
    alternate-background-color: #f8fafc;
    border: 1px solid #e2e7f0;
    border-radius: 9px;
    gridline-color: transparent;
    selection-background-color: #e7efff;
    selection-color: #172033;
    outline: none;
}
QTableWidget::item {
    border-bottom: 1px solid #edf0f4;
    padding: 6px 8px;
}
QHeaderView::section {
    background: #f7f9fc;
    color: #475467;
    border: none;
    border-bottom: 1px solid #dfe4ec;
    padding: 9px 8px;
    font-weight: 700;
}
QScrollBar:vertical {
    background: transparent;
    width: 10px;
    margin: 2px;
}
QScrollBar::handle:vertical {
    background: #c5ccd8;
    border-radius: 4px;
    min-height: 32px;
}
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical,
QScrollBar::add-page:vertical, QScrollBar::sub-page:vertical {
    background: transparent;
    height: 0px;
}
QProgressBar {
    background: #e9edf4;
    border: none;
    border-radius: 4px;
    min-height: 8px;
    max-height: 8px;
    text-align: center;
}
QProgressBar::chunk {
    background: #2f6fed;
    border-radius: 4px;
}
QToolTip {
    background: #172033;
    color: #ffffff;
    border: none;
    padding: 6px;
}
"""


class Card(QFrame):
    def __init__(
        self,
        title: str,
        hint: str = "",
        parent: QWidget | None = None,
    ) -> None:
        super().__init__(parent)
        self.setProperty("uiCard", True)
        self.body = QVBoxLayout(self)
        self.body.setContentsMargins(20, 18, 20, 20)
        self.body.setSpacing(12)
        heading = QLabel(title)
        heading.setProperty("sectionTitle", True)
        self.body.addWidget(heading)
        if hint:
            description = QLabel(hint)
            description.setProperty("sectionHint", True)
            description.setWordWrap(True)
            self.body.addWidget(description)


def field_label(text: str) -> QLabel:
    label = QLabel(text)
    label.setProperty("fieldLabel", True)
    return label


def set_status_tone(label: QLabel, tone: str) -> None:
    label.setProperty("statusTone", tone)
    label.style().unpolish(label)
    label.style().polish(label)


def make_path_label(text: str) -> QLabel:
    label = QLabel(text)
    label.setProperty("pathValue", True)
    label.setWordWrap(True)
    label.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse)
    return label
