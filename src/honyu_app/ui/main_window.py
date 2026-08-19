from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QFrame,
    QHBoxLayout,
    QLabel,
    QListWidget,
    QListWidgetItem,
    QMainWindow,
    QStackedWidget,
    QVBoxLayout,
    QWidget,
)

from honyu_app.application.create_excel_export import CreateExcelExportService
from honyu_app.application.preview_excel_export import PreviewExcelExportService
from honyu_app.application.review_extraction import ReviewExtractionService
from honyu_app.application.shared_folder import SharedFolderController
from honyu_app.infrastructure.excel.excel_recalculator import ExcelComRecalculator
from honyu_app.infrastructure.excel.workbook_inspector import XlsxTemplateInspector
from honyu_app.infrastructure.excel.workbook_validator import XlsxWorkbookValidator
from honyu_app.infrastructure.excel.xml_cell_writer import XlsxXmlCellWriter
from honyu_app.infrastructure.pdf.labsolutions_parser import LabSolutionsParser
from honyu_app.services.database_service import DatabaseService
from honyu_app.ui.pages.excel_export_page import ExcelExportPage
from honyu_app.ui.pages.database_page import DatabasePage
from honyu_app.ui.pages.extraction_review_page import ExtractionReviewPage
from honyu_app.ui.pages.pdf_registration_page import PdfRegistrationPage
from honyu_app.ui.theme import APP_STYLESHEET, Card


PAGES = (
    ("01", "PDF 등록", "LabSolutions PDF와 분석 범위를 등록합니다."),
    ("02", "추출 결과 검토", "Peak와 물질 매핑을 검토하고 DB에 저장합니다."),
    ("03", "DB 조회", "저장된 분석 배치와 수정 이력을 조회합니다."),
    ("04", "Excel 생성", "검증된 Area를 기존 Excel 양식에 안전하게 반영합니다."),
    ("05", "설정 및 로그", "연결 설정과 처리 기록을 확인합니다."),
)


def _placeholder(title: str, message: str) -> QWidget:
    page = QWidget()
    page.setObjectName("pageBody")
    layout = QVBoxLayout(page)
    layout.setContentsMargins(28, 24, 28, 28)
    card = Card(title, message)
    status = QLabel("이 기능은 후속 Phase에서 연결됩니다.")
    status.setProperty("statusTone", "neutral")
    card.body.addWidget(status)
    card.body.addStretch(1)
    layout.addWidget(card)
    layout.addStretch(1)
    return page


class MainWindow(QMainWindow):
    def __init__(
        self,
        shared_folder_controller: SharedFolderController,
        parser: LabSolutionsParser,
        database: DatabaseService,
    ) -> None:
        super().__init__()
        self.setWindowTitle("혼유 분석업무 자동화")
        self.resize(1440, 900)
        self.setMinimumSize(1100, 720)
        self.setStyleSheet(APP_STYLESHEET)

        root = QWidget()
        root.setObjectName("appRoot")
        shell = QHBoxLayout(root)
        shell.setContentsMargins(0, 0, 0, 0)
        shell.setSpacing(0)

        sidebar = QWidget()
        sidebar.setObjectName("sidebar")
        sidebar.setFixedWidth(236)
        sidebar_layout = QVBoxLayout(sidebar)
        sidebar_layout.setContentsMargins(18, 22, 18, 18)
        sidebar_layout.setSpacing(14)

        brand = QHBoxLayout()
        brand.setSpacing(11)
        mark = QLabel("HG")
        mark.setObjectName("brandMark")
        mark.setAlignment(Qt.AlignmentFlag.AlignCenter)
        brand_text = QVBoxLayout()
        brand_text.setSpacing(1)
        title = QLabel("혼유 분석 자동화")
        title.setObjectName("brandTitle")
        subtitle = QLabel("ANALYSIS WORKSPACE")
        subtitle.setObjectName("brandSubtitle")
        brand_text.addWidget(title)
        brand_text.addWidget(subtitle)
        brand.addWidget(mark)
        brand.addLayout(brand_text, 1)
        sidebar_layout.addLayout(brand)
        sidebar_layout.addSpacing(14)

        menu_label = QLabel("업무 메뉴")
        menu_label.setObjectName("sidebarMeta")
        sidebar_layout.addWidget(menu_label)
        self.navigation = QListWidget()
        self.navigation.setObjectName("sidebarNavigation")
        self.navigation.setSpacing(1)
        for number, label, _ in PAGES:
            self.navigation.addItem(QListWidgetItem(f"{number}    {label}"))
        sidebar_layout.addWidget(self.navigation, 1)
        mode = QLabel("LOCAL DATABASE  •  v0.1")
        mode.setObjectName("sidebarMeta")
        sidebar_layout.addWidget(mode)
        shell.addWidget(sidebar)

        workspace = QWidget()
        workspace_layout = QVBoxLayout(workspace)
        workspace_layout.setContentsMargins(0, 0, 0, 0)
        workspace_layout.setSpacing(0)
        top_bar = QFrame()
        top_bar.setObjectName("topBar")
        top_bar.setFixedHeight(84)
        top_layout = QVBoxLayout(top_bar)
        top_layout.setContentsMargins(28, 15, 28, 13)
        top_layout.setSpacing(2)
        self.page_title = QLabel()
        self.page_title.setObjectName("pageTitle")
        self.page_subtitle = QLabel()
        self.page_subtitle.setObjectName("pageSubtitle")
        top_layout.addWidget(self.page_title)
        top_layout.addWidget(self.page_subtitle)
        workspace_layout.addWidget(top_bar)

        self.pages = QStackedWidget()
        self.pages.setObjectName("pageStack")
        registration_page = PdfRegistrationPage(shared_folder_controller, parser, database)
        review_page = ExtractionReviewPage(ReviewExtractionService(database))
        database_page = DatabasePage(database)
        preview_service = PreviewExcelExportService(database, XlsxTemplateInspector())
        create_service = CreateExcelExportService(
            database,
            preview_service,
            XlsxXmlCellWriter(),
            XlsxWorkbookValidator(),
            ExcelComRecalculator(),
        )
        excel_page = ExcelExportPage(database, preview_service, create_service)

        registration_page.extraction_ready.connect(review_page.load_batch)
        registration_page.extraction_ready.connect(lambda _: self.navigation.setCurrentRow(1))
        review_page.batch_saved.connect(excel_page.load_batch)
        review_page.batch_saved.connect(database_page.refresh_batches)
        review_page.batch_saved.connect(lambda _: self.navigation.setCurrentRow(3))
        review_page.excel_requested.connect(excel_page.load_batch)
        review_page.excel_requested.connect(lambda _: self.navigation.setCurrentRow(3))
        database_page.review_requested.connect(review_page.load_batch)
        database_page.review_requested.connect(lambda _: self.navigation.setCurrentRow(1))
        database_page.excel_requested.connect(excel_page.load_batch)
        database_page.excel_requested.connect(lambda _: self.navigation.setCurrentRow(3))

        self.pages.addWidget(registration_page)
        self.pages.addWidget(review_page)
        self.pages.addWidget(database_page)
        self.pages.addWidget(excel_page)
        self.pages.addWidget(_placeholder("설정 및 로그 준비 중", "연결 정보와 처리 로그 화면이 들어올 자리입니다."))
        workspace_layout.addWidget(self.pages, 1)
        shell.addWidget(workspace, 1)

        self.navigation.currentRowChanged.connect(self._change_page)
        self.navigation.setCurrentRow(0)
        self.setCentralWidget(root)

    def _change_page(self, index: int) -> None:
        if not 0 <= index < len(PAGES):
            return
        self.pages.setCurrentIndex(index)
        self.page_title.setText(PAGES[index][1])
        self.page_subtitle.setText(PAGES[index][2])
