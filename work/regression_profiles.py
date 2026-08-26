from pathlib import Path
from tempfile import TemporaryDirectory

from honyu_app.application.preview_excel_export import PreviewExcelExportService
from honyu_app.application.review_extraction import ReviewExtractionService
from honyu_app.infrastructure.database.mock_database_service import MockDatabaseService
from honyu_app.infrastructure.excel.workbook_inspector import XlsxTemplateInspector
from honyu_app.infrastructure.pdf.labsolutions_parser import LabSolutionsParser


cases = (
    (
        "혼유",
        Path(r"C:\Users\양세경\Desktop\분석프로그램\혼유 74-119@완료.pdf"),
        Path(r"C:\Users\양세경\Desktop\분석프로그램\(혼유) (53,54) 74-119 미입력.xlsx"),
        74,
        119,
    ),
    (
        "1컬럼혼유",
        Path(r"C:\Users\양세경\Desktop\1컬럼혼유 120, 123-130 병합완료.pdf"),
        Path(r"C:\Users\양세경\Desktop\(1컬럼) 120, 123-130.xlsx"),
        120,
        130,
    ),
    (
        "(알콜2) IBA,1-BTOH",
        Path(r"C:\Users\양세경\Desktop\분석프로그램\알콜(2) 74-119.pdf"),
        Path(r"C:\Users\양세경\Desktop\분석프로그램\(알콜2) 74-119 빈양식.xlsx"),
        74,
        119,
    ),
    (
        "MEK",
        Path(r"C:\Users\양세경\Desktop\분석프로그램\MEK 74-119.pdf"),
        Path(r"C:\Users\양세경\Desktop\분석프로그램\(MEK) 74-119 입력안됨.xlsx"),
        74,
        119,
    ),
    (
        "(혼유-G2) THF,CFM,벤젠,클로로벤젠",
        Path(r"C:\Users\양세경\Desktop\분석프로그램\G2-혼유 655-686.pdf"),
        Path(r"C:\Users\양세경\Desktop\분석프로그램\(혼유-G2) THF,CFM,벤젠,사염화탄소,클로로벤젠 655-686.xlsx"),
        655,
        686,
    ),
    (
        "이소아밀,n-프로필 아세테이트",
        Path(r"C:\Users\양세경\Desktop\분석프로그램\(이소아밀,n-프로필 아세테이트) 611-631 병합.pdf"),
        Path(r"C:\Users\양세경\Desktop\분석프로그램\(이소아밀,n-프로필 아세테이트) 611-631.xlsx"),
        611,
        631,
    ),
    (
        "셀로솔브",
        Path(r"C:\Users\양세경\Desktop\분석프로그램\(셀로솔브) 681-690 병합.pdf"),
        Path(r"C:\Users\양세경\Desktop\분석프로그램\(셀로솔브) 681-690.xlsx"),
        681,
        690,
    ),
    (
        "(혼유-G3) 1,2-디클로로에틸렌,퍼클로로에틸렌,프로판,에탄",
        Path(r"C:\Users\양세경\Desktop\분석프로그램\G3 혼유 695-696.pdf"),
        Path(r"C:\Users\양세경\Desktop\분석프로그램\(혼유-G3) 695-696 빈양식.xlsx"),
        695,
        696,
    ),
    (
        "초산",
        Path(r"C:\Users\양세경\Desktop\분석프로그램\(초산) 637-666.pdf"),
        Path(r"C:\Users\양세경\Desktop\분석프로그램\(초산) 637-666 빈양식.xlsx"),
        637,
        666,
    ),
    (
        "ACN",
        Path(r"C:\Users\양세경\Desktop\분석프로그램\ACN 656-666.pdf"),
        Path(r"C:\Users\양세경\Desktop\분석프로그램\(ACN) 656-666 빈양식.xlsx"),
        656,
        666,
    ),
    (
        "에틸렌글리콜",
        Path(r"C:\Users\양세경\Desktop\분석프로그램\E.G 599-680.pdf"),
        Path(r"C:\Users\양세경\Desktop\분석프로그램\(에틸렌글리콜) 599-680 미입력.xlsx"),
        599,
        680,
    ),
    (
        "B.C",
        Path(r"C:\Users\양세경\Desktop\분석프로그램\(B.C) 570-588.pdf"),
        Path(r"C:\Users\양세경\Desktop\분석프로그램\(B.C) 570-588 빈양식.xlsx"),
        570,
        588,
    ),
    (
        "디에틸에테르",
        Path(r"C:\Users\양세경\Desktop\분석프로그램\디에틸에테르 152,153@완료.pdf"),
        Path(r"C:\Users\양세경\Desktop\분석프로그램\(디에틸에테르) 152,153.xlsx"),
        152,
        153,
    ),
)

for analysis_type, pdf, xlsx, start, end in cases:
    batch = LabSolutionsParser().parse(
        pdf,
        analysis_type=analysis_type,
        analysis_no_start=start,
        analysis_no_end=end,
    )
    with TemporaryDirectory() as temp:
        database = MockDatabaseService(Path(temp) / "regression.db")
        review = ReviewExtractionService(database)
        review.complete_review(batch)
        saved = review.save_batch(batch)
        result = PreviewExcelExportService(database, XlsxTemplateInspector()).preview(
            saved.batch_id, xlsx, "A"
        )
        print(
            analysis_type,
            result.can_generate,
            result.mapped_count,
            result.excluded_count,
            result.error_count,
            flush=True,
        )
