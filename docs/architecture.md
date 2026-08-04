# Architecture

UI → application use cases → service protocols → infrastructure adapters 순서로 의존합니다.

UI는 Mock DB와 Supabase를 구분하지 않습니다. 두 구현은 동일한 `DatabaseService` 계약을 따릅니다.

PHASE 1에서는 PDF parser와 DB 구현을 연결하지 않습니다. 기존 parser가 전달되면 `PdfParser` 계약 뒤에 배치합니다.

PHASE 2의 공유폴더 탐색은 `SharedFolderService` 계약 뒤에 배치합니다. UNC가 우선이며 Z:는 보조 경로일 뿐 저장 기준으로 기록하지 않습니다. 폴더가 없을 때는 생성하지 않습니다.

PHASE 3의 `MockDatabaseService`는 로컬 SQLite를 사용합니다. 운영에서는 같은 `DatabaseService` 계약의 Supabase 구현으로 교체하며 UI와 application 계층은 변경하지 않습니다. SQLite 파일을 UNC 또는 Z: 공유폴더에 두지 않습니다.

## PHASE 6 Excel 입력 미리보기

- `infrastructure/excel/workbook_inspector.py`: XLSX ZIP/XML을 읽기 전용으로 분석하고 셀 값 유형·수식·스타일 ID를 제공합니다.
- `application/preview_excel_export.py`: STD A/B, 회수율, 작업자 분석번호, 물질 열, DIBK 적용 Area 상위 2개를 목표 셀로 변환합니다.
- `services/excel_template_service.py`: 템플릿 분석 구현을 UI·업무 로직에서 분리합니다.
- `ui/pages/excel_export_page.py`: 배치·STD 방식·원본 Excel 선택과 입력 미리보기 표를 제공합니다.
- PHASE 6은 원본 Excel을 수정하지 않습니다. 실제 복사·XML 숫자 입력·COM 재계산은 PHASE 7의 책임입니다.
- 수식 셀 입력, 목표 셀 중복, 작업자 행 0/복수 매칭, 필수 시트 누락은 생성 차단 오류입니다.

## PHASE 7 Excel 생성

- `infrastructure/excel/xml_cell_writer.py`: 원본 XLSX ZIP의 부품을 그대로 복사하고 승인된 비수식 셀의 숫자 값만 기록합니다.
- `infrastructure/excel/workbook_validator.py`: 입력 셀 외 값·수식·스타일, 시트 구조, 병합, 인쇄 설정, 차트·그림 부품이 보존됐는지 검사합니다.
- `infrastructure/excel/excel_recalculator.py`: 설치된 Microsoft Excel을 COM으로 열어 `CalculateFullRebuild` 후 저장합니다.
- `application/create_excel_export.py`: 미리보기 재검증 → 임시 XLSX 입력 → 사전 구조 검증 → Excel 재계산 → 사후 구조 검증 → 최종 파일 승격 → export job 저장 순서를 보장합니다.
- 어느 단계든 실패하면 임시 파일을 최종 이름으로 승격하지 않습니다. 원본과 같은 경로 및 기존 결과 파일 덮어쓰기도 차단합니다.
- Excel COM을 시작할 수 없는 PC에서는 생성이 중단되며 재계산 실패 점검본만 남습니다.
