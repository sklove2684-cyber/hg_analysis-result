# 혼유 분석업무 자동화

Windows에서 LabSolutions 분석 PDF를 검토하고, 검토 완료 데이터를 저장한 뒤 기존 Excel 양식으로 내보내기 위한 프로그램입니다.

현재는 PHASE 7까지 구현되었습니다. 입력 미리보기 후 원본을 보존한 채 결과 Excel을 생성하고, Microsoft Excel 전체 재계산과 구조 검증을 수행합니다.

## PDF 추출 및 검토

- LabSolutions 8열 Peak Table을 표 셀 기준으로 읽습니다.
- 빈 물질명 잔 피크도 Peak 번호, RT, Area와 함께 보존합니다.
- 전체 Peak 개수는 고정 검증값으로 사용하지 않습니다.
- 미등록 물질은 자동 추정하지 않고 검토 완료를 차단합니다.
- DIBK 다중 Peak는 합치지 않고 모두 보존하며, Excel에는 Sample별 적용 Area 상위 2개만 입력합니다.
- 검토 완료 데이터만 Mock DB에 저장합니다.

## Mock DB

- SQLite 파일은 `%LOCALAPPDATA%\HonyuAutomation\mock_db\honyu_mock.db`에 저장합니다.
- SQLite 파일을 회사 공유폴더에 두지 않습니다.
- `area_raw`는 수정하지 않고 `peak_corrections`에 revision 이력을 추가합니다.
- 동일한 SHA-256 `file_hash`는 다시 저장할 수 없습니다.

## 공유폴더 동작

- UNC 기본 경로를 먼저 확인합니다.
- UNC에 연결할 수 없으면 Z: 보조 경로를 확인합니다.
- 두 경로가 모두 실패하면 Excel 저장 기능만 사용할 수 없는 상태로 표시합니다.
- 작업장 및 `26상`/`26하` 기간 폴더는 자동 생성하지 않습니다.
- 최근 작업장, 연도, 반기, 최종 저장 폴더는 PC의 로컬 설정에 저장합니다.

## 개발 실행

```powershell
py -3.11 -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -e ".[dev]"
honyu-automation
```

## 원칙

- 원본 PDF와 Excel은 수정하지 않습니다.
- UI는 서비스 인터페이스를 통해 업무 로직을 호출합니다.
- Mock DB와 향후 Supabase 구현은 같은 `DatabaseService` 계약을 사용합니다.
- PDF 전체 Peak 개수는 잔 피크 때문에 고정 검증값으로 사용하지 않습니다.
- 숨겨진 미사용 Excel 영역의 기존 `#REF!`는 오류 판정 대상에서 제외합니다.

## PHASE 6~7 Excel 미리보기 및 생성

- DB 분석 배치, 원본 Excel, STD 방식 A/B를 선택할 수 있습니다.
- 각 Peak의 적용 Area, 목표 시트·셀, 기존 셀 유형, 수식 여부, 처리 상태를 확인합니다.
- 작업자 Sample은 `area!A37:A183` 분석번호의 마지막 숫자로 유일 매칭합니다.
- DIBK는 모든 피크를 DB에 보존하고 적용 Area 상위 2개만 Excel 두 슬롯에 매핑합니다.
- 수식 셀 충돌이나 작업자 행 불일치가 있으면 생성 가능 상태를 차단합니다.
- 결과 경로를 고른 뒤 오류 없는 미리보기에 대해서만 Excel 생성 버튼이 활성화됩니다.
- XLSX ZIP/XML에서 승인된 입력 셀에 정수 Area만 기록하므로 수식·서식·차트는 다시 만들지 않습니다.
- 임시 복사본을 사전 검증한 뒤 Microsoft Excel COM으로 전체 재계산하고, 다시 구조를 검증한 경우에만 최종 파일명으로 승격합니다.
- 실패 시 최종 파일은 만들지 않으며 `_검증실패_...xlsx` 또는 `_재계산실패_...xlsx` 점검본을 남깁니다.
- 같은 이름의 기존 결과 파일은 덮어쓰지 않습니다.
