# 메탄올A Excel 프로필 검증 (2026-08-28)

## 실제 파일

- PDF: `(메탄올)A 237-320.pdf`
- Excel: `(메탄올)A 237-320.xlsx`
- 분석종류: `메탄올A`
- canonical material: `Methanol`

## 확인한 Excel 구조

- 필수 시트: `검량선`, `LOD(area입력)`, `회수율`, `std`
- 양식 식별 헤더: `LOD(area입력)!E2 = Methanol`
- STD Area: `LOD(area입력)!F4:F8`
  - 방식 A: STD1, STD2, STD3, STD4, STD5
  - 방식 B: STD1, STD2, STD3, STD4, STD6
- 회수율 Area: `회수율!B28:B36`
- 일반시료 분석번호: `LOD(area입력)!A19:A80`
- 일반시료 Area: `LOD(area입력)!F19:F80`

## 실제 PDF/Excel 결과

- PDF 페이지: 25
- 전체 Peak: 93
- UNKNOWN_MATERIAL: 0
- 선택된 PDF STD Methanol RT: 2.030~2.032분
- Excel 입력: 16
- 입력 제외: 77
- 검증 오류: 0
- 일반시료 매핑:
  - `237` -> `LOD(area입력)!F19 = 1685`
  - `238` -> `LOD(area입력)!F20 = 1930`
- 방식 A 마지막 STD: `F8 = 198028` (STD5)
- 방식 B 마지막 STD: `F8 = 245337` (STD6)
- 잘못된 `(혼유) 601-690.xlsx` 선택: `TEMPLATE_PROFILE_MISMATCH`로 차단
- 생성 파일: 숫자 Area 입력 및 구조 검증 통과
- 원본 보존 검증: 수식, style id, 병합 범위, 차트/도형/미디어 동일

## 자동 테스트

- 메탄올A registry의 `excel_profile_key = methanol`
- Methanol/메탄올/MeOH canonical 정규화 회귀
- STD 방식 A/B 셀 선택
- PDF STD RT 기준 복수 Peak 최근접 선택(큰 Area 우선 금지)
- 실제 PDF + Excel preview 및 생성
- 잘못된 혼유 양식 차단
- 전체 테스트: `158 passed, 31 skipped, 0 failed` (189 collected)
