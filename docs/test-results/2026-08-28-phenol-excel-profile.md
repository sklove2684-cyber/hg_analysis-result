# 페놀 Excel 프로필 실파일 검증

## 실제 파일

- PDF: `(페놀) 256-305.pdf`
- Excel: `(페놀) 256-305.xlsx`
- 비교용 잘못된 양식: `(혼유) 601-690.xlsx`

## 등록 프로필

- 분석종류: `페놀`
- profile key: `phenol`
- canonical material: `Phenol`
- PDF aliases: `Phenol`, `페놀`
- 필수 시트: `검량선`, `LOD(area입력)`, `회수율`, `std`
- 양식 식별: `LOD(area입력)!E2 = Phenol`
- STD Area: `LOD(area입력)!F4:F8`
- 회수율 Area: `회수율!B28:B36`
- 일반시료 Area: `LOD(area입력)!F19:F122`
- 분석번호: `LOD(area입력)!A19:A122`
- STD 방식 A: STD1, STD2, STD3, STD4, STD5
- STD 방식 B: STD1, STD2, STD3, STD4, STD6
- Peak 선택: 선택된 STD 세트에서 실제 Phenol RT를 얻어 가장 가까운 Phenol Peak 1개 선택

## 실제 양식 분석번호

`256, 257, 277, 278, 281, 282, 287, 288, 293, 294, 295, 296, 302, 303, 304, 305`

## 실제 PDF 기준 RT 및 입력 결과

- STD1~STD4 Phenol RT: `2.754`
- STD5~STD6 Phenol RT: `2.756`
- Phenol Peak: 15개(STD 6개, 회수율 9개)
- Methanol Peak: 53개, Excel 입력 0개
- UNKNOWN_MATERIAL: 0개
- 방식 A/B 모두 Excel 입력 14개, 입력 제외 54개, 검증 오류 0개
- 실제 일반시료에는 Phenol Peak가 없으므로 일반시료 Area의 원본 `N.D`를 유지

## 보존 및 차단 검증

- 입력 Area는 숫자 형식으로 기록됨
- 기존 수식과 style ID가 동일함
- 병합 범위가 동일함
- 차트, drawing, media 바이너리가 동일함
- 생성 결과의 네 시트를 렌더링해 레이아웃 보존 확인
- 혼유 양식 선택 시 `TEMPLATE_PROFILE_MISMATCH`로 생성 차단
- COM 강제 재계산은 사용하지 않음

## 자동 테스트 결과

- 페놀 관련 대상 테스트: 80 passed, 2 skipped
- 전체 테스트: 170 passed, 31 skipped, 0 failed
