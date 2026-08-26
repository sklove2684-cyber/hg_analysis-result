# 디에틸에테르 STD A/B 및 PDF STD RT 선택 검증

## 목적

- 디에틸에테르에 공통 STD 방식 A/B 적용
- Area 크기나 Area 차이로 STD 농도를 판단하지 않음
- 일반 시료의 복수 Diethyl ether Peak를 실제 PDF STD RT 기준으로 선택

이 문서는 같은 날짜에 작성된 이전 5% Area 판정 결과를 대체한다.

## 적용 규칙

- 방식 A: `STD1, STD2, STD3, STD4, STD5`
- 방식 B: `STD1, STD2, STD3, STD4, STD6`
- 5% Area 차이 및 `STD_LEVEL_REVIEW_REQUIRED` 판정은 완전히 제거했다.
- Area 크기로 STD 농도 중복을 판단하거나 다른 STD 번호를 대체하지 않는다.
- 뒤쪽의 불완전한 재확인 STD 세트는 기존 `DUPLICATE_STD_SET` 제외 규칙을
  유지한다.

## STD RT 결정 방식

프로젝트에는 PDF에서 동적으로 STD 대표 RT를 구하는 기존 공통 함수가 없었다.
평균이나 중앙값을 새로 계산하지 않고, 선택된 A/B STD 번호 순서에서
`Diethyl ether` 명명 Peak가 정확히 1개 확인되는 첫 STD의 실제 RT를 그대로
사용한다.

STD RT를 확인할 수 없으면 Area가 큰 Peak로 대체하지 않고
`STD_TARGET_RT_NOT_FOUND` 오류로 생성하지 않는다.

## 실제 파일

- PDF: `디에틸에테르 152,153@완료.pdf`
- Excel: `(디에틸에테르) 152,153.xlsx`
- 비교 양식: `(혼유) 601-690.xlsx`

실제 선택 기준은 `STD1`의 Diethyl ether RT `1.305`였다. 이는 코드 고정값이
아니라 해당 PDF에서 읽은 값이다.

| Sample | RT | Area |
|---|---:|---:|
| STD1 | 1.305 | 26,539 |
| STD2 | 1.305 | 62,642 |
| STD3 | 1.305 | 63,347 |
| STD4 | 1.305 | 280,075 |
| STD5 | 1.305 | 534,342 |
| STD6 | 1.305 | 705,737 |
| 뒤쪽 재확인 STD2 | 1.305 | 61,218 |

## 실제 PDF 검증 결과

### 방식 A

- `F4:F8`: STD1, STD2, STD3, STD4, STD5
- STD6: `STD_METHOD_A_NOT_SELECTED`
- 뒤쪽 STD2: `DUPLICATE_STD_SET`
- mapped: 14
- excluded: 30
- error: 0
- 생성 가능: 예

### 방식 B

- `F4:F8`: STD1, STD2, STD3, STD4, STD6
- STD5: `STD_METHOD_B_NOT_SELECTED`
- 뒤쪽 STD2: `DUPLICATE_STD_SET`
- mapped: 14
- excluded: 30
- error: 0
- 생성 가능: 예

## Peak 선택 자동 테스트

- Diethyl ether Peak 1개: 그대로 선택
- Diethyl ether Peak 여러 개: PDF STD RT와 가장 가까운 1개 선택
- 더 큰 Area라도 STD RT에서 더 멀면 제외
- 고정 RT `1.305`를 사용하지 않는지 확인하기 위해 STD RT `1.250` 테스트 사용
  - RT `1.252`, Area `123`: 선택
  - RT `1.305`, Area `9,999`: `MATERIAL_RT_NOT_CLOSEST` 제외
- STD RT를 구할 수 없으면 Area 최대 Peak로 대체하지 않음

## Excel 구조 검증

- 실제 PDF + Excel 방식 A 생성 성공
- 숫자 Area 입력 확인
- 기존 수식, 스타일, 병합셀, 차트/드로잉/미디어 보존 확인
- `(혼유) 601-690.xlsx` 선택 시 `TEMPLATE_PROFILE_MISMATCH` 유지

## 회귀 테스트

- 정상 디에틸에테르 방식 A/B
- 실제 디에틸에테르 PDF 방식 A/B
- 단일/복수 Peak RT 선택
- 기존 다른 분석종류의 STD A/B 및 프로필 결과 유지

- 전체 pytest: `118 passed, 29 skipped, 2 deselected` (34.70초)
- 실패: 0
- 디에틸에테르 실제 파일 회귀: `3 passed`
- 기존 12개 실파일 프로필: 기존 mapped/excluded/error 수치 유지
- 디에틸에테르 실파일: `mapped 14 / excluded 30 / error 0 / 생성 가능`

## 주의사항

- STD RT 선택에 평균/중앙값을 사용하지 않는다.
- 현재 실제 PDF의 STD2와 STD3 Area가 비슷하더라도 이것만으로 농도 중복을
  판정하지 않는다.
