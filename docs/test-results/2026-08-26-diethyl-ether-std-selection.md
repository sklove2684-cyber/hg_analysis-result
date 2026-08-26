# 디에틸에테르 STD A/B 선택 및 예외 검증

## 목적

- 디에틸에테르 프로필의 고정 `STD1, STD2, STD4, STD5, STD6` 선택 제거
- 공통 STD 방식 A/B 복구
- 실제 PDF의 STD2/STD3 중복 농도 의심 상태를 자동 대체 없이 검토 오류로 처리

이 문서는 `2026-08-26-diethyl-ether-profile.md`의 고정 STD 선택 결과를 대체한다.

## 적용 규칙

- 방식 A: `STD1, STD2, STD3, STD4, STD5`
- 방식 B: `STD1, STD2, STD3, STD4, STD6`
- 선택된 인접 표준의 Diethyl ether Area 차이가 5% 이하이거나 Area 순서가
  역전되면 `STD_LEVEL_REVIEW_REQUIRED` 오류를 발생시킨다.
- 오류가 발생해도 다른 STD 번호로 대체하지 않는다.
- 뒤쪽의 불완전한 재확인 STD 세트는 기존 `DUPLICATE_STD_SET` 제외 규칙을
  유지한다.

PDF Peak Table의 `Conc.` 값은 모든 STD에서 `0.000`이므로 농도값을 직접
확정할 수 없다. 따라서 실제 PDF에서 확인 가능한 대상 물질 Area의 근접성과
순서를 이용해 사용자 검토가 필요한 상태만 탐지한다.

## 실제 파일

- PDF: `디에틸에테르 152,153@완료.pdf`
- Excel: `(디에틸에테르) 152,153.xlsx`
- 비교 양식: `(혼유) 601-690.xlsx`

실제 STD Diethyl ether RT는 모두 `1.305`였다.

| Sample | Area |
|---|---:|
| STD1 | 26,539 |
| STD2 | 62,642 |
| STD3 | 63,347 |
| STD4 | 280,075 |
| STD5 | 534,342 |
| STD6 | 705,737 |
| 뒤쪽 재확인 STD2 | 61,218 |

STD2와 STD3 Area 차이는 약 1.11%로 5% 검토 기준 이내다.

## 실제 PDF 예외 처리 결과

### 방식 A

- `F4:F8`에 미리보기된 STD: STD1, STD2, STD3, STD4, STD5
- STD6: `STD_METHOD_A_NOT_SELECTED`
- 뒤쪽 STD2: `DUPLICATE_STD_SET`
- mapped: 14
- excluded: 30
- error: 1 (`STD_LEVEL_REVIEW_REQUIRED`)
- 생성 가능: 아니오

### 방식 B

- `F4:F8`에 미리보기된 STD: STD1, STD2, STD3, STD4, STD6
- STD5: `STD_METHOD_B_NOT_SELECTED`
- 뒤쪽 STD2: `DUPLICATE_STD_SET`
- mapped: 14
- excluded: 30
- error: 1 (`STD_LEVEL_REVIEW_REQUIRED`)
- 생성 가능: 아니오

두 방식 모두 STD2/STD3을 그대로 표시하고, STD4/STD5/STD6을 임의 대체하지
않았다. 실제 Excel 생성 서비스도 검토 오류가 해소될 때까지 결과 파일을 만들지
않는 것을 확인했다.

## 자동 테스트

- 정상 디에틸에테르 방식 A가 STD1~STD5를 `F4:F8`에 매핑
- 정상 디에틸에테르 방식 B가 STD1~STD4, STD6을 `F4:F8`에 매핑
- STD2/STD3 Area가 근접한 경우 `STD_LEVEL_REVIEW_REQUIRED`
- 실제 PDF의 방식 A/B가 STD 번호를 대체하지 않는지 검증
- 실제 PDF로 Excel 생성 시 검토 오류로 차단되는지 검증
- 정상화된 표준 세트로 실제 Excel의 수식, 스타일, 병합, 차트 보존 검증
- `(혼유) 601-690.xlsx` 선택 시 `TEMPLATE_PROFILE_MISMATCH` 유지
- Diethyl ether RT `1.305` 근접 Peak 선택 유지

## 테스트 결과

- 디에틸에테르 단위 테스트: 통과
- 디에틸에테르 실제 파일 회귀 테스트: `4 passed`
- 전체 테스트: `117 passed, 29 skipped, 2 deselected` (24.26초)
- 실패: 0
- skip/deselect: 현재 PC에 없는 기존 외부 `TEST` fixture 의존 테스트
- 실파일 프로필 회귀:
  - 기존 12개 프로필: 기존 mapped/excluded/error 수치 유지
  - 디에틸에테르: 의도대로 `mapped 14 / excluded 30 / error 1`

## 주의사항

- PDF 자체에 실제 표준 농도값이 들어오지 않으므로 오류는 “중복 농도 확정”이
  아니라 “중복 농도 또는 검량선 순서 이상 의심”이다.
- 현재 실제 PDF는 사용자 확인 없이 자동 생성하지 않는다.
- 정상적인 A/B 표준 Area 흐름에서는 기존 Excel 셀 매핑과 생성 기능을 그대로
  사용한다.
