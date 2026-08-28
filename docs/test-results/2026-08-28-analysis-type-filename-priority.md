# 분석종류 파일명 우선 판별 회귀 검증

## 변경 목적

- PDF 파일명에서 분석종류가 명확히 판별되면 추출 후 Method Filename 또는 검출 물질이 이를 덮어쓰지 않도록 했다.
- 파일명으로 판별할 수 없을 때만 Method Filename과 검출 물질을 보조 근거로 사용한다.

## 실제 파일 검증

- PDF: `(페놀) 256-305.pdf`
- 추출 결과: Sample 53개, Peak 68개, 경고 0개
- 검출 물질: Methanol 53 Peak, Phenol 15 Peak
- 최종 분석종류: `페놀`
- Methanol 처리: Excel 입력 0개, `MATERIAL_NOT_SUPPORTED_FOR_ANALYSIS` 35개
- `UNKNOWN_MATERIAL`: 0개

PDF 파일명과 내부 물질을 함께 전달해도 최종 분석종류가 `페놀`로 유지됨을 확인했다.

## 동일 PDF 배치 교체 검증

- 동일 PDF를 먼저 `메탄올A` 배치로 저장한 뒤 기존 교체 기능으로 `페놀` 재추출 결과를 저장했다.
- 신규 중복 배치를 만들지 않고 기존 batch ID를 유지했다.
- DB에는 배치 1개만 남고 analysis_type, Sample, Peak가 모두 `페놀` 결과로 교체됐다.
- 원본 file_hash는 동일하게 유지됐다.

## 자동 테스트

- 파일명 우선: 페놀 + MeOH, 메탄올A + DMF, 초산 + Formic acid, IPA + 2-BTOH
- 파일명 판별 불가 시 내부 물질 기반 fallback
- UI 추출 완료 처리에서 페놀 배치 및 배치 코드 유지
- 실제 페놀 PDF의 비대상 Methanol 제외 및 동일 PDF 원자적 배치 교체

## 결과

- 변경 관련 테스트: 45 passed, 2 skipped
- 전체 테스트: 194 collected, 163 passed, 31 skipped, 0 failed
- 참고: pytest cache 디렉터리 쓰기 권한 경고 1건은 테스트 결과에 영향 없음
