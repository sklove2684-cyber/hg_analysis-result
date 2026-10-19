# Recovery set IDs mapped to Excel slots

Recovery samples retain their original names and replicate numbers in storage.
Before mapping peaks, the preview scans all recovery samples in PDF/sample order.
Exactly three distinct positive source IDs must each contain one low, mid and high
sample. Their first appearance determines logical Excel slots 1, 2 and 3.

Examples covered by cell-level tests: 1/2/3, 3/4/5, 7/8/9, 2/4/6, 5/6/7,
10/20/30, and non-sorted first appearance 9/2/7. Missing or duplicate levels
produce RECOVERY_SET_INCOMPLETE; two or four sets produce
RECOVERY_SET_COUNT_MISMATCH. Invalid/nonpositive IDs remain errors. A batch
without recovery samples retains the existing STD/worker-only behavior.
Recovery blank aliases are excluded from this grouping.

The database save/replace/read/preview test verifies that source IDs 3/4/5 remain
unchanged in storage and produce the same target cells and values as 1/2/3.
Local real-PDF regressions compare all preview statuses, target cells and applied
areas before/after renumbering complete 1/2/3 extractions to 3/4/5, 7/8/9 and
2/4/6 for mixture and one-column mixture PDFs.

## Incomplete existing alcohol fixture

`알콜(2) 74-119.pdf` contains 저1/중1/고1, 저2/중2/고2 and 중3/고3;
저3 is absent. Under the requested validation this file must block export.
Its regression now expects RECOVERY_SET_INCOMPLETE and verifies that the
existing worker peak RT selection and cells remain unchanged. Preview counts
are 54 mapped, 485 excluded, one validation issue. No recovery cells are mapped.

## Requested 394–432 file

The user reported `혼유 394-432(수정)@완료.pdf` with source IDs 3/4/5 and
191 mapped, 346 excluded, six issues. These are user-provided baseline figures,
not a locally reproduced measurement. The file is absent from the project TEST
fixtures and no matching batch exists in the inspected local application DB.
Its PDF/Excel paths are required to re-extract, replace that batch, and measure
actual before/after counts. No unrelated live DB batch was replaced.

## Validation

`python -m unittest discover -s tests` on 2026-10-19:
305 discovered, 266 passed, 39 skipped due to unavailable actual-data fixtures,
zero failures/errors (110.511 seconds). Unit tests: 233 passed.
Application data for tests was isolated under runtime-validation/recovery_20261019.
`git diff --check` passed.
