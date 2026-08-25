# Peak Table continuation verification - 2026-08-25

## Scope

- Parser: `LabSolutionsParser`
- Real PDF: `TEST/혼유 120-167 병합.pdf`
- Branch baseline: `main` at `a9b81f2`
- Python: project `.venv` Python 3.12

## Real PDF verification

The available PDF was rendered and pages 35 through 42 were visually inspected.

- Sample `2` is created once and contains Peak 1 through 82 from pages 39 and 40.
- Sample `3` is created once and contains Peak 1 through 65 from pages 41 and 42.
- Pages 40 and 42 do not create separate Sample records in SQLite.
- Parsing completes without `Peak Table을 찾을 수 없습니다` on continuation pages.
- A continuation page containing only an eight-column `Total` row is accepted and adds no duplicate Sample or Peak.
- A headerless continuation table containing numeric Peak rows is appended to the previous Sample.
- New Sample Information still starts a new Sample; Blank, STD, recovery, and QC classification is unchanged.

The supplied file differs from the originally reported page pair for Sample `126`:
page 35 is `125B`, page 36 is `126`, and page 37 is `126B`. The parser correctly
keeps `126` and recovery blank `126B` as separate samples.

## Test commands and results

### New targeted tests

```text
python -m unittest tests.unit.test_sample_classification tests.regression.test_labsolutions_sample_pdf.Honyu120167ContinuationRegressionTests -v
Ran 13 tests in 12.126s
OK
```

### PDF regression and shared-folder compatibility tests

```text
python -m unittest tests.unit.test_shared_folder_service tests.regression.test_labsolutions_sample_pdf -v
Ran 28 tests in 44.970s
OK
```

### Full test suite

```text
python -m unittest discover -s tests -p test_*.py -v
Ran 138 tests in 60.528s
OK
```
