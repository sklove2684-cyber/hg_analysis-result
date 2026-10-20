# Heavy metal recovery-only exports (2026-10-20)

## Changes

- Search every sheet for BLANK plus 검출량/검출농도 headers and a matching
  element with three rows each for 저/중/고. No fixed sheet names, injection,
  correction, average or recovery-percent headers are required.
- Na and K match NaOH and KOH labels. General elements match Element-number
  labels (their suffixes do not determine row positions).
- Extract PDF G# and persist it through DB save/replace/read. Existing DBs migrate
  atomically, retaining old values with an unknown group; re-extraction supplies G#.
- Match explicitly numbered Excel regions to PDF G#. Ambiguous/missing regions
  block generation. If the same region exists on several sheets, use the saved
  active sheet only when it provides a unique candidate; otherwise block.
- Reject formula targets, merged targets, duplicate target cells and incomplete
  12-value element/region sets.
- Heavy metal creation writes only the selected inputs. It preserves workbook.xml
  and skips COM recalculation/save even if global recalculation was enabled, so
  formula caches and unrelated sheets remain byte-for-byte unchanged. Formulas
  remain intact for Excel to calculate normally when opened.

## Actual files

Input template: `0. 틀-수산화.xlsx`; PDF: `Na,K 321-352.pdf`.
Both were found in the existing desktop network-share mirror's 중금속 folder.
The source template has 상반기 틀 / 230106 / intensity; its saved active sheet
is 230106. Both of the first two sheets contain regions 1 and 2.

PDF G#: 1. Selected sheet: 230106, selected region: 1.

| Element | Label | BLANK targets | Low/mid/high targets |
| --- | --- | --- | --- |
| Na → NaOH | B61 | E61:E63 | F61:F69 |
| K → KOH | B70 | E70:E72 | F70:F78 |

Total: 6 BLANK inputs, 18 recovery inputs, 24 writes, 0 preview errors,
0 TEMPLATE_PROFILE_MISMATCH. No ordinary sample inputs.

Actual extraction → DB save → preview → CreateExcelExportService → output
succeeded, with all 24 values verified. The ZIP part list is unchanged and only
xl/worksheets/sheet2.xml differs. After removing the 24 authorized value cells,
the whole remaining sheet XML is identical. Every other ZIP part (including the
other two worksheets, workbook, styles, charts, drawings and shared strings) is
byte-for-byte identical. The workbook validator reports no structural errors.
Region 2 remains unchanged.

Native Excel verification: the output opened successfully through the normal
read-only Workbooks.Open path (no recovery mode requested). The recovery range
calculated in memory and exported to a visual preview successfully. Closing
without saving retained the exact output SHA. No repair/recovery operation was
used. Visual inspection confirmed the original layout/merges and the populated
24 inputs. Artifact Tool's renderer terminated with a native access violation;
Excel's native range-to-PDF rendering was used for visual verification instead.

Final generated workbook: outputs/20261020/수산화_회수율_검증결과.xlsx.

## Tests

New coverage includes arbitrary sheet names, absent unrelated sheets/calculation
headers, both input header aliases, general metals on arbitrary sheets, G#=2,
missing/ambiguous groups, active-sheet disambiguation, formula and merged targets,
empty PDF recovery sets, multi-group PDF/DB roundtrips, legacy DB migration and
actual file creation with exact package/cell preservation.

Final full suite: `python -m unittest discover -s tests` — 316 discovered,
277 passed, 39 skipped for unavailable actual-data fixtures, zero failures/errors
(103.753 seconds). The actual hydroxide test ran and passed. `git diff --check`
also passed.
