# 0363 — macOS native verification queue (before W3)

Owner decision D1 (2026-10-26): routes migrated out of `litchi-iwa` whose host
version has native evidence are proven by the Tier-B differential (record
0363 §4). Every other migrated route is listed here and must pass a native
open, save, close and reopen in Pages, Numbers or Keynote on macOS before W3
deletes the host. Each wave appends its rows; the macOS run records the result
and the artifact SHA-256.

| Wave | Format | Route | Reason queued | Fixture | Result | Artifact SHA-256 |
|---|---|---|---|---|---|---|
| W0-C | Numbers | `litchi_numbers::Package::edit_table_merges` | merge and unmerge writes have no native receipt (the receipt covers a native-authored merge read); host delegates, tier A (record 0369) | `test-data/iwork/numbers/table-merges-native.numbers`: merge A3:B4, unmerge B11:C12 | pending | - |
| W0-C | Pages | `litchi_pages::Package::edit_body_table_merges` | merge and unmerge writes have no native receipt (the receipt covers a native-authored merge read); host delegates, tier A (record 0369) | `test-data/iwork/pages/body-table-merges-native.pages`: merge A2:A3, unmerge C4:D4 | pending | - |
| W0-C | Keynote | `litchi_keynote::Package::edit_slide_table_merges` | merge and unmerge writes have no native receipt (the receipt covers a native-authored merge read); host delegates, tier A (record 0369) | `test-data/iwork/keynote/slide-table-merges-native.key`: merge A2:A3, unmerge B4:C5 | pending | - |
| W1-A | Numbers | `litchi_numbers::Package::edit_table_topology` | row and column insertion: no host native evidence, and the host oracle itself changed (record 0370 host defect fixes: (1) tile `numrows` now keeps the native row-count convention instead of the row extent, (2) a column insertion now splices the pre-BNC cell-offset table with the BNC one, (3) a header/footer insertion now moves the table formula owner's total/body extents); tier A on `basic.numbers` (numbers_table_insert_row_parity, numbers_table_insert_column_parity), tier B on 140 corpus routes (numbers_table_topology_corpus_parity). W3 must not delete the host before this run confirms the corrected output opens, saves and reopens cleanly | `test-data/iwork/numbers/basic.numbers`: insert body row 1, insert body column 1; `custom-number-native.numbers`: insert header row 0, insert header column 0, insert footer row 0 | pending | - |
| W1-A | Numbers | host `numbers::Draft::remove_table_row` (still host-only, not migrated) | host output changed by record 0370 defect fix (1): row removal writes tile `numrows` in the native row-count convention | `test-data/iwork/numbers/basic.numbers`: remove body row 0 | pending | - |
