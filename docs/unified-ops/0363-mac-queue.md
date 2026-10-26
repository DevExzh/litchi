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
| W1-A | Numbers | `litchi_numbers::Package::edit_table_topology` | row and column insertion: no host native evidence; tier A on `basic.numbers` (numbers_table_insert_row_parity, numbers_table_insert_column_parity), tier B on 140 corpus routes (numbers_table_topology_corpus_parity); the host gained three documented defect fixes (tile `numrows`, pre-BNC offsets, table extents; record 0370) | `test-data/iwork/numbers/basic.numbers`: insert body row 1, insert body column 1; `custom-number-native.numbers`: insert header row 0 | pending | - |
| W0-B | Numbers | `litchi_numbers_wire::formula_graph::next_owner_id` (formula-owner allocation of table duplicate; adopted by the W1-A duplicate route) | output differs from the host by design: the host allocates internal owner identifiers that collide with owner-map entries (a repeated map entry), the leaf allocates above the map (record 0368a; `tests/parity/numbers_owner_allocation.rs`) | `test-data/iwork/numbers/chart-arrangement-native.numbers` (duplicate any table; host map repeats an identifier), `test-data/iwork/numbers/shared-model-native.numbers` (clone primitive) | pending | - |
| W1-D | shared | `litchi_iwa_text_wire::paragraph_style::effective` | storage-scoped character and paragraph property reads (24 properties): no host native evidence; tier A read parity on native fixtures and on host-derived variations (record 0374) | `test-data/iwork/keynote/basic.key`, `keynote/multi-slide.key`, `pages/basic.pages`, `numbers/basic.numbers`: every uniformly styled storage | pending | - |
| W1-D | shared | `litchi_iwa_text_wire::paragraph_style::plan` | storage-scoped property set/reset (24 properties): no host native evidence; derived variations tier B (host UUIDs random), in-place rewrites and resets tier A (record 0374) | same fixtures: the first two uniform storages of each, every property set, rewritten and reset | pending | - |
