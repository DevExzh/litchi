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
