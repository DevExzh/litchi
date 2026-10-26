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
| W1-B | Pages | `litchi_pages::Package::edit_body_drawable_properties` | image, movie, audio and shape frame, flip, Original Size, drawable-property and movie-playback writes have no native receipt (only Pages audio playback does: `pages/audio-playback-native-resaved.pages`); tier A, the leaf output is byte-identical to the host's (record 0371) | `test-data/iwork/pages/image-adjustments-native.pages` image 0 (geometry, vertical flip, Original Size after a 10×10 resize, properties); source-built movie and audio (`Draft::create_with_text` + `add_body_movie`/`add_body_audio`: movie geometry, horizontal flip, Original Size, properties, playback; audio position, properties; source-built rectangle (`add_body_rectangle`): geometry, vertical flip, properties) | pending | - |
| W1-B | Pages | `litchi_pages::Package::body_drawable_properties` | read route of the same family; the source-built movie and audio inputs have no native counterpart (record 0371) | as above | pending | - |
| W1-B | Pages | `litchi_pages::Package::edit_body_text` | whole-body set and clear writes have no native receipt; tier A, byte-identical to the host's `set_body_text`/`clear_body` (record 0371) | `test-data/iwork/pages/basic.pages`: set to "Replaced body — ünïcode"; clear | pending | - |
| W1-B | Pages | `litchi_pages::Package::edit_section_text` | the host's body-absolute `replace_body_text` maps to a section-relative replace; no native receipt; tier A (record 0371) | `test-data/iwork/pages/body-sections-unicode.pages`: section 2, UTF-16 1..3 → "Ω inserted" | pending | - |
