# 0527 focused row-arena test review

This handoff contains tests only. The production candidate remains unapplied;
root owns baseline and candidate builds, tests, captures, and retention.
OLE2/OOXML remains the active priority and ODF stays deferred.

## Bound source and draft

The tests are based on accepted revision
`67028ab6037ae6eef15af92a0d540285c3c5362c`. The frozen 0526 production draft
is `row-primary-arena.patch`, SHA-256
`28269e6051f9ee9453ccfd590a0f89269c400529d973a84f0ac908eb6730c5e9`.
Its candidate source hashes are:

| File | Base SHA-256 | Draft SHA-256 |
| --- | --- | --- |
| `crates/litchi-xlsx/src/raw/worksheet/edit/codec/snapshot/model.rs` | `71e23d40e8ef116ae833b51fb77c11067461fc6d5fb69f05fb2bfbe61644a0bd` | `fea4369470a24600451bc96e7992e8a346b1c61a689ebc57d8e400da001ce736` |
| `crates/litchi-xlsx/src/raw/worksheet/edit/codec/snapshot/scan.rs` | `137a317c696e65043027007fbd04edf47f8407563bf0799b852c9be0d54996f8` | `883ee4dcd9d848185401278665c631970712453bb6b5a73010228da3204a96a1` |
| `crates/litchi-xlsx/src/raw/worksheet/edit/codec/snapshot/write/sheet_data.rs` | `71bf1f126205c97b9ca964c1c6f927481becf2195833488470b2823266b426a5` | `4d9841a90d7b76f946bc33f394abb9f35586179952873e9447c19c7d79a64064` |

## Patch contents

`tests.patch` (SHA-256 `b9c3e20595217f5e5941d297744edbaa6c60a8c264f76fb906b19d13786d7c48`)
adds four baseline-compatible focused tests:

- `ordinary_and_provenance_writers_match_for_nonsemantic_multi_primary_bytes`
  calls both raw writer paths directly, proves nonempty omission provenance,
  compares exact unchanged B1/A3 bytes, and verifies all repeated primary
  spans and interleaved comment/`extLst` bytes are handled. It intentionally
  makes no semantic-parser claim for the duplicate-primary fixture.
- `valid_formula_value_provenance_forces_omission_and_matches_complete_parse`
  uses a separate source-backed-valid `<f>+<v>` fixture with prefixes,
  comments, an unchanged B1 owner, and an untouched row. It forces omission,
  compares ordinary/provenance bytes, and uses the complete parser as the
  semantic oracle across the retained store fields.
- `style_only_multi_primary_and_empty_owner_rewrite_is_lossless` checks the
  style-only branch, duplicate primaries, opaque `extLst` bytes, empty B1,
  and an empty row by exact raw slices only.
- `snapshot_scan_preserves_malformed_error_order_after_prior_primary_spans`
  reruns the existing malformed attribute cases after a populated multi-span
  row and compares typed and display errors with the legacy pipeline.

`arena-only-tests.patch` (SHA-256
`8ebb971e50c7968356154e77cf0828c09f868f324fbc0ab6e7e619f5b64799d6`) is
candidate-only and must be applied after `row-primary-arena.patch` (and may be
applied after `tests.patch`). It adds:

- `snapshot_scan_resolves_each_row_arena_range_to_its_own_payload_markers`,
  which resolves per-cell ranges through each row arena and checks exact
  cross-row marker bytes, multiple primaries, prefix aliases, empty cells, and
  empty rows; and
- `snapshot_writer_rejects_primary_range_outside_owning_row`, which exercises
  the draft's checked `Range` lookup and requires the typed
  `Error::Invalid("worksheet cell primary span range exceeds its owning row")`
  refusal.

The candidate-only patch is separate so the baseline lane never sees fields
that do not exist before the production change. Both patches use focused
existing modules and add no production behavior.

## Validation performed

Only patch construction, static diff checks, and `git apply --check` were used;
no Rust build, test, benchmark, capture, or production-source edit was run for
this handoff. The temporary source copies were under
`/tmp/litchi-goal-0527-tests` (a symlink to the owned minimal working copy
`/home/zhuhe/litchi-goal-0527-tests-work` while `/tmp` user quota was full).
They are safe to remove after root records this review.
