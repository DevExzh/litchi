# 0392 — W1-G comments gate repair on the 295b candidate

- **Base:** `295b421bcd3d9ec0b0b9208e4039de22623685bc`
- **Implementation commits:** `5fc2f905a0` and `56020e9214` on
  `uo2/iwa-comments-gate-295b-repair`.
- **Status:** scoped Numbers/shared-engine repair is complete and remains
  unintegrated. Coordinator review, Pages/Keynote clear-safety proof,
  final integrated performance, and native acceptance remain pending.

## Scope and root causes

The 295b full gate exposed five regressions in the existing Numbers comment
route. The repair stays in `litchi-iwa-table/src/comments_native.rs`; the
Numbers facade continues to publish the selector-first API and the shared
root operation.

1. Numbers comment roots live in the DataList component, not necessarily in
   the selected table-model component. Existing-list admission now validates
   the selected roots through the source census without rejecting this valid
   multi-component profile.
2. `ListKeys::Reuse` applies to a newly-created Numbers list. An existing
   list still allocates from its source `next_list_id` watermark, avoiding a
   key-2 collision on populated lists.
3. A single registered author is reused. When a multi-author Numbers graph
   has no existing `litchi-iwa` author, creation allocates a new generated
   author and appends it to the existing registry. Payload-only native author
   registries retain their header shape.
4. Existing table-list headers vary between aggregate-only, generic
   `FieldInfo`, and keyed `FieldInfo` mirrors. List metadata updates now
   preserve those source shapes while checking selected reference counts and
   removing only the selected object mirror.
5. A destructive clear admits exactly one type-3056 message with matching
   `ArchiveInfo` metadata, rejects empty table-info candidates, and rejects
   unknown top-level comment payload fields before mutation. Unknown fields
   remain supported by non-destructive read/text paths.

No public selector, convenience API, ledger disposition, or native evidence
claim was added by this repair. Unsupported reply authoring and native
acceptance remain as previously documented in the W1-G records and queue.

## Regression and focused verification

The original gate failures now pass without weakening the existing assertions:

| Check | Result |
|---|---:|
| `comment_authors` generated-author reuse/new-author cases | 2 passed |
| `comment_corpus` corpus and workbook round trips | 2 passed |
| `table_data_list_reader_integration` | 25 passed |
| `table_cell_comment_reply` | 31 passed |
| `table_cell_comment_metadata_native` | 5 passed |
| `numbers_focused_comment_reads` | 17 passed |

The clear-loss controls retain their exact typed outcomes and source bytes:

- noncanonical metadata framing → `InvalidSource`;
- an opaque sibling message in the selected storage object →
  `UnsupportedDependency`;
- an empty type-6000 table-info candidate → `InvalidSource`;
- unknown top-level type-3056 payload field → typed refusal.

The valid root-only clear, inverse, candidate reopen, locality, no-op, and
reply-bearing refusal cases remain green. Host-backed Pages/Keynote clear
baseline and negative proof is tracked separately by the coordinator before
integration.

## Scoped gates

Passed on the clean implementation source (release/debug settings are stated
per command):

- `CARGO_BUILD_JOBS=2 CARGO_INCREMENTAL=0 CARGO_PROFILE_DEV_DEBUG=0
  CARGO_PROFILE_TEST_DEBUG=0 cargo fmt --all -- --check`;
- `cargo clippy --offline -p litchi-iwa-table --all-targets --all-features
  -- -D warnings`;
- `cargo clippy --offline -p litchi-numbers --all-targets --all-features
  -- -D warnings`;
- `RUSTDOCFLAGS=-Dwarnings cargo doc --offline --no-deps --all-features
  -p litchi-iwa-table -p litchi-numbers`;
- `python3 tools/check_crate_boundaries.py`;
- `python3 tools/iwa_exit_ledger.py --check`;
- focused `tools.test_check_crate_boundaries`, `tools.test_iwa_exit_ledger`,
  and `tools.test_no_3rdparty_references`;
- locked Numbers and iwa-protos fuzz-workspace `cargo check --bins` checks.

The full workspace glossary check completed against `target/api-glossary` with
`RUSTC_BOOTSTRAP=1`. The affected generated documents were then checked
explicitly with `--profile iwa-table`, `--profile numbers-minimal`, and
`--profile numbers-full`; all three passed. The coordinator should retain the
full-workspace command output/exit receipt with the integrated handoff.

## Evidence boundaries and remaining qualification

This repair has no native E3/E4 evidence and makes no new host/leaf
performance claim. Existing comments performance captures predate this repair
and do not qualify it. The coordinator must run the final host-vs-leaf
instruction-cost matrix on the integrated candidate after the Pages/Keynote
clear proof and any other wave merges, then update `PROGRESS.md` and perform
the complete 53-row gate on one trunk commit.
