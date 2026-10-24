# IWA protobuf and source-request inventory dispositions

Status: source and inventory disposition for the 0362 final sweep.  The
inventory rows below were reviewed against the frozen rustdoc/source boundary
at `49c25ef67f62a8bb943e7c6a0bf464e49ae87d07`.  This is a set of exact public
paths and source witnesses, not a crate-wide `internal/proto` exception.

## Canonical names repaired in this leaf

`litchi_iwa_protos::numbers_table_physical_sort_codec::RowMove` is the one
public value for a physical row move.  The former `HeaderRowMove`,
`HeaderRowEdit` and `TileRowEdit` aliases are removed.  Header planning and
rewrite functions now take `&[RowMove]`; the Keynote package adapter, the
codec tests, and the fuzz target use `RowMove` directly.  The value remains a
pair of `u32` local indexes and the rewrite algorithms, validation order,
allocation accounting and wire bytes are unchanged.

`litchi_iwa_protos::numbers_table_cell_storage_codec::ControlCellSpecTableReferenceEdit`
is the one public precondition value for a DataStore reference rewrite.  The
former `CommentStorageTableReferenceEdit` alias is removed.  The four
comment-storage/table-model preparation and execution functions and the
Numbers comment consumer now pass the canonical type.  The field route still
selects DataStore field 19 (`comment_storage_table`) and keeps the same
insert/replace preconditions and source-preserving rewrite behavior.

## Remaining `edit_type` rows: source-backed wire requests

These types are compact values handed to a source-byte codec.  None is a
format `Package`, a read-only projection, or a staged package `Edit`; none has
`commit`, `rollback`, `is_changed`, `snapshot`, or a package publication path.
The inventory may retain the exact row with the stated `wire_request`
disposition, rather than renaming it to a package-family `Edit`.

| Exact public path | Source witness and role |
| --- | --- |
| `litchi_iwa_protos::keynote_media_lifecycle_codec::SlideLifecycleEdit` | `keynote_media_lifecycle_codec.rs:689-759`; borrowed selector-free lists of identifiers. The owner resolves selectors and package witnesses first. `prepare_slide_lifecycle_rewrite` at `1103-1140` retains the source bytes; `rewrite_slide_lifecycle_with_report` at `1151-1158` emits a rewritten payload and report. |
| `litchi_iwa_protos::keynote_media_lifecycle_codec::BuildLifecycleEdit` | `keynote_media_lifecycle_codec.rs:761-773`; one drawable identifier rewrite passed to `rewrite_build_with_report` at `1169-1203`. |
| `litchi_iwa_protos::keynote_media_lifecycle_codec::BuildChunkLifecycleEdit` | `keynote_media_lifecycle_codec.rs:775-805`; optional build/UUID edge rewrites passed to `rewrite_build_chunk_with_report` at `1214-1225` and its source-byte continuation. |
| `litchi_iwa_protos::keynote_media_lifecycle_codec::node_cache::SlideNodeBuildCacheEdit` | `keynote_media_lifecycle_codec/node_cache.rs:93-116`; `#[doc(hidden)]` unit request to invalidate native cache scalars. Preparation at `202-232` retains source bytes and `commit` at `137-176` emits bytes/report. |
| `litchi_iwa_protos::numbers_table_cell_storage_codec::TableDataListEntryRefCountEdit` | `numbers_table_cell_storage_codec.rs:1302-1334`; three scalar fields (key, expected count, replacement count), converted into `TableDataListEntryMutation` at `1404-1430`. |
| `litchi_iwa_protos::numbers_table_cell_storage_codec::HeaderSizeEdit` | `numbers_table_cell_storage_codec.rs:1702-1739`; one source-indexed size mutation. Plans store these values at `1898-1905`; `plan_header_storage_bucket_sizes` and execution validate and rewrite source bytes. |

The Keynote consumers prove the boundary: the slide package passes these
values to the codec (`litchi-keynote/src/package/slide_media_lifecycle.rs:1907-1928`
and `2525-2554`, `slide_audio_creation.rs:1866-1880`, and
`slide_media_lifecycle/graph.rs:933-1020`) and then replaces archive message
bytes.  No codec request publishes a package or creates a public commit.

## Remaining `edit_type` and `transaction_type` rows: protobuf schema

The following exact paths are generated schema messages from
`TSWPArchives.proto`; they are not Rust editing objects.  The source schema
at `crates/litchi-iwa-protos/src/protos/TSWPArchives.proto:574-695` defines
`UndoTransaction` as a persisted native undo payload containing union entries,
storage-action groups, flags and version.  Its nested messages contain wire
fields such as text insertion locations, character indexes, attribute indexes,
paragraph data and object/date references.  `UndoTransactionWrapperArchive`
embeds that payload, and command archives reference it as a field.  The crate
root describes itself at `src/lib.rs:3-9` as the raw generated schema/codegen
boundary, and re-exports the generated `tswp` module at `2532-2536`.

These genuine protobuf names and messages remain unchanged so their schema
vocabulary and wire layout stay stable.  They expose generated fields and
encoding methods only; they do not expose package staging, commit, rollback,
snapshot or output methods.  Each is listed explicitly to avoid a broad
module-level exception:

| Exact public path | Wire role |
| --- | --- |
| `litchi_iwa_protos::tswp::UndoTransaction` | Native undo payload containing repeated union entries and action groups. |
| `litchi_iwa_protos::tswp::undo_transaction::GenericTransaction` | Three generic wire values (`value1`, `value2`, `value3`). |
| `litchi_iwa_protos::tswp::undo_transaction::TextTransaction` | Text insertion location/length and optional deleted string. |
| `litchi_iwa_protos::tswp::undo_transaction::CharIndexTransaction` | Character index/count and optional object reference. |
| `litchi_iwa_protos::tswp::undo_transaction::ReplaceCharIndexTransaction` | Attribute and character indexes. |
| `litchi_iwa_protos::tswp::undo_transaction::AttributeIndexTransaction` | Attribute/character indexes, object reference and string. |
| `litchi_iwa_protos::tswp::undo_transaction::InsertAttributeTransaction` | Attribute index/count, object reference and string. |
| `litchi_iwa_protos::tswp::undo_transaction::InsertNilTransaction` | Attribute index and count. |
| `litchi_iwa_protos::tswp::undo_transaction::CharDeltaTransaction` | Character index, attribute index and signed delta. |
| `litchi_iwa_protos::tswp::undo_transaction::ParagraphDataTransaction` | Attribute index, two paragraph values and character index. |
| `litchi_iwa_protos::tswp::undo_transaction::ObjectDolcTransaction` | Object-DOLC flags and optional object reference. |
| `litchi_iwa_protos::tswp::undo_transaction::CtDateTransaction` | Attribute index and optional date. |
| `litchi_iwa_protos::tswp::undo_transaction::UnionTransaction` | Undo-kind discriminator, table and one nested wire transaction. |

## Decision

After the alias removal, the remaining iwa-protos inventory rows have the
exact dispositions above.  The only source/API repairs in this leaf are the
four public aliases and every Rust, fuzz, test and boundary-tool consumer of
those aliases.  No protobuf schema message is renamed, and no source request
is promoted into a package edit family.  A fresh rustdoc inventory should
therefore contain no `HeaderRowMove`, `HeaderRowEdit`, `TileRowEdit` or
`CommentStorageTableReferenceEdit` path, while the listed wire-request and
wire-schema rows remain traceable to their source witnesses.

## Validation receipt

The committed source patch was checked before commit `91dcf9c0d0` with ten
staged paths (`127 insertions, 48 deletions`), `git diff --cached --check`, the
large-evidence and binary guards, and a source/tool search that found none of
the four removed names.  The locked root commands used the frozen-49
`Cargo.lock` (139,591 bytes, SHA-256
`52ed0cacac15b5e440349bfc0f4aca2665bcf50d59d4321624e596bb410783c6`):
`cargo test --locked --offline -p litchi-iwa-protos --lib --tests` reported
972 library tests passed and every integration binary passed (no aggregate
count is asserted), and `cargo check --locked --offline -p litchi-keynote
-p litchi-numbers --lib` passed.  `python3 -m unittest
tools.test_check_crate_boundaries` passed all 1,144 tests.

The separate fuzz manifest had no pinned lock.  Its offline command,
`cargo check --offline --manifest-path
crates/litchi-iwa-protos/fuzz/Cargo.toml --bin
keynote_table_physical_sort_codec`, resolved its own ignored 22,763-byte lock
(SHA-256
`e1523f0d57989de8b315d44ce639d2ebf7b05919fae030675d4c923c1c04b02b`) and
passed; that lock was not committed.  The capped environment used the
coordinator cargo shim, `LITCHI_MEM_MAX=20G`, `CARGO_BUILD_JOBS=1`,
`RUST_TEST_THREADS=1`, offline mode, and target
`/home/zhuhe/code/litchi-targets/uo2-iwa-protos-alias-repair` (822 MiB after
validation); the temporary directory was
`/home/zhuhe/code/litchi-worktrees/.coord/iwa-protos-alias-repair-tmp`.
Session receipts were 79655, 98489,
35753 and 75652 respectively; no durable command-log hash was retained.  An
earlier sibling-lock attempt (session 79727, lock prefix `6d7895e`) was
SIGINT before any test completed and is diagnostic only, not an acceptance
receipt.  The alias-only declarations preserve the codec layout and core
algorithms, so this leaf introduces no performance-path change.
