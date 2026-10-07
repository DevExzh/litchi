# 0952 source and contract review

Reviewed against checkout HEAD `2d6833be51615486ef3c2a83d4791fd2a3f223c4`.
This is a read-only source and retained-metadata review. It makes no native-time,
allocation, speedup, or adoption claim and introduces no source change.

## Retained 0951 route

The measured case is `xlsx_source_backed_cell_values_one_percent_edit_save`.
`tools/perf-baseline/src/lib.rs:44748-44752` opens it through
`SourceBackedEditor::from_read_at_with_limits_and_cache_limits`, which supplies
no `ExecutionContext`. `crates/litchi-opc/src/source_backed.rs:5537-5552`
therefore creates `PartCache` with no `write_session`; the publication reaches
the serial `index.write_to` arm at `:12191-12203`.

The retained reports expose these exact corpus facts:

- medium: 17 archive members, 4,231,168 uncompressed payload bytes, six source
  payload materializations, twelve untouched members;
- dense-sparse: 17 archive members, 4,265,472 uncompressed payload bytes, six
  materializations, twelve untouched members;
- both reports mark `deterministic_task_count` as `not_applicable` because the
  route has no explicit execution context.

The metadata does not expose the number of regenerated Deflate tasks, their
logical sizes, or a wave width. The 12-member untouched oracle must not be
converted into a Deflate-task count.

## Compression path and invariants

`crates/litchi-xlsx/src/cell_values/source.rs:599-650` builds the topology
plan and delegates publication to OPC. In
`crates/litchi-opc/src/source_backed.rs:9596-9685`, changed payloads are
materialized as shared logical bytes after the original-first XML audit; exact
byte no-ops skip the audit and remain source-preserving. At
`:12086-12115`, changed entries retain the source member's compression method,
while unchanged entries are `Copy` actions.

`crates/soapberry-zip/src/preserve.rs:2361-2373` consumes a scheduled result
or calls `generated_entry` once. The Deflate branch at `:2543-2565` constructs
one `ReusableDeflateState` and performs one logical-payload compression pass.
The zlib-rs `longest_match` symbol is therefore work for a changed Deflate
member; this source path shows no same-member duplicate compression.

The direct precompressed branch at
`crates/litchi-opc/src/source_backed.rs:10001-10021` is restricted to an
authorized addition payload. Edited worksheet replacements use the logical
shared-payload path above, so reusing the old compressed bytes would publish
stale data and violate CRC/size and source-publication semantics.

## Candidate boundary

The next bounded action is evidence collection, not a source rewrite: first
run the authorized native scoped profile to confirm the codec/runtime-dispatch
attribution. Then, if the exact corpus has multiple balanced Deflate tasks,
qualify the already implemented managed route on the same updates using
`from_read_at_with_limits_and_cache_limits_and_execution_context` and record
Deflate task count, total/largest/remainder bytes, selected wave width, and
correctness/refusal outcomes. `PreservationPlan::deflate_wave` already applies
worker, remainder, and per-task floors (`preserve.rs:430-455`), and OPC charges
CPU tasks/worker permits before compression (`source_backed.rs:12141-12151`).
If the task census is one dominant member or yields width one, there is no
bounded parallel candidate for this corpus.

This keeps unchanged-member passthrough, source freshness, complete original
XML auditing, no-op behavior, typed limits/refusals, deterministic ordering,
and cancellation boundaries intact. It does not propose a compression-level
change, guessed precompressed reuse, or ambient/global parallelism.

## Prior evidence that remains binding

- `docs/performance/0618-zip-writer-deflate-state-reuse.md:186-208` rejected
  sharing one compressor across members after +14.1% cycles, +9.9% instructions,
  and 271 extra minor faults per publication.
- `docs/performance/0618-zip-writer-deflate-state-reuse.md:210-222` priced the
  mini-archive parse/framing round trip at 0.15% per regenerated member and
  retained its structural checks.
- `docs/performance/0662-parallel-changed-member-deflate.md:150-199` binds
  explicit session-owned scheduling, ordered errors, cancellation, and no
  ambient/global pool.
- `docs/performance/0578-zip-passthrough-is-already-bounded.md:119-133`
  confirms production edited members use `RegeneratedEntry::new_shared` plus
  Deflate while unchanged members use `Copy`; its precompressed-token route is
  for authorized cross-document additions, not edited unchanged members.
