# 0381 — W2-D: relocate the focused Numbers comment-read tests

- Status: implementation slice on `uo2/iwa-w2d`, cut from
  `feat/unified-ops-api` at `c650ba53f34faf334cf256fbc98fdb4b83d64335`.
  Build and validation gates are intentionally pending the coordinator's
  build-hold clearance; this record is not a merge or W3 deletion approval.
- Date: 2026-10-28.
- Plan: [0363](0363-iwa-monolith-exit-plan.md), W2-D; parity policy
  [0369](0369-iwa-exit-parity-harness-and-ledger.md). The historical
  `79e209d0c9` rename is a donor fragment only; it is not merged here.

## 1. Scope and ownership

This slice relocates exactly one host integration target containing 17 focused
Numbers comment-read tests:

| Before | Destination | Owner |
|---|---|---|
| `crates/litchi-iwa/tests/numbers_focused_comment_reads.rs` | `crates/litchi-numbers/tests/numbers_focused_comment_reads.rs` | W2-D relocation, reviewed by the Numbers owner |
| `crates/litchi-numbers/tests/support/numbers_comment_fixture.rs` | unchanged | existing Numbers test-support owner |
| `crates/litchi-numbers/tests/fixtures/comment-edit-root.numbers` | unchanged | existing Numbers fixture owner |

The compatibility target
`crates/litchi-iwa/tests/numbers_comment_compatibility_reads.rs`, every
`crates/litchi-iwa/tests/parity/**` target, all other host tests/examples, the
ledger, and the macOS queue are unchanged.

The only source adaptations are owner-relative paths:

- the shared fixture module now uses `#[path = "support/..."]`;
- the permanent fixture is included from `/tests/fixtures/...` under the
  Numbers crate manifest directory;
- the boundary ratchet now locates the focused read target in
  `crates/litchi-numbers/tests/`.

The test keeps the current selector-first `At` API. No public API, production
code, compatibility seam, dispatch route, or capability ledger row changed.

## 2. Coverage and overlap disposition

The moved target compares selector/A1 projections with an independent wire
oracle that reads comment storage and author records directly. It retains
coverage for shared roots/replies, multitable and segmented lists, malformed
references, missing field-info headers, unknown wire data, source-order
metadata, finite timestamps and signed zero, optional authors, bounded
metadata, repeated public IDs, and exact-source preservation.

The existing leaf targets were reviewed before moving the file:

- `table_data_list_reader_integration.rs` covers lower-level table-data-list
  and mutation/rejection paths using its own synthetic package builder;
- `table_cell_comment_reply.rs` covers reply mutation, inverse, ownership,
  and transaction limits;
- `table_cell_comment_metadata_native.rs`, `comment_authors.rs`, and
  `comment_corpus.rs` cover native metadata/authors and corpus mutation.

Those are adjacent contracts, not exact duplicates of the moved focused
wire-oracle read cases. No assertions were removed. Keeping the independent
oracle avoids turning a leaf-only self-round-trip into host parity evidence.

## 3. Evidence boundary and fixtures

This is test ownership relocation only. It claims no native application
acceptance, no new parity tier, and no migration-ledger progress. The existing
6,178-byte `comment-edit-root.numbers` fixture remains in the leaf crate; the
other focused graphs are generated in memory by the existing support module.
No fixture was copied, no provenance changed, and no `3rdparty/` path was
introduced. The host remains available as the differential oracle elsewhere.

## 4. Deferred verification

Per the coordinator build hold, no Cargo command, test, formatter, Clippy,
boundary checker, fuzz lock, or benchmark has been run on this branch yet.
After clearance, the targeted Numbers integration test and the normal scoped
format/Clippy, rustdoc/glossary, boundary, ledger, and affected fuzz-lock
checks must run on the final integrated trunk. The historical parity/cost
targets remain in `litchi-iwa` and are not part of this relocation gate.

No native queue entry is added or cleared by this slice.
