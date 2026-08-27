# Change 0331: XLSB external-link lossless provenance and allocation hardening

## Scope

This change hardens the source-backed XLSB external-link path while preserving
wire content that the semantic model does not own. It is an incremental
prerequisite for full source-backed external metadata parity, not that parity
itself.

## Implementation

- Workbook, DDE, and OLE `BrtSupNameBits` payload provenance is retained
  privately. Writers patch only the modeled bit masks; ignored and reserved
  bits survive source-backed edits, while authored entries use canonical bits.
- Provenance is excluded from semantic equality. A source-backed no-op is byte
  exact, including the original external-cache and future opaque frames.
- Opaque external-cache/future frames are copied and emitted in their original
  relative order for safe nonstructural edits. The fixture uses valid cache
  record payloads (table start, row and increasing-column cell records, and
  table end) plus a separate unknown frame, positioned between the modeled
  records.
- Formula lengths are rejected before parser-owned byte copying. Source and
  opaque-record copies use owned `Vec` storage inside the `Arc`-shared
  `SourceState`; `Patch` independently owns its before/after `Vec` images.
  Source and opaque copies, record growth, and output reservations are
  fallible, and output size is checked before output allocation.
- The DDE relationship-id setter returns a typed refusal instead of mutating
  the DDE server field. `SheetRange` has private validated fields, removing
  the former public panic path.

## Evidence

Validation was serialized with `CARGO_BUILD_JOBS=1` and the isolated
`/dev/shm/litchi-0331-target` Cargo target:

- Focused external-link tests: 11/11 passed.
- Full library tests: 548/548 passed.
- Integration tests: 121 passed, with exactly
  `checked_in_unique_standard_drawing_corpus_transfers_every_anchor` skipped.
- Strict Clippy, `RUSTDOCFLAGS="-D warnings"` rustdoc, the minimal facade
  XLSB feature check, crate-scoped formatting, and the diff check passed.

The workspace-wide formatting check was not used as proof because the
committed unrelated `crates/litchi/src/presentation/prs.rs` has an existing
formatting mismatch; the crate-scoped check passed. These results make no
performance, RSS, or OOM claim.

## Residuals

- Aggregate external-link semantic budgets are not implemented yet.
- `HashSet::with_capacity` remains an ordinary bounded allocation.
- Exact noncanonical record-header provenance is not claimed; only retained
  opaque frames and source-backed no-op bytes receive that preservation claim.
- Ordinal opaque anchors can change semantic adjacency after structural edits.
- Full source-backed external metadata parity remains future work.
- Patch and semantic-clone allocation scope is not fully hardened yet.
- The known drawing-corpus test remains skipped as stated above.
