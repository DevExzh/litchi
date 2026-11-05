# 0362rh — ODF replay base and source-freshness errors

This slice audits and repairs the public ODF replay paths on coordinator
integration base `532cd8634b786827f1a3191f21ce6bc4889a2e69`.  A recorded patch
whose base bytes do not match the recipient is a `PatchConflict` and maps to
`invalid_input`.  A positional source that changes after open is a
`SourceChanged` condition and maps to `io_error`.  A well-formed artifact or
route that is outside a family remains a typed refusal.  The redaction audit
(item 22) is outside this slice.

## Audit and repairs

The following public callers were reachable through the package facades or
format-owned durable/history APIs.  They previously collapsed a stale patch
base into `InvalidFormat` (or a format-local generic invalid result).  The
implementation now returns `PatchConflict`, and focused tests match the typed
variant.

| Public route | Previous failure | Repair and regression |
|---|---|---|
| ODT `Package` durable replay, sealed replay, and `History::commit` (`crates/litchi-odt/src/transaction.rs:4161,4915,2856`) | `InvalidFormat` from `durable_source_mismatch` or `stale_history_error` | Both helpers return `litchi_core::Error::PatchConflict`; `tests/internal/packaged_transactions.rs` checks durable, sealed, and stale-history replay. |
| ODT embedded-resource GC plan (`crates/litchi-odt/src/package/resource_gc.rs:195`) | Generic invalid source mismatch | Returns `PatchConflict`; `tests/embedded_resource_gc.rs` checks forward and inverse stale recipients. |
| ODG joined apply, adjacent `Patch::then`, and durable replay (`crates/litchi-odg/src/package/snapshot.rs:1215,4933,5696`) | Generic invalid source/lineage result | Each exact-base mismatch returns `PatchConflict`; `tests/package_semantics.rs` covers joined replay, composition, and durable stale input. |
| ODM `History::record`, durable replay, and sealed replay (`crates/litchi-odm/src/transaction.rs:1327,1422,1480`) | Generic invalid source result | All three return `PatchConflict`; `tests/capability_transactions.rs` covers stale history and durable replay, and `tests/advanced_transactions.rs` covers durable replay. |
| ODB durable and sealed replay (`crates/litchi-odb/src/authoring/durable.rs:407`) | Generic invalid source result | The shared source mismatch wrapper returns `PatchConflict`; `tests/advanced_transactions.rs` checks both directions. |
| ODB composed sub-edits (`crates/litchi-odb/src/authoring/composition.rs:74,85`) | Generic invalid lineage/source result | Both exact-base guards return `PatchConflict`; the test constructs a public `SubEdit` with a foreign payload to exercise the source guard. |
| ODI `SemanticPatch::apply_flat/apply_package` (`crates/litchi-odi/src/semantic.rs:667,675`) | One `InvalidFormat` branch conflated stale bytes and wrong artifact kind | Same-kind byte mismatch returns `PatchConflict`; a wrong artifact kind returns `Refused(UnsupportedContent)` with fixed, content-free text. `tests/semantic_planning.rs` covers both routes. |
| ODF Formula `Patch::apply` and `History::push` (`crates/litchi-odf-formula/src/facade/mod.rs:1041,1335`) | Generic invalid source/chain result | Both return `PatchConflict`; `tests/fifth_wave.rs` checks stale history application and duplicate chain insertion. |

Private lower-level patch decoders and replay helpers were not renamed solely
for message consistency when the public owner already checks the same base.
Malformed target content, invalid wire vocabulary, limits, signed/protected
sources, and unsupported content remain their existing typed errors.

## ODP tail-append source wrapper

`crates/litchi-odp/src/package/source_append.rs` has a format-owned
`Error::SourceChanged { expected, observed, progress, source }` variant.  Its
display is limited to opaque source-version tokens and progress, while
`Error::source()` exposes an optional typed diagnostic through the boxed
`SourceChangedSource` wrapper.  Helper conversions preserve freshness from:

* `check_source`, `effective`, `source_version`, and
  `member_materialized_size` during apply, prepare, and commit;
* common insertion-plan publication errors; and
* the common source-content scan's publication and verified-member errors.

Patch-structure and byte-proof failures continue to use the existing
`Error::PatchConflict { operation, reason }`.  The source-append unit test uses
a bounded `ReadAt` adapter with a revision counter and verifies that a changed
source is reported as `SourceChanged` by both replay and preparation.  A
focused regression also verifies `Untouched` progress for a stale-before-output
failure, `Complete` progress for a stale-after-output failure, and freshness
retention for a callback-scoped source-member error.

The follow-up keeps `SourceChanged` as the primary classification while
retaining the typed core, source-member, or publication diagnostic through
`SourceChangedSource`.  This preserves a secondary callback failure when a
source verification failure wins precedence; `Error::source()` exposes the
retained typed chain.

The retained wrapper is heap-indirected on the error path.  On the capped
lane, `size_of::<litchi_odp::tail_append::Error>()` is 104 bytes with
alignment 8, equal to the 104-byte/8-aligned pre-follow-up base at
`7dbfa217fa0c338a175e290f5b3f278be888aab9`; the unboxed wrapper would have
made the largest result error at least 144 bytes and failed strict Clippy's
`result_large_err` check.  `SourceChangedSource` itself is 96 bytes and is
allocated only when retaining a stale diagnostic, so successful operations
keep the original result ABI and perform no diagnostic allocation.

## Common ODF publication boundary

The previous common conversion flattened `SourceContentPublicationError` and
`SourceContentInsertionError` into `Core` strings or a bare
`SourceChanged`.  That lost sink progress and the replay callback's typed
source-member diagnostic on `SourceBackedPresentation::save`, and it also
turned a stale source discovered in a replay callback into a generic core
error on `effective`.  The common owner error now has boxed
`SourceContentPublication` and `SourceContentInsertion` variants.  Boxing
keeps the ordinary package error layout bounded; the allocation is made only
when the source-backed operation fails.  `save` retains the complete typed
insertion chain, including progress, replay, source-member and secondary
callback errors.

`SourceBackedPresentation::effective_with_odf_error` is used by the ODP edit
paths before the ordinary core-facing `effective` conversion.  Its conversion
to `Error::SourceChanged` preserves the insertion progress and typed chain
when a replay callback reports a stale source.  The core-facing result still
has only the shared `SourceChanged` carrier, so a caller that requests the
read-only projection receives the correct `io_error` classification while the
typed chain remains available through the source-backed ODF/save route.

The common `Error` remains 48 bytes and 8-aligned before and after this
repair.  The preceding cached build measured 48 bytes, and the repaired
build measured 48 bytes; `SourceContentInsertionError` is 96 bytes and
`SourceContentPublicationError` is 64 bytes.  The new common variants hold
boxes, so their larger typed payloads do not enlarge the successful-path
package error value.

The ODP facade now registers `tail_append::Error` as a leaf source and
classifies its `SourceChanged` and `PatchConflict` variants as `io_error` and
`invalid_input`.  A public source-backed open/edit/commit/replay regression
exercises both outcomes against a revisioned `ReadAt` source and an unrelated
recipient, rather than constructing only the leaf variants.  Its common
publication classifier also checks an exact typed `SourceChanged` callback
before classifying an archive content fault, sink, or transport wrapper; an
arbitrary callback error does not replace the primary archive or sink code.
The facade regression constructs an archive content fault with a retained stale
callback and verifies that the ODF wrapper still reports `io_error`.

## Validation receipt

The isolated worktree was `/home/zhuhe/code/litchi-worktrees/uo2-replay-error-audit`
at `a5c83d3a2f90e5e1a3bba1889b44d9d6a22f8418` before this layout follow-up.
The focused lane reused
`(deleted-coord-artifact)`, ran with one
Cargo job and one test thread, and remained below the 20 GiB target cap.
Results:

* ODT `embedded_resource_gc`: 12 passed;
* ODG `package_semantics`: 24 passed;
* ODM `advanced_transactions` and `capability_transactions`: 9 and 8 passed;
* ODB `advanced_transactions`: 20 passed;
* ODI `semantic_planning`: 14 passed;
* ODF Formula `fifth_wave`: 5 passed; and
* ODP source-append progress regression: 1 passed.

The follow-up's ODP stale-source filter passed all 4 freshness and diagnostic
tests, and the complete `package::source_append::tests` filter passed all 18
source-append tests.  The common conversion regression and the public ODP
facade stale/conflict regression also passed.  The repository pins
`rustc 1.95.0` in `rust-toolchain.toml`.  The Clippy results below are
exploratory host checks from `rustc 1.98.1`; this sparse checkout had
temporarily omitted the tracked toolchain file, so these results are not the
required 1.95 gate result.  Under 1.98.1, strict Clippy passed for
`litchi-odf-common` and `litchi-odp`.  The strict `litchi-ops` command still
encountered two pre-existing Rust 1.98 lint failures in `litchi-sheet-text`
(`byte_char_slices` and `chunks_exact_to_as_chunks`) and two pre-existing
`litchi-ops` failures (`chunks_exact_to_as_chunks` and `question_mark`); the
affected package passed with those four baseline lints explicitly allowed.
The coordinator must rerun the affected strict checks with the pinned 1.95.0
toolchain.
Direct rustfmt on all changed Rust files and `git diff --check` passed.

Strict `cargo clippy --lib -- -D warnings` passed for ODT, ODG, ODM, ODB,
ODI, ODF Formula, and ODP.  Broader all-target checks were limited by absent
corpus fixtures in this intentionally partial audit checkout; the affected
integration targets above were run after temporarily restoring only their
small tracked fixtures, which were removed before commit.

## False positives and boundaries

The audit searched public apply/apply-family callers and their source-backed
wrappers.  These findings were intentionally left unchanged:

* private `Patch::apply` functions that are reached only after a public exact
  source precondition; their malformed-target errors are not public stale-base
  diagnostics;
* source-backed ODF common errors whose conversion is now handled by the
  typed publication/insertion boundary above; the core-facing projection
  intentionally retains only the shared freshness carrier; and
* genuinely unsupported content, malformed XML/wire data, resource limits,
  cancellation, and source I/O errors unrelated to freshness.

The XLSB stale-source path belongs to W4b and is excluded here.  Legacy
non-ODF findings are recorded for the separate legacy wave.  No performance
claim is made by this record; the coordinator should run the focused tests and
full gate in its existing capped build lane.
