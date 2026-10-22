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

## Validation receipt

The isolated worktree was `/home/zhuhe/code/litchi-worktrees/uo2-replay-error-audit`
at `a5c83d3a2f90e5e1a3bba1889b44d9d6a22f8418` before this layout follow-up.
The focused lane reused
`/home/zhuhe/code/litchi-worktrees/.coord/zip-gate-repair/target`, ran with one
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
tests, and the complete `package::source_append::tests` filter passed all 16
source-append tests.  Direct rustfmt on the changed Rust file and
`git diff --check` passed.

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
* source-backed ODF common publication errors that already preserve
  `SourceChanged` and are mapped by the common owner error; and
* genuinely unsupported content, malformed XML/wire data, resource limits,
  cancellation, and source I/O errors unrelated to freshness.

The XLSB stale-source path belongs to W4b and is excluded here.  Legacy
non-ODF findings are recorded for the separate legacy wave.  No performance
claim is made by this record; the coordinator should run the focused tests and
full gate in its existing capped build lane.
