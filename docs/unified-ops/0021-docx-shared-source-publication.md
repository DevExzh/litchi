# 0021 — Shared source-backed DOCX owners and borrowing publication

Date: 2026-10-08. Status: sharing prerequisite complete; W2 package commits open.

## Change

Source-backed `Package: Clone + Send + Sync + 'static` now shares the existing
OPC source lineage, index, lazy caches and execution context. DOCX tolerance
reports, property-report deduplication and main-document admission are shared
alongside the existing paragraph-index cache. Cloning does not reopen the
archive, read payloads or duplicate retained payload reservations.

All 32 public package stream publishers now borrow `&self`, including scoped
story, layout, variable, paragraph-copy, sanitization/redaction and SVG paths.
They pass a cheap inner OPC share into its existing consuming publication
engine. Tail-append plan ownership is unchanged. The original package remains
readable and can publish again after success or a sink failure. Redaction
retains the original source and still exposes no inverse API.

The admission memo lock covers checking already-retained main-document bytes,
preventing duplicate admission work between concurrent shares. No source I/O,
callback or publication occurs under that DOCX memo lock. Property repairs are
recorded once across all shares. Existing source-change, cancellation, budget,
protection, signature and XML-publication checks remain in their original
engines; a share does not authorize publishing a stale source.

Borrowing changes the reservation lifetime visible to callers: a retained
package keeps its cached payload reservations. Existing managed tests now drop
the package explicitly before asserting final zero usage, preserving those
assertions. A dedicated test proves cloning adds no reservations and the last
retaining source owner's drop releases its cache.

## Verification

New coverage checks source cloning without payload reads, repeated identical
publication, retry after sink failure, unchanged original reads, source drop
survival, exact no-op bytes and stale-source refusal before output. Concurrent
shares publish the same no-op artifact and aggregate core-property repairs only
once. Managed shares retain and release reservations at the expected lifetime.

- Full-feature DOCX unit/integration coverage: 1,992 passed; one existing corpus
  test ignored. The initial full run exposed one remaining consuming-lifetime
  assertion in `source_backed_managed`; its six-test target passed after adding
  the explicit source drop. All other targets passed in that full run.
- `cargo check -p litchi-docx --all-features --lib --tests --examples`: passed.
- Warning-denied Clippy for the DOCX library: passed.
- Warning-denied DOCX doctests: 81 passed; 31 existing examples ignored.
- Warning-denied glossary rustdoc: 3 profiles, 8 banned-name checks,
  0 complete transaction families; passed.
- `cargo test -p litchi-ops --all-features --lib`: 388 passed.
- `cargo check -p litchi -p litchi-py --all-features --lib`: passed.
- Crate-boundary gate: 66 packages, 267 internal dependency declarations,
  11 existing explicit debt items; passed.
- Focused formatting and `git diff --check`: passed.

Read-only review checked shared source identity, sidecar lifetime, the admission
memo, reservation behavior and all publisher receivers. Its findings were
addressed before final verification. No iWork crate or base-checkout file was
modified. Native SVG tests used the two external fixture copies recorded by
path, size and SHA-256 in record 0019.

## Remaining work and limits

These borrowing stream APIs are transitional. `document::Edit`/`Commit` still
carry a document snapshot and must retain the complete package and expose the
canonical package/patch/diagnostics triad. Source-overlay package publication,
legacy writer removal, limit/error vocabulary and all remaining format waves
remain open. The all-format CRUD, Python, CLI, parity and performance goal is
unchanged; this record does not complete W2 or a full transaction family.

Opening adds three shared sidecar allocations. The tests establish shared
payload ownership and correct accounting, not timing budgets. Paired timing
and allocation evidence remains required for the completed ownership wave.

## Cleanup

Removed generated build output, root lockfile, the two temporary external
fixture copies and this turn's check logs after recording their results. No
new worktree was created. Pre-existing worktrees and tracked standalone
lockfiles remain unchanged.
