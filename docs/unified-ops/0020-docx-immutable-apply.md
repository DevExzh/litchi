# 0020 — Cheap eager DOCX package shares and immutable apply

Date: 2026-10-08. Status: eager apply migrated; W2 package commits remain open.

## Change

The eager `Package::apply(&document::Patch)` now returns a new `Package` and
retains its source unchanged on both success and failure. The public mutating
`apply_document_patch` was removed, not aliased. The existing main-part proof,
transfer-graph checks, signature/protection/MCE checks and publication audit
run through a private engine on the new share. Legacy commit/history methods
use that same private engine until their own migration; there is no second
validation implementation.

`Package: Clone + Send + Sync` shares physical graph, metadata slots and retained
legacy writer state. Scalar security/font configuration and dirty flags are
copied; paragraph/admission caches remain allocation-keyed shared caches.
Mutation detaches through a private copy-on-write holder. The shared OOXML
metadata values and DOCX authoring tree gained structural `Clone` implementations
so this detachment preserves pending authoring state instead of dropping it or
reopening a source. Cloning the package itself does not traverse that tree.

An unchanged patch retains all shared state. The save rollback guard checks
whether the writer is modified before requesting mutable access, so saving an
unchanged share does not clone the writer. Failed saves preserve retryable dirty
state only on the affected clone. As before, custom OPC `Part` implementations
control their own clone/interior-mutation semantics; built-in payloads share
immutable allocations.

The error context for the removed patch method now uses operation `apply`.
No security default is relaxed: an encrypted source still refuses this plaintext
OPC transaction path even for a no-op, and plain output still requires explicit
authorization. Exact no-ops on signed packages remain byte-identical; changed
patches refuse until explicit unsigning. Unsigning one clone preserves the
signature state of its source.

## Verification

New and updated tests cover:

- `Package: Clone + Send + Sync + 'static` and pointer sharing of physical,
  core/custom metadata and writer state on no-op apply/save.
- Successful application and inverse publication retain readable source and
  target packages; stale application leaves the target unchanged.
- The shared admission memo is safe when alternating changed/source snapshots.
- A cloned dirty writer and cloned dirty core/custom metadata detach on edits;
  failure retains both clones' dirty state, retry succeeds, and saving the
  changed clone does not clear or change the source's state/output.
- Core-property replacement and removal on clones preserve the source slot.
- Signed no-op byte equality, changed signed refusal, explicit unsign isolation,
  encrypted-source clone policy, and refusal without writing plaintext bytes.

- `cargo check -p litchi-docx --all-features --lib --tests --examples`: passed.
- DOCX and shared OOXML full-feature unit/integration tests: 2,443 passed
  (1,989 DOCX and 454 shared OOXML); one existing corpus test ignored.
- Warning-denied Clippy for both touched libraries: passed.
- Warning-denied doctests for both touched libraries: 81 passed; 31 existing
  examples remain ignored.
- `cargo test -p litchi-ops --all-features --lib`: 388 passed.
- `cargo check -p litchi -p litchi-py --all-features --lib`: passed.
- Crate-boundary gate: passed for 66 workspace packages and 267 internal
  dependency declarations; 11 existing explicit debt items remain.
- Focused formatting and `git diff --check`: passed.
- Warning-denied glossary rustdoc: 3 profiles, 8 banned-name checks,
  0 complete transaction families; passed.

Read-only review identified the transitive authoring Clone requirements and
requested sidecar/security isolation coverage. Those requirements are implemented
and covered by the passing checks. No iWork crate was modified.

The non-ignored native SVG tests use the two temporary external fixture copies
identified by path, size and SHA-256 in record 0019. The base checkout is read
only; fixture copies and build artifacts are not committed.

## Remaining work and limits

This is an immutable primary publication entry, not completion of ADR 0034's
owner model. Public legacy writers and scoped mutation methods still exist.
`document::Edit`/`Commit` still publish a main-document `Snapshot`; they must be
bound to the package and expose the canonical package/patch/diagnostics triad.
The source-backed owner still uses stream publishers and requires a retained
package publication representation before it can use the same contract.
Limit/error vocabulary and the remaining per-format waves are still required.
The overall all-format unified read/write, Python, CLI, parity and performance
acceptance goal is unchanged.

Pointer checks prove sharing and isolation, not timing budgets. A changed eager
publication still structurally clones its OPC graph; mutation of shared legacy
authoring state clones that tree. Opening now allocates shared metadata/writer
holders. No performance improvement or completed performance gate is claimed;
paired timing and allocation evidence remains required for the ownership wave.

## Cleanup

Removed generated build output, root lockfile, temporary external fixture copies
and this turn's test/check logs. No new worktrees were created. Pre-existing
independent worktrees and tracked standalone lockfiles remain unchanged.
