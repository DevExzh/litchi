# 0019 — Owned eager DOCX read projection

Date: 2026-10-08. Status: projection migrated; W2 ownership work continues.

## Change

`Package::document()` now returns `Document`, with no package lifetime
parameter. The read-only projection retains a shared physical OPC snapshot;
it remains usable after the package is dropped and observes its original
content and relationships after later package edits. Source-backed DOCX
already had an owned projection and is unchanged.

The eager facade stores its physical owner in a private copy-on-write holder.
Creating a projection shares that graph. A legacy mutable path detaches before
writing, using the existing OPC structural clone and payload sharing rules.
The save rollback guard retains the old graph by reference count and restores
it on failure. Clean metadata slots no longer request mutable graph access,
so an unchanged save keeps the graph shared and the output byte-identical.
Dirty slots retain their existing stage/commit behavior and failure retry.

`DocumentPart` can retain an owned main-part snapshot; its raw borrowed
constructor remains a secondary low-level form. Its `part()` borrow is now
tied to the part wrapper. The semantic document uses the owned form, copies
only main-part metadata/relationships, and shares built-in payload allocations.
Related parts resolve through the retained OPC graph. Existing custom `Part`
implementations still control their own clone and interior-mutation behavior;
no new guarantee is made for arbitrary custom implementations.

Style and raw-part views borrow the owned document rather than the package.
Two test callers now bind the document before borrowing styles. The ops
Markdown adapter and typed test helpers no longer name a document lifetime.
An unnecessary lifetime coupling in header image resolution was removed.
The runnable document example demonstrates reading after package drop.

## Evidence

The ownership regression tests in `paragraph_index_reuse` establish:

- `Document: Send + Sync + 'static`.
- Two projections and their source package share the same physical graph.
- Successful publication detaches the new graph; the old projection retains
  its original main-document content and related core-properties payload.
- The projection survives dropping both the package and another projection.
- Failed raw edits and failed saves leave the old projection and original
  graph intact, and the legacy writer remains retryable.
- An unchanged save preserves exact source bytes and graph pointer identity.

- `cargo test -p litchi-docx --all-features --tests`: 1,986 passed
  (1,264 unit and 722 integration); one existing external-corpus test ignored.
- `cargo check -p litchi-docx --all-features --lib --tests --examples`: passed.
- `cargo clippy -p litchi-docx --all-features --lib --no-deps -- -D warnings`:
  passed.
- Warning-denied DOCX doctests: 80 passed; 31 existing ignored examples.
- `cargo test -p litchi-ops --all-features --lib`: 388 passed.
- `cargo check -p litchi -p litchi-py --all-features --lib`: passed.
- Warning-denied rustdoc glossary gate: 3 profiles, 6 banned-name checks,
  0 complete transaction families; passed (includes minimal/full DOCX).
- Crate-boundary gate: passed for 66 workspace packages and 267 internal
  dependency declarations; 11 existing explicit debt items remain.
- Focused formatting and `git diff --check`: passed.

Read-only review found no correctness blocker. Its recommendations to decouple
header image lifetimes and guard clean metadata staging were implemented.
The initial test compilation also caught the style-view lifetime callers and
a test-only error-conversion mismatch; both were fixed before the passing run.

The first broad integration run found missing external SVG fixtures, not an
assertion against the changed behavior. Temporary copies from the base
checkout's `3rdparty` directory supplied the two existing fixture paths:

| Fixture | Bytes | SHA-256 |
|---|---:|---|
| Open-XML-SDK `TestFiles/svg.docx` | 74104 | `1ce06ef88f89c5e14d1d22cd8e4b3a5f70951e7bd53e19053d4942873333d21d` |
| LibreOffice `ooxmlexport/data/tdf164835_nonDummyLineHeight.docx` | 302210 | `569572dc2ec5dfee9334504ae331482ebb8188b80a0d3c93d44b8884cda29283` |

No fixture bytes or generated build artifacts are committed. The base checkout
was read only. No new worktree was created.

## Limits and remaining work

This step does not yet make the public package immutable. Existing mutable
writers, scoped mutations, and publication methods remain to migrate into
package-bound edits. `document::Commit` still carries a main-part snapshot;
`into_package`, immutable `apply`, the canonical limit/error policies and the
complete family gate remain required. Record 0018's other W2 tasks and the
all-format unified-operations objective remain open.

Graph identity tests establish structural sharing, not latency budgets. A
projection still clones the main part's metadata/relationships; a mutation of
a shared physical owner still structurally clones the OPC graph. No measured
performance improvement or completed performance acceptance is claimed. Paired
benchmarks remain required for the completed ownership wave.

## Cleanup

Removed this turn's generated build output, root lockfile, temporary fixture
copies and test/check logs. No temporary worktrees were created. Pre-existing
independent worktrees and tracked standalone lockfiles are unchanged.
