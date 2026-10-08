# 0028 — Materialize the retained effective OPC graph

`PreparedTopology::to_opc_package()` is the explicit owning conversion needed
by the future source-backed DOCX effective adapter. It reads the effective
catalog: retained replacements and additions, removed Parts, effective root
and Part relationships, and each Part's effective content type. Unchanged
payloads use one serial decoder session; retained replacement allocations are
shared. The owning result survives dropping the candidate and later changes to
its physical source.

The operation checks base and transfer source/context fences before, during and
after conversion, including error paths. A stale source wins over a concurrent
payload-read failure. Managed bases and managed XML/precompressed donor sources
are refused before cold reads: materialization cannot detach their admissions
into a bare owning graph. No-op candidates use the established owning converter.
The result retains ingress repair reports, non-Part metadata and borrowed-ingress
signature policy. This conversion is a whole-package operation and does not
preserve ZIP framing; retained publication remains the physical preservation path.

## Verification

- Full-feature OPC unit/integration tests: 1,092 passed, one existing ignored.
- Warning-denied OPC doctests: ten passed.
- Warning-denied OPC library Clippy: passed.
- Full-feature facade and Python library compilation: passed.
- Glossary rustdoc: three profiles, 16 banned names, zero complete families;
  passed.
- Crate boundaries: 66 packages, 267 internal dependency declarations,
  11 existing explicit debt items; passed.
- Focused formatting and diff whitespace checks: passed.

Tests cover graph edits and root redirection,
shared replacement storage, result lifetime, no-op/changed cold Parts, stale
base and donor sources, cancellation, managed XML and precompressed transfers,
read-error precedence, signature policy and ingress repair reports.

Read-only review initially suspected an unchanged XML replacement could retain
an unfenced managed donor. A concrete regression attempt was rejected by the
existing destination-lineage check before preparation. Replacement tokens must
belong to the base lineage/context; cross-source XML is allowed as an addition.
The speculative planner change was reverted, and legal managed XML additions
were added to the conversion regression coverage. The reviewer found no
remaining effective-catalog, source-fence or reservation issue.

## Next integration boundary

Source-backed DOCX still uses its physical OPC package. The next owner migration
must route reads through one effective adapter and call this converter when a
candidate exists. It must use logical revision keys (or per-owner invalidation)
for main-document admission and paragraph caches. Source XML proof failures
must not fall back to physical bytes. Specialized publishers that have not
migrated must refuse effective candidates rather than publish the original.

Review also found that preparing an empty topology currently disables opt-in
read-ahead. DOCX must not eagerly prepare one at open without resolving that
behavior. Its physical package accessor, unified signature/unsign path and
`DocxOwner::finish` fences need explicit handling of effective candidates.

This record does not complete canonical source DOCX Edit/Commit, the glossary
family gate, all-format CRUD, Python/CLI work or the original goal's acceptance
and performance requirements. No edit-path allocation or throughput claim is
made.

## Cleanup

Removed workspace build output, generated root Cargo.lock and temporary check
logs. No fixture or worktree was created. The base checkout, iWork crates and
pre-existing worktrees were preserved.
