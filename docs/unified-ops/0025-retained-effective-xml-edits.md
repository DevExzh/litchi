# 0025 — Immutable XML edits of retained effective candidates

Date: 2026-10-08. Status: chained source-Part publication; canonical source
DOCX owner migration remains open.

## Change

`PreparedTopology::edit_xml(&self, partname, callback)` issues a checked splice
transaction against the current effective XML and returns a new retained
candidate. The callback receives a fresh logical revision and its transaction;
the owner checks their identity again after the callback so replacing the
transaction with a foreign one cannot authorize another source. Callback or
validation failure leaves the original owner unchanged.

The result is flat against the original physical archive. It replaces the
selected physical member's final overlay, preserving all other final overlays,
additions, omissions, relationships and content types. It does not publish or
reopen an intermediate ZIP, or retain a chain of parent candidates. Logical
no-ops return a share of the current candidate. When all physical replacements
are reversed and no omissions/additions remain, publication copies the exact
original artifact.

Effective Part records now retain source XML provenance. Unchanged records,
relationship graphs, content-type maps and added-member names share immutable
allocations. Shared retention owns additions, appended entries, transfer guards
and their reservations. Candidate-local catalog/overlay vectors have separate
reservations; new Part metadata and payload reservations follow their retaining
owners. Retention charges its backing vectors and names when taking custody.
Shared object bounds remain conservative; they are not allocator telemetry.

Part and archive totals are retained from the existing topology validator.
Each revision updates them against the effective payload and checks aggregate
and entry limits before publishing a candidate. Operation counts, metadata
sizes and output bounds use checked arithmetic. Base and transfer source
fences and execution checks remain active before/after edits and at publication.

## Verification

- Full-feature OPC unit/integration run after the accounting changes: 1,077
  passed, one existing ignored test.
- Final XML publication target: 25 passed, including one additional aggregate
  archive-limit regression (1,078 distinct passing OPC tests across these runs).
- New coverage includes chaining, forks, no-op retention, exact-artifact inverse,
  source-owner drop, preservation of additions/removals/relationships, callback
  rollback, swapped transactions, source changes during callbacks, exact aggregate
  part/archive limits, exhausted memory, cancellation and 32 sequential revisions
  with stable retained memory followed by final-owner release.
- Final XML/prepared-topology targets: 58 passed after adding the transfer
  vector and Arc header allowance identified in final review.
- Existing transferred-precompressed test now edits the effective candidate,
  drops the parent, publishes, then verifies foreign-source-change refusal.
- DOCX managed document/paragraph targets: 52 passed. XLSX form-control scalar
  lifecycle: 32 passed.
- Warning-denied OPC library Clippy and nine runnable/compile-fail doctests:
  passed after splitting catalog and retained reservations.
- Full-feature facade and Python library compilation: passed.
- Glossary rustdoc: three profiles, 14 banned names, zero complete transaction
  families; passed.
- Crate-boundary gate: 66 packages, 267 internal dependency declarations,
  11 existing explicit debt items; passed.
- Focused formatting and `git diff --check`: passed.

Read-only review identified missing aggregate-limit checks and reservation
separation concerns during development. Explicit aggregate checks and tests,
retained backing-allocation charges, candidate-local vector reservations and
checked count arithmetic address those findings.

## Remaining work and limits

This entry handles existing physical XML Parts whose current payload is either
unchanged source or retains a checked source XML token. Added Parts and authored
replacement payloads without such provenance fail closed; no physical original
or lineage is fabricated for them. General topology changes against a retained
candidate and logical provenance for added/authored XML remain required.

The existing topology planner cannot change an existing physical Part's content
type independently. A future general topology-apply path must validate against
the effective content-type map when it introduces that operation, rather than
reuse the physical catalog's classification.

This is an OPC publication primitive, not the canonical source DOCX transaction
family. Source `Package::edit` still needs integration with retained candidates,
owned projections, immutable `apply`, and package-owning commits. Full package
lineage, snapshot history/composition, `LimitKind`, remaining format waves,
Python/CLI, acceptance and performance work remain required by the original goal.

The implementation shares payloads and graphs, but each changed candidate still
copies catalog handle slots and final overlay metadata. Resolving a Part scans
the effective catalog; this does not yet establish the changed-closure cost bar.
Persistent/indexed catalog storage and paired measurements remain follow-ups.
There is no timing or allocation-improvement claim, and no completed performance
gate or W2 milestone is claimed.

## Cleanup

Removed generated build output, root Cargo.lock and temporary check logs after
recording results. No new worktree or fixture copy was created. Base checkout,
iWork files, existing worktrees and standalone tracked lockfiles were preserved.
