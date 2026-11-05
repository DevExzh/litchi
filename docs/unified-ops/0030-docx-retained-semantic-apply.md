# 0030 — Immutable semantic apply on retained DOCX owners

Source-backed `Package::apply(&document::Patch)` returns a new owner without
serializing or reopening a ZIP. Main-document snapshots and raw XML edits now
read the effective graph. Sequential patches preserve unrelated graph changes,
old owners remain readable, and inverse patches can restore the original archive
byte for byte. Exact no-ops retain the current owner, including signed sources.
Changed signed sources still refuse publication authority.

The apply boundary verifies exact source bytes and physical identity, rejects
paragraph transfers requiring dependency publication, and inspects the effective
settings relationship and payload for protection/tracked-revision policy. It
rechecks base and donor source/context fences on success and failure. Managed
snapshots can reuse matching retained XML proofs; original signed no-op graphs
use the existing physical snapshot fallback. Opening a physical owner still does
not eagerly prepare an empty topology or change its read-ahead policy.

OPC `PreparedTopology::apply_source_xml` checks both expected and target tokens
against the original physical Part and checks expected bytes against the current
effective Part. It preserves other graph changes and refuses foreign, stale or
unproven targets. `SourceXmlPart::with_replacement_xml` validates a complete XML
document under source-read rules, retaining declarations, BOMs and producer
whitespace. It shares an incoming Arc payload, accounts managed output capacity,
and retains original physical authority without a chain of intermediate revisions.
Malformed XML, DTDs, unsupported encodings, limits and cancellation remain refused.
These are OPC provenance primitives, not DOCX semantic/protection authorization.
As in 0029, explicit `with_prepared_topology` adoption and its publication or
owning conversion remain raw OPC ingress, equivalent to opening caller-supplied
package bytes. They do not establish semantic edit authorization. Protection
validation described here belongs to the semantic `apply` boundary.

A regression exposed revision edits dropping unmanaged source identity. Direct
revision dispositions and replay now use the identity-preserving rewrite helper.
The existing managed refusal stays: that rewrite cannot retain managed admission.

Raw snapshots currently cannot retain independent donor contexts. Review found
that copying a managed donor's main XML would detach those obligations. Snapshot
and edit capture therefore refuse any effective topology with managed donors,
including when the base itself is managed. Read-only Document views keep their
retained data handle and remain supported. OPC exposes managed-base/donor status;
its owning conversion shares the same detection logic.

## Verification

- `cargo test -p litchi-docx -p litchi-opc --all-features --lib --tests`:
  3,100 passed, two existing ignored.
- Final semantic/topology targets after donor-retention review: 18 and 47
  passed; managed signed retained-no-op regression passed.
- Warning-denied DOCX/OPC doctests, including final rerun: 93 passed,
  31 existing ignored.
- Warning-denied DOCX/OPC library Clippy, including final rerun: passed.
- Unified operations library: 388 passed.
- Full-feature facade and Python library compilation: passed.
- Glossary rustdoc: three profiles, 16 banned names, zero complete families;
  passed.
- Crate boundaries: 66 packages, 267 internal dependency declarations,
  11 existing explicit debt items; passed.
- Formatting and diff whitespace checks: passed.

New tests cover managed/unmanaged chaining, stale and foreign preconditions,
exact inverse publication, graph retention, effective protection settings,
signed no-op/changed application, declaration/BOM/revision spelling, malformed
replacement XML, cancellation and managed reservation release. Temporary SVG
fixtures are the two files with sizes/hashes recorded in 0019.

## Remaining work and limits

This is not completion of the canonical source transaction family: `edit()`
still returns a raw `document::snapshot::Edit`, and its commit does not own a
package. Package-owning source Edit/Commit and specialized effective reads and
publishers remain required. Added/unproven main-document Parts and dependency
transfers are refused by this apply route. Existing specialized physical routes
continue to refuse effective owners.

Preconditions use physical identity plus exact effective bytes, not a distinct
logical version that rejects every ABA sequence. Unmanaged snapshot construction
still copies/scans XML; the full-replacement fallback validates XML again. No
performance improvement or budget compliance is claimed. Other formats,
Python/CLI completion, transaction-family gates, performance and acceptance
requirements remain active.

## Cleanup

Removed temporary SVG fixtures, workspace build output, generated root Cargo.lock
and check logs. No new worktree was created; the base checkout, iWork files and
pre-existing worktrees were preserved.
