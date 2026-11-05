# 0197 — Immutable XLS package ownership

## Change

`litchi_xls::Package` retains either the eager parsed reader or the existing
positional CFB/semantic owner. `Package::workbook()` returns an owned, cheaply
cloned read-only `Workbook`; worksheet handles outlive their parent projection.
The previous generic reader is now `raw::Reader`, with no root compatibility
alias. Native consumers and the performance harness use that explicit boundary.
The facade delegates its existing read routes to the native package.

Byte and native path ingress remain eager, preserving BIFF5 and password-to-open
coverage. The facade filesystem route adopts its lazy owner without reading
payloads or rebuilding CFB. Explicit reader materialization retains typed errors
and checks source identity even on cache hits. Metadata caches are published
after their final source fence. Eager metadata keeps the prior facade soft-failure
behavior; all-stream physical preflight precedes its access. A repeated catalog
preflight preserves metadata repair ordering without a second CFB parse.

Exact `to_bytes`, `write_to`, and atomic `save` preserve source bytes. Signed
artifacts can be published exactly, while focused editing refuses signatures.
Unlocked encrypted sources refuse ordinary publication; explicit source export
returns the original ciphertext. Publication never passes an opened workbook
through the fresh authoring Writer. Atomic save checks source identity before
staging and immediately before publication.

The on-demand focused edit adapter intersects retained eager input/entry/aggregate
bounds and CFB admission defaults, or retained positional input and shared CFB
bounds, with requested edit limits. Authorized facade signature removal uses the
same retained profile through an explicit raw transformed-bytes handoff.
`SharedOleFile::limits()` exposes its selected profile without reparsing.

## Verification

The combined XLS, XLS/encryption facade, and CFB suite passes 2,768 tests and
doctests, with three ignored. New acceptance tests cover exact output, projection
lifetime, lazy adoption/read counts and CFB identity, cached source fences,
signed/encrypted refusal, atomic publication rollback, and narrow edit profiles.
Existing corpus comparisons cover both facade routes and tolerance/metadata
behavior. All seven owner acceptance tests also pass against the final fence and
limit changes. Warning-denied all-target Clippy and rustdoc pass. The XLS glossary
passes all 12 banned names. The boundary audit passes for 66 packages and 268
internal dependency declarations, retaining 11 explicit debt items. Performance
harness library/binary consumers compile with XLS source attribution enabled.
Scoped rustfmt and diff checks pass.

The rewritten base `32ad64e320` remains an ancestor. No iWork crate changes or
performance acceptance are claimed. Repository lockfiles remain unchanged;
Cargo uses an external copied lockfile with the cached cc 1.6.0 dependency.

## Remaining scope

This is an ownership migration, not completed XLS primary transactions. Package
Edit/Commit/Patch/Diagnostics, Builder creation, canonical worksheet catalog and
query coverage, unified XLS CRUD/publication, positional OpenOptions unification,
performance acceptance, broader Tier 2/Tier 3, Python/CLI and full-goal acceptance
remain open. Focused edit snapshots and raw reader escapes are explicit interim
boundaries. Limits describe physical/operation dimensions, not total memory.
