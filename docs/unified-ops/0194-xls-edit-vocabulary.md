# 0194 — XLS cell-edit vocabulary

## Change

The existing XLS cell-values engine exposes a concrete `Edit`; the
`Transaction` type/alias and duplicate `Snapshot::transaction()` entry are
removed. Internal semantic replay and transfer planning, native consumers,
and the numeric measurement harness use the canonical entry.

`Commit::into_parts()` returns the focused `(Snapshot, Patch)` pair;
diagnostics remain available separately before consumption. Exact and
semantic patches use `is_noop()`. Publication diagnostics expose `changed()`
using an exact comparison of the committed before/after bytes, including
structural/resource edits that do not count as scalar cell changes. This adds
one byte comparison on changed publication; no performance improvement or
complete performance acceptance is claimed.

The XLS glossary profile rejects the removed spellings. This is a focused
engine migration, not a declaration that the complete transaction family is
canonical: `Snapshot` and `Commit::snapshot()` still need integration into the
immutable `Package` owner and read-only `Workbook` projection. The glossary
therefore does not register the family as fully conformant.

## Verification

The final native XLS and XLS/encryption facade suites pass 2,178 tests and
doctests, with one ignored. Existing scalar, structural, resource, durable
replay, exact-no-op, source-backed numeric, signature and encryption coverage
remains green. The no-op regression now checks diagnostics before consuming
the pair, and changed scalar/structural cases assert `Diagnostics::changed()`.

Final warning-denied all-target Clippy passes. The semantic patch's public
operation count retains `len()`; a narrowly scoped Clippy expectation explains
why its canonical no-op method is `is_noop()` instead of `is_empty()`.
Warning-denied rustdoc and the XLS glossary profile pass (four banned
spellings). The glossary verifier's 21 Python tests pass. The performance
harness library/binaries and native-resave XLS consumer compile. The crate
boundary audit passes for 66 packages and 268 declarations, with 11 existing
debt items. Scoped rustfmt and diff checks pass. Root and performance lockfiles
remain unchanged; builds use an external copied lockfile with cached cc 1.6.0.
The rewritten base remains an ancestor of this branch. No iWork crate changes
or performance acceptance claim are included.

## Next required integration

Before package integration, preserve ingress policies through mutation:
`cell_values::Edit::commit` reconstructs its `PackageEditor` with
`Limits::default()`, and `Patch::apply` reopens through `Snapshot::from_bytes`.
The current `Inner` does not retain the object/CFB profiles accepted by the
configured constructors. Source-backed publication has analogous default
reopen paths. Tightened bounds must survive those transitions, including
inverse/replay; this is a prerequisite to claiming the package policy contract.
[Record 0195](0195-xls-retained-limits.md) addresses these retained profiles.

The native cell engine already has guarded scalar/structural changes, exact
patch replay, and source-backed numeric publication. The public reader still
has distinct eager and positional paths, with different BIFF version coverage.
A canonical package must preserve those read paths, signed/encrypted exact
publication, source fences, and retained limits while delegating edits to this
engine. Creation must publish a package through a builder; ordinary byte and
atomic-path publication must live on the package. The facade then needs typed
sheet/cell projections, scalar writes, transaction coalescing, history and save
routes over that native owner.

No new XLS facade write capability or performance acceptance is claimed by
this migration. Tier 2 read/write completion, broader DOC CRUD, tier 3,
Python/CLI, and all remaining goal gates stay open.
