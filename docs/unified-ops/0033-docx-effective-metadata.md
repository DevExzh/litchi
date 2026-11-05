# 0033 — Metadata reads from retained DOCX owners

Source-backed `Package::metadata()` now resolves core properties against the
retained effective graph. A package returned by canonical document commit can
read metadata without materializing or reopening the ZIP. Property payload
replacement, relationship redirection, addition and removal are visible, while
older owners keep their original results.

The shared OOXML property reader exposes `read_prepared_with_report`. Its
PreparedPartView graph adapter uses the same inspector and decoder as eager and
physical source reads. The iterator factory is fallible, so effective capture
errors propagate instead of becoming absent properties. Graph iteration remains
metadata-only; only the relationship-selected payload is decoded. External or
wrongly typed targets retain the existing refusal policy. Producer-compatible
graph/XML repairs, dialect handling and absent properties retain the established
shared-reader behavior.

Base and retained donor source/context checks surround graph inspection, payload
capture and decoding, including error returns. The data handle retains donor
reservations during decoding. DOCX performs its final source fence before
merging the repair report. Adoption already creates fresh per-owner report state;
clones retain the existing once-per-owner repair deduplication. Metadata values
remain ordinary owned semantic results under the existing property-reader
contract; this does not change semantic allocation accounting.

## Verification

- Full-feature DOCX/common-OOXML/operations unit and integration tests:
  2,858 passed, one existing ignored (2,015 DOCX, 455 common OOXML,
  388 operations).
- Warning-denied DOCX/common-OOXML doctests: 84 passed, 31 existing ignored.
- Warning-denied DOCX/common-OOXML/operations library Clippy: passed.
- Full-feature facade and Python library compilation: passed.
- Glossary rustdoc: three profiles, 22 banned names, zero complete families;
  passed.
- Crate boundaries: 66 packages, 267 internal dependency declarations,
  11 existing explicit debt items; passed.
- Formatting and diff whitespace checks: passed.
- Rewritten base remains 32ad64e320 and is an ancestor of this branch.

New tests cover canonical main-document commit followed by metadata read,
property replacement and redirection to a newly added Part, removal/absence,
old-owner stability, dialect repair deduplication, no unrelated/main payload
reads during selected metadata access, source changes during property reads,
managed donor cancellation and final reservation release. Shared tests compare
retained/eager results and repair reports for Strict/Transitional relationships,
duplicate properties, wrong content types and external relationships.

## Remaining work

Embedded objects, ink, variables, layout and other specialized reads/transactions
still need effective adapters. This change does not remove those typed refusals
or complete the metadata write family. Caller-selected semantic limits, broader
API/policy migration, other formats, Python/CLI, performance and full acceptance
remain open. No measured performance improvement or complete transaction-family
conformance is claimed.

## Cleanup

Removed the temporary SVG fixtures recorded in 0019, workspace build output,
generated root Cargo.lock and check logs. No new worktree was created. The base
checkout, iWork files and pre-existing worktrees were preserved.
