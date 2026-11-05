# 0036 — Typed OPC source-limit errors

OPC exports SourceLimitKind in place of ReadResource. DOCX, PPTX, XLSX and XLSB
also re-export this physical dimension kind beside SourceLimits. There are no
compatibility aliases. The shared error variants now use the ADR 0034 shapes:

- `OpcError::Limit { kind, observed, maximum }`
- `OpcError::InvalidLimit { kind, value, maximum }`

Archive error translation, deferred decode-error reproduction, source-backed
checks, leaf callers, unified error classification, facade tests and active
fuzz/performance consumers use the renamed types and fields. Exceeded limits
still classify as LimitExceeded; invalid profiles remain InvalidArgument.
Diagnostics contain the typed dimension and numeric facts. Display uses
“source limit” and invalid profiles report their allowed positive range.

## Invalid-profile bounds

Input bytes report `usize::MAX - 1`, matching the existing addressability rule.
Unrestricted u64 byte ceilings report `u64::MAX`; usize count ceilings report
`usize::MAX`. Source-part splice dimensions that already reject u64::MAX report
`u64::MAX - 1`. No accepted range or default was widened or narrowed.

Cross-field validation reports the offending dimension and the actual related
bound: relationship parts and parts are bounded by archive members; members are
bounded by total archive entries; relationship-target bytes are bounded by XML
attribute bytes. The members/entries case previously pointed at the lower
entries value. It now correctly reports ArchiveMembers, its value and the
maximum allowed by total entries. Validation order remains unchanged.

Three new tests check zero/over-maximum refusals, exact upper-bound setter
acceptance without unrelated cross-field constraints, all four inconsistent
profiles, splice finite bounds, canonical Debug field order and deferred error
copying. Existing ingress boundary tests migrate their typed assertions.

## Enforcement and verification

OPC profiles ban both retired variant names and the old fields at their exact
variant scopes. OOXML profiles also ban ReadResource (XLSB uses the existing
root-only scope to preserve unrelated nested vocabulary). No complete focused
transaction family is registered by this change.

- Full-feature OPC/DOCX/operations unit and integration suites: 3,505 passed,
  two existing ignored (1,098 OPC, 2,019 DOCX, 388 operations).
- Targeted PPTX limits, XLSB source-backed and facade ingress suites: 81 passed.
  The facade covers DOCX/PPTX/XLSX/XLSB limits and ZIP detection arbitration.
- Warning-denied OPC/DOCX/PPTX/XLSX/XLSB doctests: 115 passed,
  44 existing ignored.
- Warning-denied six-crate library Clippy: passed.
- Workspace all-feature library compilation, including facade/Python: passed.
- Glossary: 11 profiles, 62 banned-name rules, zero complete families; passed.
  Checker tests: 21 passed.
- Crate boundaries: 66 packages, 267 internal dependency declarations and
  11 existing explicit debt items; passed.
- Active perf-baseline bins and OPC fuzz targets compile.
- Deterministic OPC smoke: one valid seed, six malformed seeds, 22 main
  boundaries, six additional boundaries and two payload reopen checks passed.
  No seeds were written. This is bounded smoke coverage, not a fuzz campaign.
- Independent read-only review found no concrete error-mapping or bound issue.
- Changed Rust formatting and diff whitespace checks: passed.
- Rewritten base remains 32ad64e320 and is an ancestor of this branch.

## Remaining work

This completes the physical profile's kind and primary error vocabulary, not
the full limits contract. Dimension/setter names, OpenOptions envelopes,
semantic limits, operation-specific splice errors and other format families
still need migration. Specialized DOCX transactions, other unified milestones,
Python/CLI and full acceptance remain part of the active goal. No performance
improvement is claimed.

## Cleanup

The six temporary test fixtures and their hashes are listed in record 0035;
they were copied from the base checkout for validation and removed afterward.
Workspace build output, generated root/OPC-fuzz lockfiles and temporary logs
were removed. No worktree was created; base, iWork and pre-existing worktrees
were preserved.
