# 0037 — Canonical OPC source dimensions

The shared physical SourceLimits profile now uses the ADR 0034 common nouns.
Getters and consuming builder setters use `max_<dimension>`; typed failures
identify the corresponding SourceLimitKind variant. There are no aliases.

| Previous dimension | Canonical dimension | Kind |
|---|---|---|
| archive_members | members | Members |
| archive_entry_bytes | entry_bytes | EntryBytes |
| archive_total_bytes | aggregate_bytes | AggregateBytes |
| xml_depth | depth | Depth |
| xml_events | events | Events |

InputBytes/input_bytes was already canonical. Defaults, numeric bounds,
validation order and ZIP resource mapping are unchanged. In particular the
shared input ceiling remains 512 MiB and XML depth remains 256. Events now
has the diagnostic label “XML events”; it already applies beyond relationship
parts, including checked content-types, styles, drawing and source XML.

API_GLOSSARY.md registers all 24 OPC dimensions and their meanings. Distinct
bounds retain their names: members excludes directories, archive_total_entries
includes them; aggregate_bytes counts declared ZIP payloads, total_part_bytes
counts materialized OPC payloads. Aggregate relationship XML events remain a
separate dimension. The table does not claim that every semantic reader already
enforces every physical-profile bound.

OPC, OOXML consumers, operations detection, tests, fuzz harnesses and active
performance tools use the new names. Performance records describing this
shared profile also use max_depth/max_events. Separate validation, custom-data,
stream and MCE limits retain their owner-local vocabulary for later migration.
The XLSX custom-data adapter preserves the minimum of the shared and local
limits while using each profile's own methods.

## Enforcement and verification

The minimal/full OPC glossary profiles ban all five old methods under both
SourceLimits and SourceLimitsBuilder, and all five old variants under
SourceLimitKind. Scope is exact so unrelated local profiles are unaffected.
No complete focused transaction family is claimed.

- Full-feature OPC/DOCX/operations unit and integration suites: 3,505 passed,
  two existing ignored (1,098 OPC, 2,019 DOCX, 388 operations).
- Targeted XLSX drawing, pivot and custom-data suites: 141 passed; PPTX
  source-backed cross-copy suite: 58 passed.
- Warning-denied OPC/DOCX/PPTX/XLSX/XLSB doctests: 115 passed,
  44 existing ignored.
- Warning-denied six-crate library Clippy: passed.
- Active perf-baseline bins and the two changed DOCX fuzz targets compile.
- Deterministic OPC smoke: one valid seed, six malformed seeds, 22 main
  boundaries, six additional boundaries and two payload reopen checks passed.
  No seeds were written. This is bounded smoke coverage, not a fuzz campaign.
- Crate boundaries: 66 packages, 267 internal dependency declarations and
  11 existing explicit debt items; passed.
- Changed Rust formatting and diff whitespace checks: passed.
- Workspace all-feature library compilation, including facade/Python: passed.
- Glossary: 11 profiles, 92 banned-name rules, zero complete families; passed.
  Checker tests: 21 passed.
- Independent read-only review checked mappings and mixed-profile consumers;
  remaining documentation and active tool vocabulary findings were addressed.
- Rewritten base remains 32ad64e320 and is an ancestor of this branch.

## Remaining work

OpenOptions envelopes, semantic limit profiles, specialized DOCX transactions,
other format families, Python/CLI and full acceptance remain required by the
active goal. This increment does not claim a performance improvement or full
completion of the limits contract.

## Cleanup

The six temporary fixtures are the base-checkout copies listed with hashes in
record 0035. They were removed after verification, together with generated build
output, root/fuzz lockfiles and temporary logs. No worktree was created; base,
iWork and pre-existing worktrees were preserved.
