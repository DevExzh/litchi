# 0362hd — ODF flat Parts raw policy

- Date: 2026-10-22.
- Parent: W8 Part A candidate `8d6abdab69`.
- Status: implemented source slice; coordinator integration and the final
  workspace inventory remain required.

The public raw `litchi_odf_common::flat::Parts` helpers now require a focused
`flat::Limits` carrier. It combines the caller-selected `SourceLimits` and
`SemanticLimits` without creating a second `OpenOptions` or selecting an
implicit owner policy. `split`, `merge`, `to_bytes`, and `merge_package` all
receive that carrier. `merge_package` admits an independently supplied ZIP
against its physical source ceiling before copying it.

A flat owner already admitted its XML and packaged projection, so it uses
`merge_package_projection` for republishing. That route uses a separate,
hard-bounded synthetic-package profile and the retained semantic policy. It
does not charge a generated ZIP's framing bytes against the original flat XML
source ceiling or the published flat-document output ceiling. `Parts::to_bytes`
checks the logical XML part dimensions and uses the same bounded ZIP scratch
profile. `Parts::merge` checks the resulting flat bytes against content and
output limits before reserving or publishing its candidate.

The scanner checks source bytes before parsing, semantic depth while scanning,
UTF-8 and XML structure before publication, and uses checked lengths plus
fallible reservations for child ranges, part XML, replacements, package copies,
the bounded ZIP sink, and the final flat result. Raw child-range allocation is
bounded by the already admitted source buffer; it is not charged to semantic
paragraph/table/sheet/slide unit counts. XML bytes and unchanged source ranges
remain exact, including BOM and root metadata handling.

All consumers transfer their retained policy: ODT flat projection, ODP flat
materialization and republish, ODS flat materialization and republish, and ODG
flat tree publication. ODG retains the flat source bytes together with the
focused limits carrier, so later republish cannot silently fall back to
workspace defaults. The ODS facade reconstructs the facade options around the
retained ODF subprofile when reparsing a new flat owner.

Focused tests cover exact and one-under source/content limits, output and
nesting refusal before publication, external-package source refusal, and an
independently generated ZIP larger than its exact flat source that succeeds
through the projection route. The common flat tests pass 9/9; the ODT, ODP and
ODS library suites pass 807, 221 and 694 tests respectively. The facade flat
ODF integration suites pass 7 `flat_odf_unified`, 7 `odp_unified_write`, 16
`ods_unified_write`, and 8 `odt_unified_parts` tests. The capped all-feature
check for `litchi-odf-common`, `litchi-odt` and `litchi-ops` also passes.

No native performance qualification was run in this slice. The coordinator
must merge W4/W4b and the remaining W8 work, rerun the workspace inventory and
glossary, and qualify any performance-sensitive projection route on matched
binaries.
