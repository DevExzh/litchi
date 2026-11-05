# 0362ul44 — ODF raw-codec limits

- Basis: integration `30aaafbf98e097aeec8cf6fe3df6caaae62a37d6`.
- Scope: source-only API migration. No Cargo build, native run, benchmark, or
  acceptance claim was made in this leaf.

## Migrated surfaces

The public variable-size ODF codec entry points now require the owning limit
policy, and all current consumers were updated in the same change:

- ODS tracked-change parse and XML writers use the tracked-change `Limits`.
- ODT document scripts, font faces, list styles, and note configurations use
  retained `XmlStreamLimits` for parsing, serialization, and XML mutation.
- ODF-common annotation builders and writers carry `XmlStreamLimits`; ODS,
  ODP, and ODT annotation readers/writers pass their package policy.
- ODF formula serialization and element XML use the formula `Limits`, with
  output checks performed while emitting.
- ODS database-range parse/source/filter and writer entry points require
  `XmlStreamLimits`; package catalogs and commits retain the archive policy.
  Builder-only ingress uses the owner default.
- ODS style-protection extraction and managed-style rewrites require limits;
  protection snapshots retain the package policy through parse and commit.

## Review follow-up (`c1775daa5ef553dae3351e9680c519376bb5acd7` preserved)

The follow-up keeps the prior diagnostic tip intact and closes the frozen
review gaps in one source change:

- Formula owned ingress derives codec limits from the caller's XML and
  semantic profile instead of reinstating `codec::Limits::default()`.
- DataPilot and database-range location, typed parsing, replacement, and
  rendering now carry the required XML policy.  The database-range and
  DataPilot nested parsers use the shared bounded stream admission pass for
  non-default profiles, while the ordinary owner path keeps its single-pass
  typed performance.
- Protection's fused and sequential location paths share event, depth,
  aggregate-attribute, token, and source-byte admission.  ODS, ODP, and ODT
  annotation scanners apply caller event/depth/attribute policy rather than
  top-level-only or fixed nesting checks.
- ODT script, font, list, and note scanners use the shared bounded admission
  path for narrowed policies, and their writers avoid eager capacity
  reservation before the final configured output check.  Database/DataPilot
  writers admit output atomically and retain exact-limit behavior.

Focused source tests cover tiny event/depth/attribute budgets, exact output
budgets, nested database/DataPilot input, formula owner-profile retention, and
annotation admission.  No Cargo build, native run, benchmark, or acceptance
claim was made.

The changes preserve existing typed validation, source-preserving mutation,
and opaque-content refusal behavior. No aliases or limit-less public twins
were added. The only default-policy calls are fresh-authoring/builder or test
ingress paths; package-backed projections pass their retained policy.

## Verification boundary

`git diff --check` and source parsing via `rustfmt --check` were run. Cargo,
native timing, profiling, and performance qualification remain for the owner’s
bounded follow-up command; this record does not assert compilation or runtime
acceptance.
