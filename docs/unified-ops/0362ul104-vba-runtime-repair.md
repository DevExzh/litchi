# 0362ul104 — VBA runtime admission and manifest repair

This source-only repair starts from the frozen OOXML runtime candidate
`ea5c95e212e1f8e934a0485eb0b1f51355f13ba5` (tree
`769c360f14d24784b3cc25bc807875a75686e486`).  The coordinator's common
runtime receipt identified four VBA failures; no compiler, Cargo, rustdoc,
native, or performance command was run in this leaf.

Recipient admission now checks retained part metadata before parsing
`[Content_Types].xml`.  A deferred graph whose selected payload exceeds a
tighter edit policy therefore returns the VBA `PartBytes` limit identity
without hydrating the payload.  The modeled per-part pass also precedes the
retained-archive metrics pass, so simultaneous part and entry ceilings keep
the selected VBA limit identity while archive member and aggregate accounting
remain intact.

The noncanonical content-types replay test now spells out the expected Excel
fixture manifest independently: it replaces the known plain workbook
override with the required macro-enabled main-part and VBA-project overrides.
It still checks exact bytes, including the recipient/source spelling marker.
The structural-limit fixtures now set their `parts` and `relationship_parts`
ceilings alongside `members`, so they exercise the intended refusal kinds
instead of rejecting the limit configuration itself.

Source inspection, formatting, and repository-size/binary checks passed in this
leaf.  Cargo, compiler, and runtime execution remain intentionally pending for
the coordinator's focused and common runtime rerun.
