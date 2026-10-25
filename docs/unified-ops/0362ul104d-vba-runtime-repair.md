# 0362ul104d — VBA runtime limit and manifest repair

This source-only repair starts from the frozen OOXML runtime candidate
`e130751fc503099c4d5fff74a5746b5099d625e1` (tree
`0251a7e0945eb18d30b80779526ecbe5e0b2fd25`).  The coordinator's common
runtime receipt identified three remaining VBA failures; no compiler, Cargo,
rustdoc, native, or performance command was run in this leaf.

The noncanonical content-types oracle now independently removes the known
plain workbook override and inserts the required macro-enabled main-part and
VBA-project overrides at the manifest close.  This keeps a recipient-only
comment before the generated fragments and still checks every byte exactly.

The structural-limit tests use a tracked Excel source fixture with directory
entries and count the serialized candidates through the existing ZIP reader,
independently of the production retained-metrics planner.  They use exact
candidate budgets and one-under limits that satisfy all `SourceLimits`
cross-field invariants while exercising the intended refusal kinds.  The
replay test consequently distinguishes the member and total-entry ceilings
instead of constructing an impossible policy.

Source inspection, formatting, and repository-size/binary checks passed in this
leaf.  Cargo, compiler, and runtime execution remain intentionally pending for
the coordinator's focused and common runtime rerun.
