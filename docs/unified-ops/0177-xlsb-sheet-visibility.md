# 0177 — XLSB sheet visibility

## Contract

The unified XLSB sheet registry now admits `hidden` and `visibility` alongside
`name`. The shared closed visibility vocabulary is `visible`, `hidden`, and
`veryHidden`; its normal case-insensitive coercion applies. Boolean `hidden`
assigns the ordinary hidden state or visible state. Supplying both properties
in one assignment is refused as an ambiguous duplicate of the same state.

Visibility selects through the complete tab catalog, including chart tabs.
Changed hiding of dialog/macro tabs is refused until their view codecs are
modeled. A tab without
a physical sheet payload must remain very hidden. Name writes still require
a worksheet and retain their existing name-identity/dependency admission.
Combined name/state assignments publish through one native structural edit.
A visibility query must leave at least one visible tab; failures leave the
caller snapshot unchanged. Native policy and dependency checks remain the
publication authority. Hiding an active tab relocates the workbook view to
the next visible tab with wraparound and updates worksheet/chart selected
flags. Hiding a selected non-active tab also clears its selection. Other view
fields, including the first displayed tab, remain unchanged. Active relocation
across multiple workbook views and unsupported target-view layouts are typed
capability refusals. A hide followed by a reveal in two published facade
operations does not reactivate the original tab; its inverse patch restores
the complete original view state.

The read projection keeps its established sparse representation: visible
state and false `hidden` are omitted from the format map. Hidden and very-hidden
states are explicit. Empty assignments and exact state no-ops retain the
original bytes. Signed facade writes still require explicit signature
consumption before editing.

## Evidence

Independent implementation, regression-test, benchmark and review agents
completed their work in isolated worktrees, which were removed after integration.
The base branch is already contained in this branch (zero base-only commits).

Verification passed:

- 1,159 native XLSB tests/doctests with `sign,vba-inspection`, including fourteen
  new visibility regressions and the runnable setter example.
- 1,217 combined facade tests/doctests with `docx,xlsx,pptx,xlsb,encryption`,
  followed by 24 final focused XLSB tests after the last review fixes.
- Warning-denied native and combined-facade all-target Clippy, facade rustdoc,
  native rustdoc JSON and the XLSB glossary profile (31 banned names).
- Minimal facade Clippy, crate boundaries (66 packages, 268 declarations),
  generated OfficeCLI parity check and fourteen parity-generator tests.

The root and shared performance lockfiles remain unchanged. Verification used
an external copied root lockfile with cached `cc` 1.6.0 through Cargo's
`resolver.lockfile-path`; neither repository lockfile was rewritten.

The paired harness at `tools/perf-baseline/unified-xlsb-visibility` measures
one native edit against one unified query fanout. Two nine-round ABBA matrices
cover eager/source owners and 1, 10, 100 and 1,000 target tabs. The JSON records
include all timing samples, generated corpus identities, source-file hashes,
build command, compiler and binary SHA-256. No binary is committed.

Results: [after](results/0177-xlsb-visibility-after.json) and
[repeat](results/0177-xlsb-visibility-repeat.json). Every native/unified output
and property-alias comparison passed. One-tab ratios were 1.059/1.075 eager
and 1.031/1.040 source, meeting the measured 10% overhead threshold. At 1,000
tabs the ratios were 1.180/1.176 eager and 1.145/1.146 source, below 2×.
However, unified time grew roughly 35–38× from 100 to 1,000 tabs, so linear
scaling is **not** accepted. Native publication also grows superlinearly;
profiling its catalog/publication work and facade target admission is a
required performance follow-up. These measurements cover query fanout, not
1,000 separate facade commands.

A full 10,000-tab matrix was stopped after more than eighteen minutes without
a completed report. It supplies no accepted timing evidence. The complete
1→10k scaling requirement remains open; the harness retains the 10k default
so the measurement can be resumed in a dedicated performance follow-up.

## OfficeCLI delta

The reference `ExcelHandler.Set.Sheet.cs` changes the workbook-level sheet
state and refuses hiding the last visible sheet. Those semantic behaviors are
retained here. OfficeCLI additionally accepts loose truthiness and spellings
such as `very` or `none`, and falls back on truthiness for unknown strings.
Litchi uses the existing strict boolean/closed-token coercion instead; unknown
values and conflicting aliases fail before native publication. This is the
ADR 0034 strict-default policy, not silent compatibility fallback.

## Remaining scope

A query fanout shares one native structural edit. Separate sheet commands in
closure/JSON batches still publish at each sheet-operation boundary; homogeneous
sheet-command coalescing and its 1k/10k acceptance remain follow-ups.

This change does not close worksheet creation/removal/order, formula authoring,
formatting, or the all-format, Python, CLI and acceptance work in the full goal.
