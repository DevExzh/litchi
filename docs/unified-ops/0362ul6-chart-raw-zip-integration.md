# 0362ul6 — reviewed chart, raw and ZIP API integration

Status: source assembled on `uo2/coordinator-replay-integration` through
`0dcc767c68`. Final acceptance remains open. Only the comprehensive redaction
audit in item 22 is deferred; its rules still apply to these changes.

## Reviewed source

Spreadsheet chart builders (`341164e0bf`, merged at `1f2d0509b5`) now use
shared `chart::builder` and host `builder::chart` paths. The old authoring
aliases are removed and all consumers migrate together. Review confirms
unchanged chart models and codec behavior. Record 0362zab reports 2,675
passing drawing/XLSX/XLSB tests and consumer, example, harness and glossary
checks on the 52ed graph. The standalone chart harness lock adds required
local dependency edges; all 122 registry identities and checksums stay fixed.

Legacy raw boundaries (`6a55d251d2`, merged at `db433b2474`) require explicit
limits, preserve narrower retained source policies, and expose format-owned
typed errors. Review returned two real defects before merge: the typed error
classification/redaction gaps repaired in 0362zaca, and XLS revision-log
allocation before limit admission repaired in 0362zacb. The latter checks
declared stream length before materialization; the counting-reader regression
proves that refusal performs no additional body reads. Focused native,
facade and feature checks pass as recorded. Existing raw DOC cursor
mutability remains governed by 0351 §2.1.6 and 0359a §4; no new exception is
introduced and immutable Package outputs remain unchanged.

ZIP constructors (`0e3cfc0e1c`, merged at `0dcc767c68`) take explicit archive
and cache limits through their sole canonical constructors. The review return
0362zaea removes the overlooked lazy-cache aliases, preserves ODF's retained
archive profile, and repairs a standalone fuzz consumer. All 631 ZIP tests
pass with one existing ignore; focused ODF retained-limit and cache-setter
tests and affected consumer checks also pass. Records 0362zae/0362zaea retain
the actual command coverage and distinguish interactive results from saved
logs.

## Composed checks and dependency pin

At the chart merge, 77 focused tooling tests pass. Workspace formatting then
finds three existing IWA alias-migration formatting leftovers in two files.
Commit `2b83731236` fixes those lines; `cargo fmt --all -- --check` passes.
The failed and successful logs remain outside Git under
`.coord/final-acceptance-2026-10-24/` as
`fmt-chart-1f2d0509b5.log` (1,788 bytes) and
`fmt-iwa-alias-correction.log` (empty successful output). This formatting
result precedes the later raw and ZIP merges and is not a final gate result.

The raw repair adds a direct `litchi-ops -> litchi-vba` dependency edge.
The coordinator independently compares the parsed root lock graphs and
adopts the reviewed leaf lock without dependency resolution or regeneration:

| Pin | Bytes | SHA-256 |
| --- | ---: | --- |
| Previous | 139591 | `52ed0cacac15b5e440349bfc0f4aca2665bcf50d59d4321624e596bb410783c6` |
| Current assembled source | 139606 | `4100b1fbb272dd24842f33a2d6f353a3f3a97dfa7e435c4649a0e1a029bfaea1` |

All 518 registry package identities remain unchanged. The only package-graph
difference is that local dependency edge. Both lock inputs and the reduced
`legacy-raw-pin-transition.json` receipt remain outside Git under
`.coord/final-acceptance-2026-10-24/locks/`. Historical frozen-49c captures
and their 52ed input remain untouched. The six-version-different 6d graph
remains diagnostic and is not substituted for either pin.

## Remaining work

CFB canonical required-limit constructors and their consumers are still under
validation. DrawingML is returned for retained owner budgets, replay
admission before allocation, typed limit and conflict errors, and preservation
of the ordinary no-op fast path. Neither unfinished lane is accepted by this
record. After those repairs and review, fresh all-68 inventory/dispositions,
the complete 53-step gate, actual corpus replay and native performance
qualification must bind the assembled source and its reviewed lock.

Cleanup thirty-eight removes five merged leaves while preserving every
ignored payload and branch; its verification and external evidence are in
the cleanup record. Nothing is pushed.
