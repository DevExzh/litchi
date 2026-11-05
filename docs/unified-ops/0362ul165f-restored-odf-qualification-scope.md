# 0362ul165f — restore missing ODF qualification scope

The final qualification inventory now contains 28 required rows. Independent
audit finds three omitted groups in the previous 25-row list: common flat
ODF parts, the two original W6 ODT performance residuals, and four W6 ODS
source-backed routes. None has final-C2a closure. The omission is corrected;
the historical results and code-generation explanation do not waive the
remaining work. Candidate C and protocol P remain frozen at
`33a0a63d3553712c1276da592f7846a0c7dcabd6` during the live full65 gate.

## Restored requirements

| Added row | Required coverage | Why existing rows do not close it |
| --- | --- | --- |
| ODFflatParts | The nine hda split/publication/merge/proof/owner routes, with typed foreign-source, depth and bounded-output checks and ODT/ODS/ODP/ODG projection fidelity | The reviewed raw-codec probe covers Archive, manifest/XML, writers and raw facade access; it does not exercise common `flat::Parts`. The hda receipt establishes historical allocation/fidelity evidence only. |
| W6ODTresiduals | Original ODT set and flat ODT open/get workloads, with original pre-regression baseline and default release settings | The 1.070× and 1.090× residuals remain unresolved. W8 opens packaged ODT, and its later baseline already contains the W6 changes. The two-route `codegen-units=1` diagnosis does not close the default-release result. |
| W6ODSsource | `cells1`, `cells100`, `dde` and `sheet_metadata` through `SourceBackedSpreadsheet` | W8's packaged ODS operations do not exercise these source-backed families. No final-C2a mapping or execution receipt was found. |

ODP tail append already has a required qualification row; it is not duplicated.
Its original heap residual remains subject to that row's final evidence. The
audit does not require indiscriminately rerunning every historical benchmark.
It restores explicit requirements and named residuals without later closure.

The corrected external v3 catalog preserves v2 unchanged. It also corrects the
stale claim that ODF raw codec lacks a protocol: reviewed leaf `8e400cbaee`
provides that source preparation, with integration, compilation and actual
qualification still pending. New flat-parts and W6 preparation workers retain
the original probes, fixtures, limits, measurement boundaries and baselines.
No substituted workload or inferred owner exception is authorized.

## Sheet source preparation and protocol composition

Sheet v6 passes independent source review for XLSX source families, XLSX raw,
XLSF3 and XLSB. Its future-P patch migrates both the XLSF3 runner and tracked
test consumer from the historical harness reference to C2a. The public parser,
non-diagnostic admission and source-identity checks are exercised directly;
wrong sources, stale harnesses and duplicate routes fail closed. Four positive
lane checks, eleven negative guards and the three existing provenance tests
pass. All seventeen packet seals match, and actual `git apply --check` passes
against C2a without modifying it. Future-P identities remain unbound until
the actual consumer commit exists.

This brings source-ready preparations to 23 of the corrected 28 rows. ODG/ODP
v3 is under independent review; the three restored rows are in preparation.
All 28 final-C2a executions remain pending. Source readiness does not imply
compilation, correctness, fidelity, allocation or native acceptance.

Root extends the previous external composition preview with the reviewed
legacy v3 four-consumer patch, sheet v6 two-consumer patch and ODF approved
overlay-reuse change. All three apply cleanly. The preview contains 26 files,
including twenty Python files that parse successfully. This is patch
composition only: no real current/signed binding, archive or integrated test
result is generated. The P1/P2 ordering and normalized signed successor patch
from 165e still apply. ODG/ODP v3 and the restored lanes are outside this preview.

## Gate and space reclamation

The same full65 service remains active with 33 verified completed checks.
Thirty facade log chunks were observed at this checkpoint; the facade row
has no terminal result yet. The complete 132-chunk Ops result from 165d is
unchanged. No competing heavy build, archive or native qualification starts.

Root removes three exact retired crypto test/debug targets after independently
checking their terminal receipts, logs, preserved integration ancestry,
absent retired ref/worktree, unchanged directory identities and visible
process holders. The guard encounters two inaccessible service processes,
`systemd` and `(sd-pam)`, already explicitly classified by the existing
cleanup policy; root additionally pins their process start ticks. Process
visibility is therefore not claimed to be complete. No visible holder or
unclassified denial remains. Receipts and logs are rehashed after removal.

The removed caches contain 3,841,607,899 apparent bytes and 3,859,337,216
allocated bytes. Active targets, native baselines, original flat/W6 inputs
and evidence remain intact. This removes no additional worktree or ref.

The refreshed census verifies 75 heads and 62 worktrees. All seventy excluded
refs are unchanged. The four merged unified controls and active ODF leaf are
still required; the leaf is now `df15ec337d3eea3b62b1fb082157e3b0c1e1e220`
with committed ODG/ODP build-handoff repairs awaiting independent review.
Cumulative removal remains 277 worktrees and 439 refs. No other unused
unified-derived leaf is found.

## External evidence

Paths are relative to `(deleted-coord-artifact)`.

| Artifact | Bytes | SHA-256 |
| --- | ---: | --- |
| `odf-w6-flat-coverage-audit-C2a-20261026.json` | 14,226 | `5d0961bc7c12d7c7d53bba343bb9f4e1dd1fe9788d037d89fd96d35f8854bbd3` |
| `final-native-performance-reconciliation-20261026-v3.json` | 51,688 | `4942e52435768ea43bb241bf247b377569b09181cb99cc272b3136acd48e1b34` |
| `sheet-c2a-native-prep-20261026-v6-independent-review-20261026.json` | 4,566 | `51cbfb7c59e63c6f6a1a8d93283b9215af7d9d59b940fd70999efbf0225a85d5` |
| `protocol-consumer-composition-C2a-20261026-v2/receipt.json` | 13,511 | `eb51441b5e9eca6b077acb1e17a4c763c113d2c4fe89780c3602c75d227cd4ef` |
| `target-cache-reclamation-audit-20261026/proposal.json` | 5,734 | `89b704beecc497d1a9651b3d285a1450f6337e2c4211ad59bfabb5e9706ed982` |
| `target-cache-reclamation-audit-20261026/root-removal-receipt.json` | 3,803 | `b6068d31c6c5be0ebb1259cf354ef2ba6ebe056bb4c140b509c45885f18f626f` |
| `branch-worktree-165f-20261026/observation.json` | 18,187 | `73cd5fc52e0481ab4e2fdb925b7fa84c2e7df24b579cfcc42d71db032cc0ba8b` |

The full gate, raw census, corpus proof, all final qualification, protocol
integration and final control cleanup remain required. Only memo 22 is
deferred. Nothing is pushed.
