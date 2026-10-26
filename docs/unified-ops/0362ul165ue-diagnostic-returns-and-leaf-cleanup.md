# 0362ul165ue — Diagnostic returns and leaf cleanup

The immutable-7c remaining-check diagnostic finishes all twenty selected
checks: thirteen pass and seven fail. Every runtime supplement passes,
including 58 PPTX minimal-feature chunks, PPTX batch equivalence, OOXML
metadata, facade transactions and shared OOXML error classification. No resource,
source-provenance or process guard stops the run. These are diagnostic
results, not acceptance of a final candidate.

Root rehashes all twenty logs and Cargo traces, verifies the unchanged source
and four locks, checks target bounds and empty post-exit process samples,
and verifies all 58 nested PPTX chunk logs. The terminal result agrees:
twenty attempted checks, seven ordinary failures, no hard stop.

The failed rows remain recorded:

| Row | Finding and disposition |
| --- | --- |
| 38 | Stale Python signed individual no-op assertion; repaired and merged in 165t. Signed batches still refuse immediately. |
| 41 | Standalone fuzz lock files absent under locked mode; fresh lock preparation is required before rerun. |
| 46 | 711 RTF canonical-definition findings and 36 non-RTF findings; reviewed leaves 165uc/165ud await Rust validation. |
| 48 | Mixed Python exception identities, stale DrawingML adapter identity, missing evidence registrations and missing rustc on PATH. Source repairs 165uf/165ug are merged; launch environment and rerun remain. |
| 49 | Shared ZIP dev edge and optional operations VBA edge lack policy registration; repaired in 165ug. |
| 50 | The systemd launch omits the rustup toolchain directory from PATH. Bind it explicitly before rerun. |
| 51 | Five protocol/compiler validation records lack matrix dispositions; precise nonmeasurement reasons are registered in 165ug. |

The canonical RTF and Markdown source repairs preserve the existing bans;
ODF, XLSB and Keynote fixes correct canonical checker identities without
restoring aliases. Independent static reviews find no remaining source
issue, but do not replace their pending compiler and runtime checks.

## Branch disposition

Root remains on `feat/unified-ops-api`. The fresh census verifies all seventy
excluded office-derived refs unchanged. The registry and tooling leaves are
merged at `a94459d439` and `cb6e2b0aa1`; their clean worktrees and refs are
removed after checking ancestry and process dependencies. Cumulative cleanup
reaches 286 worktrees and 448 refs. Active source leaves and native baseline
or protocol inputs remain until their dependent work finishes. Superseded
fuzz source worktrees can be retired after the v3 replacement is accepted;
their evidence packets must remain.

## Evidence

Paths below are relative to `/home/zhuhe/code/litchi-worktrees/.coord/`.

| Artifact | SHA-256 |
| --- | --- |
| `remaining20-api-glossary-failure-root-verification-20261026/receipt.json` | `5399ac0291f94eb6273fc8ba9e209c296f8e5b1c4ec2172afa63e2191f707d13` |
| `remaining20-tooling-tail-root-verification-20261026/receipt.json` | `650b5f642b1d12ca962684835a8b327ce5d36d9bfba3df9e3f454eb13f7518a6` |
| `remaining20-terminal-root-verification-20261026/receipt.json` | `f2f6d4d97f2c05c5d5d80d21994e034b49e005d1fac49436eb09ff97bdd77e19` |
| `remaining20-terminal-root-verification-20261026/result-addendum.json` | `e5e2c19bc2595a878261b17742c484ad1e69be0474f0c0d6584e041b7c3b750f` |
| `branch-worktree-165ue-20261026/receipt.json` | `e2a5c92f39db4a1a027c4057ecc534ae487c9b806960e6a32b4beb2292e591c7` |
| `gate-registry-165ug-retirement-20261026/receipt.json` | `9f69a2edad52202673c6ec372ba8a5ed52c47224764c8e91023a732a7182bcd0` |
| `gate-tooling-165uf-retirement-20261026/receipt.json` | `6d19b44e8d8778e40dcd0bc5cd892dd8bcaa576ce43432f7a070dfb6993e6f1b` |

All 29 native groups, actual final-source binding and the fresh final all-65
gate remain required. Nothing is pushed.
