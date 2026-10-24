# 0362ugu — Composed acceptance

The reviewed XLS lazy metadata, property-set and worksheet-feature packet
ends at source `3ae7b831406b01a801d5e75b171db0974243a13c` and receipt
`9d38e2c4f28e28be4532823e66a13de5637bb916`. It merges at
`2dfc4f134d94195dd9b41f2d806887b38ef10c58`. The coordinator independently
verified the source-bound logs, commands, locks and configuration in 0362ugea,
including stable candidate-error and concurrent-source-change controls for
both admission and worksheet rendering. The composed OLE benchmark retains
all three fallible DOC, XLS and PPT metadata accessor migrations.

## First frozen attempt

The clean diagnostic worktree at `2dfc4f134d94195dd9b41f2d806887b38ef10c58`
started the full 53-step safe gate with no selection filters. Its first step,
`cargo fmt --all --check`, failed after 81.780 seconds. The external log is
`.coord/final-acceptance-2026-10-24/full-gate-2dfc4f134d/00-fmt-attempt-1792815904599553077-3713346.log`
(10,644 bytes, SHA-256
`929e9bf021a568ce3217a2f9b7185cb9c2c9a95dfe910a4583938dbd3a5c088f`).
The reduced receipt in `results/0362ugu/first-gate-attempt.json` retains the
full selection, source/lock identity, target and process-exit checks. The
target was 4,706,726,230 bytes; no owned builders remained. No later gate
step ran, and this failed attempt is not acceptance evidence.

Running `cargo fmt --all` on the composed source repairs ten files. Changes
are whitespace-only except rustfmt's removal of braces around one expression
in an XLSX hyperlink filter closure. Its expression and behavior are unchanged.
The full gate will be restarted on the corrected source.

Separate build-free native protocol checks on the same frozen source found
stale reviewed identities. The six-module command covering the native runner,
ODP, W8 provenance/original baseline, XLS F3 and XLSX source-family batching
ran 63 tests with 11 errors, all in W8 probe identity checks. Its external
`native-protocol-tests-2df.log` is 14,075 bytes, SHA-256
`ee015827821a3c6aa231c93481384729d9255c35e287b1e451d5ed24221b426d`.
The iWork native protocol suite ran 20 tests with one fixture-identity error;
`iwa-native-protocol-tests-2df.log` is 1,273 bytes, SHA-256
`507aeed85661b6baa8d18f0f9c9fb3c893f2c4a3595a7901c301855be0aa673f`.
Both logs are under `.coord/final-acceptance-2026-10-24/`. These checks remain
required corrections; source identities will only be refreshed after review
of the intervening source and fixture changes. No native timing was run.
