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


## ODG baseline prerequisite

A fresh external baseline endpoint at
`.coord/odg-joined-baseline-ad524` retains clean tracked source
`ad52437f5e5e2c71ea476eedd5ddcf61aee2021d`. Offline lock generation under
Rust/Cargo 1.95.0 produced the 139,592-byte `6d7895e6...` lock. That original
artifact is retained as
`.coord/odg-joined-baseline-ad524-evidence/Cargo.lock.baseline-generated-6d7895e6`.
It validates against the baseline, correcting an earlier untested assumption
that it was incompatible. Its dependency resolution differs from the
candidate, so it is not the lock selected for the matched comparison.

Seeding the baseline with the candidate's 139,591-byte ignored lock,
SHA-256 `52ed0cacac15b5e440349bfc0f4aca2665bcf50d59d4321624e596bb410783c6`,
passes offline locked all-features metadata. The coordinator independently
traversed the ODG dependency graphs and verified equality of all 228 external
package name/version/source/checksum rows. The baseline production tree and
tracked status remain unchanged. This is dependency preparation only; no
native binary, timing or acceptance result is claimed.

The external evidence directory is
`.coord/odg-joined-baseline-ad524-evidence/`. Its source-bound metadata is
regenerated with `cargo metadata --offline --locked --all-features
--format-version 1 --manifest-path <endpoint>/Cargo.toml`, after sourcing the
capped `.coord/odg-native-protocol/env.sh` and overriding target/TMP to that
external evidence directory. The baseline endpoint above uses the seeded
candidate lock; the candidate capture used clean diagnostic source `2dfc4f134d`.

- `baseline-metadata-seeded-candidate-lock.json`: 3,093,862 bytes, SHA-256
  `62ee160686d542e080da8bcf40a13a0532311fc3f3a2f14fee669402b94b81a5`.
- `candidate-metadata-all-features.json`: 3,057,334 bytes, SHA-256
  `10b9dd7a55e854223a44340f68b08367de7d60bf7866159f8d9cfa37b70e5385`.

These large metadata artifacts remain outside Git. The separate ODF
shared-triad baseline remains `173dd5330476713714871007316794b9a3a1bdaa`;
ODG preparation does not qualify that separate lane.


## Composed strict consumer check

After the XLS merge and formatting repair, clean source
`97817cb75f5bd569cfaf13f73331b4bcd82d22ae` passes the selected all-target,
all-feature strict Clippy step for `litchi-cli`, `litchi-ops`, `litchi` and
`litchi-py`. The exact command and source/lock/environment identity are retained
in `results/0362ugu/composed-strict-check.json`. It completed in 167.272
seconds with an empty log, no owned builders after exit, and stable target
size 4,706,734,564 bytes. This covers the composed XLS fallible consumers and
PPTX refusal classification. It is one selected step, not the full 53-step
gate, and predates the separate W8 baseline-adapter binding correction.


## Refreshed candidate

IWA fixture binding `4dae391375` merges at `494f399c74`; review confirms
that the fixture generator slice is byte-identical across its test-consumer
migration. W8 binding and pristine-baseline consumer correction
`d714aee19d` merges at `b6732d02b241705ba153c1a2cb9ff03bf0a13f93`.
The W8 adapter removes only the three deferred XLSB resolution chains that
the original f2 API does not expose, in addition to its four existing result
accessor translations. The resulting baseline probe is byte-identical to the
previously reviewed d39 adapter; its portable regression pins that size/hash.
The 109-route workload and timing/proof boundaries remain unchanged.

On clean `b6732d02b241705ba153c1a2cb9ff03bf0a13f93`, the six-module native
protocol command recorded above now passes 65 tests, with zero failures or
skips. The external `native-protocol-tests-b673.log` is 752 bytes, SHA-256
`933f6d49522fcc92db37fe66e6e4fd7ac6133cd9dd4a76c04199e6556ab31e42`.
The explicit iWork native protocol command passes 20 tests; its 119-byte
`iwa-native-protocol-tests-b673.log` has SHA-256
`992669421ab4a90e4dc1b7b52d73d5d61c3518a59586510f8174fcd0d9e06051`.
Both logs are under `.coord/final-acceptance-2026-10-24/`. These are protocol
checks, with no native measurement claim.

The full 53-step gate restarted on the same clean frozen source, using
`.coord/replay-integration/gate-safe/run_safe_gate.py` without `--only` or
`--start-at`, target `.coord/budget-validation/target`, and state directory
`.coord/final-acceptance-2026-10-24/full-gate-b6732d02b2`. Formatting passes
in 70.764 seconds and the composed strict consumer step passes in 17.547
seconds. The remaining steps continue. This is an in-progress acceptance
attempt, not a full-gate pass. The separate all-68 rustdoc inventory and quiet
native qualification remain required.
