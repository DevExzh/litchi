# 0362ul165g — build handoff review

Independent review returns ODG/ODP v3 for two failures in the actual composed
prepare/build/seal interfaces. Passing isolated tests does not establish an
executable qualification protocol. The author is repairing both on the
existing leaf; frozen C/P and the live full65 run remain unchanged.

## Two concrete returns

ODG's v3 binder places its fixture inside the C2a source checkout. The bounded
build stage can produce its receipt, but sealing copies that fixture path to
the final manifest, whose validator requires the fixture to be outside both
source roots. The public flow consequently fails before final admission.
The successor must use an external copy with the same fixed 738-byte identity.

ODP's build producer calls the shared streaming subprocess helper without its
required 4 MiB output limit. The helper defaults to ODG's 64 MiB limit, which
is recorded in the real build receipt. ODP sealing correctly rejects that
receipt because its monitor policy requires 4 MiB. The successor must pass
the ODP limit explicitly and test the produced receipt through sealing and
final validation. It must retain failed-build and changed-input refusals.

The reviewer reproduces these failures through the public interfaces using
controlled executable callbacks, without Cargo or native execution. The 45
focused ODG/ODP tests and 31 ODF raw/reuse tests pass but miss these two
composition defects. The v3 packet and independent receipt remain immutable.
Neither row is source-ready or qualified.

## Reusable root gate verification

The committed [`verify-final-gate.py`](results/0362ul165g/verify-final-gate.py)
checks the frozen driver/plan, regenerates the reviewed command sequence,
and validates recorded command order, source and locks, environment, log and
target-trace hashes, target limits and completed-process samples. It cannot
launch or restart the gate. Partial mode requires the same service to be
currently active; terminal mode requires all 65 rows and successful service
exit, the final result and the row-37 preservation barrier.

The initial partial invocation verifies the existing 33 rows. Ten negative
controls reject incomplete terminal coverage, sequence and command changes,
effective failures, phase gaps, source drift, remaining builders, lost live
guards, changed logs and changed target traces. This is not a new gate pass.
Nested chunk completeness/logs and preserved OfficeCLI bytes still need their
separate checks, as do the raw census and every native qualification row.

Independent review confirms the real partial CLI and terminal refusal, then
finds one false rejection: the approved watcher records an initial size before
its polling loop, so a valid short command can have zero subsequent periodic
samples. The successor accepts a nonnegative integer sample count, requires a
nonnegative integer initial maximum below 19 GB and the exact limit, and
rejects measurement errors. Existing resource and monitor-error checks remain.
Twelve focused controls cover the valid zero-sample case and eleven refusal
cases. The original tested source and receipts remain preserved.

Independent successor review exercises all twelve controls and the actual
partial and terminal CLIs again. The partial invocation verifies 33 rows on
the live service; terminal mode refuses incomplete coverage without creating
output. The narrow correction passes review and leaves source/plan/state
unchanged. No terminal gate acceptance is inferred.

The facade suite subsequently finishes. Root runs the reviewed successor
against the new row snapshot and verifies 34 completed top-level checks.
All 79 facade chunk logs are rehashed against their producing records; their
command sequence, successful exits and resource checks pass. The logs report
209 passing tests in total, including 19 library tests and 20 doctests. The
same gate continues to core tests; the remaining checks, including isolated
workspace rustdoc, have no terminal acceptance yet.

Run the verifier with a fresh external output directory; omit
`--allow-partial` for terminal verification. Its fixed inputs are the exact
C2a gate artifacts, not a generic or rebindable acceptance policy. The v2
external source and corrected committed source have identical bytes.

## Restored ODF lanes also require repair

Flat-parts v1 passes nine isolated tests and preserves its historical inputs,
but independent review returns four admission/provenance gaps: unrelated
nonempty bytes are accepted as the candidate lock, copied probe inputs are not
resealed, child output is unbounded and report validation can bypass the run
receipt, and the symlink check runs after resolution. The 36 synthetic legs
exercise allocation/fidelity orchestration only. The native timing obligation
still needs its own executable protocol, preserving the original allocator
control. This packet is not source-ready for final qualification.

W6 v1 recovers the original f755/e4a baselines and generated workloads, but
its proposed C2a lock overlay is not admitted: C2a uses path-vendored packages
while the historical trees use registry dependencies. There is no checked
dependency/vendor closure. The recipe also checks two origins and then clones
all four endpoints from C; its candidate overlay assertion requires changes
to side-workspace locks whose copied bytes are already identical, causing a
deterministic preparation failure. Later build/run stages lack the required
source, lock, executable, bounded-output and producer-receipt checks. Seven
policy and two output-validator tests passing do not close these failures.
Both restored-lane packets are returned for concrete repairs; no input waiver
or performance exception is recorded.

## Evidence

Paths are relative to `/home/zhuhe/code/litchi-worktrees/.coord/`.

| Artifact | Bytes | SHA-256 |
| --- | ---: | --- |
| `final-C2a-three-lanes-source-only-20261026-v3-independent-review-20261026/review-receipt.json` | 5,057 | `33e2f87ff483367611456d546a9739d4f5248a2389bc9a72483c1ffd2b435477` |
| `final-gate-C2a-root-verifier-20261026/verify.py` | 8,722 | `8f4c0ac2ea83de890b4eb267e92c599751bf928efce49e3aac4179145208ecf4` |
| `final-gate-C2a-root-verifier-20261026/partial-1/verification.json` | 25,186 | `31848d58c5018a20f160f6fe1280f0bd06bb18a9a5f2eaade5feeffad0e1fadd` |
| `final-gate-C2a-root-verifier-20261026/negative-controls.json` | 1,935 | `e6fb6b87882d92dc0e7c4356d6da279d073175d3938ac930e79f3b275fb01687` |
| `final-gate-C2a-root-verifier-20261026-independent-review-20261026/review-receipt.json` | 6,090 | `5ed4639169ad3acc89df84dfa1acbfea1570f92d65d99bff0a466a2e1c742cca` |
| `final-gate-C2a-root-verifier-20261026-v2/verify.py` | 9,037 | `29a6831801de026e994fe7b56156dc326a4004a3f8b23e66be122f1b605ab0a7` |
| `final-gate-C2a-root-verifier-20261026-v2/guard-controls.json` | 2,039 | `851d7b6365b7b846f71f1ee2086d99b2ab81a74897d1055763c158b035c08801` |
| `final-gate-C2a-root-verifier-20261026-v2-independent-review-20261026/review-receipt.json` | 4,762 | `a9c2c2b988597ef7db946e2cadb24c5e87a4caeefff58ad124854918c3906d0c` |
| `final-gate-C2a-root-verifier-20261026-v2/partial-2-facade-complete/verification.json` | 25,908 | `ae9d25c5c431c46d00cb4b522b5f3ac8e057a558d440d3b96af7ccfc6e1d4b00` |
| `final-gate-C2a-root-verifier-20261026-v2/facade-79-chunks-verified.json` | 56,814 | `18e28b1e1842f665e4f00eb3046665876ac153752ed836e943bdfed416eb31c7` |
| `odf-flat-parts-C2a-prep-20261026-v1-independent-review-20261026/review-receipt.json` | 8,482 | `20c6888e7546ed7c95c89887b2344103280290ef8493c9bf864f47e53b2bbaad` |
| `w6-odt-ods-C2a-prep-20261026-v1-independent-review-20261026/review-receipt.json` | 11,542 | `76747562bb685631e432fd03fd7d2f9344b51bd57568710e23afcb57f83cd704` |

Flat-parts and W6 ODT/ODS repairs remain required. The corrected 28-row scope
remains required, with no final-C2a native acceptance yet. Only memo 22 is
deferred. Nothing is pushed.
