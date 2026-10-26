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

The actual partial invocation verifies the existing 33 rows. Ten negative
controls reject incomplete terminal coverage, sequence and command changes,
effective failures, phase gaps, source drift, remaining builders, lost live
guards, changed logs and changed target traces. This is not a new gate pass.
Nested chunk completeness/logs and preserved OfficeCLI bytes still need their
separate checks, as do the raw census and every native qualification row.

Run the verifier with a fresh external output directory; omit
`--allow-partial` for terminal verification. Its fixed inputs are the exact
C2a gate artifacts, not a generic or rebindable acceptance policy. The tested
external source and committed source have identical bytes.

## Evidence

Paths are relative to `/home/zhuhe/code/litchi-worktrees/.coord/`.

| Artifact | Bytes | SHA-256 |
| --- | ---: | --- |
| `final-C2a-three-lanes-source-only-20261026-v3-independent-review-20261026/review-receipt.json` | 5,057 | `33e2f87ff483367611456d546a9739d4f5248a2389bc9a72483c1ffd2b435477` |
| `final-gate-C2a-root-verifier-20261026/verify.py` | 8,722 | `8f4c0ac2ea83de890b4eb267e92c599751bf928efce49e3aac4179145208ecf4` |
| `final-gate-C2a-root-verifier-20261026/partial-1/verification.json` | 25,186 | `31848d58c5018a20f160f6fe1280f0bd06bb18a9a5f2eaade5feeffad0e1fadd` |
| `final-gate-C2a-root-verifier-20261026/negative-controls.json` | 1,935 | `e6fb6b87882d92dc0e7c4356d6da279d073175d3938ac930e79f3b275fb01687` |

The flat-parts packet has entered independent review. Original W6 ODT/ODS
preparation continues. The corrected 28-row scope remains required, with no
final-C2a native acceptance yet. Only memo 22 is deferred. Nothing is pushed.
