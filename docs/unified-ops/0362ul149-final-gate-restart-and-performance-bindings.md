# 0362ul149 — restart the interrupted final gate; bind PPTX measurements to C

Root remains on `feat/unified-ops-api`. Candidate C remains immutable at
`f8cd398a770f7363f628be40edb3ce4219566c16`, tree
`7d57ecfc2b5f46f395f2eb49164ce82bc0c7eaba`. Protocol P remains at
`53451cc8d22499c3532502b3c986339a7206afdd`. Neither source is changed here.

## Interrupted execution

The session handle from 148 is missing, and a fresh process inspection finds
no surviving final-gate driver or build for that attempt. Its state has three
passing compiler receipts: changed-path formatting, minimal PPTX checking,
and OOXML ops Clippy. Root verifies their physical logs and target traces,
unchanged source/locks, and empty post-exit build/process observations. The
next row has a log but no completed receipt. There is no final result.

This is an incomplete run; its termination cause is unknown. All original
files remain at `(deleted-coord-artifact)`. The root
observation is `root-interruption-observation.json`, 4,353 bytes, SHA-256
`e71bab6657120dc238f43d0ad90c313721d9aebe73b35708533f6d94a89985b7`.
The observation also lists a separate read-only `rg` process matching the
source path; it is not a gate or build process.

## Durable restart

Root regenerates a fresh packet with the same reviewed v6 driver. Source,
tool identities, guard policy, order, and all 65 commands match the previous
plan; only the external state/target paths change. The driver's required
plan-hash dry-check passes. The entire gate starts again, without importing
the partial attempt's rows, in the user service
`litchi-unified-final65-cf8cd-v6-r2.service`.

The new external packet is `(deleted-coord-artifact)`:

| Artifact | Bytes | SHA-256 |
|---|---:|---|
| `packet/plan.json` | 183,048 | `959328f67b2cc6a89fadd2c519ddd8d089e1776c8173ebf2fb7a3f112f308839` |
| `root-launch.json` | 2,394 | `aa9f5302dec3991fe01b837654209b7aeaed2c087ae286c7ed06b5ac600f9143` |

The launch receipt contains the exact `systemd-run --user` command. The
driver remains the 23,875-byte file recorded in 148, SHA-256
`3e7aba2bdfc65f434c8730cb5ffa12db301583eb033b6f759da9976134d9ec01`.
The service is observed active with a live driver. This records a running
gate, not a terminal success. The full53 documentation rows remain isolated,
and the single build slot stays assigned to this gate.

## Measurement preparation

The notes packet is rebound from C780 to C in a new external directory,
`(deleted-coord-artifact)`. Its driver changes only candidate
commit/tree and the packet target prefix. The preparer and tests remain
byte-identical. Root checks the normalized four build rows and unchanged
guards, runner, result contract and subsequent measurement workload. All
eight harness blobs and both endpoints' three lock roles match their pins.
All 16 focused tests pass; the real source-only inspection reports
`source-bound-runtime-pending`.

| Artifact | Bytes | SHA-256 |
|---|---:|---|
| `run_f5.py` | 45,143 | `f088a28f664d2ef39da413ab189424674aa2c5781b8fd107a9905031a0179ef2` |
| `run-plan.json` | 18,226 | `c5fe203f7ad554d9c8b97f92e5651d309ade2bab339b44c95668ab437d64a593` |
| `static-inspection.json` | 8,949 | `4eff7b43ae21d28fe1b11c692148c43618d42166551fdb6448890af905392ff4` |
| `rebind-delta.json` | 1,690 | `53661a11ae5046bff27fc43c8d8510c85eea20560e8d07b2b3064e59c574de5f` |

The physical XML replay-scan packet is separately rebound at
`(deleted-coord-artifact)`. Its five code/probe/test
files are byte-identical to the reviewed v3 packet. The ten scoped source
blobs also match C780; commands, workloads and baseline a8 are unchanged.
Both endpoint source/lock observations pass, as do all eight protocol tests
and the exact file-inventory/hash guard. The new manifest is 4,917 bytes,
SHA-256 `8ff93931f64b6d78cbf52f576767925fc6fd2ffb1861b2c19127ca9ca52f765c`.
The adjacent root-validation receipt has SHA-256
`16d5b32c0e415c9507f1f1a5efd7a00a28592b67b0e7b1c649e7d1cfa47a0f0c`.

Both previous packets remain intact. Independent review confirms both
rebindings, their source/lock/tool identities and unchanged workloads, and
reruns the 16 and eight tests successfully. Neither new packet has built or
measured anything; no performance acceptance is inferred from preparation.

## Remaining qualification

The raw-census v11 provenance repair passes six source-only tests and
independent inspection of the real producer seam. Normal PATH lookup names
the executed tracing wrapper, and receipts separately identify its underlying
memory-cap shim. The C-bound plan still needs its actual capture binding,
manifest and preflight before execution.

Current-native v3 passes ten source-only tests but is returned because its
three-file historical-diff restriction rejects a shared final P containing
the other reviewed family bindings. Its successor must accept an explicitly
pinned common P while retaining exact current-native policy/manifests and
narrow installer writes. DrawingML preparation remains under repair.

Full acceptance gates, raw inventory review, native fidelity and quiet ABBA
measurements, final integration into trunk and retirement of the remaining
controls/baselines remain open. Only audit 22 is deferred. Nothing is pushed.
