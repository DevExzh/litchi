# 0978 variable-layout decoder evidence packet

**Status: sealed failed-capture packet.** This packet records bounded offline
work for the 0978 variable-layout scanner, lifecycle ledger, decoder guard, and
scoped reducer replay, plus the one authorized Office run. The final
pre-capture plan and helper seal were reviewed before the separate root GO. The
serial run stopped at child 3 after its loss gate reported four terminal lost
samples: two children gated successfully, children 4--12 were not started, and
there was no retry or replacement. This packet makes no qualification,
performance, source-owner, causality, or optimization claim.

The governing offline policy/design review is
[`independent-design-review-v3.json`](review/independent-design-review-v3.json).
It requires every valid sample to remain in a disjoint lifecycle and
attribution ledger, keeps startup/transition/non-user-context/missing-unwind
periods explicit, separates physical DSO ownership from target-source
eligibility, and requires variable-layout parsing, lifecycle accounting,
conservation, and mutation controls before any future native GO. The copied
[`protocol-design-v1.md`](design/protocol-design-v1.md) is a proposal that
makes this boundary explicit; it is not a GO or a sealed integration schema.

## Selected offline evidence

The raw scanner controls v2 pass against the retained 0977 child inputs. Its
scanner source is `tools/raw_stack_0978.py` at SHA-256
`5f844b82a40426662ffaed5ba21dd0687fefa3f98ecc233dc572eeb3041017b7`, and the
receipt is `receipts/raw-stack-scanner-controls-0978-v2.json`. The independent
mutation receipt records 10 tests passed with scanner and test hashes unchanged
before and after. The child-1 and child-2 replay receipts are separate: child
1 has 7,200 nonzero-stack sample records; child 2 has 7,201 sample records,
including one valid 72-byte zero-stack variant. All raw records are counted
and the replay is offline; payload presence is not an unwind or owner claim.
The selected raw review is
`review/independent-raw-stack-tests-v1-final.json`; provisional raw reviews
remain historical.

The lifecycle helper is frozen at SHA-256
`6296e2dd20f911b2d989f1a8c7b9e337c9a34efc5c0f04831a50111192b9e7d0`. Its
author correction, lifecycle review v4, raw/lifecycle/integration controls,
and retained child-2 build-ledger binding are copied. Review v4 says the
lifecycle helper is ready for separately owned decoder/reducer replay. The
final reparse receipt (`receipts/lifecycle-stream-v2-receipt.json`) binds the
helper and reparse script to the external lifecycle stream and records zero raw
loss, 9,191 records, and 7,201 sample descriptors. The stream itself remains
external.
The selected lifecycle review is `review/lifecycle-ledger-review-v4.json`, and
the final decoder guard review is
`review/focused-decoder-lifetime-guard-review-v2.json`.

The final reducer source is copied at SHA-256
`1c40193d0debcea9feb696d82f029cc058b10e39e20ceb2183e2444223fc756c`.
The v4 final review (`review/independent-reducer-v4-final-review-v1.json`) and
scoped approval v3 pass a string-only owner-bounds
correction over the validated v2 packet proof; the packet-validation v4 result
remains the bound offline evidence. The earlier v2 source and review are kept
as the correction's inherited lineage, while `reduce_0978_v3.py` is an
external pre-label snapshot and is not selected. No 0978 cohort, speedup,
timing, source ownership, causality, qualification, or optimization claim
follows from this offline proof.

The final decoder guard source is copied at SHA-256
`9596e246c7c627fb3810e7c6c6d3dda12af58b62477f1f1233654c4256e2dac9`.
Its guard receipt and independent review v2 pass the full-lifetime event-list
check without changing the canonical event contract. The retained fresh result
was generated with the original decoder SHA-256
`866723324881a57ce7c4282cdfd2dabe497397b1612e8d7c4f975cb7e9fe2134`; the
final v3 review validates the guard and does not rerun perf or decode. The
fresh child-2 result, original plan, and packet-validation v4 are compact
proof bindings; their raw packet and sidecars remain external.

The native-boundary review v4 and control-boundary equivalence v5 pass the
static retained-control and pre-capture checks. The copied runner and updated
plan-binding helpers bind decoder v3 and reducer v4, and the copied plan v3
and helper seal v3 bind those sources. Final readiness v5 and root GO then
authorized one serial Office run. The run result is recorded below; no retry or
selective replacement followed the first loss-gate failure.

The retained child-2 diagnostic remains bounded forensic evidence only. The
direct corrected `/usr/bin/perf script` receipt is the authoritative completed
offline run: return code 0, 7,201 samples, total period 7,201,000,000, 76
empty callchains, 945 marker samples, marker period 945,000,000, and raw and
terminal loss zero. The focused-decoder process receipt is preserved as
historical lineage: it ended with SIGTERM during its redundant raw perf-script
dump, and the correction receipt withdraws its earlier natural-rejection
interpretation. That interrupted process is distinct from the direct corrected
script run and the final offline diagnostic.

## Native capture outcome

The bounded custody summary is
[`analysis/native-capture-summary-v1.json`](analysis/native-capture-summary-v1.json).
The final readiness review, native-boundary review, root GO, exact runner
command, runner result, and terminal capture manifest are copied under
`review/` and `receipts/`. Children 1 and 2 reached their decoder and runner
gates with 955 and 950 marker samples respectively, raw and terminal loss zero.
Child 3 produced no decoder packet or raw-stack sidecar; its terminal report
recorded four lost samples, so the loss gate failed. The reducer was invoked
once offline and rejected the incomplete manifest with return code 2 without
writing an output. The independent post-capture audit records child 3's raw
`PERF_RECORD_LOST` counter values as 1 and 3 (two records, total 4) and its terminal
`PERF_RECORD_LOST_SAMPLES` count as four; those views are retained separately,
not added together. The remaining children were not launched.

## Evidence custody

The compact packet includes the policy review, protocol proposal, scanner and
independent mutation controls, lifecycle correction and review, final decoder
guard proof, final reducer validation, scoped approvals, native-boundary/
equivalence receipts, final pre-capture plan/seal, root authorization and
failed-capture custody, replay receipts, child-2 diagnostics, and bounded
helper sources. The retained 0977 plan, focused decoder, binary, raw captures,
and prior receipts are referenced through the sealed 0977 packet descriptor in
[`provenance/dependency-0977-v1.json`](provenance/dependency-0977-v1.json);
they are not duplicated here.

The full perf data, mmap dumps, corrected perf-script text, DWFL logs, native
packets, forensic sample ledgers, lifecycle streams, build-ledger outputs,
derived ledgers, target executable, decoder shared object, and other compiled
artifacts remain external. Their exact size, SHA-256, role, and regeneration
references are in
[`provenance/external-artifact-descriptor-final-v1.json`](provenance/external-artifact-descriptor-final-v1.json).
No raw capture, full ledger, compiled executable, shared object, or file at or
above 1 MiB is copied into this packet.

The decoder integration outputs, lifecycle/reducer WIP variants, provisional
reviews, and superseded scanner/replay variants remain external historical
lineage. They are not presented here as approved 0978 qualification evidence.
The independent post-capture audit and temporary-cleanup receipt are copied;
the final descriptor and packet manifest bind the completed custody set. The
manifest excludes itself and the descriptor excludes both manifest and
descriptor self-reference.

## Reproduction note

[`provenance/reproduction-v1.md`](provenance/reproduction-v1.md) lists the
read-only offline commands, the historical capture command, and their
fresh-output requirements. It records the native command for provenance; do not
rerun sealed paths, retry child 3, or replace children. The packet was assembled
without rerunning the workload, rebuilding, or changing the source tree.
