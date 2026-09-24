# Reference-metadata performance capture analysis

This is the single authorized frozen-source capture (session `2878`) using three warmups and fifteen fresh child samples in both `evaluate` and `parse-evaluate`. The host context is retained in `capture-context-before.json`; timing was not isolated from the host.

## Capture and verification

- Baseline: `049c09cdde3978593149079c4257df047a3fa419`, 840 rows across 28 controls.
- Candidate: frozen checkout `/var/tmp/litchi-reference-metadata-gates-7d9uic94/checkout`, 3,810 rows across 28 controls and 99 metadata cases.
- Contract: `87f20c419139bf721221d06c8714edd900ae7a57cbedecf070878e4f6cf64d8a`; gate lock: `58b4be6c…40a3e3`; candidate freeze SHA-256: `e9d8dbac4eb9d9fe964693ccb801e74a031f7c223b3595e5003de03e36f02288`.
- Built-in verification: PASS. Preflight, source/profile stability, raw receipt presence, allocator balance, target cleanup, typed outputs/failures, and read bounds all passed.

## Matched controls

- Latency delta range: `-3.854%` to `+5.865%`.
- RSS delta range: `-4.989%` to `+4.311%`.
- Positive latency review flags over 5%: `1`; RSS flags over 5%: `0`.
- Allocator-call, byte-throughput, work, and resolver-read mismatches: `0` matched groups each.

The one positive latency review flag is `reference-conditional-256x4-sumifs` in `evaluate` at `+5.865%`; it is retained as an observation rather than treated as a causal regression. Its fifteen child p50s span `77,935..91,085` ns/repeat for baseline and `78,500..92,890` for candidate, with medians `83,295` and `88,180`; `parse-evaluate` is `-1.179%`. A simple independent 100,000-resample bootstrap over those fifteen child medians gives a descriptive 95% interval `[+0.528%, +6.831%]`. Work (`1,839`), reads (`1,792`), allocator calls (`42`), and bytes/repeat (`52`) are identical, so the timing observation has no matching accounting change. The host was not isolated: the pre-capture load average was `[0.91, 1.93, 2.25]` on 32 available CPUs, and the retained process snapshot showed no matching cargo, rustc, boundary, or evidence workload.

## Candidate metadata lanes

The candidate-only matrix has `99` cases and `198` phase groups. `196` groups have zero resolver reads; the two groups with one read are the `SHEET(ABS(range))` computed-scalar lanes, where the child `ABS` performs the one read before metadata evaluation. Resource, cancellation, shape/refusal, and source-policy rows remain typed and preflight-checked.

## Host and limits

The pre-capture load average was `[0.91064453125, 1.92626953125, 2.2451171875]` with `32` CPUs available to the capture process. No matching build/test/boundary workload appeared in the recorded process snapshot, but this was not an isolated host. RSS observations therefore remain workload-sensitive; the profile does not attribute them to the evaluator or allocator without further evidence.

Raw receipts, environment/source manifests, preflight logs, the generated performance report, and this analysis are retained under `results/`.
