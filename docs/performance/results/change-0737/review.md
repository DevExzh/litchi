# 0737 lifecycle probe review

## Scope

This review covers the lifecycle qualification requested by the 0736 audit. It
is a review of the probe contract and the proposed capture schedule; it does
not authorize a production edit, a native run, or a new performance claim.
The accepted owner and the complete 0735 preservation oracle remain the
comparison baseline. The archived 0735 probe must remain available as an
independent anchor.

The review is conditional until the lifecycle probe source, its independent
validator, and the capture receipts are available. A green process exit or a
top-level `oracle_ok` field is not enough evidence for this experiment.

## Required exact-oracle contract

Every arm must bind the same source bytes, expected-output bytes, replacement
set, fixture, operation, and policy. The report must carry and the validator
must compare the source and expected hashes, replacement digest, inventories,
changed-length proof, expected oracle, corruption-control results, and all
sample indices. For every measured sample, the validator must require the
complete 0735 oracle projection, including every nested boolean, an empty
`failure_reasons` list, the semantic witness, directory metadata and raw
directory policy fields, stream-path and stream-byte comparisons, CLSID
preservation, and `oracle_ok == true`. It must reject missing, defaulted,
truncated, or extra-shaped oracle values rather than treating them as an
acceptable compact result.

The output digest and output inventory for a given fixture and sample index
must agree across arms wherever the owner operation is deterministic. Timing
fields are allowed to differ; correctness fields are not. Every 50-sample arm
must contain exactly indices `0..49` in recorded order, with the complete
ordered sample sequence retained for later window diagnostics. A strict arm
must also retain an explicit receipt for each validated warmup. The receipt
must state its index, operation result digest/inventory, full-oracle status,
and lifecycle arm; a warmup count alone cannot prove that validation ran.

For strict-drained, the implementation may drop the full Rust witness only
after the full `output_sample`/`oracle_for_output` path has returned and all
checks have passed. The receipt is derived from that result and must preserve
the exact visible oracle projection used by the sealed fixture. The semantic
payloads skipped by JSON serialization remain source-audited evidence rather
than independently reconstructible packet fields; the report must state that
limitation. A separate payload fingerprint is optional and is not a required
control, since it would add work to both strict arms without proving the
validation ordering by itself.

The source must make the lifetime boundary explicit: the owner output is
timed first; validation and receipt construction are outside that interval;
the full witness, output inventory, reopen snapshots, and output bytes are
dropped before the next owner call in strict-drained. Strict-retained must
keep the complete current `Sample` witness across iterations while using the
same receipt path. Source/expected inventories and the expected oracle may
remain process-owned in both arms. A lexical helper returning only the
compact receipt is preferred so a compiler lifetime extension cannot silently
keep the full witness alive.

The legacy arm must preserve the original 0735 lifecycle exactly: three owner
only warmups, no per-output oracle during warmups, and full measured
`Sample`-witness retention. Strict-retained and strict-drained must validate
every warmup with the same full oracle path used for measured outputs. The
owner timer and allocation region boundaries must remain unchanged; oracle,
serialization, and receipt work must not enter `whole_ns` or an allocation
region. Allocation outputs must receive the same full oracle after the region
closes.

The independent audit should include negative controls that alter one field at
a time: top-level and nested oracle booleans, semantic witness content,
failure reasons, expected/source/output hashes, receipt digest, sample index,
sample count, warmup receipt, arm name, and fixture identity. Each alteration
must be rejected. This catches validators that accept a malformed oracle or
silently trust a report's own summary.

## Arms and schedule

The proposed native qualification has five 50-sample arms: the archived 0735
binary, independently rebuilt `legacy-a`, independently rebuilt `legacy-b`,
`strict-retained`, and `strict-drained`. All use three warmups and 50 measured
samples. The two rebuilt legacy arms are A/A controls, not treatments; their
purpose is to expose build/codegen or binary custody effects. Archive,
legacy-a, and legacy-b must be source/oracle-equivalent before any lifecycle
comparison is interpreted.

Use nine rounds per fixture (three cycles with three repeats), both PPT
fixtures, serial execution, CPU 12 pinning, and rotated arm/process order.
The native plan is therefore 90 processes (5 arms × 9 rounds × 2 fixtures).
Add the fresh strict-drained one-sample/three-warmup control as 18 more
processes (9 rounds × 2 fixtures). Its first-sample result is a process-level
control; it must not be presented as a 50-sample distribution or compared by
tail precision with the 50-sample arms.

The allocation plan is separate: archived/legacy no-warmup controls plus
strict-drained zero-warmup, strict-drained three-warmup, and strict-retained
three-warmup arms, three repeats per fixture. This is 30 processes (5 arms × 3
rounds × 2 fixtures). Report every allocation field, keep the region boundary
unchanged, and make no latency claim from allocation counters. The schedule
manifest must make the `samples/warmups` pair explicit for every arm; an
ambiguous label such as `strict-drained1/3` must be normalized to separate
numeric fields.

The total planned capture is 138 processes. Each process receipt must record
fixture, arm, binary identity, source/probe hash, round, repeat, execution
order, CPU affinity, lane, sample count, and warmup count. Pair comparisons
must use the predeclared round/fixture process pair as the unit. Individual
samples are ordered diagnostics inside a process, not independent bootstrap
units. Retain full-window p50, mean, p95, p99, and maximum; first-ten,
middle-thirty, and last-ten windows are descriptive and explicitly
non-independent. Do not choose a favorable window or replace process-pair
statistics with within-process samples.

The primary lifecycle contrast is strict-drained versus strict-retained:
same binary, owner operation, validated warmup count, full oracle cadence,
fixture, and receipt construction, differing only in whether the full sample
witness survives to the next owner call. Legacy versus strict-retained also
changes warmup validation and therefore cannot isolate retention. The fresh
control addresses process-local trajectory; the zero-versus-three allocation
control addresses combined validated-warmup and allocator state. These arms
can establish sensitivity to named lifecycle conditions, but they cannot by
themselves establish why the 0735 candidate regressed.

No candidate reintroduction is admissible until the unchanged-owner A/A and
oracle/lifecycle controls pass. A candidate comparison must use a newly frozen
source, binaries, fixtures, schedule, and exact same preservation constraints.

## Review disposition

The design is sound if the probe and validator implement the exact contract
above. Before capture, review the new probe diff against the archived 0735
source and reject the run if it changes the public operation, oracle scope,
timing boundary, allocation boundary, or fixture policy. After capture, the
packet must preserve raw JSON/stderr, manifests, binary/source hashes,
negative-control receipts, and an exact artifact inventory. Production must
remain byte-identical to the accepted baseline throughout.

After the final controls-probe source review and the strict receipt-schema
contract update, there is no remaining pre-freeze validation blocker. The
hidden skipped-payload limitation is recorded as an evidence boundary, while
the source ordering and exact visible-oracle checks are sufficient for this
qualification.

## Current implementation findings before freeze

The controls probe has the lifecycle loop and explicit `drop(sample)` for
strict-drained. Its `strict_receipt` is `serde_json::to_value(sample)`, so the
receipt omits semantic witness payloads marked `serde(skip)` and retains the
visible semantic witness in both strict arms. This is a documented audit
limitation of JSON evidence, not a validation bypass: `output_sample` invokes
the unchanged full `oracle_for_output` and `semantic_reopen_check` before
`strict_receipt`, and strict-retained/drained share that path. Source custody,
the real-fixture parity qualification, and exact visible-oracle equality to
the sealed fixture establish the validation ordering. The packet must state
that the offline JSON cannot reconstruct hidden heap ownership; a new payload
fingerprint is not required because it would add work to both controls without
proving the ordering by itself.

The controls probe adds `lifecycle`, `warmup_receipts`,
`retained_witness_count`, and a per-sample count to the default legacy JSON
shape. That means its legacy output is not byte/schema-identical to the
archived 0735 output even though the measured owner is unchanged. Either keep
the legacy serializer exactly compatible and put lifecycle metadata in a
separate envelope, or bind the extended schema explicitly and correct the
README/contract claim. The archive remains the strongest legacy anchor.

The Python contract now validates lifecycle names, warmup receipt
count/content, retained-witness counts, receipt field sets, phase/allocation
fields, and exact visible oracle equality. Keep the field-mutation negative
controls listed above. The plan's 108 native and 30 allocation processes,
including the 18 one-sample fresh controls, are internally consistent; the
schedule must still retain the explicit numeric lane/arm/sample/warmup
contract for every row.
