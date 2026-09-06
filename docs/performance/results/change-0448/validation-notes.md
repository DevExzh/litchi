# Validation and scope

The standalone release harness passes 381 tests, zero failed, one existing ignored.
Three new tests check minimum-service boundary/overflow behavior, exact short-read
bytes and nominal-counter equivalence across policies, and CLI policy requirements,
values and duplication. The existing real PPTX lifecycle equivalence test now
includes both policies and checks exact output and final caller-budget release.
No wall-clock microtest substitutes for the frozen full-lifecycle service-floor
verification.

The frozen report oracle retains all prior strict cache, budget, source, output,
phase-order and timing checks. It validates the policy setting, nominal pacing
arithmetic and counter deltas, then checks each API duration against the combined
fixed-plus-transfer floor. Producer refusal and semantic equality gates remain
Rust evidence; Python verifies reports, not an independent Office round trip.

The two policies share the exact same build, source/destination corpus identities,
range cap, fixed delay and transfer-rate target. Source reads, returned bytes,
short-read counts, histograms and nominal pacing counters must match every sample.
The comparison measures delay-model calibration. Minimum service includes elapsed
wrapped-source work and fixed-wait overshoot; it is not an additive transport
model. Targets are not actual sleep observations. No physical bandwidth, cold I/O,
shared-link concurrency, native compatibility or scaling result follows.

The sink retains full output. Operation allocator attribution is unavailable;
managed phase-boundary gauges and whole-process RSS measure different scopes and
are not allocation high-water evidence. CPU profiles include setup, warmups,
gates, diagnostics and reporting while omitting blocked sleep. Unknown symbols
and missing callchains remain in profile summaries rather than being discarded.

All-feature/all-target harness and no-iWork workspace release checks pass, as do
warning-denied harness rustdoc, explicit two-file rustfmt and crate boundaries.
Strict Clippy retains the same 29 rendered harness diagnostics with zero new
diagnostics. Its exact baseline receipt/log/driver/source manifest are imported
from sealed 0447; the candidate differs only in the two named tooling modules.
Production Rust and the prior sealed bundle are unchanged.

Four 1/0 pilots pass before the actual timestamped freeze. Twenty deliberate
report corruptions reject, including missing/unknown policy, inconsistent service
floor, pacing availability/arithmetic, output identity, timer and budget changes.
The frozen gate requires both plain median improvements to reach 5%, every API
service-floor check to pass, and explicit review of all 5% paired/repeat triggers.
No raw evidence or threshold is selected after observing the formal samples.

The final eight reports retain 240 samples and all four profiles pass. Plain API
p50 falls 19.299%/19.284%, exceeding the frozen 5% gate in both repeats. All 30
absolute 5% paired triggers are lower latency and explicitly retained in the
decision; no positive paired, RSS or repeat trigger exceeds 5%. All source work
and nominal counters match. Every service floor passes. Both profile summaries
have zero missing callchains and zero unparsed lines; unknown symbols remain.
The initial complete portable replay passes without corrections or discarded
captures. No failure drafts have been created in 0448; strict inherited lint is
the sole expected nonzero check.

Portable verification also passes from exported copies before and after cleanup.
Five precleanup and six postcleanup custody corruptions reject, for 31 total
report/custody corruption probes. The only owned copied binary directory was
removed after passing sealed precleanup replay: 459,741,448 regular-file bytes.
Both Cargo target directories and the pinned user GOAL.md are preserved. Probe
scratch copies remove themselves before receipts are written. Final sealing covers
180 members plus SHA256SUMS (181 bundle files). The full non-iWork goal remains
active; no native/default coverage status is promoted by this tooling batch.
