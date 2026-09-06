# Validation and evidence boundaries

The full standalone release harness passes 378 tests, zero failed, one existing
ignored. Five new tests cover transfer rounding and overflow, successful capped
reads versus empty/EOF/zero-cap cases, source failures, counter overflow and
delta regression, and explicit CLI limits/provider mismatches. The existing
PPTX lifecycle equivalence test now includes a paced range source alongside
bytes, warm file and unpaced short range sources, checking exact output and
final budget release.

The candidate adds zero strict Clippy diagnostics relative to the retained
29-diagnostic harness baseline. The previous lint receipt/log/driver and full
source manifest are retained; only the two named tooling files differ.
All-feature/all-target harness and no-iWork workspace checks, warning-denied
harness rustdoc, explicit two-file rustfmt and crate boundaries pass. Production
Rust, public document APIs, dependencies, unsafe-code policies, accepted ADRs and
core/format I/O behavior are unchanged. No production retest claim is made from
these harness gates alone.

The first pilot used 3 samples/1 warmup, while the unchanged copied report oracle
accepts 1/0 controls or 30/3 formal runs. Its workload succeeded but its oracle
rejected that sample contract. The report, resource log, failed receipt and
original capture/gates scripts remain. Four subsequent 1/0 controls pass before
freezing; no retained formal sample existed yet. All 17 new pacing and inherited
output/timer/budget report corruptions reject before freeze.

The new report verifier validates transfer configuration, availability, u64
counters, read-count bounds, sum-of-ceiling bounds and exact adjacent snapshot
deltas. It then removes only the validated additive pacing fields from a copy
and invokes the unchanged 0431 provider oracle, preserving all prior cache,
source, budget, output, phase-order and timer checks. That is independent report
arithmetic verification, not an independent Office application round trip or
re-execution of producer refusal gates. Fixture/output identities are frozen
from valid pilots and checked across every formal report.

Requested transfer delay is not actual sleep time, physical network bandwidth,
a syscall count or a shared-link rate limit. The model has separate fixed and
transfer sleeps; OS scheduling may dominate small requests. The same underlying
logical reads and returned bytes must match across configurations, repetitions
and retained samples. CPU profiles contain untimed work and omit blocked sleep.
Operation allocator metrics are unavailable; managed boundary gauges and whole-
process RSS do not establish an operation high-water allocation bound. The sink
retains full output. No cold-cache, native-breadth, scaling, production speedup or
bounded-total-memory claim is authorized.

All 8 formal reports (240 samples) and 4 profile reports pass. Source read
counts, requested/returned bytes, short reads and histograms match exactly
across configurations, repeats and every sample. Twelve absolute 5% repeat
flags remain: three increases (paced plain publication p95/p99 and paced
media-rich open p99) and nine decreases. No API p50 or process-RSS repeat
crosses 5%. This is a descriptive simulation baseline, not a production latency
regression or stable-tail claim; all raw vectors and flagged comparisons remain.

Profile derivation initially rejected four unpaced sample headers without
callchains. They are now retained as missing-callchain self samples, with their
full 0.024307% sampled-period weight. The initial parser draft and failure remain.
The paced profile has no missing callchains; both exports have no other unparsed
lines. Unknown-symbol samples are retained separately. A verifier adaptation
also retained an incorrect zero pilot count; replay rejected it, and it was
corrected to the actual 1/0 pilot contract. Raw observations and frozen files
were unchanged in both corrections.

Requested transfer pacing totals 2.767330 ms for plain and 1,923.768399 ms for
media-rich in each retained sample. The additional per-read sleep, OS scheduling
and API work remain in measured latency; the observed paced-minus-unpaced
increase is not identified with these requested delays. Future timer-model
calibration should assess combining pacing deadlines before interpreting a
small-payload result as an ideal link-rate effect. No physical network behavior
is inferred.

The first exported-bundle probe caught a checkout-relative import in the verifier:
reusing capture.command imported the capture module's repository-root assumption.
The verifier now reconstructs the expected argv independently, without importing
the workload driver. That failed exported probe and verifier draft remain;
TemporaryDirectory removed the failed scratch export. Frozen files and raw
captures were unchanged.

The corrected exported bundle verifies before and after cleanup. Five
pre-cleanup and six post-cleanup custody/argv/seal/cleanup corruptions reject,
for 28 deliberate corruption probes including the 17 report mutations. Cleanup
removes only `/tmp/litchi-goal-0447-binaries` (459,723,864 regular-file bytes),
preserving both Cargo target directory identities and the pinned user-owned
GOAL.md digest. Final replay needs no original executable, Git, Cargo, perf or
workload execution. The full goal remains active; this batch adds no semantic
or native certification and does not alter representative-index statuses.
