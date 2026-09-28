# 0816 results review

**Disposition: pass with bounded interpretation.** This review checks the
retained 0816 reports and the official offline analysis after capture. It did
not build, execute a benchmark, run a workload, or use a profiler. The review
read `analysis.json`, `scaling.csv`, `scaling.md`, `observer.csv`,
`raw-audit.json`, `qualification-audit.json`, all three receipt sets, and the
draft batch report.

The final offline validator passes with cleanup checked: 648 reports, 13,320
samples, 72 scaling rows, six quality gates, and the 72-row independent raw
audit. The independent raw reader also passes byte-for-byte replay of its
retained output. Post-cleanup replay exposed two unfrozen reader-interface
errors recorded in `execution-notes.md`: the cleanup marker name was corrected
to `binaries_verified_before_removal`, and binary receipt comparison was
separated from the derived `custody_verified` marker. The cleanup witness,
binary identities, raw measurements, and numerical outputs were unchanged;
the corrected `--final` replay passes.

## Cardinality, schemas, and semantic oracles

The retained lanes have the frozen cardinalities:

* qualification: 72 reports and 72 samples, accepted before native capture;
* native: 432 reports and 12,960 samples from six blocks × 30 samples; and
* observer: 144 reports and 288 samples from two blocks × two samples.

Every report matches its receipt's route, shape, state, worker width, 64 KiB
task floor, and source settings. Local `0/0` reports use
`litchi.execution-baseline.v1`; nonzero cap/delay reports use
`litchi.execution-range-baseline.v1`. All member order, byte, SHA-256,
resource-limit, worker/I/O release, and CPU-task checks pass. The independent
audit reconstructs both 32-member payload shapes, confirms 72 paired curves,
and confirms output parity across lanes. Qualification timing is not imported
into native statistics, and observer timing is kept separate from native
timing.

The source counters are conserved across the controls:

* Large CFB returns 8,388,608 bytes. The local arm makes 32 calls with no
  short reads; the capped and delayed arms make 128 calls, request
  20,971,520 bytes, return the same 8,388,608 bytes, and record 96 short
  reads.
* Mixed CFB returns 8,130,560 bytes. The local arm makes 32 calls; the capped
  and delayed arms make 125 calls, request 20,320,256 bytes, and record 93
  short reads.
* Fresh large and mixed Parts make 64 logical calls and return 146,041 and
  143,782 bytes respectively, with no short reads. Primed Parts record zero
  timed source calls, requested bytes, returned bytes, and short reads in all
  three source arms.

All timed observer samples end with zero active reads, and maximum simultaneous
reads never exceeds the requested width. Mixed CFB and Parts remain at a
maximum of one simultaneous read under the 64 KiB task floor. Delayed large
CFB and fresh Parts reach a maximum of eight in their wider observer cases.
Primed Parts stay at zero. These are source-observer maxima, not native worker
occupancy or a measurement of how long a read remained active.

The capped arm is a real short-read control: its returned bytes and logical
output match local and delayed arms while its call/request counts change as
expected. Delayed and capped source counters are equal. The request histogram
has a single bucket for sizes zero through 64 bytes, so it cannot identify how
many calls were nonempty or turn the configured 250 µs sleep into a measured
sleep total. The delay is a caller-side in-memory simulation; it is not a
network RTT, filesystem latency, or physical cold-source measurement.

## Numerical cross-check

The report table agrees with the official `scaling.csv` p50 values and the
width-eight paired speedups and intervals for all 18 route/shape/state/source
families. Representative rows are:

* delayed large/fresh CFB: 40.047576 ms at width one, 5.440918 ms at width
  eight, speedup 7.363746 with interval [7.342908, 7.403719];
* delayed large/fresh Parts: 10.556259 ms to 1.860380 ms, speedup 5.674399
  with interval [5.497420, 5.726218];
* mixed delayed CFB: 39.114396 ms to 39.112992 ms, speedup 1.000266 with
  interval [0.999577, 1.001977];
* mixed delayed Parts: 10.531064 ms to 10.530473 ms, speedup 0.999988 with
  interval [0.999621, 1.000295]; and
* primed large Parts: 2.995 µs to 234.101 µs for local and 3.035 µs to
  247.712 µs for delayed, giving width-eight speedups 0.012739 and 0.011908.

The mixed curves are an admission-policy result: the one 4 KiB member is
below the 64 KiB task floor, and every mixed observer case sees one maximum
simultaneous read. Their near-one width ratios do not show a scheduler defect.
The primed Parts curves are a cache-hit control: zero timed source reads leave
the width-dependent scheduling cost visible, so their negative scaling cannot
be attributed to delayed source service.

The cap-to-local p50 controls stay close to one within shared-host variation.
Delayed-to-capped ratios rise when the operation performs source reads and
stay near one for primed Parts. These ratios describe this harness's source
arms and do not estimate a general remote-service cost.

## Analysis recovery and corrected binding

The first reader output is retained under `analysis-attempt-0/` with its
`recovery.json` provenance record. The current `analysis.json`, `scaling.csv`,
`scaling.md`, `observer.csv`, and `raw-audit.json` are bound to the corrected
readers. No capture, workload, report, receipt, or raw audit data was rerun or
changed. The corrected offline validator and independent raw audit pass with
648 reports, 13,320 samples, and 72 paired curves.

The first reader implemented the lower bootstrap endpoint through a percentile
call that selected sorted index 249. The corrected reader uses the frozen
explicit sorted indices 250 and 9749. Every one of the 72 speedup intervals
and the 144 source-control intervals has the same numeric endpoints
after correction because the adjacent bootstrap order statistics are tied.
All p50, p95, p99, mean, ratio, CPU, RSS, and source-counter values are also
unchanged. The corrected `scaling.md`, `observer.csv`, and `raw-audit.json`
are byte-identical to the archived outputs; the current analysis records the
corrected raw-audit reader identity.

The second reader correction makes model admissibility explicit. Eight of the
18 family fits, repeated across 32 of 72 width rows, have an unconstrained
Amdahl serial fraction outside [0, 1]. Corrected output therefore reports
`amdahl_fit_valid=true` for 40 rows and `amdahl_fit_invalid=true` for 32 rows,
with the unconstrained fraction, clamped fraction, residuals, and an explicit
reason retained. This changes interpretation metadata only; it does not mark a
capture or workload failure and does not change any measured or derived timing
number.

## Quantiles, CPU, RSS, and statistical limits

The raw reports retain p50, p95, and p99 for every native row. Each is a
nearest-rank process quantile over 30 samples in one child; p99 is therefore
the largest observed sample. The published row is the median of six process
blocks. Bootstrap intervals apply to paired p50 width and source-control
ratios, with 10,000 choice-based resamples, seed 816816, and endpoints
250/9749. They are not confidence intervals for p95, p99, CPU, RSS, or the
source delay.

The machine output flags 21 of 72 width rows with negative scaling, 40 with a
p99/p50 tail flag, and 54 with at least one greater-than-five-percent block
spread. The component spread counts are 26 p50, 41 p95, and 47 p99 rows. CPU
spread is flagged in 11 rows; no row receives the packet's RSS block-spread
flag. These flags retain shared-host variation rather than establishing a
before/after regression.

CPU is process-wide `ProcessCPUTime` around the operation and has a slightly
wider boundary than wall time. It includes clock and process scheduling
effects; delayed sleeps reduce CPU/wall while short cached operations can make
clock overhead prominent. RSS is whole-child peak memory from the capture
wrapper, including corpus construction, setup, warmups, priming, verification,
and teardown. It is not an operation-local allocation measurement. CPU, RSS,
tail, and spread values therefore support scope diagnostics only.

The corrected Amdahl fits are numerically finite, but 32 of 72 rows are marked
`amdahl_fit_invalid` because their unconstrained serial fraction is outside
the model-admissible [0, 1] interval; 40 rows remain admissible. The same
split is eight invalid and ten admissible fits among the 18 route/shape/state/
source families. The width-level apparent fraction lies outside [0, 1] in 21
rows. Primed Parts reach unconstrained values around 92–101 and width-two
apparent values above 128. The constrained fits, unconstrained fits, residuals,
and invalid reasons remain descriptive diagnostics; they do not partition
measured CPU time or justify a production optimization.

## Interpretation correction for the draft report

The draft's headline phrase that “primed Parts make zero source reads” must be
read as “zero source reads during the timed operation after the preload and
observer reset.” Priming itself reads and populates the source-backed cache.
The later source-counter paragraph and the observer table use the narrower
timed scope; retaining that qualifier in the headline avoids an end-to-end
source-read claim. Likewise, “one active read” and “eight active reads” should
be understood as maximum simultaneous observer reads, with active reads zero
after each operation. The recovery changed reader metadata only; no measured
timing, counter, quantile, ratio, interval, or table number changed.

The evidence supports a finite-budget, warm in-memory, low-level CFB/Parts
provider-scaling result. It does not establish physical cold behavior, real
network performance, native-format CRUD, cross-session contention, allocation
savings, or a universal scaling law.
