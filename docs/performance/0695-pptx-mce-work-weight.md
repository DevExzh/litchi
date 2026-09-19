# 0695 — weight remaining PPTX MCE work before expanding reuse

Status: completed attribution experiment; production unchanged.
`performance_claim: none`. Production revision: `df44030d6`.

The five remaining presentation XML passes are **not the largest remaining
MCE opportunity on the measured real deck**. Their combined isolated median
is 83.46–84.87 µs, versus 2.512–2.538 ms for the complete 18-call MCE sequence.
The presentation/full ratio of median-of-leg-medians is 3.33%; independent
hardware-counter slopes put it at 3.32–3.39% of cycles and 3.51–3.52% of
instructions. Slide 11 alone measures 1.082–1.094 ms, about 43.16% of the
full isolated sequence time.

These are kernel attribution ratios, not additive capture fractions or
achievable end-to-end speedups. They change the next priority: investigate
shared per-element namespace/context ownership work before extending
presentation-only reuse. No production optimization is introduced in this
batch. The [packet](results/change-0695/README.md) retains the reproducible
probe, exact inputs, raw samples, counter repetitions, source bindings,
independent audit and cleanup receipts.

## Corpus and timing

The retained 0693 first-real-capture trace identifies exactly 18 successful
MCE calls over 14 input owners: presentation XML three times, slides 1–13 once
each, then presentation XML twice. The pre-capture setup call is excluded.
Current PPTX callsites match the traced production source; 0694 changed only
element-name ownership in shared MCE. The preparation driver verifies the
historical instrumentation separately from current production source hashes.
No historical timing is treated as a fresh measurement.

All selected XML comes byte-for-byte from the retained LibreOffice
`slide-section-test.pptx` fixture. Five presentation calls process 11,785 bytes
(4.19% of the sequence's 280,963 bytes); the 13 slide calls process 269,178
bytes. Slide 11 is 121,498 bytes (43.24% of total sequence input). Pointer
identity is only historical per-process evidence. The new probe loads each
unique path once and reuses its owner index for repeated calls.

There are 16 cases: 13 individual slides and three groups. Each case has four
native process legs on CPU 12, with reversed case order in odd legs, ten warmup
batches and 200 measured batches of four executions. The 12,800 recorded batch
samples cover 51,200 sequence executions. Input loading, identity hashes,
printing and sample-vector allocation are outside timers. Default capability
construction, MCE processing, output destruction and loop overhead are inside.
The host is shared; no quiescence or cold-cache claim is made.

All individual and group results follow. Ranges are the four observed leg
medians, not population confidence intervals. Per-leg means, p95/p99 and every
raw sample remain in the packet. Since samples average four executions,
p95/p99 are batch-average tails, not single-call tails. The maximum observed
inter-leg median spread is 2.72% (slide 3).

| Case | Calls per sequence | Input bytes | Median range (µs) |
| --- | ---: | ---: | ---: |
| all | 18 | 280,963 | 2512.093–2537.890 |
| presentation | 5 | 11,785 | 83.457–84.868 |
| slides | 13 | 269,178 | 2416.434–2433.700 |
| slide1 | 1 | 2,589 | 21.242–21.455 |
| slide2 | 1 | 17,006 | 148.820–149.200 |
| slide3 | 1 | 3,038 | 25.018–25.698 |
| slide4 | 1 | 31,153 | 273.659–275.541 |
| slide5 | 1 | 17,833 | 157.648–160.996 |
| slide6 | 1 | 3,037 | 25.293–25.488 |
| slide7 | 1 | 28,449 | 251.182–253.606 |
| slide8 | 1 | 2,674 | 22.055–22.329 |
| slide9 | 1 | 2,676 | 22.128–22.358 |
| slide10 | 1 | 3,037 | 25.088–25.376 |
| slide11 | 1 | 121,498 | 1082.204–1094.441 |
| slide12 | 1 | 29,160 | 263.830–267.016 |
| slide13 | 1 | 7,028 | 60.065–61.140 |

The sum of individual-source medians (with the five-call presentation group)
is 98.08% of the separately measured full-sequence median; the two group
medians sum to 99.77%. Keep that mismatch: isolated allocator/cache/branch
states differ. Do not normalize the rows to force an additive partition.
The full-sequence probe also omits all intervening semantic consumers and
their allocation lifetimes. In particular, it is not a replacement for the
0694 native capture/commit measurement.

## Hardware and remaining work

Three independent 10/210-sample counter pairs, each with ten sequence
executions per sample, estimate per-sequence counters by subtracting the short
run from the long run and dividing by 2,000. All recorded event running
percentages are 100%; raw counters and every derived slope are retained.
The slopes retain sample handling and one output row per ten executions;
subtracting startup does not remove work proportional to sample count.

| Group | Cycles per sequence | Instructions per sequence | Task clock (ms) |
| --- | ---: | ---: | ---: |
| Full 18 calls | 11.297–11.319 million | 49.688–49.698 million | 2.497–2.520 |
| Presentation 5 calls | 0.375–0.383 million | 1.742–1.747 million | 0.0836–0.0853 |
| Slides 13 calls | 10.892–10.925 million | 47.931–47.997 million | 2.425–2.432 |

Branch, branch-miss, cache-miss and page-fault counts remain explicit in
`counter-summary.json`. Full-sequence slopes include about 33 page faults per
execution; this is a measured cost, not a memory bound. No new allocation or
RSS measurement is claimed.

A separate fresh full-sequence sampling profile reports 32.22% self samples
in MCE `start`, 6.46% in `Inherited` destruction, 6.16% in attribute value
decoding/normalization, and 3.27% in `Ctx` destruction. It reports no lost
samples. Its denominator is isolated MCE work, unlike 0694's 11.60% `start`
share in an open-plus-capture profile. Inline namespace-`Arc` clone annotations
suggest a follow-up, but imperfect DWARF call chains do not establish an exact
clone-only cycle share.

The smallest next hypothesis is to avoid `Namespaces::with_local` cloning the
existing namespace owner when the local declaration vector is empty. The
current empty branch returns `Ok(self.clone())`, immediately followed by
replacement of the old owner. It may pay a redundant increment/decrement pair
per element. This batch does not prove that removing it improves public
workflows; the next experiment must compare native end-to-end timings,
allocations, counters, custom profiles and exact refusal/output behavior.
Broader lifetime changes to `Inherited`, frame state or shared namespace layers
need a separate ownership proof. Presentation reuse remains a possible later
improvement, with a much smaller measured kernel weight on this deck.

## Validation and scope

The release probe is built with `--locked --offline` and warnings denied.
The successful build binds all 7,196 tracked crate Rust sources and Cargo
manifests, plus the exact probe and lockfile. All 33 previously read goal/ADR
constraint hashes remain unchanged. Production, public APIs, validation order,
resource policies and iWork sources are unchanged.

All nine validation gates pass: probe formatting, warning-denied Clippy and
rustdoc, plus the six repository boundary/claim/coverage/non-iWork gates. Output SHA-256,
length and ownership repeat across all native legs and agree between member
and group invocations. This is deterministic codec evidence on the selected
inputs, not new native Office interoperability or complete XML validation.
No production test suite is rerun solely for this measurement-only batch;
0694's production verification remains separately scoped to that commit.

The first probe compile failure concerned diagnostic hash formatting; the
successful source fixes that outside the timer. Raw build history is retained.
The independent audit verifies corpus/trace/build bindings, all native samples,
profile slopes and final gates. Cleanup removes only the owned 0695 target and
raw-profile directories, preserving the workspace lock and permanent evidence.
No CRUD coverage or registered performance claim is promoted.
