# 0699 — marker-free refusal-path attribution

Status: diagnostic complete; production unchanged. `performance_claim: none`.
Baseline revision: `e85f6a461e`.

The [0698 candidate](0698-mce-borrowed-inherited-namespaces.md) remains rejected.
This batch explains why its early-refusal regression should not be attributed
directly to borrowed namespace handling and identifies a different measured
hot loop. The early-name fixture takes the unchanged marker-free MCE return;
it never enters the modified start handler for those slide payloads. Its
separate XML reader detects the duplicate attribute. [Source path review](results/change-0699/path-review.md)
records the call chain and its limits.

Fresh diagnostic binaries compare current production with the exact rejected
codec in an isolated sparse checkout. The main source is unchanged. The new
probe adds single-case selection and a profile loop, so its code layout differs
from the original 0698 executable. Neither improved pairs nor further outliers
replace that experiment's recorded decision.

## Repeated timings

Twelve isolated process legs use ABBA, BAAB, ABBA order with three cases,
300 samples and ten warmups each. Case order reverses on alternate legs.
A separate four-leg ABBA run uses the original ten-case matrix. Together these
retain 22,800 timed captures, 76 case/leg distributions and 38 paired comparisons.
Fixture construction, result inspection, error formatting and snapshot drop
are outside each capture timer. All ten prepared graph, archive and exact error
identities are checked against the 0698 matrix. The host is shared and warm;
no cold-cache or quiescence claim is made. Per-process distributions and pair
spread expose uncertainty without treating samples within one process as
independent evidence about process-to-process variation.

| Isolated case | Pair 1 | Pair 2 | Pair 3 | Pair 4 | Pair 5 | Pair 6 |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| early-name-error | +1.70% | +0.06% | +9.36% | +1.19% | +0.87% | +1.13% |
| late-root-error-mce | +1.20% | +1.57% | +0.87% | +2.52% | +0.11% | +2.67% |
| small-valid | +0.77% | +0.54% | +0.83% | +0.15% | +1.59% | +2.17% |

The isolated early-error baseline medians span 34.40–34.93 µs. Five candidate
legs span 34.91–35.04 µs; one is 37.62 µs. That pair costs +9.36% (+3.22 µs),
with mean +9.61%, p95 +9.93% and p99 +12.85%. A fresh process thus reproduces
the outlier without needing earlier cases to execute in the matrix. This
rules out prior timed matrix cases as a necessary condition, but does not
identify the cause. Fixture construction still prepares all cases in both modes.

| Full-matrix case | First paired median | Second paired median |
| --- | ---: | ---: |
| early-name-error | +0.28% | +9.53% |
| generated-12x8-valid | -1.10% | +2.26% |
| late-missing-relationship | -1.11% | +3.32% |
| late-missing-relationship-mce | -1.03% | +1.59% |
| late-root-error | -0.61% | +3.97% |
| late-root-error-mce | -0.77% | +1.77% |
| mixed-conformance | -1.28% | +3.27% |
| notes-invalid-tail | -2.25% | +0.32% |
| slide-raw-overlimit-16m-to-64m | -0.02% | -0.67% |
| small-valid | +0.01% | +0.15% |

The second matrix early-error candidate is 37.91 µs versus baseline 34.61 µs:
+9.53% (+3.30 µs), with mean +7.18%. The final baseline p95 itself rises to
44.38 µs, limiting the paired tail interpretation. The seventh and remaining
>5% flag is an MCE missing-relationship p99 cost of +31.49% (+48.51 µs).
All seven flags remain in [triggers.json](results/change-0699/triggers.json).
The source path and repeated outliers justify investigating process-dependent
layout or state; neither proves a particular allocator, ASLR, hashing or host
mechanism. No regression is dismissed as noise.

## Instruction and profile attribution

The early-error self profiles show `process_markup_compatibility` at
7.25% baseline and 7.80% candidate, plus libc `memcmp` at 42.89% and 43.96%.
No start or Inherited destructor entry is sampled in those early-error profiles.
The MCE-marked late-root positive control does sample start (15.22% baseline,
13.34% candidate) and the baseline Inherited destructor (4.09%). These are
separate profile denominators, not capture-time fractions or additive speedups.
Profiles include setup, assertions, error formatting, result destruction and
a running checksum; the native timers do not.

[Bounded disassembly and annotations](results/change-0699/annotations.json)
identify the marker search loop at offsets 0xc0–0xdf of the processor symbol.
It loads length 0x3b (59), calls bcmp, advances the input pointer one byte and
repeats. Baseline instruction samples place 257 of 261 processor samples in
that loop; candidate samples place 288 of 290 there. These counts are sampled
instruction evidence, not a count of comparisons or an estimate of the time
saved by removing them. They establish the loop's weight within this symbol;
they do not assign every libc memcmp sample to this call site.

The processor symbol has the same 6,305-byte size in both binaries but moves
from 0x13d0f0 to 0x13d080. Its marker-search operations remain the same apart
from relocated addresses. The start symbol shrinks 18,083 → 17,833 bytes.
Changed layout is observable, but causal attribution of the slow process legs
requires another experiment.

Three (11,000 minus 1,000)/10,000 early-error counter slopes show candidate
cycles +0.90%/+1.35%/+0.17%, task clock +0.90%/+1.34%/+0.13%, and instructions
−0.077%/−0.061%/−0.157%. Branch misses change +6.67%/+3.00%/−1.34%; cache
misses −24.49%/+51.09%/+78.84%. These diagnostic samples are variable and do
not establish memory or cache savings. One baseline page-fault slope is zero,
so its percentage comparison is null. Raw values and all slopes are retained.

## Decision and next experiment

Keep production unchanged and retain the diagnostic harness. The next coherent
hypothesis is to replace the byte-by-byte namespace URI window scan with the
existing safe substring-search dependency. `memchr` is already a direct shared
OOXML dependency. This follows measured unnecessary work rather than a new
manual SIMD implementation. Preserve input-limit ordering, output-limit checks,
Cow ownership, complete Reports, and the exact behavior when the URI appears
anywhere—including comments, text or malformed XML. Search parity must cover
short buffers, near matches, boundary positions, arbitrary bytes and repeated
prefixes before a new representative A/B decision.

Namespace frame ownership should remain unchanged for that experiment. The
rejected borrowed-view timings are not evidence of a retained production gain.
A marker-search change must be measured on real edit/no-op workflows as well as
marker-free, MCE-positive and refusal controls; this diagnostic does not itself
establish the broader ROI.

## Reproduction and verification

The [packet README](results/change-0699/README.md) describes the sequence.
The environment records Rust 1.95 and the shared AMD EPYC 9R45 host. Receipts
bind 8,579 tracked crate files, accepted ADR/goal hashes, root build inputs,
identical probe files and both binary hashes. The only candidate source change
is the already-reviewed codec witness, applied in the disposable checkout.
The root and isolated source states were verified restored before gates.

Twenty CLI checks cover bounded counts, unknown cases and successful minimal
runs on both binaries. The nine applicable gates are workspace/probe formatting,
warning-denied probe Clippy and the six repository evidence checks. No
production Rust changed, so existing production tests are not represented as
fresh executions. The focused retained-source result from 0698 remains the
prior exact-source test evidence. No new fuzzing, native Office or Miri claim
is made. Final audit, review and cleanup receipts accompany this record.

All nine gates, twenty CLI checks and the complete pre-cleanup audit passed.
Independent review retains the diagnostic interpretation and leaves the 0698
rejection unchanged. Cleanup removed exactly the detached worktree, target,
frozen binaries and raw profile directory, preserving the root lockfile and
all production sources. The post-cleanup audit and four final documentation
gates are recorded in [final-validation.json](results/change-0699/final-validation.json).
