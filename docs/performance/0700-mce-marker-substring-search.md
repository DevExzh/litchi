# 0700 — substring search for the MCE namespace marker

Disposition: rejected; production predicate restored before commit. `performance_claim: none`.
Baseline revision: `6dfe64739f`.

[0699](0699-marker-free-refusal-attribution.md) identified a hot marker-presence
loop: the processor compares a 59-byte window, advances one byte and repeats.
This batch tests the existing safe `memchr::memmem::find` API in place of that
predicate. The crate already depends on memchr. No namespace-frame ownership,
XML reader, streaming API, cache, public API or dependency changes are part of this candidate.

The input-limit check remains before search. If no exact URI bytes occur, the
same output-limit check, borrowed input and empty Report must be returned.
Any occurrence must route to the existing parser, including comments, text,
malformed XML or a final-position occurrence. No validation is added to the
marker-free path. Exact bytes, Cow ownership, whole Reports, typed errors and
error precedence remain binding.

## Method

The packet freezes baseline and candidate native, allocation, refusal and
shared-MCE oracle binaries separately. Initial A/A completes before edits;
A/B/B/A then uses those frozen binaries on CPU 12. Thirteen PPTX workflows
cover one edit, no-op and two edits across the real deck, same-length
marker-stripped control, generated input and two notes fixtures. The native
phase timers cover capture, clone, text change, commit and apply; initial open,
target search, serialization and reopen checks are outside. They do not measure
complete save latency.

All initial A/A total medians differ by −0.86% to +0.45%. This is a warm shared
host measurement, not a quiescent or cold-input claim. Native and allocator
instrumentation are separate. Per-leg raw distributions, tail metrics and
all regression triggers remain visible; a pooled mean cannot waive a cost.

The shared oracle compares exact output identity, ownership, Report and errors
across real Office XML, deterministic mutations and synthetic cases. Existing
ordinary/opaque, real-part and declaration-heavy controls are supplemented by
short, repeated-prefix and late-marker search controls. The larger goal and
all accepted ADR constraints retain their previously read, unchanged hashes.

The [packet README](results/change-0700/README.md) records reproduction,
provenance and the final artifact inventory.


## Initial workflow and refusal results

Both baseline and candidate pass all 101 focused MCE tests, including the five
new marker-routing and limit cases. The initial matrix retains 7,800 native
workflows. All 78 allocation comparisons have identical counts, requested
bytes and measured live-byte metrics.

| Workflow / input | Baseline medians (ms) | Candidate medians (ms) | Paired change |
| --- | ---: | ---: | ---: |
| one-real | 11.0420 / 11.0090 | 11.3998 / 11.2813 | +3.24% / +2.47% |
| one-control | 6.0218 / 5.9074 | 5.2205 / 5.1607 | -13.31% / -12.64% |
| one-generated | 1.5967 / 1.5532 | 1.3895 / 1.4021 | -12.97% / -9.73% |
| one-notes-poi | 1.0364 / 1.0903 | 0.9067 / 0.8974 | -12.51% / -17.69% |
| one-notes-lo | 1.6449 / 1.6284 | 1.4178 / 1.4248 | -13.81% / -12.51% |
| noop-real | 5.1831 / 5.1822 | 5.3216 / 5.2833 | +2.67% / +1.95% |
| noop-control | 2.8813 / 2.8572 | 2.5045 / 2.4810 | -13.08% / -13.16% |
| noop-generated | 0.7651 / 0.7536 | 0.6729 / 0.7252 | -12.04% / -3.76% |
| noop-notes-poi | 0.4706 / 0.4724 | 0.4048 / 0.4034 | -13.98% / -14.60% |
| noop-notes-lo | 0.6040 / 0.6088 | 0.5154 / 0.5108 | -14.67% / -16.11% |
| two-real | 11.3416 / 11.4059 | 11.7626 / 11.7387 | +3.71% / +2.92% |
| two-control | 6.1166 / 6.2024 | 5.3438 / 5.2524 | -12.63% / -15.32% |
| two-generated | 1.8111 / 1.7836 | 1.5854 / 1.5919 | -12.46% / -10.75% |

| Real one-edit phase | Baseline medians (ms) | Candidate medians (ms) |
| --- | ---: | ---: |
| capture | 4.7642 / 4.7416 | 4.9009 / 4.8498 |
| clone | 0.0166 / 0.0166 | 0.0164 / 0.0166 |
| settext | 1.0163 / 1.0190 | 1.0608 / 1.0522 |
| commit | 5.1370 / 5.1357 | 5.3264 / 5.2660 |
| apply | 0.0914 / 0.0913 | 0.0920 / 0.0912 |

| One-edit allocation diagnostic | Baseline | Candidate |
| --- | ---: | ---: |
| alloc_calls | 143,848 | 143,848 |
| realloc_calls | 4,870 | 4,870 |
| requested_bytes | 12,430,874 | 12,430,874 |
| peak_above_start | 459,170 | 459,170 |
| net_live_change | 195,730 | 195,730 |

The real deck regresses in all three workflow shapes; the other ten median
pairs improve. Fifteen native >5% flags are retained. One is a whole-workflow
metric: real no-op p99 +8.10% (+424.43 µs) in the first pair. The other fourteen
are phase metrics, including real no-op capture p95 +5.01% and p99 +9.31%.
Clone, apply and no-op commit tails also flag. None is hidden by a group mean.

| Case | Baseline medians (µs) | Candidate medians (µs) | Paired change |
| --- | ---: | ---: | ---: |
| small-valid | 247.00 / 253.77 | 220.88 / 220.19 | -10.58% / -13.23% |
| generated-12x8-valid | 632.22 / 631.28 | 547.52 / 548.17 | -13.40% / -13.17% |
| early-name-error | 35.58 / 35.47 | 19.79 / 19.74 | -44.38% / -44.35% |
| late-root-error | 80.75 / 80.19 | 64.20 / 64.13 | -20.49% / -20.02% |
| late-root-error-mce | 150.72 / 150.21 | 147.23 / 146.95 | -2.32% / -2.17% |
| late-missing-relationship | 72.83 / 72.84 | 61.47 / 61.05 | -15.60% / -16.18% |
| late-missing-relationship-mce | 142.64 / 142.66 | 144.41 / 144.01 | +1.24% / +0.94% |
| notes-invalid-tail | 237.75 / 232.48 | 192.11 / 196.54 | -19.20% / -15.46% |
| mixed-conformance | 114.88 / 116.23 | 98.14 / 97.93 | -14.57% / -15.75% |
| slide-raw-overlimit-16m-to-64m | 22559.52 / 22531.62 | 270.76 / 271.48 | -98.80% / -98.80% |

All ten expected refusal/control results and graph identities match. The
16–64 MiB padded-slide case still reaches the same typed refusal; faster marker
search does not relax its later 16 MiB validation boundary. One refusal tail
flags: MCE late-root p99 +43.74% (+70.08 µs). The longer independent follow-up below
supplements these original results, which remain retained.

## Shared XML controls

The shared oracle reports zero mismatches across 192 cases, five profiles and
both binaries (1,920 invocations: 1,046 successful MCE results and 874 typed
error results, counting both sides). Corpus SHA-256 is
`a6f8abb6b2c504defa6bd879bcb6e31392a690976ba941d1e4d9a082f1bebf16`.
This is exact parity with the existing processor, not a new Office-validator
or native-application claim.

| Control / case | Profile | First median delta | Second median delta |
| --- | --- | ---: | ---: |
| synthetic: opaque | baseline | +0.58% | +2.26% |
| synthetic: opaque | opaque | +0.22% | +0.58% |
| synthetic: opaque | opaque-many | -0.43% | +1.31% |
| synthetic: ordinary | baseline | +0.99% | +6.16% |
| synthetic: ordinary | opaque | +0.23% | +1.31% |
| synthetic: ordinary | opaque-many | +3.77% | +0.08% |
| real part: docx | baseline | +1.22% | +1.24% |
| real part: docx | opaque | +1.91% | +2.09% |
| real part: docx | opaque-many | +2.76% | +3.34% |
| real part: pptx | baseline | +0.79% | +1.13% |
| real part: pptx | opaque | +1.86% | +0.78% |
| real part: pptx | opaque-many | +1.93% | +2.11% |
| real part: xlsx | baseline | +1.70% | +1.15% |
| real part: xlsx | opaque | +2.91% | +0.54% |
| real part: xlsx | opaque-many | +2.19% | +1.79% |
| declaration: declared | baseline | -0.42% | +1.15% |
| declaration: declared | opaque | -0.61% | -0.03% |
| declaration: declared | opaque-many | +0.02% | -2.45% |
| declaration: mixed | baseline | +3.95% | +3.35% |
| declaration: mixed | opaque | +3.21% | +2.67% |
| declaration: mixed | opaque-many | +0.65% | +1.86% |

These 37,800 isolated timings use 300 samples after ten warmups per leg.
The ordinary/default second pair exceeds 5% on all four primary statistics;
its first p99 pair is +9.12%. The DOCX many-name first p99 pair is +22.91%.
No declaration-control metric exceeds +5%. Most marked real-part medians have
a small positive cost, consistent with keeping the marked-path tradeoff open.

## Marker-search controls

Five additional deterministic cases use the same 300/10 AA/ABBA method and
exact output, ownership and Report identity checks, for 9,000 timings.

| Marker case | First median delta | Second median delta |
| --- | ---: | ---: |
| tiny-marker-free | +0.00% | +0.00% |
| tiny-valid-marked | +44.68% | +44.68% |
| large-near-prefix-marker-free | -99.04% | -99.05% |
| long-late-comment-hit | -96.68% | -96.44% |
| root-comment-hit | +45.83% | +42.86% |

Tiny marked medians rise from 470 ns to 680 ns in both pairs (+210 ns);
root-comment hits cost +220/+210 ns. All primary metrics for those marked
cases exceed 40%. Tiny marker-free medians remain 30 ns, although one mean
flags +6.18% (+1.73 ns). The 30 ns marker-free
measurements are close to timer resolution. These isolated timings are not
whole-document latency estimates; the consistent ~210 ns marked cost is a
concrete setup tradeoff. Repeated near-prefix and
late-comment cases show large search gains, without waiving that cost.

## Code, stack and counters

Native text grows 2,594,926 → 2,601,990 bytes (+7,064), data 60,880 → 61,152
(+272), and BSS 520 → 1,336 (+816). The processor symbol grows 6,305 → 7,565
bytes. Its recorded stack reservation grows 0x248 → 0x320 (+216 bytes).
Bounded assembly replaces the repeated bcmp window loop with the memmem search
setup/call path, including FinderBuilder::build_forward_with_ranker. Namespace
frame ownership is unchanged. This per-call reservation is not a whole-program
stack bound or a per-XML-depth allocation.

The single open-plus-capture counter slope changes cycles +0.82%, instructions
+0.19%, branches −0.24%, branch misses −8.25%, cache misses −6.32%, page faults
−10.06% and task clock +0.78%. It uses a different denominator from the edit
timers and is diagnostic rather than a repeated latency estimate. Whole-child
peak RSS is 5,656 → 5,712 KiB; no memory saving is inferred.

## Candidate validation

All seven integration gates pass on the frozen candidate: formatting, locked
all-feature checks, warning-denied Clippy, default tests (1,219 passed, two
ignored), all-feature Office tests (4,143 passed, 33 ignored), PPTX facade
tests (45 passed), and warning-denied rustdoc. No tests failed. These counts
include overlapping suites and must not be added as unique test coverage.
The baseline and candidate focused MCE suites each pass the same 101 tests.
The five added tests cover byte-search dispatch, borrowed ownership and limit
precedence; the shared oracle and marker controls provide additional exact
output, Report and error checks.

All six repository evidence gates also pass: crate boundaries, strict and
structural performance claims, report classification, CRUD coverage, and
non-iWork verification. Cargo-fuzz/nightly availability is recorded separately;
this packet does not claim a new fuzz campaign or native Office GUI validation.

## Independent follow-up and disposition

After all seven integration and six evidence gates finished, four additional
ABBA legs ran 300 samples after ten warmups: 7,200 native workflows across six
cases and 12,000 captures across the full ten-case refusal matrix. The separate
audit passes all 24 native and 40 refusal rows and 138 comparisons. Initial
measurements remain intact.

| Follow-up case | First median delta | Second median delta |
| --- | ---: | ---: |
| native: noop-real | +2.09% | +3.44% |
| native: one-control | -13.06% | -13.62% |
| native: one-generated | -10.50% | -10.97% |
| native: one-notes-poi | -13.52% | -13.18% |
| native: one-real | +2.22% | +1.47% |
| native: two-real | +2.42% | +1.71% |
| refusal: early-name-error | -43.99% | -44.44% |
| refusal: generated-12x8-valid | -14.70% | -14.38% |
| refusal: late-missing-relationship | -17.39% | -17.81% |
| refusal: late-missing-relationship-mce | +0.58% | +0.07% |
| refusal: late-root-error | -21.43% | -21.72% |
| refusal: late-root-error-mce | -1.73% | -2.92% |
| refusal: mixed-conformance | -16.64% | -16.93% |
| refusal: notes-invalid-tail | -17.60% | -16.67% |
| refusal: slide-raw-overlimit-16m-to-64m | -98.82% | -98.80% |
| refusal: small-valid | -12.90% | -11.93% |

Baseline native median drift is −0.37% to +1.12%; refusal drift is −1.72%
to +0.60%. Real-work costs repeat in both candidate pairs. The follow-up
trigger file preserves 37 flags, including extrema; 12 concern mean, median,
p95 or p99. Eleven are native phase p99 flags, including one-edit commit
+7.21% (+374 µs), two-edit capture +6.97% (+337 µs), two-edit commit
+5.11/+6.44% (+273/+342 µs), and two-edit set-text +12.87% (+158 µs).
The remaining primary trigger is small-valid refusal p99 +19.25% (+50 µs)
in the second pair. No whole-native-workflow primary statistic exceeds 5%
in this follow-up. The initial marked-root refusal p99 trigger does not repeat.

Reject the direct memmem replacement. It removes substantial marker-free
search work, including the large over-limit scan, while preserving semantics.
However, real edit/no-op workflows repeatedly slow by about 1.5–3.4%, tiny
marked XML adds about 210 ns, the real-part controls mostly get slower, and
the processor adds 216 bytes of stack reservation with no allocation saving.
The goal's 5% threshold calls for review rather than automatic rejection;
this decision weighs the recurring common marked-path cost against the
mechanism-control gains. It does not claim every tail outlier is causal or
that processor layout alone explains the real-work regression.

The measured candidate, source patch, binary hashes and validation receipts
remain in the packet. Only the five regression tests are retained in source;
production returns to the baseline predicate. These are rejected-candidate
measurements, not shipped speedups.

A subsequent experiment can test a lower-setup exact search using the existing
safe first-byte `memchr` primitive and `starts_with`, possibly in a private
outlined helper. Preserve arbitrary-byte matching, all limits and error
precedence. The repeated-prefix and marked controls must price that algorithm
as well; no improvement is assumed and no such change is implemented here.

The rejected codec was preserved and restored through `reject-candidate.py`;
`rejection.json` binds the candidate and final 602-file source maps. The final
source differs from baseline only in the five added tests. A fresh retained
focused run passes all 101 MCE tests with warning denial. The full integration
receipts describe the frozen candidate; the retained-focused receipt describes
the restored production codec plus the new tests.

The complete pre-cleanup audit passes, including exact script/source/binary
bindings, all raw matrices and independent control audits, patch replay,
three focused receipts and all candidate gates. Cleanup removes exactly the
batch target, binary and profile directories plus its generated marker-control
archive. The workspace Cargo.lock is preserved; unrelated scratch is untouched.
