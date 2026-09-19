# 0698 — borrow temporary inherited namespace views

Status: candidate rejected; production codec restored. `performance_claim: none`.
Baseline revision: `8f16c57b2`.

The borrowed-view candidate improves representative successful workflows, but
is **not retained**: early duplicate-name refusal exceeds the +5% review
threshold in both measurement rounds (+15.12% initially, +7.84% in the
follow-up). The production codec is restored byte-for-byte to baseline.
Three regression tests and the complete experiment remain. All candidate
timings below describe the rejected implementation, not shipped gains.

The hypothesis follows [0697](0697-mce-context-ownership-attribution.md):
`Inherited` temporarily clones namespace owners already held by a parent parser
frame. Borrowing those owners for the current start event should remove two
reference-count increment/decrement pairs when present. The child context and
the namespace boundary retained by each child frame must continue to own their
state. No allocation reduction is assumed.

The [packet](results/change-0698/README.md) retains frozen before/after native,
allocator, refusal and shared-MCE oracle binaries by hash, reproducible drivers,
raw samples, semantic parity results, review and cleanup receipts. All accepted
ADR and goal constraints are bound to the previously read 33-file set. This
experiment adds no public API, persistent cache, executor or resource-policy
change.

The candidate must preserve the pre-local inherited namespace scope, the nearest
emitted ancestor, namespace rebinding and hoisting through removed compatibility
wrappers. Parent `AlternateContent` bookkeeping must finish before a borrowed
view is taken, and the owned child frame must be constructed before `close`
can mutate or reallocate the frame vector. Exact output bytes, `Cow` ownership,
the whole `Report`, typed errors and error precedence remain binding.

Measurements use the existing thirteen PPTX one-edit/no-op/two-edit scenarios,
ten refusal/control cases, the shared real/mutated/synthetic XML oracle, and
isolated default/custom-capability controls including declaration-heavy XML.
Native timing and allocator instrumentation use separate binaries. Initial
baseline A/A precedes edits; A/B/B/A uses frozen binaries on CPU 12. These are
warm measurements on a shared AMD EPYC 9R45 host using Rust 1.95;
[environment.json](results/change-0698/environment.json) records the toolchain
and host details. Native edit-phase totals exclude initial
open, save and reopen verification and do not establish complete-save latency.

The three new regression cases passed with the unchanged baseline codec before
the candidate was applied. Both versions pass all 96 tests selected by `mce::`.
They establish exact output and whole-Report behavior through selected
alternate content, prefix rebinding, a default namespace reset and unwrapped
content; visible and skipped opaque scopes; and nested error precedence.

Frozen native assembly removes the separate 118-byte `Inherited` destructor.
`start` shrinks from 18,083 to 17,833 bytes, while `Ctx` destruction remains
118 bytes. The `start` stack reservation grows from `0x598` to `0x618`, a
128-byte cost in this build. These are symbol/frame observations, not a
whole-program memory or stack bound; allocation and RSS results are separate.


## Initial native comparison

The 13 cases each have six native legs: baseline A/A followed by A/B/B/A,
100 samples after five warmups per leg. This retains 7,800 timed workflows and
468 phase/leg summaries. Native timing uses the default system allocator;
allocation diagnostics use separate binaries. The initial A/A total medians
range from −2.10% to +1.94%, and means from −2.06% to +1.97%. The real-deck
A/A medians are 11.1535/11.0952 ms. Raw legs, means, p95/p99 and within-leg
bootstrap median intervals are retained; intervals do not describe all
shared-host or between-process uncertainty.

| Workflow / input | Baseline medians (ms) | Candidate medians (ms) | Paired change |
| --- | ---: | ---: | ---: |
| one-real | 11.0486 / 11.0387 | 10.8052 / 10.8203 | -2.20% / -1.98% |
| one-control | 6.2094 / 5.9608 | 5.9963 / 6.0108 | -3.43% / +0.84% |
| one-generated | 1.5698 / 1.5652 | 1.5432 / 1.5671 | -1.69% / +0.12% |
| one-notes-poi | 1.0357 / 1.0349 | 1.1020 / 1.0498 | +6.40% / +1.45% |
| one-notes-lo | 1.6415 / 1.6429 | 1.6570 / 1.6507 | +0.94% / +0.48% |
| noop-real | 5.1733 / 5.1631 | 5.0746 / 5.1468 | -1.91% / -0.32% |
| noop-control | 2.8629 / 2.9068 | 2.8793 / 2.8758 | +0.57% / -1.07% |
| noop-generated | 0.7763 / 0.7579 | 0.7502 / 0.7482 | -3.37% / -1.28% |
| noop-notes-poi | 0.4707 / 0.4717 | 0.4729 / 0.4697 | +0.45% / -0.44% |
| noop-notes-lo | 0.6029 / 0.6043 | 0.6033 / 0.6104 | +0.06% / +1.00% |
| two-real | 11.3456 / 11.2920 | 11.1530 / 11.1033 | -1.70% / -1.67% |
| two-control | 6.0653 / 6.0808 | 6.0955 / 6.0963 | +0.50% / +0.26% |
| two-generated | 1.7860 / 1.7838 | 1.7571 / 1.7697 | -1.62% / -0.79% |

| Real one-edit phase | Baseline medians (ms) | Candidate medians (ms) |
| --- | ---: | ---: |
| capture | 4.7676 / 4.7670 | 4.6535 / 4.6514 |
| clone | 0.0164 / 0.0163 | 0.0160 / 0.0160 |
| settext | 1.0182 / 1.0174 | 1.0066 / 1.0176 |
| commit | 5.1420 / 5.1420 | 5.0301 / 5.0405 |
| apply | 0.0911 / 0.0913 | 0.0903 / 0.0902 |

| One-edit allocation diagnostic | Baseline | Candidate |
| --- | ---: | ---: |
| alloc_calls | 143,848 | 143,848 |
| realloc_calls | 4,870 | 4,870 |
| requested_bytes | 12,430,874 | 12,430,874 |
| peak_above_start | 459,170 | 459,170 |
| net_live_change | 195,730 | 195,730 |

All 78 allocation comparisons match across all three repeats: allocation and
reallocation calls, requested bytes, baseline/current/peak live bytes,
reallocation requested bytes, peak above the starting point and net live change.
Requested bytes already include full replacement sizes for reallocations;
those values must not be added again. The borrowed view removes reference-count
traffic, not heap allocations.

The initial native matrix has 31 >5% review triggers. Four are whole-workflow
metrics on `one-notes-poi` in the first candidate pair: median +6.40%, mean
+6.46%, p95 +6.95%, and p99 +6.68%; the second median pair is +1.45%.
The other 27 are phase-level metrics. The complete trigger list is retained;
no geometric mean hides the notes case. Longer follow-up measurements address
that case and real/control stability separately.

## Refusals and semantic parity

The initial refusal matrix has ten cases, six legs and 100 samples per leg:
6,000 samples with exact expected error and graph identities. First-pair review
triggers include early-name median +15.12% (5.33 µs), late-root median +6.05%
(4.85 µs), notes-invalid-tail p95 +5.23%, and raw-overlimit median +33.00%
(7.44 ms). Their second-pair median changes are +0.61%, −0.94%, −1.60% and
approximately zero respectively. These initial costs remain reported even if
later runs differ; the follow-up uses new processes and larger samples.

The shared oracle compares 192 inputs under five capability/limit profiles in
both frozen binaries: **1,920 invocations, zero mismatches**. The corpus has
36 real XML parts from twelve DOCX/XLSX/PPTX archives, 144 mutations and twelve
synthetic inputs. It compares exact output SHA/length, borrowed/owned `Cow`,
whole `Report` and typed `Debug` errors. Its corpus hash is
`a6f8abb6b2c504defa6bd879bcb6e31392a690976ba941d1e4d9a082f1bebf16`.
This proves parity with the existing processor on that corpus; it is not an
independent XML validator or a new native Office GUI interoperability claim.

## Shared-kernel controls

The following 300-sample AA/ABBA controls use the frozen oracle's separate
native timer. Input loading, capability construction and hashing are outside;
processing and output destruction are inside. Default, one-name (`opaque`) and
4,096-name (`opaque-many`) extension profiles expose capability-dependent costs.
These are individual XML kernels, not complete DOCX/XLSX/PPTX workflows.

| Control | Profile | First median pair | Second median pair |
| --- | --- | ---: | ---: |
| synthetic: ordinary | baseline | -8.83% | -6.39% |
| synthetic: ordinary | opaque | -4.30% | -4.95% |
| synthetic: ordinary | opaque-many | -4.68% | -3.93% |
| synthetic: opaque | baseline | -8.35% | -9.30% |
| synthetic: opaque | opaque | -10.72% | -9.23% |
| synthetic: opaque | opaque-many | -8.60% | -8.85% |
| real part: docx | baseline | -2.89% | -2.52% |
| real part: docx | opaque | -1.55% | -0.20% |
| real part: docx | opaque-many | -0.82% | -0.47% |
| real part: xlsx | baseline | -0.21% | -0.72% |
| real part: xlsx | opaque | +3.22% | +0.21% |
| real part: xlsx | opaque-many | +1.63% | -0.63% |
| real part: pptx | baseline | -2.77% | -3.68% |
| real part: pptx | opaque | -0.12% | +0.29% |
| real part: pptx | opaque-many | -0.41% | -0.86% |
| declaration: declared | baseline | -4.61% | -4.35% |
| declaration: declared | opaque | -4.25% | -3.15% |
| declaration: declared | opaque-many | -4.90% | -4.90% |
| declaration: mixed | baseline | -4.35% | -4.21% |
| declaration: mixed | opaque | +1.51% | +5.73% |
| declaration: mixed | opaque-many | +0.40% | -2.33% |

The ordinary/opaque synthetic matrix contains 10,800 timed samples; real-part
controls contain 16,200; declaration-heavy/mixed controls add 10,800. Output
identities and complete Reports repeat across all legs. Real parts are the
same deterministic marker-bearing members selected in 0696: DOCX numbering,
XLSX sheet1 and PPTX slide1. The declaration generator uses 1,000 redeclaring
siblings, adding an inherited child to each sibling for the mixed case.

One synthetic opaque/one-name p99 pair costs +21.41% despite improved medians.
No real-part metric exceeds +5%. Mixed declaration/one-name triggers in the
second pair are median +5.73%, p95 +5.80% and p99 +6.61%; that case is included
in a separate follow-up. All timing controls run separately from Cargo and
repository gates. No universal XML or tail-latency improvement is claimed.

## Counters, code and stack

The native open-plus-capture prefix has a different denominator from the edit
phase timers. A single 210-minus-10 iteration slope gives cycles −1.25%,
instructions −0.24%, branches −0.49%, branch misses +3.29%, cache misses +0.29%,
page faults −3.42% and task clock −1.22%. These are diagnostic counter samples,
not independently repeated latency estimates. Whole-child peak RSS is
5,664 → 5,560 KiB in the recorded process; no general RSS saving is inferred.

Native `.text` is 2,594,926 → 2,594,502 bytes (−424), data stays 60,880 bytes,
and BSS is 520 → 936 bytes (+416). Together with the +128-byte `start` stack
reservation, these costs remain explicit. The private `Frame`, `Ctx` and
namespace graph retain their owned representation. The event handler is not
recursive per XML element, but the recorded frame reservation is not a proof
of a whole-program stack bound.


## Longer follow-up

After all integration and evidence gates completed, a separate ABBA matrix
used the same frozen binaries with 300 samples after ten warmups per leg.
It adds 4,800 native workflows, 12,000 refusal samples and 1,200 mixed/opaque
XML timings. Four byte-exact oracle identity runs accompany the XML timings.
These samples supplement, rather than replace, the initial matrices.

| Kind / case | First paired median | Second paired median |
| --- | ---: | ---: |
| native: noop-real | -1.26% | -1.12% |
| native: one-control | -1.59% | +1.91% |
| native: one-notes-poi | +1.74% | +1.46% |
| native: one-real | -2.45% | -2.30% |
| refusal: early-name-error | +0.61% | +7.84% |
| refusal: generated-12x8-valid | -0.96% | +1.90% |
| refusal: late-missing-relationship | -1.08% | +1.55% |
| refusal: late-missing-relationship-mce | -2.15% | -1.56% |
| refusal: late-root-error | -0.72% | +2.44% |
| refusal: late-root-error-mce | -1.96% | -0.93% |
| refusal: mixed-conformance | -0.74% | -2.89% |
| refusal: notes-invalid-tail | -2.27% | -0.08% |
| refusal: slide-raw-overlimit-16m-to-64m | +0.01% | -0.11% |
| refusal: small-valid | -2.06% | +2.26% |
| oracle: mixed-opaque | -1.37% | -1.80% |

The real one-edit medians are 11.0998/11.0435 → 10.8276/10.7892 ms;
no-op medians are 5.1916/5.1721 → 5.1262/5.1143 ms. The marker-stripped
control changes sign between pairs. POI notes retain a small measured cost:
1.0352/1.0403 → 1.0531/1.0555 ms, or +17.97/+15.22 µs. Its initial +6.40%
median does not repeat, but the positive cost is not dismissed.

Early duplicate-name refusal remains flagged in the second pair: median
+7.84% (+2.75 µs), mean +7.78%, p95 +8.31% and p99 +6.61%. The first pair
is +0.61% at the median. This is an unresolved measured refusal-path cost,
not evidence of a universal improvement or proof of host noise. Exact errors
and prepared graph identities remain equal. The large initial over-limit
median regression does not repeat (+0.01%/−0.11%); the final baseline leg
instead has +32.38% p95 and +31.70% p99 relative to the first baseline leg.
That variability limits any tail-latency conclusion. Mixed/opaque XML
medians improve 1.37%/1.80% in the follow-up, while the original positive
control pair remains retained.

[Follow-up triggers](results/change-0698/followup-triggers.json) contain
21 flags because this driver includes min/max as well as median, mean, p95
and p99. Six concern those four primary statistics: the four early-name
refusal metrics above, no-op clone p99 +25.58% (+4.76 µs), and POI notes
apply p95 +17.29% (+5.29 µs). Fifteen are min/max observations. No follow-up
whole-native-workflow metric exceeds +5%. The initial 31 native and 12 refusal
flags, plus the synthetic/declaration control flags, remain available in
their original records.

## Verification

All seven integration gates passed on the measured candidate: formatting, all-feature checking,
warning-denied Clippy, default tests, all-feature tests, facade tests and
warning-denied Rust documentation. The default suite records 1,214 passes and
two ignored tests; the all-feature suite records 4,138 passes and 33 ignored;
the facade suite records 45 passes. These are separate suite totals with
overlapping coverage, not a count of distinct tests. The new focused tests
passed on both baseline and candidate (96 each).

[Quality receipts](results/change-0698/quality-summary.json) retain command,
source and log bindings. Fresh tool availability records show no cargo-fuzz or
nightly toolchain, so this batch makes no new fuzzing or Miri claim.

All six repository evidence gates also passed: crate boundaries, strict and
structural performance claims, report classification, CRUD coverage index and
the non-iWork gate. [Evidence receipts](results/change-0698/evidence/results.json)
bind their commands and logs.

The full pre-cleanup packet audit passes after correcting verifier bugs in
refusal-row unpacking, statistics definitions and the focused source-subset
comparison. The initial audit failure is preserved; no production source or
measurement was changed to obtain the passing result. The verifier now also
requires both exact 96-test receipts and the complete 29-driver hash census.


## Final disposition

[Independent review](results/change-0698/code-review.md) rejects the candidate.
A faster successful real-deck path does not offset the repeated reachable
refusal regression under this experiment's acceptance threshold. The +128-byte
handler stack cost reinforces the need for a better revision. No host-noise
explanation or smaller-error-path exemption is used to waive the result.

[rejection.json](results/change-0698/rejection.json) binds the measured candidate
source map, the retained final source map and the exact candidate codec witness.
The production codec matches baseline byte-for-byte; only the three regression
tests remain as Rust changes. Their baseline focused run tested this
exact codec/test combination with 96 passes, and a fresh post-restoration
[retained focused run](results/change-0698/focused-retained.json) also passes
all 96 tests. Full integration and oracle results above validate the measured
candidate; these focused receipts validate the retained combination. Rejection does not turn
candidate timings into a production performance claim.

The next experiment should investigate the duplicate-name refusal path and
handler code layout before attempting to retain borrowed views. Persistent
frame/context ownership remains unchanged. The measured source, raw data,
initial failures, corrected audit and rejection decision are retained for
reproduction.

The rejected-state audit and its independent review pass. Cleanup removed
exactly the batch target, frozen binaries, raw profiling directory and generated
control archive; the workspace lockfile is preserved. The post-cleanup audit
and final documentation gates are recorded in
[final-validation.json](results/change-0698/final-validation.json).
