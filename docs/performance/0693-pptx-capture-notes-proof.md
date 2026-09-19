# 0693 — reuse capture-local slide proofs during notes validation

Status: retain the measured capture-only optimization.
`performance_claim: none`. Baseline: `5e89851b9`.

The real 13-slide one-edit phases fall from **16.93–16.95 ms to 11.96–12.00 ms**
median, a paired **29.11–29.42% reduction**. Real no-op phases improve
34.22–34.68%, and two-slide edit phases improve 28.83–28.93%. Capture reuses
its already processed slide XML to establish the later notes root proof,
removing another MCE pass per slide without retaining processed XML.

The change has a deliberate refusal tradeoff: a late bad root or missing
relationship now follows full validation of earlier slides. The measured
three-slide refusal fixtures add **43.95–48.72 µs**, or **125–161%** for the
marker-free fast exits and **42–47%** for MCE-bearing variants. These costs are
retained individually below. The coordinator and independent reviewer accept
them under the existing finite scanner/catalog limits in exchange for the
measured gains on common real, control and generated workflows. There is no
claim that every refusal becomes faster.

The [packet](results/change-0693/README.md) retains frozen source, dependency,
probe, corpus and binary bindings, raw samples, separate allocator diagnostics,
profiles, trace/restoration receipts, and code/evidence reviews. The broader
GOAL remains active; iWork is excluded. No coverage entry or registered claim
is promoted by this phase-only change.

## Native scope and results

Thirteen workflows have two fresh pre-edit A/A legs and four A/B/B/A legs:
**7,800 native samples**, with 100 samples and five warmups per process. The
paired changes below are b0/a2 and b1/a3; every leg and bootstrap median
interval remains separate in the JSON. CPU 12 is pinned on a shared Linux
host, with warm caches and no demonstrated host-wide quiescence. Initial A/A
total-median drift ranges from −1.67% to +2.50%; initial no-op-real p99 drift
is +6.80%. The final no-op-generated baseline a2 is elevated (0.9082 ms versus
0.8301 ms in a3), so its −17.08% pair is not a clean isolated estimate. Its
other pair remains −9.57%. Small notes-fixture differences are near host drift.

Timers cover capture, working clone, text edit, commit, apply and their
enclosing total. File reads, corpus construction, initial package open, target
selection, save and reopen are outside these timers. This is not a complete
open/edit/save benchmark. Inputs are `slide-section-test`, its marker
counterfactual, generated 12-slide × 8-text-box content, POI `prProps` notes,
and LibreOffice `tdf131082` notes. Two-edit cases use distinct slides. The
counterfactual changes marker namespaces and ZIP metadata; it is a mechanism
control, not a semantic-equivalence or untouched-payload oracle.

| Workflow / input | Baseline medians (ms) | Candidate medians (ms) | Paired change |
| --- | ---: | ---: | ---: |
| one-real | 16.9501 / 16.9288 | 11.9637 / 12.0003 | -29.42% / -29.11% |
| one-control | 6.7287 / 6.7097 | 6.0850 / 5.9865 | -9.57% / -10.78% |
| one-generated | 1.7174 / 1.7078 | 1.5660 / 1.5542 | -8.82% / -8.99% |
| one-notes-poi | 1.0583 / 1.0661 | 1.0464 / 1.0443 | -1.13% / -2.05% |
| one-notes-lo | 1.6953 / 1.6825 | 1.6474 / 1.6429 | -2.83% / -2.36% |
| noop-real | 8.4877 / 8.5614 | 5.5828 / 5.5926 | -34.22% / -34.68% |
| noop-control | 3.4486 / 3.4344 | 2.8722 / 2.8839 | -16.71% / -16.03% |
| noop-generated | 0.9082 / 0.8301 | 0.7531 / 0.7507 | -17.08% / -9.57% |
| noop-notes-poi | 0.4828 / 0.4881 | 0.4787 / 0.4735 | -0.85% / -2.98% |
| noop-notes-lo | 0.6280 / 0.6258 | 0.6426 / 0.6118 | +2.32% / -2.24% |
| two-real | 17.3739 / 17.2411 | 12.3643 / 12.2531 | -28.83% / -28.93% |
| two-control | 6.8684 / 6.8664 | 6.2661 / 6.1006 | -8.77% / -11.15% |
| two-generated | 1.9027 / 1.9206 | 1.7718 / 1.7713 | -6.88% / -7.77% |

| Real one-edit phase | Baseline medians (ms) | Candidate medians (ms) |
| --- | ---: | ---: |
| capture | 7.6278 / 7.6326 | 5.1832 / 5.1948 |
| clone | 0.0166 / 0.0165 | 0.0163 / 0.0165 |
| settext | 1.0806 / 1.0783 | 1.0946 / 1.0950 |
| commit | 8.1213 / 8.1021 | 5.5750 / 5.5882 |
| apply | 0.0953 / 0.0948 | 0.0927 / 0.0927 |

| One-edit allocation diagnostic | Baseline | Candidate |
| --- | ---: | ---: |
| alloc_calls | 270,242 | 190,338 |
| realloc_calls | 6,950 | 4,870 |
| requested_bytes | 19,915,826 | 13,815,905 |
| peak_above_start | 463,159 | 459,170 |
| net_live_change | 195,730 | 195,730 |

Requested bytes already include the full new size of successful reallocations;
do not add the realloc-request counter again. All three repeats of every
allocator metric agree exactly. Across all 78 case/phase comparison groups,
net live bytes are unchanged and no measured peak-above-start increases. The
real one-edit total peak falls only 3,989 bytes (463,159 → 459,170), while
79,904 allocation calls and 6,099,921 requested bytes are removed.

No total median, mean, p95 or p99 exceeds the 5% regression trigger. Their
largest increases are all on one no-op LibreOffice pair: +2.32%, +2.20%,
+2.01% and +2.54%, respectively. There are **28 phase review triggers**:
one median, one mean, seven p95 and nineteen p99. The no-op LibreOffice apply
phase in b0/a2 rises 14.46% at median (2.32 µs) and 14.71% at mean (2.39 µs).
The largest p99 trigger is two-generated clone, +55.16% (7.05 µs); one-control
clone p99 adds 11.46 µs (+51.83%). Every trigger is retained in
`native-review-triggers.json`; these observations provide no universal tail
latency guarantee.

## Refusal and control guardrails

The separate native probe times only `Package::opened_presentation()` on an
immutable prepared package. Fixture authoring, mutation, graph hashing,
result formatting, assertions and snapshot destruction are outside the timer.
Ten cases have two A/A and four A/B/B/A legs: **6,000 samples**, each checking
its expected typed result. Graph hashes cover payloads, names, content types,
part relationships and root relationships. The duplicate-name case freezes its
full typed Debug value per process, and all legs additionally require exact
metadata/error equality. Valid controls assert slide count and first name.

| Case | Baseline medians (µs) | Candidate medians (µs) | Paired change |
| --- | ---: | ---: | ---: |
| small-valid | 257.91 / 254.10 | 246.54 / 245.83 | -4.41% / -3.25% |
| generated-12x8-valid | 702.01 / 701.76 | 626.52 / 624.17 | -10.75% / -11.06% |
| early-name-error | 35.20 / 35.25 | 35.04 / 35.16 | -0.45% / -0.24% |
| late-root-error | 35.03 / 35.07 | 79.54 / 79.02 | +127.06% / +125.31% |
| late-root-error-mce | 110.20 / 111.73 | 158.92 / 159.02 | +44.21% / +42.33% |
| late-missing-relationship | 27.61 / 27.81 | 72.17 / 72.61 | +161.43% / +161.09% |
| late-missing-relationship-mce | 103.33 / 104.58 | 151.61 / 151.96 | +46.73% / +45.30% |
| notes-invalid-tail | 259.33 / 253.36 | 233.79 / 228.87 | -9.85% / -9.67% |
| mixed-conformance | 112.33 / 112.66 | 113.18 / 114.10 | +0.76% / +1.28% |
| slide-raw-overlimit-16m-to-64m | 22563.47 / 22553.49 | 22554.86 / 22541.21 | -0.04% / -0.05% |

The late-refusal regressions repeat in both pairs; they are not dismissed as
noise. Earlier name errors still stop additional proof work, and raw XML above
the notes slide ceiling still produces the legacy generic root refusal. Mixed
conformance rises only 0.76–1.28%; notes-tail refusal improves 9.67–9.85%.
Restricting reuse to owned MCE output would discard measured common
marker-free gains, so the retained implementation uses the broad proof path.
The proof prefix stops after its first invalid classification. There is no
extra cumulative proof-work allowance: the existing catalog, per-part XML,
notes depth/node/attribute and byte limits remain the finite bounds.

## Mechanism, memory and semantic boundaries

`processed_xml_with_source` preserves two `Part::blob()` observations: the
first supplies the existing 64 MiB part preflight, and the second is the exact
slice passed to MCE. A borrowed reference to that second slice is the proof
witness, keeping it alive rather than storing an unowned address. After a
successful slide root and name projection, the full notes scanner checks the
already processed XML, including the distinct 16 MiB slide limit and processed
byte limit, in Transitional-then-Strict order. Only the raw witness and optional
conformance survive; the processed Cow and scanner inventory are dropped.

A private notes loader retains all presentation inventory, relationship,
content-type, conformance, master/theme, notes resource, orphan and snapshot
materialization checks. It disables hints when inventory lengths differ.
Otherwise the resolver compares the current slice pointer and length at the
original notes-validation position. Missing or mismatched hints use the old
parser; a matching invalid classification reproduces the old generic error.
Raw classification depends on bytes and fixed limits, not a relationship ID,
so equal-length reordered hints are safe after that source check. All later
slide roots and deferred name/error ordering remain intact.

The temporary trace confirms **31 → 18 MCE calls** per real/control capture
and **31 → 19** for generated input. Real input processed by those calls falls
from 550,141 to 280,963 bytes; owned output falls from 634,121 to 324,063 bytes.
The commit capture has 280,892 input and 323,993 output bytes. Five main-part
calls remain; generated input has twelve slide calls and two additional
unmapped resource-path calls. All generated/control calls return borrowed
output. Pointer mappings are local to one process; no cross-process address
comparison is used. All instrumented source bytes were restored exactly.

The trace measures **24 bytes per proof** and **48 bytes per capture entry** on
this 64-bit host. The extra reservation is 24N requested logical bytes: 312
bytes for thirteen slides. During projection, the two vectors total 72N bytes
before names and parser/MCE scratch; capture entries are consumed before notes
loading and proofs are dropped immediately afterward. The ordinary immutable
default-policy catalog permits at most 4,096 entries. The actual reservation
follows the second catalog read, whose absolute existing `MAX_SLIDES` bound is
100,000; caller-custom limits and changing foreign parts must not be described
as universally capped at 4,096. Allocator capacity and full live-byte/RSS
measurements remain separate from this layout arithmetic.

The optional proof-vector reservation uses `try_reserve_exact`; failure disables
reuse, and a successful reserve covers the complete prefix without growth.
Only that reservation has recoverable fallback. The inherited XML scanner has
infallible allocations, which can now abort before a later document refusal;
that abort ordering is not guaranteed unchanged. Reusing successful MCE output
also removes a later independent MCE allocation-failure opportunity. These are
explicit resource differences, not changes to typed document limits or
validation. No proof, processed XML, cache, lock or public API state survives
the capture operation.

The native open-plus-capture counter slope, `(210 calls − 10 calls) / 200`,
falls from **159.17M to 108.71M instructions** and **39.74M to 28.78M cycles**.
Branches fall from 32.29M to 21.95M and branch misses from 133,595 to 110,528.
Page faults fall from 211.32 to 175.08 per slope operation. Whole-child peak
RSS is 5,540 → 5,632 KiB (+1.66%), with one diagnostic process per build;
there is no RSS improvement or bound claim. Native text grows by 7,044 bytes
(2,588,542 → 2,595,586). Sampling, allocator counters and traces explain work;
they are not substitutes for the native phase samples. Cold caches, remote
sources, concurrent scaling and full save workflows are outside this packet.

## Verification and retained setup history

Ten new tests cover processed-root T/S retry and limit masking; exact witness,
equal-byte/different-allocation, changed-byte and different-length fallback;
the two source observations; strict-with-notes and MCE graphs; root/name/notes
error precedence; foreign inventory lookalikes; separate 16/64 MiB limits;
independent package sources; and full notes-load swapped equal-length sentinel
proofs. Existing no-op, source sharing, untouched-payload/relationship,
publication, patch and reopen tests remain in the owner suites and probes.

Final validation passes: workspace formatting; PPTX all-feature/all-target
check; warning-denied all-feature library Clippy; **918 default** and
**932 all-feature tests**, each with two existing ignored examples;
**45 facade tests**; and warning-denied rustdoc. The narrow facade check uses
normal Rust warnings because its unrelated unused helper warnings predate this
batch, as recorded in 0692. Owner checks and rustdoc retain warning denial.
No cargo-fuzz command or nightly toolchain is installed; no fuzz/Miri run is
claimed. The accepted-ADR mapping is in the packet design; all 33 frozen
GOAL/ADR constraint hashes remain unchanged.

Setup history remains visible. The initial probe lock-copy checksum mistake,
malformed-fixture serialization attempt and corrected notes-tail oracle are
archived separately. Eight-case refusal receipts were expanded with two MCE
variants; that supplemental baseline was built from exact HEAD source bytes,
then all candidate files were restored byte-for-byte. The initial direct
resolver test exposed identity filtering in its caller rather than the helper;
the guard moved into the resolver, the full notes-load test was added, and all
candidate native, allocator, refusal and profile measurements were repeated.
The superseded source and results remain under `initial-identity-helper`.
A trace-only qualification lint failure is retained under `trace-initial`;
its source restoration passed and it contributed no performance samples.

All six repository boundary/claim/coverage/non-iWork gates and the source/data
evidence audit pass. The post-cleanup seal reruns the audit and final report,
coverage, non-iWork and structural-claim checks. Cleanup removes only this
batch's named target, staged binaries, raw profile directory and generated
marker-control archive, preserving the workspace Cargo.lock and shared targets.
The final evidence review and receipts are linked from the packet.
