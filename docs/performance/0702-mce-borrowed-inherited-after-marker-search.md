# 0702 — borrowed inherited namespaces after marker-search isolation

Disposition: retained with scoped performance evidence. `performance_claim: none`.
Baseline revision: `72d6f4500d`.

[0698](0698-mce-borrowed-inherited-namespaces.md) rejected temporary borrowed
namespace views despite improved real-edit medians because early-name refusal
regressed. [0699](0699-marker-free-refusal-attribution.md) proved that refusal
bypasses the changed MCE start handler and takes the marker-free window scan.
[0701](0701-mce-first-byte-marker-search.md) replaced that scan with a private
outlined first-byte helper. This materially changes the previously blocking
path, making a fresh comparison of the ownership candidate appropriate.
It does not establish that the old regression is fixed or predict a gain.

The retained 0701 open-plus-capture profile has MCE `start` at 9.37% self time,
the largest individual symbol share in that run. The earlier isolated marked
XML attribution identified temporary namespace Arc clone/drop pairs in the
`Inherited` view. This experiment removes only those temporary owners. Parent
Ctx and child-frame namespace ownership remain owned. The marker helper and
its dispatch predicate stay unchanged.

## Semantics and constraints

`Inherited` borrows the parent's namespace and emitted-boundary Arc references.
The parent borrow begins after any mutable AlternateContent bookkeeping and
ends before stack mutation. Opaque, selected, skipped and unwrapped branches
must preserve exact error ordering and output. `Inherited::after` still clones
the owner that a child Frame must retain. No borrowed parent reference escapes
into a frame. No parser, limit, namespace-hoisting or ownership contract changes.

All 33 recorded goal/ADR constraints are revalidated before edits. ADRs
0001/0002/0024 keep this private to the existing shared owner; ADR 0003 retains
edit and patch behavior; ADRs 0005/0031 retain resource and execution policy;
ADR 0006 retains preservation and refusal; ADR 0008 requires frozen differential
and gate evidence. ADRs 0010/0011 package ownership is unchanged. No cache,
worker, unsafe code, API, dependency or retained namespace representation is added.

## Method

The packet binds the complete 602-file Rust source census, lockfiles, build
inputs, fresh environment and exact codec-only source diff. The 104 existing
MCE tests include the three namespace/hoisting regressions retained from 0698
and the marker-search tests from 0700/0701. The identical test source runs
against both production states; no redundant tests are added to this retry.

Native timings cover capture, working clone, text editing, commit and apply
on prepared packages, excluding initial opening, setup, saving and semantic
assertions. Thirteen workflows use two A/A and four ABBA legs, 100 samples and
five warmups. Separate allocation, ten-case refusal, shared XML and marker
controls retain their own denominators. The 192-case/five-profile oracle
compares exact output, ownership, Report and errors. After seven integration
and six repository gates finish, seven native cases and the full refusal matrix
run a separate four-leg, 300-sample/ten-warmup follow-up.

The host is shared and warm. Every >5% review trigger remains explicit;
no cold-cache, quiescence, aarch64, remote I/O, full-save or parallel-scaling
claim applies. Retention requires practically useful representative gains
without repeating the previous blocking refusal cost. Assembly, allocations,
RSS and common marked-path controls must support the decision independently.

Both focused suites pass all 104 tests with warning denial. The test file is
byte-identical between baseline and candidate, as are the marker helper and
its dispatch call. The candidate changes only `codec.rs`; its source SHA-256 is
`a5b5b0aca3ec5a392bc7ae1ea6ca482653bd0a72cb9bdd4c30b8ea9faff87bee`.
Initial native A/A median drift ranges from −1.60% to +3.14%; this shared-host
variation remains part of the retained evidence.

## Initial native and refusal results

| Workflow / input | Baseline medians (ms) | Candidate medians (ms) | Paired change |
| --- | ---: | ---: | ---: |
| one-real | 11.1498 / 11.1727 | 10.7508 / 10.6699 | -3.58% / -4.50% |
| one-control | 5.2096 / 5.1847 | 5.1696 / 5.0265 | -0.77% / -3.05% |
| one-generated | 1.3772 / 1.3876 | 1.3642 / 1.3617 | -0.95% / -1.87% |
| one-notes-poi | 0.9046 / 0.9136 | 0.9215 / 0.8883 | +1.88% / -2.77% |
| one-notes-lo | 1.4396 / 1.4592 | 1.4140 / 1.4114 | -1.78% / -3.27% |
| noop-real | 5.3695 / 5.4457 | 4.9883 / 4.9914 | -7.10% / -8.34% |
| noop-control | 2.5478 / 2.5549 | 2.4476 / 2.4489 | -3.93% / -4.15% |
| noop-generated | 0.6743 / 0.6794 | 0.6644 / 0.6608 | -1.48% / -2.74% |
| noop-notes-poi | 0.4115 / 0.4061 | 0.3998 / 0.4022 | -2.85% / -0.97% |
| noop-notes-lo | 0.5183 / 0.5249 | 0.5093 / 0.5075 | -1.74% / -3.31% |
| two-real | 11.6030 / 11.4816 | 10.8301 / 10.8760 | -6.66% / -5.27% |
| two-control | 5.3757 / 5.3274 | 5.1982 / 5.2352 | -3.30% / -1.73% |
| two-generated | 1.5991 / 1.6014 | 1.6329 / 1.5669 | +2.12% / -2.15% |

| Real one-edit phase | Baseline medians (ms) | Candidate medians (ms) |
| --- | ---: | ---: |
| capture | 4.7942 / 4.8106 | 4.6395 / 4.5849 |
| clone | 0.0164 / 0.0169 | 0.0162 / 0.0163 |
| settext | 1.0391 / 1.0277 | 0.9967 / 0.9977 |
| commit | 5.2056 / 5.2177 | 4.9996 / 4.9746 |
| apply | 0.0911 / 0.0914 | 0.0891 / 0.0893 |

| One-edit allocation diagnostic | Baseline | Candidate |
| --- | ---: | ---: |
| alloc_calls | 143,848 | 143,848 |
| realloc_calls | 4,870 | 4,870 |
| requested_bytes | 12,430,874 | 12,430,874 |
| peak_above_start | 459,170 | 459,170 |
| net_live_change | 195,730 | 195,730 |

| Case | Baseline medians (µs) | Candidate medians (µs) | Paired change |
| --- | ---: | ---: | ---: |
| small-valid | 218.86 / 223.26 | 219.95 / 220.53 | +0.50% / -1.22% |
| generated-12x8-valid | 548.72 / 543.28 | 557.83 / 547.88 | +1.66% / +0.85% |
| early-name-error | 18.81 / 18.59 | 18.44 / 18.53 | -1.99% / -0.32% |
| late-root-error | 64.14 / 63.10 | 63.30 / 63.88 | -1.32% / +1.23% |
| late-root-error-mce | 143.37 / 142.21 | 138.95 / 137.83 | -3.09% / -3.08% |
| late-missing-relationship | 61.48 / 60.81 | 61.19 / 61.28 | -0.47% / +0.79% |
| late-missing-relationship-mce | 141.75 / 140.08 | 136.99 / 135.27 | -3.36% / -3.43% |
| notes-invalid-tail | 204.31 / 195.53 | 193.45 / 191.84 | -5.31% / -1.89% |
| mixed-conformance | 98.26 / 97.49 | 98.34 / 96.44 | +0.09% / -1.07% |
| slide-raw-overlimit-16m-to-64m | 156.87 / 161.30 | 167.31 / 168.56 | +6.66% / +4.50% |

These initial matrices contain 7,800 native workflows and 6,000 refusal/control
captures. All exact workflow, graph and refusal identities match. All 78
allocation comparisons have zero deltas in every recorded metric and repeat.
Requested bytes include replacement sizes for reallocations; no allocation
reduction is inferred from removing Arc reference-count operations.

The old early-name refusal regression does not appear in either initial pair.
Fifteen native primary-statistic review flags remain: the only total-workflow
flag is two-generated p99 +14.40% (+232 µs); its capture p99 is +36.86%
(+213 µs). Other flags concern clone/apply/commit/set-text phases, including
real one-edit clone p99 +41.99% (+7.58 µs).

Eight refusal metrics exceed 5%. The over-limit case flags first-pair median
+6.66% (+10.44 µs), mean +6.07%, p95 +10.03% and p99 +11.61%; its
second-pair p99 is +8.04%. Other p99 flags are generated-valid +5.43%
(+30.14 µs), late-root +9.95% (+6.90 µs), and notes-invalid-tail +44.50%
(+94.98 µs). The same later 16 MiB validation limit and typed over-limit
refusal remain in force. These regressions require the longer follow-up and
explicit review rather than being hidden by the real-work gains.

## Shared XML and marker controls

The oracle reports zero mismatches across 192 cases, five profiles and both
binaries (1,920 invocations: 1,046 successful results and 874 typed errors).
The deterministic corpus SHA-256 is
`a6f8abb6b2c504defa6bd879bcb6e31392a690976ba941d1e4d9a082f1bebf16`.
Exact output bytes/hash/length, Cow ownership, whole Report and error identity
agree. This is parity with the previous processor, not independent native
Office validation. Secondary XML controls add 37,800 isolated timings and
marker controls 9,000, each using 300 samples and ten warmups per process.

| Family / case | Profile | First median delta | Second median delta |
| --- | --- | ---: | ---: |
| synthetic: opaque | baseline | -8.26% | -7.11% |
| synthetic: opaque | opaque | -10.98% | -8.75% |
| synthetic: opaque | opaque-many | -8.09% | -9.83% |
| synthetic: ordinary | baseline | -4.52% | -5.29% |
| synthetic: ordinary | opaque | -6.60% | -3.61% |
| synthetic: ordinary | opaque-many | -5.34% | -6.60% |
| real part: docx | baseline | +0.18% | +1.28% |
| real part: docx | opaque | -1.99% | -1.67% |
| real part: docx | opaque-many | -1.82% | -1.21% |
| real part: pptx | baseline | -6.82% | -1.89% |
| real part: pptx | opaque | -3.80% | -3.74% |
| real part: pptx | opaque-many | -2.84% | -3.64% |
| real part: xlsx | baseline | -4.64% | -3.90% |
| real part: xlsx | opaque | -4.12% | -2.27% |
| real part: xlsx | opaque-many | -3.91% | -3.12% |
| declaration: declared | baseline | -3.15% | -2.29% |
| declaration: declared | opaque | -1.28% | -2.11% |
| declaration: declared | opaque-many | -3.18% | -2.38% |
| declaration: mixed | baseline | -4.56% | -5.48% |
| declaration: mixed | opaque | -3.94% | -2.44% |
| declaration: mixed | opaque-many | -3.30% | -3.95% |
| marker: large-near-prefix-marker-free | baseline | -15.58% | -1.77% |
| marker: long-late-comment-hit | baseline | +0.44% | +0.22% |
| marker: root-comment-hit | baseline | -6.38% | -4.26% |
| marker: tiny-marker-free | baseline | +0.00% | +0.00% |
| marker: tiny-valid-marked | baseline | +4.65% | +4.65% |

The only shared XML primary-statistic flags are PPTX many-name p95 +39.47%
and p99 +9.77% in the second pair. Marker controls flag tiny marked first-pair
mean +10.46% and p99 +6.12%, plus tiny marker-free second-pair mean +6.74%.
Tiny marked medians rise 430 → 450 ns (+20 ns) in both pairs; root-comment
medians improve. Nanosecond marker-free means are near timer resolution.
The first-byte search itself is unchanged in this batch, so its control
variation is not attributed to a new search algorithm.

## Code, stack and counters

The MCE start symbol shrinks 18,599 → 18,192 bytes (−407), but its recorded
stack reservation grows 0x598 → 0x618 (+128 bytes). Native text shrinks
2,598,022 → 2,597,494 bytes (−528), data remains 60,872 bytes, and BSS
grows 1,488 → 2,032 bytes (+544). The processor and marker helper remain
7,183 and 170 bytes at the same addresses in both binaries; relocation and
call-target differences mean their disassembly is not byte-identical.
This does not establish peak-stack or whole-program memory savings.

The single `(210 − 10) / 200` open-plus-capture counter slope changes cycles -3.96%, instructions -0.13%, branches -0.24%, branch-misses +0.98%, cache-misses +0.69%, page-faults -8.20%, task-clock -3.99%. Whole-child RSS is 5,560 → 5,636 KiB. These diagnostic counters have a different denominator from native phase timings; they do not prove a repeated latency gain or RSS bound.

The longer follow-up includes two-generated in addition to the six established
native cases because its initial total p99 flags +14.40%. This seventh case
was added before follow-up measurements began; the initial data and completed
drivers remain unchanged. The same full ten-case refusal matrix checks the
new over-limit trigger and the earlier early-name blocker.

Fresh open-plus-capture profiles place `start` at 9.29% baseline and 8.83%
candidate self time. Baseline `Inherited` destruction accounts for 3.02%
self time; that symbol is absent from the candidate profile listing. Static
start-handler assembly has nine locked-increment and fourteen locked-decrement
sites before, versus six and eight after. These are instruction-site counts,
not operations per XML element or an attribution of all elapsed-time savings.

## Candidate validation

All seven integration gates pass on the frozen candidate: formatting, locked
all-feature checks, warning-denied Clippy, default tests (1,222 passed, two
ignored), all-feature Office tests (4,146 passed, 33 ignored), PPTX facade
tests (45 passed), and warning-denied rustdoc. No tests failed. The suites
overlap, so these are not additive unique-coverage counts. Both separate
focused MCE suites pass the same 104 tests.

All six repository evidence gates pass: crate boundaries, strict and structural
performance claims, report classification, CRUD coverage, and non-iWork
verification. Tool availability is recorded; no cargo-fuzz/nightly campaign or
native Office GUI run is claimed. The expanded follow-up starts only after
all thirteen integration/evidence gates have finished.

## Expanded independent follow-up

After all seven integration and six evidence gates finished, four ABBA legs
ran 300 samples after ten warmups: 8,400 native workflows over seven cases and
12,000 captures over ten refusal/control cases. The independent audit passes
all 28 native rows, 40 refusal rows and 156 comparisons. The initial data is
preserved unchanged.

| Follow-up case | First median delta | Second median delta |
| --- | ---: | ---: |
| native: noop-real | -4.43% | -4.29% |
| native: one-control | +1.32% | -5.21% |
| native: one-generated | -0.45% | -0.13% |
| native: one-notes-poi | -2.02% | -3.12% |
| native: one-real | -3.72% | -4.77% |
| native: two-generated | -1.57% | -1.23% |
| native: two-real | -3.74% | -1.15% |
| refusal: early-name-error | +0.75% | -0.21% |
| refusal: generated-12x8-valid | +1.60% | -0.38% |
| refusal: late-missing-relationship | +2.53% | +0.34% |
| refusal: late-missing-relationship-mce | -2.68% | -6.36% |
| refusal: late-root-error | +2.67% | -0.12% |
| refusal: late-root-error-mce | -2.58% | -6.68% |
| refusal: mixed-conformance | +1.63% | -1.29% |
| refusal: notes-invalid-tail | -1.82% | +1.43% |
| refusal: slide-raw-overlimit-16m-to-64m | +0.42% | -0.82% |
| refusal: small-valid | +1.98% | -1.81% |

Native baseline median drift ranges from −1.91% to +3.40%, the largest on
one-control; that control's mixed candidate pair is not treated as a consistent
gain. Refusal drift ranges from −2.04% to +3.25%. Real one-edit medians improve
3.72/4.77% (413/536 µs), real no-op 4.43/4.29% (233/226 µs), and real
two-edit 3.74/1.15% (426/133 µs). These reinforce the initial real-work result
without turning every small control delta into a claim.

The initial generated two-edit total/capture p99 flags do not recur; its
longer total medians improve 1.57/1.24%. The over-limit refusal changes
+0.42/−0.82% (+0.68/−1.35 µs), with no primary metric above +5%.

Twenty-four follow-up triggers include extrema. Seven concern primary
statistics. Four are clone p99: one-generated +24.98% (+3.89 µs),
two-generated +50.87/+14.78% (+6.71/+1.89 µs), and two-real +14.47%
(+2.79 µs). The remaining three concern early-name refusal in the first pair:
mean +5.14% (+0.96 µs), p95 +55.05% (+10.35 µs), p99 +15.40%
(+3.92 µs). Its median is +0.75% (+0.14 µs), and the second-pair median
is −0.21%, with no primary flag. Thus the previous sustained median cost
is not reproduced, but an intermittent refusal-tail cost remains. It is
not dismissed as noise or hidden by the real-work gains.

The raw early-name follow-up shows 27/300 samples at or above 25 µs in b0:
indices 0–24, 76 and 286. The other legs have five (a0), six (b1) and three
(a3) such samples. Thus the first candidate leg's slow samples cluster near
its timed beginning despite ten warmups. Every sample remains included; this
observation neither proves a cause nor authorizes trimming the distribution.

## Disposition

Retain borrowed inherited namespace views on the 0701 marker-search baseline.
Real one-edit and no-op medians improve about 4% in both longer pairs, while
real two-edit medians improve 1.15–3.74%. Shared marked XML controls also
mostly improve, with exact semantic parity and unchanged allocation metrics.
The previous sustained early-name median regression is absent, and the new
over-limit and generated total/capture triggers do not repeat in the longer
run. This combination changes the retention decision from 0698; it does not
retroactively turn that rejected experiment into a valid retained result.

Accept the explicit costs: 128 additional bytes of start-handler stack
reservation, 544 bytes of native BSS, tiny marked median +20 ns, shared XML
and clone tails, and first-leg early-name mean/p95/p99 regressions. The source
removes temporary reference-count traffic rather than allocations; no memory
reduction or universal refusal-path improvement is claimed. The early-name
tail remains a follow-up target, with no proven attribution to the source
change or to host noise. Its unchanged lexical-search route must be preserved
in any subsequent investigation.

The production change is confined to `codec.rs`; the 104 existing focused
tests remain unchanged. Claims apply only to this corpus, build, shared host
and timed prepared-package workflows. Full-save, cold-cache, native Office,
remote-I/O and parallel-scaling gains are not established here.

The complete pre-cleanup audit passes source/probe/binary bindings, patch
replay, raw statistics, independent marker/follow-up checks, both focused
receipts and all candidate gates. Cleanup removes exactly the batch target,
binary and profile directories plus its generated marker-control archive.
The workspace Cargo.lock is preserved; unrelated scratch is untouched. Raw
results, original lockfiles, source witnesses and reproduction scripts remain.
