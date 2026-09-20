# 0701 — first-byte search for the MCE namespace marker

Disposition: retained with scoped performance evidence. `performance_claim: none`.
Baseline revision: `7e4c245673`.

[0700](0700-mce-marker-substring-search.md) rejected direct memmem search:
large marker-free inputs improved, but real edit/no-op medians slowed and tiny
marked inputs paid about 210 ns of setup. This experiment uses the existing
safe first-byte search and exact prefix comparison in a private outlined helper.
It tests whether skipping impossible start positions can preserve the scan gain
without the general substring searcher's per-call setup. No gain is assumed.

The helper searches only positions where the complete 59-byte namespace can
fit. After a first-byte hit it checks the full exact URI, then advances one
byte on a miss. It treats input as arbitrary bytes. With a fixed URI, each
iteration advances and the worst-case work remains linear in input length,
with up to one fixed-size comparison per candidate position. Repeated prefixes
and short marked inputs are explicit performance controls.

Input limits precede search. Absent marker bytes retain the output-limit check,
borrowed XML and default Report; any exact occurrence enters the existing XML
processor, including text, comments and malformed input. Frame ownership,
streaming processing, public APIs and dependencies are unchanged.

## Constraints and measurement scope

| Contract | Batch scope |
| --- | --- |
| ADR 0001/0002/0024: layers and ownership | Private shared-MCE helper; existing dependency only |
| ADR 0003: edit/patch contracts | Exact native graph, patch and refusal checks retained |
| ADR 0005/0031: resources and execution | No cache, retained allocation, worker or execution provider added |
| ADR 0006: preservation and refusal | Exact bytes, ownership, Report, errors and limit order retained |
| ADR 0008: evidence | Frozen baseline/candidate builds, focused tests, shared oracle and gates |
| ADR 0010/0011: package ownership | No archive or package change |

All 33 recorded goal/ADR constraints were revalidated before edits. The packet
binds 602 Rust source files, build inputs, lockfiles, fixture hashes and fresh
host/toolchain records. The sole batch Cargo/performance lane runs serially.
The host remains shared and warm; no cold-cache or quiescence claim applies.

The native probe times capture, working clone, text editing, commit and apply
on prepared packages. Initial opening, fixture setup, save and semantic checks
are outside that timer. Thirteen workflows use two A/A legs and four ABBA legs,
100 samples and five warmups. Allocator instrumentation runs separately; the
refusal matrix includes ten cases. A longer independent follow-up uses six
native cases and the full refusal matrix with four legs, 300 samples and ten
warmups. Shared XML and marker controls use their own isolated denominators.

Preserve every review trigger above 5%; do not conceal individual regressions
in an aggregate. Native Office GUI, aarch64, cold-cache, remote I/O and parallel
scaling are outside this local predicate experiment. No new fuzz campaign is
claimed unless fresh availability and execution evidence proves one.

## Source and baseline validation

The baseline and candidate each pass the same 104 focused MCE tests with
warnings denied. The three new tests exercise 16,442 dispatch checks against
the previous window predicate: arbitrary-byte surroundings at all starts
through 128, every byte substitution at each of 59 URI positions, and repeated
candidate/prefix sequences. Absent markers must preserve exact borrowed bytes
and default Report; present markers must enter the owned-output/error path.
Existing tests retain the separate exact error and limit-order assertions.
The source patch binds the original 101-test baseline witness and the shared
104-test baseline/candidate focused source; those are distinct receipts.

Initial A/A total medians range from −2.66% to +2.05% across the 13 native
workflows. This observed shared-host variation is retained in the packet;
performance judgments require the balanced candidate pairs and longer
independent follow-up rather than treating a single small delta as a gain.

## Initial native and refusal results

| Workflow / input | Baseline medians (ms) | Candidate medians (ms) | Paired change |
| --- | ---: | ---: | ---: |
| one-real | 11.0218 / 11.0381 | 11.1427 / 11.1643 | +1.10% / +1.14% |
| one-control | 5.9679 / 5.9536 | 5.1761 / 5.2022 | -13.27% / -12.62% |
| one-generated | 1.5710 / 1.5637 | 1.3858 / 1.3835 | -11.78% / -11.52% |
| one-notes-poi | 1.0455 / 1.0516 | 0.9058 / 0.9218 | -13.37% / -12.34% |
| one-notes-lo | 1.6530 / 1.6529 | 1.4322 / 1.4342 | -13.36% / -13.23% |
| noop-real | 5.1818 / 5.1865 | 5.2668 / 5.3310 | +1.64% / +2.79% |
| noop-control | 2.8698 / 2.8900 | 2.5189 / 2.5063 | -12.23% / -13.28% |
| noop-generated | 0.7576 / 0.7582 | 0.6796 / 0.6766 | -10.30% / -10.76% |
| noop-notes-poi | 0.4780 / 0.4755 | 0.4092 / 0.4105 | -14.40% / -13.68% |
| noop-notes-lo | 0.6135 / 0.6089 | 0.5218 / 0.5179 | -14.96% / -14.95% |
| two-real | 11.2729 / 11.2438 | 11.4673 / 11.4256 | +1.72% / +1.62% |
| two-control | 6.1431 / 6.1471 | 5.3565 / 5.3227 | -12.80% / -13.41% |
| two-generated | 1.8325 / 1.7783 | 1.5876 / 1.5817 | -13.36% / -11.06% |

| Real one-edit phase | Baseline medians (ms) | Candidate medians (ms) |
| --- | ---: | ---: |
| capture | 4.7366 / 4.7116 | 4.8144 / 4.8305 |
| clone | 0.0170 / 0.0170 | 0.0163 / 0.0163 |
| settext | 1.0269 / 1.0607 | 1.0152 / 1.0109 |
| commit | 5.1437 / 5.1509 | 5.2009 / 5.2141 |
| apply | 0.0919 / 0.0916 | 0.0905 / 0.0899 |

| One-edit allocation diagnostic | Baseline | Candidate |
| --- | ---: | ---: |
| alloc_calls | 143,848 | 143,848 |
| realloc_calls | 4,870 | 4,870 |
| requested_bytes | 12,430,874 | 12,430,874 |
| peak_above_start | 459,170 | 459,170 |
| net_live_change | 195,730 | 195,730 |

| Case | Baseline medians (µs) | Candidate medians (µs) | Paired change |
| --- | ---: | ---: | ---: |
| small-valid | 252.64 / 254.94 | 221.61 / 218.96 | -12.28% / -14.11% |
| generated-12x8-valid | 641.10 / 638.58 | 559.82 / 539.81 | -12.68% / -15.47% |
| early-name-error | 35.85 / 35.61 | 18.80 / 18.75 | -47.55% / -47.33% |
| late-root-error | 80.69 / 81.16 | 65.42 / 62.41 | -18.92% / -23.10% |
| late-root-error-mce | 150.13 / 151.42 | 145.06 / 142.79 | -3.38% / -5.70% |
| late-missing-relationship | 73.32 / 73.77 | 62.83 / 59.85 | -14.30% / -18.86% |
| late-missing-relationship-mce | 142.92 / 144.45 | 142.75 / 140.11 | -0.12% / -3.00% |
| notes-invalid-tail | 232.61 / 246.44 | 196.72 / 197.82 | -15.43% / -19.73% |
| mixed-conformance | 116.23 / 117.22 | 100.47 / 95.73 | -13.56% / -18.33% |
| slide-raw-overlimit-16m-to-64m | 22534.72 / 22543.82 | 165.28 / 162.68 | -99.27% / -99.28% |

The initial native matrix contains 7,800 timed workflows and the refusal matrix
6,000 captures. All exact native and refusal identities match. The 16–64 MiB
padded case retains the same later 16 MiB validation refusal; faster search
does not relax that boundary. All 78 allocation comparisons retain identical
metrics. Requested bytes include replacement sizes for reallocations and must
not be added to realloc bytes again. No allocation saving is claimed.

Fifteen native primary-statistic review triggers remain in the packet, including
real one-edit commit p99 +5.52/+6.72% (+289/+351 µs), and real no-op total p99
+6.12% (+321 µs) in the second pair. The other flags concern individual small
phases and are not hidden by whole-workflow improvements.

## Code and profile diagnostics

Native text grows 2,594,902 → 2,597,998 bytes (+3,096); data changes
60,848 → 60,840 (−8) and BSS 552 → 1,552 (+1,000). The processor symbol
grows 6,305 → 7,183 bytes (+878), while its local reservation shrinks
0x248 → 0x238 (−16 bytes). The separate helper is 170 bytes and pushes
56 bytes on its full-search path, in addition to its return address and
nested search calls. The smaller caller reservation is not a peak-stack
saving. Bounded disassembly retains the outlined call, valid-start search,
exact comparison and one-byte advancement, without FinderBuilder setup.

The single `(210 − 10) / 200` open-plus-capture counter slope changes cycles -0.15%, instructions -0.00%, branches -0.40%, branch-misses +0.45%, cache-misses -2.50%, page-faults -2.87%, task-clock -0.15%. Whole-child RSS is 5,652 → 5,764 KiB. These are diagnostic runs with a different denominator from native phase timers; they do not establish a repeated latency gain or memory bound.

## Shared XML and marker controls

The shared oracle has zero mismatches across 192 cases, five profiles and
both binaries (1,920 invocations: 1,046 exact successful results and 874 exact
typed-error results). Corpus SHA-256 is
`a6f8abb6b2c504defa6bd879bcb6e31392a690976ba941d1e4d9a082f1bebf16`.
This establishes parity with the existing processor; it is not an independent
Office validator. The secondary controls add 37,800 isolated timings and the
five marker cases add 9,000, each with 300 samples and ten warmups per leg.

| Family / case | Profile | First median delta | Second median delta |
| --- | --- | ---: | ---: |
| synthetic: opaque | baseline | -0.93% | -1.02% |
| synthetic: opaque | opaque | +0.04% | -1.33% |
| synthetic: opaque | opaque-many | +1.18% | -0.36% |
| synthetic: ordinary | baseline | +0.16% | +1.22% |
| synthetic: ordinary | opaque | -0.91% | +0.87% |
| synthetic: ordinary | opaque-many | -0.08% | -0.86% |
| real part: docx | baseline | -3.17% | +0.30% |
| real part: docx | opaque | +0.86% | -3.57% |
| real part: docx | opaque-many | +0.14% | +0.09% |
| real part: pptx | baseline | -1.62% | -1.38% |
| real part: pptx | opaque | -1.52% | -0.80% |
| real part: pptx | opaque-many | -1.63% | +0.53% |
| real part: xlsx | baseline | -0.63% | -0.68% |
| real part: xlsx | opaque | -0.72% | -1.40% |
| real part: xlsx | opaque-many | -3.10% | -0.30% |
| declaration: declared | baseline | -2.61% | -1.33% |
| declaration: declared | opaque | -1.79% | -1.32% |
| declaration: declared | opaque-many | -0.82% | -2.18% |
| declaration: mixed | baseline | -0.78% | +0.72% |
| declaration: mixed | opaque | +2.12% | +3.62% |
| declaration: mixed | opaque-many | +2.54% | +2.26% |
| marker: large-near-prefix-marker-free | baseline | -87.72% | -85.67% |
| marker: long-late-comment-hit | baseline | -97.24% | -97.46% |
| marker: root-comment-hit | baseline | -4.08% | -4.08% |
| marker: tiny-marker-free | baseline | +0.00% | +0.00% |
| marker: tiny-valid-marked | baseline | -8.51% | -7.45% |

Tiny marked medians improve from 470 ns to 430/435 ns (−40/−35 ns),
and root-comment medians improve 490 → 470 ns. This candidate avoids the
roughly 210 ns marked setup cost measured for 0700. Tiny marker-free medians
remain 30 ns; means flag +5.24/+6.27%, at nanosecond scale near timer resolution.
Repeated near-prefix scans and long late-comment hits improve substantially;
those gains are mechanism controls, not whole-document timings.

No synthetic or declaration primary statistic exceeds +5%. The real-part
controls flag DOCX opaque p99 +8.20% in the first pair and DOCX many-name p99
+25.26% in the second. Mixed opaque declaration medians cost +2.1–3.6%;
declared-only medians improve. All flags remain in the raw comparison records.

## Candidate integration validation

All seven integration gates pass on the frozen candidate: formatting, locked
all-feature checks, warning-denied Clippy, default tests (1,222 passed, two
ignored), all-feature Office tests (4,146 passed, 33 ignored), PPTX facade
tests (45 passed), and warning-denied rustdoc. No tests failed. These suites
overlap; the counts must not be added as unique test coverage. The separate
baseline and candidate focused MCE runs each pass 104 tests.

All six repository evidence gates also pass: crate boundaries, strict and
structural performance claims, report classification, CRUD coverage, and
non-iWork verification. Fresh tool availability records show no cargo-fuzz or
nightly campaign was run. The independent longer follow-up starts only after
these gates and all integration checks have finished.

## Longer independent follow-up

After all seven integration and six evidence gates finished, four ABBA legs
ran 300 samples after ten warmups: 7,200 native workflows over six cases and
12,000 captures over the complete ten-case refusal matrix. The independent
audit passes all 24 native and 40 refusal rows and 138 comparisons. Initial
measurements remain intact.

| Follow-up case | First median delta | Second median delta |
| --- | ---: | ---: |
| native: noop-real | -7.14% | -0.27% |
| native: one-control | -14.52% | -14.21% |
| native: one-generated | -11.74% | -11.28% |
| native: one-notes-poi | -14.51% | -14.77% |
| native: one-real | +0.07% | +0.68% |
| native: two-real | -0.61% | +0.34% |
| refusal: early-name-error | -48.17% | -48.28% |
| refusal: generated-12x8-valid | -12.69% | -14.35% |
| refusal: late-missing-relationship | -16.71% | -18.55% |
| refusal: late-missing-relationship-mce | -1.97% | -3.31% |
| refusal: late-root-error | -21.87% | -22.20% |
| refusal: late-root-error-mce | -4.64% | -5.78% |
| refusal: mixed-conformance | -15.65% | -18.14% |
| refusal: notes-invalid-tail | -17.00% | -22.22% |
| refusal: slide-raw-overlimit-16m-to-64m | -99.28% | -99.28% |
| refusal: small-valid | -13.69% | -13.27% |

The no-op real baseline drifts −6.87% between its two baseline legs. Its
first −7.14% comparison therefore does not establish a candidate speedup;
the second is −0.27%. Other native baseline medians drift −0.41% to +0.11%.
Real one-edit changes +0.07/+0.68% (+8/+76 µs), and two-edit changes
−0.61/+0.34% (−69/+38 µs). The initial real-work median and commit-tail
regressions do not recur at the same magnitude. No real-deck gain is claimed.

All ten refusal medians improve in both longer pairs. Refusal baseline drift
ranges from −0.65% to +4.60%, the latter on notes-invalid-tail; the large
marker-free gains remain well beyond that variation. Seventeen follow-up
triggers include extrema. Only three concern primary statistics, all clone
p99: generated +50.39% (+6.45 µs), notes-bearing +43.31% (+2.75 µs),
and one-real +7.50% (+1.68 µs). No whole-native-workflow or refusal primary
statistic exceeds +5% in this follow-up. These tails remain part of the
retention review even though whole-workflow medians improve or remain close.

## Disposition

Retain the private first-byte search with a scoped result. It accelerates
marker-free and notes-bearing workflows by about 11–15% in the longer
controls, removes most of the large rejected-buffer scan, and avoids 0700's
tiny marked setup penalty. Exact semantics and allocation behavior match.
The longer real-deck edit comparisons remain within ±0.7%; initial real
costs, shared DOCX tails, tiny marker-free means and clone-tail outliers are
retained rather than averaged away. This is not a universal MCE or Office
CRUD speedup, a memory reduction, or a full-save throughput claim.

The source remains one private helper, one predicate replacement and three
regression tests. The next investigation is the remaining marked XML parser
work and real-workflow tail variation; do not infer their cause from the
helper's smaller caller reservation or synthetic scan savings.

The complete pre-cleanup audit passes source/probe/binary bindings, exact patch
replay, raw statistics, control and follow-up audits, focused receipts and all
candidate gates. Cleanup removes exactly the batch target, binary and profile
directories plus its generated marker-control archive. The workspace Cargo.lock
is preserved; unrelated scratch is untouched. The packet retains the original
lockfiles, source witnesses, raw outputs and reproduction scripts.
