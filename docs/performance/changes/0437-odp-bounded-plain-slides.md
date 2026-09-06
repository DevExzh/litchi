# 0437: bounded fresh ODP titled-slide publication

The new plain-slide sink API removes whole-presentation model, content-string,
and archive-Vec retention from fresh ODP creation. Keep it as an opt-in bounded
publication facility: across 64, 4,096, and 8,192 slides, all 180 streaming
allocator samples peak at 420,352 bytes above operation entry. Large Builder
peak is 23,973,505 bytes, so this measured peak is 98.247% lower. Streaming
normal p50 is 1.412–1.614 times the candidate Builder's. No latency or RSS
improvement is claimed.

Production commit: `4560c9478`.
Before-build source revision: `203c50a848204ed51013b0555f61a3c7e75644c9`.
[Evidence and reproduction](../results/change-0437/README.md),
[summary](../results/change-0437/summary.json),
[retention decision](../results/change-0437/retention-decision.json).

## Change and boundaries

`litchi_odp::streaming::{stream_plain_slides_to,try_stream_plain_slides_to}`
consumes an ordered source of borrowed/owned `PlainSlide<S>` values and writes
to a caller-owned sequential sink. It retains one source item and a reusable
bounded page fragment. Titles, body text, page numbering, geometry, escaping,
and text controls follow the established plain Builder grammar. Rich slide
fields, existing-document append, and general presentation preservation remain
separate APIs and workloads.

The common `GeneratedXmlEnvelope::try_new_with_prelude` accepts fixed balanced
element-only prelude children before a final open insertion path. The old
`try_new` contract is unchanged. The common owner still validates XML, budgets,
ZIP publication, and manifest bindings. ODP retains its format grammar.

The formal provider uses a 4,096-byte fragment window. Its modeled Memory
reservation excludes caller inputs and extra ZIP/auditor allocations. Work
counts content shell/slide XML and fixed styles/meta XML, excluding manifest,
ZIP framing, compression, and auditing. XML depth has its own limit; execution
Depth is not charged. Cancellation is cooperative across at most a 256-byte
ordinary-text span. The source can be polled once beyond its slide limit to
distinguish exact exhaustion from excess input. Errors retain typed causes
and accepted sink-byte progress; failed output must be discarded.

The shared XML audit's ambiguous adjacent-control spacing refusal is preserved
and tested against Builder. Default limits reconstruct through checked
constructors. The proposed 32,768-slide shape exceeded Builder's default
250,000-attribute limit and was explicitly excluded before baseline capture;
no ceiling was relaxed. Native rendering was unavailable. Fixed auxiliary
parts reproduce the existing Builder, including its legacy style references;
decoded geometry/text parity is not visual compatibility evidence.

## Matched results

CPU 2, one worker, Rust 1.98.1 release builds with frame pointers/debug symbols;
AMD EPYC 9R45, Linux 7.0.0-1011-aws, Rust system allocator, 132,553,797,632
bytes physical memory and 4,096-byte pages. Storage identity is unavailable.
The retained build commands use four Cargo jobs with incremental disabled;
three warmups and 30 samples per fresh process. A1/B1/C1/C2/B2/A2 covers three
roles, two modes, three sizes, and two repeats: 36 reports/1,080 samples. Six
large normal profiles follow. The eighteen pilots and two preparatory profiles
remain separate. Both authoring roles create fresh input Strings inside the
operation timer and release operation buffers before it stops.

| Slides | Candidate Builder p50 R1 / R2 (ms) | Streaming p50 R1 / R2 (ms) | Streaming / Builder R1 / R2 |
| ---: | ---: | ---: | ---: |
| 64 | 0.251591 / 0.252051 | 0.355252 / 0.358132 | 1.412 / 1.421 |
| 4,096 | 12.624643 / 12.641038 | 20.370000 / 19.995559 | 1.614 / 1.582 |
| 8,192 | 25.504111 / 25.179240 | 40.138854 / 39.981843 | 1.574 / 1.588 |

The same-API buffered p50 deltas are +3.561/+3.202% (tiny),
+2.676/+1.494% (medium), and +2.102/+2.491% (large). One buffered control
flag crosses 5%: tiny R1 p99 rises 5.874%. All 36 cross-API normal latency and
throughput comparisons cross their adverse 5% thresholds. These are retained
tradeoffs, not averaged away. None of the repeat comparisons crosses 5%.
All 37 review flags and all 144 matched comparisons are in the summary.

| Slides | Builder peak above entry | Streaming peak above entry | Builder / streaming allocation calls | Builder / streaming requested allocation bytes |
| ---: | ---: | ---: | ---: | ---: |
| 64 | 592,188 | 420,352 | 2,985 / 1,323 | 2,170,651 / 1,739,284 |
| 4,096 | 11,988,609 | 420,352 | 180,405 / 73,899 | 32,749,223 / 5,392,276 |
| 8,192 | 23,973,505 | 420,352 | 360,631 / 147,627 | 63,817,229 / 9,103,252 |

These allocator values are identical across samples and repeats for the
named shape/role; the reported before/after Builder allocation values agree. Every operation's
live-byte delta is zero. Large requested allocation bytes fall 85.735% and
allocation calls fall 59.064%. Total requested allocation bytes still grow with
slide count; constant measured peak does not mean allocation-free operation.

GNU-time peak RSS ranges from 84,529,152 to 84,729,856 bytes over all 36
processes. No RSS comparison or repeat crosses 5%; no RSS reduction is claimed.
RSS includes setup, corpus construction/readback, warmups, samples, and binary
hashing, unlike the operation allocator region. An initial analysis omitted
tab-indented RSS lines; the reader was corrected to require exactly one value.
Original analysis artifacts remain in `versions`, and corrected summaries
reuse the unchanged raw measurements.

## Profiles and follow-up

| Whole-executable counter | Before Builder | After Builder | Streaming |
| --- | ---: | ---: | ---: |
| Cycles | 5,979,686,981 | 6,053,924,698 | 8,268,119,977 |
| Instructions | 23,610,712,638 | 23,641,795,045 | 29,975,205,327 |
| Branches | 5,221,699,616 | 5,227,417,827 | 6,199,522,533 |
| Branch misses | 10,611,961 | 9,955,167 | 13,490,860 |

Streaming self samples include SHA hashing 12.85%, execution `consume` 11.00%,
Deflate longest-match 10.78%, Deflate medium 6.88%, XML audit 5.39%, attribute
iteration 3.84%, and fragment-shape validation 3.17%. All three records report
zero lost samples; addr2line warnings remain retained. L1 readings are zero
without a load denominator; LLC is unavailable. These are whole-executable
profiles: the independent Python oracle and symbolization run afterward,
outside perf/GNU time, despite the older receipt scope string's overbroad
wording. No operation-only causal fraction or parallel Amdahl estimate follows.

The next concrete hypothesis is whether repeated fixed slide-markup Work
charges can be batched with exact scalar fallback at limits, while retaining
validation and typed progress. Its benefit remains unmeasured. Broader append,
Part addition, repackaging, real-producer/native, cold/range, and scaling
coverage remain open. The original non-iWork program is not complete.

## Verification

The final full release suites pass 349 ODP tests, 462 common ODF tests (one
existing ignored), and 364 standalone all-feature tests (one existing ignored).
Focused tests cover zero/empty/title/body distinctions, exact Builder member
parity, Unicode/control handling, malformed preludes, inclusive/one-under XML
and provider limits, hierarchical resource fallback, cancellation, lazy source
errors, short/interrupted/zero/partial sinks, and memory release.

Production all-target/all-feature Clippy passes with only the preexisting
common `large_enum_variant` lint allowed. The standalone strict comparison
retains the same 57 rendered diagnostics across 17 distinct message/file
findings, with none in changed ODP modules or changed harness lines. Docs,
minimal features, formatting, and crate boundaries pass. Failed development
attempts and resolutions remain retained. The bundle's portable replay checks
all source/build/artifact bindings, recomputes summaries and the retention
decision, and exercises independent corruption probes before and after
removing the twelve task scratch directories.
