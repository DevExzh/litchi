# Change 0694 — borrow expanded MCE element names until ownership is needed

performance_claim: none

## Decision and measured scope

The candidate removes temporary namespace/local-name Strings from ordinary
shared-MCE start-element decisions. An owned Name remains for nonempty custom
extension profiles and ignorable-element preservation/process lookup; their
existing HashSet and pattern matching algorithms are unchanged. The same
QName validator runs at the same position, including inside opaque descendants.
No persistent cache, source identity, API, output grammar or execution policy
changes. The production diff is confined to one private function.

On the pinned warm shared host, real-deck one-edit phase totals improve
**4.12–5.26%**, no-op **2.58–3.93%**, and two-slide edits **3.43–3.78%**.
The native instruction/cycle diagnostic supports the allocation mechanism.
These results and the final correctness gates justify retaining the narrow
change. They do not establish faster Office workflows generally,
full save latency, cold I/O, remote input or concurrent scaling.

The [packet](results/change-0694/README.md) retains commands, exact source/build/
corpus bindings, native samples, allocations, profiles, refusals and reviews.
Baseline is `829bed696`; native and allocator executables are separate.
The borrowed-name change follows 0693's removal of repeated slide processing:
fresh baseline MCE `start` still accounts for 12.48% of native self samples.

## Native results and allocation mechanism

The same 13 workflows run two A/A legs followed by A/B/B/A, each with 100 samples
and five warmups: 7,800 native samples. Timer scope is capture, working clone,
text edit, commit and publication after package materialization and target
selection. Semantic checks and serialization/reopen validation are outside the
timers. Marker controls deliberately replace the MCE URI and are counterfactual
mechanism inputs, not proof of equivalent semantics. Individual legs and
bootstrap intervals remain in the packet; there is no pooled aggregate claim.

| Workflow / input | Baseline medians (ms) | Candidate medians (ms) | Paired change |
| --- | ---: | ---: | ---: |
| one-real | 11.9960 / 12.2711 | 11.5021 / 11.6256 | -4.12% / -5.26% |
| one-control | 6.0266 / 5.9966 | 5.9630 / 6.0042 | -1.05% / +0.13% |
| one-generated | 1.5634 / 1.5538 | 1.5621 / 1.5562 | -0.08% / +0.16% |
| one-notes-poi | 1.0428 / 1.0473 | 1.0402 / 1.0487 | -0.25% / +0.13% |
| one-notes-lo | 1.6481 / 1.6561 | 1.6522 / 1.6600 | +0.25% / +0.24% |
| noop-real | 5.5381 / 5.5814 | 5.3950 / 5.3620 | -2.58% / -3.93% |
| noop-control | 2.9079 / 2.8560 | 2.9424 / 2.8823 | +1.19% / +0.92% |
| noop-generated | 0.7619 / 0.7616 | 0.7603 / 0.7594 | -0.21% / -0.29% |
| noop-notes-poi | 0.4816 / 0.4726 | 0.4719 / 0.4715 | -2.01% / -0.23% |
| noop-notes-lo | 0.6056 / 0.6101 | 0.6064 / 0.6129 | +0.14% / +0.46% |
| two-real | 12.3282 / 12.2277 | 11.8617 / 11.8086 | -3.78% / -3.43% |
| two-control | 6.1722 / 6.1212 | 6.1528 / 6.0989 | -0.32% / -0.36% |
| two-generated | 1.7892 / 1.7812 | 1.7711 / 1.7776 | -1.01% / -0.21% |

| Real one-edit phase | Baseline medians (ms) | Candidate medians (ms) |
| --- | ---: | ---: |
| capture | 5.1976 / 5.3012 | 4.9500 / 5.0058 |
| clone | 0.0166 / 0.0164 | 0.0164 / 0.0165 |
| settext | 1.1021 / 1.1322 | 1.0732 / 1.0643 |
| commit | 5.5961 / 5.7198 | 5.3659 / 5.4464 |
| apply | 0.0943 / 0.0926 | 0.0920 / 0.0922 |

| One-edit allocation diagnostic | Baseline | Candidate |
| --- | ---: | ---: |
| alloc_calls | 190,338 | 143,848 |
| realloc_calls | 4,870 | 4,870 |
| requested_bytes | 13,815,905 | 12,430,874 |
| peak_above_start | 459,170 | 459,170 |
| net_live_change | 195,730 | 195,730 |

The real edit eliminates **46,490 allocation calls** and **1,385,031 requested
bytes**; all three repeats are retained. Reallocation count, end-of-phase live
storage and peak-above-start for this case are unchanged. Requested bytes include
the full replacement size of reallocations; realloc-requested is not added a
second time. Counters measure logical allocation requests, not RSS. Native
latency never uses the counting allocator.

## Costs and uncertainty

Initial A/A total medians vary from −1.58% to +1.74%; mean and p95 changes remain
within 1.60%. A no-op POI p99 falls 6.44% across A/A. Baseline real one-edit
A/B/B/A legs differ by 2.29%, so report both candidate pairings rather than
calling a single percentage exact. Marker-free total medians remain within
about 2.1%, and no total p50/mean/p95/p99 candidate metric triggers +5%.

Twenty individual phase metrics exceed +5% and remain in
`native-review-triggers.json`. No-op real clone p50 rises 6.44% (+1.015 µs),
and no-op control clone p50 rises 5.65% (+0.880 µs), in the first pair.
The largest relative tail is no-op LibreOffice clone p99 +109.19% (+5.940 µs);
real one-edit clone p99 increases 30.93% (+5.540 µs). This is no universal tail
improvement claim, and these costs are not averaged away.

The ten-case public capture/refusal matrix retains 6,000 samples and identical
exact graph/error metadata. Median changes range from −4.36% to +3.36%.
Its sole +5% trigger is generated-valid capture p99 +17.49% (+111.921 µs) in
one pair. It remains a disclosed tail observation, not a new error contract.

Native open-plus-capture counter slopes use `(210 − 10) / 200`: instructions
fall 3.28%, cycles 2.57%, branches 4.00%, cache misses 2.85%, and task-clock
2.52%. Branch misses rise 3.95% and page faults rise **10.41%** (182.410 →
201.395 per operation). Single-child peak RSS is 5,696 → 5,700 KiB (+0.07%).
These diagnostics do not prove improved faults or an RSS bound. Native text
shrinks 920 bytes (2,595,586 → 2,594,666); sampled MCE `start` remains the top
self symbol at 11.60%. Repeated presentation work and the rest of the parser
remain open; the five presentation passes must be priced by their byte/cycle
weight before another reuse design is prioritized.

Secondary synthetic MCE controls use 1,000 ordinary or opaque-descendant
nodes, three profiles (no opaque names, one name, 4,096 names), A/A and A/B/B/A,
and 300 samples per process. All 10,800 exact-output-checked samples are retained.
Default-profile medians improve 6.39–7.76%; opaque descendants with explicit
profiles improve 6.12–9.71%. Ordinary elements under a **nonmatching explicit
profile regress 2.39–6.11%** (one name) or **3.27–5.36%** (4,096 names).
These costs concern the shared processor, not a measured whole XLSX/media
workflow. The public custom-profile path is retained and receives additional
real-part controls below; no claim of universal profile improvement is made.


The supplemental real-part control uses the largest marker-bearing identity
case per format in the frozen oracle corpus: a 56,460-byte DOCX numbering
part, a 19,886-byte XLSX worksheet and a 2,906-byte PPTX slide. These are shared
processor measurements with synthetic nonmatching extension names, not full
conditional-formatting, protection or media workflows. It retains 16,200
samples and verifies exact output identity across every leg.

| Real XML family | Profile | Paired median change |
| --- | --- | ---: |
| DOCX | default | -0.36% / -2.56% |
| DOCX | one nonmatching name | +2.98% / +1.96% |
| DOCX | 4,096 nonmatching names | +2.40% / +0.93% |
| XLSX | default | -1.76% / -4.37% |
| XLSX | one nonmatching name | +3.80% / +1.06% |
| XLSX | 4,096 nonmatching names | -1.81% / +1.48% |
| PPTX | default | -7.35% / -7.16% |
| PPTX | one nonmatching name | +1.90% / +3.80% |
| PPTX | 4,096 nonmatching names | -0.16% / +3.97% |

Real-part custom-profile medians span −1.81% to +3.97%, so the synthetic
ordinary-element cost is not a universal 6% consumer regression. The PPTX
one-name first pair nevertheless has mean +39.97% and p99 +15.54%. One
3,072,415 ns sample appears in that candidate leg (median 26,320 ns); the
second pair has mean +4.25% and p99 +2.36%. Keep the outlier and both pairs,
without asserting an unverified scheduling cause or dropping it from the mean.

Disposition: retain the small shared change for the measured default-profile
workflow and allocation gains, accepting the disclosed nonmatching-profile
and tail costs. Custom extension lookup stays hashed, with no linear scan or
new retained state. Neither this control nor the primary PPTX matrix supports
a universal latency improvement or a full XLSX-feature performance claim.

## Correctness and reproducibility

The [design](results/change-0694/design.md) maps the accepted ADR constraints.
All 33 GOAL/ADR hashes remain unchanged. Existing finite limits, fallible
reservations, namespace installation, directive validation, opaque subtree
handling, output filtering and report accounting remain at their original
sites. Removing infallible String allocations removes those particular abort
opportunities; it adds no recoverable-allocation guarantee.

Four new tests cover expanded default/prefixed names with matching/mismatched
extension profiles, inherited opaque content, namespace rebinding and directive
shadowing, and QName-before-directive refusal. All 85 focused MCE tests pass.
The initial run caught two new test expectations that omitted filtered MCE
directives from ignored-attribute counts; the corrected complete Reports pass.
The failed run is retained and no production behavior changed to satisfy it.

The oracle baseline was built from exact baseline source after archiving the
candidate codec/tests; both files were restored byte-for-byte in a finally
block. A copied oracle lockfile initially failed `--locked` before compilation;
an offline lock refresh is retained. Both successful oracle builds use the
same final oracle lockfile. It is independent of the main probe lockfile.

The exact-output differential covers **192 cases × five profiles** (1,920
binary invocations): 36 XML members from 12 DOCX/XLSX/PPTX archives, 144
deterministic mutations and 12 synthetic boundary cases. All output
SHA-256/length/ownership/Report records and typed Debug refusals match.
Each side has 523 accepted profile/case combinations and 437 refusals. The
opaque-descendant invalid QName case refuses identically in all five profiles.
This compares existing behavior, including the marker-free borrowed bypass;
it is not an independent full XML-conformance validator.

All final Rust gates pass: workspace formatting; all-feature/all-target checks
and warning-denied library Clippy for the common owner plus DOCX/XLSX/PPTX;
**1,208 default tests** (two existing ignored), **4,132 all-feature tests**
(33 existing ignored), **45 facade tests**, and warning-denied rustdoc.
The facade preserves the preceding batches' normal-warning policy for unrelated
pre-existing unused helpers. Owner checks retain warning denial. No cargo-fuzz
command or nightly toolchain is installed; no fuzz/Miri execution is claimed.

All six repository boundary, strict/structural claim, report-classification,
coverage and non-iWork gates pass. Final source/data audit, independent evidence
review and exact owned-scratch cleanup are retained with the post-cleanup seal
in the packet. The workspace lock and shared build targets are preserved.
