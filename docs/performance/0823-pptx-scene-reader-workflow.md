# 0823 — PPTX Scene reader workflow trial

**Reject the candidate and retain the baseline production source.** Direct
`Reader` plus `NamespaceResolver` passes correctness and allocation checks, but
no eligible workflow reaches the frozen 3% paired p50 benefit threshold.
The real-file edit improves 2.394% (ratio 0.976062, 95% bootstrap interval
0.974798–0.984336). This is a measured result for the rejected candidate,
not a shipped speedup. All 19 rows remain visible below.

## Candidate and measured basis

The sole trial change was `crates/litchi-pptx/src/shape/reader.rs`, replacing
`NsReader` event transport in `Scanner::scan` with the same quick-xml 0.41
`Reader` and `NamespaceResolver`. Declaration pushes, deferred scope pops,
namespace error conversion, byte offsets, handlers, and resource/refusal order
were preserved. The original scanner was retained as a test-only differential
oracle. No API, dependency, cache, fingerprint, validation, or durability
contract changed. The complete candidate is archived; production is restored
byte-for-byte to base `b76786208d04310b4a033b39d70de439a64fcf69`.

An independent reread of the sealed 0822 profiles found 636/624 exact-owner
samples with `Scene::read_with`, 128/157 with `NsReader::process_event` under
Scene, and 124/116 with `NamespaceResolver::resolve_event` under Scene.
These inclusive counts overlap and are not additive phase fractions or
predicted savings. Fingerprint materialization remains required by the exact
revision contract; digest memoization already handles shared parent payloads.
The earlier duplicate-catalog candidate was below its measured noise floor.
No fingerprint shortcut or variable-size semantic cache was introduced.

## Workloads and decision rule

Six synthetic fixtures (tiny, medium, large, vendor extensions, Unicode vendor
extensions, and four valid attributes) each cover capture, commit, and full
in-memory lifecycle. Capture times opened-presentation capture; commit stages
the edit outside the clock; lifecycle includes capture, edit, commit, apply,
and sequential serialization. The fixture/output/semantic identities are
bound to the sealed 0806 qualification evidence through the 0813 probe.

The real-file workload opens `test-data/ooxml/pptx/shapes.pptx` (68,822 bytes,
SHA-256 `19fde9b87e33dd1a95fdbba0cf6abc2278bf03874f4665c7f8b88b6afe4a2571`)
before timing the ordinary public transaction, `set_shape_text(0, 0, marker)`,
commit, and apply sequence. The marker is `litchi-perf-0638-ordinary-save`.
The returned snapshot is dropped within the timed helper. Serialization,
readback, output hashing, and package destruction are outside the clock.
Every sample must reproduce the admitted 0821 output (68,284 bytes, SHA-256
`38c7fd3cc5037316e0a3a9c8cb6e6b013a95dd03b39bbd7fb357fc3cc3b9e5cf`),
six slides, complete text digest, and target shape text. There is no fsync or
atomic path save in this timed real-file region.

The frozen rule requires at least one of twelve synthetic commit/lifecycle
rows or real/direct to have paired median p50 ratio at most 0.97 and upper
95% bootstrap endpoint below 1.00. Six synthetic capture rows are negative
controls. Any of all nineteen rows with lower endpoint above 1.05 vetoes
adoption. Allocation calls, allocated bytes, net live bytes, and peak above
entry may not increase; RSS increases above 5% trigger review. This trial has
zero qualifying benefits, zero latency vetoes, zero allocation violations,
and zero RSS review triggers. The failed benefit gate alone rejects it.

## Native results

Each value is the median of six per-process nearest-rank p50 statistics;
ratios are medians of paired block ratios, not ratios of the displayed
medians. Times are milliseconds. Intervals use 10,000 bootstrap resamples,
seed 823823, sorted endpoints 250 and 9749, reset for each metric family.

| Scenario | Before p50 (ms) | After p50 (ms) | Paired after/before | 95% interval |
| --- | ---: | ---: | ---: | --- |
| synthetic/tiny/capture | 0.229351 | 0.228186 | 0.994921 | 0.987902–0.999082 |
| synthetic/tiny/commit | 0.207351 | 0.206421 | 0.995148 | 0.992732–1.000964 |
| synthetic/tiny/lifecycle | 1.406507 | 1.392022 | 0.989802 | 0.987827–0.992467 |
| synthetic/medium/capture | 0.427647 | 0.426522 | 0.997060 | 0.995065–1.003495 |
| synthetic/medium/commit | 0.290812 | 0.289102 | 0.996750 | 0.990528–0.998965 |
| synthetic/medium/lifecycle | 1.980500 | 1.962505 | 0.990261 | 0.984262–0.998348 |
| synthetic/large/capture | 16.049971 | 16.261857 | 1.012828 | 1.006000–1.020260 |
| synthetic/large/commit | 1.259636 | 1.252136 | 0.994412 | 0.989443–0.997103 |
| synthetic/large/lifecycle | 25.767209 | 25.593958 | 0.993038 | 0.992780–0.997843 |
| synthetic/vendor/capture | 0.512267 | 0.507428 | 0.995192 | 0.984021–0.999824 |
| synthetic/vendor/commit | 0.319732 | 0.318202 | 0.995779 | 0.990142–1.000756 |
| synthetic/vendor/lifecycle | 2.134261 | 2.118565 | 0.993314 | 0.987766–0.994833 |
| synthetic/unicode-vendor/capture | 0.513233 | 0.512943 | 0.998739 | 0.995665–1.003835 |
| synthetic/unicode-vendor/commit | 0.321572 | 0.317997 | 0.989178 | 0.985459–1.002139 |
| synthetic/unicode-vendor/lifecycle | 2.142266 | 2.125221 | 0.992766 | 0.988008–0.995215 |
| synthetic/valid-4attr/capture | 0.489312 | 0.491198 | 1.000261 | 0.996654–1.007262 |
| synthetic/valid-4attr/commit | 0.312372 | 0.311297 | 0.997182 | 0.992313–0.999713 |
| synthetic/valid-4attr/lifecycle | 2.094046 | 2.078596 | 0.991838 | 0.990600–0.993996 |
| real/real/direct | 1.439947 | 1.405687 | 0.976062 | 0.974798–0.984336 |

Large synthetic capture is 1.283% slower with an interval above 1.00; it does
not reach the 5% veto but is not a benefit. All eligible synthetic point
improvements are below 3%. The real-file row is the largest eligible point
improvement and still misses the adoption threshold.

Real-file p95 is 1.451532→1.416682 ms, p99 1.455732→1.418167 ms, and mean
1.439894→1.405523 ms. Median process RSS is 5,118→5,262 KiB; RSS is a whole
process metric, not retained edit memory. Per-process p50/p95/p99/mean/RSS,
paired intervals, and all spread flags for every row are retained in
[analysis.json](results/change-0823/analysis.json) and
[native.csv](results/change-0823/native.csv).

No row has a native p50, p95, or mean process spread above 5%. Nine rows have
p99 process spread above 5% in at least one leg, and eleven rows have at
least one process with p99/p50 above 1.05. RSS spread flags occur in multiple
rows, including real/direct in both legs. These short-process tail and RSS
observations do not establish a tail-latency or memory saving. No aggregate
hides the row-level results.

## Allocation and code generation

All nineteen rows have identical observed allocation calls, allocated bytes,
net live bytes, and peak above entry before/after. The real edit retains
7,682 calls, 2,719,494 allocated bytes, 158,985 net live bytes, and 305,970
peak bytes above entry. The complete four-metric table is in
[analysis.md](results/change-0823/analysis.md). Observer timings are excluded
from native timing claims.

The exact native real-file binary's `Scene::read_with` symbol grows from
10,276 to 11,485 bytes. Its direct `NsReader::process_event` call disappears.
Neither listing exposes a direct named `resolve_event` or namespace `push`
call, so their absence cannot establish that their underlying work vanished.
The archived disassembly supports a transport/code-generation change, not a
causal instruction cost or a sufficiently large workflow improvement.

## Verification, failures, and cleanup

Both production versions pass fmt, all-target/all-feature check, tests,
warning-denied Clippy, warning-denied rustdoc, and the crate-boundary audit.
Baseline totals are 1,241 passed/0 failed/3 ignored; candidate totals are
1,249 passed/0 failed/3 ignored. Eight added differential tests cover rich
DrawingML semantics, BOM offsets, namespace shadowing/restoration, malformed
XML, reserved namespaces, declaration ceilings, and scanner limits.

Both packet-local probes pass eighteen total quality commands across default
and allocator features. Three initial probe-quality failures are retained with
full source snapshots and logs: feature-disabled dead code, a test-build
unused helper, and ambient global-allocator callbacks contaminating manual
counter tests. Feature/test gating and test mutex isolation were repaired
before freeze; timed corpus bodies and release counter arithmetic are intact.

Qualification has 38 reports/38 samples; native has 228 reports/6,840 samples;
allocation has 76 reports/228 samples. Fresh processes run serially on CPU 12
of the recorded AMD EPYC 9R45 host with three native warmups and thirty
samples. Six native block orders and two allocation block orders are
counterbalanced. Release builds use locked offline dependencies, opt-level 3,
thin LTO, one codegen unit, debug level 1, unwind, and two build jobs.

The first native attempt was interrupted by an unexpected analysis-agent
commit that changed HEAD. The source-state guard stopped it. Its nine report
artifacts, logs, and exact commit patch are archived separately and excluded
without timing analysis; root restored the owned local commit's parent while
preserving files and restarted the entire native matrix with unchanged frozen
inputs. No successful subset was reused. Git mutations remained root-owned
for the remaining trial.

The first offline analysis invocation failed on an origin-field schema
mismatch. Its original log and complete reader sources are retained; lock
identity validation was also aligned with the frozen hash-only descriptors.
The corrected analysis and independent raw audit agree for all 342 admitted
reports and 7,106 samples. Codegen and final cleanup validation pass.
Cleanup verifies all eight binaries and removes only the owned target:
9,687 files and 3,238,577,403 logical bytes. Candidate and raw evidence remain
archived, while production has no net change.

This is one host, one admitted real file, and a fixed synthetic corpus with
warm filesystem/provider state. It proves no cold-cache, cross-format,
producer-wide, tail, RSS, full-save, or historical speedup. Future work should
revisit required versus repeated Scene construction/commit compaction, with
preservation and bounded-cache constraints intact; this rejected transport
rewrite is not a basis for claiming broader gains.

Replay and exact commit custody are described in
[the packet](results/change-0823/README.md).
