# 0824 — exact-root proof for PPTX compaction

**Adopt the candidate.** On the admitted real-file public shape-text edit,
paired median p50 improves **10.407%**: after/before ratio **0.895935**, with
95% bootstrap interval **0.892216–0.899349**. Process-median p50 changes from
1.436052 to 1.286512 ms. Allocation calls fall from 7,682 to 7,384 and allocated
bytes from 2,719,494 to 2,697,806. All nineteen rows pass the frozen latency and
allocation guards. Several synthetic rows are slower; the complete results
and limits follow. This is not a full-save or cross-format speedup.

## Change and correctness argument

The source allowlist is `crates/litchi-pptx/src/opened/transaction.rs` and
`crates/litchi-pptx/src/opened/xml.rs`, based on
`7b268927bfe3ed9c9000bfd567fc4ce656ec0cd9`. The existing compactor records the
complete document-root spans in source and output during its event pass.
It compares those bytes exactly and separately verifies exact copying of
outside declarations, processing instructions, and comments. The BOM is
preserved. Outside ASCII whitespace is removed under the existing policy;
invalid outside content and DTDs retain their existing refusals.

When compaction changes only that whitespace, the successful staged Scene
validation proves the unchanged root's semantics. The existing Weak allocation
record still identifies the exact validated staged payload. A valid record
plus the root proof removes two duplicate `compaction_scene` reads; a missing
or stale record retains the required initial read and removes only the second.
A root change or unavailable proof keeps the original semantic comparison.
Byte-identical compaction keeps its existing behavior, including initial
validation when the staged record is absent. Debug builds rederive semantic
equality. No cache, retained parsed state, dependency, public API, emitted-byte,
fingerprint, source-authorization, resource-limit, or durability change is added.
Final snapshot capture and publication validation still run on the final bytes.

The mechanism follows ADRs 0003, 0005, 0006, and 0032: eliminate duplicate work
using a local exact-byte proof while retaining validation, preservation, and
bounded ownership. Root ranges and the verdict are transient constant-size
state. Independent [source review](results/change-0824/candidate-review.md)
and the complete before/after candidate are retained.

## Measured basis and workload scope

The sealed 0822 profiles contain 636/624 exact-owner Scene-qualified samples;
301/304 also contain `compaction_scene`. These are overlapping inclusive stack
counts, not additive phase fractions or predicted savings. The independent
profile reader rechecks the original compressed/decoded evidence. Required
fingerprint work remains unchanged; no digest or CRC shortcut is introduced.

The admitted real input is `test-data/ooxml/pptx/shapes.pptx`, 68,822 bytes,
SHA-256 `19fde9b87e33dd1a95fdbba0cf6abc2278bf03874f4665c7f8b88b6afe4a2571`.
The byte-basis reader proves that after replacing the selected text with
`litchi-perf-0638-ordinary-save`, the published slide differs only by removal
of CRLF after its XML declaration. The timed public transaction performs
`set_shape_text(0, 0, marker)`, commit, and apply. The returned publication
snapshot is dropped inside the helper. Open, serialization, readback, output
hashing, package-owner destruction, and path-save/fsync are outside the clock.
Every sample reproduces the admitted 0821 output: 68,284 bytes, SHA-256
`38c7fd3cc5037316e0a3a9c8cb6e6b013a95dd03b39bbd7fb357fc3cc3b9e5cf`,
with six slides, the complete text digest, and target shape text checked.

Six fixed synthetic fixtures (tiny, medium, large, vendor, Unicode vendor,
and four valid attributes) each cover capture, commit, and lifecycle.
Capture precedes compaction and is a negative control; commit stages outside
the clock; lifecycle includes in-memory serialization. Probe timed bodies,
corpus, and locks match the final 0823 probes; only packet identity changed.
Synthetic output oracles remain bound to the sealed historical qualifications.

The frozen benefit rule requires a paired median p50 ratio at most 0.97 and
upper interval below 1.00 in one of thirteen eligible commit/lifecycle/real
rows. Any of nineteen rows with a lower interval above 1.05 vetoes adoption.
Allocation calls, allocated bytes, net live bytes, and peak above entry may
not increase. Paired process RSS increases above 5% require review. There is
one qualifying benefit, zero latency vetoes, zero allocation violations, and
zero RSS review triggers.

## All native timing rows

Times are milliseconds, each the median of six per-process nearest-rank p50
values. Ratios are medians of paired block ratios, not ratios of displayed
medians. Intervals use 10,000 resamples, seed 824824, sorted zero-based endpoints
250 and 9749, reset for each metric family.

| Scenario | Before p50 (ms) | After p50 (ms) | Paired after/before | 95% interval |
| --- | ---: | ---: | ---: | --- |
| synthetic/tiny/capture | 0.229226 | 0.228716 | 0.997557 | 0.994500–1.020506 |
| synthetic/tiny/commit | 0.207556 | 0.206666 | 0.995998 | 0.994748–0.999761 |
| synthetic/tiny/lifecycle | 1.407902 | 1.397652 | 0.993913 | 0.990330–0.996771 |
| synthetic/medium/capture | 0.425832 | 0.431507 | 1.011316 | 1.009936–1.016247 |
| synthetic/medium/commit | 0.290892 | 0.290687 | 0.999295 | 0.997484–1.000498 |
| synthetic/medium/lifecycle | 1.979289 | 1.979154 | 0.998797 | 0.996136–1.003514 |
| synthetic/large/capture | 16.136125 | 16.725542 | 1.036363 | 1.034552–1.044513 |
| synthetic/large/commit | 1.261646 | 1.267396 | 1.005191 | 0.999137–1.007173 |
| synthetic/large/lifecycle | 25.875311 | 26.197253 | 1.012505 | 1.008338–1.024666 |
| synthetic/vendor/capture | 0.511013 | 0.521293 | 1.022539 | 1.011212–1.027563 |
| synthetic/vendor/commit | 0.319217 | 0.320967 | 1.006086 | 1.003927–1.007458 |
| synthetic/vendor/lifecycle | 2.136445 | 2.140640 | 1.001965 | 0.997188–1.005611 |
| synthetic/unicode-vendor/capture | 0.513552 | 0.524898 | 1.021207 | 1.016350–1.023558 |
| synthetic/unicode-vendor/commit | 0.319826 | 0.320707 | 1.003955 | 1.001156–1.010912 |
| synthetic/unicode-vendor/lifecycle | 2.142321 | 2.153520 | 1.005617 | 1.000551–1.008742 |
| synthetic/valid-4attr/capture | 0.489402 | 0.503923 | 1.029132 | 1.024277–1.033094 |
| synthetic/valid-4attr/commit | 0.312566 | 0.313082 | 0.999105 | 0.997944–1.004020 |
| synthetic/valid-4attr/lifecycle | 2.095550 | 2.102475 | 1.003216 | 1.002497–1.007081 |
| real/real/direct | 1.436052 | 1.286512 | 0.895935 | 0.892216–0.899349 |

Large synthetic capture is 3.636% slower, with interval 1.034552–1.044513;
large lifecycle is 1.251% slower. Medium/vendor/Unicode/valid-attribute capture
controls also slow down. These are retained observations, below the frozen
veto threshold. The timed capture path does not invoke compaction, so this
trial cannot attribute its change to the shortcut itself. No synthetic benefit
is claimed, and the controls are not subtracted from the real-file result.

Real-file process-median p95 changes 1.450482→1.297442 ms, p99
1.458122→1.307317 ms, and mean 1.435489→1.287716 ms. Median process RSS is
5,166→4,934 KiB. These do not establish tail-latency or RSS savings: the real
candidate p99 process spread is 41.663%, and both legs have RSS spread above
5%. Across all native rows, none has p50 or mean process spread above 5%, one
has p95 spread above 5%, thirteen have p99 spread above 5%, sixteen contain a
process with p99/p50 above 1.05, and seven have RSS spread above 5%.
All per-process statistics, paired intervals, and flags remain visible in
[analysis.json](results/change-0824/analysis.json) and
[native.csv](results/change-0824/native.csv).

## Allocation observations

The real edit saves 298 allocation calls and 21,688 allocated bytes. Net live
bytes remain 158,985 and peak bytes above entry remain 305,970. All eighteen
synthetic rows have identical observed calls, allocated bytes, net live bytes,
and peak above entry on both legs. The complete four-metric table is retained
in [analysis.md](results/change-0824/analysis.md). Instrumented elapsed times
are excluded from native timing claims. No retained-memory or RSS reduction
is asserted.

## Validation, retained failure, and cleanup

Baseline quality passes 1,241 tests and candidate quality passes 1,253 tests,
each with zero failures and three ignored. Both versions pass formatting,
all-feature/all-target compilation, warning-denied Clippy and rustdoc, and
crate-boundary checks. The two probes pass eighteen quality commands across
default and allocator features. Twelve added tests cover exact-root/BOM/empty
root bounds, declarations/PI/comments, unknown markup, fallback for inner
whitespace and attribute changes, invalid outside/root content, cold/staged
read counts, and all six slides of the admitted real fixture. Existing
namespace/MCE, resource-limit, stale-allocation, and refusal tests also pass.
This does not claim a broad real-producer or active-MCE differential corpus.

The initial candidate quality attempt failed an incorrect newly added test
expectation: an empty staged-read map still requires one initial validation
read even for byte-identical compaction. Only that expectation and its message
were corrected from zero to one. The original source, freeze, review, build,
qualification, failed receipt, source census, and logs are archived under
`pre-repair-0` and `quality-after-0`. The first baseline release outputs were
verified and removed (1,031 files; 768,731,181 logical bytes), then a fresh
freeze, complete baseline build, and baseline qualification ran. No candidate
release build or comparative timing existed before that restart. The successful
full candidate quality then preceded its build and qualification. A dedicated
recovery audit verifies the exact test-only difference and original failure.

The admitted trial retains 38 qualification reports/38 samples, 228 native
reports/6,840 samples, and 76 allocation reports/228 samples. Processes run
serially on CPU 12 of the recorded AMD EPYC 9R45 host, with six counterbalanced
native blocks (three warmups and thirty samples) and two allocation blocks
(no warmup, three samples). Builds use offline locked Cargo, opt-level 3,
thin LTO, one codegen unit, debug level 1, unwind, two jobs, and no incremental
compilation. There is no exclusive-host claim.

Analysis, independent raw arithmetic, profile/fixture basis, failure recovery,
and final validation all pass. No offline reader invocation failed. Cleanup
verifies all eight admitted binaries and removes only the owned target:
9,686 files and 3,238,494,566 logical bytes. Raw evidence, candidate sources,
failed-attempt evidence, and hashes remain available for offline replay.

This result is limited to one host, one admitted real file, a fixed synthetic
corpus, and warm filesystem/provider state. It establishes no cold-cache,
cross-format, broad-producer, full-save, fsync, tail, RSS, or historical speedup.
The broader non-iWork performance goal remains open. Next work should measure
the effect in complete ordinary-save workflows and revisit the remaining
required Scene/fingerprint costs with their contracts intact.

See the [packet README](results/change-0824/README.md) for replay and exact
staged/committed custody checks.
