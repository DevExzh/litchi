# 0721 — DOCX structural scan fusion pilot

The candidate is rejected and all production source is restored byte for byte.
Fusion removed a structural traversal and improved edit medians, but it
regressed both read controls in every pair and failed one edit-mean gate.
No production optimization is retained. The performance program remains active.

The 64 top-level captures retain 9,600 native samples (6,400 primary and
3,200 read-control samples) and 48 allocator samples. Output parity passes.
One of 64 primary hard gates and all 16 read-control hard gates fail. In the
first pair, generated edit mean improves only 0.68%, below the required 3%;
its p95/p99 regress 20.87%/20.36%. No slow samples were removed.

| Scenario | Paired p50 delta range | Paired mean delta range |
| --- | ---: | ---: |
| Generated edit | −14.51% to −11.26% | −14.78% to −0.68% |
| NumberedList edit | −6.25% to −5.01% | −6.00% to −5.11% |
| Generated lifecycle | −0.78% to +1.54% | −0.66% to +1.47% |
| NumberedList lifecycle | −0.37% to +0.29% | −1.12% to +1.31% |
| Generated paragraph listing | +6.72% to +8.60% | +7.14% to +8.33% |
| Pinned-media eager paragraph count | +5.59% to +6.05% | +5.69% to +5.93% |

Deltas are candidate/baseline minus one; negative latency deltas are faster.
These ranges describe the rejected candidate on this host, not retained gains.
[Every pair and raw-derived comparison](results/change-0721/measurements.md)
remains visible. The primary lane flags three native tails and one
within-variant repeat drift above 5% (generated candidate edit mean, 15.21%).
Read controls flag 15 tails and seven tail/max repeat drifts, with no p50/mean
repeat drift above 5%. These flags do not exclude observations.

All allocator request-count and requested-byte gates pass. Aligned net-live
bytes are unchanged. Peak above region start is also unchanged except for
NumberedList edit: 21,310 → 23,790 bytes (+11.64%) in both allocator pairs.
[Memory diagnostics](results/change-0721/memory-diagnostics.json) retain every
aligned vector and formula. This is a separate memory cost, not process RSS.

The candidate combines the two structural reader traversals inside
`DocumentBody::from_xml`. Alternative-format anchors and outer block ranges
retain separate parser state, namespace predicates, depth/node/anchor limits,
resource labels and capture rules. A range error is deferred until the alt
parser and its first MCE selection succeed. Both ordered MCE calls and the
final body capture remain separate. Read-only callers retain their borrowed
reader path through the extracted range state.

The public oracle compares 19 XML controls and two complete packages, including
edited output, byte for byte. The candidate passes the DOCX all-feature,
all-target tests, formatting, warning-denied Clippy, doctests and rustdoc.
The benchmark harness passes 535 tests with one ignored. Independent legacy
scanner copies test namespaces, capture suppression, malformed markup, error
precedence and explicit resource boundaries. This does not prove identical
host allocator-exhaustion scheduling: interleaved storage changes allocation
lifetimes.

Temporary debug tracing has exact public output parity and deterministic
repeated stderr in both lanes. All 21 document boundaries preserve ordered
MCE inputs and outputs and structured metadata/ranges. At the two representative
boundaries, reader calls including EOF change as follows:

| Corpus | Main XML bytes | Baseline structural reads | Candidate structural reads | MCE input counts |
| --- | ---: | ---: | ---: | --- |
| Generated medium | 21,517 | 2,820 | 1,410 | 0, 200 |
| NumberedList | 4,563 | 356 | 178 | 0, 5 |

Observer callbacks are counted separately and are not reader calls. These are
logical parser traversals, not physical I/O, instructions or memory copies.
MCE and final body reads are outside those structural counts. On refusals,
the baseline range counter is recorded after its initial admission checks,
so its rejecting event is excluded. Work-reduction comparisons here use only
complete successful structural passes.

The host reports AMD EPYC 9R45, Linux 7.0.0-1012-aws, Rust 1.95.0
and the system allocator. Captures bind CPU 12; the filesystem root is on
ext4. Full build, source, environment and corpus identities are retained.

The native plan has two ABBA cycles, with 100 warmups and 200 samples per
corpus/phase/stage on CPU 12. Every edit p50/mean must improve at least 3%,
and every lifecycle p50/mean must regress no more than 3%. The first cycle
also uses separate allocator binaries at three samples and zero warmups,
collected after the native stages in the first cycle’s variant order;
request count and requested bytes must regress no more than 3%. Two read
controls have their own 3% p50/mean regression gate. Tails and repeat drift
above 5% remain visible; samples and pairs are never discarded.

Read controls are generated-medium semantic paragraph listing and the warm,
pinned `docx-source-backed-media` eager paragraph count. The latter selector
ignores `--ooxml-file`; that mismatch was caught before capture and the plan
binds its actual 16,793,036-byte corpus. Its eight top-level invocations each
launch 100 warmup, 200 priming and 200 measured filesystem children. The ordinary-save
NumberedList selectors do consume the named fixture.

Edit timing covers `owner.edit` with open outside. Lifecycle covers open,
edit and atomic save. Owner destruction, output readback and cleanup are
outside both clocks and allocation regions. The phases are not additive.
Allocator peaks and net-live counts are per-operation diagnostics. No RSS,
leak, cold-cache, hardware-counter, throughput or scaling claim is made.

The first four native captures completed before a read-driver command check
stopped execution: its expected ordinary-save command omitted NumberedList's
fixture argument. The archived failure contains no read-control samples.
The validator was corrected, its new hash frozen, and capture resumed using
the same four native reports. No samples were replaced or resampled.

The capture-time analyzers had integration errors in receipt inventory,
aggregate disposition, and filesystem priming counts. Separate analysis-only
copies correct those errors; the original scripts, exact diffs and hashes
remain retained. Capture scripts, gate formulas, thresholds and samples are
unchanged. The final analyses bind the restored baseline source.

The candidate passed correctness checks before rejection. Nine bounded
corruption checks verify the retained validators, and the final repository
evidence checks and cleanup are recorded in the packet. The frozen candidate
and its differential tests remain under `candidate/` for reproduction.

The next bounded hypothesis is to preserve the original read-only scanner
implementation while investigating writer-local fusion. The current data
localizes a read-path regression to this candidate but does not prove a compiler
inlining cause or promise that a different implementation will pass.

[Evidence and reproduction](results/change-0721/README.md).
