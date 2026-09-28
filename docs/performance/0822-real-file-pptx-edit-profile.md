# 0822 — real-file PPTX edit CPU profile

The real-file public PPTX edit profile retains **23 reports and 4,549 samples**.
The ordinary direct probe has a median process p50 of **1.428437 ms**. The
wrapper control is indistinguishable within its paired interval; frame pointers
add a small measurable perturbation. Two profiles retain 2,971 and 2,996
exact-owner samples. XML attributes, decompression, and SHA-256 lead the sampled
leaves. Production and the ordinary-save harness are unchanged; this packet
identifies work to investigate and adopts no optimization.

## Operation and evidence boundary

The input is the admitted `test-data/ooxml/pptx/shapes.pptx`, 68,822 bytes,
SHA-256 `19fde9b87e33dd1a95fdbba0cf6abc2278bf03874f4665c7f8b88b6afe4a2571`.
The expected output is the 0821 real-PPTX publication, 68,284 bytes,
SHA-256 `38c7fd3cc5037316e0a3a9c8cb6e6b013a95dd03b39bbd7fb357fc3cc3b9e5cf`.

The packet-local probe uses the ordinary-save public edit sequence:
`opened_presentation_transaction` → `set_shape_text(0, 0, marker)` → `commit`
→ `apply_opened_presentation_commit`. The marker is
`litchi-perf-0638-ordinary-save`. The transaction and commit must both report a
change. The returned publication snapshot is dropped inside the edit boundary,
matching the original harness; the package owner remains alive until afterward.

Each iteration opens a fresh package from the path before starting the clock.
Serialization, output hashing, semantic reopening, and owner destruction occur
after the clock. Serialization produces an owned byte vector before the
package is dropped; semantic verification then uses that vector. The reference
is independently read from the sealed 0821 output. No pre-timing snapshot capture warms the new owner's edit memo.

The probe is diagnostic evidence for this public edit sequence, not the
ordinary-save benchmark executable. It contains no save/fsync in the measured
region and changes no production API, preservation rule, or durability default.
The earlier durability comparison and this edit profile are separate workloads;
their medians and sampled percentages cannot be combined into phase fractions.

## Controls and capture

| Arm | p50 (ms) | p95 (ms) | p99 (ms) | Mean (ms) | RSS (KiB) |
| --- | ---: | ---: | ---: | ---: | ---: |
| control (ordinary/direct) | 1.428437 | 1.438872 | 1.452577 | 1.428771 | 5,122 |
| ordinary/wrapped | 1.428187 | 1.441948 | 1.448183 | 1.429399 | 4,940 |
| fp/wrapped | 1.443077 | 1.452523 | 1.459002 | 1.443994 | 5,142 |

Values are medians of six process statistics, not pooled sample quantiles.
Wrapped/control p50 ratio is **1.000137**, 95% bootstrap interval
**[0.997252, 1.002430]**. FP/wrapped is **1.009094**,
**[1.005815, 1.012671]**. Neither comparison is a production speedup.
There are no latency spread or p99/p50 flags above 5%. Whole-process RSS
spread exceeds 5% for wrapped (5.62%) and fp (5.26%); it is not edit-region
allocation or retained-memory evidence.

The native controls assess direct versus wrapped call overhead and the effect
of compiling with frame pointers. Their order rotates across six deterministic counterbalanced blocks on CPU 12
of the retained AMD EPYC 9R45 host. Each report is a fresh process with three
warmups and thirty measured edits; filesystem/provider caches are warm,
without a physical cold-cache claim.
Nearest-rank p50/p95/p99 are computed within each process; medians and paired
ratios summarize the six blocks. Bootstrap intervals use 10,000 resamples,
seed 822822, with sorted endpoints 250 and 9749.

Perf captures use the frame-pointer binary, user cycles, and a separate exact
wrapper symbol. Only stacks carrying that symbol in the expected executable
contribute to operation attribution. Whole-process samples, unresolved frames,
lost events, and unattributed samples retain separate denominators. Sampled
leaf rankings are descriptive; inclusive stacks overlap and do not yield an
Amdahl serial fraction or a causal cycle cost.

## Findings and next action

| Sample census | Repeat 0 | Repeat 1 |
| --- | ---: | ---: |
| Whole process | 6,164 | 6,160 |
| Exact edit owner | 2,971 | 2,996 |
| Outside exact owner | 3,193 | 3,164 |
| Empty stacks (included outside owner) | 0 | 1 |
| Owner-qualified unknown leaves | 2 | 3 |
| Owner-qualified unknown interiors (excluding leaf) | 2 | 2 |
| Wrapper self leaves | 0 | 0 |

No lost-event lines appear in the retained recording or decode logs. Maximum
observed stack depths are 27/26 frames; 23/20 whole-process stacks end at an
unknown frame. No explicit truncation marker appears, but complete unwinding
is not proven. The full depth histograms are retained in `stack-diagnostics.json`. Every
owner-qualified sample belongs to exactly one self-leaf row. Outside-owner
samples include all untimed setup and verification; this partition is not an
elapsed edit fraction. Incomplete unwinding can also hide an owner, so sampled
coverage is not a proof of complete CPU attribution.

| Selected owner-qualified descendant self leaf | Repeat 0 | Repeat 1 |
| --- | ---: | ---: |
| `quick_xml::events::attributes::IterState::next` | 184 | 200 |
| `zlib_rs::inflate::inflate_fast_help` | 172 | 169 |
| `sha2::sha256::x86_sha::compress` | 166 | 173 |
| `quick_xml::name::NamespaceResolver::resolve_event` | 136 | 119 |
| `quick_xml::reader::ns_reader::NsReader<R>::process_event` | 129 | 128 |
| `litchi_pptx::shape::reader::Scene::read_with` | 111 | 108 |

The full leaf ranking, including libc and other parser leaves, is retained in
`frame-audit.json` and the analysis. Observed decompression and SHA stacks run
through `package_fingerprint_with_memo` → `capture_internal` →
`opened_presentation_with_limits`. Shape-reader stacks also occur beneath
`compaction_scene` → `Transaction::commit`.

The next bounded source review should trace initial snapshot fingerprint
materialization and repeated XML/shape reads, then distinguish required
source authorization and validation from avoidable work. The profiles do not
show that any validation can safely be omitted. No low-level rewrite or
performance benefit is inferred from these leaf counts; a candidate still
requires preservation admission and matched end-to-end before/after evidence.

## Verification and limits

Exact current source identity reuses the committed production six-gate
result: 641 tests passed, zero failed, one ignored. The packet-local probe
passes all five fresh gates (fmt, locked offline release check, all three
focused tests, warning-denied Clippy, and warning-denied rustdoc).
Two fresh release binaries have frozen source/lock/build receipts. The unique
mangled profiling owner matches the exact demangled address and size; its
assembly retains a frame-pointer prologue and call boundary.

Three qualification reports cover nine samples; eighteen native reports cover
540 samples; two profiler reports cover 4,000 separately classified samples.
Every sample checks full output byte equality, SHA-256, size, reopened complete
text, slide count, and target shape text against the independent admitted
reference. The first frame-audit attempt rejected a legitimate empty stack;
its script and failure log are retained, and the corrected audit accounts for
that stack explicitly without changing any capture.

Offline analysis, independent native/frame audits, stack diagnostics, and final
post-cleanup validation pass with 23 reports/4,549 samples. Six preliminary reader-schema failures are recorded by agent attestation;
their original console logs were not retained. Final raw-data replay passes.
The duplicate analysis write attempt was refused without altering the
successful outputs.
Cleanup verifies both exact executables before removing the owned target:
3,684 files and 2,054,019,575 logical bytes. Lossless compressed perf/frame
archives remain; uncompressed copies were removed only after verification.

The checked-in fixture and one host do not establish all producers or document
sizes. Frame-pointer compilation and sampling can perturb code generation and
scheduling. Native probe timings and profiled timings remain separate. No
production optimization is adopted without a subsequent before/after experiment.

Cleanup, replay commands, and exact commit custody are recorded in
[the packet](results/change-0822/README.md).
