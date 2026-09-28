# 0822 independent results review

Status: **bounded capture and offline evidence accepted; final replay marker and
root-owned cleanup/seal remain**. This review covers the retained 0822 JSON
receipts, native reports, raw and decoded profile artifacts, `root-audit.json`,
`frame-audit.json`, `analysis.json`, and the current main report. It used
read-only inspection and arithmetic only. It ran no Cargo command, workload,
release binary, profiler, decoder, or packet reader, and it made no production
or probe-source change.

The source custody is the unchanged base
`353aa00a7da2795e6b4c28708138a103a143917e`. The packet uses the canonical
`test-data/ooxml/pptx/shapes.pptx` input (68,822 bytes,
`19fde9b87e33dd1a95fdbba0cf6abc2278bf03874f4665c7f8b88b6afe4a2571`) and the
sealed 0821 default output (68,284 bytes,
`38c7fd3cc5037316e0a3a9c8cb6e6b013a95dd03b39bbd7fb357fc3cc3b9e5cf`). The
timed operation is the public transaction through
`apply_opened_presentation_commit`; `to_bytes`, output hashing, semantic
reopening, and package destruction are outside the timer. Complete byte
equality and independently reopened semantic checks bind every sample to the
0821 output. This is transitive publication parity for an edit profile; it is
not new save, rename, sync, or fsync evidence.

**Capture accounting.** The three qualification reports contain nine samples,
the 18 native reports contain 540 samples, and the two profile reports contain
4,000 separately classified samples. The packet therefore retains **23 reports
and 4,549 samples**; the qualification plus native lanes are **21 reports and
549 samples**. All report families reached their expected counts, and the
retained verification fields require the exact output bytes, size, digest,
complete text, six-slide count, and target shape marker. `analysis.json` is
`status: accepted` with every verification flag true, including the exact
owner-DSO filter, raw/non-inline frame retention, zero-owner guard, and the
omission of shipping-latency, historical-timing, and Amdahl claims.

The retained identity hashes used for this cross-check are:

| Evidence | SHA-256 |
| --- | --- |
| `analysis.json` | `e6dad13caeafd6ee58ec095049173c3490618e9b62fa6accf56eff3591219d94` |
| `root-audit.json` | `5e101b847bc8bc752a0b50ce59e1eb9b8bf28daed0f22a3a30e943bec750ca1f` |
| `frame-audit.json` | `67efb11ef7d128e738acaadd8e5e5cce8fd6a87ff403df46db10532b908c775c` |
| `perf/frame-owner-counts.json` | `3a4f905bb13e7b70874c3713dea43d3627a0b37c6e394ef64d7e880410716dee` |

**Native controls.** The independent root audit recomputes the six-block,
nearest-rank per-process statistics and the matched-block bootstrap. Values
below are medians over the six blocks; time is in milliseconds and RSS is the
whole-process maximum resident set size.

| Arm | p50 | p95 | p99 | Mean | RSS (KiB) |
| --- | ---: | ---: | ---: | ---: | ---: |
| control / ordinary direct | 1.428437 | 1.438872 | 1.452577 | 1.428771 | 5,122 |
| wrapped / ordinary wrapper | 1.428187 | 1.441948 | 1.448183 | 1.429399 | 4,940 |
| fp / frame-pointer wrapper | 1.443077 | 1.452523 | 1.459002 | 1.443994 | 5,142 |

The paired `wrapped/control` p50 ratio is **1.000137**, with a 10,000-resample
95% interval of **[0.997252, 1.002430]**. The `fp/wrapped` diagnostic ratio is
**1.009094**, with interval **[1.005815, 1.012671]**. The first interval
includes one and the second describes frame-pointer build perturbation; neither
is a production speedup. No latency spread or p99/p50 tail flag exceeds five
percent. The only spread flags are whole-process RSS for wrapped (5.62%) and
fp (5.26%); these are not edit-region allocation or retained-memory evidence.

**Profile census.** Both profile repeats use `cycles:u` at 997 Hz with the
exact `pptx_edit_profile_0822::edit_region_0822` owner in the frame-pointer
binary. The independent frame audit reports:

| Repeat | Whole samples | Exact owner | Outside owner | Owner sample share | Empty stacks | Unknown leaves | Unknown interiors, excluding leaf | Wrapper self leaves |
| ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: |
| 0 | 6,164 | 2,971 | 3,193 | 48.199% | 0 | 2 | 2 | 0 |
| 1 | 6,160 | 2,996 | 3,164 | 48.636% | 1 | 3 | 2 | 0 |

Owner-qualified periods are 12,918,080,937 of 26,748,595,054 in repeat 0
(48.294%) and 13,027,190,866 of 26,732,076,750 in repeat 1 (48.732%). These
are separate whole-process attribution denominators, not an edit-phase
fraction. Setup, serialization, readback, verification, and unwinding limits
remain outside the owner-qualified partition. No owner symbol was found in a
different DSO, no owner stack repeated the owner, and every qualified stack
had a descendant.

The decoder retains unknown-frame diagnostics separately: 84/73 unknown frame
records affect 32/35 whole-process samples, while the owner-qualified counts
above identify only the qualified unknown leaves and interiors. There are no
lost-event lines, status lines, or malformed-frame lines in either retained
recording/decode path. Repeat 1's single empty stack is retained in the
whole-process and outside-owner denominator. The first frame-audit attempt is
also retained; it failed closed by asserting that every stack was non-empty.
The corrected audit explicitly accounts for the empty stack and passes without
changing a capture. The original failure is visible in
`frame-audit-attempt-0.py` and `frame-audit-run.log`; the corrected pass is
recorded in `frame-audit-attempt-1.log`.

The selected owner-qualified descendant leaves agree with the complete
analysis ranking:

| Descendant leaf | Repeat 0 | Repeat 1 |
| --- | ---: | ---: |
| `quick_xml::events::attributes::IterState::next` | 184 | 200 |
| `zlib_rs::inflate::inflate_fast_help` | 172 | 169 |
| `sha2::sha256::x86_sha::compress` | 166 | 173 |
| `quick_xml::name::NamespaceResolver::resolve_event` | 136 | 119 |
| `quick_xml::reader::ns_reader::NsReader<R>::process_event` | 129 | 128 |
| `litchi_pptx::shape::reader::Scene::read_with` | 111 | 108 |

This is a selected table, not the complete ranking. The full retained ranking
includes libc and other parser leaves; for example, repeat 1 has
`__memcmp_evex_movbe` at 137 samples, intentionally omitted from the selected
table. The wrapper itself has zero self-leaf samples, so the main report should
call these rows **owner-qualified descendant leaves**, rather than implying
that they are wrapper self leaves.

**Bounded next source review.** The profiles justify a narrow source-tracing
pass in this order:

1. Trace `package_fingerprint_with_memo` and `capture_internal` under
   `opened_presentation_with_limits`. SHA-256 paths through this chain reach
   the owner at 127 samples in each repeat; representative decompression paths
   reach it at 101/104 and 71/65 samples. Check memo invalidation, package
   limits, authorization, and why this materialization is required by the
   public transaction.
2. Trace `Scene::read_with` and `raw_shape_span` beneath
   `Transaction::set_shape_text` (60/56 and 38/44 samples respectively).
   Verify target selection, namespace handling, raw-span validation, and the
   preservation conditions before considering any alternate lookup.
3. Trace `compaction_scene` and `Transaction::commit`, including namespace
   resolution. The representative commit path has 51/52 samples, and the
   namespace-resolver commit path has 57/43. Verify that compaction and all
   relationship/shape checks remain required for the admitted output.
4. Trace `scan_processed_xml` through `SlidePart::finish_from_processed` and
   `capture_internal` (41/42 samples). Establish whether notes/processed XML
   work is required for this package and public edit contract.

These are investigation priorities, not optimization recommendations. The
next review must preserve the public call sequence, limit and validation
semantics, full output bytes, complete semantic reopening, and the independent
0821 reference. No sampled leaf count supplies a causal cost, a phase
fraction, or permission to omit validation.

**Quality and report review.** The packet's production reuse records the
committed 641 passed, zero failed, one ignored test result across 28 suites;
the fresh packet probe has five passing quality gates. The retained
`reader-check.log` records both independent audits and accepted 23/4,549
analysis counts. Its `final: false` marker is consistent with the current
release state, and the duplicate write refusal in `reader-attempt-0.log` shows
that the sealed `analysis.json` was not overwritten. The main report currently
retains `FINAL_REPLAY_PENDING`; root should remove it only after the terminal
replay/validation gate.

Before that final edit, the main report should make three small wording
corrections. It should use `control` consistently for the ordinary/direct arm,
describe the selected leaf table as descendant leaves, and state that the
serialized byte vector is owned before the package is dropped outside the
timer; package destruction need not wait for the later semantic checks. Its
explicit omission of libc from the selected table and its statement that this
packet supplies no new save/fsync evidence are otherwise correct.

The measurements and parity evidence therefore support this packet as a
bounded diagnostic source-review input. They do not support a production
change or optimization adoption. The remaining release work is the root-owned
terminal replay, marker removal after that replay, cleanup, and final seal.

**Closure addendum.** Root's final post-cleanup validation and staged seal now
pass all 178 paths. The main report's `FINAL_REPLAY_PENDING` marker has been
removed and the wording corrections above have been applied. Root's complete
raw replay also passes, with the successful analysis outputs and measured
artifacts unchanged.

The final audit records a limitation in the reader-alignment history: six
earlier `--write` attempts failed, but their original console logs and
intermediate reader files were not retained. The retrospective
`reader-failure-attestation.json` records only the observed error/order facts;
it invents no times or exit codes. The main report, reader notes, and execution
notes disclose this limitation. It limits reconstruction of those failed
alignment attempts, but does not alter the accepted final analysis or raw
capture evidence.
