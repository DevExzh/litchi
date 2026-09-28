# 0829 — current PPTX edit CPU phases

Fresh current-source profiling completes **23 reports and 4,549 verified
samples**. Transaction capture accounts for the largest observed exact-owner
CPU-sample bucket, followed by commit/application and text replacement. This is
diagnostic attribution on one real file. Production is unchanged; no new
optimization, historical speedup or full-save benefit is claimed.

The [0828 symbol-selector failure](0828-pptx-edit-phase-profile.md) remains sealed.
This attempt rebuilds both executables and admits all four wrappers using exact
address-bounded disassembly. Independent preflight checks the live symbol
receipts before qualification. No historical executable or timing is reused.

## Scope and method

The base is `90b9466b99639d411f9b0a89208aad21525e0006`. The input is the
68,822-byte checked-in `test-data/ooxml/pptx/shapes.pptx`, SHA-256
`19fde9b87e33dd1a95fdbba0cf6abc2278bf03874f4665c7f8b88b6afe4a2571`.
Every output must equal the separately admitted 0821 reference: 68,284 bytes,
SHA-256 `38c7fd3cc5037316e0a3a9c8cb6e6b013a95dd03b39bbd7fb357fc3cc3b9e5cf`.
Every warmup and measured iteration verifies full bytes, hash, complete text,
slide count and selected shape text through public reopen/readback.

Each iteration opens a fresh package before the clock. The measured public
sequence constructs the opened-presentation transaction, changes shape `(0, 0)`
to the fixed marker, commits and applies the commit. The returned publication
snapshot drops inside the measured helper. Package-owner drop, serialization,
hashing and semantic readback occur afterward. No snapshot or digest memo is
warmed on the sample owner before its timer.

The AMD EPYC 9R45 host has 32 logical CPUs. Capture processes run serially on
CPU 12. Release builds use opt-level 3, debug level 1, thin LTO, one codegen unit,
two build jobs, unwind panics and no incremental compilation. The second build
adds `-C force-frame-pointers=yes`.

Three qualification reports contain three samples each. Native timing uses six
counterbalanced blocks across ordinary/direct, ordinary/wrapped and
frame-pointer/wrapped arms, with thirty samples and three warmups per report.
Two separate perf captures each execute 2,000 edits without warmup, using
`cycles:u` at 997 Hz and frame-pointer call chains. Qualification, unprofiled
native timings and profiler-observed executions remain separate.

The four named wrappers are whole edit, transaction capture, shape-text
replacement and commit/application. Admission pairs unique raw/demangled
`nm` rows by address, size and text-symbol type, checks positive non-overlapping
64-bit ranges, and verifies each emitted owner heading, instruction bounds,
frame-pointer prologue and call instruction. The independent reader rechecks
this evidence without invoking a binary or disassembler.

## Native instrumentation controls

Values are medians across six process statistics, in milliseconds except RSS.
Within each thirty-sample process, p50/p95/p99 use nearest rank; across-process
medians use the ordinary midpoint. Paired ratios use matching blocks, so they
need not equal ratios of the displayed medians. Confidence intervals use 10,000
bootstrap resamples, seed 829829, and sorted zero-based endpoints 250 and 9,749.

| Arm | p50 ms | p95 ms | p99 ms | Mean ms | RSS KiB |
| --- | ---: | ---: | ---: | ---: | ---: |
| Ordinary / direct | 1.270951 | 1.285112 | 1.297836 | 1.273473 | 4,976 |
| Ordinary / wrapped | 1.271472 | 1.287292 | 1.299532 | 1.274032 | 5,008 |
| Frame pointer / wrapped | 1.302777 | 1.314802 | 1.317762 | 1.304234 | 4,994 |

| Paired p50 diagnostic | Ratio | 95% interval |
| --- | ---: | --- |
| Wrapped / direct | 1.000715 | 0.996008–1.003981 |
| Frame pointer / wrapped | 1.023752 | 1.020221–1.027284 |

The wrapper interval includes 1.0; the frame-pointer build adds approximately
2.375% in this control. These are instrumentation effects, not production
regressions or improvements. Four greater-than-5% spread flags remain visible:
direct p99 (1.084799), and RSS for direct (1.099595), wrapped (1.061489), and
frame-pointer (1.089286). No p50/p95/mean spread or aggregate p99/p50 flag fires.
RSS is whole-process maximum residence, including work outside the edit timer;
no allocation improvement is inferred.

## Sampled phase attribution

A sample qualifies only when its stack contains exactly one whole-edit owner
in the expected executable. Exact nested phase symbols partition qualified
samples into mutually exclusive capture, set-text, publish, unclassified and
ambiguous buckets. Counts and sampled-event periods are independently conserved.
The table reports CPU samples, not operations or milliseconds.

| Population | Repeat 0 | Repeat 1 |
| --- | ---: | ---: |
| Whole process | 5,861 | 5,866 |
| Exact edit owner | 2,670 | 2,685 |
| Outside exact owner | 3,191 | 3,181 |
| Transaction capture | 1,196 | 1,193 |
| Text replacement | 607 | 612 |
| Commit/application | 867 | 879 |
| Unclassified within edit | 0 | 1 |
| Ambiguous within edit | 0 | 0 |

The one unclassified leaf is `Transaction::set_shape_text`; its stack contains
the whole-edit owner but no exact nested phase marker. It remains unclassified.
Owner-qualified sampled periods are 11,608,467,032 and 11,673,631,535; these
are perf sampling weights, not exact hardware-counter totals for the operation.
Whole-process execution also includes opening, saving, hashing, output
verification and startup, so samples outside the owner cannot be assigned to
the edit.

Selected disjoint **self-leaf** groups within each phase are below. Prefix rows
sum only leaf symbols with that literal module prefix. The full leaf census,
including every other symbol, remains in the machine-readable analysis.
Inclusive frames and call paths overlap and are not added to these counts.

| Phase / self-leaf group | Repeat 0 | Repeat 1 |
| --- | ---: | ---: |
| Capture / `zlib_rs::*` | 331 | 355 |
| Capture / `sha2::*` | 133 | 131 |
| Capture / `quick_xml::*` | 201 | 219 |
| Capture / `litchi_pptx::notes::*` | 103 | 101 |
| Set-text / `quick_xml::*` | 303 | 314 |
| Set-text / `Scene::read_with` | 70 | 59 |
| Set-text / `raw_shape_span` | 34 | 31 |
| Publish / `quick_xml::*` | 209 | 216 |
| Publish / `sha2::*` | 29 | 28 |
| Publish / `litchi_pptx::notes::*` | 51 | 54 |

Initial capture contains all observed zlib self leaves within the exact edit
owner. This is consistent with the cold complete-package fingerprint path and
its deferred payload decoding. It is not proof of bytes inflated or of repeated
inflation: package clones share the deferred payload once-cell. Set-text is
concentrated in XML scanning, Scene and raw-span work; publish retains XML and
validation work. None of these counts authorizes removing required validation.

No empty stack, malformed frame, explicit lost-event marker or explicit
truncation marker is observed. Whole-process unknown frames total 86/97 across
43/41 samples; 32/30 stacks end in unknown frames. Qualified owner interiors
have no unknown frames. Maximum observed depths are 28/26. There are no
repeated-owner or wrong-executable owner samples. Missing callers and unmarked
truncation still cannot be ruled out; complete unwinding is not proven.

## Interpretation and next work

The result prioritizes initial capture for further PPTX investigation, while
showing substantial remaining set-text and publication work. Complete source
fingerprinting, payload limits, preservation and publication validation remain
mandatory. The earlier duplicate-catalog hypothesis has only 11/8 capture and
11/10 publish self leaves for `slide_references`; child work and potential
savings are not determined by that leaf count. A direct Scene-span shortcut
also remains unproven because the raw scanner enforces additional refusal rules.

The [source opportunity review](results/change-0829/opportunity-review.md)
retains these candidates and the separate XLSX semantic readback investigation.
A useful next step is to localize the known XLSX allocation amplification in a
public one-cell edit, while keeping any PPTX candidate subject to a fresh matched
workflow trial. This profile adopts no candidate and does not weaken durability.

Observed phase sample counts are neither wall-clock fractions nor an Amdahl
model. Frame-pointer perturbation, one small fixture and one host limit the
inference. There is no fresh allocation, copied-byte, read-byte, inflate-byte,
syscall, cold-cache, concurrency, external-Office or broad-producer evidence.
The archive directory declares 48 members and 154,250 logical bytes; that is
metadata, not a measured decode/copy total.

## Verification and retention

Five fresh probe gates pass: formatting, release check, three tests,
warning-denied Clippy and warning-denied rustdoc. Historical production/tool
quality is reused from sealed 0827 with exact source, lock, normative-input and
committed-blob checks. Its six harness gates and 641 passing tests, zero failures
and one ignored test across 28 summaries are reused evidence, not new tests.
All 9,197 production files, 87 harness files and 35 normative inputs remain
unchanged. The three unrelated workspace files are preserved.

Independent native arithmetic, frame partition and stack diagnostic readers
agree with the primary analysis. Every reader attempt retains its source
snapshot, command, console and exit status. All seven attempts through
pre-cleanup validation passed; no quality, build, symbol, capture or decode
command failed. Raw perf data and decoded frames are retained as lossless,
hash-verified deterministic gzip; redundant uncompressed copies are removed.

Post-cleanup final validation also passes, making eight retained passing reader
attempts. Exact binary verification precedes removal of the owned target:
3,688 files / 2,054,144,787 logical bytes. No filesystem scratch was used.
Final seal verification binds the complete packet and this report.

Offline replay after cleanup:

```sh
python3 -B docs/performance/results/change-0829/validate.py --final
python3 -B docs/performance/results/change-0829/seal.py --check-head
```

Evidence: [analysis](results/change-0829/analysis.json),
[native replay](results/change-0829/root-audit.json),
[independent frame census](results/change-0829/frame-audit.json), and
[stack diagnostics](results/change-0829/stack-diagnostics.json).
