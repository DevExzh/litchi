# 0814 — native attribution after direct event-result matching

Completed diagnostic: the scanner remains the leading sampled leaf, with
most scanner self samples at two post-dispatch payload moves. Production is
unchanged; no new speedup or adoption is claimed.

The baseline is `5eb8629254`, which retains 0813's direct event-result dispatch.
All 9,196 production file hashes match sealed 0813 after; all 35 previously
read normative inputs remain unchanged. The prior batch improved large-capture
p50 by 10.544% and large lifecycle by 7.269% on its own frozen workflow trial.
Those historical timings are not pooled into this diagnostic.

The question is what native work remains after the three pre-dispatch event
copies were removed. Ordinary, profiling-wrapper, and frame-pointer-wrapper
builds use the same exact probe and production source. Six counterbalanced
blocks cross tiny/medium/large capture, with thirty samples and three warmups
per process. Two subsequent large-capture perf processes each measure 100
samples with cycles:u at 499 Hz, CPU 12, and frame-pointer stacks. This yields
56 reports and 1,820 verified outputs. Wrapper/control and fp/wrapper ratios
measure instrumentation perturbation, not source optimization benefits.

Exact binaries remain available through decode and bounded scanner/inspector
disassembly for all three variants. Raw perf data and non-inline frames are
retained as deterministic gzip members. Independent readers check the sealed
source/output/semantic oracles, six-block bootstrap statistics (seed 814814,
10,000 draws, sorted endpoints 250/9749), owner-qualified stacks, unresolved
frames, lost events, overlapping inclusive counts, and tail/RSS diagnostics.

Six production quality gates are reused by exact identity to sealed 0813
(1,241 passed, zero failed, three ignored); probe formatting, all 36 tests,
and warning-denied Clippy run freshly. The diagnostic has no allocation lane,
Callgrind recapture, cross-format or concurrency inference, or adoption gate.
Guest instruction counts, sampled cycles, and instrumented timings cannot be
used interchangeably. iWork is excluded; the broad goal remains active.

[Protocol](results/change-0814/protocol-review.md) and
[packet](results/change-0814/README.md) describe the frozen execution sequence.

All 56 reports / 1,820 measured outputs pass the sealed source, output, and
full semantic oracles. Main and independent raw readers agree exactly on
paired p50 statistics and native frame counts.

| Shape | Ordinary p50 ms | Wrapper/control paired ratio (95% interval) | FP/wrapper paired ratio (95% interval) |
| --- | ---: | --- | --- |
| tiny | 0.228732 | 1.002120 [1.000741, 1.007084] | 1.005104 [0.997263, 1.007530] |
| medium | 0.426422 | 1.012834 [1.006596, 1.016642] | 0.999433 [0.997355, 1.010088] |
| large | 16.0685455 | 1.040429 [1.034036, 1.044374] | 1.006262 [0.999645, 1.013582] |

Ordinary values are medians of six process p50s; ratios are medians of paired
block ratios. The wrapper adds 4.043% to large-capture p50. The additional
frame-pointer ratio is 1.006262 with an interval crossing one. Both are build
perturbation observations, not optimization benefits or evidence that the
instrumented build has the same phase costs as ordinary execution. The AMD
EPYC 9R45 host, CPU 12, Rust/Cargo 1.95.0, release optimization 3, thin LTO,
and one codegen unit are recorded in the packet.

Six process-spread flags exceed 5%: tiny control p99 11.755% and RSS 7.667%;
tiny wrapper p95 6.361%, p99 7.687%, and RSS 11.400%; medium wrapper RSS
7.482%. All mean, p95, p99, RSS, and paired distributions remain retained.
No p50 spread exceeds 5%. Ordinary median RSS is 5,074 / 5,254 / 18,710 KiB
for tiny / medium / large. This packet makes no tail or memory benefit claim.

| Large-capture repeat | Whole-process samples | Exact-owner samples | Unknown interior frames | Lost-event lines |
| --- | ---: | ---: | ---: | ---: |
| 0 | 2867 | 851 | 2 | 0 |
| 1 | 2851 | 853 | 1 | 0 |

Scanner self leaves account for 150/142 samples, hardware SHA for 91/84,
Reader event handling for 69/67, and element inspection for 39/59. The scanner
appears inclusively in 729/716 owner stacks and the inspector in 239/284;
these nested counts overlap. UTF-8, attribute iteration/duplicate checking,
memchr, and memory comparison remain visible. The unknown interiors are
retained; no sample is discarded to improve apparent coverage.

The exact fp scanner is 2,246 bytes; ordinary and wrapper scanners are each
2,320 bytes. Inspectors are 3,798 bytes with fp and 3,860 in the other builds.
All six symbol ranges are bound to the actual binary hashes. The old
pre-dispatch aggregate chain is absent. Inside the Start and Empty arms,
`movups 0x18(%rax),%xmm1` occurs at fp offsets `0x25b` and `0x33b`. They receive
66/62 and 59/52 self samples respectively, or 125/150 and 114/142 of scanner
self samples. Both arm-local 32-byte payload moves also exist in ordinary
and wrapper assembly. These are distinct from the three pre-dispatch
40-byte copies removed in 0813.

The offset reader joins every scanner/inspector leaf to an instruction boundary
in the exact fp assembly. Other owner leaves are counted separately (662/652),
so the selected-symbol and other-leaf totals reconstruct 851/853 owner samples.
Sampling skid, the small sample, and measured wrapper perturbation preclude
assigning causal cycle savings to either instruction.

The assembly stage initially refused an ambiguous `inspect_element` suffix:
OPC and PPTX each export a matching symbol. The original driver and partial
control scanner artifacts are preserved. A separate recovery driver narrows
the selectors to fully qualified PPTX mangled prefixes and records all six
ranges from the unchanged live binaries. No build, workload, perf capture,
or decode is rerun. The aggregate checks the exact recovery delta and retained
partial hashes along with the frozen original driver.

A bounded follow-up is to test borrowing Start/Empty payloads directly in the
existing success patterns, preserving event bodies, resolver timing, typed
errors, limits, and buffered oracles. This remains a hypothesis: it needs
ordinary-build assembly confirmation and a fresh public workflow trial.
Namespace/attribute fusion remains deferred for ordering risk. The broader
native-producer, cross-format, cold-cache, range-source, non-seek output, and
scaling evidence gaps remain open; this scanner diagnostic does not close them.

[Independent results review](results/change-0814/results-review.md),
[next-step source review](results/change-0814/next-step-review.md),
[numerical replay](results/change-0814/analysis.json), and
[exact-binary offsets](results/change-0814/offset-audit.json) delimit the evidence.

Owned-target cleanup verifies all three executable identities and removes
3,557 files / 2,074,159,275 logical bytes. Final offline validation
passes with the retained binary witnesses. Production and the three unrelated
workspace files remain unchanged.
