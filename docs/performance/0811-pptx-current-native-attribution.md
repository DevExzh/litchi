# 0811 — current PPTX capture native attribution

Completed diagnostic; production is unchanged and no new speedup is claimed.

This diagnostic follows retained change 0810 at `2894fbd628`. All 9,196
production file hashes match the tested 0810 after source. The goal, scenario
taxonomy, and accepted ADR identities are unchanged. The exact six-file public
workflow probe and its sealed source/output/semantic oracles are reused; no
historical timings are pooled. Production remains unchanged.

The question is which work remains significant in native capture after direct
event transport removed the scanner's event-wrapper call. Two static reviews
separate guest instruction attribution from native cost. In particular,
Callgrind's software SHA-256 cost cannot stand in for native hardware-SHA cost,
and short root/name projections cannot be described as complete repeated scans.

Three fresh release binaries share production and probe inputs: ordinary
control, a non-inlined capture wrapper, and that wrapper with frame pointers.
Six counterbalanced orders cross tiny, medium, and large capture. Each process
uses CPU 12, thirty samples, and three warmups: 54 reports/1,620 samples.
Profile/control and fp/profile paired ratios quantify build perturbation;
they are not optimization benefits. Quantiles use nearest rank; bootstrap uses
10,000 six-block resamples with seed 811811 and sorted endpoints 250/9749.

Two subsequent native profiles use `cycles:u` at 499 Hz, frame-pointer stacks,
100 large captures each, no warmup, and the exact
`namespace_uri_probe::capture_region_0793` owner. Raw data and non-inline frames
are decoded while exact executables remain, then retained as deterministic gzip
with compressed and decompressed hashes. The complete packet is 56 reports and
1,820 measured outputs. Whole-process samples, owner samples, unresolved frames,
and lost-event diagnostics remain explicit; nested frame counts overlap.

Sealed 0810's six production gates are reused by exact source identity
(1,241 tests passed, zero failed, three ignored). The fresh probe lane runs
formatting, all 36 tests, and warning-denied all-features/all-targets Clippy.
There is no production candidate, allocation lane, Callgrind recapture, or
adoption gate. No phase fraction, native cycle share, tail-latency gain, RSS
gain, or universal performance claim follows from this packet.

[Protocol and receipts](results/change-0811/README.md),
[guest-cost review](results/change-0811/opportunity-review.md), and
[source-path review](results/change-0811/source-opportunities.md).
The broader OLE2/OOXML performance goal remains active; iWork is excluded.

## Results and interpretation

All 56 reports/1,820 measured outputs pass the sealed source, output, and
semantic oracles. The main reader and independent raw audit agree exactly on
paired timing statistics and native frame counts; aggregate custody and
chronology checks pass.

| Shape | Ordinary p50 (ms) | Profile/control paired p50 ratio (95% bootstrap interval) | FP/profile paired p50 ratio (95% bootstrap interval) |
| --- | ---: | ---: | ---: |
| tiny | 0.230741 | 1.005159 [1.001110, 1.006412] | 1.025545 [1.022907, 1.036667] |
| medium | 0.441233 | 1.009636 [1.001000, 1.014502] | 1.048026 [1.037391, 1.059546] |
| large | 17.753220 | 1.002151 [0.999541, 1.009719] | 1.069250 [1.061059, 1.078472] |

Ordinary p50 values are medians of six process p50s. Ratios are medians of
six paired block ratios. Forced frame pointers materially perturb large
capture (+6.925%); sampled rankings cannot establish ordinary-build phase
costs. Tiny fp p99 spread is 36.12%, and RSS spread exceeds 5% for tiny
control (9.23%), tiny fp (6.53%), and medium fp (6.17%). These four flags,
all p95/p99 values, and all resource observations remain in the evidence.
Ordinary median RSS is 5,082/5,292/18,636 KiB for tiny/medium/large.

| Native repeat | Whole-process samples | Exact-owner samples | Unknown interior frames | Lost-event lines |
| --- | ---: | ---: | ---: | ---: |
| 0 | 3,068 | 960 | 4 | 0 |
| 1 | 3,073 | 957 | 0 | 0 |

The scanner is the leading owner self leaf (217/254 samples), followed by
hardware SHA (89/90). Reader event handling (76/38), element inspection
(53/71), UTF-8 validation (52/62), and attribute iteration/duplicate checking
remain visible. Scanner-inclusive counts are 831/835 and overlap with
nested work. Guest software-SHA instruction costs cannot predict native
hardware-SHA savings.

The next investigation should isolate scanner event handling and the
namespace/attribute walks before choosing a candidate. The review suggests
an event-local shared attribute traversal as a bounded hypothesis, but 0806's
rejected iterator experiment makes a fresh end-to-end benefit test essential.
Neither the leading scanner leaf nor those attribute leaves establish causal
savings. Short root/name projections have no demonstrated material opportunity.
Any future change must preserve namespace timing, validation/refusal ordering,
limits, cancellation, semantic proofs, and the buffered differential oracles.

[Detailed results review](results/change-0811/results-review.md) records
instrumentation, tails, RSS, coverage, and the guest/native distinction.
[Machine-readable analysis](results/change-0811/analysis.json) and
[independent audit](results/change-0811/root-audit.json) retain all results.
The first offline reader attempt rejected the probe completion descriptor;
its descriptor selection was repaired to validate both the summary and sibling
completion receipt. No raw receipt, frozen driver, or workload was changed or
rerun. No additional disassembly experiment was added to this diagnostic.

Owned target cleanup verified all three binary hashes and the source map, then
removed 3,557 files / 2,074,170,355 logical bytes.
Compressed evidence and independent replay inputs remain retained.
