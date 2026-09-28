# 0821 independent real-file durability results review

Status: **fresh capture, admission, offline replay, and independent raw audit
pass; cleanup and seal remain root-owned**. This review is based on the
retained packet descriptors, admission witnesses, completion receipts, direct
JSON inspection, `analysis.json`, `root-audit.json`, and the completed main
report. It ran no Cargo command, release binary, exporter, workload, profiler,
or offline reader.

## Scope and custody

The packet is the 24-row matrix on base
`8312aaa29b59f73d2a7409cb501828989320a8c9`: three checked-in real files
(DOCX, XLSX, and PPTX), `lifecycle` and `atomic_publish`, and the four
policies `default`, `full`, `file-only`, and `no-sync`. The packet and source
reviews establish the paired forward/reverse format/phase order, rotating
policy order, six native process blocks, two observer blocks, and the
nearest-rank/paired-bootstrap definitions. Native and observer measurements
remain separate. Default and explicit full are the durability control pair;
the weaker routes are configuration observations and do not weaken the
full-durability default.

Quality is reused through the exact committed 0820 repair receipts and seal.
The 0821 quality adapter records six reused gates, 641 passed tests, zero
failed, one ignored, and no Cargo command. It also records fresh source,
lock, normative-input, and packet bindings. The fresh release builds,
artifact export, admission, qualification, and captures are separate 0821
evidence; this reuse is not a second quality run.

## Artifact and admission evidence

The independent artifact audit is `ok: true` for six cases (the three
generated controls and the three real inputs). Each case has five untimed
policy outputs. Semantic edits, published bytes, untouched ZIP member bytes,
member order, and archive comments are retained by the audit and preservation
witnesses. The policy output digest vectors are equal within each case, while
the admitted source and output identities remain format-specific.

`artifact-admission.json` is accepted for all 24 planned selectors and binds
the format, phase, input, source digest, published digest, output digest, and
edit outcome before timing. Qualification admission is also accepted for all
24 selectors. The first admission invocation is retained in
`execution-errors.json`: root ran it after the artifact auditor but before
the standalone ZIP-preservation witness, so it exited with
`FileNotFoundError: zip-preservation.json`. That attempt did not publish
admission, start timing, or alter an output or frozen input. The subsequent
preservation witness and `admission-1` pass on the same fresh exports.

## Capture counts and lane separation

| Lane | Reports | Samples | Blocks | Samples/block |
| --- | ---: | ---: | ---: | ---: |
| Qualification | 24 | 24 | 1 | 1 |
| Native | 144 | 4,320 | 6 | 30 |
| Observer | 48 | 144 | 2 | 3 |
| **Total** | **216** | **4,488** | — | — |

All three fresh release builds and the artifact exporter exited zero. The
qualification, native, and observer completion receipts report the expected
child, report, and sample cardinalities; all capture children reached
terminal success before reader authorization. Observer reports retain the
32 empty procfs controls and allocator/process diagnostics. Qualification and
observer evidence are not pooled with native elapsed values, and observer
controls are retained without subtraction.

Direct inspection of the retained report JSON found one result per report,
the expected policy route (`default` omits `save_durability`, while the other
three carry `full`, `file-only`, or `no-sync`), admitted edit outcomes, the
caller-named real input, one source hash per format, and identical published
hash vectors within each format. No report cardinality, policy, source, or
admission mismatch was found in this static pass.

## Derived analysis and raw audit

`analysis.json` and `analysis.md` report `status: accepted` with the expected
216 reports and 4,488 samples. The derived native rows use nearest-rank
p50/p95/p99 inside each 30-sample process block, then the median over six
blocks. The bootstrap uses seed 821821, 10,000 resamples, and sorted ranks 250
and 9749 for both absolute p50 and matched-block policy/default ratios.

The retained `root-audit.json` independently reports pass for all 144 native
reports, 4,320 native samples, 24 rows, and 24 absolute plus 24 paired
intervals. Its reader source reopens every native report, recomputes the
nearest-rank quantiles, block medians, spread ratios, tail flags, and the same
bootstrap endpoints. The resulting rows are:

| Format / phase | Policy | p50 ms | p95 ms | p99 ms | p50 CI95 ms | Ratio/default | Ratio CI95 | Flags |
| --- | --- | ---: | ---: | ---: | --- | ---: | --- | --- |
| DOCX / lifecycle | default | 5.222937 | 5.357742 | 5.427153 | [5.190397, 5.234002] | 1.000000 | [1.000000, 1.000000] | p95,p99 |
| DOCX / lifecycle | full | 5.205017 | 5.348802 | 5.448429 | [5.194142, 5.217192] | 0.998016 | [0.994463, 1.001612] | p99 |
| DOCX / lifecycle | file-only | 3.377918 | 3.498148 | 3.531484 | [3.354512, 3.391062] | 0.647016 | [0.642290, 0.651661] | p99 |
| DOCX / lifecycle | no-sync | 0.242312 | 0.270827 | 0.277566 | [0.237571, 0.244777] | 0.046316 | [0.045694, 0.046926] | p99,tail |
| DOCX / atomic_publish | default | 5.024506 | 5.147036 | 5.195161 | [5.012495, 5.039531] | 1.000000 | [1.000000, 1.000000] | — |
| DOCX / atomic_publish | full | 5.026841 | 5.158131 | 5.221202 | [5.015881, 5.046556] | 1.000865 | [0.998006, 1.003675] | p99 |
| DOCX / atomic_publish | file-only | 3.195921 | 3.307942 | 3.328727 | [3.178031, 3.208451] | 0.635722 | [0.631614, 0.639438] | — |
| DOCX / atomic_publish | no-sync | 0.083996 | 0.102591 | 0.105336 | [0.083165, 0.085165] | 0.016755 | [0.016542, 0.016911] | p95,p99,tail |
| XLSX / lifecycle | default | 5.410838 | 5.547759 | 5.609569 | [5.388018, 5.420233] | 1.000000 | [1.000000, 1.000000] | — |
| XLSX / lifecycle | full | 5.415628 | 5.536084 | 5.598423 | [5.377573, 5.434692] | 1.000870 | [0.993522, 1.007268] | — |
| XLSX / lifecycle | file-only | 3.593988 | 3.804905 | 3.868600 | [3.571693, 3.603903] | 0.663721 | [0.660877, 0.667434] | p99,tail |
| XLSX / lifecycle | no-sync | 0.534127 | 0.602153 | 0.735094 | [0.531463, 0.538373] | 0.099025 | [0.098052, 0.099607] | p95,p99,mean,tail |
| XLSX / atomic_publish | default | 4.948680 | 5.055026 | 5.111111 | [4.932830, 4.961260] | 1.000000 | [1.000000, 1.000000] | p99 |
| XLSX / atomic_publish | full | 4.937820 | 5.040411 | 5.092766 | [4.916925, 4.968411] | 0.998569 | [0.993300, 1.004181] | — |
| XLSX / atomic_publish | file-only | 3.104771 | 3.201396 | 3.223816 | [3.088166, 3.124916] | 0.628414 | [0.624003, 0.630893] | — |
| XLSX / atomic_publish | no-sync | 0.099605 | 0.129920 | 0.154821 | [0.098876, 0.099985] | 0.020091 | [0.019980, 0.020255] | p95,p99,mean,tail |
| PPTX / lifecycle | default | 7.420598 | 7.567738 | 7.588133 | [7.404678, 7.442219] | 1.000000 | [1.000000, 1.000000] | p99 |
| PPTX / lifecycle | full | 7.403842 | 7.541149 | 7.622004 | [7.389048, 7.419233] | 0.998390 | [0.993476, 1.000696] | p95,p99 |
| PPTX / lifecycle | file-only | 5.569784 | 5.648839 | 5.691139 | [5.548264, 5.602194] | 0.750027 | [0.747717, 0.754906] | p95,p99 |
| PPTX / lifecycle | no-sync | 2.078195 | 2.151996 | 2.154001 | [2.070806, 2.113081] | 0.280308 | [0.278623, 0.284738] | p99 |
| PPTX / atomic_publish | default | 5.618328 | 5.756464 | 5.806049 | [5.602344, 5.647374] | 1.000000 | [1.000000, 1.000000] | p99 |
| PPTX / atomic_publish | full | 5.608739 | 5.767435 | 5.796780 | [5.605513, 5.614909] | 0.997913 | [0.993732, 1.001472] | p99 |
| PPTX / atomic_publish | file-only | 3.780639 | 3.868800 | 3.899599 | [3.765384, 3.796880] | 0.673256 | [0.670654, 0.673441] | — |
| PPTX / atomic_publish | no-sync | 0.296617 | 0.351747 | 0.355367 | [0.293047, 0.301437] | 0.052739 | [0.052206, 0.053536] | p95,p99,mean,tail |

The 24 derived p50 values match the main report's six-by-four p50 table
after converting nanoseconds to milliseconds (`atomic_publish` is named
`path save` there). The main report also correctly records 17 rows with 27
spread flags, no p50 spread flag, and six tail flags. All six full/default
intervals include 1.0; the twelve weaker-policy intervals lie below 1.0.
Those ratios describe synchronization-policy configuration on this host and
do not support a default change or a per-syscall cost claim.

The observer records independently reproduce the main report's six resource
rows: allocation calls / allocated bytes / region peak above entry are
DOCX lifecycle `2,617 / 1,258,172 / 634,214`, DOCX atomic publish
`380 / 590,203 / 568,921`, XLSX lifecycle `3,542 / 4,561,313 / 1,103,104`,
XLSX atomic publish `126 / 983,559 / 561,522`, PPTX lifecycle
`11,402 / 5,814,055 / 939,897`, and PPTX atomic publish
`674 / 641,339 / 597,823`. The four policy vectors are equal within each
row. These are observer diagnostics, include observer overhead, and are not
subtracted from native timing.

## Durability interpretation and limits

The exact `save_durability` and `atomic_publication_steps` fields carry the
policy attribution. The generic `timing_scope` text describes the phase
boundary and must not be used to infer synchronization behavior for a weaker
policy. `default` and `full` retain the full temporary-file and parent-
directory synchronization route; `file-only` retains temporary-file sync;
`no-sync` omits both. Policy ratios, once replayed, are matched within the
same block and case and remain configuration attribution rather than
optimization claims.

Each destination starts absent and the filesystem cache is warm. These
captures do not measure physical cold I/O, replacing an existing destination,
permission-copy latency, crash or power-loss recovery, or cross-host/storage
behavior. The three small inputs and one host bound the result. Lifecycle and
atomic-publication values are phase-specific and are not additive or
subtracted to price synchronization. Output equality and ZIP preservation do
not prove crash durability.

## Replay corrections and release gate

The first analysis replay is retained in `reader-attempt-0.log`; it failed
closed on the old quality-reuse seal count/map check. The second replay
accepts all 216 reports and 4,488 samples. The first aggregate validator
attempt is retained in `validator-attempt-0.log`; it expected an extra nesting
for the admission-attempt witness. The bounded validator correction reads the
`admissions` object emitted by the accepted analysis, and the subsequent final
validator passes. Neither correction changes a raw report, frozen driver, or
derived numeric value.

The main report now contains the timing intervals, ratio conclusions,
resource observations, and replay audit. It accurately keeps the warm-cache,
absent-destination, one-host, small-corpus, no-crash, and no-existing-
replacement limits, and preserves the distinction between generic
`timing_scope` text and policy-specific synchronization fields. Its only
remaining marker is `CLEANUP_PENDING`; cleanup and the final 0821 seal remain
root-owned release checks. This review therefore approves the packet's
capture and analysis conclusions for that final cleanup/seal gate.
