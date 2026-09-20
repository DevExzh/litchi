# 0721 — rejected DOCX structural scan fusion pilot

The candidate is rejected and production is restored exactly. All 16 read-control
gates and one primary edit-mean gate fail. No production optimization is retained.
See the [report](../../0721-docx-structural-scan-fusion-pilot.md),
[every paired measurement](measurements.md), and [independent review](review.md).

The candidate explores a shared traversal for the DOCX alternate-content and
active document-block range work. iWork is outside this packet. The semantic
oracle and trace lane are diagnostic checks; they do not establish a native
speedup, allocator-failure equivalence, hardware result, RSS result,
cold-cache result, throughput result, or scaling result.

The frozen timing design has two primary corpora (`generated-medium` and
`NumberedList.docx`), two phases (`edit` and `lifecycle`), and eight
ABBA stages. It retains 48 primary children: 32 native and 16 allocator.
Each native stage also has two independent read-consumer controls, for 16
control children. Every top-level child is a fresh process with 200 samples
and 100 warmups in the native lanes; the allocator lane has three samples and
no warmups. The pinned filesystem control has its own documented isolated
internal child schedule; those internal processes belong to that control's
scope and are not pooled with primary timings. Each of its eight invocations
launches 100 untimed warmup-priming children, then 200 priming children and
200 timed warm children, for 500 inner children and 200 timed queries per
invocation. Across the eight invocations that is 4,000 inner children and
1,600 timed queries. The other read control lists paragraphs from the
generated-medium semantic corpus. The allocator lane is
instrumentation-bearing and is not latency-comparable to the native lane.
Edit, lifecycle, publication, and allocation observations are separate
measurements and are never pooled.

The primary decision requires exact normalized output parity, at least three
percent edit improvement in p50 and mean for every native pair and corpus,
at most three percent lifecycle regression, and at most three percent
allocator regression in request counts and requested bytes for the first two
pairs. The read controls require their own parity and non-regression gates.
Tail changes over five percent and repeat drift over five percent remain
visible review flags. A final candidate disposition is valid only when all
required gates pass. A rejected candidate must leave an explicit baseline
disposition and `source-final.json` bound to the restored baseline.

The packet also retains a 19-case public oracle and two package corpora, plus
the repeated TRACE0721 diagnostic. The trace proves the observed relationship
between the two reader owners and the fused observer on the retained cases; it
does not turn reader-event counts into a timing claim. Source maps, build
records, oracle probes, trace patch manifests, script hashes, fixture hashes,
child receipts, stdout/stderr, and raw reports remain part of custody.

The packet retains all 20 successful orchestrator commands, 48 primary
invocations and 16 read-control invocations. Four archived quality preparation
failures and one read-driver validation failure remain as provenance. The
validation failure occurred before read samples; the four completed A1 primary
captures were retained when the driver resumed. No samples were rerun.

The candidate's five implementation files and differential test file remain
under `candidate/`; original implementation files are under `baseline/`.
`disposition.json` and `source-final.json` prove baseline restoration. DOCX
formatting, tests, Clippy, doctests and rustdoc pass on the candidate, as do
535 benchmark harness tests (one ignored). Final-source repository evidence
gates, nine corruption checks and exact post-cleanup replays complete validation.

Replay the terminal audit after cleanup:

```sh
python3 -B docs/performance/results/change-0721/audit.py
```

The terminal audit replays `analyze-final.py --check` and
`read-controls-analyze.py analyze` in a temporary directory. `pilot.py` and
`read-controls.py` remain the frozen capture scripts; the two analysis outputs
bind those original capture hashes and their replacement analyzer hashes.
The analyzer replays are Python custody checks, and the retained correction
diffs are applied only to disposable copies. The audit does not build Cargo
targets or execute a native benchmark binary. It accepts a live binary or one exact
cleanup witness containing its absolute path, SHA-256, and byte count, so the
reports remain replayable after owned binaries are gone.

`artifact-manifest.json` seals every packet file except itself. Verify it with:

```sh
python3 -B docs/performance/results/change-0721/artifact-seal.py --check
python3 -B docs/performance/results/change-0721/memory-diagnostics.py --check
```

For a fresh capture, use an isolated checkout at the recorded baseline revision
and archive prior output files. Install the verified baseline snapshots, run
`build.py baseline` and `oracle.py baseline`, then install all six candidate
snapshots and run the candidate quality/build/oracle steps. `trace-run.py
baseline` and `trace-run.py candidate` restore the exact candidate after their
diagnostic captures. Freeze the recipes, plans, binaries and sources before
`capture.py`; it runs both native/read cycles and then the allocator cycle.
`analyze-run.py initial` computes decisions, after which the explicit source
disposition must be applied before `analyze-run.py final`. Captures refuse
existing child reports. These recipes bind the recorded absolute paths; adapt
paths only in a separate copy and retain the amended recipes and hashes.

The original capture-time analyzers remain frozen. `analyze-final.py` and
`read-controls-analyze.py` are the corrected analysis-only entry points; their
exact diffs and hashes are retained in `analyzer-corrections.json`. They preserve
the capture-script receipt bindings and the original gate formulas. Cleanup
removes only this batch's four owned roots and retains eight binary hash/size
witnesses. The overall performance program remains active.

Verbatim logs, unified-diff context lines and the frozen trace recipe retain
their captured whitespace. The staged whitespace check excludes those evidence
artifacts; editable documentation and remaining scripts pass it.
