# 0729 — public DOC attribution evidence

This unchanged-source packet follows sealed 0728 at commit `949a4a3037`. Both
DOC fixtures, exact edited outputs and preservation witnesses remain bound to
that baseline. Four routes in one performance-diagnostics-enabled binary
separate ordinary execution, outer clocks, profiled implementation, and semantic
timestamp callbacks. Existing profiled APIs have a documented strict-editor
lifetime difference; the experiment does not assume they have identical cost.

`plan.json` contains the exact prospective order of 72 processes across three
cycles and three rounds: two cases, four routes, three warmups and fifty measured
lifecycles each. `freeze.json` binds the source, fixtures, ancestry, final probe,
qualification, executable and analysis code before capture. The manifest retains
all commands, raw stdout/stderr, times, exit codes and start/end bindings.

`analysis.json` reports all per-process distributions, same-owner outer and
semantic fractions, and 54 matched route comparisons with 108 central metrics.
`analysis.md` and `report-summary.json` present process ranges and control flags.
`audit.py` independently checks custody, exact event sequences, span arithmetic,
nonnegative outer residuals, positive denominators, output semantics and all
statistics. `negative-checks.json` records fifteen actual-analyzer corruption
controls. No samples are removed or adaptively repeated.

Final qualification uses `build-1.json`, `quality-0/manifest.json`, and
`qualification-0/manifest.json`: formatting, seven unit tests, warning-denied
Clippy and rustdoc, plus eight route smoke processes. Initial failed build
source/logs are retained. Synthetic preflight records are temporary schema
integration data and are never published as measurements. Source and probe
review is in [source-review.md](source-review.md) and
[probe-review.md](probe-review.md); the latter retains the corrected review
finding about signed residuals.

Replay after cleanup from the repository root:

```sh
python3 -B docs/performance/results/change-0729/source-guard.py
python3 -B docs/performance/results/change-0729/analyze.py
python3 -B docs/performance/results/change-0729/report-tables.py
python3 -B docs/performance/results/change-0729/audit.py
python3 -B docs/performance/results/change-0729/negative-checks.py
python3 -B docs/performance/results/change-0729/artifact-seal.py --check
```

`cleanup.json` binds the executable before removing the two owned target/binary
roots. Eight terminal checks then reproduce the results exactly without the
binary. The artifact manifest covers every packet file except itself. Rebuilds
and new measurements belong in a new packet, not over this sealed evidence.

Whole timing includes local owner destruction and output extraction. Returned
output survives for untimed oracles. The recorder uses a fixed-capacity stack;
trace formatting and validation occur outside whole time. Recorder calibration
is never subtracted from measured work. Outer windows are sequential and bounded
by whole; inner fractions use their own sample's parent window. Semantic phases
do not imply nested CFB phases. No new allocator, RSS, cold-cache, I/O, concurrency,
Office, cross-platform, or production speedup claim follows.

[Performance record and next step](../../0729-doc-public-phase-attribution.md).
The broader non-iWork goal remains active, with no registered claim or CRUD
coverage promotion.

[Result review](result-review.md) recommends a bounded, private DOC handoff
prototype with direct ordinary-route retention and memory gates. All eight
post-cleanup terminal checks pass with exact result hashes; production remains
unchanged and the 17 larger-fixture control flags remain part of the evidence.
