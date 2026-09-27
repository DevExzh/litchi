# 0789 known-page RSS accounting evidence

[Interpretation and limitations](../../0789-rss-accounting-calibration.md).
The rejected 0787 candidate remains rejected; this packet makes no library
change or adoption decision.

Offline replay from the repository root:

```sh
python3 -B docs/performance/results/change-0789/analyze.py --check
python3 -B docs/performance/results/change-0789/accounting_summary.py --check
python3 -B docs/performance/results/change-0789/validate.py --require-final-seal
```

`validate.py` checks 9,196 production files against the capture baseline. Run
it from this batch's source revision. Add `--check-workspace` only in the
original owner's workspace to also verify preserved unrelated files and the
pre-existing worktree inventory. The ordinary replay does not require those
untracked user files. Binary custody uses exact cleanup witnesses after the
temporary target is removed. No replay command executes a native probe.

Capture commands, already executed exactly once on the retained matrix:

```sh
python3 -B docs/performance/results/change-0789/build.py
python3 -B docs/performance/results/change-0789/capture.py matrix
python3 -B docs/performance/results/change-0789/capture.py traces
python3 -B docs/performance/results/change-0789/capture-detail.py
```

These are provenance commands, not instructions to overwrite this packet.
Drivers intentionally refuse an existing target or output lane. A new
experiment needs a fresh packet/target and its own frozen source, plan, host,
and binary receipts. Root ran builds and native children serially; offline
agents did not run workloads.

- `plan.json`, `frozen-inputs.json`, `build.json`, `host.json`: original protocol,
  source/tool identity, strict builds, and normal/sanitized binary identities.
- `probe.c`, `probe-review.md`: bounded anonymous mapping and worker protocol.
- `qualification/`: six refusal/ACK checks and three sanitized successful runs.
- `matrix/`: 272 child reports and raw six-phase observations, with the full
  and identity observer lanes kept distinct.
- `traces/`: four original syscall-path traces; rusage fields are abbreviated.
- `trace-detail-plan.json`, `capture-detail.py`, `trace-detail/`: four separate
  verbose traces exposing exact self/wait4 resource counters. The capture code
  is the frozen original with `strace -v` and a separate supplement entrypoint.
- `analyze.py`, `analysis.json`: complete original 276-child replay, 840 full
  and 816 identity snapshots, exact matrix/protocol/checksum/source custody.
- `accounting_summary.py`, `accounting-summary.json`: independent raw-counter
  summary across all 280 children, 864 full and 816 identity snapshots, and
  exact verbose-trace checks. No data are pooled into a performance estimate.
- `source-accounting-review.md`, `sources.json`: versioned primary-source and
  installed-header evidence, including vendor-source limitations.
- `final-review.md`: independent source/protocol/numerical review.
- `replay-corrections.json`: analysis and trace-format corrections; raw captures
  and the original frozen experiment remain unchanged.
- `cleanup.json`, `seal.json`: verified removal of owned temporary executables
  and final artifact/document hashes.

Known-page checks do not observe every transient peak. GNU time, self usage,
wrapper usage, phase residency, allocation peaks, and launcher history remain
different measurements. Two repeats support descriptive observations only.
