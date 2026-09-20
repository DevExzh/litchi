# 0719 producer edit/save correctness gate

This packet verifies the strengthened untimed output oracle in the existing
producer-shaped XLSX edit/save selectors. It is harness correctness evidence,
not an optimization or before/after performance comparison. Medium and dense
corpora retain their existing sheet-0 numeric edit targets and identities.

`run.py build` builds the frozen source. `run.py quality` runs standalone
harness formatting, tests, warning-denied Clippy and rustdoc. `run.py capture`
runs two reversed shape repetitions, each with one warmup and three samples.
Those small captures prove executable admission and repeatable output; their
elapsed samples are not a latency baseline or a speedup claim. `run.py evidence`
runs the repository's six evidence gates. Every command has an exit receipt,
source manifest and hashed log; raw corpus/result JSON is retained.

The production crates are unchanged. `baseline-source.json` identifies the
0718 source state; `source.json` identifies the final benchmark source.
`constraints.json` revalidates the goal and all previously read accepted ADRs.
The source timer still covers planning, set/commit and sequential publication;
opening, sink setup, untimed verification and remaining commit/sink destruction
are excluded. Publication consumes the editor and its returned temporary
snapshot is dropped before the timer closes.

See [the report](../../0719-producer-edit-output-oracle.md) and the
[next DOCX experiment](next-experiment.md). That experiment is a reviewed next
step, not an implemented optimization in this batch.

Replay from the 0719 commit with:

```sh
python3 -B docs/performance/results/change-0719/audit.py
python3 -B docs/performance/results/change-0719/artifact-seal.py --check
```

Capture commands intentionally refuse to overwrite retained logs. For a fresh
capture, use the recorded command arguments and separate output paths, or copy
the frozen runner/input scripts to a new evidence directory at the same depth.
The final `build.receipt.json` supplies the exact locked release build command.
The three `initial-*` directories retain superseded builds, the refused
workbook-change assumption, and the passing-tests/failed-lint attempt. None is
pooled into final captures. Cleanup removes the owned Cargo target; its binary
identity remains recorded for post-cleanup replay.
