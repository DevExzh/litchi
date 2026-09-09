# 0486 DOCX replay caller evidence

See [results-review.md](results-review.md) for findings, limitations, and the
unimplemented candidate. [caller-summary.json](caller-summary.json) contains
period-weighted inclusive attribution for four profiles. Raw perf data and
script exports are retained as verified gzip archives under `callers/baseline1`.

Successful source-stable gates: environment-host2, callers-baseline1,
summarize-callers1, crud-index1, and report-claims1. The failed environment-host1
argument-parsing attempt remains retained. No production sources changed, so no
new Cargo test campaign was run. Python syntax and gate-log hashes were checked;
synthetic parsing checks confirmed period weighting, duplicate-frame handling,
and rejection of malformed records. Existing 0485 correctness evidence remains
the applicable production validation.

`seal_bundle.py` verifies the exact artifact inventory and seven external
references, including the synchronized program documents and retained capture
helpers. Run without creating any new artifacts:

```sh
python3 -B docs/performance/results/change-0486/seal_bundle.py
```

The seal is an integrity check, not an independent performance claim validator.
`summarize_callers.py` performs the archive, report, and build checks; its
successful output and gate are retained. It creates its output exclusively and
does not overwrite the sealed summary. The empty scratch directory was removed;
0484/0485 external executable evidence remains retained.
