# 0512 current XLSX attribution evidence

The [change record](../../changes/0512-xlsx-commit-attribution.md) states the
observations, instrumentation caveats, unchanged-source scope and next action.

From the repository root, replay retained evidence with:

```sh
python3 -B docs/performance/results/change-0512/verify.py
python3 -B docs/performance/results/change-0512/analyze.py
```

`verify.py` validates source/build/fixture custody, report/catalog bindings,
raw durations, sink identities, native repeat drift, exact commit profile
boundaries and whole-child hardware quality. Its output matches `summary.json`
and `replay-after-cleanup.json`. `analyze.py` reproduces the disjoint commit
children and overlapping diagnostic contexts in `attribution-summary.json`.
`SHA256SUMS` binds every retained artifact except itself; run `sha256sum -c
SHA256SUMS` from this directory for the byte inventory.

The capture order, corpus selection, sample counts and limits were frozen in
`plan.json`; exact invocations and artifact hashes are in the per-lane receipts.
`run.py` and `capture.py` use exclusive file creation to prevent replacement of
retained measurements. A fresh recapture needs an empty evidence directory,
owned scratch path and the recorded base source/toolchain; record its fresh
source and executable hashes rather than overwriting this batch. Executables
and target files were removed after capture, while source manifests, build
logs, raw profiles, raw sample vectors and replay tools remain.

No production/harness code changes or speedup claim are made. The broad goal,
operation-memory/save-phase measurements, cold/provider coverage and scaling
remain open. OLE2/OOXML retain priority; ODF is deferred and iWork excluded.
