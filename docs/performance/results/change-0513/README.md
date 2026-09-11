# 0513 XLSX operation measurement evidence

The [change record](../../changes/0513-xlsx-operation-allocation.md) describes
this harness-only measured enabler and its native overhead observations.

Replay the retained evidence from the repository root:

```sh
python3 -B docs/performance/results/change-0513/verify.py
```

The verifier checks source/build/fixture custody, report/catalog bindings,
raw durations and negative corrupt/short vectors, corpus/sink identity,
normal unavailable versus allocator measured fields, aligned allocation
vectors, native ABBA deltas/drift/RSS and the exact save-helper profile edges.
Its output is retained in `summary.json` and `replay-after-cleanup.json`.
`SHA256SUMS` binds every retained artifact except itself. From this directory,
run `sha256sum -c SHA256SUMS` to verify the byte inventory.

`plan.json` records the predeclared protocol. `candidate.patch` records the
exact two-file harness change against the base revision. The per-lane
receipts retain exact commands, source/binary hashes, timings and artifact
hashes. Normal captures use 100 samples and three warmups in serial ABBA
order; allocator captures use ten samples and one warmup, candidate only.
The three-sample profile toggles the private save helper after the oracle.
Allocator and Callgrind timing/RSS never enter native comparisons.

`run.py`, `build-allocator.py`, `capture.py`, `check.py` and
`remaining-lanes.py` record the serial execution recipe. Their exclusive
output creation prevents overwriting existing captures. A fresh recapture
needs an empty evidence directory, a distinct owned scratch directory and
reconstructed base/candidate sources. Rebuild and bind new binaries rather
than treating historical executable hashes as reproducible guarantees.
The cleanup receipt records removal of this batch's executables and target
files after all running-process references were checked; raw evidence remains.

There is no format speedup or allocation-reduction claim. This batch
establishes operation allocation baselines and an exact commit/save profiling
boundary for the next measured XLSX production change. OLE2 and OOXML retain
priority until their full optimization goal is complete; ODF remains deferred
and iWork excluded.
