# 0783 — current PPTX lifecycle phase attribution

This diagnostic measures the current production source without changing it.
The inherited 0780 generated corpora, marker, public calls and semantic readback
remain identical. Package ingress, corpus generation, readback and destruction
of retained package/output owners are outside the lifecycle clock.

The `phase-timing` feature adds consecutive timestamps after capture, staging,
commit, apply and serialization. Stack-held timestamps partition the total;
report maps are populated after the clock. These durations include boundary
clock overhead and temporary drops within each statement. In particular apply
includes dropping the returned snapshot. They are descriptive phase timing,
not CPU profiles or attribution of the historical 0780 regression.

The feature-off control keeps the inherited lifecycle clock sequence. Six
alternating blocks pair control and phase builds across tiny/medium/large
corpora. Each process runs three warmups and thirty measured operations pinned
to CPU 12. Native measurements remain separate from allocator instrumentation.
RSS is whole-process peak and includes setup/readback. No physical cold-cache,
range source, scaling, concurrent workload or native Office claim is made.

`plan.json` is written before builds and capture. No process is retried or
excluded. Reports check every output by reopening and comparing exact generated
text, and retain deterministic source/output SHA-256 identities. The unmodified
production source and complete probe/build commands are bound in `build/`.

Build and capture require the recorded source revision and a fresh target and
packet output tree; retained evidence must never be overwritten. Materialize
`probe-src/Cargo.toml.template` for the checkout, keep the retained lock, then
run `python3 -B build.py` and `python3 -B capture.py`. Replay from repository root:

```
python3 -B docs/performance/results/change-0783/validate.py
python3 -B docs/performance/results/change-0783/analyze.py --check
```

No production optimization or before/after speedup is claimed. OLE2/OOXML remain
active, ODF deferred, and iWork excluded.
