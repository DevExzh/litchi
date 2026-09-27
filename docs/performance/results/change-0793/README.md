# 0793 — current PPTX capture allocation attribution

This diagnostic investigates allocation sources after retained change 0792.
There is no production candidate in this packet. The 0785 probe's fixture,
publication/readback oracle, schema, and tool name are retained. An optional
non-inlined wrapper surrounds only the public capture call and keeps its
result visible inside that wrapper to preserve a stack boundary.

Four diagnostic binaries separate direct/wrapped capture from direct/wrapped
allocation instrumentation. Two repetitions over five shapes use three samples
and no warmup, yielding 40 control reports and 120 samples. Ten separate
heaptrack children each capture one sample. Each trace is decoded without
backtrace merging or shortened templates, once whole-process and once filtered
by the capture wrapper, retaining allocation-count stacks and size histograms.
All native work and decoding run serially; CPU 12 is used for child captures.

The frozen plan requires exact output/semantic parity, equivalent operation
counters across wrappers, whole-process count conservation, exact owner-stack
filtering, a nonzero public-capture descendant, and equality between owner
allocation count and successful operation allocations plus reallocations.
This frozen rule fails in all ten traces because allocation_calls already includes
reallocations. The [post-capture explanation](counter-semantics.md) retains a
supplementary comparison without replacing the failed gate. No owner allocation
fractions are authorized.

Elapsed times are diagnostic observations only. No native timing comparison,
CPU phase fraction, optimization result, or memory reduction is claimed.
Heaptrack's intercepted whole-process peak is not an operation peak or RSS.
Nested stack costs overlap and must not be added as disjoint shares.

Build/capture provenance order is `build.py`, `quality.py`, `capture.py controls`,
`capture.py heaptrack`, and `decode.py`. These scripts refuse existing outputs;
reproduction requires a new output directory and owned target at the captured
source with matching lock/tool identities. Never overwrite this packet.
Offline replay is:

```sh
python3 -B docs/performance/results/change-0793/validate.py --require-final-seal --check-workspace
```

Omit the workspace option when unrelated files/worktrees legitimately differ.
Production and sealed documentation must match this packet's commit. Cleanup
records all four executable identities before deleting the owned build target.
The prior production suite is source-bound inherited evidence; only the
unchanged-oracle diagnostic probe gets fresh tests here. No new CRUD coverage,
native-producer, cold/range, concurrency, or broad goal completion follows.
