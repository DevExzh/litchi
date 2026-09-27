# 0788 cached-Part memory attribution evidence

This packet diagnoses the whole-child RSS increase that rejected change 0787.
It does not reconsider adoption. The exact historical candidate is rebuilt
only for measurement and the complete production baseline is restored afterward.

`plan.json` fixes four distinct lanes before capture: 120 native children,
20 qualification children, 64 diagnostic children with phase handshakes enabled
or disabled, and 16 heaptrack children. These total 220 reports and 5,092
measured samples. The enabled diagnostic lane retains 432 externally read
phase snapshots. All children are serial; source-observer, phase, and profiler
measurements are never pooled with ordinary native measurements.

`candidate-binding.json` binds the previous candidate archive, patch, and
quality evidence. Each build's `source.json` records every production file,
the unchanged standalone native tool, and the archive-only probe. Those source
maps must equal the corresponding historical 0787 maps. `quality.json` records
fresh checks of the new probe; it does not claim a fresh rerun of the historical
production test suite.

Root-only execution order is `quality.py`, `build.py before`, qualification
before, `candidate.py apply`, `build.py after`, qualification after, the native,
memory and heaptrack capture modes, then `candidate.py restore`. Drivers refuse
to overwrite capture output. Do not run native children concurrently with
builds, profilers, or CPU-heavy offline analysis.

`source-review.md` explains lifecycle and instrumentation boundaries.
`references.json` records primary Linux accounting documentation. Whole-child
high-water RSS, point-in-time residency, and intercepted allocation peaks have
different scopes. Anonymous mappings are not automatically classified as heap
or worker stacks. Two diagnostic repeats do not establish a confidence interval
or a general causal conclusion.

Offline replay from the relocated repository:

```sh
python3 -B docs/performance/results/change-0788/analyze.py --check
python3 -B docs/performance/results/change-0788/memory_analysis.py --check
python3 -B docs/performance/results/change-0788/heap_analysis.py --check
python3 -B docs/performance/results/change-0788/validate.py --require-final-seal
```

The final validator includes the separate memory/heap replays, live production
restoration, executable cleanup witnesses, and the complete payload seal.
`capture-correction.json` preserves the baseline qualification sidecar-label
correction without rewriting any raw evidence. `heap-histograms/` derives
allocation-size counts offline from the original compressed traces.

See [the change report](../../0788-cached-part-memory-attribution.md) for scoped
results, counter disagreement, limitations, and the next measurement requirement.

To recapture rather than replay, create a clean checkout of `origin.json.base`
at the recorded owned-worktree path and a fresh owned target. Copy only this
packet's runner/planning inputs, `probe-src/`, and `workspace-Cargo.lock` into
an empty change-0788 packet in that checkout; do not copy output lane/build/
quality directories. Copy the retained workspace lock to the checkout root and
recreate the three recorded third-party links if running the boundary gate.
The archived `Cargo.toml.template` substitutes `@SRC@` with that checkout path.
The exact historical candidate is already present in change-0787 at the base
revision. Preserve this published packet, freeze any host/path changes as a
new experiment, and follow the root-only execution order above. Reusing the
same source does not imply reproducing the observed timing or RSS values.

[Independent raw review](results-review.md) checks the payload oracle, full
native paired matrix, phase identity/counters, allocation scope, correction,
and restoration. Its result is diagnostic-only and leaves the 0787 rejection
authoritative.
