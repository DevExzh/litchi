# 0839 — cached ordered-Part scheduling adopted

[Change report](../../0839-cached-part-scheduling-requalification.md).
Fresh current-source qualification admits the candidate: sixteen eligible cached
cases improve paired p50 by 96.68–98.99%, with zero frozen latency/RSS,
phase-residency or operation-allocation regression flags. The historical 0787
rejection remains unchanged. Five uncertain tail point-estimate flags remain
visible; this is not a whole-Office or universal tail-speedup claim.

## Evidence and replay

- `origin.json`, `host.json`, `plan.json`, `plan-amendment.json`: base, unchanged
  normative/unrelated hashes, environment, immutable cases/guards, and the
  pre-formal-capture clarification of independent metric reducers.
- `candidate.patch`, `source-before`, `source-after`, `candidate-source.json`:
  exact three-file change and source identities. `candidate-review.md` is
  pre-adoption review; its historical caution is superseded only by this fresh
  trial, not by reinterpretation of 0787.
- `probe`: byte-identical 0788 native workload with current relative dependencies.
- `allocator`: separate safe workload and isolated existing System allocator
  wrapper; no timing is measured. Counter source is byte-identical to the
  shared current observer. Unit tests isolate synthetic callbacks from the
  ambient test runner; normal builds install the global wrapper.
- `quality-*`, `allocator-quality-*`, `freeze-*`, `build-*`, `admission.json`:
  final source/build/quality bindings. Passing allocator legs are `before-v2`
  and `after`; the initial `before` test failure is retained with draft source.
- `runs`: 1,414 independent process reports / 28,450 samples. The native,
  source-observer, memory, allocation, qualification and preflight populations
  remain separate. Each process has argv, terminal receipt and hashed artifacts.
  Memory gzip files retain raw acknowledged proc snapshots.
- `traces`, `trace-analysis.json`: six additional untimed one-sample reports and
  successful clone/clone3 counts. Trace durations are not native evidence.
- `readers.py`, `qualify.py`, `reader-tests.json`: independent deterministic
  payload reconstruction, report/resource/source validation, phase identity and
  custody checks, fourteen rejected mutations and four statistical checks.
- `analysis.json`, `paired.md`, `paired.csv`: complete individual results, raw
  block vectors, intervals, flags and frozen decision. `analyze.py` replays the
  decision; `tables.py` renders its tables.
- `selection-review.md`, `review-resolution.json`, `review.md`: target choice,
  resolved pre-capture review concerns, and final review.
- `close.py`, `closure.json`, `cleanup.json`, `seal.json`: command/failure ledger,
  exact final source, marker-checked temporary-root removal and owned-file seal.

From the recorded checkout, retained replay requires no native build:

```sh
python3 -B docs/performance/results/change-0839/analyze.py
python3 -B docs/performance/results/change-0839/close.py verify
```

Custody retains the original absolute command/build paths. `readers.py` also
accepts report paths directly for use in a relocated checkout. A fresh native
experiment must use new output directories and fresh source/build bindings;
`capture.py` deliberately refuses to overwrite existing evidence. Relative
Cargo manifests, locks, deterministic generators, source archives and complete
argv/environment records make those inputs reviewable and reproducible.

The initial allocator test composition and reader-test exception-name failure
are preserved; corrected checks passed before formal capture. Both baseline
and candidate have full fresh OPC/probe gates. Seven retained executables are
hash-verified before cleanup. No ambient library instrumentation, executor,
compression-policy change or unrelated workspace edit belongs to this batch.
