# 0720 — current DOCX structural scan qualification

This packet retains an untimed observation of the actual
`DocumentBody::from_xml` boundary. No production optimization is retained and
no performance improvement is claimed. See the [report](../../0720-docx-structural-scan-qualification.md)
and [independent review](design-review.md).

The 19 XML controls and two package corpora produce byte-identical public
results before and after temporary instrumentation. Both trace runs also have
byte-identical stderr. The generated-medium body is 21,517 bytes and each
structural pass visits 1,410 events; NumberedList is 4,563 bytes and each pass
visits 178. Each representative body runs two MCE selections, with inputs
0/200 and 0/5 respectively. This supports qualification of one removed lexical
walk; it does not establish native savings or allocator-failure equivalence.

## Artifacts and boundaries

- `oracle/` extends the retained 0712 public diagnostic; its README explains
  lexical prefix probes, facade scope and non-exhaustive limits coverage.
- `instrument.py` applies exact-string patches guarded by original hashes.
  `trace.fragment` is copied to owned scratch and included as a private module.
  `instrumentation.json` binds original/transformed files and the fragment.
- `baseline*`, `trace*` and their receipts bind full Rust source censuses,
  binary identities, public outputs, commands and raw stderr. `CASE0720` labels
  generator setup separately. `TRACE0720` records begin/end pairs even when
  parsing returns early. `analysis.json` binds metadata and per-call vectors.
- Alt event counts increment after successful reads; range event counts after
  initial depth/node/classification guards. Refusal counters therefore differ
  in admission point. Successful complete passes include EOF in both counters.
  Consumed positions are source traversal counters, not bytes read from disk.
  A `completed` flag covers the wrapper including MCE selection; its stage and
  raw-range fields distinguish structural EOF from a later selection refusal.
  Alt metadata is retained and hashed as debug strings, not structurally parsed
  by the analyzer; this is not a complete equivalence proof.
- `initial-trace-build/` preserves a failed build caused by the tracer's SHA-256
  display helper. Its stale baseline binary hash is not a successful trace
  binary. The corrected trace build and captures are separate.
- `constraints.json`, `motivation.json` and `environment.json` bind accepted
  constraints, prior recommendations, generator source, fixture and toolchain.
- `quality.json` and logs record scoped formatting/Clippy plus six repository
  evidence gates. No library-wide test or native performance run is claimed.
- `cleanup.json` witnesses removal of owned scratch and the build target;
  `artifact-manifest.json` seals all packet paths and hashes.

## Replay and fresh capture

For retained evidence, from repository root:

```sh
python3 -B docs/performance/results/change-0720/analyze.py
python3 -B docs/performance/results/change-0720/artifact-seal.py --check
```

For a fresh capture, first archive the prior run outputs or use an isolated
checkout containing the recipe and oracle but no existing receipt/log/output
files. The runner refuses receipt reuse. Then, serially:

```sh
python3 -B docs/performance/results/change-0720/run.py baseline-build
python3 -B docs/performance/results/change-0720/run.py baseline
python3 -B docs/performance/results/change-0720/instrument.py apply
# Retain scratch manifest.json as instrumentation.json before restoration.
python3 -B docs/performance/results/change-0720/run.py trace-build
python3 -B docs/performance/results/change-0720/run.py trace
python3 -B docs/performance/results/change-0720/run.py trace-repeat
python3 -B docs/performance/results/change-0720/instrument.py restore
python3 -B docs/performance/results/change-0720/analyze.py --write
python3 -B docs/performance/results/change-0720/quality.py
```

The recipe intentionally binds this checkout path and these source hashes.
Adapt the scratch path in a separate checkout before capturing and retain that
recipe. `restore` rejects unknown source changes rather than overwriting them.
The analyzer also validates the archived initial failure in this retained
packet; fresh captures may keep that historical subdirectory as provenance.
