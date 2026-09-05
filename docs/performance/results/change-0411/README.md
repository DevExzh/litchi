# Change 0411 evidence

Single-revision diagnostic evidence for the six existing XLS open/list/one-cell
selectors. No speedup claim. Captured revision:
`44edf790669a0aa4dc0aff73af6f7b5f5e709b6d`.

The source-backed path uses an instrumented in-memory source. The
`InstrumentedSource::read_at` observer dominates the CPU profile; these times are not a plain
`OwnedSource` or filesystem-source proxy. Each normal/allocator invocation is a
fresh process, while warmups and retained samples share that process. Global
filesystem isolation defaults in the raw report do not apply to these cases.

- `capture/protocol.json`: protocol declared before measurement.
- `capture/build-identity.json`: clean source, compiler/flags and binary hashes.
- `capture/commands.json`: exact runtime argv, launcher PID, UTC start and exit.
- `capture/normal-*.json.zst`: four normal reports, 500 samples/case, 20 warmups.
- `capture/allocator-*.json.zst`: two allocation reports, 30 samples/case, three warmups.
- `capture/verification.json`: sample, identity, digest, locality and allocation validation; per-repeat statistics and exact IID median intervals.
- `capture/resources.json`: whole-process RSS, multiplexed PMU events and profile identities.
- `capture/*profile.svg`: whole-process flame graphs for each one-cell family.
- `attribution/`: period-weighted CPU subsets, parser and identity/source bindings.
- `checks/`: passing gates plus the initial diagnostics and rejected corrupt-report checks.

Raw perf data, perf text and larger raw reports are losslessly compressed.
`artifact-manifest.json` binds both published bytes and uncompressed originals.
Numeric summaries, catalogs and command journals remain directly readable.
Allocation calls include reallocations; live/high-water counters are absolute
process snapshots. RSS and PMU counts include setup, clones, validation and
reporting. PMU events are scaled for multiplexing. No isolated operation peak,
exact L1/LLC, physical-I/O, request-size distribution or scaling claim follows.

From a checkout containing this bundle:

```sh
python3 docs/performance/results/change-0411/verify-artifacts.py
python3 docs/performance/results/change-0411/verify-capture.py --root docs/performance/results/change-0411/capture --repo-root . --output /tmp/litchi-goal-0411-reverified.json
python3 docs/performance/results/change-0411/verify-resources.py --root docs/performance/results/change-0411/capture --repo-root . --output /tmp/litchi-goal-0411-resources-reverified.json
```

The capture verifier uses repository `tools/perf_compare.py` and
`tools/validate_perf_corpus_binding.py`. Python and `zstd` are required.
The resource verifier checks two profile reports, six PMU reports and six RSS
captures. The attribution parser additionally verifies event frequency,
call-chain mode, clean revision and source bindings. To replay attribution,
decompress the desired `*profile-script.stdout.zst` and `*profile-self.stdout.zst`
into a temporary directory and use the arguments in `attribution/commands.json`;
adapt the input/output paths and materialize the two profile JSON files with
`build-identity.json` and `commands.json` as described in `attribution/README.txt`.
The retained JSON reports preserve original
capture paths as historical identity fields.

To collect new evidence, use a clean detached checkout of the captured revision
and the exact standalone-workspace build command/environment in
`capture/build-identity.json`. `checks/build.py` and `capture/capture.py` retain
the original scripts and absolute task paths; adapt paths before a rerun. Run
normal, allocator, profile and counter phases sequentially with no competing
build/test/benchmark workload. A new run has a new measurement identity and
must not replace these raw results.
