# Paired DOCX formatting batches

This isolated harness belongs to `tools/perf-baseline` and builds only the
DOCX dependency closure. Run from the repository root:

```sh
CARGO_NET_OFFLINE=true cargo build --release --locked --manifest-path tools/perf-baseline/unified-docx-batch/Cargo.toml
taskset -c 0 tools/perf-baseline/unified-docx-batch/target/release/unified-docx-batch-baseline 9 > /tmp/paired.json
```

The four frozen route names are `docx_eager_leaf_format_batch_v1`,
`docx_eager_unified_format_batch_v1`, `docx_source_leaf_format_batch_v1`, and
`docx_source_unified_format_batch_v1`. Source owners use in-memory input;
filesystem source-revision checks are outside this scenario. The generated
8,620-byte DOCX has a frozen SHA-256 checked at startup. Changing its contents
or route semantics requires new names.

Each leg starts with a prepared owner. One direct run's bold property alternates
at 1, 10, 100, 1,000 and 10,000 operations. Both legs retain all native commits;
the unified leg also builds its required per-item results, counted as facade
overhead. The native 4,096-operation limit splits the last case into three
commits in both legs. Command construction/cloning, open, save and output
destruction are outside the measured interval. Path parsing, property admission,
staging, result assembly and commit remain inside it. Before timing, the runner
checks exact output bytes, native/unified operation counts, result count and
all per-item change flags. Each scenario warms both legs, then measures nine
ABBA rounds by default (18 samples per leg). Array pairs identify successive
rounds: leaf samples `2r,2r+1` surround unified samples `2r,2r+1`.

`without-target-cache.patch` supplies an uncached lookup control for record 0078.
It disables only reuse of the last proven target; cache storage remains. It is
an isolating control, not a reconstruction of a historical commit. Apply it
only to the matching working source, rebuild, measure, reverse it, and rebuild:

```sh
git apply --check tools/perf-baseline/unified-docx-batch/without-target-cache.patch
git apply tools/perf-baseline/unified-docx-batch/without-target-cache.patch
# Build and measure the control with the same commands above.
git apply -R tools/perf-baseline/unified-docx-batch/without-target-cache.patch
# Rebuild before measuring the candidate.
```

Do not run validation/build jobs concurrently with final timing. Record source
revision and dirty-source fingerprint, lockfile, compiler, machine, binary size
and SHA-256. Retain binaries outside the checkout or remove them after recording
that identity. These timings cover one tiny repeated-target workload. They do
not establish allocation/RSS, other documents/families, filesystem owners,
best-effort, single-set, save, or full-goal performance acceptance.
