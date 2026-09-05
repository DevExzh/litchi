# 0423: matched owned and source-backed PPTX lifecycles

This bundle adds a current, same-revision baseline for plain and media-rich
cross-presentation slide copying through two public APIs. The source-backed
selectors now include opening, planning and publication in one operation, with
independent media, XML, relationship and raw ZIP preservation checks. The
original phase-only selector retains its boundary.

The implementation revision is `6ca9962c7818e173538a40ed45f3a3c32cc1aa6a`. The
[frozen protocol](protocol.json) defines 16 fresh processes: two API roles,
two corpora, two repeats, and separate normal (100 samples / 10 warmups) and
allocator (30 / 3) lanes. Each lane/corpus runs owned R1, source R1, source R2,
owned R2, serialized on CPU 2 with one worker. Functional 1 / 0 checks are
excluded. No heaptrack or CPU sampling profile is added here.

The [result table](result-table.md) and [machine-readable summary](summary.json)
retain per-role distributions, repeat drift, logical source/sink observations,
V3 allocator vectors and whole-process RSS. Normal repeat thresholds are 5%
for p50/mean, 10% for p95 and 15% for p99. Unstable rows remain descriptive and
withhold acceptance-grade statistics. There is no source-backed speedup,
memory-reduction, prior-version delta or zero-copy claim. The APIs retain
different artifacts at the operation boundary, and source read-adapter overhead
is included. See [design](design.md), [resource scope](resource-review.md),
[validation](validation.md) and [next work](next-work.md).

All 16 captures and full replay passed. Owned media normal repeats exceeded
the p50/mean drift limits (+7.083% / +7.034%), so those statistics remain
descriptive. The other three corpus/API pairs passed all declared drift limits.
The eight R1 mutation suites rejected 88 probes; portable replay covers all
16 original reports and repeats those eight suites.

Replay the retained bundle with Python 3 from any location:

```sh
python3 -B docs/performance/results/change-0423/portable-replay.py
python3 -B docs/performance/results/change-0423/check-portable-tool-binding.py
```

Replay exports only this bundle and its four hash-pinned shared validators.
It requires neither the original worktree nor the copied binaries. The
binding check also changes a copied validator and requires rejection. Reports,
journals and catalogs retain their original bytes. Compressed logs are
lossless; `compression.json` records original/stored sizes and SHA-256 values.
`SHA256SUMS` inventories the final bundle. Replay updates its own check receipt;
run on a copy if preserving the directory's byte-for-byte inventory is desired.

To recapture, create a fresh directory with the scripts, protocol and pinned
tools, then build the exact source revision and run the matrix. Commands below
assume the repository is the working directory:

```sh
repo_root=$(pwd)
bundle="$repo_root/docs/performance/results/change-0423"
batch_root=/tmp/litchi-0423-recapture
source_root=/tmp/litchi-0423-recapture-source
mkdir "$batch_root"
cp "$bundle"/*.py "$bundle/protocol.json" "$batch_root/"
cp -a "$bundle/replay-tools" "$batch_root/"
git worktree add --detach "$source_root" 6ca9962c7818e173538a40ed45f3a3c32cc1aa6a
python3 -B "$batch_root/build.py" candidate "$source_root" \
  --root "$batch_root" --binary-prefix /tmp/litchi-0423-recapture \
  --target-dir "$repo_root/tools/perf-baseline/target"
python3 -B "$batch_root/capture.py" --root "$batch_root" --repo-root "$source_root"
python3 -B "$batch_root/summarize.py" --root "$batch_root" --repo-root "$source_root"
```

The build/capture drivers refuse an existing result or run directory. The
source checkout must remain clean and both binary identities must stay fixed.
Use the retained validation driver for the focused tests; all CPU workloads
must stay serialized. This batch is a measurement enabler with explicit
follow-up work; the broader non-iWork goal remains incomplete.
