# 0078 — Paired DOCX batch scaling and target reuse

The preceding checkpoint made verified progress on atomic formatting runs.
This checkpoint adds paired eager/source timing routes under `tools/perf-baseline`
and removes repeated target-resolution work inside a compatible run. The full
unified API goal remains active.

## Change and safety boundary

A formatting run remembers its last exact direct-path spelling and proven
paragraph/run selection. A repeated spelling reuses that selection. These
setters only change property markup, so the positional proof stays valid within
this sealed family. Different spellings resolve normally. Structural, text,
query and metadata barriers discard the run and its cache. A capacity flush
also starts a fresh cache. Property admission still runs for every item; no
assignment or repeated key is merged. Owner/source fences remain on cache hits,
staging, commit and final publication. Cache misses use the existing planning
fence without a redundant outer check.

Read-only review confirmed this proof boundary. Existing byte-oracle, barrier,
inverse, duplicate-key and 4,097-operation tests pass. A new regression changes
the physical source before the second operation on the cached path and verifies
fatal refusal without consuming the third item.

## Measurement method and result

The [isolated harness](../../tools/perf-baseline/unified-docx-batch/README.md)
uses frozen, new route names and a fixed 8,620-byte one-paragraph DOCX. Both
legs start with prepared in-memory owners. The leaf leg uses the primary native
Edit/Commit API; the unified leg uses atomic command batches. Both retain all
native commits until after timing; unified per-item result assembly remains
included as facade overhead. Owned command payloads are prepared before timing.
Open, save and output destruction are excluded. Each invocation first verifies
exact output bytes, operation totals, result counts and every changed flag.

The uncached control disables only the target-lookup shortcut using the retained
patch; it still stores the cache. It isolates the optimization, rather than
reconstructing a historical commit. Exploratory measurements used borrowed
payloads; those are not included in the final evidence. The final harness uses
owned payloads in both versions and makes no full-performance-acceptance claim.

After validation jobs finished, two invocations per version ran on CPU 0 in
control/candidate/candidate/control order. Every invocation measured nine ABBA
rounds per owner/count: 36 samples per leg after pooling the two invocations.
The machine was an AMD EPYC 9R45, Linux x86_64, Rust 1.95.0. Raw arrays, p95/p99,
paired-block bootstrap intervals, source/build identity and binary hashes are
retained in [the small evidence files](../performance/results/unified-ops-0078/).
Confidence bounds describe these runs, not other hosts or workloads.

| Owner | Operations | Control unified/leaf | Candidate unified/leaf | Candidate median |
|---|---:|---:|---:|---:|
| Eager | 1 | 1.309× | 1.296× | 49.09 µs |
| Eager | 10 | 1.472× | 1.129× | 173.95 µs |
| Eager | 100 | 1.581× | 1.136× | 1.484 ms |
| Eager | 1,000 | 1.518× | 1.078× | 14.547 ms |
| Eager | 10,000 | 1.509× | 1.026× | 137.692 ms |
| Source | 1 | 1.117× | 1.124× | 100.97 µs |
| Source | 10 | 1.315× | 1.112× | 231.78 µs |
| Source | 100 | 1.455× | 1.138× | 1.603 ms |
| Source | 1,000 | 1.441× | 1.095× | 15.143 ms |
| Source | 10,000 | 2.809× | 1.025× | 138.765 ms |

The 1,000-operation ratio intervals are 1.075–1.083 eager and 1.091–1.097 source.
Both use one native commit. At 10,000 operations both legs split into three
commits at the native capacity boundary. Candidate elapsed time grows roughly
linearly from 1,000 to 10,000. The source control's 10,000-operation median was
381.539 ms; avoiding repeated resolution reduces it to 138.765 ms. These results
support the goal's 2× batch target for this exact repeated-target scenario.
They do not establish single-set overhead, varied targets, representative large
corpora, filesystem owners, best-effort, other families, allocation/RSS or the
full performance acceptance bar. Those remain required work.

## Validation and reproducibility

The combined unit/batch/transaction run passed 452 tests. After the final fence
simplification and new source-revision regression, all 25 batch tests passed.
DOCX all-target Clippy and harness Clippy passed with warnings denied. The
crate-boundary gate passed for 66 workspace packages; scoped formatting and
`git diff --check` passed. The rewritten base remains an ancestor of HEAD.

Build/run instructions, frozen corpus identity and the reversible control patch
are in the harness README. Regenerate statistics with:

```sh
python3 tools/perf-baseline/unified-docx-batch/summarize.py docs/performance/results/unified-ops-0078
CARGO_NET_OFFLINE=true cargo test -p litchi-ops --no-default-features --features docx --lib --test docx_batches --test docx_transactions
CARGO_NET_OFFLINE=true cargo clippy -p litchi-ops --no-default-features --features docx --all-targets -- -D warnings
```

`provenance.json` records HEAD plus a fingerprint of the dirty Rust/TOML source,
the locked dependency resolution, toolchain, build command, and binary file
names, sizes and SHA-256 values. The control source was restored before final
validation and cleanup. Binaries are not committed.

Both required staged evidence-size/binary-directory checks printed nothing.
The artifact audit covered 350 changed/untracked paths with no prohibited
artifacts. The final candidate source hash matches the recorded binary build.
`git add -A` still fails to create the linked worktree's `index.lock` because
Git metadata is read-only. No commit was created. Changes remain in the checkout;
the full goal remains active.

Cleanup removed both build trees, scratch results and both measurement binaries.
The intentionally added harness lockfile and small JSON evidence remain. No
temporary worktree was created.
