# 0091 — Reuse the DOCX body handle during write admission

DOCX target admission now loads one body handle for the resolved target set.
Previously each selected node reacquired the same memoized native body index.
Every node still proves its address and native transaction position, and every
source-freshness fence remains. The cache exists only inside one admission call
and cannot survive an edit or be reused by the next operation. Structural
operations using individual write selections retain their original behavior.
There is no public API change. The full goal remains active.

## Correctness and bounded work

The new eager/source regression resolves paragraph and run fanout, then counts
body loads during admission only. Both target kinds require exactly one load
for multiple nodes. All 692 DOCX tests/doctests passed, covering nested/mixed
selection refusals, ordered edits, transaction poisoning, patch inversion,
command text normalization and formatting barriers. All-target Clippy passed
with warnings denied. Reproduce with:

```sh
CARGO_NET_OFFLINE=true cargo test -p litchi-ops --no-default-features --features docx
CARGO_NET_OFFLINE=true cargo clippy -p litchi-ops --no-default-features --features docx --all-targets -- -D warnings
```

## Paired performance

The frozen 0089 harness ran nine native/unified ABBA rounds per invocation.
An isolating control patch disables only admission's shared body handle. The
outer invocation order was control/candidate/candidate/control on CPU 0, with
no concurrent builds or validation. All byte-equivalence, operation-count,
frozen-corpus and 10k typed-refusal checks passed for both binaries.

The ranges below span the two per-invocation upper-median unified/native ratios:

| Targets | Eager control | Eager candidate | Source control | Source candidate |
|---:|---:|---:|---:|---:|
| 1 | 1.417–1.418 | 1.422–1.424 | 1.248–1.260 | 1.250–1.258 |
| 10 | 1.297–1.325 | 1.204–1.230 | 1.234–1.255 | 1.151–1.155 |
| 100 | 1.278–1.281 | 1.128–1.135 | 1.219–1.226 | 1.098–1.107 |
| 1,000 | 1.277–1.279 | 1.126–1.126 | 1.219–1.230 | 1.104–1.106 |
| 10,000 | Refused | Refused | Refused | Refused |

This reduces measured fanout overhead in this tiny generated memory-owner
workload. It does not fix the single-paragraph review trigger, establish
allocation/RSS acceptance or prove command-batch performance. A native
composite replacement limit still refuses 10,000 targets. Other documents,
filesystem owners, remaining text policy, formats and Python remain open.

Samples, summaries and source/compiler/binary provenance are in
[unified-ops-0091](../performance/results/unified-ops-0091/). The
[harness README](../../tools/perf-baseline/unified-docx-text-fanout/README.md)
describes timing boundaries and control reproduction. The patch is retained;
binaries are excluded, with their size, SHA-256 and build command recorded.
The candidate source was restored before final checks.

Scoped formatting, diff hygiene, control-patch applicability and rewritten-base
ancestry checks passed. Both required staged evidence checks printed nothing;
an audit of 396 changed/untracked paths found no prohibited artifacts. Staging
failed at the read-only linked Git metadata path
`/home/zhuhe/code/litchi/.git/worktrees/unified-ops-api/index.lock`, so no commit
was created. All checkpoint processes finished before cleanup. Temporary
binaries, logs, source backup, build output and the generated untracked root
lockfile were removed; intentional harness lockfiles and reduced evidence
remain. No temporary worktree was created.
