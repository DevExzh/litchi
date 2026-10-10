# 0090 — Borrow DOCX immediate write assignments

Immediate DOCX setters now retain borrowed property keys and values during
admission and execution. The input Props already lives through the whole write,
so cloning its strings into static storage was unnecessary. Deferred formatting
plans explicitly convert assignments into owned storage after establishing
formatting eligibility. Text barriers therefore also avoid ownership copies in
the rejected formatting plan. Property order, preflight conversions, source
fences, transaction rollback and error attribution are unchanged.

The lifetime is internal; no new public API or unsafe code is introduced.
This advances allocation discipline without claiming that the entire setter
or property admission is allocation-free: assignment vectors, coercion and
leaf staging still allocate. The full goal remains active.

## Verification

The new regression uses an owned, dynamically constructed key and a large
owned text value. It checks that admitted assignments point at the input
storage, then checks that explicitly owned assignments remain valid after the
input is dropped. The full DOCX suite passed 691 tests/doctests, including
command batches, deferred formatting, fanout, transactions and patch inversion.
All-target Clippy with warnings denied passed.

```sh
CARGO_NET_OFFLINE=true cargo test -p litchi-ops --no-default-features --features docx
CARGO_NET_OFFLINE=true cargo clippy -p litchi-ops --no-default-features --features docx --all-targets -- -D warnings
```

The unchanged frozen 0089 harness was rebuilt and run twice, each with nine
ABBA rounds, after all validation jobs finished. All corpus, byte equivalence,
ledger count and 10k typed-refusal checks passed. Samples, summary and binary/
source provenance are in [unified-ops-0090](../performance/results/unified-ops-0090/).
Reproduction uses the [same harness](../../tools/perf-baseline/unified-docx-text-fanout/README.md)
with `/tmp/litchi-unified-0090-target` as the build directory.

| Targets | Eager unified/native | Source unified/native |
|---:|---:|---:|
| 1 | 1.397–1.427 | 1.259–1.263 |
| 10 | 1.298–1.332 | 1.234–1.247 |
| 100 | 1.259–1.313 | 1.214–1.235 |
| 1,000 | 1.264–1.305 | 1.222–1.233 |
| 10,000 | Typed refusal, not timed | Typed refusal, not timed |

There is no clear timing improvement over 0089, and that earlier checkpoint
is a historical baseline, not an interleaved before/after control. The
single-write review trigger remains exceeded. This record establishes the
ownership-copy reduction through storage identity and lifetime checks;
it does not claim a measured total allocation count or speedup. Further
profiling, representative workloads and the remaining full-goal milestones
are still required.

Scoped formatting, diff hygiene and rewritten-base ancestry checks passed.
Both required staged evidence checks printed nothing; the audit of 387
changed/untracked paths found no prohibited artifacts. Staging remains blocked
by read-only linked Git metadata at
`/home/zhuhe/code/litchi/.git/worktrees/unified-ops-api/index.lock`; no commit
was created. Checkpoint scratch/build output and the generated untracked root
lockfile were cleaned after all jobs completed. Intentional harness lockfiles
and small evidence remain. No temporary worktree was created.
