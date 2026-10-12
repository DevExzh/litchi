# 0221 — PPT planning over the retained native scene

The first direct PPT shape setter parsed a slide into a temporary cache, then
started a native run without transferring that cache. The second setter parsed
the same slide again before retaining it. The initial plan now transfers its
native scene and ordinal indexes into the run; subsequent setters reuse them.
Repeated-target and structural barriers still replan against the new snapshot.

Text-only planning uses the cached native shape directly. Text boxes and auto
shapes expose borrowed text equal to their native aggregate projection, avoiding
an existing-body clone just to determine whether the value changed. Placeholder
and other shape text follows the existing native aggregate method. Geometry
planning retains its checked node/anchor path. No detached format model is added.

Every PPT handle constructor obtains the presentation from its package. For a
retained owner, both share the same source, so the facade source fence now checks
the authoritative package once. A released handle still checks its retained
presentation. Entry and exit fences, native publication checks and strict stale
source/error policies remain in place.

## Verification

The new source-probe regression fails on the prior implementation (two probes)
and passes with one probe, including stale-source rejection for retained and
released presentations. The new allocation regression measures only admission
of the second setter, after the first setter starts a run. Previously it used
19,333 requested bytes for a short existing body and 248,657 for a 16-KiB body.
After cache transfer and borrowed comparison, both use equal allocations and
bytes. Publication and fixture/open work are excluded from this regression.

The final all-feature facade suite passes 1,605 tests/doctests. Fifty-two
PPT-only tests pass, covering the source probe, read/write/structure, coalescing,
capabilities and allocation regressions. All-target warning-denied Clippy,
warning-denied rustdoc, the 66-package crate-boundary check, OfficeCLI capability
delta check and its 15 Python tests pass. Both repository lockfiles are
unchanged; Cargo used external copied locks (root validation uses cached
cc 1.6.0, harness retains its exact lock contents). No public API shape changed.

## Paired measurements

Runtime source: `6587364862`. [Samples, build identity and exact proofs](results/0221-ppt-planning.json)
cover all 28 existing aggregate-v2 routes. Every source hash, output hash,
operation count, inner-chunk count, commit count and publication-history proof
matches record 0220. The native publication API and benchmark calls are unchanged.

Each pair uses one warmup and five samples, native then unified. Runtime
sources were committed, with only documentation uncommitted. No tests or builds
overlapped timing; release LTO is disabled. Opening, fixture and command-vector
preparation, serialization and verification are outside timing. String-path
parsing, planning, native publication and patch destruction are included.
The default 256-MiB snapshot profile is used through 4096 operations; 10k retains
the explicit finite 1-GiB profile.

| Operations | Eager native ms | Eager unified ms | Ratio | Positional native ms | Positional unified ms | Ratio |
| ---: | ---: | ---: | ---: | ---: | ---: | ---: |
| 1 | 0.109 | 0.118 | 1.076 | 0.114 | 0.131 | 1.153 |
| 10 | 0.149 | 0.174 | 1.168 | 0.150 | 0.182 | 1.214 |
| 100 | 0.538 | 0.720 | 1.337 | 0.557 | 0.806 | 1.446 |
| 256 | 1.221 | 1.691 | 1.385 | 1.215 | 1.858 | 1.529 |
| 1,000 | 4.311 | 6.018 | 1.396 | 4.326 | 7.004 | 1.619 |
| 4,096 | 17.743 | 25.369 | 1.430 | 17.432 | 28.560 | 1.638 |
| 10,000 | 127.618 | 190.147 | 1.490 | 122.252 | 187.553 | 1.534 |

The default-profile **1k batch is within the 2× budget for both owners**:
1.396 eager and 1.619 positional, versus 1.749 and 2.080 in record 0220.
Unified 1k improves 1.26× eager / 1.31× positional; 10k improves 1.20× /
1.26×. Single-operation eager overhead is 7.6%; positional is 15.3%, so the
positional 10% single-write trigger remains open.

Linear scaling across native rollover remains open. Native 1k→10k growth is
29.60× eager / 28.26× positional; unified is 31.60× / 26.78×. The native
4096-operation bound and three publications at 10k are unchanged. Passing the
relative 1k budget does not establish the full batch-scaling requirement.

These are local synthetic same-slide text closure edits on warm owners, with
blocked leg order. Cold opens, command JSON conversion, geometry, broad corpus
and complete allocation qualification are not established by this matrix.

This turn created no temporary worktrees. Owned target directories, copied
locks, regression logs, probe binary and raw measurement copies were removed
after evidence reduction. No compiled artifacts are committed.


## Remaining work

The full unified-operations goal remains active. Native rollover/publication
cost, broader CRUD and format coverage, Python/CLI completion, and complete
performance/corpus qualification remain required.
