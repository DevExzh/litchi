# Retained evidence — change 0676

Record: [`../../0676-execution-budget-composition.md`](../../0676-execution-budget-composition.md).
Authority: accepted ADR 0031 as selected by change [0652](../../0652-owner-decisions-for-the-third-wave.md),
with the landed write-side prerequisite from [0662](../../0662-parallel-changed-member-deflate.md).

This packet records the read-session budget-composition implementation. It is
correctness and resource-admission evidence only. No latency, throughput,
scaling or benchmark claim is registered.

## Provenance

| | |
| --- | --- |
| base commit | `d2d7e2dcc` (landed 0662) |
| branch | `perf/0676-execution-budget-composition` |
| worktree | `/home/zhuhe/code/litchi-worktrees/0676` |
| build concurrency | `CARGO_BUILD_JOBS=2` |
| target directory | `/tmp/litchi-0676-debug0` |
| scheduler | private pools are lazy; caller facilities are explicit; no global pool |
| performance claim | none |

## Contents

| path | what it records |
| --- | --- |
| [`decision.json`](decision.json) | scoped decision, evidence and limitations |
| [`gates.txt`](gates.txt) | targeted checks and test results |
| [`log-sections.md`](log-sections.md) | four paragraphs for coordinator rollups |

## Scope

The implementation covers the three retained read sessions named by ADR 0031:
ZIP `OpenSession`/`ParallelReadSession`, CFB `SharedOleBulkRead`, and
source-backed ordered Part reads. `Workers`, `IoConcurrency` and `CpuTasks`
compose through the hierarchical budget. Private pools retain worker permits
for their pool lifetime; caller facilities use operation-scoped worker
reservations. Every operation passes its admitted width into task submission,
including waves on a cached private pool. Serial paths still admit one worker
and one positional-read permit. Deterministic request/output preflight and
worker/I/O refusal happen before the cumulative `CpuTasks` charge; CFB charges
per admitted batch. The existing constructors retain their no-op, unbounded
behavior.

The final four-crate debug0 run passed 1,877 tests with two pre-existing
ignored tests and no failures; its raw output is retained at
`/tmp/litchi-0676-debug0-tests-final4.log` during review. The tests prove
deterministic order, caller-facility use, shared-root narrowing, measured
cached-pool read width, zero-I/O refusal before payload reads, and release after
serial failure. They do not claim performance movement or complete the separate
ordinary-CRUD API surface.
