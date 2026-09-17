# Log sections for change 0676

Four paragraphs for the coordinator to merge into the shared performance logs.
This packet deliberately does not edit the shared rollups.

## For `HOTSPOTS.md`

## 0676 — read-session budgets now compose across ZIP, CFB and source-backed reads

Record: [0676](../../0676-execution-budget-composition.md), implementing the
remaining read-side boundary accepted by [ADR 0031](../../../adr/0031-execution-context-budgets.md)
on top of [0662](../../0662-parallel-changed-member-deflate.md). The shared
budget adds `IoConcurrency`; ZIP `OpenSession`, CFB `SharedOleBulkRead` and
source-backed ordered Part reads now admit `Workers` and positional-read
permits before payload I/O, charge cumulative `CpuTasks` after deterministic
preflight and worker/I/O admission, narrow deterministically under a shared
root and route borrowed tasks through an explicit caller facility when
supplied. ZIP and CFB private pools remain lazy and retain worker permits for
their lifetime; every task facility submits only the admitted operation width,
including later waves on a cached pool. Focused tests establish byte order,
active read-width bounds, zero-I/O refusal, caller routing and release on
serial failure. No latency or throughput claim is made; ordinary CRUD remains
outside the explicit session scope.
[Record and limitations](0676-execution-budget-composition.md);
[retained evidence](results/change-0676/README.md).

## For `GOAL_AUDIT.md`

## 0676 — explicit read sessions now honor the shared execution budget

Record: [0676](0676-execution-budget-composition.md). The change completes the
accepted execution-context composition boundary without changing ordinary CRUD
entry points. Every managed read that reaches payload I/O admits at least one
`IoConcurrency` permit and refuses with a typed resource error before source
bytes when the dimension is exhausted. Worker width is bounded by the shared
hierarchical root, CPU work is charged after deterministic preflight and
admission but before payload tasks, and caller-owned borrowed-task facilities
replace private pools where attached. Ordered result slots and existing source
fences preserve deterministic bytes and error selection; focused ZIP, CFB and
source-backed tests cover lazy construction, active-width bounds, caller
routing, shared-root narrowing and release paths. No hidden global executor is
introduced, and no performance claim follows from this record.
[Record](0676-execution-budget-composition.md);
[retained evidence](results/change-0676/README.md).

## For `REPORT.md`

## 0676 — execution budgets compose across the three explicit read boundaries

The ZIP, CFB and source-backed read sessions now share `Workers`,
`IoConcurrency` and cumulative `CpuTasks` through the hierarchical budget. A
private pool is built only after a qualifying batch and keeps its admitted
worker width for the session; an attached caller facility runs borrowed tasks
with operation-scoped worker permits, and cached pools submit only the current
operation width. Serial reads still reserve one worker and one positional-read
slot, and a zero positional-read budget refuses before the first payload read.
Tests retain input order, typed failures, cancellation fences, active CFB read
width and release behavior. This is an admission and correctness change for
opt-in low-level sessions; ordinary constructors and CRUD remain serial, and
the record registers no speedup or latency result.
[Record](0676-execution-budget-composition.md);
[retained evidence](results/change-0676/README.md).

## For `ADR_COMPLIANCE.md`

## 0676 — ADR 0031 read-side composition is implemented at all three named sessions

Record: [0676](0676-execution-budget-composition.md), authority accepted ADR
0031 and [0652](../../0652-owner-decisions-for-the-third-wave.md) decision 6.
`litchi-core` owns the runtime-neutral `IoConcurrency` resource and scoped
facility; `soapberry-zip` keeps its standalone mirrored trait; `litchi-opc`
bridges caller facilities privately. ZIP and CFB pools are lazy, source-backed
ordered reads keep their existing scoped-thread path when no facility is
attached, and all three paths reserve before payload reads, preserve typed
refusals and release operation permits. Cached private pools and caller
facilities submit only the admitted operation width. The additive execution
builder leaves old constructors unbounded, and no ordinary CRUD signature
exposes a scheduler. Focused tests prove caller-facility determinism,
shared-root narrowing, active CFB read bounds, zero-I/O admission and serial
release. No performance claim is made; process-wide thread probes and latency
measurements are outside this prerequisite record.
[Record](0676-execution-budget-composition.md);
[retained evidence](results/change-0676/README.md).
