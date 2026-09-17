# 0676: compose execution budgets across the existing read sessions

Change 0676 completes the read-side portion of the accepted execution-context
decision from [0652](0652-owner-decisions-for-the-third-wave.md), on top of
landed change 0662. The shared `Budget` now has a ninth dimension,
`Resource::IoConcurrency`, alongside `Workers` and `CpuTasks`. The additive
`Limits::with_execution_io` builder sets all three dimensions while the
existing six-argument constructor and two-argument `with_execution` remain
source-compatible and leave positional-read admission unbounded.

The explicit ZIP `ParallelReadSession` now builds its private pool lazily and
can run borrowed tasks on the caller's `ScopedWorkers` facility. The OPC
`OpenSession` adapter admits a shared worker width and positional-read width
before the first payload read, holds private-pool worker permits for the pool's
lifetime, and uses operation-scoped permits for a caller facility. CFB bulk
reads use the same distinction: a private pool retains its worker reservation,
while a caller facility receives an operation-scoped reservation. Both paths
consume one `CpuTasks` unit per requested stream/member.

Source-backed ordered Part reads add `Workers` and `IoConcurrency` to their
existing scheduler admission, narrow deterministically when the shared root
has fewer permits, consume `CpuTasks` once per prepared request, and route
borrowed tasks through the caller facility when present. The private scoped
thread path remains available otherwise. Every serial path reserves one worker
and one positional-read permit before payload I/O; a zero `IoConcurrency`
budget refuses before the first source read. Reservations release on operation
completion, cancellation and failure, with private-pool worker reservations
released when their owning session is dropped.

The focused tests cover ordered output, caller-facility routing, lazy pool
construction, shared-root worker narrowing, zero-I/O refusal before source
reads, CPU-task charging, and release on serial failure. Targeted format, test,
lint and documentation gates pass. This record makes no latency or
throughput claim and retains no benchmark table; the change establishes
correctness and resource-composition behavior only. Ordinary CRUD APIs remain
unchanged and no global executor is introduced.

## Scope

The managed and explicit low-level paths are covered: ZIP bulk reads through
`OpenSession`, CFB `SharedOleBulkRead`, and
`SourceBackedPackage::read_parts_ordered`. Existing callers that construct
`Limits` without opting into finite execution dimensions retain unbounded
`Workers`, `IoConcurrency`, and `CpuTasks` semantics. The ordinary package
constructors and CRUD routes remain serial and outside this opt-in boundary.

## Verification boundary

The tests prove admission order, deterministic output and typed refusal. They
do not establish a speedup, a cross-platform scheduler bound, or a process-wide
thread-count measurement. Those claims remain intentionally absent from this
record.
