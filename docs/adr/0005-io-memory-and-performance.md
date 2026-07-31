# ADR 0005: I/O, memory, and measured performance

- Status: Accepted
- Date: 2026-07-31

## Input and lazy state

The foundational input contract is immutable positional `ReadAt`, not shared
`Read + Seek`. Paths, moved byte owners, mmap-like owners, remote range sources,
and borrowed byte scopes adapt to it without exposing source generics on a
document. A source has stable snapshot identity/version; mutation during a read
returns `SourceChanged`.

Opening performs container, relationship/catalog, security, and mandatory
structural validation. Semantic payloads load lazily into thread-safe weighted
caches. Clean parsed values are evictable; active handles pin them; dirty edit
state is never silently evicted. Cache behavior is semantically invisible.

Every operation charges a hierarchical resource budget supplied by an execution
context. Production-safe desktop, server, and trusted-batch profiles are finite.
Callers may raise specific configurable limits but cannot bypass integer,
nesting, decompression, or structural safety ceilings. Limit errors identify the
resource, observed value, limit, and object path.

Scratch storage is an explicit capability. Litchi never spills decrypted or
sensitive content to plaintext temporary files automatically. Supported scratch
providers include memory, encrypted temporary storage, and caller-defined
stores; absence yields a typed resource error.

## Output

Ordinary save creates a fresh artifact. Filesystem replacement uses a sibling
temporary artifact, validation/finalization, flush/fsync as supported, and atomic
replacement. Cancellation leaves the destination untouched and removes the
temporary artifact. Caller-owned non-atomic sinks report incomplete output and
bytes written.

Every finalized document supports a sequential non-seekable sink by planning
sizes and layout first or using explicit scratch storage. Preserve-mode save
raw-copies unchanged compressed ZIP entries or CFB streams when possible.

Random-access `Edit` and forward-only `stream::Writer` are separate APIs.
Streaming writers consume and release flushed rows, paragraphs, slides, or parts
and make revisiting them impossible through ownership. Existing-document append
is a restricted tail-only transaction that still writes a new artifact.

Async APIs exist only at genuine suspension boundaries. In-memory CRUD and pure
calculation stay synchronous. Core crates have no Tokio dependency or boxed
future hot path; runtime adapters are optional. CPU parallelism is opt-in through
an execution context controlling scheduling, affinity, cancellation, thread and
memory budgets. There is no hidden global Rayon pool.

## Measurement contract

Representative small, large, sparse, media-heavy, encrypted, and malformed
corpora gate open latency, lazy lookup, concurrent reads, disjoint writes,
patching, and save. Track peak resident memory, allocations, copied bytes,
decompression, cache misses, lock/contention time, CPU utilization, and scaling.
Optimization decisions require profiles, flame graphs, and statistical evidence;
intuition alone is not accepted.

Instrumentation is an opt-in runtime-neutral observer with optional tracing and
profiling adapters. It never records document content, credentials, or sensitive
paths by default.
