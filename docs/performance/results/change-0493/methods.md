# Managed DOCX read-ahead methods

This batch follows the request-amplification measurements in change 0492.
The candidate is now owned by `litchi-opc`, with an explicit policy forwarded
by the DOCX source-backed expert API. Existing constructors select exact reads.
This is source/I/O workstream A of `docs/GOAL.md`; it does not close the broader
CRUD, borrowed-lifetime, native-producer, or worker-scaling requirements.

## Comparison

Both arms open a fresh managed DOCX package, load the main document, extract
text, and drop the document and package. Both use the same finite execution
context, payload-cache limits, source bytes, transport, and executable. The
control selects exact reads. The candidate selects a 4,096-byte forward-start
window. Separate normal and instrumented allocator executables prevent the
allocator observer from being confused with ordinary performance.

The operation clock and allocator region include production window construction
and destruction. This differs from the 0492 benchmark-private adapter, whose
window was allocated before timing. Corpus generation, source and transport
construction, and execution-context construction are outside the clock. The
returned text survives the clock; hashing, semantic comparison, source-version
comparison, trace extraction, and post-drop budget checks occur afterward.
Diagnostic snapshots inside the clock are part of the measured operation.
The physical trace buffer is preallocated outside the clock, so eliminating
provider calls does not produce artificial savings from observer-buffer growth.

The pinned 0188 corpus has 200 paragraphs, eight 2 MiB media payloads, 20 ZIP
members, and 16,793,036 archive bytes. Its SHA-256 is
`a4a2e4921235a6da6b38e31d26ddcca1301909885e37330ab4f83ecc0c4e04f4`.
The expected 10,000-byte text has SHA-256
`ad4fe690f0ef2281ad8e64a78d1f4d64e7c8625d672b3ac2f7e9fbc28a82f4af`.
This synthetic corpus is not native-producer evidence.

The provider is a caller-supplied in-memory transport model. Arms use either
zero service delay or 1 ms fixed delay plus 104,857,600 bytes/second under the
minimum-service policy. The maximum transport range is 65,536 bytes. Provider
calls are physical only relative to this adapter stack; they are not disk
reads or network packets. No ambient network client is added.

## Accounting and interpretation

The package's `requested_bytes` counts logical forward requests, including
hits. Its `returned_bytes` counts accepted bytes from successful physical
fills. A counter below the transport records actual provider ranges separately.
Cache hits need not consume any new input bytes. Overfetch consumes managed
`InputBytes`; the allocated window retains a `Memory` reservation. Publication
through the package permanently disables read-ahead and releases its
window before exact ZIP traversal. This includes source-XML capture,
materialization, splice preparation, inverse publication, and artifact restore. A source artifact retains the source and
its context, but does not own the archive read-ahead allocation.

Error-path diagnostics are not a replacement for resource accounting: a
physical read can charge accepted bytes before a later cancellation/source
fence refuses the operation. Successful benchmark samples require agreement
between physical returned bytes and managed input charges, unchanged source
identity, exact text, and released package memory after drop.

Median and nearest-rank p95/p99 are reported per role, transport, and repeat.
With 30 retained samples, p99 is the observed maximum; it is not a precise
population-tail estimate. Bootstrap intervals describe each sampled median.
GNU time provides one whole-child peak RSS observation per process,
not an operation-level memory distribution. Allocation counts, reallocation
counts, allocated bytes, and incremental peak live bytes remain separate.
Adverse changes above 5% require explicit reporting, including local-source
overhead. Do not combine unlike transport arms into an overall speedup.

The machine is shared. CPU affinity and an advisory capture lock serialize
this lane's work; they do not reserve the CPU against unrelated jobs. No
multiworker or cold-filesystem claim follows from this comparison. Request
elimination explains one serial I/O component; it does not establish parallel
scaling or a program-wide order-of-magnitude gain.

## Architecture constraints

The ADR inventory in `adr-refresh.json` matches the accepted records already
read for this performance program. ADRs 0001/0002/0010/0011/0024 keep physical
ZIP ownership inside OPC and expose only a bounded policy through the expert
DOCX surface. ADR 0005 requires explicit opt-in, bounded window memory,
hierarchical physical-input charging, cancellation, and measured evidence.
ADRs 0003/0006 preserve source fences, exact no-ops, typed refusals, untouched
physical members, and sequential-output behavior. ADR 0008 requires runnable
validation and honest scope. No archive implementation dependency or runtime
handle is introduced into ordinary CRUD APIs.

## Concurrent publication

Forward reads share one window and one admission lease. Provider `read_at` and
`version` callbacks execute without an adapter mutex held. A publication request
permanently closes forward admission, drains an existing forward lease, and
releases the buffer and its memory reservation. Readers arriving during that
drain use exact ranges. Same-thread callback reentry into an admitted forward
operation or publication transition receives a typed refusal before another
provider callback. The state is bounded per package; it has no thread-local
registry or hidden executor. Managed readers queued behind a forward fill
check their context before waiting and poll at 10 ms intervals, matching the
existing payload-cache waiter. Unmanaged waiters have no cancellation token.
Panic cleanup restores fill ownership and permits
window release. Synchronous providers must return for an outstanding fill to
drain; this adapter does not impose a preemptive deadline on caller code.

## ADR compliance matrix

| Constraint | Implementation and evidence |
| --- | --- |
| 0001, 0002, 0010, 0011, 0024: public layers and physical ownership | Private OPC adapter; DOCX forwards validated policy and diagnostics. Final boundary check and warning-denied facade/substrate builds. |
| 0005: explicit bounded resources | Window is opt-in, at most 64 KiB, allocated after Memory admission; accepted physical bytes consume InputBytes. Unit tests exercise exhausted/shrinking budgets, cancellation, short reads, EOF, and overflow. |
| 0003: immutable source identity and atomic publication | Snapshot lineage/version checks surround forward use; publication permanently closes admission. Integration tests cover stale-source refusal and exact physical output. |
| 0006: preservation and fail-closed behavior | Exact traversal before materialization, source-XML/auth capture, overlays, topology, splice, inverse, and restore. Existing and added preservation tests exercise stored/Deflate unknown members and no-op restoration. |
| 0005: caller-controlled execution | No runtime, global worker pool, filesystem, or network provider added. Per-package synchronization has bounded state and deterministic gated race/callback tests. |
| 0008: verification and support scope | Final-source gate receipts, separate normal/allocator builds, raw pilot/formal observations, strict oracles, independent reviews, cleanup authentication, and a sealed inventory. Native and multiworker scope remain unproven. |
