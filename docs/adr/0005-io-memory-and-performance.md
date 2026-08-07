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

Generated IWA protobuf bindings are also kept minimal. Prost runtime type-name
metadata remains disabled because no production path consumes it, while the
first production archive-header seam uses exact-version Buffa 0.9.1 lazy views
behind a private codec. This is a staged runtime migration, not permission to
generate the complete schema corpus eagerly for every format.

Physical iWork ingress bounds the two ZIP name spellings independently before
copying either one. Local and central names, extras, and comments are charged
cumulatively; compressed sizes are rejected before payload materialization.
For legacy packages, catalog, component-catalog, and detection paths reject a
nested `Index.zip` from its declared uncompressed size before decompression.
Limit failures retain the resource kind and exact observed and maximum values.
This layer does not yet promise a semantic object path in every physical ZIP
diagnostic.

Buffa lazy decoding of untrusted IWA bytes is always preceded by a
schema-directed common wire-tree preflight. One aggregate policy bounds scanned
bytes, fields, nesting, repeated metadata items, deferred-message occurrences,
and a conservative decoded-memory envelope before a lazy view is constructed.
The adapter then visits every deferred archive-header child exactly once,
checks proto2 required presence, and projects directly into the existing
physical metadata with fallible destination reservations; generated
`to_owned_message` is not a production ingress path. Buffa 0.9.1 still uses
ordinary infallible `Vec` growth for some internal lazy metadata, so the
preflight bounds hostile amplification but does not claim typed recovery from
global allocator exhaustion or a language-level exact resident-memory bound.
Strict contracts requiring either property must use a streaming handwritten
cursor or a corrected Buffa runtime.

The next production Buffa projection is deliberately smaller than a canonical
format schema root. A derived five-file projection reads only repeated field 3
of `TSWP.StorageArchive`, is hard-capped at 32 KiB of generated Rust, disables
unknown retention, and exposes only borrowed text fragments through a private
wrapper. Common-wire preflight bounds root bytes, fields, field type,
fragment count, UTF-8 bytes, and the conservative repeated-view allocation
before stock Buffa 0.9.1 runs. Other length-delimited fields remain opaque by
design, so this projection is suitable only for call sites whose policy needs
semantic text rather than eager validation of every unrelated known child.
The caller-owned source remains authoritative, and no owned Buffa projection
or lazy re-encoding participates in preservation.

Core archive metadata is projected into core-owned `FieldPath`, `FieldInfo`,
and closed-enum wrappers. Optional presence and unknown signed enum values are
retained exactly. Preflight charges both the transient Buffa representation and
the neutral destination vectors, including unknown closed-enum records, before
publication.

Lazy re-encoding is not the preservation boundary. Original source-backed
header bytes and common raw spans remain authoritative for exact no-ops,
unknown fields, duplicate occurrences, and non-canonical encodings. Buffa
encoding is used only for the canonical header created after a semantic change.

Borrowed IWA wire readers use the common source-bound `WireView<'a>` and
`WireFieldView<'a>` when interpreting recognized fields: one borrowed source
and compact spans avoid per-field slice metadata, payload ranges are sliced
through validated spans, and schema-owned key/length framing can be required
without changing the permissive unknown-field parser. Singular wire overlays
index base and overlay field numbers once and emit one exact-capacity output,
so sparse updates do not repeatedly reparse a growing message. Source-built
Pages, Numbers, and Keynote chart updates also locate their single chart
payload with one linear scan and
no temporary index allocation before decoding or invoking a mutation callback.
These are allocation-shape and safety improvements; representative allocation,
latency, and throughput measurements remain governed by the measurement
contract above and are not claimed by this slice.

Reference-line graph updates likewise avoid a full generated-Prost round-trip:
bounded raw fields are merged by repeated-field occurrence and only recognized
values are replaced, so unknown graph bytes are copied once at their original
nesting positions. The candidate field collection is validated before
publication. This reduces avoidable graph allocations while remaining a
structural optimization rather than a measured throughput claim.

Instrumentation is an opt-in runtime-neutral observer with optional tracing and
profiling adapters. It never records document content, credentials, or sensitive
paths by default.
