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

Keynote is the first concrete format owner to consume that projection in a
production package path. Ingress counts all parsed IWA objects before one
fallible exact reservation, stores only `(identifier, component, object)`
locators, sorts them once, rejects duplicate identities, and performs later
lookups by binary search. Slide records, semantic slides, builds, and text
storages reserve fallibly from validated source counts. The text adapter is
invoked only for graph-reachable typed storage payloads and receives the
smaller of the physical message ceiling, the wire hard ceiling, and the
remaining package-wide semantic text and fragment budgets. Streaming wire
preflights count slides and used build/drawable references before generated
Prost vectors are materialized; they also charge retained slide names and
build/transition identifiers before semantic ownership conversion. Aggregate
storage, fragment-range, reference, and UTF-8 counters include a content-free
semantic path in every limit failure. Common-wire byte, field, nesting, and
work ceilings are translated into the same format-owned counted diagnostic.
Text extraction performs a checked sizing pass and one fallible destination
reservation instead of building a temporary vector of cloned strings. These
are bounded allocation and lookup-shape guarantees for the migrated fields;
ignored nested fields still materialized by the generated Prost graph remain
bounded only by the physical message profile and require a later focused
projection. No throughput or RSS improvement is claimed without a
representative benchmark.

The concrete Numbers package uses the same bounded lookup shape without
sharing Keynote's format graph. It counts all component objects before exact
fallible reservation, stores one compact `(identifier, component, object)`
locator and at most one primary-message classification per object, rejects
duplicate global identities, sorts both arrays once, and resolves later
references by binary search. This replaces package-wide linear object lookup
and the previous one-index-entry-per-message amplification. Checked read
options combine physical archive limits with non-zero hard-bounded ceilings for
objects, rooted sheets, semantic tables, and rooted reference occurrences.
Sheet/reference/table counts are charged before their semantic result vectors
grow; structured table output is fallibly reserved one item at a time under
the same table ceiling. Legacy type-6000 model discrimination still requires a
complete bounded parse because genuine type-6000 table-info payloads are valid
false positives; the object and physical message ceilings bound that fallback
until its schema family receives a lazy Buffa projection.

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

## 2026-08-08 amendment: direct Keynote settings projection

The focused Keynote settings reader uses the existing schema-directed Show and
SlideTree preflight, including the caller's slide-reference ceiling, before a
private Buffa settings projection is forced. Unlike the full Show projection,
it does not allocate or retain the slide-node identifier collection and does
not initialize the package's full semantic slide cache. This is a bounded
allocation-shape statement only. The format and codec validation layers may
both scan the payload, and no O(1), single-pass, latency, RSS, allocation-count,
or throughput result is inferred without measurement.

## 2026-08-08 amendment: bounded Pages section-name rewriting

The Pages section-name transaction retains the package's original limits and
checks input/output package bytes, entry and aggregate bytes, IWA object and
message counts, retained name bytes, protobuf bytes, fields, nesting, and
rewrite work. Fallible reservations precede owned copies, size arithmetic is
checked, and the complete candidate is reopened under the same limits before
publication. A no-op shares the existing `Arc` and avoids package reassembly
and reparsing.

For a changed exact package, the implementation locates the selected native
section privately, performs a bounded canonical-wire preflight, replaces only
length-delimited field 26, preserves the complete IWA object header with
`replace_message_preserving_header_with_limits`, recompresses one component,
and reassembles the source catalog. Untouched ZIP members and their raw local
and central records remain exact except for central-directory offsets that
must move when the changed member length changes.

No generated Buffa or Prost message is materialized for this preservation
rewrite. That is deliberate: raw validated field records are the authority for
unknown fields, duplicate ordering, encoded keys, and length headers. This is
an allocation-shape and boundedness statement only; it makes no O(1),
single-pass, latency, RSS, allocation-count, or throughput claim without the
measurement protocol above.

## 2026-08-08 amendment: retained semantic text accounting

The neutral iWork aggregate now measures the UTF-8 bytes of every owned string
that survives in the archive-free result. Keynote accounting includes the show
title and owned unknown animation/transition identifiers; known static effect
labels consume no owned-text budget. Its failure observation follows the same
title, slide content, additional storage, and speaker-notes order exposed by
the public semantic model.

Pages now separates rendered text length from retained text. Section names,
headings, paragraphs, and storage text are charged once, while synthesized
rendering separators and temporary `Option<Box<str>>` slots are not charged as
UTF-8. A rejection reports the checked observed byte count (or `usize::MAX` on
arithmetic overflow) through the focused crate and root facade rather than
fabricating `limit + 1`. Exact-limit and one-under regressions lock these
rules. This is a correctness and boundedness result; it makes no latency,
allocation-count, throughput, or peak-RSS claim.

## 2026-08-16 amendment: bounded Numbers root, sheet, and storage views

The focused Numbers package removes four eager application-payload decodes from
its production package reader. The type-1 document root is projected through
the existing strict `numbers_sheet_order_codec` view; standard and form-based
type-2/3 sheets use the bounded name/drawable preflight; and compatibility
storage text uses the shared bounded `ValidatedStorage` path from
`litchi-iwa-text-wire`. The semantic reader therefore retains only the root
sheet references, borrowed sheet name and drawable spans, and the text needed
by its public diagnostic. It does not construct a generated document, sheet,
form-sheet, or storage object for these paths.

Each selected payload is bounded before publication by input bytes, traversed
fields, nesting, aggregate wire work, and the caller's remaining semantic
references or text output. Sheet names, drawable references, and joined storage
text are checked before owned allocation; malformed storage candidates retain
the established compatibility skip behavior, while malformed rooted document
and sheet ownership fails atomically. Standard/form sheet parity, duplicate
references, missing objects, and strict-versus-Buffa disagreement remain
failure cases. Raw component bytes remain the preservation authority, so the
private Buffa views neither retain unknown fields nor encode replacement data.

This is an allocation-shape and boundedness improvement, not a measured
latency, RSS, or complete Numbers Buffa-laziness claim. Table, tile, formula,
sidecar, and other native graph paths remain separate migration work and may
still use generated Prost values behind their own limits.
