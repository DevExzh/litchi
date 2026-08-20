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

## 2026-08-16 amendment: bounded Numbers names dependency and pivot guards

The focused Numbers names transaction removes its remaining production
generated-message reads from the changed-only dependency guards. The rooted
calculation-engine route is inspected through the existing strict
`numbers_table_cell_dependency_codec`: its calculation-engine envelope,
dependency tracker, and formula-owner dependency records are borrowed
Buffa-checked snapshots. The raw field-3 root reference and the repeated
tracker field-6 records are still cross-checked against object metadata,
local-reference framing, and declared paths. The volatile name-dependency
field that is intentionally opaque in the sidecar is checked with a narrow
raw-wire presence scan; an empty coordinate set does not count as a
dependency.

The pivot guard is intentionally narrower than a table-cell read. It scans
only `TST.TableModelArchive` field 85, requires one canonical local
`TSP.Reference` when present, and does not force or recursively walk the table
data-store graph merely to decide whether a rename is supported. This keeps
the conservative native `O(T²)` rooted-table traversal, while the transaction
charges the complete over-approximation (`selected changes × rooted topology`
plus the quadratic table term and object term) against `WireWork` before any
changed component is scanned or rewritten.

Every selected dependency/pivot payload receives bounded input bytes, fields,
work, references, text, and nesting options. The format maps codec resource
errors into content-free `names::LimitKind` values, preserves fallible
collection growth, and rejects malformed, duplicate, wrong-wire, non-local,
or metadata-inconsistent routes before publication. The raw component remains
the preservation authority; private Buffa views retain no unknown fields and
never encode replacement bytes. The package-wide hard ceilings remain 512 MiB
input/output, 1,000,000 fields, depth 64, and 16,000,000 rewrite-work units.

This is a bounded allocation-shape and production-boundary result. It does not
claim a measured latency, RSS, allocator-count, or throughput improvement,
and it does not make the complete Numbers table/formula/sidecar graph
generated-message-free. The ordinary Numbers manifest still retains its
compatibility Prost paths where unrelated table extraction and editing code
requires them.

## 2026-08-17 amendment: bounded Numbers rich-text payload envelope

The Numbers table extractor now treats its bounded raw-wire preflight as the
authoritative projection of the small type-6218
`TST.RichTextPayloadArchive` envelope. The preflight scans the complete
message, requires one canonical length-delimited local storage reference in
field 1 and one cell-owner value in field 3, and validates the nested local
reference framing before returning the storage identifier. The extractor
forwards that identifier directly to the existing bounded
`litchi-iwa-text-wire` storage path. It no longer constructs a generated
`RichTextPayloadArchive` with Prost merely to read the same reference and
compare it with the already-validated projection.

This removes one redundant generated allocation and one duplicate parse from
the rich-text path without changing its source authority. Unknown envelope
bytes remain owned by the original component and are neither retained in the
private projection nor re-encoded. Storage validation and materialization still
charge physical bytes, fields, nesting, references, wire work, UTF-8 text, and
the caller's remaining aggregate semantic text budget before owned output is
published; document projection retains the existing strict validation/text
length parity check. The focused ratchet excludes only test fixtures and
rejects a production `RichTextPayloadArchive::decode`; it intentionally does
not claim that the broader Numbers extractor or crate is Prost-free.

Current verification is 298 passing and four ignored tests in the all-feature
Numbers library suite, including nine focused rich-text projection tests,
16/16 document-reader integration tests, and 267/267 boundary-policy tests.
The live boundary graph remains 64 workspace packages,
239 internal declarations, and the unchanged 14 ordered `litchi-iwa`
migration debts. Existing formula/rich-text and basic Numbers application
fixtures remain the native semantic oracles; this read-path-only change does
not alter their bytes or require a new native mutation claim. This is a bounded
allocation-shape improvement, not a measured latency/RSS result or a complete
table, tile, formula, comment, or host-editor migration.

## 2026-08-21 amendment: OPC exact-source authorization

For OPC, direct byte-identical no-op publication has one authority: the owning
package or source-backed object must retain its exact source artifact and an
unrevoked exact-source authorization. Preservation provenance, ZIP indexes,
and reconstructed graph equality are planning evidence only; none of them can
authorize exact passthrough or a normalizing full-writer fallback. Any mutable
OPC seam revokes that authorization. A changed owned source is publishable only
through a proven preservation plan; if physical framing or opaque members cannot
be preserved, publication returns a typed capability refusal before output.
