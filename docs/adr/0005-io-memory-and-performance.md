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
state is never silently evicted. Cache behavior is semantically invisible; see
[Caches and retained state](#caches-and-retained-state) for the two kinds of
retained state that are not evictable caches.

Every operation charges a hierarchical resource budget supplied by an execution
context. Production-safe desktop, server, and trusted-batch profiles are finite.
Callers may raise specific configurable limits but cannot bypass integer,
nesting, decompression, or structural safety ceilings. Limit errors identify the
resource, observed value, limit, and object path. A budget may be charged in
chunks through a lease; see [Budget leases](#budget-leases).

Scratch storage is an explicit capability. Litchi never spills decrypted or
sensitive content to plaintext temporary files automatically. Supported scratch
providers include memory, encrypted temporary storage, and caller-defined
stores; absence yields a typed resource error.

## Output

Ordinary save creates a fresh artifact. Filesystem replacement uses a sibling
temporary artifact, validation/finalization, flush/fsync as supported, and atomic
replacement. Cancellation leaves the destination untouched and performs
best-effort identity-aware cleanup of the temporary artifact. Caller-owned
non-atomic sinks report incomplete output and
bytes written. "Flush/fsync as supported" is the default `Full` durability of
every ordinary save; a caller may select a weaker level for one save (see
[Save durability levels](#save-durability-levels)).

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

### Save durability levels

Full durability is the default and is what every ordinary `save` does, with
unchanged bytes, and with unchanged system calls on success and on every
failure before the replacement: the staged temporary artifact is flushed and
synchronized before the atomic replacement, and the destination's parent
directory is synchronized after it where the platform supports that. A caller
may select a weaker level for one save, and only by passing an explicit
`litchi_core::Durability` to a `*_with_durability` method: `FileOnly` skips the
parent-directory synchronization, and `NoSync` also skips the temporary
artifact's synchronization. The level is a per-call argument. It is never read
from the environment, a global or thread-local value, or a package's stored
save preferences, so selecting one changes no preference and revokes no
exact-source authorization.

Every level keeps the rest of this section: the sibling temporary artifact,
every check the route makes before publishing, validation and finalization
before replacement, the one atomic replacement, a destination left untouched
by any failure or cancellation before that replacement, best-effort
identity-aware cleanup of the temporary artifact, and typed errors. A level
never attempts a step it skips, so it cannot fail at one; only `Full` can
report that the destination was replaced before its directory synchronization
failed. What a weaker level gives up is durability across an operating-system
crash or power loss after the save has returned: after `FileOnly` the
destination names what it named before the save or the complete new artifact,
and the replacement itself may be lost; after `NoSync` nothing is promised,
and on a filesystem that does not write file data before the rename that
exposes it the destination may be truncated. A process crash needs no
synchronization at any level. The `litchi_core::Durability` documentation
states each level's guarantees and is the normative list.

### Exact-output streaming of iWork packages

`Package::write_to` on the focused Pages, Numbers, and Keynote packages is the
exact-output seam. It streams the retained ZIP artifact to a caller-owned
`Write` sink without allocating another package-sized output buffer. It
tolerates `Interrupted`, tracks the offset reached by conforming partial
writes, and reports zero-write, over-report, and ordinary sink failures through
typed, redacted `WriteError` values carrying `bytes_written`. `write_to` does
not flush or sync its sink and does not rename, atomically replace, or durably
publish a filesystem path; that is `save`'s job.

### Durable filesystem publication of iWork packages

`litchi-iwa-archive` owns the one reusable durable filesystem-publication
operation for iWork packages. `litchi_pages::Package::save`,
`litchi_numbers::Package::save`, and `litchi_keynote::Package::save` each
delegate to it; no format keeps a parallel temporary/rename implementation.
These methods publish the already validated exact artifact and add no semantic
selection, mutation, candidate construction, reopen/readback, or locality
work. The read-only `litchi` iWork facade adds no save API of its own.

The helper validates the destination before invoking its one-shot writer
callback, rejects symbolic-link and non-regular destinations, captures
permissions from an existing regular file, and creates the staging file as a
private sibling in the destination directory. The callback writes only to that
staging file; each `Package::save` invokes `Package::write_to` once with the
staging `File` and never first materializes `to_bytes` or a second
package-sized buffer. A callback, stream, validation, or allocation error
before replacement is returned through the typed save error, and the
destination is not replaced. Cleanup is identity-aware: the private temporary
is removed only while its pathname is still proved to name this operation's
staging file; otherwise cleanup is disarmed rather than unlinking an entry
whose identity is uncertain. The same rule applies to pre-replacement
permission hardening, flush, file-sync, revalidation, and replacement failures.

The publication sequence is callback/write, private Unix staging mode `0600`,
`File::flush`, staging-file `sync_all`, final destination and staging-path
identity revalidation, and same-directory atomic replacement. Only after
replacement does the helper apply the existing destination's ordinary
permissions through the still-open published file descriptor, synchronize that
file when permissions were applied, and synchronize the parent directory where
supported. New Unix destinations remain mode `0600`; inherited Unix
set-user-ID and set-group-ID bits are cleared.

Replacement is the commit point. A post-replacement permission or
published-file synchronization failure, or a parent-sync failure other than
Unix `InvalidInput`/`Unsupported`, is an explicit committed error; those two
parent-sync kinds mean the capability is unavailable and return success.
Callers use `was_committed()`/`is_committed()` and the typed stage and I/O
accessors to learn that replacement already committed, so blind retry is
unsafe. The save errors distinguish staging/write failure from publication
failure. Their `Display`/`Debug` forms redact destination and temporary paths,
package bytes, and lower-layer diagnostic text; `Error::source()` and
`Error::io_error()` intentionally still expose the original operating-system
error to a caller that asks for it.

Windows receives the best-supported same-directory atomic replacement and file
synchronization, but no stronger parent-durability or POSIX-rename guarantee;
reparse points are rejected and broader Windows metadata/ACL transfer is not
portable. Other targets fail before the writer runs. The contract assumes a
trusted, caller-controlled parent and ordinary local-filesystem rename/sync
semantics; network and userspace filesystems, concurrent same-identity
pathname replacement, and platform metadata can weaken identity, atomicity,
cleanup, or crash-durability guarantees.

## Caches and retained state

"Cache behavior is semantically invisible" describes evictable parsed values
behind a handle. Two other shapes of retained state are admitted, each under
its own rules.

### Retained digest memos

An immutable snapshot, and a package facade that publishes snapshots, may
retain a memo that is not an evictable weighted cache of parsed values but a
bounded table of digests over bytes the owner already holds. [ADR
0032](0032-snapshot-derived-value-memos.md) widens the values such a memo may
hold. Every one of the following holds, and the implementation proves rather
than asserts it:

- The memo is keyed on the identity of an allocation the owner already holds
  and retains a strong reference to it, so a key cannot be recycled by a
  different payload and a hit proves byte identity.
- Every entry names an allocation the owning snapshot's own package holds; the
  process retains no payload solely because a memo names it, and a memo
  carried onto another package is re-projected onto that package's own
  allocations rather than inherited.
- Adoption: a facade that retains such a memo across operations adopts the
  memo of the snapshot each publication produces and releases it at every
  mutation that produces no snapshot to adopt, so the memo never outlives the
  graph it describes. This names when a facade must refresh its memo, not the
  only point at which it may take one: a facade holding no memo may fill the
  slot from a capture of its current, unmutated graph, provided every kept
  entry is re-projected onto the facade's own allocations (entries naming
  allocations the facade's package does not hold are dropped).
- A miss is an ordinary recomputation, so no value, refusal, limit or
  published byte depends on which entries a memo holds. The memo is rebuilt or
  projected, never mutated in place, and it holds no limit and raises no error
  the unmemoized path would not raise.
- Its resident cost is bounded by the part count, is stated where the memo is
  implemented, and is charged through fallible reservation, so exhaustion is a
  typed resource error and not an abort.
- A reused value is re-derived by a `debug_assert!` in test and debug builds,
  so the owning crate's suite proves value identity rather than only
  compilation.

This does not weaken the eviction rules for semantic payload caches and does
not permit ambient process-wide state.

### Retained state that is not a cache

State that a public value owns after the operation that produced it has
returned, whose bytes a caller therefore holds for as long as it keeps that
value (for example a whole serialized document held inside a public value), is
a retention policy rather than a cache. It is declared, bounded and
observable:

- its ceiling is a member of the operation's own finite limit policy, of the
  same kind as the already-enforced live-memory ceilings there, intersected
  with every other member when two policies meet;
- the value exposes what it is holding, and exposes a way to release it
  without changing what the value means;
- the retained allocation is one an operation already made, not a new copy of
  it; retention is a decision not to free.

Where a path has an execution context, that context is the authority for the
ceiling. Where a path has **no** execution context, as the PPTX
opened-presentation path does not, the operation's finite limit policy carries
the ceiling. That is an interim location, not a second budget system: when a
path gains an execution context under
[ADR 0031](0031-execution-context-budgets.md), its retention ceilings move
there.

Exceeding such a ceiling is **not** a refusal when the work the retention
avoids is always available by recomputation. The typed limit error protects a
caller from work or memory it did not ask for; there is nothing to protect it
from when the alternative is the path the library takes anyway. A retention
ceiling that is exceeded therefore falls back to recomputation silently and
observably (the value reports that it is holding nothing), and the ceiling
never changes any result, refusal or published byte. Ceilings whose
alternative is *not* recomputation keep their typed `Limit` refusal; this
reading applies only to retention that is a pure optimization.

The scratch-storage rule is the reason the fallback is recomputation rather
than a spill: a retained document archive is document content, and litchi
never writes it anywhere automatically.

## Budget leases

An operation may charge a budget through a `Lease` (`Budget::lease`,
`ExecutionContext::lease`): it claims up to a stated chunk of one resource at
a time, with the same atomic check-and-add per hierarchy level as any other
charge, and hands the units out locally. That is charging the budget, not a
second budget system. A lease never holds more than one chunk: a claim near a
limit shrinks to the room left, and units a refund takes back beyond one chunk
return to the budget at once.

What stays exact: no level is ever above its limit, even transiently; when
every charge a holder makes of a resource goes through one lease, that holder
is refused on the same charge, at the same level and with the same resource,
observed value and limit as exact accounting (a charge it makes outside that
lease, directly or through another lease, sees the lease's unspent units as
used); a refused claim changes nothing; and the units a lease holds but has
not handed out return to the budget when it is released or dropped, including
on an error, a cancellation or a poisoned operation, so every counter then
settles at exactly what was handed out.

What becomes rough, and only for other holders of the same budget: the
pre-claimed units count as used, so another holder can be refused earlier than
exact accounting would refuse it, by up to one chunk per other open lease, and
a refusal it receives reports usage that includes the pre-claims. Each lease's
chunk size is stated with its reason by the code that opens it. Reservations,
`consume` and every charge made outside a lease are unchanged. ADR 0031's
reservation dimensions (`Workers`, `IoConcurrency`) are admission-time
reservations, not charges, and are not leased.

## Measurement contract

Representative small, large, sparse, media-heavy, encrypted, and malformed
corpora gate open latency, lazy lookup, concurrent reads, disjoint writes,
patching, and save. Track peak resident memory, allocations, copied bytes,
decompression, cache misses, lock/contention time, CPU utilization, and scaling.
Optimization decisions require profiles, flame graphs, and statistical evidence;
intuition alone is not accepted.

The allocation-shape, bounded-accounting, and indexing rules in the rest of
this ADR are correctness and boundedness properties. They are not latency,
throughput, allocation-count, zero-copy, or peak-memory results; any such claim
needs the measurement above.

Instrumentation is an opt-in runtime-neutral observer with optional tracing and
profiling adapters. It never records document content, credentials, or sensitive
paths by default.

## IWA decoding and preservation

### Generated bindings and lazy views

Generated IWA protobuf bindings are kept minimal. Prost runtime type-name
metadata stays disabled because no production path consumes it. Production
archive-header and focused projections use exact-version Buffa 0.9.1 lazy views
behind private codecs in `litchi-iwa-protos`. The complete schema corpus is not
generated eagerly for every format, and generated Buffa types, decode contexts,
and codec errors never cross a crate's public boundary. `litchi-iwa-protos`
keeps its normal `prost` dependency for the generated owners that still use it;
`litchi-keynote`, `litchi-numbers`, and `litchi-pages` use Prost only as a
dev-dependency for test oracles and fixture builders.

Buffa lazy decoding of untrusted IWA bytes is preceded by a schema-directed
common wire-tree preflight. One aggregate policy bounds scanned bytes, fields,
nesting, repeated metadata items, deferred-message occurrences, and a
conservative decoded-memory envelope before a lazy view is constructed. The
archive-header adapter then visits every deferred child exactly once, checks
proto2 required presence, and projects directly into physical metadata with
fallible destination reservations; generated `to_owned_message` is not a
production ingress path. Buffa 0.9.1 still uses ordinary infallible `Vec`
growth for some internal lazy metadata, so the preflight bounds hostile
amplification but does not give typed recovery from global allocator
exhaustion or a language-level exact resident-memory bound. A contract that
needs either property must use a streaming handwritten cursor or a corrected
Buffa runtime.

Core archive metadata is projected into core-owned `FieldPath`, `FieldInfo`,
and closed-enum wrappers. Optional presence and unknown signed enum values are
retained exactly. Preflight charges both the transient Buffa representation and
the neutral destination vectors, including unknown closed-enum records, before
publication.

The recurring codec pattern is: a strict handwritten router owns framing,
required and duplicate field checks, canonical scalar and Boolean validation,
and unknown/group handling, and charges every resource; the private Buffa view
is forced afterwards only for parity. The view retains no unknown fields and
never encodes replacement bytes.

### Preservation authority

Lazy re-encoding is not the preservation boundary. Original source-backed
header bytes and raw spans remain authoritative for exact no-ops, unknown
fields, duplicate occurrences, and non-canonical encodings. Buffa encoding is
used only for a canonical header created after a semantic change. A rewrite
replaces only the selected fields of a validated raw record and copies
everything else from the source; no generated Buffa or Prost message is
materialized for a preservation rewrite, because raw validated field records
are the authority for unknown fields, duplicate ordering, encoded keys, and
length headers.

### Physical ZIP ingress

Physical iWork ingress bounds the two ZIP name spellings independently before
copying either one. Local and central names, extras, and comments are charged
cumulatively; compressed sizes are rejected before payload materialization.
For legacy packages, catalog, component-catalog, and detection paths reject a
nested `Index.zip` from its declared uncompressed size before decompression.
Limit failures retain the resource kind and exact observed and maximum values.
This layer does not promise a semantic object path in every physical ZIP
diagnostic.

### Borrowed wire readers

Borrowed IWA wire readers use the common source-bound `WireView<'a>` and
`WireFieldView<'a>` when interpreting recognized fields: one borrowed source
and compact spans avoid per-field slice metadata, payload ranges are sliced
through validated spans, and schema-owned key/length framing can be required
without changing the permissive unknown-field parser. Singular wire overlays
index base and overlay field numbers once and emit one exact-capacity output,
so sparse updates do not repeatedly reparse a growing message. Source-built
Pages, Numbers, and Keynote chart updates locate their single chart payload
with one linear scan and no temporary index allocation before decoding or
invoking a mutation callback.

Reference-line graph updates avoid a full generated-Prost round-trip: bounded
raw fields are merged by repeated-field occurrence and only recognized values
are replaced, so unknown graph bytes are copied once at their original nesting
positions. The candidate field collection is validated before publication.

### Focused text projection

The storage-text Buffa projection is deliberately smaller than a canonical
format schema root. A derived five-file projection reads only repeated field 3
of `TSWP.StorageArchive`, is hard-capped at 32 KiB of generated Rust, disables
unknown retention, and exposes only borrowed text fragments through a private
wrapper. Common-wire preflight bounds root bytes, fields, field type, fragment
count, UTF-8 bytes, and the conservative repeated-view allocation before stock
Buffa 0.9.1 runs. Other length-delimited fields remain opaque by design, so
this projection suits only call sites whose policy needs semantic text rather
than eager validation of every unrelated known child. The caller-owned source
remains authoritative.

### Retained semantic text accounting

The neutral iWork aggregate measures the UTF-8 bytes of every owned string that
survives in the archive-free result. Keynote accounting includes the show title
and owned unknown animation/transition identifiers; known static effect labels
consume no owned-text budget. Its failure observation follows the same title,
slide content, additional storage, and speaker-notes order exposed by the
public semantic model.

Pages separates rendered text length from retained text. Section names,
headings, paragraphs, and storage text are charged once, while synthesized
rendering separators and temporary `Option<Box<str>>` slots are not charged as
UTF-8. A rejection reports the checked observed byte count (or `usize::MAX` on
arithmetic overflow) through the focused crate and root facade rather than
fabricating `limit + 1`.

## Keynote reads

Keynote package ingress counts all parsed IWA objects before one fallible exact
reservation, stores only `(identifier, component, object)` locators, sorts them
once, rejects duplicate identities, and performs later lookups by binary
search. Slide records, semantic slides, builds, and text storages reserve
fallibly from validated source counts. The text adapter runs only for
graph-reachable typed storage payloads and receives the smaller of the physical
message ceiling, the wire hard ceiling, and the remaining package-wide semantic
text and fragment budgets. Streaming wire preflights count slides and used
build/drawable references before generated Prost vectors are materialized, and
charge retained slide names and build/transition identifiers before semantic
ownership conversion. Aggregate storage, fragment-range, reference, and UTF-8
counters include a content-free semantic path in every limit failure.
Common-wire byte, field, nesting, and work ceilings are translated into the
same format-owned counted diagnostic. Text extraction performs a checked sizing
pass and one fallible destination reservation instead of building a temporary
vector of cloned strings. Ignored nested fields still materialized by the
generated Prost graph are bounded only by the physical message profile.

The focused settings reader uses the schema-directed Show and SlideTree
preflight, including the caller's slide-reference ceiling, before a private
Buffa settings projection is forced. Unlike the full Show projection, it does
not allocate or retain the slide-node identifier collection and does not
initialize the package's full semantic slide cache.

`Slide::plain_text` computes a checked aggregate UTF-8 length for its semantic
text values, reserves one output `String` when representable, and appends
directly in source order without the intermediate `Vec<String>` used by
`all_text()`. Empty text storages are filtered, while empty modeled values keep
their separators.

## Numbers reads

### Object index

The Numbers package uses the same bounded lookup shape without sharing
Keynote's format graph. `Index::from_components` counts every physical object
before checking `max_objects` and before reserving locator/type storage. It
stores one compact `(identifier, component, object)` locator and at most one
primary-message classification per object, sorts both arrays once, and
resolves later references by binary search. Checked read options combine
physical archive limits with non-zero hard-bounded ceilings for objects, rooted
sheets, semantic tables, and rooted reference occurrences. Sheet/reference/
table counts are charged before their semantic result vectors grow; structured
table output is fallibly reserved one item at a time under the same table
ceiling.

Duplicate identifiers within one component are rejected. Across distinct
physical components, only source-authoritatively exact copies of one object
identifier are admitted, and an admitted alias group coalesces to one logical
locator and one primary-type entry; `Package::object_count` still reports the
physical count. `ArchiveObject::same_content_ignoring_offsets` compares decoded
`ArchiveInfo`, exact raw messages, header and payload lengths, and retained
original and original-canonical header bytes without allocation; only
component-relative header and payload offsets are ignored, so payload,
metadata, raw-header, or framing divergence fails closed. Lookup comparison
work is based on the deduplicated logical locator set, while index allocation,
population, sorting, and comparison topology are charged from the physical
count.

Legacy type-6000 model discrimination requires a complete bounded parse
because genuine type-6000 table-info payloads are valid false positives; the
object and physical message ceilings bound that fallback. Strict shape
classification ignores TableInfo false positives, while schema-shaped legacy
models hit the table budget before Prost decode/materialization and cannot
commit rejected candidate budgets.

### Root, sheet, and storage views

The type-1 document root is projected through the strict
`numbers_sheet_order_codec` view; standard and form-based type-2/3 sheets use
the bounded name/drawable preflight; and compatibility storage text uses the
shared bounded `ValidatedStorage` path from `litchi-iwa-text-wire`. The reader
retains only the root sheet references, borrowed sheet name and drawable spans,
and the text needed by its public diagnostic, and constructs no generated
document, sheet, form-sheet, or storage object for these paths. Each selected
payload is bounded before publication by input bytes, traversed fields,
nesting, aggregate wire work, and the caller's remaining semantic references or
text output. Malformed storage candidates keep the compatibility skip
behavior, while malformed rooted document and sheet ownership fails
atomically. Standard/form sheet parity, duplicate references, missing objects,
and strict-versus-Buffa disagreement are failures.

### Table extraction

The table extractor treats its bounded raw-wire preflight as the authoritative
projection of the type-6218 `TST.RichTextPayloadArchive` envelope: it scans
the complete message, requires one canonical length-delimited local storage
reference in field 1 and one cell-owner value in field 3, validates the nested
reference framing, and forwards the storage identifier to the bounded
`litchi-iwa-text-wire` storage path. Production never calls
`RichTextPayloadArchive::decode`. Storage validation and materialization charge
physical bytes, fields, nesting, references, wire work, UTF-8 text, and the
caller's remaining aggregate semantic text budget before owned output is
published; document projection keeps the strict validation/text length parity
check.

Native type-6002 `TST.Tile` payloads go through the no-`Vec` transactional
`numbers_table_cell_storage_codec::decode_tile_with_visitor` path; production
never calls generated `Tile::decode` or `TileRowInfo::decode`. Row callbacks
expose payload slices borrowed from the caller-owned component source. The row
visitor applies a prefix materialized-cell guard before materializing a cell.
After a semantic cell error is retained, later rows continue through strict
wire validation so the decoder can return its full `DecodeReport` without
growing the semantic table. The extractor then charges that complete report,
charges the aggregate materialized-cell count, and only then surfaces the
retained semantic error. A later malformed wire error therefore overrides an
earlier semantic cell error, while an aggregate materialized-cell limit takes
precedence over a retained semantic error. Each table candidate owns a local
table and candidate budget; rejected candidates publish neither cells nor
retained output, although bounded decode work remains an aggregate admission
cost.

`TST.TableDataList` and `TST.TableDataListSegment` payloads (types 6005, 6201,
and 6011) go through a strict handwritten router and the private Buffa
projection, for both `Package` and `Document` extraction and for the generic
text registry. Known repeated entry/segment fields must carry
length-delimited wire shape before a candidate is admitted to full decode.
Production retains only the final semantic table values needed by a candidate,
plus bounded segment-id/key tracking where required, with fallible reserves.
The text adapter stages only non-empty borrowed strings, copies them with
fallible reservations, and publishes the owned vector only after the complete
decode succeeds. Required-field, duplicate, canonical-wire, UTF-8, reference,
segment-range, ownership, and candidate-publication checks apply. Input,
output, field, nesting, and aggregate work are bounded by the `WireLimits`
hard ceilings (512 MiB input, 512 MiB output, 1,000,000 fields, depth 64,
16,000,000 work units), with tighter per-payload and semantic text budgets
applied before fallible allocation.

Read-only table-cell storage queries (string-value resolution, entry
presence, and rich-text entry counts) use the borrowed root/segment snapshots and strict streaming visitors of the
doc-hidden `litchi_iwa_protos::numbers_table_cell_storage_codec`. Each admitted
payload has explicit source-byte, field, work, nesting, reference, and
selected-text limits, and successful reports merge into one operation-local
field/work/reference/text budget across root candidates, ordered segments, and
rich-text package scans. Failed resource checks stop the operation; ordinary
malformed candidates conservatively charge their complete source length before
compatibility selection continues. No partial strings, counts, or package
mutation are published after a later segment failure.

### Comments

`TSD.CommentStorageArchive` payloads go through a strict handwritten codec and
a private Buffa parity projection. `Package` keeps strict comment validation
and resolution; `Document` deliberately skips comment resolution. Per-cell
comment values are materialized only after strict validation, with bounded,
fallible reserves and candidate-local publication. The bounded comment-cell
scan-work counter spans the whole model/cell scan rather than resetting for
each table.

The comment-reply read is source-immutable and publishes no partial result.
Package ingress keeps the physical archive and semantic object/reference/text
ceilings; each strict comment-storage decode keeps typed byte, field, work,
nesting, reference, and text ceilings. Census and output vectors use fallible
reservations, and storage duplicate detection and selected reply lookup use
hash indices, so the census/projection path is linear in recognized storage
and reply counts. Root selection and the census may inspect the same payload in
separate bounded passes; there is no single aggregate ledger across them.

### Formulas

A bounded owned copy of the admitted formula source (`FormulaArchiveBytes`) is
the exact-byte preservation authority and is never re-encoded. Strict
schema-directed preflight charges bytes, fields, nesting, work, and
repeated-entry/node counts before a private Buffa lazy per-node/event stream is
admitted. The stream is forced for every selected node or event used by
rendering; it materializes no generated repeated field or `LazyRepeatedView`
and never calls `to_owned_message`. The canonical Prost builders and decoder
are a test-only oracle for differential fixtures and parity checks.

### PackageMetadata inspection

The PackageMetadata visitor (`inspect_package_metadata_with_visitor`) derives
package/effective-message limits and `WireLimits` for input, output, fields,
work, nesting, components, and references, and a strict two-pass
preflight/visitor charges them before observations are accepted. Visitor
results stay borrowed; byte, field, work, nesting, and allocation failures are
typed and publish no partial observation. Admission is narrower than permissive
Prost decoding for the selected projection: selected fields require canonical
singular required and nonzero values, valid UTF-8, canonical Booleans, complete
UUIDs, and valid nested references, so malformed versioned records fail even
when a query would otherwise ignore them. The codec validates only this
projection, not the complete `PackageMetadata` schema; unknown noncanonical
keys or lengths are rejected, while unknown noncanonical varint values and
balanced unknown groups are accepted.

## OPC exact-source authorization

For OPC, direct byte-identical no-op publication has one authority: the owning
package or source-backed object must retain its exact source artifact and an
unrevoked exact-source authorization. Preservation provenance, ZIP indexes,
and reconstructed graph equality are planning evidence only; none of them can
authorize exact passthrough or a normalizing full-writer fallback. Any mutable
OPC seam revokes that authorization. A changed owned source is publishable only
through a proven preservation plan; if physical framing or opaque members cannot
be preserved, publication returns a typed capability refusal before output.

## iWork edit transactions

### Common resource contract

Every focused iWork edit transaction (Pages, Numbers, and Keynote) follows one
resource contract:

- **One operation-local budget.** A private transaction budget (for example
  `TransactionBudget`, `CaptionBudget`, `PlaybackBudget`, `GeometryBudget`)
  carries every phase: physical source and catalog inventory, rooted selection
  and ownership census, strict codec prepare/execute reports, Metadata
  UUID/external-reference/save-token transitions, archive serialization and
  exact sizing, Snappy and Deflate compression, ZIP reassembly, candidate
  reopen, semantic readback, locality verification, and retained exact
  source/target patch artifacts. Residual limits pass from phase to phase;
  no phase resets to full limits. Independent source and candidate scans are
  charged independently, nesting is merged by maximum, and a borrowed payload
  already charged by an enclosing phase is not charged again. Counters cover
  input, intermediate and final output, package entries, payloads, objects,
  messages, items, fields, work, nesting, references, components, additions,
  allocations, retained bytes, and scratch bytes, with checked arithmetic;
  overflow fails closed.
- **Preflight before allocation.** Before Snappy decompression or
  `Archive::parse` allocates a selected component, the owner derives the
  decoded archive extent and physical object/message inventory from the parsed
  source catalog, charges it, and then cross-checks the actual decompressed
  length. Name selection charges the compared source and selector bytes for
  every candidate; position selection charges the inspected prefix. Exact
  source/patch and no-op artifact comparisons charge equal-length byte scans.
- **Prepare, then execute.** Codecs expose a prepare step that validates the
  source, sizes the exact output, and returns execution requirements without
  allocating the candidate, and an execute step that replays the residual
  limits and performs one fallible output reservation, then strictly scans the
  candidate. `Archive::encoded_len` measures exact source-authoritative IWA
  serialization, `SnappyStream::maximum_compressed_len` gives a bounded
  compression ceiling, and `PreparedReassembly` exposes exact ZIP execution
  requirements before the final allocation. A codec's allocation, scratch, and
  retained counters are logical/requested bounds, not allocator telemetry.
- **Staged private candidates, atomic publication.** Native archive, Metadata,
  compression output, and final ZIP values may be staged as separate private
  candidates (later stages can consume identifiers produced by earlier ones),
  each charged before allocation. Nothing is published until every stage, the
  complete candidate reopen, semantic readback, and locality verification
  succeed; any limit, allocation, or verification failure leaves the immutable
  source and the published package unchanged. There is no single global
  output-free preflight or single allocation per transaction, and the
  aggregate intermediate-output ceiling may exceed the final physical package
  ceiling because member and final artifact limits are enforced again by
  reassembly and reopen.
- **No-ops share the source.** An exact no-op shares the source snapshot and
  constructs no candidate. Exact patch artifacts are bounded by the package
  input and output limits.
- **Typed failures.** Output, allocation, and limit failures keep typed
  limit/error mapping into the format's content-free `LimitKind` vocabulary.

### Per-operation rules

These add to or narrow the common contract:

- **Pages section names.** The package's original limits cover input/output
  package bytes, entry and aggregate bytes, IWA object and message counts,
  retained name bytes, protobuf bytes, fields, nesting, and rewrite work. A
  changed name replaces only length-delimited field 26 of the selected
  section, recompresses one component, and reassembles the source catalog;
  untouched ZIP members and their raw local and central records stay exact
  except for central-directory offsets that must move.
- **Pages footnotes.** `rewrite_custom_mark_wire` checks the selected/after
  field arithmetic instead of saturating, and overflow is `WireFields`.
  Candidate verification runs `native_footnotes` before semantic readback, so
  aggregate text and entry ceilings surface as `TextBytes` and `Entries`
  (`TextTooLarge` and `TooManyBodyStorages` map to those categories). One
  `FootnoteSemanticBudget` is shared across source and candidate projections,
  and staged text and custom-mark copies use fallible exact reservations.
- **Pages drawable order.** `pages_drawable_order_codec` keeps complete raw
  reference records and interleaved root fields as the preservation authority,
  including unknown balanced groups and overlong unknown scalar framing. It
  rejects noncanonical known keys, lengths, and values, missing or zero
  identifiers, duplicates, and non-permutation edits. The outer IWA
  object-length prefix may be canonicalized by the archive serializer.
- **Numbers names.** The calculation-engine dependency route is inspected
  through the strict `numbers_table_cell_dependency_codec`; the raw field-3
  root reference and repeated tracker field-6 records are cross-checked
  against object metadata, local-reference framing, and declared paths, and
  the opaque volatile name-dependency field is checked with a narrow raw-wire
  presence scan (an empty coordinate set is not a dependency). The pivot guard
  scans only `TST.TableModelArchive` field 85, requires one canonical local
  `TSP.Reference` when present, and does not walk the table data-store graph.
  The transaction charges the complete over-approximation (`selected changes ×
  rooted topology` plus the quadratic table term and object term) against
  `WireWork` before any changed component is scanned or rewritten. For the
  save-token sidecar, the selector count is bounded by
  `RewriteOptions.max_components`, the report keeps `additions = 0`, and the
  caller precharges decompressed type-11006 payload bytes times the changed
  semantic-operation upper bound for visitor locator matching (compressed
  Snappy bytes are charged separately).
- **Numbers comment clear.** The ownership pass charges complete BNC cell
  payload bytes before the handwritten cell parser runs. Archive-owner
  inspection accumulates header work and reference occurrences across every
  physical component. The comment graph census applies one aggregate
  semantic-reference ceiling before allocating its lookup sets and uses
  bounded `HashSet` membership rather than quadratic scans.
- **Numbers comment replies.** Comment-storage and metadata codecs may receive
  conservative source-sized option profiles where the shared budget has no
  residual adapter, and preparation can allocate private plan storage before
  execution requirements are returned.
- **Numbers split cell controls and Pop-Up Menus.** The exact-byte filter can
  serialize the source `Archive` once for every structurally changed source
  member before deciding whether that member edit is a no-op; that
  serialization and its comparison bytes are charged. The same output-free
  `Archive::encoded_len` preflight and immediate actual-length replay cover the
  split scalar-control serializer. A split-component read reuses one cached
  `RegistryFacts`/physical census and is bounded only by package-ingress
  limits.
- **Numbers display formats.** The Number and Percentage rewrites preflight the
  bounded member count, maximum compressed bytes, compression-frame
  allocations, and checked work totals; compressed members, edit lists, and
  changed-name lists use fallible reservations, and changed names borrow the
  immutable package catalogs. Conservative owner aggregate charges saturate to
  the applicable ceiling before authorization. Dense format-list append updates
  the entry and `next_list_id` cursor in one prepared rewrite. Candidate
  readback keeps the caller's original message ceiling rather than raising it
  to the candidate size.
- **Keynote chart captions.** Caption input is capped at 64 MiB and copied
  only after a fallible exact `String` reservation. The existing-caption
  replacement stages the native text candidate and then the Metadata
  save-token candidate. Every internal chart-caption codec traversal (source
  scan, sizing, raw emission, candidate verification) is charged.
- **Keynote physical table sort.** Bounded by source/candidate package,
  entry/payload/object/message/reference, wire, row/column/key-cell,
  BNC-offset, retained/scratch, allocation, component, and reassembly
  ceilings. The permutation is planned without cloning each BNC payload.
- **Keynote chart legend visibility.** The private codec applies its own
  bounded message, field, work, nesting, output, retained-byte, allocation, and
  scratch ceilings inside the package and archive limits.
- **Keynote placeholder and slide-number visibility.** Before allocating its
  output, the slide router charges exact
  `6 * source.len() + output_len + 2 * parsed_fields` work.

### Package-scale indexing

Package-scale paths use indexes instead of repeated quadratic scans, charging
retained memory and work to the transaction budget. Numbers comment metadata
and reply validation use reserved ID/UUID/location indexes and one inbound
census; the scalar display-format owners and persisted sort share a
duplicate-rejecting physical-entry index. Pages footnote validation borrows
metadata until preflight, uses fallible visitor growth, and reuses one sorted
package graph census. Keynote chart-caption, movie-playback, table-title,
table-header, and table-appearance verification use fallible catalogs and
reference-frequency maps. The Keynote build-order owner performs fallible
preflight before retaining build/chunk references, preserves complete nested
reference bytes, reserves chunk groups before publication, and uses checked
global chunk offsets after a permutation; its exact no-op classification
requires both semantic identity and exact artifact equality.
