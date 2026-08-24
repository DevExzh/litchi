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

Buffa lazy decoding of untrusted IWA bytes in focused format-owned and
`litchi-iwa-core` ingress paths is preceded by a schema-directed common
wire-tree preflight. Legacy host compatibility decoders remain migration debt:
the `litchi-iwa` registry calls `archive_codec::decode_archive_info` and
`decode_message_info` directly with bounded options, without this common
preflight. One aggregate policy bounds scanned bytes, fields, nesting, repeated
metadata items, deferred-message occurrences, and a conservative decoded-memory
envelope before a lazy view is constructed.
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

The 2026-08-17 gate recorded 298 passing and four ignored tests in the
all-feature Numbers library suite, including nine focused rich-text projection
tests,
16/16 document-reader integration tests, and 267/267 boundary-policy tests.
The live boundary graph remains 64 workspace packages,
240 internal declarations, and the unchanged 14 ordered `litchi-iwa`
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

## 2026-08-20 amendment: focused Numbers type-6002 Tile production-decode exit

The focused Numbers extractor now routes native type-6002 `TST.Tile` payloads
through the no-`Vec` transactional
`numbers_table_cell_storage_codec::decode_tile_with_visitor` path. A strict
handwritten canonical router owns root and row framing, required and duplicate
field checks, canonical scalar and Boolean validation, unknown/group handling,
and aggregate resource accounting. Private Buffa lazy views are forced only
for parity after the strict route; no generated Tile value is published by
production. Row callbacks expose payload slices borrowed from the caller-owned
component source. The production ratchet rejects both generated
`Tile::decode` and `TileRowInfo::decode` (including `tst::Tile::decode` and
`tst::TileRowInfo::decode`); Prost remains a test-only oracle and fixture
builder.

The row visitor applies a prefix materialized-cell guard before attempting cell
materialization. After a semantic cell error is retained, later rows continue
through strict wire validation so the decoder can return its full
`DecodeReport`, without growing the semantic table. The extractor then charges
that complete report, charges the aggregate materialized-cell count, and only
then surfaces the retained semantic error. A later malformed wire error
therefore overrides an earlier semantic cell error, while an aggregate
materialized-cell limit takes precedence over a retained semantic error. Each
table candidate owns a local table and candidate budget; rejected candidates
publish neither cells nor retained output, although bounded decode work remains
an aggregate admission cost.

This turn also hardens legacy type-6000 admission: strict shape classification
still ignores TableInfo false positives, while schema-shaped legacy models hit
the table budget before Prost decode/materialization and cannot commit rejected
candidate budgets.

This removes the duplicate Tile parse and row-buffer copies from the production
path. The decoder's row payloads are source-borrowed, but the change is not an
end-to-end zero-copy claim. The 2026-08-20 evidence records 221
`litchi-iwa-protos` tests, 311 passed and four ignored Numbers unit tests, and
all 14 integration binaries passed (120 tests total), including
`document_reader` 16/16 and `tile_reader_integration` 2/2. Boundary is 271;
the live audit is 64 workspace packages, 240 internal declarations, and 14
unchanged ordered debts. The final temp-corpus nightly ASan smoke ran 100
runs with eight seeds. A read-only native Numbers ZIP and directory hash/UI
gate also passed against its known hashes.

No measured performance, latency, RSS, native save, or native mutation claim
follows. This is not a complete extractor, Numbers, or monolith exit: eight
production generated decodes remain in the extractor, the older `litchi-iwa`
Tile paths remain, and debt 015 plus all 14 ordered migration debts are
unchanged.

## 2026-08-21 amendment: bounded Numbers TableDataList and Segment ingress

The focused Numbers extractor now routes native `TST.TableDataList` and
`TST.TableDataListSegment` payloads through a strict handwritten wire router
and the private Buffa lazy-view projection. The same strict ingress is used by
both Package and Document extraction. Document intentionally skips comment
resolution, as required by its projection contract; Package retains strict
comment validation. Production retains only the final semantic table values
needed by a candidate, plus bounded segment-id/key tracking where required,
with fallible reserves. It does not claim zero-copy operation or the absence
of all `Vec`/`HashSet` storage.

The route preserves required-field, duplicate, canonical-wire, UTF-8,
reference, segment-range, ownership, and candidate-publication checks. Input,
output, field, nesting, and aggregate work remain bounded by the package
ceilings of 512 MiB, 512 MiB, 1,000,000 fields, depth 64, and 16,000,000 work
units, with tighter per-payload and semantic text budgets applied before
fallible allocation. Unknown bytes remain owned by the raw component and are
not re-encoded by the private view. FormulaArchive and the legacy `litchi-iwa`
comment path remain explicit migration debt; five generated extractor decodes
remain (four table-model decodes and one FormulaArchive decode). The former
generated comment decode is no longer part of this extractor path.

The 2026-08-21 evidence records 237 `litchi-iwa-protos` tests, 49 focused
extractor tests, 14 TDL integration tests, and 281 boundary tests. The live
boundary graph remains 64 packages, 240 declarations, and 14 ordered debts.
The TDL fuzz corpus has 11 seeds and completed 100 nightly ASan runs without
an artifact. A disposable native Numbers 1,200-by-8 artifact was created,
saved, closed, and reopened through the application; strings, formula
results, an intentional formula error, rich text, and a persisted comment were
observed without a repair dialog. Its archive listing contains DataList
members, but no parsed type-6011 Segment was established. This is correctness
and boundedness evidence only: it makes no native segment, native
save-mutation, performance, host-exit, or complete-monolith-retirement claim.

## 2026-08-21 amendment: bounded Numbers CommentStorage migration

The focused Numbers extractor now routes `TSD.CommentStorageArchive` payloads
through a strict handwritten CommentStorage codec and a private Buffa lazy
projection. Package keeps strict comment validation and resolution;
Document deliberately skips comment resolution. Per-cell comment values are
materialized only after strict validation, with bounded, fallible reserves and
candidate-local publication. The private projection is a parity view, while
the caller-owned raw bytes remain the preservation authority.

FormulaArchive and the legacy `litchi-iwa` comment path remain explicit
migration debt. Five generated extractor decodes remain (four table-model
decodes and one FormulaArchive decode); the former generated CommentStorage
decode is no longer part of the production extractor path. Legacy monolith and
dependency edges, plus all 14 ordered migration debts, remain unchanged.

Evidence on 2026-08-21 records 237 `litchi-iwa-protos` tests, 324 passing
Numbers library tests with four ignored (including 53 focused extractor
tests), 14 integration tests, and 281 boundary tests. The CommentStorage fuzz
corpus has 20 seeds, and its nightly ASan smoke completed 100 runs cleanly
without an artifact. A disposable native Numbers 1,200-by-8 document with
comment, formula, and rich-text content was saved, closed, and reopened
without a repair dialog. Its archive listed DataList members, but no parsed
type-6011 Segment proof was established. This is bounded correctness evidence
only: it makes no native segment or save-mutation, performance, host-exit, or
complete-monolith-retirement claim.

## 2026-08-23 amendment: Keynote plain-text allocation shape

`Slide::plain_text` now computes a checked aggregate UTF-8 length for its
semantic text values, reserves one output `String` when representable, and
appends directly in source order without first constructing the intermediate
`Vec<String>` used by `all_text()`. Empty text storages remain filtered, while
empty modeled values retain their separators. This is structural allocation
shape and ordering evidence only; it makes no allocation-count, latency, RSS,
throughput, or peak-memory claim and is not a measured performance result.

## Current present status

The recorded detached baseline reports 64 workspace packages, 239 internal
dependency declarations, and 13 ordered migration debts; its commit/tree
provenance is unavailable in this checkout. Its focused protocol codec check
records 33 passed tests; the full `litchi-iwa-protos` suite records 267 passed
and 0 failed. The baseline Numbers library gate records 356 passed tests with
four ignored. The baseline `litchi-iwa` library gate records 1,442 passed
tests; the Pages all-target gate records 149 passed tests, and the Keynote
all-target gate passes. These detached-baseline metrics are historical
evidence, not current-wave gate results. The prior XLS certification gate
remains at 2,146 passed across BIFF15/CFB254/XLSB603/XLS1274.

Native no-repair reopen evidence covers Keynote transition/title and Pages
body/background paths. A native Numbers rich/formula/comment reopen retained
the exact `Persistent TDL comment`; its archive inspection recorded 24
type-6005 entries and 0 type-6011 entries. This does not establish the
type-6011/Segment path, whose proof remains withheld.

FormulaArchive and the legacy `litchi-iwa` comment path remain explicit
migration debt; legacy monolith/dependency edges and all 13 ordered debts
remain open. These checks provide focused correctness/boundedness and native
no-repair evidence only: no native segment/save-mutation, performance,
host-exit, or complete-monolith-retirement claim follows. The dated gate
counts above remain historical evidence.

## 2026-08-23 boundary-audit disposition

At committed HEAD `4e84abdf`, the Cargo metadata snapshot contains 64 workspace
packages, 239 internal dependency declarations, and 13 ordered policy debts.
Those figures are inventory evidence only: the complete
`tools/check_crate_boundaries.py` run is not green and reports 20 focused
Pages table-lock flat-alias findings: 18 findings for the six compatibility
aliases introduced by `093b82f40` (three public sites each), plus two
pre-existing findings for `TableLockState` and the shared `TableSelector`
export. This disposition changes no policy edge, debt item, ownership
assignment, or verification gate.

## 2026-08-23 follow-up: aggregate comment-scan scope

Committed HEAD `22e8c58ca` only changes where the bounded Numbers comment-cell
scan-work counter is accumulated: the counter now spans the model/cell scan
instead of resetting for each table. The change does not alter Cargo metadata
or the boundary-checker topology, and it does not turn the historical
`4e84abdf` 20-finding audit above into a current-wave gate. No policy edge,
debt item, ownership assignment, or verification result changes here.

## 2026-08-23 amendment: typed Pages footnote limit preservation

The committed `33656a7f0` change keeps the Pages footnote rewrite seam's
resource failures typed. `rewrite_custom_mark_wire` now checks the
selected/after field arithmetic instead of saturating; overflow is reported
as `WireFields`. Candidate verification runs `native_footnotes` before
semantic readback so aggregate text and entry ceilings can surface as
`TextBytes` and `Entries`; `map_package_error_with_kind` maps `TextTooLarge`
and `TooManyBodyStorages` into those categories. The focused source test
`semantic_package_limits_keep_typed_footnote_error_categories` covers
`TextTooLarge { observed: 9, limit: 8 }` to `TextBytes` and
`TooManyBodyStorages { actual: 5, limit: 4 }` to `Entries`; no test execution
result is claimed here.

This is bounded error classification and checked-arithmetic evidence only.
It does not claim package-wide accounting, allocation or performance
measurements, native round-trip behavior, migration-debt retirement, host
exit, or `litchi-iwa` monolith deletion.

## 2026-08-23 amendment: current HEAD bounded follow-up

The prior dated evidence and scope notes remain historical as written. At
committed HEAD `4f2c14484` (parent `609d44049`), the post-`b208dbd77` changes
are bounded source, test, deprecation, and fuzz-harness hardening:

- `d131176b4`, `94b07ffc5`, `03cee9a79`, and `7a497d1ef` retain and ratchet the archive
  production Buffa lazy-view boundary by rejecting generated eager helpers;
- `6b705d39b`, `f8060a43f`, and `447cb3ec3` keep duplicate unselected/shared
  Keynote slide owners rejected before read or staging and add focused
  coverage;
- `33656a7f0` and `8254df4e9` preserve typed Pages footnote wire and
  semantic/rewrite limit categories;
- `cd2dc0ca5` deprecates only the legacy raw-ID Pages section-clear API;
- `d3b1bd571` and `5abf10bb4` add mixed reorder/replacement cache-publication
  coverage;
- `95356f2a8` rejects case-insensitive worksheet-name duplicates;
- `543f74a37` makes Numbers table-header growth fallible per iterator item;
- `cca10f240` redacts native reference identifiers from diagnostic formatting;
- `7effedb57` adds a deprecation boundary to legacy raw-ID Numbers
  cell-comment APIs while retaining compatibility and migration-host paths;
- `609d44049` fixes BIFF8 supplementary-character length accounting and adds
  validation-string round-trip coverage;
- `4f2c14484` charges nested table-lock references before declaration-map
  reservation and adds a low-ceiling budget case;
- `b5a9d5994` bounds CFB fuzz input to 4 MiB, while `1446dc8e7` bounds RTF
  fuzz inputs to 1 MiB and parser/opaque work; both supply campaign recipes
  only.

No test/build/Cargo execution, sanitizer campaign, native application
result, performance/allocation measurement, dependency-edge or ordered-debt
change, host-exit, or monolith-deletion result is admitted by this amendment.

## 2026-08-23 amendment: current HEAD truth audit

The preceding `4f2c14484` note is historical. At committed HEAD
`51ca8eae7` (parent `9366d0a00`), the intervening source changes remain narrow:

- `b316320c9` marks the `litchi-iwa::raw` compatibility facade deprecated and
  makes the boundary checker require that marker. Its policy update only
  clarifies ordered debt 9's reason and exit condition; it removes no edge or
  debt item.
- `e4533c17d` lets focused and migration-host Keynote chart-title setters take
  `AsRef<str>`. `ChartTitleEdit::set` still copies through the existing
  64-MiB visible-title ceiling before staging. The owned-input cases are
  source-level tests, not an executed test gate.
- `9366d0a00` only removes needless `&Cell` borrows in XLS test calls.
- `51ca8eae7` changes Pages footnote edit publication to use fallible exact
  reservations for staged text and custom-mark copies and adds typed mapping
  cases. This is edit-local allocation hardening, not package-wide accounting
  or a peak-memory bound.

The archived HEAD metadata still contains 64 workspace packages, 239 internal
dependency declarations, and 13 ordered debts. The complete boundary checker
is not green at this HEAD: it reports 22 findings (two Keynote
slide-transition validator exports, 18 Pages table-lock alias sites, and the
two existing `TableLockState`/`TableSelector` aliases). These are inventory
and source-audit observations only. No test/build/Cargo execution, native
application result, performance measurement, Buffa/Prost or generated-schema
boundary change, dependency-edge/debt retirement, host exit, or monolith
deletion follows from these commits.

## 2026-08-23 amendment: current HEAD extension after the truth audit

The preceding `51ca8eae7` audit is historical. At current committed HEAD
`ff782f2f5` (parent `dab07a09f`), `dab07a09f` extends the existing bounded
Keynote chart-graph ownership scan to count both non-style owners and title
stand-ins, rejecting a shared title stand-in before read or edit and adding a
focused malformed-graph case. It changes no chart-title wire authority or
package-wide limit profile. `ff782f2f5` marks the three legacy
`KeynoteSlideInfo` native text-storage ID fields deprecated and adds a source
ratchet; the migration-host compatibility value remains available.

These are source-level ownership/deprecation changes. No test command, native
application result, performance measurement, Buffa/Prost projection change,
dependency-edge/debt retirement, host exit, or monolith deletion is admitted.

## 2026-08-23 amendment: current HEAD extension after cache and stale-metadata fixes

The preceding `ff782f2f5` note is historical. At current committed HEAD
`8cddd44af`, `09286432e` charges the empty body-table metadata inventory
before declaration reservation and rejects stale field-9 table declarations
as invalid source rather than treating them as an ordinary selector miss.
The focused and integration cases are source-level coverage only.
`8cddd44af` adds a test-only moved/replaced-entry cache invalidation case; it
does not change production cache behavior. These changes do not establish a
package-wide memory, latency, or performance result.

The current Cargo inventory remains 64 workspace packages, 239 internal
declarations, and 13 ordered debts. The complete boundary checker remains
non-green with 22 findings (two Keynote slide-transition validator exports,
18 Pages table-lock alias sites, and the two existing
`TableLockState`/`TableSelector` aliases). No test execution, native result,
Buffa/Prost or generated-schema change, dependency-edge/debt retirement, host
exit, or monolith deletion follows.

## 2026-08-23 amendment: current HEAD extension after deprecation and alias guards

The preceding `8cddd44af` note is historical. At the then-current committed
HEAD `546df8f04`, `406cff111` adds source ratchets requiring the retained
Numbers cell-comment and Pages section-text compatibility methods to stay
deprecated;
it adds boundary unit cases but changes no runtime format path. `546df8f04`
rejects aliased Keynote slide/show/node identities before transition reads or
staging and adds malformed-topology coverage. The existing bounded owner-work
profile is retained; no package-wide or performance result is implied.

No test command, native application result, Buffa/Prost or generated-schema
change, dependency-edge/debt retirement, host exit, or monolith deletion is
admitted by these source and boundary-ratchet changes.

## 2026-08-23 amendment: Wave45 source-only follow-up at HEAD 97972ce74

The preceding `546df8f04` note is historical. At committed HEAD
`97972ce74` (parent `875fe2a48`), the intervening Wave45 changes remain
bounded source hardening and staged budget work:

- `6d7f1a701` repairs the Pages native-message provenance guard markers and
  adds a direct guard check. The supplied direct Pages guard evidence is
  `16/16`; this is scoped guard evidence, not a full workspace gate.
- `1bad9156e` tightens the strict Numbers TableDataList/Segment route probe:
  known repeated entry/segment fields must carry length-delimited wire shape
  before a candidate is admitted to full decode. This is source-level wire
  hardening; no test execution is recorded here.
- `875fe2a48` stages private `BundleExtractionLimits`,
  `BundleExtractionBudget`, `BundleExtractionContext`, and
  `BundleExtractionTransaction` for cumulative fields/work/text/output
  charging and atomic publication. The seam is explicitly `dead_code`-allowed
  pending neutral-decoder cutover and does not widen the 6005/6201/6011
  compatibility routes. Review was limited to rustfmt/diff; Cargo was blocked
  before this repair, so no Cargo or test result is claimed.
- `97972ce74` repairs the Keynote soundtrack projection digest and adds a
  scalar provenance guard. The supplied soundtrack-focused `cargo check`
  passes; the full `litchi-iwa-protos` test remains blocked by the
  `table_info` compile failure.

These changes establish no native or type-6011 result, full-workspace
verification, Buffa/Prost exit, generated-schema retirement, dependency/debt
change, host exit, monolith deletion, or performance measurement.

## 2026-08-24 amendment: accepted bounded follow-up at `2bbf3c64`

The historical tails above are retained. The accepted chain from `7468fdfc4`
through `2bbf3c64` remains narrowly scoped:

- `7468fdfc4` makes an exact-byte `EntryStore::replace_data` replacement a
  copy-on-write no-op and its focused case asserts shared state and payload
  `Arc` identity (`Arc::ptr_eq`); it adds no wire or package-wide memory
  measurement.
- `322032c10` deprecates the three legacy raw-ID Keynote chart-title methods
  and adds boundary source ratchets requiring the typed selector methods;
  the compatibility methods remain available.
- `031678a31` covers the strict alternate type-`6201` TableDataList route,
  rejects a non-canonical duplicate scalar before public projection, and
  explicitly uses only synthetic type-`6011` segment fixtures; it is not
  native Segment evidence.
- `2bbf3c64` shares one `FootnoteSemanticBudget` across source and candidate
  projections, with the exact-cap case rejecting observed `19` at limit `14`;
  fallible reservations remain local and this is not package-wide accounting.

The current lightweight checks are `git diff --check` and the boundary-policy
Python suite (`python3 -m unittest tools.test_check_crate_boundaries`, 314
tests). No full-workspace Cargo gate is claimed: the surrounding checkout has
a dirty/staged snapshot and `Cargo.lock` is not tracked by this parent. The
clean archived Buffa/protos reports remain historical only (33 focused
protocol tests and a full `litchi-iwa-protos` run at 267 passed/0 failed; a
later strict Numbers/Buffa report recorded 237 passing protos tests). None of
these four commits changes generated schema ownership or removes Prost.

Prior reports supply only historical native no-repair and fuzz-smoke evidence:
the Numbers comment/formula/rich-text reopen recorded 24 type-6005 and zero
type-6011 entries, and the TableDataList/CommentStorage smokes recorded 100
ASan runs with 11/20 seeds. They do not establish native type-6011 support or
a current campaign. No performance/allocation result, Prost removal,
host-edge or ordered-debt retirement, migration-host exit, or monolith
deletion follows from this amendment.

## 2026-08-23 amendment: Wave52 Numbers formula projection (bounded resource record)

The preceding dated evidence and scope notes remain historical. The current
Numbers formula slice moves the production fallback away from the per-cell
generated `FormulaArchive` graph. A bounded owned copy of the admitted formula
source (`FormulaArchiveBytes`) remains the exact-byte preservation authority;
the production path does not re-encode that source. Strict schema-directed
preflight charges the bounded bytes, fields, nesting, work, and
repeated-entry/node counts before a private Buffa lazy per-node/event stream
is admitted. The stream is forced for every selected node or event used by
rendering. It does not materialize generated repeated fields or a generated
`LazyRepeatedView`, and it does not call `to_owned_message`.

Within the `litchi-numbers` formula path, the canonical Prost-generated
builders and decoder remain a `cfg(test)`-gated oracle for differential
fixture construction and parity checks. Commit `2c538d496` moves the direct
`litchi-numbers` `prost` dependency from normal dependencies to
dev-dependencies; it does not remove Prost from the workspace or from other
owners.

The exact isolated resource record used a detached snapshot plus a temporary,
uncommitted bypass of the unrelated Pages provenance guard. The formula codec
suite passed 36/36, the direct compatibility-renderer suite 7/7, the raw
formula-envelope suite 9/9, and the scalar/fallback suite 9/9.
`cargo check -p litchi-iwa-protos -p litchi-numbers --all-targets` passed.
Strict `litchi-iwa-protos` library Clippy passed; strict `litchi-numbers`
library Clippy stopped only at two untouched deprecated `object_count` calls
and one untouched `manual_contains` finding. The boundary-policy suite passed
319/319, with the tracked audit retaining 20 Pages findings and zero Numbers
findings.

Numbers 14.4 opened the read-only native fixture
`/private/tmp/litchi-wave52-formula-native.2dgQQv/source.numbers` without a
repair warning. The 1,200 by 8 table exposed 2,398 formulas, result `5` at E2,
and the intentional division-by-zero at F2; inspecting and cancelling the
formula editor left SHA-256
`e0fb395b0e819583f14d72f1d4916ce482ff35dce9de895d3dfd26869888bbc9`
unchanged. The Rust `read_numbers` example also reported one rooted sheet, one
1,200 by 8 table with 9,600 materialized cells, and one compatibility table,
with the same post-read hash. This is bounded reader and application-acceptance
evidence only: no exact RSS, allocation, latency, performance, Rust/native
formula parity, native save/reopen mutation, publication change, host exit,
debt retirement, or monolith deletion is claimed.

## 2026-08-24 amendment: Wave53 Keynote slide-background resource record

Commit `0b12df5e1` applies the resource contract to the focused Keynote
slide-background seam. The borrowed codec charges aggregate source scanning,
selected fields, nesting, projection/iterator work, typed rewrite work,
readback, and exact output size before the corresponding allocation. Fallible
reservations cover handwritten output buffers and metadata/style rewrites;
failed limit checks publish no partial package. Unknown source bytes are
retained only when the selected typed rewrite can prove that its framing and
nested fields remain safe. This is bounded operation-local accounting, not a
package-wide peak-memory, allocation, latency, or throughput measurement.

The focused manifest topology is unchanged. `litchi-keynote` keeps its direct
`prost` edge in `dev-dependencies` for test-only compatibility oracles;
`litchi-iwa-protos` retains normal `prost` because other generated owners still
use it; and the earlier `litchi-numbers` formula migration retains its direct
`prost` edge only in `dev-dependencies`. Wave53 introduces no normal
dependency edge and does not claim workspace-wide Prost freedom or generated
schema retirement.

The frozen focused resource evidence is recorded in ADR 0008: 25/25 focused
Keynote package cases, 19/19 background-codec cases, 13/13 package-metadata
codec cases, and 9/9 migration-host background cases. Strict Keynote and
protos Clippy passed; the three-crate all-target check passed with the
checkout's existing warnings. These results establish bounded correctness and
limit behavior for this seam, not a general performance result.

## 2026-08-24 amendment: Wave54 Numbers table-cell storage reader resource record

Commit `b7f720872` moves the three read-only Numbers table-cell storage
helpers `resolve_table_string_values`, `table_data_list_has_entries`, and
`rich_text_payload_entry_count` off owned generated `TableDataList` and
`TableDataListSegment` graphs. The doc-hidden
`litchi_iwa_protos::numbers_table_cell_storage_codec` now supplies borrowed
root/segment snapshots and strict streaming visitors. Its private Buffa lazy
views are forced only as parity checks; generated repeated storage does not
escape. Generated messages remain in the unchanged mutation/writer paths and
test oracles.

Each admitted payload has explicit source-byte, field, work, nesting,
reference, and selected-text limits. The host merges successful reports into
one operation-local field/work/reference/text budget across root candidates,
ordered segments, and rich-text package scans. Failed resource checks stop the
operation; ordinary malformed candidates conservatively charge their complete
source length before compatibility selection continues. Visitor output is
staged with fallible reservations, and no partial strings, counts, or package
mutation are published after a later segment failure.

This is a finite reader-resource contract, not a measurement of peak memory,
allocation count, latency, throughput, or package-wide performance. No
manifest edge changes, generated-schema retirement, or workspace-wide Prost
claim follow from this reader cutover.

## 2026-08-24 amendment: Wave55 PackageMetadata registry reader resource record

Commit `02eeb3acc29801838e3981236fd1e41ecbc44fee` moves only the three private
`litchi-iwa` PackageMetadata registry queries
`component_identifier_for_entry` (157 non-definition callers),
`component_identifier_for_object_uuid` (14), and `component_uuid_identifiers`
(41), 209 callers in total, through the doc-hidden
`litchi_iwa_protos::package_metadata_codec` borrowed visitor. Mutation, writer,
allocator, save-token, and data-reference routes remain on generated Prost
messages.

The reader derives package/effective-message limits and `WireLimits` for input,
output, fields, work, nesting, components, and references. A strict two-pass
preflight/visitor charges those finite resources before observations are
accepted. Visitor results stay borrowed; the UUID query uses fallible
`HashSet` staging, and byte, field, work, nesting, and allocation failures are
typed and mapped without publishing partial observations. This is a bounded
resource contract, not a performance, peak-RSS, or allocation measurement.

Admission is intentionally narrower than permissive Prost decoding for the
selected projection: selected fields require canonical singular required and
nonzero values, valid UTF-8, canonical booleans, complete UUIDs, and valid
nested references. Malformed versioned records therefore fail even when the
three query results would otherwise ignore them. The codec validates only this
projection, not the complete `PackageMetadata` schema; unknown noncanonical
keys or lengths are rejected, while unknown noncanonical varint values and
balanced unknown groups are accepted. No manifest/public-API, package-owner,
workspace Prost-retirement, host, or monolith claim follows from this reader
cutover.

## 2026-08-24 amendment: Wave56 Numbers exact-alias resource record

Commit `e4952f7eeaa196263a28fac6cf86510c3e11a2f6` admits only
source-authoritatively exact copies of one Numbers object identifier in
distinct physical components. `Index::from_components` counts every physical
object before checking `max_objects` and before reserving locator/type storage.
It then sorts the physical locators deterministically, rejects duplicate
identifiers within one component, compares cross-component candidates, and
coalesces an admitted alias group to one logical locator and one primary-type
entry. `Package::object_count` continues to report the physical count.

`ArchiveObject::same_content_ignoring_offsets` performs the alias comparison
without allocation. It compares decoded `ArchiveInfo`, exact raw messages,
header and payload lengths, and retained original and original-canonical
header bytes; only component-relative header and payload offsets are ignored.
Consequently, payload, metadata, raw-header, or framing divergence fails
closed even when a neutral decoded projection would otherwise compare equal.

Lookup comparison work is based on the deduplicated logical locator set,
while index allocation, population, sorting, and comparison topology are
charged from the physical count. The already-admitted component/message bytes
remain part of package ingress and candidate-reopen accounting. This is a
finite operation-local accounting contract, not a measurement or claim for
package-wide peak RSS, allocation count, latency, or throughput. The slice
changes no manifest edge and retires no Prost-generated schema or workspace
dependency.

## 2026-08-24 amendment: Wave57 Numbers name-publication resource record

Commit `35c2ae281d40bc16c659c30dd3690a702a750ea0` extends the bounded Numbers
name-publication operation through the PackageMetadata save-token sidecar.
Follow-up accounting hardening is commit
`ba57356166e82ff37ee1cd5223b68acd484aac42`. It removes detached, unmetered
candidate-size `Budget` work: candidate sizing and selector comparisons now
share the reported `Budget`, a test-only aggregate proves every successful
`Budget` charge equals `RewriteReport.work_bytes`, and max-minus-one fails
before candidate allocation.
Source scanning charges physical package entries, the metadata member,
identifier/effective-locator matching, fields, nesting depth, work, selected
component count, and output bytes. The operation charges native rewrite and
metadata rewrite work together and performs exact output sizing, fallible
reservation, and candidate verification before publication. A candidate
allocation is made only after the relevant preflight checks succeed.

The save-token selector count is bounded by `RewriteOptions.max_components`;
the report keeps `additions = 0` because token updates are not registry-object
additions. Root token advancement and selected current-component updates are
therefore accounted as operation-local rewrite work without overloading the
registry-addition counter. Unknown root/component fields are copied from the
source, while selected known fields are replaced or canonically appended only
after the full source and selector checks pass. This is a finite resource
contract for this operation, not a package-wide memory, allocation, RSS,
latency, throughput, or performance claim. Before native rewrite/allocation,
the Names caller also precharges decompressed type-11006 payload bytes times
the changed semantic-operation upper bound for visitor locator matching;
compressed Snappy bytes are separate publication accounting.

No manifest edge changed. This amendment does not retire generated schemas or
Prost ownership: the hidden codec uses the existing Buffa projection as a
parity boundary, while unrelated generated mutation/writer paths and normal
workspace Prost owners remain. It makes no workspace-wide dependency-removal
claim.

## 2026-08-24 amendment: Wave60 Numbers root comment-clear resource record

The Wave60 implementation series culminates in commit
`895ef17848516cf201e717239c44c119eae5da87`. The changed-clear transaction
keeps the existing package, archive, semantic-reference, wire-byte, field,
nesting, work, materialized-cell, text, and output ceilings. Selected tile,
root-list, comment-storage, PackageMetadata, and archive-header operations use
their explicit finite codec profiles; fallible reservations precede retained
census collections and publication artifacts.

The global ownership pass charges complete BNC cell payload bytes before the
handwritten cell parser runs. Archive-owner inspection accumulates header work
and reference occurrences across every physical component instead of resetting
the effective budget at each object. The comment graph census applies one
aggregate semantic-reference ceiling before allocating its lookup sets, and
source identity, reply, list, storage, cell-key, and UUID checks use bounded
`HashSet` membership rather than repeated quadratic scans. Candidate archive
mutation validates canonical object framing, rejects multi-message storage
objects, preserves supported raw headers/unknown fields, and refuses opaque
future owners before deletion.

The PackageMetadata codecs preflight source scanning, exact output sizing,
selected-field work, candidate verification, and root-map/combined-removal
support before their output allocation. The current comment-clear route uses
the dedicated ownership visitor and save-token rewrite because an admitted
storage object must have no metadata ownership; it does not claim to delete an
arbitrary registered resource graph. Native component compression, Metadata
rewriting, ZIP reassembly, and candidate reopen remain separately staged
fallible buffers under their existing package/archive ceilings. This amendment
therefore records bounded operation-local work, not a single-allocation design
or a measurement of peak memory, allocation count, latency, throughput, or
RSS.

No manifest edge changed. Generated schemas and normal Prost owners elsewhere
remain, and this resource record makes no package-wide performance,
workspace-wide dependency-removal, or publication-completeness claim.

## 2026-08-24 amendment: Wave61 Pages body-table lock resource record

Commit `ca3fbd21a24f7195ef9b2d8d169e286339fc274e` strengthens the finite
resource contract for the selector-first Pages body-table lock transaction.
Before Snappy decompression or `Archive::parse` can allocate the selected
component, the owner derives the complete decoded archive extent and physical
object/message inventory from the already parsed source catalog, charges that
inventory to the shared operation budget, and then cross-checks the actual
decompressed length. Canonical archive-object framing is still validated
after parse.

Name selection charges the compared source and selector bytes for every
candidate table; position selection charges the inspected prefix. Exact
source/patch and no-op artifact comparisons now charge equal-length byte
scans instead of performing unreported linear work. These charges join the
existing finite source, package-entry, payload, field, nesting, reference,
object/message, output, compression, ZIP reassembly, and candidate-reopen
ceilings. Overflow and limit failures remain typed and occur before the
selected mutation is published.

This is an operation-local preflight improvement, not a measurement of peak
RSS, allocation count, latency, or throughput. It does not establish a
package-wide allocation census or claim that every generic ZIP/catalog
candidate allocation has moved behind one common reservation. No manifest,
generated-schema, Buffa, or Prost ownership changed, and no workspace-wide
performance or dependency-removal claim follows.

## 2026-08-24 amendment: Wave62 Pages body-table title resource record

The Wave62 ownership and hardening series is commits
`48f203aae56e43133fd931accfa6558661594ba0`,
`a92f8f11a50c709877b8d1f0a158da72114dc4ab`,
`a7088be4dd9fde9b2e843473b093839e7b7629b3`, and
`a1c1e83a3edad808bacefa648fe0c1bdd53308f5`. The title transaction reuses the
body-table lock `WireBudget` for one operation-local accounting envelope.
Before the editable selected archive is allocated, it charges the source
catalog, selected compressed member, decoded archive extent, and physical
object/message inventory. Name and position selection, rooted graph
inspection, reference metadata, and cross-component title-style scans are
charged through the same budget; the final style scans are allocation-free.

The strict title codec reports and charges selected wire bytes, fields,
nesting, work, and references. The package precharges rewritten payload,
compression, ZIP-entry, complete package, and candidate verification bounds.
Candidate bytes are charged before reopening the candidate `SourceCatalog`,
and verification continues with the same transaction budget. Exact output
sizing and fallible reservations remain the allocation boundary for the
focused codec and publication artifacts; overflow and limit failures are
typed and occur before publication.

This is a finite operation-local resource contract, not a measurement or
claim about package-wide peak memory, allocation count, RSS, latency,
throughput, or performance. Generic ZIP/catalog construction still has its
own bounded internal allocations; this amendment does not claim one global
allocation for every candidate stage. No manifest edge, generated schema,
Buffa owner, or Prost owner changed, and no workspace-wide dependency-removal
claim follows.

## 2026-08-24 amendment: Wave63 Pages body-table header resource record

The Wave63 resource owner culminates in `7bc68903e`, following the shared
semantics, strict codec, focused package, host cutover, boundary, and
dependency-neutral edits `6f35a3d88`, `dd9e7f25c`, `58f04cd90`, `2bf7598d7`,
`c93598d22`, `7261dd1c4`, and `30207fa47`.

The hidden neutral `table_header_settings_codec` derives finite source-byte,
field, work, nesting, and candidate-output ceilings. Strict preflight validates
the selected scalar projection, measures exact rewritten payload size before
output allocation, performs one fallible output reservation, preserves unknown
source spans, and verifies candidate readback. Package reference/dependency
scans charge aggregate and optional field-local references before ownership
decisions; non-local, external, data-reference, duplicate, malformed, and
unsupported group/dependency cases are rejected rather than normalized.

The Pages package joins those charges to source-catalog/member bytes, decoded
archive extent, physical object/message inventory, selection, compression,
root-preview deletion, ZIP reassembly, and candidate-reopen ceilings.
Candidate reservation and verification remain operation-local and fallible;
generic ZIP/catalog staging retains its own bounded allocations. This is a
finite resource contract, not a package-wide peak-memory, allocation-count,
RSS, latency, throughput, or performance claim. No manifest edge, generated
schema, Buffa owner, or normal Prost owner changed.

## 2026-08-24 amendment: Wave64 Keynote chart-caption resource record

Commit `514b82bdf658d78ea0154f4fc075b1b85488f31c` reuses the finite
Keynote package, semantic, archive, ZIP, and wire ceilings for existing-caption
replacement. Caption input is capped at 64 MiB and copied only after a
fallible exact `String` reservation. Chart-caption and caption-info projections
receive payload-sized `DecodeOptions` bounded by the package field, nesting,
and rewrite-work limits. The selected text storage uses the existing strict
text-wire rewrite limits, exact output sizing, fallible reservations, archive
framing checks, Snappy compression limits, package-output ceiling, and complete
candidate reopen.

The ownership census visits the already prepared finite package object,
message, and reference inventory and retains no unbounded caller-controlled
graph collection. Exact patch artifacts remain bounded by the package input
and output limits; unchanged commits reuse the exact source artifact. Preview
deletion and storage/node component rewrites are counted in transaction
diagnostics, but the public diagnostics do not expose native identifiers.

These are operation-local bounds inherited from the focused package and codec
seams, not a claim that every comparison is represented by one aggregate
counter or that every candidate stage uses a single allocation. No latency,
throughput, peak-memory, RSS, or package-wide allocation measurement was made.
No manifest edge, generated schema, Buffa owner, normal Prost owner, or
workspace-wide dependency changed.

## 2026-08-24 amendment: Wave65 Keynote chart-caption graph resource record

Commit `f0bbe079b094b3652751f6ab7c89b2dc64fac6a9` extends the focused
chart-caption resource contract from text replacement to the canonical graph
lifecycle. The graph codec, Metadata codec, archive parser, Snappy path, and
package owner retain finite source, payload, field, nesting, work, reference,
object, text, output, and candidate-reopen ceilings. Creation and stand-in
removal account for graph-object additions, UUID/Metadata registration,
selected-component save-token work, archive edits, preview invalidation,
compression, ZIP reassembly, and candidate verification.

The graph and metadata codecs perform bounded preflight, exact selected
payload sizing, and fallible reservations for their own outputs. The package
also bounds physical package members, decoded archive objects/messages, graph
ownership census, and Metadata scans; unknown source spans are retained and
ambiguous or unsupported ownership is rejected. Exact artifacts are reused
for no-ops and retained for inverse application.

This implementation does not claim one aggregate transaction budget across
graph encoding, Metadata rewriting, archive serialization, Snappy
compression, ZIP reassembly, candidate reopening, and `ExactArtifacts`.
Those stages remain separately bounded and fallible. Consequently this is an
operation-local resource record, not a package-wide allocation, peak-memory,
RSS, latency, throughput, or performance measurement. No manifest edge,
generated schema, Buffa owner, or normal Prost owner changed.

## 2026-08-24 amendment: Wave66 Keynote chart-caption hardening resource record

Commit `cd76c394e2cd6e360cd01e8bd61cb7594d837701` keeps the Wave65
chart-caption lifecycle while adding strict source-authority work at its
selected boundaries. Canonical IWA framing checks, chart/reference/theme
wire scans, physical dependency attribution, Metadata component and external-
reference inspection, save-token rewriting, MessageInfo reference transition,
candidate reassembly, and complete reopen all remain under the existing
finite package, archive, wire, object, field, nesting, reference, work, and
output ceilings.

The strict Metadata and chart-caption codecs measure their selected outputs
before fallible allocation and verify candidate readback. The package scans a
finite physical catalog, rejects ambiguous owners instead of retaining an
unbounded graph, and preserves exact source/target artifacts for inverse
application. Some-to-some replacement adds the Metadata sidecar as one
authorized changed component; graph transitions authorize aggregate and
field-local reference lists without exposing raw identifiers publicly.

Wave66 still does not provide one aggregate transaction counter or a single
candidate allocation across archive parsing, dependency census, Metadata
inspection/rewrite, Snappy compression, ZIP reassembly, candidate reopen, and
exact artifacts. Each stage remains bounded and fallible, but this is not a
package-wide allocation, peak-memory, RSS, latency, throughput, or performance
claim. No manifest edge, generated schema, Buffa owner, normal Prost owner, or
workspace dependency changed.

## 2026-08-24 amendment: Wave67 Keynote chart-caption aggregate transaction-budget resource record

Commit `688ddbb0bd6ca4e7addc15ac8fbd12b0032fc5a9` replaces the
Wave66 collection of independently bounded chart-caption stages with one
private aggregate `CaptionBudget` for the focused transaction. The budget
charges the finite package catalog and ownership census, strict graph and
Metadata codec reports, archive serialization, Snappy input and maximum
compressed extent, ZIP reassembly planning and execution, complete candidate
reopens, and exact source/target artifacts. Checked counters cover input,
intermediate and final output, fields, work, nesting, components, references,
objects, messages, additions, allocations, retained bytes, and scratch bytes;
overflow fails closed.

The supporting seams are output-free where publication needs a global plan.
`Archive::encoded_len_with_limits` measures exact source-authoritative IWA
serialization, `SnappyStream::maximum_compressed_len` provides a bounded
compression ceiling, and `PreparedReassembly` exposes exact ZIP execution
requirements before final allocation. The strict PackageMetadata save-token
and addition-plus-save-token routes now expose prepare reports and execution
requirements; prepare performs source scan, sizing, precharge, and semantic
validation without allocating the candidate output, while execute enforces
the residual limits and performs the single codec output reservation.

Existing-caption replacement still uses two private staged candidates: the
native text candidate followed by the Metadata save-token candidate. Both
stages and both allocations are charged before publication, and a later
failure leaves the package source unchanged, but Wave67 does not claim one
candidate allocation for the whole transaction. The aggregate intermediate-
output ceiling is transaction-local and intentionally larger than the final
physical package ceiling; member and final artifact limits are enforced again
by reassembly and complete reopen.

This is a focused chart-caption transaction contract, not a package-wide
peak-memory, allocation-count, RSS, latency, throughput, or performance
measurement. No manifest edge, generated schema, Buffa owner, normal Prost
owner, or workspace dependency changed.

## 2026-08-24 amendment: Wave69 shared chart-caption edge resource record

Commit `cd382f95dad52fecdfdb17a7fbb1565c489e6de1` gives the shared
chart-caption edge finite input, output, field, work, and nesting ceilings.
The host helper derives those ceilings from the owning package's finite IWA
stream and archive-message limits and maps codec byte, output, field, work,
nesting, and allocation failures back to the existing typed host error
surface.

The strict codec meters source projection, selected-field validation, exact
output sizing, raw emission, and candidate readback. Rewrite preflight covers
the complete emission and verification passes before the single exact output
reservation. Its reports expose measured input/output bytes, fields, work,
maximum depth, allocation, retained/scratch, and changed-state facts; direct
tests prove exact ceilings and max-minus-one output, field, work, and depth
failure before output allocation.

These are operation-local codec and helper bounds. Wave69 does not establish
one aggregate transaction budget across chart graph ownership, archive
serialization, Snappy compression, ZIP reassembly, Metadata, or package
publication, and it makes no peak-memory, RSS, latency, throughput, or
zero-copy claim. No manifest dependency, generated schema, Buffa owner,
normal Prost owner, or workspace dependency changed.

## 2026-08-24 amendment: Wave70 Pages drawable-order resource record

Commit `9b44e20ac0d3696aeb28f94d923612157446a83c` gives the hidden
`pages_drawable_order_codec` a bounded source projection and rewrite for the
complete `TP.DrawablesZOrderArchive` repeated `TSP.Reference` field. Strict
preflight charges input and exact output bytes, fields, work, nesting, and
references, including the full source scan and identifier lookup work. The
rewrite performs one output allocation only after exact sizing and all
limits have passed; the host then charges its transactional archive
candidate/readback/reopen path separately under its existing package limits.

The codec keeps complete raw reference records and interleaved root fields as
the preservation authority, including unknown balanced groups and overlong
unknown scalar framing. It rejects noncanonical known keys, lengths, and
values, missing or zero identifiers, duplicates, and non-permutation edits.
The outer IWA object-length prefix may be canonicalized by the existing host
serializer policy; this record does not promise raw outer-prefix preservation.

These are operation-local resource bounds. They are not a package-wide
allocation, peak-memory, RSS, latency, throughput, or zero-copy measurement,
and no aggregate budget across archive serialization, compression, ZIP
reassembly, or native application save is claimed. No manifest, generated
schema, Buffa owner, normal Prost owner, or workspace dependency changed.

## 2026-08-25 amendment: Wave73 Keynote movie-caption resource record

Commit `6ed7ee4e277c78a99f1b779a2b5530d7267beddf` gives the hidden movie-caption
edge codec finite input, exact output, field, work, and nesting ceilings. The
codec charges strict projection, selected-field validation, exact sizing, raw
emission, and candidate readback before its single exact output reservation;
its report includes input/output bytes, fields, work, maximum depth,
allocation, retained/scratch, and change facts.

The admitted `Some -> Some` package mutation reuses the private chart-caption
physical transaction: bounded catalog and storage work, prepared Metadata
save-token execution for every changed current component, archive sizing,
Snappy compression, ZIP reassembly, preview deletion, complete candidate
reopen, and exact locality verification. Split slide/storage and slide-node
components are not undercounted as one semantic target: each changed current
component receives the same new root save token and diagnostics include the
Metadata member.

Selection and the global ownership census remain separately governed by the
package's finite semantic and wire limits rather than one newly claimed
end-to-end allocation counter. This record therefore makes no package-wide
single-allocation, peak-memory, RSS, latency, throughput, zero-copy, or
performance claim. No manifest dependency, generated source schema owner,
normal Prost owner, or workspace dependency edge was removed; the private
Buffa projection remains an implementation detail of `litchi-iwa-protos`.

## 2026-08-25 amendment: Wave71 Pages exact-output resource record

Commit `580a5343a2c75a1c1b185a5cc8aff4a87e2a5c11` makes
`litchi_pages::Package::write_to` the public exact-output seam. The writer
streams the retained ZIP artifact to a caller-owned `Write` sink without
allocating another package-sized output buffer. It tolerates `Interrupted`,
tracks the offset reached by conforming partial writes, and reports zero-write,
over-report, and ordinary sink failures through typed redacted `WriteError`
values carrying `bytes_written`.

The Pages-to-host section bridge now uses one-pass `FocusedCandidateWriter`.
It charges each emitted chunk before retaining it and preserves the existing
`3S+2T` work accounting, so a limit failure does not leave a package-sized
candidate allocation behind. This is an operation-local bridge accounting
record, not a package-wide allocation or peak-memory proof.

`write_to` does not flush or sync its sink and does not rename, atomically
replace, or durably publish a filesystem path. No package-wide output-buffer,
allocation, zero-copy, RSS, latency, throughput, or performance claim follows;
those policies remain with the caller and later publication work. No manifest
edge, generated schema, Buffa owner, normal Prost owner, or workspace
dependency changed.

## 2026-08-25 amendment: Wave72 shared chart-caption resource record

Commit `997e0bf55e8a1381296a4d446f83dd5217a15385` makes every internal
chart-caption codec traversal charge its source scan, sizing, raw emission, and
candidate-verification work. Unknown overlong scalar values and complete
unknown source groups are retained byte-for-byte; their keys and length
framing, selected known fields, and exact nesting depth remain bounded and
strict. Pages and Numbers retarget paths preserve raw `ArchiveInfo` headers and
the exact aggregate/`FieldInfo` reference sets, rejecting shared or
misattributed `CaptionInfo`/storage owners before publication. Keynote merges
the residual codec report into `CaptionBudget` and precharges candidate reopen
work; output and allocation failures retain typed limit/error mapping.

These are operation-local source, work, nesting, output, allocation, and
candidate-accounting facts. They do not establish a package-wide memory,
peak-RSS, latency, throughput, zero-copy, or performance measurement, nor a
single allocation for every format transaction. No manifest edge, generated
schema, Buffa owner, normal Prost owner, or workspace dependency changed.
