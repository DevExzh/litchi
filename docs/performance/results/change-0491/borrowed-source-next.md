# Borrowed source next-implementation contract

Status: design contract only. This records the next production slice for the
genuine non-static borrowed-byte gap found in change-0491. It does not claim
that the current pilot/build/capture artifacts are performance evidence, and
it does not authorize a production implementation in this turn.

## Gap and selected shape

`litchi-core::ReadAt` (`crates/litchi-core/src/source.rs:62`) is the immutable
positional source contract. `SliceSource<'a>` (`source.rs:193`) already adapts
caller-owned bytes without copying and carries a stable `SourceVersion`. The
gap is above that seam: `litchi_opc::SourceBackedPackage`
(`crates/litchi-opc/src/source_backed.rs:5184`) stores an
`Arc<dyn ReadAt>` with an effectively static lifetime, and its
`IndexedArchive<SourceReader>` cannot lend bytes from arbitrary positional
providers. `litchi_docx::source_backed::Package`
(`crates/litchi-docx/src/source_backed.rs:101`) consequently has no scoped
read-only document view over a non-static byte slice.

The first implementation should add a separate scoped path rather than
genericize `SourceBackedPackage`:

```text
litchi_opc::BorrowedSourceBackedPackage<'a>
    owns PhysPkgReader<'a> + the admitted source catalog
litchi_docx::BorrowedPackage<'a>
    owns the OPC view and exposes read-only semantic queries
```

Its first public constructor is `from_bytes(&'a [u8], ReadLimits)`. The
source bytes remain caller-owned for the whole view lifetime. `to_owned` is an
explicit conversion after a complete semantic read; it is the only place this
route may copy the source or decoded payload. The view has no edit, publication,
writer, or source-artifact capability. XLSX and PPTX are later slices, not
implicit scope expansion.

This shape follows ADR 0001 (specialized zero-copy paths are explicit), ADR
0003 (immutable `Send + Sync` snapshots, scoped borrowed views, explicit
conversion), ADR 0005 (stable positional sources, hierarchical limits, lazy
state, and measured claims), ADR 0010 (archive grammar stays below the
facade), ADR 0011 (OPC owns the physical package), and ADR 0024 (borrowed
package inputs remain source-bound and read-only).

## Authoritative code to reuse

The new owner must use the existing `litchi-opc` admission path. The relevant
seams are:

- `crates/soapberry-zip/src/office.rs:7444`,
  `LazyArchiveReader::read_stored_borrowed`, validates local headers, data
  descriptors, declared sizes, layout spans, encryption flags, and CRC before
  returning a source slice. It returns `None` only for a valid but
  borrow-ineligible case such as Deflate, conservative ZIP64/layout metadata,
  or a non-empty zero CRC.
- `crates/litchi-opc/src/pkgreader.rs:25`, `ArchiveAccess`, and
  `read_structural_member` at line 74 already express the structural byte
  policy: validated borrowed Store first, then an existing shared/owned read
  for an eligible non-borrowed member. Reuse this helper unchanged for
  `[Content_Types].xml` and relationship XML.
- `PackageReader::source_catalog` in
  `crates/litchi-opc/src/pkgreader.rs` (around line 1191), including
  `load_part_catalog` (line 972), `classify_part_members` (line 1018),
  `load_rels_lazy`, `ContentTypeMap`, `PartNameIndex`, and
  `NonPartMember`. This is the shared relationship, content-type, part-name,
  duplicate/case-conflict, and non-part admission logic. The validation-only
  twin is `source_catalog_for_validation` at line 1244.
- `crates/litchi-opc/src/phys_pkg.rs:53`, `PhysPkgReader<'data>`, and
  `blob_for_borrowed` at line 361. The latter checks
  `ReadResource::PartBytes` from central metadata before borrowing and does
  not consume `Parts` or aggregate materialization budget because it retains
  no payload allocation.
- `SourceBackedPackage::from_read_at_inner` in
  `crates/litchi-opc/src/source_backed.rs` remains the source-freshness and
  managed-budget reference. Preserve its version/length capture, current
  checks around mandatory work and publication, cancellation checks, catalog
  `Resource::Objects` reservation, and managed payload reservation rules.

Do not call `PackageReader::from_phys_reader` for this view. That constructor
uses `load_parts_eager` and materializes every admitted part into
`Arc<Vec<u8>>`; it is the wrong ownership shape. Do not fork
`classify_part_members`, relationship walking, CRC checks, or ZIP parsing in a
second weaker parser.

## Admission, validation, and refusal contract

The borrowed owner should call `PackageReader::source_catalog` against the
`LazyArchiveReader<'a>` held by `PhysPkgReader<'a>` (the physical reader's
crate-private `archive()` seam is sufficient). Convert each resulting
`DeferredPart` into a private descriptor keyed by its `PackURI` and preserve
its admitted content type and relationships. A descriptor does not need an
`IndexedArchive::EntryId`; payload access goes through the validated physical
reader by `PackURI`.

The catalog's exact existing phases remain authoritative:

1. Check archive-member and relationship-part counts, locate the unique
   content-types member, check `ContentTypesBytes`, and parse it through
   `read_structural_member`.
2. Load package relationships through `load_rels_lazy`, including its
   relationship XML ledger limits.
3. Walk relationship targets and classify all members through
   `load_part_catalog` in `Deferred` mode. Enforce `Parts`, member-name bytes,
   content-type presence, relationship-source rules, duplicate/case-conflict
   names, and non-part policy.
4. Retain only the admitted catalog. Ordinary payloads stay cold at open.

`Deferred` catalog admission intentionally does **not** charge aggregate
ordinary-part bytes. It does not read ordinary payloads or prove their CRC.
At a later part access, run the existing `PhysPkgReader::blob_for_borrowed`
admission for a borrowed Store result; its per-part `PartBytes` check is still
required. If the operation materializes or streams decoded bytes, use the
existing validated `blob_for`, `with_verified_decoded_reader`, or equivalent
source-backed path so `Parts`, `PartBytes`, `TotalPartBytes`, `Memory`,
`Objects`, cancellation, and source freshness are charged at the established
phase. Never claim that catalog-open CRC validation covers a deferred ordinary
payload.

The result type must distinguish a valid but borrow-ineligible member from a
failure. A suitable internal contract is:

```text
borrowed_bytes(part) -> Result<Option<&'a [u8]>>
  Some(bytes): validated Store slice, pointer-identical to the source range
  None: valid member but no borrowable layout/compression/provenance
  Err: missing, malformed, encrypted, invalid CRC/size, stale source, or limit
```

An optional higher-level `read_validated` operation may explicitly choose the
existing bounded owned/stream path after `None`, but it must never catch an
`Err` and retry with guessed offsets. In particular:

- encrypted flags and malformed local headers/descriptors/sizes/CRC remain
  typed ZIP/OPC refusals; they never fall through to owned bytes;
- Deflate may use the existing verified decompression or streaming path, and
  its result must be described as decoded/materialized, never zero-copy;
- archive-level ZIP64, conservative layout metadata, and non-empty zero CRC
  remain `None`/a typed `BorrowedUnavailable` outcome for a borrowed-only
  operation. If a separate explicit owned operation supports that archive,
  it must invoke the existing reader and its limits, not reinterpret the
  physical offsets as Store bytes;
- a missing member remains `PartNotFound`, even though a valid member may
  return `None` from borrowed access.

For mandatory catalog members, retain the existing `read_structural_member`
decision exactly. It propagates typed errors from `read_stored_borrowed`; it
does not add a new catch-and-fallback branch. Its established shared/owned
path for a valid, non-borrowed structural member is catalog behavior and is
not a promise that semantic part access can borrow the member.

## Budget and freshness contract

The borrowed source allocation is caller-owned and must not be double-counted
as newly retained `Memory`. The view's retained catalog, descriptor vectors,
name indexes, relationship values, and physical metadata still consume the
same finite object/allocation policy. Before publishing a managed view,
reserve the catalog object envelope equivalent to the source-backed path:
one package catalog owner plus one unit per admitted physical archive member,
with checked `Resource::Objects` arithmetic. Retain that reservation with the
view and release it on drop.

Borrowed Store access charges the per-member declared `PartBytes` admission,
but no materialization `Parts`, `TotalPartBytes`, or `Memory` charge. A
materialized Deflate/owned result must use the existing weighted cache or an
explicit one-shot callback that has the same reservation and cancellation
semantics. Do not copy `PartCache` into a reduced cache with weaker flight,
eviction, or reservation behavior. A managed result must not expose an
unreserved `Arc<Vec<u8>>` escape analogous to `PartData::into_arc`.

Construct one `SliceSource<'a>` at ingress and capture its stable version and
length before opening. The slice's immutable identity makes its freshness
check constant, but the package must still perform the same phase fences
around catalog and payload publication; a future non-static `ReadAt` provider
must replace those fences with `SourceSnapshot::ensure_current`, not be
smuggled through the static `Arc<dyn ReadAt>` API. A slice lifetime protects
memory safety; it does not waive limit, CRC, or admission checks.

## Implementation sequence

1. Add a private/common OPC catalog owner that accepts a
   `LazyArchiveReader<'a>` through `ArchiveAccess`, calls
   `PackageReader::source_catalog`, and converts `DeferredPart` values into
   borrowed-view descriptors. Keep the existing source-backed and eager
   owners on their current paths.
2. Add `BorrowedSourceBackedPackage<'a>` with `from_bytes` and explicit
   limits/context variants. Retain the physical reader, catalog, object
   reservation, and immutable source identity. Expose only read-only part
   lookup and a borrowed/explicit-owned payload operation.
3. Add focused OPC tests first. They must prove catalog error parity with
   `SourceBackedPackage`, source pointer identity for eligible Store members,
   typed refusal behavior, and budget counters before any DOCX adapter uses
   the new owner.
4. Add `litchi_docx::BorrowedPackage<'a>` as a read-only adapter over the
   borrowed OPC owner. Start with main-document relationship and content-type
   checks plus full-text extraction. Keep semantic output ownership explicit
   where parsing requires it; do not add edit or publication methods.
5. Add an explicit `to_owned` conversion only after the scoped semantic view
   has completed its checks. Reopen or materialize through existing validated
   APIs and preserve all typed failures and limits.
6. Extend other OOXML formats only after the DOCX gates and measurement
   contract pass. Do not genericize `SourceBackedPackage` or expose archive
   implementation types as a shortcut.

## Test gates

The implementation is not ready for production review until all of these pass:

- Lifetime/API: a local `Vec<u8>` can create a view from `&bytes`; the view
  cannot outlive `bytes` in compile-fail coverage; dropping the physical
  reader does not invalidate a returned slice while the source remains alive.
- Catalog parity: valid packages and typed failures match the existing
  source-backed path for missing/duplicate/case-conflicting names, malformed
  relationships, missing content types, untyped referenced parts, junk, and
  relationship-source violations. Exercise archive-member, relationship,
  content-types, member-name, `Parts`, and catalog XML limits.
- Deferred work: opening does not read ordinary payloads. Verify this with
  source range/counter instrumentation and assert that ordinary CRC checks
  happen only when that part is accessed.
- Borrowing/refusal: eligible Store returns source pointer/range identity and
  no cache allocation; Deflate takes the explicit validated decoded path;
  missing, encrypted, bad local headers/descriptors, wrong sizes, bad CRC,
  stale source, and limit failures preserve their typed errors. Cover
  archive-level ZIP64, conservative layouts, and zero CRC without guessed
  slices or an implicit owned retry.
- Budget: catalog objects/indexes are retained and charged; caller-owned
  source bytes are not charged as a second input allocation; borrowed access
  does not consume materialization budgets; explicit decoded access enforces
  `Parts`, `PartBytes`, `TotalPartBytes`, `Memory`, `Objects`, cancellation,
  cache/flight, and eviction behavior.
- DOCX semantics: full-text output, error precedence, and source immutability
  match the existing source-backed oracle. No edit/publication/source-artifact
  method is reachable from the borrowed type. `to_owned` is explicit and its
  copy is observable in accounting.
- Repository gates: `cargo fmt --check`, targeted `cargo test -p
  litchi-core -p litchi-opc -p litchi-docx`, rustdoc/clippy, boundary checks,
  malformed-ZIP regressions, and source-mutation/freshness tests.

## Measurement contract

The change-0491 current pilot/build/capture artifacts are before evidence and
do not establish a speed or memory claim. Compare the existing owned
`SourceBackedPackage` path, the new borrowed `&[u8]` path, FileSource, and
instrumented short/delayed positional providers on the same semantic corpus.
Load each fixture before timing; the borrowed route must contain no ingress
`Vec::to_vec` or equivalent archive copy.

Report p50/p95/p99 latency, allocation count/bytes, peak RSS, logical range
calls and returned bytes, source-copy bytes, borrowed pointer/range identity,
decoded/decompressed bytes, cache/object/memory counters, and exact semantic
digest. Separate these claims: no ingress copy, Store-member zero-copy,
Deflate decoded streaming/materialization, and provider I/O behavior. Include
both Store and ordinary Deflate fixtures, frozen warmups and samples, and
fresh child processes for cold measurements. No hardware/network cold-start
claim follows from an in-memory slice benchmark.

## Unresolved risks

- The public lifetime-aware provider shape for non-slice `ReadAt` sources is
  still open. It must preserve stable identity and freshness without
  propagating a source lifetime through the existing lifetime-free semantic
  handles.
- The current structural helper permits its established validated read path
  after a borrow-ineligible `None`; the borrowed semantic API must keep that
  choice explicit so ZIP64 and zero-CRC policy cannot drift between catalog
  and payload operations.
- A retained borrowed view has no normal decoded cache yet. Choosing a
  one-shot callback versus a weighted borrowed/owned cache requires measured
  reuse patterns and an exact reservation design; a partial cache clone is
  prohibited.
- DOCX semantic readers may need owned strings or temporary XML state even
  when the physical Store bytes are borrowed. The API and measurements must
  report those semantic allocations separately from archive ingress and must
  not broaden the zero-copy claim.
- ZIP64, encrypted packages, and CRC-unverifiable members need fixture-level
  parity before the borrowed view is published. Until those gates are fixed,
  the typed refusal boundary is the safe contract.
