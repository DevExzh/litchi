# ADR 0030: Lazy OPC part decode behind the fallible package accessors

- Status: **Proposed — awaiting human review. Not accepted, not normative, and
  deliberately absent from the accepted table in [README](README.md).** No code
  may cite this record as authority until a human accepts it.
- Date: 2026-09-15
- Supersedes: nothing. Amends: nothing.
- Raised by: [change 0610](../performance/0610-opc-lazy-part-decode-design.md),
  which carries the site enumeration, the predicted retention and the sizing
  probe. Gate 1 of [change 0581](../performance/0581-opc-package-retention.md)
  requires this record to exist before candidate C2 or C2′ may be implemented.

## Context

`OpcPackage::open` decompresses and retains **every** admitted part before any
caller has asked for one. `PackageReader::load_parts_eager`
(`crates/litchi-opc/src/pkgreader.rs:852`) inflates the whole admitted set in a
single bulk read (`:913`) and hands each payload to `PartFactory::load_shared`,
which stores it as `Arc<Vec<u8>>` inside an `XmlPart` or a `BlobPart`
(`crates/litchi-opc/src/part.rs:269`, `:167`). The package therefore holds the
compressed source archive *and* the sum of every decompressed payload at the
same time.

Change 0581 priced that: 2.00× to 11.19× the archive across a six-fixture real
corpus, 254× the bounded path's peak on a 64 MiB media fixture and 506× at
128 MiB, and it established that every documented DOCX, XLSX, PPTX and XLSB
open-then-save example reaches this path. It declined to implement, for one
reason: `Part::blob(&self) -> &[u8]` is public, infallible and borrowing, so a
lazy part behind it has no error channel, and deferring the decode would move
the `ReadResource::PartBytes` and `ReadResource::TotalPartBytes` refusals out of
`open()`. It froze two candidates — C2, a fallible `Part::blob`, and C3, routing
the format crates onto `SourceBackedPackage` — and six admission gates.

Change 0610 measured how much of the eager decode is wasted, on the real fixture
corpus:

- The eager decode inflates **5,077 parts and 45,562,463 bytes** across 334
  OOXML fixtures whose archives total 15,323,729 bytes (2.97×).
- An `OpcPackage` open-then-edit-one-part-then-publish needs the payload of
  **zero** other parts: across 325 fixtures and 4,896 parts, ablating any other
  part's payload leaves the published output unchanged outside that part's own
  member.
- The documented XLSX editor route — `Workbook::open` → `edit()` → hide a tab →
  `commit()` → `to_bytes()` — needs **exactly one** part decoded,
  `/xl/workbook.xml`, on all 33 corpus fixtures that admit the operation: 33 of
  550 parts and 60,239 of 8,274,037 inflated bytes, **0.73%**.

The remaining 99.27% is decompressed, charged, retained until the package drops,
and never read. This record proposes the narrowest change that stops paying for
it.

## Decision

### The seam

Parts opened from an **owned** source archive hold their payload behind a
`std::sync::OnceLock<Arc<Vec<u8>>>` plus a shared handle to the retained source
and that part's ZIP member identity. The payload is decoded on first access and
at most once per package. Borrowed ingress (`OpcPackage::from_bytes*`), streamed
ingress with no retained buffer, and parts authored in memory keep today's eager
representation, because there is nothing to decode from later; the lazy
representation is available exactly when `source_archive` is `Some`, which is
exactly when `authorize_owned_source` (`package.rs:1204`) ran.

### `Part::blob` keeps its signature; the fallible accessors force the decode

`Part::blob(&self) -> &[u8]` (`part.rs:52`) does **not** change. Candidate C2
would have made it fallible and migrated 1,078 call sites across seven crates;
this record rejects that cost. Instead the invariant is:

> **No `&dyn Part` is handed to a caller outside `litchi-opc` before its payload
> has been decoded.**

Every public route to a part except one is already fallible and forces the
decode before it returns: `main_document_part` (`package.rs:521`), `get_part`
(`:552`), `get_part_mut` (`:589`) and `part_by_reltype` (`:621`) all return
`Result<&dyn Part>` today and gain no new signature. `OnceLock::get` and a
fallible fill both take `&self`, so forcing works through the shared borrow
`get_part` already has.

The one exception is `iter_parts(&self) -> impl Iterator<Item = &dyn Part>`
(`package.rs:717`), which is infallible. It gains a sibling:

```rust
pub fn try_iter_parts(&self) -> impl Iterator<Item = Result<&dyn Part>>
```

which forces each part's decode as it yields it. `iter_parts` itself is retained
only for the metadata shapes — sites that read `partname()`, `content_type()` or
`rels()` and never `blob()` — and its item type narrows to a name-and-metadata
view that has no `blob()` at all, so a site that needs bytes cannot compile
against it. Inside `litchi-opc`, the publication plan reads payloads through a
crate-private `decoded_blob(&self) -> Option<&[u8]>` that returns `None` for a
part that was never forced, and treats `None` as proof of pristineness rather
than as a reason to decode.

`part_count` (`:726`) and `contains_part` (`:1041`) stay infallible: both answer
from the part map without touching a payload.

### Refusal relocation: `PartBytes` and `TotalPartBytes`

Gate 2 of 0581 requires that these two typed refusals still fire at `open()`
with the same resource, observed value, limit and object path, or that a
reviewed decision relocate them. **The open-time refusal is retained and does
not move.** The reader already charges both limits twice, against two different
quantities:

- **Against the declared central-directory size, before any decompression.**
  `check_declared_part_bytes_for_member` (`pkgreader.rs:1128`) reads
  `archive.metadata(name)?.uncompressed_size()` and charges
  `ReadResource::PartBytes` per part and `ReadResource::TotalPartBytes` in
  aggregate. `load_parts_eager` runs this as a batch pass at `:897`, *before*
  the bulk read at `:913`. The regression test
  `eager_declared_part_preflight_rejects_before_bulk_reader_invocation`
  (`pkgreader.rs:2149`) asserts that the decompression closure is never invoked
  when this pass fails.
- **Against the actual inflated size, after decompression** (`:917-927`), as a
  backstop for a central directory that under-declares.

This record moves **only the second charge**, and only for parts that are not
decoded. The declared-size pass stays exactly where it is, so every package that
is refused at `open()` today because its declared sizes exceed the limits is
still refused at `open()`, with the identical `OpcError::ReadLimit { resource,
actual, maximum }` (`error.rs:23`) and the identical `Display` text.

**What happens when the recheck fails after open.** A part whose declared size
passed and whose inflated size exceeds `max_part_bytes`, or whose inflation
pushes the running actual total past `max_total_part_bytes`, is refused at the
first access that decodes it. The refusal is the same `OpcError::ReadLimit`
value the eager path produces, returned from `get_part`, `get_part_mut`,
`main_document_part`, `part_by_reltype` or a `try_iter_parts` item instead of
from `OpcPackage::open`. The package remains open and every other part remains
readable; the failing part has no payload and never acquires one. The decode is
**not** retried on a later access: the failure is recorded in the cell so the
refusal is stable and idempotent, because a limit failure that alternates with
success across accesses would make the error identity depend on call order.

Two consequences are relaxations, and the reviewer is asked to accept or refuse
them explicitly:

1. **A package whose central directory under-declares one part's size now opens
   where it previously did not.** The refusal still exists but is reached only
   by touching that part. A caller that opens a package to read its content
   types, or to count its parts, and never touches the offending payload, now
   succeeds. The declared-size ceiling still bounds what the archive may claim,
   so the *bounded-resources* property is preserved; what changes is *when* a
   forged declaration is caught.
2. **The aggregate actual charge is over decoded parts, not over all parts.** A
   package whose honest declared total is within the limit but whose true
   inflated total exceeds it is refused today and, under this record, is refused
   only once enough parts have been decoded to cross the limit. Since the
   declared total is charged in full at open and an honest ZIP declares an
   uncompressed size no smaller than the truth, the only packages affected are
   those with a forged central directory.

The alternative the reviewer may prefer — charging the *declared* aggregate as a
hard reservation held for the package's lifetime, so nothing about the open-time
refusal changes at all — is already implemented at the physical layer:
`reserve_declared_parts` / `commit_actual_parts` / `release_declared_parts`
(`phys_pkg.rs:779`, `:860`, `:849`) exist and are used by `blob_for` (`:325`)
for exactly this reserve-then-reconcile shape. Adopting it costs nothing new in
code but keeps the aggregate reservation at its declared, not actual, size for
the package's lifetime, which is the conservative reading of ADR 0005's
"hierarchical resource budget". This record prefers reconciling on decode and
records the alternative rather than deciding it unilaterally.

### The error channel for a first-access decode failure

No new error variant is introduced. A first-access decode failure surfaces as
one of the variants `OpcError`'s existing `From<soapberry_zip::Error>` mapping
(`error.rs:343-384`) already produces at open: `ReadLimit` for a limit,
`Allocation` for an allocation failure, `Cancelled` for cancellation,
`IoError` through `map_io_error`, and `ZipError(String)` for a corrupt Deflate
stream or a CRC mismatch. `OpcError` is `#[non_exhaustive]` (`error.rs:8`), so
no enum change is required and no caller's `match` breaks.

What does change is the *set* of variants `get_part` and `get_part_mut` can
return. Today they can only return `OpcError::PartNotFound`; under this record
they can also return the decode family. Callers that treat a `get_part` error as
"absent" must be audited; change 0610 enumerates them. Error text is unchanged
for every variant, so no diagnostic string moves.

ADR 0005 requires limit errors to "identify the resource, observed value, limit,
and object path". `OpcError::ReadLimit` carries resource, observed value and
limit but not an object path, at open or at first access alike; this record does
not change that and does not claim to satisfy the object-path clause any better
or worse than the eager path does.

### `Clone`, `Send` and `Sync`

`OpcPackage` is `#[derive(Clone)]` (`package.rs:80`) and is `Send + Sync`
structurally, because `parts` is typed `Box<dyn Part + Send + Sync>` and `Part`
is declared `Part: PartClone + Send + Sync` (`part.rs:49`). Nothing in the struct
uses interior mutability today.

`std::sync::OnceLock<T>` is `Sync` when `T: Send + Sync` and `Clone` when
`T: Clone`, so both properties survive; `Cell`, `RefCell` and `OnceCell` would
cost `Sync`, and `RwLock` has no `Clone` impl at all and would break the derive
outright. `OnceLock<Arc<Vec<u8>>>` is therefore the required primitive, and
**gate 6 of 0581 is satisfied structurally rather than by test**.

Cloning a package clones each part's cell — a decoded clone stays decoded, an
undecoded clone stays undecoded — and `Arc`-shares the source handle, preserving
the property `clone_shares_owned_source_but_revocation_is_independent`
(`package.rs:1666`) already tests, that clones share rather than duplicate the
source allocation. The shared handle carries the shared
`Arc<Mutex<PartBudget>>`, so a clone's decodes charge the same budget as the
original's; a part decoded independently in two clones is charged twice, which is
the behaviour `blob_for`'s doc comment already documents for repeated
materialization.

### The publication plan

Change 0593 made `PublicationPlan::from_package` decide a relationship
collection pristine by `Arc::ptr_eq` against the preservation provenance
(`pkgwriter.rs:145-150`, `:170-174`), and decide `[Content_Types].xml` from
provenance metadata alone (`:139-144`, `:168-169`). Neither decision reads a part
payload. The remaining payload read is `source_blob_retained` (`:829-831`), which
compares the planned blob against the provenance blob, pointer first and bytes
second.

Under this record a part that was never decoded is pristine **by construction**:
its payload cannot have been replaced, because no caller ever held it. The plan
takes `PreservationAction::Copy` for such a member without decoding it and
without comparing anything. This is not new authority: it is the same
planning-evidence use of provenance that ADR 0005's 2026-08-21 amendment permits
and that 0593 already relies on, and it does not touch
`exact_source_authorized`, does not widen who may take whole-archive
passthrough, and does not remove the retained source archive.

### Migration plan

1. Land the lazy representation, `try_iter_parts` and the crate-private
   `decoded_blob` inside `litchi-opc` with `iter_parts` unchanged, so no crate
   outside `litchi-opc` compiles differently. The package is still eagerly
   decoded at this step; only the plumbing exists. Gates: byte-identical corpus
   output and identical refusal identity.
2. Make the publication plan treat a never-decoded part as pristine, and stop
   the eager decode for owned-source ingress. This is the step that changes
   behaviour and must clear every gate in change 0610.
3. Narrow `iter_parts`'s item type to the metadata view and migrate the
   payload-reading sites among the 259 in-scope production `iter_parts()` sites
   to `try_iter_parts`. Change 0610 brackets that group at 16–88 by two
   mechanical scopes and puts it at about 25 by an independent full-context
   review; the rest need no behavioural change. The narrowed item type makes
   the difference a compile error rather than a review question, so the
   estimate does not have to be exact before the work starts.
4. Leave the 1,078 `.blob()` sites untouched. They are downstream of accessors
   that already force the decode.

C3 — routing the format crates onto `SourceBackedPackage` — remains the end
state and is **not** authorized by this record. It needs a general source-backed
save, which does not exist, and it would add the ADR 0003 typed readback that
the eager path does not pay today.

## Consequences

- A package opened from an owned source retains the compressed archive plus an
  index plus only the payloads someone asked for. On the measured corpus that is
  0.73% of today's inflated bytes for an XLSX one-tab edit and save.
- A decode failure becomes visible later than it does today, at the accessor
  rather than at the open. Two open-time refusals become first-access refusals,
  both of them only reachable with a central directory that under-declares.
- `Part::blob` stays infallible, so no call site outside `litchi-opc` changes
  its error handling, and the facade's panic-free requirement is met by the
  decode-before-handing-out invariant rather than by a new `Result`.
- `iter_parts` stops being able to reach payloads. That is the point: it is the
  only infallible route to a part, so it must not be one.
- Exact-source no-op publication is untouched: it never reads a part payload.
- Cost: the lazy representation adds per-part member identity and a cell to
  every part of an owned-source package, so a package all of whose parts are
  read is slightly *larger* and no faster than today. The design is a bet that
  the measured working set is small, and change 0610 is the evidence for it.

## Verification

Acceptance requires, beyond the six gates 0581 already states:

1. Byte-identical published output for every OOXML fixture in `test-data` across
   every mutation scenario, by whole-stream digest, before and after — the
   1,344-row oracle shape change 0593 used.
2. Every refusal preserved by identity, at open and at first access: the
   corpus's existing open refusals, save refusals, and a constructed archive
   whose central directory under-declares a part, which must refuse with the
   same `OpcError::ReadLimit` value at whichever point this record specifies.
3. A test that a never-decoded part publishes the same bytes as the same
   package with that part force-decoded.
4. A test that `OpcPackage` is still `Clone + Send + Sync` and that a clone
   shares rather than duplicates the source allocation, extending
   `package.rs:1666`.
5. A test that a failed first-access decode is stable across repeated accesses.
6. Peak-retained-byte and allocation counts on 0581's three axes, showing the
   predicted factors were met and that no axis regressed.
