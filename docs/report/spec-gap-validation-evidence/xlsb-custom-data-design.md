# XLSB Custom Data and Custom Data Properties owner design

Status: design handoff only. This note defines the implementation and evidence
boundary for MS-XLSB §2.1.7.10/§2.1.7.11. It does not add a reader, writer,
transaction, fixture, or native-application support claim. The current tree
has no XLSB Custom Data owner, and the local native corpus does not yet contain
a producer-backed Custom Data example.

The feature is an **atomic package owner**, not a read-only binary inventory.
The normal operation must be able to insert, replace, rename, detach, retarget,
and remove one Custom Data storage while updating every recognized host
reference in the same source-checked transaction. A read-only projection can be
implemented first, but it is not a sufficient completion of this gap.

The governing repository rules are [`docs/GOAL.md`](../../GOAL.md), the
accepted [ADR index](../../adr/README.md), [ADR 0001](../../adr/0001-priorities-and-api-layers.md),
[ADR 0002](../../adr/0002-crate-topology.md),
[ADR 0003](../../adr/0003-snapshots-edits-and-patches.md),
[ADR 0005](../../adr/0005-io-memory-and-performance.md),
[ADR 0006](../../adr/0006-validation-security-and-compatibility.md),
[ADR 0011](../../adr/0011-ooxml-physical-package-ownership.md), and
[ADR 0024](../../adr/0024-current-topology.md). The audit entry is
[`spec-gap-audit.md`](../spec-gap-audit.md); it currently records XLSB
§2.1.7.10/§2.1.7.11 as a zero-source-match gap and records the XLSX owner as
synthetic/source-bound evidence without a native producer claim.

## Decision

The owner is split into three layers:

```text
litchi-opc
    |
    +-- litchi-ooxml-common::custom_data
    |       (X14 datastoreItem model and bounded XML codec)
    |
    +-- litchi-xlsx::custom_data
    |       (XLSX OPC graph and SpreadsheetML Connections bindings)
    |
    +-- litchi-xlsb::custom_data
            (XLSB OPC graph and BIFF12 BrtBeginExtConn14 bindings)
```

`litchi-xlsx` and `litchi-xlsb` consume the package-neutral leaf separately;
neither concrete spreadsheet crate may depend on the other. The common leaf
does not open an OPC package, inspect workbook relationships, parse
`connections.xml`/`connections.bin`, or own a transaction. The host layers own
those concerns and map their errors into their existing public error types.
This follows ADR 0001's three public layers, ADR 0002/0024's downward crate
topology, ADR 0003's source-bound transaction model, ADR 0005's explicit
bounded operations, ADR 0006's preservation rules, and ADR 0011's OPC
ownership.

The ordinary XLSB entry point should be `litchi_xlsb::custom_data` and the
package methods below. It must not expose `OpcPackage`, ZIP/archive errors,
BIFF record IDs, relationship IDs, source ranges, or a mutable XML/record tree
in a normal CRUD signature. A narrow raw/source module may expose validated
record spans to the host implementation and tests.

## Existing implementation boundary

The XLSX owner already provides the useful transaction vocabulary in
[`custom_data/mod.rs`](../../../crates/litchi-xlsx/src/custom_data/mod.rs),
[`model.rs`](../../../crates/litchi-xlsx/src/custom_data/model.rs), and
[`package.rs`](../../../crates/litchi-xlsx/src/custom_data/package.rs):
`CustomData`, `Properties`, `ExtensionList`, `RemovalDisposition`, bounded
`Limits`, source-bound `Snapshot`/`Transaction`, and reversible `Commit`/`Patch`.
Its package adapter also captures content types, workbook relationships,
properties/payload members, and the XLSX Connections `embeddedDataId` closure.
That owner is the source of API and preservation behavior to extract, not a
crate that XLSB may import.

The XLSB package already owns `OpcPackage`, `ReadLimits`, source-checked
owner transactions, and public `apply_*_patch` seams in
[`package/mod.rs`](../../../crates/litchi-xlsb/src/package/mod.rs). Its
inert Connections owner is split across
[`host/connections/package.rs`](../../../crates/litchi-xlsb/src/host/connections/package.rs),
[`model.rs`](../../../crates/litchi-xlsb/src/host/connections/model.rs), and
[`parse.rs`](../../../crates/litchi-xlsb/src/host/connections/parse.rs).
That parser currently skips balanced `BrtBeginExtConn14`/`15` collections as
unknown records, which is a safe preservation baseline but does not expose the
`irstClientCubeUrn` binding needed for this owner. The new binding lens should
therefore be a focused source-preserving adapter, not a broad rewrite of the
existing inert connection model.

The existing [XLSX Custom Data evidence](xlsx-custom-data-v2/README.md)
demonstrates the intended source/graph/limits/inverse discipline and explicitly
makes no native-producer claim. XLSB must add the BIFF12 reference closure and
must keep that evidence distinction.

## Normative scope and exact profiles

The local MS-XLSB file-structure text deliberately delegates both parts to the
XLSX specification:

| XLSB item | Local authority | Implemented profile |
| --- | --- | --- |
| §2.1.7.10 Custom Data | [MS-XLSB] §2.1.7.10, delegated to [MS-XLSX] §2.1.2 | Inert binary payload; one typed Custom Data Properties owner points to it through the exact internal `customData` relationship. |
| §2.1.7.11 Custom Data Properties | [MS-XLSB] §2.1.7.11, delegated to [MS-XLSX] §2.1.3 and §2.4.35 | XML `datastoreItem` root in the X14 namespace, required `id`, optional one `extLst`, unknown extension contents retained. |
| Host reference | [MS-XLSB] §2.4.78 `BrtBeginExtConn14.irstClientCubeUrn` | Decoded UID is a reference to `datastoreItem/@id`; all recognized occurrences participate in rename/detach/retarget. |
| Nearby extension | [MS-XLSB] §2.4.79 `BrtBeginExtConn15.irstId` | **Not** a Custom Data reference. It is a data-model connection field and must remain outside this owner. |

The admitted canonical OPC values are:

| Graph value | Required value |
| --- | --- |
| Properties content type | `application/vnd.openxmlformats-officedocument.customDataProperties+xml` |
| Payload content type | `application/binary` |
| Workbook to properties relationship | `http://schemas.openxmlformats.org/officeDocument/2006/relationships/customDataProps` |
| Properties to payload relationship | `http://schemas.openxmlformats.org/officeDocument/2006/relationships/customData` |
| Properties root | `{http://schemas.microsoft.com/office/spreadsheetml/2009/9/main}datastoreItem` |
| Properties root `id` | Required Xstring UID, fewer than 65,536 characters |
| Direct properties children | At most one X14 `extLst`; whitespace, comments, declaration, and processing-instruction policy follows the common XML codec; unknown extension descendants remain opaque |
| Relationship target mode | Internal for both admitted edges |

The local MS-XLSX part enumeration also says that Custom Data is user-defined
binary data, that a Custom Data Properties part has one UID for its associated
storage, and that the properties part is associated with the workbook and may
point to its Custom Data payload. This owner treats multiple storage pairs in a
package as a catalog: each properties part has exactly one payload edge, and
UIDs are globally unique. “At most one” is a per-properties-part cardinality
rule, not a reason to collapse the package catalog to one entry.

No Strict custom-data relationship URI is admitted by this design. Prefixes and
relationship lexical spelling are retained as source data, but a Strict or
unknown relationship URI is a typed unsupported-profile/graph error rather
than an inferred equivalent. If a later local specification or native fixture
establishes a Strict counterpart, it gets a separate profile row and tests;
the implementation must not silently broaden this table.

`PackURI` equivalence is used for package-part identity according to the OPC
owner's existing policy. Raw `Target`/`TargetMode`/`Type` tokens are retained
for source checks and exact inverse. Normalized path equality cannot hide two
ambiguous physical parts: equivalent duplicates, case-only collisions, or
multiple owners are rejected before a semantic snapshot is published.

## Host reference record and graph closure

The Custom Data graph is:

```text
workbook.bin
  -- customDataProps (internal) --> properties part
properties part
  -- customData (internal) --> inert binary payload
connections.bin
  -- BrtBeginExtConn14.irstClientCubeUrn --> datastoreItem/@id
```

The workbook edge and properties edge must each be present exactly where the
profile requires them. A missing or dangling edge, an external target, a wrong
content type, a duplicate owner, a duplicate payload edge, an orphan payload,
or an unexpected outbound relationship is a typed graph error. A payload is
inert: it is not parsed, refreshed, queried, executed, decompressed as an
Office object, or followed as a graph root.

The XLSB host needs a source-bound binding lens over `connections.bin`. It must
walk balanced BIFF12 records and admit only `BrtBeginExtConn14` (record kind
1068) in the context required by [MS-XLSB] §2.4.78. The lens records:

* the complete enclosing record range and the exact `irstClientCubeUrn`
  XLWideString range;
* the decoded UID, including the empty value;
* the owning connection ordinal/source range for diagnostics; and
* the original `Arc<[u8]>` connections part and its relationship/content-type
  source image.

`BrtBeginExtConn14` has an FRT header (the checked-in excerpt labels it an
`FRTBlank`) and variable XLWideString fields. The parser must use the complete
§2.4.78 field sequence and bounded cursor checks; it must not locate the UID by
searching for a UTF-16 substring.
If a boundary cannot be proven, the owner returns a typed malformed/unsupported
binding error and leaves the package untouched. Unknown records and unknown
fields are copied as opaque bytes when no recognized reference is changed.

The preceding `BrtBeginExtConnection`/database-properties context is validated
where §2.4.78 requires it. A structurally valid `BrtBeginExtConn14` whose
connection context is inconsistent is not promoted to an ordinary reference.
The implementation may preserve it as an opaque connection extension, but an
operation that would rename, detach, or remove the associated storage must
refuse because the dependency closure is not proven.

`BrtBeginExtConn15` (record kind 2109) is scanned only as an unrelated opaque
record. Its `irstId` field is a spreadsheet-data-model identifier, not a
Custom Data UID. Equal decoded text in `irstId` must not create an incoming
edge, and changing a Custom Data UID must not rewrite it.

For an **effective admitted** `BrtBeginExtConn14` record outside hidden AC/FRT
wrappers, every non-empty `irstClientCubeUrn` must equal one
`datastoreItem/@id`. A connection with an empty field contributes no edge.
Many `BrtBeginExtConn14` fields may legitimately point to the same UID; this
many-to-one reference cardinality is the normal shared-storage case, and
`rename`, `detach`, and `retarget` update all of those fields atomically. A
duplicate properties owner, duplicate payload edge, overlapping/duplicated
field span, non-empty effective field with no matching UID, or an ambiguous
record range is a typed graph error. A hidden or unresolved AC/FRT record is
preserved as opaque provenance and does not by itself invalidate a direct
semantic read; it is excluded from `connection_references`. If a mutation
could change or remove the UID it may contain, the unresolved record blocks
that mutation because its dependency cannot be proven. External relationships,
query/fragment targets, and host references outside the admitted
record/profile are not followed by guesswork.

### BIFF alternate-content and grammar authority

`BrtACBegin`/`BrtACEnd` (record kinds 37/38) and `BrtFRTBegin`/`BrtFRTEnd`
(record kinds 35/36) are not ordinary unknown records that the Custom Data
scanner may skip and then edit around. The scanner keeps balanced alternate-
content and future-record stacks, captures each `cver`/`RgACVer` product-
version list and FRT product/version header, and records the complete wrapper
and contained-record ranges. Because this owner does not have an application
product/version capability policy, a Custom Data owner or `irstClientCubeUrn`
inside either wrapper is classified as hidden/unresolved. It remains
byte-preserved and readable only as opaque provenance; any rename, retarget,
detach, remove, or insertion whose closure could cross that block returns a
typed MCE/unsupported error. Unbalanced, malformed, or overlapping AC/FRT
wrappers also refuse mutation. A wrapper with no Custom Data owner/reference
remains unchanged and does not block unrelated edits.

The checked-in local §2.4.78 excerpt is not a complete writer grammar: its
`BrtBeginExtConn14` layout contains `...` around `irstCulture` and
`irstClientCubeUrn`. Before implementing a binding rewrite, pin one complete
authoritative grammar in the evidence manifest. Prefer a complete local
MS-XLSB source if one is added; otherwise use the official Microsoft Open
Specifications MS-XLSB release, recording its revision, download URL, and
SHA-256. The local ellipsis is sufficient to identify the semantic field, but
is not permission to guess the preceding/following fields, FRT header flags,
reserved bytes, or future tails.

The pinned grammar must define every payload field and the exact FRT header
encoding. A UID edit reserializes the complete admitted `BrtBeginExtConn14`
payload through that grammar, preserving untouched fields and tails byte for
byte, then recomputes the XLWideString UTF-16 unit count and bytes. It also
recomputes the enclosing BIFF12 record header: both record-kind and payload
length use variable-width encodings, and the payload-length varint may change
width when the UID grows or shrinks. All source ranges after that record are
therefore planned from the complete record range, never from a fixed payload
offset. A changed payload/header that cannot be encoded under the pinned
grammar is rejected before a replacement buffer is allocated.

The enclosing connection context is part of the admission proof, not a
best-effort diagnostic. The preceding `BrtBeginExtConnection.idbtype` must be
`DBTOLEDB` for an admitted ExtConn14 block. When the ExtConn14 collection is
non-empty, its immediately preceding `BrtBeginECDbProps.icmdtype` must be
`CMDCUBE`. If the enclosing connection is associated with a PivotCache, the
ExtConn14 collection must be empty as required by §2.4.78. Missing context,
wrong source type/command type, a non-empty PivotCache collection, or a context
boundary hidden by an AC/FRT wrapper is a typed malformed/unsupported owner
state. The package may still retain an entirely opaque, untouched connection
stream, but the Custom Data owner cannot expose its references or perform
identity/removal edits until this context is proven.

## Package-neutral model and public API

The shared leaf should extract the current XLSX model and bounded codec into
`litchi_ooxml_common::custom_data` (or an equivalently focused common module).
The common public values are:

```rust
pub struct CustomData {
    pub properties: Properties,
    pub data: Vec<u8>,
}

pub struct Properties {
    pub id: String,
    pub extension_list: Option<ExtensionList>,
}

pub struct ExtensionList {
    pub xml: Vec<u8>,
}

pub struct CustomDataView { /* shared immutable properties and payload */ }

pub enum RemovalDisposition {
    RejectReferenced,
    DetachConnections,
    RetargetConnections(String),
}
```

The exact field visibility can remain private with accessors if the common
codec needs to enforce validation before allocation. `CustomData::new(id,
data)` is a convenience constructor for an inert payload. `CustomDataView`
clones cheaply by sharing immutable properties, extension XML, and payload; an
explicit `to_owned` is the allocation-bearing conversion. The binary payload
has no semantic interpretation in this owner.

The required `id` attribute is present even when its `ST_Xstring` value is
empty. Empty values are valid existing/staged property values under the
delegated XString contract and the current XLSX codec tests; they are not
silently normalized to a generated UID. An empty value cannot be the target of
a non-empty connection reference: `RetargetConnections("")` is invalid, while
`DetachConnections` explicitly writes empty reference fields. A package may
contain at most one empty UID because decoded IDs remain globally unique.

The common codec owns `parse_properties`, `write_properties`, and validated
source-preserving rewrites for the `id` and direct `extLst`. The extraction
must move the private XLSX `ST_Xstring` decoder/escaper and its source-attribute
span helpers into this downward neutral module (or an equivalent common
source-XML module). Those helpers must not import `crate::raw`, `crate::error`,
or any `litchi_xlsx` path; each host maps common codec errors at its boundary.

The XML reader validates events against bounded source ranges but keeps the
original raw source as the byte-authoritative representation. It preserves
opaque comments, processing instructions, CDATA, entities, namespace prefixes,
attribute order, declaration, and whitespace through raw-source splices; it
does not rebuild an extension subtree from a lossy parsed tree. A complete
properties document may have a leading BOM or XML declaration when those are
legal document-level tokens. A staged `ExtensionList` fragment is embedded
inside an existing element, so a leading BOM or XML declaration in that
fragment is rejected (or removed only by an explicitly named, bounded
`strip_embedded_prolog` policy); it must never become character content inside
`datastoreItem`. Comments, processing instructions, and CDATA that are legal
inside the retained extension subtree remain preserved.

XML validation is namespace-aware and bounded. The common parser must reject
malformed XML, duplicate required roots or `id` attributes, multiple direct
`extLst` children, illegal namespace bindings, and an over-limit extension
before constructing an unbounded replacement. It must not reject a legal opaque
comment or CDATA just
because its bytes contain `<?` or `]]>`; event context, not a byte substring,
determines whether a declaration or delimiter is illegal.

Namespace admission compares the decoded expanded URI, not the lexical prefix
or raw URI spelling. Thus arbitrary aliases and XML entity-escaped forms of the
X14 URI (for example an escaped slash in the namespace declaration) resolve to
the same `datastoreItem`/`extLst` namespace; the raw declaration and attribute
lexemes remain in the source image. The resolver applies XML attribute entity
decoding before comparing the URI and rejects reserved/empty bindings under
the common namespace rules.

The direct grammar is element-only at the recognized containers. A
`datastoreItem` may contain only XML whitespace plus comments/processing
instructions around zero or one direct X14 `extLst`; non-whitespace text,
CDATA, or general references directly under that root are invalid. The direct
`extLst` must satisfy the pinned `CT_ExtensionList`/direct-extension grammar:
its recognized direct children and required attributes are checked, and its
direct content is likewise element-only except for permitted XML whitespace,
comments, and processing instructions. Unknown descendants inside an admitted
extension element are opaque and may contain legal CDATA, comments, processing
instructions, and future elements; their raw bytes are not normalized or
interpreted as Custom Data references. A staged replacement is checked at the
same container boundary before the source splice.

The XLSX and XLSB modules re-export the common values and codec vocabulary,
but retain their own `Snapshot`, `Transaction`, `Commit`, `Patch`, `Limits`,
and package graph implementation. The current XLSX names may be retained as a
facade during extraction; `Part` should be documented as a semantic storage
view, and physical part-name accessors should move to a low-level/source view
rather than entering the normal XLSB CRUD API.

The proposed XLSB ordinary surface is:

```rust
impl Package {
    pub fn custom_data(&self) -> Result<custom_data::Snapshot>;
    pub fn custom_data_with_limits(
        &self,
        limits: custom_data::Limits,
    ) -> Result<custom_data::Snapshot>;
    pub fn custom_data_with_limits_and_context(
        &self,
        limits: custom_data::Limits,
        context: litchi_core::ExecutionContext,
    ) -> Result<custom_data::Snapshot>;

    pub fn edit_custom_data(&self) -> Result<custom_data::Transaction>;
    pub fn edit_custom_data_with_limits(
        &self,
        limits: custom_data::Limits,
    ) -> Result<custom_data::Transaction>;
    pub fn edit_custom_data_with_limits_and_context(
        &self,
        limits: custom_data::Limits,
        context: litchi_core::ExecutionContext,
    ) -> Result<custom_data::Transaction>;

    pub fn apply_custom_data_patch(
        &self,
        patch: &custom_data::Patch,
    ) -> Result<Self>;
    pub fn apply_custom_data(
        &self,
        commit: &custom_data::Commit,
    ) -> Result<Self>;
}

impl Workbook {
    pub fn custom_data(&self) -> Result<custom_data::Snapshot>;
    pub fn custom_data_with_limits(
        &self,
        limits: custom_data::Limits,
    ) -> Result<custom_data::Snapshot>;
    pub fn custom_data_with_limits_and_context(
        &self,
        limits: custom_data::Limits,
        context: litchi_core::ExecutionContext,
    ) -> Result<custom_data::Snapshot>;
    pub fn edit_custom_data(&self) -> Result<custom_data::Transaction>;
    pub fn edit_custom_data_with_limits(
        &self,
        limits: custom_data::Limits,
    ) -> Result<custom_data::Transaction>;
    pub fn edit_custom_data_with_limits_and_context(
        &self,
        limits: custom_data::Limits,
        context: litchi_core::ExecutionContext,
    ) -> Result<custom_data::Transaction>;
    pub fn apply_custom_data(
        &mut self,
        commit: &custom_data::Commit,
    ) -> Result<custom_data::Snapshot>;
    pub fn apply_custom_data_patch(
        &mut self,
        patch: &custom_data::Patch,
    ) -> Result<custom_data::Snapshot>;
}
```

These are proposed names, aligned with current XLSB patterns: `Package` reads
and starts detached transactions from `&self`, and `Package::apply_*` returns a
new validated package; `Workbook` starts the same detached transaction from
`&self` and mutably forwards commit/patch publication through its existing
`edit_opc`/reparse seam. `Package::apply_custom_data` is a convenience for a
commit and delegates to `apply_custom_data_patch`; it does not expose or borrow
the private `OpcPackage`. There is no `Transaction<'a>` borrow of a Package.

The context-bearing variants consume an owned
`litchi_core::ExecutionContext`. A context-enabled `Snapshot`/`Transaction`
retains its own context (or a cheap clone) for admission, record/event checks,
and final publication; no public operation stores `&ExecutionContext` or a
runtime handle. `Workbook` context-bearing forwarding methods should be added
with the same owned-context convention rather than introducing a second
execution policy. The default methods use the ordinary bounded synchronous
path.

`Snapshot` provides deterministic UID order, `len`, `is_empty`, `find(id)`,
`storages`, `limits`, and `connection_references(id)`. A storage view provides
its semantic `id`, `properties`, and borrowed inert `data`; it does not expose
package names or relationship IDs in the ordinary layer. A `StorageSelector`
may offer `Id(&str)` and checked `Index(usize)` variants for transaction verbs,
but a missing selector is `Option`/typed error rather than an unchecked index.

`Transaction` provides the existing XLSX-shaped verbs:

```rust
set(selector, CustomData) -> Result<bool>
set_data(selector, Vec<u8>) -> Result<bool>
edit_properties(selector, closure) -> Result<bool>
insert(CustomData) -> Result<StorageId>
upsert(CustomData) -> Result<StorageId>
rename(selector, id) -> Result<bool>
remove(selector) -> Result<Option<StorageView>>
remove_with(selector, RemovalDisposition) -> Result<Option<StorageView>>
commit(self) -> Result<Commit>
```

`StorageId` is an owned validated `ST_Xstring` UID (including the legal empty
value) with `as_str()` and cheap cloning.
Insertion results must not borrow the transaction or return a catalog ordinal
that can shift when another storage is inserted or removed. Callers can pass
`StorageSelector::Id(id.as_str())` to subsequent operations. Its retained
storage participates in the transaction resource budget.

`set`/`rename` of an existing UID is a graph operation: all recognized
`BrtBeginExtConn14.irstClientCubeUrn` references are staged with it. The
ordinary API has no catalog-string-only rename. `set_data` changes only the
inert payload after the complete graph is validated. `Commit` exposes
`changed`, `snapshot`, and a source-checked reversible `patch`. `Patch` exposes
`is_empty`, `before`, `after`, `inverse`, and application through the XLSB
package facade. No operation deliberately panics.

## Source ownership, opaque preservation, and exact inverse

An XLSB `Snapshot` is source-bound and cheap to clone. Its private source image
must retain at least:

* `[Content_Types].xml` bytes and the exact mappings used by the properties,
  payload, and connections parts;
* workbook bytes and its exact relationship member, including relationship IDs,
  target lexical spelling, type spelling, and empty relationship members;
* the properties XML bytes, a validated XML source proof/range for `id` and
  direct `extLst`, its `.rels` member, and the inert payload bytes and `.rels`;
* the workbook-to-connections relationship, `connections.bin` bytes, and its
  exact relationship/content-type source image;
* all inbound relationship metadata needed to prove exclusive removal, plus
  package identity/source version and signature/edit policy; and
* the decoded binding spans and their enclosing BIFF record ranges.

The implementation keeps explicit private `ReadSet` and `WriteSet` values for
each snapshot/commit. The read set includes every source member and range that
can affect admission or closure: content types, root/workbook and owner
relationships, properties XML and direct extension range, payload bytes and
relationships, connections relationship and binary bytes, every ExtConn14
record plus its AC/FRT provenance, and all inbound graph metadata. The write
set contains only changed source ranges or newly allocated/removed members.
Before candidate construction, the current package must match the entire read
set. After publication, a source audit verifies that bytes outside the write
set are unchanged. The sets may be summarized in diagnostics, but raw package
IDs and source ranges do not enter ordinary semantic selectors.

`Patch` is an opaque, process-local source proof, not a persisted interchange
format. It has no `serde`/file representation or cross-process compatibility
promise. Allocation identity is only a fast path: applying a patch to a
separately reopened package in the same process is allowed when every captured
read-set member has the same exact bytes, part identity, relationship/content-
type provenance, and profile. The apply seam then remaps XML/BIFF ranges from
those source bytes and revalidates unique anchors before writing; ambiguous or
missing remaps return `PatchConflict`. A different process, revision,
profile, or normalized-but-lexically different source is not accepted. This
byte/provenance remap contract matches existing XLSB inverse tests while
keeping source allocations authoritative when they are still shared.

Ranges are meaningful only against the exact source allocation and part
identity captured by the snapshot. A stale, foreign, truncated, or
out-of-bounds proof yields `PatchConflict`/typed invalid input. The host must
not copy a complete workbook merely to inspect one binding, and it must not
materialize a second owned copy of an unchanged payload.

An exact semantic no-op returns `Commit { changed: false }`, an empty patch,
and the original source allocations/bytes. A real edit splices only the
recognized property token, extension subtree, payload part, relationship
members, content-type entries, and BIFF12 UID fields that the graph plan
requires. It preserves unrelated package members, record bytes, opaque XML,
unknown BIFF records, comments, processing instructions, CDATA, namespace
prefixes, attribute order, whitespace, and entity lexical spelling. New
members use deterministic names/IDs without consulting time, process state,
filesystem metadata, or randomness.

The writer must read back the complete candidate package under the same limits
and compare semantic storage values **and** decoded custom-data bindings before
publication. The inverse patch must restore the original content types,
relationship XML, property XML, payload bytes, connections bytes, and empty
member presence exactly; it must not regenerate a canonical approximation.
Applying a patch to a package changed anywhere in this source closure fails
without mutating the target.

Unknown XML in `extLst` is retained as a bounded opaque subtree. Unknown BIFF12
records, including an unmodelled `BrtBeginExtConn15` extension, are retained as
opaque bytes. An opaque reference-bearing record is not silently edited: if its
dependency on a UID cannot be proven, rename, retarget, and removal return a
typed unsupported/ambiguous error before staging publication.

## MCE and effective-owner policy

The properties root is a single effective `datastoreItem`; the owner scanner
must not select whichever `datastoreItem` happens to be found in a nested
`mc:Choice` or `mc:Fallback`. The following rules are required:

1. A direct X14 root with one required `id` is admitted. Prefix aliases are
   resolved by URI, and source prefix spelling is retained.
2. A root, `id`, direct `extLst`, or reference-bearing data hidden under an
   `AlternateContent` branch is unsupported for typed mutation. The scanner
   records the branch/provenance and returns an MCE ambiguity/unsupported
   error rather than creating a second direct owner or guessing active
   semantics.
3. More than one effective properties owner or payload edge is ambiguous and
   no edit is admitted. Multiple recognized `BrtBeginExtConn14` binding fields
   may describe the same storage and are allowed; they are many-to-one
   references, not duplicate ownership, and every one is included in the
   rename/detach/retarget plan. Overlapping spans, duplicate decoding of one
   physical field, or a binding whose AC provenance is unresolved remains
   malformed/unsupported.
4. MCE or future markup inside an admitted opaque `extLst` descendant remains
   opaque. Its bytes are preserved, but nested IDs are not inferred as
   Custom Data references.
5. A package with an unrelated MCE branch elsewhere remains readable and
   editable when the Custom Data closure is unambiguous. An MCE branch that
   could add an owner, inbound edge, or competing UID causes the relevant
   mutation to refuse before any output bytes are built.

The same rule applies to BIFF12 `BrtACBegin`/`BrtACEnd`: because this owner has
no selected application capability, a binding found inside that balanced block
is hidden/unresolved rather than an effective edge. The AC wrapper and all
contained records remain in the source read set. A typed edit that could
change, remove, or add a storage across such a block is refused; an exact
no-op keeps the complete wrapper byte-for-byte.

This policy is deliberately conservative: active branch resolution is a
semantic capability, not a local-element-name search. It also prevents a
writer from appending a direct properties owner beside an ignored branch and
silently changing the effective package graph.

## Removal and identity transitions

The ordinary removal policy is explicit:

| Disposition | Behavior |
| --- | --- |
| `RejectReferenced` | Return a typed `CustomDataReferenced` error when any recognized `BrtBeginExtConn14` reference names the UID. |
| `DetachConnections` | Rewrite every recognized reference to an empty valid XLWideString, preserve all other connection record bytes, and then remove the pair if no other inbound graph edge remains. |
| `RetargetConnections(new_id)` | Require an existing, unambiguous remaining UID and rewrite every recognized reference to it atomically before removing the pair. |

Removal also checks all OPC inbound relationships to the properties and data
parts, not only the workbook owner and the properties-to-payload edge. A
shared payload, an unrelated inbound owner, a foreign/external target, or an
opaque unproven inbound edge blocks deletion. Orphaned recognized property or
payload parts are invalid rather than silently garbage-collected. A missing
semantic UID is an idempotent no-op only when the complete owner graph is
absent; a half-present or contradictory graph is an error.

Renaming stages the property `id` span and all recognized BIFF12 UID spans in
one plan. Empty `ST_Xstring` IDs are syntactically valid and may be retained or
created when they are unreferenced; an existing referenced ID cannot be cleared
without first detaching its references, and an empty ID cannot be a retarget
destination. Duplicate decoded IDs, UID length overflow, invalid UTF-16, or
unproven references reject before allocation.

An edit of only an inert data payload does not touch `connections.bin` or
property XML. An edit of only opaque extension bytes validates the complete
X14 root and graph before replacing the exact direct `extLst` range. Every
combination in one transaction is planned together, so an intermediate
state—such as deleting a referenced pair before retargeting its connections—
is never published.

## Limits and cancellation

`custom_data::Limits` is a caller policy layered beneath the package's
`litchi_opc::ReadLimits`; hard protocol ceilings cannot be raised by callers.
At minimum it needs checked setters for:

* storage count, UID character/byte count, properties XML bytes, direct
  extension XML bytes, XML depth/node/event count, and payload bytes;
* `connections.bin` bytes, BIFF record count, recognized custom-reference
  count, and one-record/aggregate replacement bytes;
* package graph nodes, inbound/outbound relationship count and relationship
  XML/record bytes; and
* final part bytes, aggregate package bytes, content-type mapping count/bytes,
  and temporary replacement bytes.

Admission checks raw source lengths and record headers before allocating
decoded strings or vectors. Every checked arithmetic operation uses overflow
handling, every fallible vector/string construction reserves under the
caller/hard cap, and a candidate replacement is precharged before building a
large temporary buffer. Final output validation repeats the limits after the
complete graph plan has been serialized; a scratch upper bound is not treated
as the final candidate size when removal shrinks a member or a reused ID has a
different digit width.

The context-enabled methods consume the existing
`litchi_core::ExecutionContext` by value and retain an owned context in the
detached snapshot/transaction. The default methods use the package's normal
bounded synchronous path and never create a hidden runtime. The parser and
writer check `context.check()` (or the equivalent existing adapter) between
BIFF records, XML events, relationship edges, and before each potentially
large reservation and before final publication. Cancellation returns the
typed cancellation error and leaves the source package and source snapshot
unchanged. A cancelled transaction is not resumable and its candidate is
dropped. No filesystem, network, query engine, refresh, macro, or external
connection provider is consulted.

## Synthetic fixture and verification plan

Until a native producer fixture is available, the host tests should build a
minimal valid XLSB OPC package through the existing physical-package seam and
reopen it through `litchi_xlsb::Package`. The fixture builder should add:

* a valid `/xl/workbook.bin` and workbook relationship member;
* one or more properties parts under a deterministic **synthetic** path such
  as `/xl/customData/props1.xml` (the path is test data, not an observed native
  producer convention), each
  with the exact X14 `datastoreItem` and a bounded inert binary payload;
* `[Content_Types].xml` overrides for the properties and payload content types;
* a properties `.rels` member with exactly one internal `customData` edge and
  an explicit empty payload `.rels` member where the package seam permits it;
* `/xl/connections.bin` with valid `BrtBeginExtConnections` framing, a valid
  external connection context, one `BrtBeginExtConn14`/`EndExtConn14` pair,
  and an `irstClientCubeUrn` equal to the properties UID; and
* unrelated opaque package members, unknown BIFF records, a balanced
  `BrtACBegin`/`BrtACEnd` block, comments/PIs/CDATA, and prefix aliases in the
  property XML.

The fixture must be assembled with the same record builder and relationship
writer used by the XLSB host tests. It must not hand-wave a four-byte BIFF
signature or rely on a filename to establish a valid part. The package must
round-trip through `Package::to_bytes`/`from_bytes` and the OPC source layer.
The fixture manifest must also pin the complete ExtConn14 grammar source used
by the binding builder; no field order or record-header width is inferred from
the local §2.4.78 excerpt's ellipses. No synthetic path, relationship order,
or generated ID is reported as native behavior.

The focused matrix is:

| Area | Required cases |
| --- | --- |
| Read/model | no owner; one valid pair; multiple deterministic pairs; empty payload; empty required UID; maximum admitted UID; borrowed/shared view; unrelated opaque members preserved |
| XML profile | canonical X14 prefix; arbitrary prefix alias; declaration/comments/PI/CDATA/entity spelling; unknown ext child; one direct extLst; duplicate extLst; wrong root/namespace; duplicate/missing `id`; legal empty `id`; embedded BOM/XML declaration rejection or explicit strip policy; raw delimiter/control/invalid namespace; over-limit XML |
| OPC graph | missing workbook owner; duplicate owners; external owner; wrong relationship/content type; missing payload; duplicate payload edge; shared payload; payload outbound edge; orphan properties/payload; equivalent path collision; empty `.rels` preservation |
| BIFF binding | empty `irstClientCubeUrn`; one UID referenced by many ExtConn14 records; all references renamed/detached/retargeted; malformed XLWideString; wrong `DBTOLEDB`/`CMDCUBE` context; non-empty PivotCache ExtConn14 collection; duplicate/overlapping spans; `BrtBeginExtConn15.irstId` equal to UID but not counted; unknown record bytes survive |
| MCE/opaque | direct owner in Choice/Fallback refuses typed mutation; balanced and unbalanced `BrtACBegin`/`BrtACEnd` and `BrtFRTBegin`/`BrtFRTEnd`; hidden ExtConn14 reference is preserved but blocks mutation; unrelated inactive branch preserved; opaque extLst MCE/PI/comment/CDATA survives; no second owner is appended beside hidden owner |
| Transactions | scalar no-op; payload replacement; extension replacement; insert/upsert; rename with multiple bindings; reject/detach/retarget removal; same-transaction mixed edits; exact inverse; stale workbook XML, relationship, content-type, property, payload, or connections source conflict; atomic failure/source reopen |
| Limits | one-under and exact limits for every XML/BIFF/graph/part/aggregate cap; final shrinking replacements; decimal-width relationship/record changes; cancellation during scan and immediately before publish; no over-limit temporary allocation |
| Publication | complete reopen and semantic/binding readback; exact source bytes for no-op/inverse; unrelated part and empty relationship member retained; signed/edit-policy refusal; deterministic generated names/IDs |

Common codec tests belong under
`crates/litchi-ooxml-common/tests/custom_data.rs`. XLSB graph, BIFF binding,
source, transaction, and package tests belong under
`crates/litchi-xlsb/tests/custom_data.rs` (with a narrow internal binding test
module beside `host/connections`). Existing XLSX custom-data tests remain the
regression suite for the extracted leaf and its XLSX-specific Connections
binding; they must not be weakened to accommodate XLSB. A malformed/opaque
fixture test should assert both the typed diagnostic and byte-for-byte source
preservation.

Every evidence run should record the fixture manifest and SHA-256, the exact
feature/command/limit settings, focused test counts, `cargo check`, warning-
denied Clippy/rustdoc where applicable, and offline schema/relationship checks.
No latency, allocation, RSS, or throughput claim is valid without a named
workload, measurement protocol, and artifact under the performance evidence
tree.

## Native evidence gap

The local specifications establish the delegated part profiles and the
`BrtBeginExtConn14.irstClientCubeUrn` rule. They do not establish that a local
Excel, LibreOffice, or other producer emits this complete XLSB graph. The
current corpus search has no pinned native file containing a Custom Data
Properties member, a Custom Data payload, or a proven `BrtBeginExtConn14`
Custom Data UID. Existing native data-model and external-connection files are
not evidence for this owner; in particular, `BrtBeginExtConn15.irstId` must not
be relabelled as `irstClientCubeUrn`.

Therefore the implementation status must remain **synthetic source/graph
evidence only** until a native producer fixture is checked in with its source
path, package member inventory, hashes, exact record/XML spans, and acceptance
behavior. The native gate should include at least one read, one no-op save,
one supported rename/detach or an explicit producer-edit refusal, and a
reopen/readback comparison. Until then, do not claim native interoperability,
producer round-trip, or compatibility with any specific Office build.

## Bounded file-ownership handoff

This design file is the only file owned by this handoff. The implementation
should be split so a coder, reviewer, and tester can work without concurrent
ownership of the same production module:

| Owner | Files/scope | Required result |
| --- | --- | --- |
| Common-leaf coder | `crates/litchi-ooxml-common/src/custom_data/{mod,model,codec}.rs`, common exports, and common codec tests | X14 model/codec, bounded XML validation, source-range rewrites, no package/host imports |
| XLSX host maintainer | `crates/litchi-xlsx/src/custom_data/{mod,package}.rs` and its existing Connections binding adapter | Rebase the current XLSX owner onto the common leaf without changing its graph, source, or inverse guarantees |
| XLSB host coder | `crates/litchi-xlsb/src/custom_data/{mod,package,transaction}.rs`, `Package` facade seam, and host error mapping | OPC catalog transaction, source closure, limits/context, commit/patch/readback, ordinary API |
| BIFF binding coder | `crates/litchi-xlsb/src/host/connections/custom_data.rs` plus the smallest `connections` parser/writer seam | Exact `BrtBeginExtConn14` span scanner/rewriter; `ExtConn15` exclusion; unknown record preservation |
| Test owner | `crates/litchi-ooxml-common/tests/custom_data.rs`, `crates/litchi-xlsb/tests/custom_data.rs`, fixture helpers/evidence manifest | Synthetic package matrix, malformed/MCE/opaque cases, exact inverse/stale/limits/cancellation gates |
| Reviewer | This design, spec citations, API/error review, graph/MCE/source/limit review | Review only after the focused gates and native-evidence status are explicit; no silent profile broadening |

The merge order is common leaf and tests, then XLSX facade migration, then the
XLSB binding lens, then the XLSB package transaction and public reexports, then
the integration matrix. Shared `lib.rs`/module export edits should be made by
the host coder owning that crate, with the common leaf API frozen before either
host writes its package adapter. No task in this handoff stages, commits, or
modifies production files.

## Completion gate

The gap is ready for review only when all of the following are true:

1. the common codec and both host adapters compile without peer-crate
   dependencies or public package-ID/source leakage;
2. a complete authoritative ExtConn14 grammar is pinned with revision/hash
   before the BIFF binding writer is enabled, and valid synthetic XLSB graphs
   read, edit, reopen, and inverse exactly;
3. malformed, duplicate, external, MCE-ambiguous, unsupported, stale,
   cancelled, and over-limit candidates fail before publication;
4. `BrtBeginExtConn15` is demonstrably excluded while every admitted
   `BrtBeginExtConn14` UID is updated atomically;
5. opaque XML/BIFF/package members and lexical source bytes survive untouched
   when outside the staged closure; and
6. the report distinguishes synthetic evidence from native producer evidence,
   with no native approval or performance claim until the missing fixture is
   pinned.

The direct specification inputs are [MS-XLSB §2.1.7.10–§2.1.7.11 and
§2.4.78–§2.4.79](../../../3rdparty/specs/%5BMS-XLSB%5D/2%20Structures/2.1%20File%20Structure.md),
[MS-XLSB records](../../../3rdparty/specs/%5BMS-XLSB%5D/2%20Structures/2.4%20Records.md),
[MS-XLSX part enumerations](../../../3rdparty/specs/%5BMS-XLSX%5D/2%20Structures/2.1%20Part%20Enumerations.md),
and [MS-XLSX `datastoreItem`](../../../3rdparty/specs/%5BMS-XLSX%5D/2%20Structures/2.4%20Global%20Elements.md).
