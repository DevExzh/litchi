# XLSX next PivotTable owner design: `pivotTableData`

Status: reviewable design only, prepared 2026-09-12. This note does not add
production support, change the feature matrix, build the crate, or make a
native interoperability claim. It records the next family for implementation
after the committed `cachedUniqueNames` diagnostic/scalar batch
(`dae6162fb`). The unresolved `CT_Items` binding for that batch remains out of
scope.

Inputs are the [XLSX spec-gap audit](../spec-gap-audit.md:335), the staged
[pivot extension design](xlsx-pivot-extensions-design.md:1), and the bounded
[newer-owner mapping review](xlsx-pivot-newer-owner-mappings.md:1).

## Decision

The named newer namespace candidates are not ready for typed emission or
editing. The checked-in MS-XLSX owner tables bind no exact `extLst/ext` owner
and `ext/@uri` for `implicitMeasureSupport`, `aggregationInfo`,
`featureSupportInfo`, `autoRefresh`, or the 2023/2025 subtotal and data-source
elements. Namespace, QName, schema type, or Open XML SDK metadata does not
fill that gap. A nearby legacy URI must not be reused.

The next supported family is therefore the audit's `pivotTableData` extension
(§2.4.63), whose physical owner, GUID, namespace, type, particles, and
non-worksheet closure are all normative and locally available. It is the next
bounded implementation target after `pivotTableServerFormats`; it is selected
because its mapping is proved, rather than because its namespace is newer.

The base ECMA-376 schema archive, specifically its nested
`OfficeOpenXML-XMLSchema-Transitional.zip/sml.xsd` member, defines the generic
`extLst`/`ext` shape and token-typed extension URI attribute. The archive is
[`ECMA-376-4_5th_edition_december_2016.zip`](<../../../3rdparty/specs/ECMA-376/ECMA-376-4_5th_edition_december_2016.zip>).
It does not bind these Microsoft payloads to product-specific GUIDs. The
owner and URI below are taken only from the checked-in MS-XLSX extension table
[`2.2 Extensions.md:2827`](<../../../3rdparty/specs/[MS-XLSX]/2%20Structures/2.2%20Extensions.md:2827>).

## Exact evidence and the blocked candidates

| Payload | Normative local evidence | Exact blocker |
|---|---|---|
| `pivotTableData` | Exact payload QName is `{http://schemas.microsoft.com/office/spreadsheetml/2010/11/main}pivotTableData`. Owner is `pivotTableDefinition/extLst/ext`; URI is `{44433962-1CF7-4059-B4EE-95C3D5FFCF73}`. `CT_PivotTableData` requires `rowCount`, `columnCount`, and `cacheId`, and has `pivotRow` min 1/unbounded; each row has `c` min 1/unbounded and each cell has exactly one `v` plus optional `x`. See [`2.2 Extensions.md:2827`](<../../../3rdparty/specs/[MS-XLSX]/2%20Structures/2.2%20Extensions.md:2827>), [`2.4 Global Elements.md:749`](<../../../3rdparty/specs/[MS-XLSX]/2%20Structures/2.4%20Global%20Elements.md:749>), and [`2.6 Complex Types.md:5494`](<../../../3rdparty/specs/[MS-XLSX]/2%20Structures/2.6%20Complex%20Types.md:5494>). | None for the bounded read/scalar-edit scope below. |
| `implicitMeasureSupport` | Exact QName is `{http://schemas.microsoft.com/office/spreadsheetml/2020/pivotNov2020}implicitMeasureSupport`; the global semantic value is boolean and the schema declaration is `xsd:boolean` at [`2.4 Global Elements.md:1093`](<../../../3rdparty/specs/[MS-XLSX]/2%20Structures/2.4%20Global%20Elements.md:1093>) and [`pivotNov2020 Schema.md:7`](<../../../3rdparty/specs/[MS-XLSX]/5%20Appendix%20A%20-%20Full%20XML%20Schema/5.29%20http---schemas.microsoft.com-office-spreadsheetml-2020-pivotNov2020%20Schema.md:7>). SDK metadata lists it as a possible cache-definition extension child at [`schemas_openxmlformats_org_spreadsheetml_2006_main.g.cs:47643`](<../../../3rdparty/Open-XML-SDK/generated/DocumentFormat.OpenXml/DocumentFormat.OpenXml.Generator/DocumentFormat.OpenXml.Generator.OpenXmlGenerator/schemas_openxmlformats_org_spreadsheetml_2006_main.g.cs:47643>). Physical cardinality is unresolved because its `ext` owner is unresolved. | No checked-in MS-XLSX owner-table row supplies the exact parent `ext` URI. The SDK child list supplies no GUID. |
| `aggregationInfo`, `featureSupportInfo` | Exact QNames are `{http://schemas.microsoft.com/office/spreadsheetml/2023/pivot2023Calculation}aggregationInfo` and `{http://schemas.microsoft.com/office/spreadsheetml/2023/pivot2023Calculation}featureSupportInfo`. Their semantic roles are a data-field aggregation and field feature support, with required attributes `aggregationType`/`sourceField` and `featureName`; the declarations are at [`2.4 Global Elements.md:1262`](<../../../3rdparty/specs/[MS-XLSX]/2%20Structures/2.4%20Global%20Elements.md:1262>) and [`2.4 Global Elements.md:1274`](<../../../3rdparty/specs/[MS-XLSX]/2%20Structures/2.4%20Global%20Elements.md:1274>). The type declarations are at [`pivot2023Calculation Schema.md:7`](<../../../3rdparty/specs/[MS-XLSX]/5%20Appendix%20A%20-%20Full%20XML%20Schema/5.40%20http---schemas.microsoft.com-office-spreadsheetml-2023-pivot2023Calculation%20Schema.md:7>) and [`pivot2023Calculation Schema.md:35`](<../../../3rdparty/specs/[MS-XLSX]/5%20Appendix%20A%20-%20Full%20XML%20Schema/5.40%20http---schemas.microsoft.com-office-spreadsheetml-2023-pivot2023Calculation%20Schema.md:35>). | No physical owner or `ext/@uri`; do not infer `dataField`, `pivotField`, or any legacy GUID from the semantic descriptions. |
| `pivotAreaReferenceSubtotals`, `pivotFieldSubtotalLineItems`, `pivotFieldSubtotals` | Exact QNames are `{http://schemas.microsoft.com/office/spreadsheetml/2023/pivot2023Calculation}pivotAreaReferenceSubtotals`, `{http://schemas.microsoft.com/office/spreadsheetml/2023/pivot2023Calculation}pivotFieldSubtotalLineItems`, and `{http://schemas.microsoft.com/office/spreadsheetml/2023/pivot2023Calculation}pivotFieldSubtotals`. The schema gives `pivotFieldSubtotals/subtotal` 0..unbounded, `pivotAreaReferenceSubtotals/subtotal` 1..unbounded, and `pivotFieldSubtotalLineItems/subtotalLineItem` 1..unbounded; subtotal types and item locations are required where their types specify them. See the declarations at [`pivot2023Calculation Schema.md:43`](<../../../3rdparty/specs/[MS-XLSX]/5%20Appendix%20A%20-%20Full%20XML%20Schema/5.40%20http---schemas.microsoft.com-office-spreadsheetml-2023-pivot2023Calculation%20Schema.md:43>), [`pivot2023Calculation Schema.md:63`](<../../../3rdparty/specs/[MS-XLSX]/5%20Appendix%20A%20-%20Full%20XML%20Schema/5.40%20http---schemas.microsoft.com-office-spreadsheetml-2023-pivot2023Calculation%20Schema.md:63>), and [`pivot2023Calculation Schema.md:81`](<../../../3rdparty/specs/[MS-XLSX]/5%20Appendix%20A%20-%20Full%20XML%20Schema/5.40%20http---schemas.microsoft.com-office-spreadsheetml-2023-pivot2023Calculation%20Schema.md:81>), with semantic roles at [`2.4 Global Elements.md:1354`](<../../../3rdparty/specs/[MS-XLSX]/2%20Structures/2.4%20Global%20Elements.md:1354>), [`2.4 Global Elements.md:1378`](<../../../3rdparty/specs/[MS-XLSX]/2%20Structures/2.4%20Global%20Elements.md:1378>), and [`2.4 Global Elements.md:1390`](<../../../3rdparty/specs/[MS-XLSX]/2%20Structures/2.4%20Global%20Elements.md:1390>). | No physical owner or `ext/@uri`; a QName and particle list cannot establish whether the payload belongs under a field, area, line-item, or another owner. |
| `autoRefresh` | Exact QName is `{http://schemas.microsoft.com/office/spreadsheetml/2024/pivotAutoRefresh}autoRefresh`; the global semantic value is boolean and the declaration is `xsd:boolean` at [`2.4 Global Elements.md:1286`](<../../../3rdparty/specs/[MS-XLSX]/2%20Structures/2.4%20Global%20Elements.md:1286>) and [`pivotAutoRefresh Schema.md:7`](<../../../3rdparty/specs/[MS-XLSX]/5%20Appendix%20A%20-%20Full%20XML%20Schema/5.41%20http---schemas.microsoft.com-office-spreadsheetml-2024-pivotAutoRefresh%20Schema.md:7>). SDK metadata again lists a possible cache-definition extension child at [`schemas_openxmlformats_org_spreadsheetml_2006_main.g.cs:47643`](<../../../3rdparty/Open-XML-SDK/generated/DocumentFormat.OpenXml/DocumentFormat.OpenXml.Generator/DocumentFormat.OpenXml.Generator.OpenXmlGenerator/schemas_openxmlformats_org_spreadsheetml_2006_main.g.cs:47643>). Physical cardinality is unresolved because its `ext` owner is unresolved. | No checked-in MS-XLSX owner-table row supplies the exact parent `ext` URI. The SDK metadata does not provide one. |
| `pivotCacheDataSource` | Exact QName is `{http://schemas.microsoft.com/office/spreadsheetml/2025/pivotDataSource}pivotCacheDataSource`; the global semantic value is a dynamic-array/formula data source at [`2.4 Global Elements.md:1366`](<../../../3rdparty/specs/[MS-XLSX]/2%20Structures/2.4%20Global%20Elements.md:1366>). The type has optional `xm:f` (max 1) or optional `x:ref`; the prose requires them to be mutually exclusive and `ref` to identify one dynamic-array cell. See [`pivotDataSource Schema.md:9`](<../../../3rdparty/specs/[MS-XLSX]/5%20Appendix%20A%20-%20Full%20XML%20Schema/5.45%20http---schemas.microsoft.com-office-spreadsheetml-2025-pivotDataSource%20Schema.md:9>) and [`2.6 Complex Types.md:10012`](<../../../3rdparty/specs/[MS-XLSX]/2%20Structures/2.6%20Complex%20Types.md:10012>). | No parent placement or `ext/@uri`; do not assume `pivotCacheDefinition` merely because the name contains “cache”. |

The negative result is also recorded with the search method and archive
boundary in [`xlsx-pivot-newer-owner-mappings.md`](xlsx-pivot-newer-owner-mappings.md:10).
Its native inventory is empty for these target tokens, so it cannot select a
URI or prove an Office producer/open-save path. The exact unresolved question
for all blocked candidates is: **which later normative owner table or
documented native compatibility profile binds each QName to its physical
`extLst/ext` parent, exact `ext/@uri`, MCE declarations, and relationship
closure?** Until that evidence exists, the candidates remain opaque and
preserved.

For the implementation sketch below, C444 means the exact `pivotTableData`
URI above, C510 means `{C510F80B-63DE-4267-81D5-13C33094786E}`, and C983 means
`{983426D0-5260-488c-9760-48F4B6AC55F4}`. These labels shorten already proved
GUIDs; they do not infer a new owner mapping.

## Proposed semantic surface

The public surface should reuse the semantic `PivotTableSelector` and the
existing semantic `PivotCacheId` wrapper. `litchi_xlsx::pivot` should re-export
the data view/edit/patch types, `PivotTableSelector`, `PivotCacheId`, and the
cell value/extra vocabulary alongside the existing exports at
[`pivot/mod.rs:23`](../../../crates/litchi-xlsx/src/pivot/mod.rs:23). The
workbook facade should add `Workbook::pivot_table_data`,
`Workbook::edit_pivot_table_data`, and the ordinary patch method beside the
existing methods at [`workbook/model.rs:879`](../../../crates/litchi-xlsx/src/workbook/model.rs:879);
the umbrella `litchi::xlsx` surface re-exports that crate at
[`litchi/src/lib.rs:368`](../../../crates/litchi/src/lib.rs:368). OPC relationship
IDs, part names, extension URIs, and XML prefixes remain private, as required
by [ADR 0001](../../adr/0001-priorities-and-api-layers.md) and the
physical/semantic ownership split in [ADR 0011](../../adr/0011-ooxml-physical-package-ownership.md).
The following is a design sketch, not an API commitment:

```rust
workbook.pivot_table_data(selector) -> Result<Option<PivotTableDataView>>
workbook.edit_pivot_table_data(selector) -> Result<PivotTableDataEdit>

PivotTableDataView {
    table_name() -> &str,
    cache_id() -> PivotCacheId,
    row_count() -> u32,
    column_count() -> u32,
    rows() -> impl Iterator<Item = PivotRowView>,
    cell(address: PivotCellAddress) -> Result<Option<PivotValueCellView>>,
}

PivotRowView {
    row_index() -> Option<u32>,
    cells() -> impl Iterator<Item = PivotValueCellView>,
}

PivotValueCellView {
    column_index() -> Option<u32>,
    kind() -> PivotCellType,
    value_text() -> &str,
    extra() -> Option<PivotValueCellExtraView>,
}

PivotTableDataEdit {
    set_value(address: PivotCellAddress, value: PivotCellValueEdit) -> Result<()>,
    set_blank(address: PivotCellAddress) -> Result<()>,
    set_extra(address: PivotCellAddress, edit: PivotValueCellExtraEdit) -> Result<()>,
}
```

`PivotValueCellView` exposes a logical row/column coordinate, the preserved
cell kind (`Boolean`, `Number`, `Error`, `Text`, `DateTime`, or `Blank`), the
cell's value text, and a typed optional `PivotValueCellExtraView`. The optional
row/column indexes report only authored `r`/`i` coordinates. An omitted `r` or
`i` is `None`; the API never substitutes source order or a dense matrix
coordinate. `cell(address)` resolves only an explicit, unique coordinate; an
omitted coordinate returns no semantic address, and a duplicate authored
coordinate is a typed diagnostic/read-only error. XML
`r`, `i`, `r:id`, and `ext/@uri` are not public selectors. Explicit attribute
presence is retained internally so that a `Clear` edit differs from an absent
source attribute.

The first edit capability is deliberately scalar:

* set the existing `v` text for `str`, `b`, and `e` cells after checking the
  `ST_SXVCellType` lexical rules; the explicit `set_blank` operation only
  writes an empty ST_Xstring to an existing required `v` and refuses to omit,
  remove, or create that element;
* set or clear existing `bc`, `fc`, `i`, `un`, `st`, and `b` extra attributes
  with `AttributeEdit::{Keep,Set,Clear}`; a missing attribute can be inserted
  only into an already-present, unambiguous `x` start tag, and `Set` refuses
  to create a missing `x` element;
* expose `in` as read-only in the first batch and accept it only for a view
  when `in < actual C510 serverFormat child count`;
* return a typed `UnsupportedValueKind` for `n` and `d` edits. The first batch
  does not parse or rewrite numeric/date-time text: the local MS-XLSX rule
  identifies numeric/dateTime semantics but does not provide a pivot lexical
  validator, and `f64`/`chrono` normalization would be an invented contract.

The [ST_SXVCellType rules](<../../../3rdparty/specs/[MS-XLSX]/2%20Structures/2.7%20Simple%20Types.md:979>)
define the boolean, error, string, date, numeric, and blank meanings. There is
no calculation, refresh, query execution, cube access, external connection
access, or rendered-output promise. Row/cell insertion, deletion, reordering,
count changes, and cache reconstruction are separate structural capabilities.

The cache/table/output relationship remains an inert protected object; this
surface never refreshes a connection or recalculates a PivotTable, consistent
with the spreadsheet boundary in [ADR 0007](../../adr/0007-office-object-models.md).

## Source and OPC read set

An opened snapshot must retain the source spans and relationship tokens needed
to prove the following graph. The package layer owns physical parts and
relationships; the XLSX layer owns the semantic checks, as required by
[ADR 0024](../../adr/0024-current-topology.md).

1. Read the workbook part and relationships, including `pivotCaches`, the
   workbook `pivotTableReferences` extension, every non-worksheet reference,
   every worksheet-owned PivotTable edge, and every worksheet table definition
   needed for the workbook-wide name index. The C983 workbook owner is the
   normative mapping in [`2.2 Extensions.md:3288`](<../../../3rdparty/specs/[MS-XLSX]/2%20Structures/2.2%20Extensions.md:3288>).
   The C983 `ext` and `pivotTableReferences` payload are each unique; a
   duplicate known owner or duplicate reference for one target is diagnostic.
2. Read the selected PivotTable XML and relationships. Its `pivotTableData`
   payload must be a direct child of the exact C444 `ext` and must follow the
   direct schema sequence `pivotRow+`; each row must contain direct `c+`, and
   each cell must contain direct `v` followed by optional `x`. Its incoming
   workbook edge must be exactly one internal PivotTable relationship, with no
   external target, query, fragment, or worksheet PivotTable edge. The
   transitional relationship is
   `http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable`
   and the strict relationship is
   `http://purl.oclc.org/ooxml/officeDocument/relationships/pivotTable`, as
   recorded in [`constants.rs:304`](../../../crates/litchi-opc/src/constants.rs:304).
   Accept both strict and transitional core namespaces and relationship types;
   do not normalize one dialect into the other. A worksheet-owned table is not
   a valid C444 owner.
3. Read the selected PivotTable's standard `rowItems`, `colItems`, `cacheId`,
   location, name, and relationship closure. Enumerate **all** PivotTable
   names, including worksheet-owned tables, and all ordinary worksheet table
   names/display names before selecting a typed view. The normative C444 name
   condition is uniqueness among **all PivotTables**. An ordinary worksheet
   table name/displayName collision is only a conservative selector policy: a
   colliding Name selector is ambiguous/refused and must use another
   unambiguous selector, but the collision does not invalidate the C444
   payload or change PivotTable-name uniqueness.
4. Resolve the workbook cache with the semantic `PivotCacheId`, then read the
   matching cache-definition part and its relationships. The cache side must
   satisfy the complete external-source, 725, and ABF5 closure below. The
   selected-table-to-cache edge is exactly one internal relationship whose
   type is either the transitional
   `http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition`
   or strict
   `http://purl.oclc.org/ooxml/officeDocument/relationships/pivotCacheDefinition`
   value; it has no query or fragment, and its target content type is exactly
   `application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml`
   (`CT SML_PIVOT_CACHE_DEFINITION`). The workbook-to-cache edge named by
   `pivotCache@r:id` has the same exact-one, internal, strict/transitional,
   no-query/no-fragment, and content-type requirements. Both edges must resolve
   to the same canonical cache-definition part used by the semantic
   `PivotCacheId`. The relationship values are recorded in
   [`constants.rs:294`](../../../crates/litchi-opc/src/constants.rs:294) and
   the content type in [`constants.rs:119`](../../../crates/litchi-opc/src/constants.rs:119).
   The selected-table edge check is bounded at
   [`server_formats/mod.rs:3538`](../../../crates/litchi-xlsx/src/pivot/server_formats/mod.rs:3538),
   and the workbook edge check at
   [`server_formats/mod.rs:1532`](../../../crates/litchi-xlsx/src/pivot/server_formats/mod.rs:1532).
5. If a cell has `x/@in`, read the same PivotTable's existing C510 payload and
   include its owner span in the read set. The C510 list must be present,
   direct, unique, and have `count == actual serverFormat child count`; a
   cell index is valid only when `in < actual count`. An absent/malformed list
   or ambiguous index leaves the cell readable but diagnostic/read-only. A
   source connection or connection part is read for closure when the external
   cache source contains a route; this design never refreshes it.

### Cache identity and external-source closure

The C444 `cacheId`, the selected table's core `pivotTableDefinition@cacheId`,
the workbook `pivotCaches/pivotCache@cacheId`, and the qualified
`x14:pivotCacheDefinition@pivotCacheId` under the exact 725 owner must be the
same semantic `PivotCacheId`. The workbook cache criteria require external
source type and a cache-ID-version extension
[`2.4 Global Elements.md:455`](<../../../3rdparty/specs/[MS-XLSX]/2%20Structures/2.4%20Global%20Elements.md:455>),
and C444 requires the cache-definition and cache-ID-version extensions
[`2.6 Complex Types.md:5494`](<../../../3rdparty/specs/[MS-XLSX]/2%20Structures/2.6%20Complex%20Types.md:5494>).

The typed C444 profile requires one direct 725 `ext` with one qualified
`x14:pivotCacheDefinition` payload in the
`http://schemas.microsoft.com/office/spreadsheetml/2009/9/main` namespace and
one direct ABF5 `ext` with one qualified
`pivotCacheIdVersion` payload in the 2010/11 namespace. The 725 payload's
`pivotCacheId` is optional in the base [`CT_PivotCacheDefinition:1707`](<../../../3rdparty/specs/[MS-XLSX]/2%20Structures/2.6%20Complex%20Types.md:1707>) schema, but it is required for this C444 typed closure;
if it is absent, duplicated, or unequal, return a diagnostic/read-only result
instead of substituting an OPC relationship ID. ABF5's
`cacheIdSupportedVersion` and `cacheIdCreatedVersion` are required unsigned
bytes with no defaults, per [`CT_PivotCacheIdVersion:5690`](<../../../3rdparty/specs/[MS-XLSX]/2%20Structures/2.6%20Complex%20Types.md:5690>).
The exact owner GUID rows are [`2.2 Extensions.md:3010`](<../../../3rdparty/specs/[MS-XLSX]/2%20Structures/2.2%20Extensions.md:3010>); direct qualified payloads, owner uniqueness, and the four-way ID equality are checked before projection.

The core cache definition must contain exactly one direct `cacheSource` with
`type="external"`. The ECMA `CT_CacheSource` choice and `connectionId` default
are schema facts; an absent/default `connectionId` is not a named route. If
the cache-source ext list contains F057
`{F057638F-6D5F-4E77-A914-E7F072B9BCA8}`, it must contain exactly one direct
qualified `sourceConnection` with required `name`; the decoded name must
resolve to exactly one workbook `connection@name`. If explicit
`cacheSource@connectionId` is present, it must resolve to exactly one
workbook `connection@id`; when both routes exist, they must resolve to the
same connection. A present but missing, duplicated, malformed, or unresolved
F057 route is a closure failure; an explicitly present empty ST_Xstring name is
allowed when that empty name uniquely resolves to exactly one workbook
`connection@name`. F057's absence is permitted for the external cache
profile. This empty-name rule is local to the F057 sourceConnection route and
does not inherit the C510 empty-name refusal. When either F057 or
`connectionId` routes to the workbook connections part, that workbook edge is
also exactly one internal relationship of the transitional
`http://schemas.openxmlformats.org/officeDocument/2006/relationships/connections`
or strict
`http://purl.oclc.org/ooxml/officeDocument/relationships/connections` type,
with no query or fragment, and its target
content type is exactly
`application/vnd.openxmlformats-officedocument.spreadsheetml.connections+xml`;
duplicates or a wrong type are diagnostic. The loader evidence is
[`server_formats/mod.rs:2900`](../../../crates/litchi-xlsx/src/pivot/server_formats/mod.rs:2900),
with relationship/content-type constants at
[`server_formats/mod.rs:70`](../../../crates/litchi-xlsx/src/pivot/server_formats/mod.rs:70),
and the F057 owner table is [`2.2 Extensions.md:3057`](<../../../3rdparty/specs/[MS-XLSX]/2%20Structures/2.2%20Extensions.md:3057>).
`sourceConnection` requires an `ST_Xstring` name below 65,536 characters at
[`CT_SourceConnection:2821`](<../../../3rdparty/specs/[MS-XLSX]/2%20Structures/2.6%20Complex%20Types.md:2821>).

For C444, C983, C510, 725, ABF5, and F057, compare `ext/@uri` as an XML
Schema `xsd:token`: apply XML Schema whitespace collapse (trim and collapse
whitespace runs) before owner matching, while retaining each raw URI lexical
span and attribute spelling for source preservation and no-op/inverse checks.
Do not replace the raw owner attribute with the collapsed comparison value.

The snapshot therefore includes raw workbook/worksheet/table/cache/connection
bytes, owner ranges for C983/C444/C510/725/ABF5/F057, MCE context, resolved
relationship edges, the all-table name index, and the selected semantic graph.
The ordinary selector is table name or position. `same_readset` may permit
unrelated package changes only when all of those closure members and owner
spans still match; a full source check remains available for an advanced
patch. This follows the snapshot, read/write-set, and atomic-commit rules in
[ADR 0003](../../adr/0003-snapshots-edits-and-patches.md).

## Validation, MCE, and closure

The reader must capture the raw `extLst/ext` owner, payload span, enclosing
namespace declarations, and MCE branch spans before the existing MCE
projection at
[`reader/package.rs:106`](../../../crates/litchi-xlsx/src/pivot/reader/package.rs:106),
[`reader/package.rs:181`](../../../crates/litchi-xlsx/src/pivot/reader/package.rs:181),
or [`reader/package.rs:242`](../../../crates/litchi-xlsx/src/pivot/reader/package.rs:242).
Unknown extension siblings, namespace declarations, prefixes, comments,
processing instructions, whitespace, and unselected MCE branches remain raw
source. A direct C444 payload outside `AlternateContent` is eligible for the
bounded scalar writer after closure succeeds. Every recognized closure payload
(C983, C510, 725, ABF5, and F057/sourceConnection) must likewise be direct,
outside `AlternateContent`, before it can establish an editable closure. A
C444 or closure payload inside an `AlternateContent` choice/fallback, an
MCE-ambiguous effective branch, or an ignorable-only branch is readable only as
diagnostic/read-only; no edit may rewrite the selected projection, silently
project closure identity, or synthesize a branch. This direct-versus-MCE
policy applies even when `process_part` can produce a selected semantic view,
and follows [ADR 0006](../../adr/0006-validation-security-and-compatibility.md).

The typed `pivotTableData` path accepts exactly one known C444 payload under
the matching URI. A missing URI, wrong URI/QName, duplicate known payload,
payload in a different extension owner, or malformed MCE branch is a typed
diagnostic/read-only result. Unknown `ext` elements are preserved. The
following normative checks are required before exposing an editable view:

* The table is a non-worksheet PivotTable. The workbook reference criteria in
  [`CT_PivotTableReferences:3861`](<../../../3rdparty/specs/[MS-XLSX]/2%20Structures/2.6%20Complex%20Types.md:3861>)
  require an `A1`-origin location, no edit/change records, a name unique among
  all workbook PivotTables, matching workbook cache identity, and no
  conditional-format records. Ordinary worksheet table names/display names
  remain in the selector index only as a conservative ambiguity policy: their
  collision does not add a normative C444 invalidity, while the all-PivotTable
  name-uniqueness condition remains normative.
* `pivotTableData` has required unsigned `rowCount`, `columnCount`, and
  `cacheId`, at least one `pivotRow`, and matching standard row/column item
  counts. The complete type is [`CT_PivotTableData:5494`](<../../../3rdparty/specs/[MS-XLSX]/2%20Structures/2.6%20Complex%20Types.md:5494>).
  The strict typed profile requires the **actual number of direct `pivotRow`
  elements to equal `rowCount`** and the actual `rowItems@count` to equal it.
* The direct particle sequence is exact: `pivotRow` children contain direct
  `c` children; each row's authored `c` count equals its required `count`, and
  the strict C444 profile requires **actual `c` count = row `count` =
  `columnCount`** for every row. `colItems@count` must equal `columnCount`.
  Every authored `r` must satisfy `r < rowCount`, and every authored `i` must
  satisfy `i < columnCount`. `r` and `i` are optional by schema, but an
  omitted coordinate is never inferred; duplicate authored `r` or `i`
  coordinates are diagnostic/refusal.
  These particles are in [`CT_PivotRow:5538`](<../../../3rdparty/specs/[MS-XLSX]/2%20Structures/2.6%20Complex%20Types.md:5538>)
  and [`CT_PivotValueCell:5574`](<../../../3rdparty/specs/[MS-XLSX]/2%20Structures/2.6%20Complex%20Types.md:5574>).
* `t` is one of the allowed `ST_SXVCellType` values and defaults to numeric
  when absent. A blank cell still has the required single `v` child, and its
  decoded ST_Xstring value must be exactly empty; an omitted `v` is invalid.
  The blank setter therefore changes only that existing `v` text to empty and
  never removes or creates the element.
  `b` values must be exactly `true` or `false`; `e` values must be exactly
  `#DIV/0!`, `#VALUE!`, `#NUM!`, `#N/A`, or `#GETTING_DATA`; and `str` values
  are capped at 65,535 decoded UTF-16 units. The value meanings and exact
  token set are [`ST_SXVCellType:979`](<../../../3rdparty/specs/[MS-XLSX]/2%20Structures/2.7%20Simple%20Types.md:979>).
* ST_Xstring values are decoded through the existing SpreadsheetML escape
  layer: XML entities are decoded by the XML reader, `_xHHHH_` sequences are
  decoded as UTF-16 units, surrogate pairs are required, and unpaired
  surrogates are rejected. On changed text, literal `_xHHHH_` is protected as
  `_x005F_xHHHH_`, XML-invalid controls are emitted as `_xHHHH_`, and XML
  escaping is applied by the owner serializer. The concrete helper is
  [`raw/strings.rs:790`](../../../crates/litchi-xlsx/src/raw/strings.rs:790) and
  [`raw/strings.rs:844`](../../../crates/litchi-xlsx/src/raw/strings.rs:844);
  the vendored ECMA base type is `ST_Xstring` as `xsd:string` in the nested
  `OfficeOpenXML-XMLSchema-Transitional.zip/shared-commonSimpleTypes.xsd`.
* `x` has at most one cell-extra record. Its optional format index, colors,
  and boolean flags follow [`CT_PivotValueCellExtra:5614`](<../../../3rdparty/specs/[MS-XLSX]/2%20Structures/2.6%20Complex%20Types.md:5614>). XML Schema booleans on the extra flags accept `true`, `false`, `1`,
  or `0`; absence remains distinguishable from an explicit default. The `bc`
  and `fc` values are validated as `ST_UnsignedIntHex` per
  [`CT_PivotValueCellExtra:5626`](<../../../3rdparty/specs/[MS-XLSX]/2%20Structures/2.6%20Complex%20Types.md:5626>).
  The ECMA `sml.xsd` defines this type as `xsd:hexBinary` restricted to four
  bytes: exactly eight hexadecimal digits after XML Schema whitespace
  normalization. It is not a variable-width hexadecimal integer. Authored
  changed values must retain that width; untouched lexical spellings remain
  source-preserved. The `in` value is
  validated as `unsignedInt` per [`CT_PivotValueCellExtra:5624`](<../../../3rdparty/specs/[MS-XLSX]/2%20Structures/2.6%20Complex%20Types.md:5624>): the local value-valid parser accepts a leading
  `+` on digits and an all-zero negative spelling such as `-0`, while rejecting
  a negative nonzero value; the strict bound remains `in < actual C510
  serverFormat count`. Raw lexical spellings are retained for no-op and
  inverse preservation. The server-format index is diagnostic only unless
  C510 has `count == actual serverFormat count` and the bound holds.
* The selected table's C444 `cacheId` equals the table, workbook, and 725
  `pivotCacheId` values under the four-way closure above. The cache and table
  relationship edges are resolved before semantic allocation; relationship
  IDs never enter the public view.

No new `ext`, URI, namespace declaration, relationship, row, cell, or MCE
branch is synthesized by the first writer. This keeps the edit within a
source span whose owner and inverse are known.

## Limits, inverse, and source preservation

The implementation should reuse the existing pivot source limits: 32 MiB
per part, 1 MiB per extension fragment, depth 256, 100,000 XML events,
1 MiB of attributes, bounded namespace declarations, and a maximum server
format list below `2^31`. The caller may lower these limits. Charge the
**actual authored row elements** and **actual authored cell elements** to
separate finite budgets; charge retained bytes, attributes, namespaces, and
events independently. `rowCount` and `columnCount` are logical unsigned
dimensions checked against scalar dimension caps, but they never trigger a
dense matrix allocation or a `rowCount * columnCount` budget charge. Count
equations are checked while streaming, before reserve or allocation. This
follows [ADR 0005](../../adr/0005-io-memory-and-performance.md).

The first transaction supports only scalar changes to existing source spans:
`v` for the safe lexical kinds and explicit Keep/Set/Clear operations for
existing cell-extra attributes other than `in`. It does not insert or remove
XML nodes, change counts or indexes, or create a missing server-format list.
A no-op transaction returns the original package bytes. A changed transaction
clones the package, patches only the selected C444 owner, reopens the
candidate with the ordinary reader, and compares the semantic target, read
set, and owner ranges before publishing.

The reversible patch stores the original and edited raw C444 owner bytes,
owner/context spans, namespace/MCE branch spans, the complete source/read-set
guards, and semantic before/after values for verification. `inverse` restores
the retained raw owner bytes under the original owner/context guards; semantic
values alone are insufficient to construct an exact inverse. Applying a
stale, foreign, signed, encrypted, or otherwise non-editable source returns a
typed error before mutation. Failed validation leaves both the source snapshot
and the output package unchanged. Untouched package members and unknown
extension content remain byte/source preserving; ZIP container
reserialization is not promised to be byte-identical. These are the atomic
and exact-inverse requirements in [ADR 0003](../../adr/0003-snapshots-edits-and-patches.md)
and [ADR 0006](../../adr/0006-validation-security-and-compatibility.md).

## Verification boundary

Synthetic fixtures are sufficient to prove the local schema and source-bound
mechanics. They should cover:

* C444 under the exact `pivotTableDefinition` owner, strict and transitional
  prefixes, a unique C983 owner and exactly one internal incoming workbook
  PivotTable edge, no worksheet edge/query/fragment, all-table name
  uniqueness, ordinary-table collision selector ambiguity without C444
  invalidation, non-worksheet closure, four-way cache-ID equality, unique
  725/ABF5 cache extensions with required attributes, and exact-one
  selected-table/workbook cache edges with the required strict/transitional
  types and cache-definition content type;
* authored-row and authored-cell counts that satisfy actual rows =
  `rowCount`, actual `c` per row = row `count` = `columnCount`, standard
  `rowItems`/`colItems` equations, direct particle order, authored `r <
  rowCount`/`i < columnCount`, omitted-coordinate diagnostics, duplicate
  coordinate refusal, all six cell kinds, the blank setter's empty-existing-`v`
  rule, boolean/error/ST_Xstring escape/control rules, `bc`/`fc`
  `ST_UnsignedIntHex`, and C510 `count == actual` plus value-valid unsignedInt
  `in < actual count` cases including `+digits` and `-0` raw preservation;
* complete external `cacheSource` routes: `type="external"`, unique F057
  `sourceConnection` when present, explicit `connectionId` resolution,
  consistency when both routes exist, an empty F057 name only when it uniquely
  resolves, exact-one workbook connections edge/type/content type when routed,
  and no query/fragment on internal edges;
* duplicate or wrong URI/QName/owner cases, worksheet ownership, malformed
  cache relationships, invalid lexical values, absent-extra tri-state edits,
  safe insertion only inside an existing `x`, read-only `n`/`d` edit refusal,
  missing C510 for `in`, XML Schema `xsd:token` whitespace collapse for every
  C444/C983/C510/725/ABF5/F057 URI while preserving raw spelling, and all
  finite-limit failures without dense-product allocation;
* direct-versus-MCE selected/fallback read-only policy for C444 and every
  closure payload, unknown extension siblings, prefix and whitespace
  preservation, no-op byte identity, changed-part locality, stale/read-set
  rejection, raw-owner/span-backed exact inverse, reopen semantic equality,
  and failure atomicity.

The current [native pivot inventory](xlsx-pivot-native-inventory.json:1) has
no target-element hits. Existing native-looking pivot files therefore cannot
support an Office producer or open/save claim for this extension. A later
native non-worksheet OLAP package containing C444, the C983 workbook
reference, the required cache closure, and raw XML/package hashes is required
before making that claim. Until then, support is limited to normative-schema
and synthetic source-bound evidence. The blocked newer namespace candidates
remain opaque, and the feature matrix should not advertise them. This evidence
boundary follows the no-support-claim checklist in
[ADR 0008](../../adr/0008-migration-and-verification.md).
