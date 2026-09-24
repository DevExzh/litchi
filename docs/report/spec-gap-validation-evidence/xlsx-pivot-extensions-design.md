# XLSX newer PivotTable extensions: validation evidence and staged design

Read-only design note prepared 2026-09-12. This note does not claim that any
of these extensions is implemented. It records the current gap, the exact
schema evidence available in the repository, and the evidence required before
an implementation can publish a changed package.

The design follows the lossless, bounded, source-bound rules in
[`docs/GOAL.md`](../../GOAL.md) and the accepted ADR hierarchy in
[`docs/adr/README.md`](../../adr/README.md). In particular, an unsupported
extension remains opaque and source-owned; a recognized edit is atomic and
source-checked; and no pivot refresh, formula evaluation, rendering, external
connection, or query execution is introduced.

## Revalidation result

The audit is correct that the named newer pivot elements have no current
XLSX implementation or feature-matrix row. The zero-match audit entry is at
[`spec-gap-audit.md:355`](../spec-gap-audit.md:355). A source search over
`crates/litchi-xlsx` also finds no target element names, and the existing pivot
facade only exports standard cache/table codecs from
[`pivot/mod.rs:17`](../../../crates/litchi-xlsx/src/pivot/mod.rs:17).

The package reader currently resolves the normal workbook, worksheet, pivot
table, cache-definition, and records relationships. It calls
`litchi_ooxml_common::mce::process_part` before parsing a table, cache, or
records part at
[`reader/package.rs:106`](../../../crates/litchi-xlsx/src/pivot/reader/package.rs:106),
[`reader/package.rs:181`](../../../crates/litchi-xlsx/src/pivot/reader/package.rs:181),
and [`reader/package.rs:242`](../../../crates/litchi-xlsx/src/pivot/reader/package.rs:242).
That is suitable for the existing semantic reader, but it is not sufficient
for a source-preserving extension edit: the extension owner, original source
span, namespace declarations, and MCE branch must be retained before the MCE
projection is used.

The revalidation found the following native-looking baselines:

* `test-data/ooxml/xlsx/ExcelPivotTableSample.xlsx` and the checked-in
  `Pivot1_*`, `Pivot2_*`, `Pivot3_*`, and `Pivot4_*` files;
* `test-data/poi/test-data/spreadsheet/ExcelPivotTableSample.xlsx`; and
* the Open XML SDK assets
  `OlapPivotA3.xlsx`, `OlapPivotC3.xlsx`, `RelationalPivotA1.xlsx` through
  `RelationalPivotB6.xlsx`, `NativePivotSourceWorkbook.xlsx`, and the
  `Pivot*.xlsx` files under
  `3rdparty/Open-XML-SDK/test/DocumentFormat.OpenXml.Tests.Assets/assets/
  TestDataStorage/v2FxTestFiles/spreadsheet/`.

An archive XML scan found no checked-in occurrence of
`pivotTableServerFormats`, `pivotTableData`, `cachedUniqueNames`,
`implicitMeasureSupport`, `aggregationInfo`, `featureSupportInfo`,
`autoRefresh`, `pivotAreaReferenceSubtotals`, `pivotCacheDataSource`,
`pivotFieldSubtotalLineItems`, or `pivotFieldSubtotals`. These fixtures are
useful standard and OLAP relationship baselines, but they are not positive
native interoperability evidence for the target extensions. That absence
gates producer-shape and Office open/save interoperability claims only. It
does not block the first schema-bounded `pivotTableServerFormats`
implementation, which can use exact checked-in MS-XLSX evidence plus
source-bound synthetic fixtures. `cachedUniqueNames` remains
diagnostic/partial because its item-count anchor is unresolved below. Native
Office files remain required before claiming Office interoperability for each
family, including at least one non-worksheet/OLAP pivot for the 2010/11
payloads and a dynamic-array/formula source for the 2025 payload.

## Exact vocabulary and ownership evidence

The extension integration rules available in the checked-in
[MS-XLSX extension text](<../../../3rdparty/specs/[MS-XLSX]/2%20Structures/2.2%20Extensions.md:2667>)
are decisive here. The pivot-table section maps only the following two target
payloads, and the cache-field section maps only `cachedUniqueNames`:

| Element | Qualified namespace | Physical owner and exact ext URI | Normative grammar and constraints |
|---|---|---|---|
| `pivotTableServerFormats` | `http://schemas.microsoft.com/office/spreadsheetml/2010/11/main` (the appendix establishes the namespace; §2.4.2 omits the target-namespace line) | Payload owner: `pivotTableDefinition/extLst/ext`, URI `{C510F80B-63DE-4267-81D5-13C33094786E}`. It is admissible only for a Non-Worksheet PivotTable resolved through the workbook `pivotTableReferences` owner, URI `{983426D0-5260-488c-9760-48F4B6AC55F4}`. The workbook reference is a relationship closure, not a second payload owner. | `CT_PivotTableServerFormats`: `serverFormat` of `x:CT_ServerFormat`, min 1/max unbounded; required unsigned `count`, equal to the child count; the collection MUST contain fewer than `2^31` elements. `CT_ServerFormat` has optional `culture` and `format` attributes of `x:ST_Xstring`; neither has a default. See §2.4.2, [CT_PivotTableServerFormats and CT_PivotValueCellExtra](<../../../3rdparty/specs/[MS-XLSX]/2%20Structures/2.6%20Complex%20Types.md:5624>), and the ECMA-376 `CT_ServerFormat` schema. |
| `pivotTableData` | `http://schemas.microsoft.com/office/spreadsheetml/2010/11/main` | `pivotTableDefinition/extLst/ext`, URI `{44433962-1CF7-4059-B4EE-95C3D5FFCF73}`. It is the `PivotValues` payload for the workbook-referenced Non-Worksheet PivotTable. | `CT_PivotTableData`: `pivotRow` min 1/max unbounded; required unsigned `rowCount`, `columnCount`, and `cacheId`. `rowCount` equals the associated `rowItems` count; `columnCount` equals the `rowItems`/`colItems` width rules; `cacheId` identifies an OLAP cache that also has the required cache-definition and cache-ID-version extensions. A row has required unsigned `count`, at least one `c`, and optional `r` within `rowCount`; each cell has exactly one `v`, optional `x`, optional `i` within `columnCount`, and optional `t` defaulting to `n`. Cell-extra `i`, `un`, `st`, and `b` default to false. See §2.4.63 and §2.6.133–§2.6.136. |
| `cachedUniqueNames` | `http://schemas.microsoft.com/office/spreadsheetml/2010/11/main` | `cacheField/extLst/ext`, URI `{4F2E5C28-24EA-4EB8-9CBF-B6C8F9C3D259}`. The payload belongs to one cache field; it MUST NOT exist unless the associated connection has `model="true"`. This is a diagnostic/partial target until its item-count anchor is resolved. | `CT_CachedUniqueNames`: `cachedUniqueName` min 1/max unbounded. Each child requires unsigned `index` and `name` (`x:ST_Xstring`); `index` is unique within the collection and MS-XLSX §2.6.127 says it is less than the ancestor `CT_Items@count`; `name` is at most 65,535 characters (implemented with the conservative decoded UTF-16-unit policy below). The `CT_Items` anchor is unresolved against the base schema: standard `CT_CacheField` exposes `sharedItems` (`CT_SharedItems`), while `CT_Items` is the pivot-field item type. Do not infer a literal `cacheField/items` child or silently substitute `sharedItems@count` until this semantic mismatch is resolved. No default. See §2.4.61 and §2.6.126–§2.6.127. |

Those URIs are not interchangeable with a namespace. The exact rows are in
[MS-XLSX §2.2.4.5](<../../../3rdparty/specs/[MS-XLSX]/2%20Structures/2.2%20Extensions.md:2827>)
and [§2.2.4.6](<../../../3rdparty/specs/[MS-XLSX]/2%20Structures/2.2%20Extensions.md:3010>).

### First-family closure: non-worksheet `pivotTableServerFormats`

The workbook relationship is part of the normative owner closure. The
workbook `extLst/ext` with URI
`{983426D0-5260-488c-9760-48F4B6AC55F4}` contains
the qualified `pivotTableReferences` element (shown with the conventional
`x15` prefix here); each qualified `pivotTableReference` has required
`r:id` and resolves through a workbook relationship to a PivotTable part.
The payload itself is still under that part's
`pivotTableDefinition/extLst/ext` with URI
`{C510F80B-63DE-4267-81D5-13C33094786E}`. The exact workbook owner mapping is
in [the MS-XLSX workbook extension table](<../../../3rdparty/specs/[MS-XLSX]/2%20Structures/2.2%20Extensions.md:3288>), and the global element and
complex-type definitions are in [the `pivotTableReference` entry](<../../../3rdparty/specs/[MS-XLSX]/2%20Structures/2.4%20Global%20Elements.md:3>) and [CT_PivotTableReferences/CT_PivotTableReference](<../../../3rdparty/specs/[MS-XLSX]/2%20Structures/2.6%20Complex%20Types.md:3861>).

A typed server-format view is therefore exposed only when all of these
conditions hold:

* the workbook extension has the exact URI above and one unique
  `pivotTableReference` whose `r:id` resolves through an **internal** workbook
  relationship of type
  `http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable`
  (or its strict `http://purl.oclc.org/ooxml/officeDocument/relationships/pivotTable`
  form) to the same PivotTable definition part as the payload. The target part
  has the PivotTable content type, has exactly one incoming workbook
  PivotTable relationship and exactly one matching `pivotTableReference`, and
  has no incoming worksheet PivotTable relationship or other worksheet edge;
* the target `pivotTableDefinition/location@ref` begins with `A1`;
* `pivotTableDefinition@enableEdit` is absent or false, and `pivotEdits` and
  `pivotChanges` are absent;
* the target PivotTable `name` is unique among workbook PivotTables;
* the workbook `pivotCaches` entry has a logical `cacheId`/`cacheID` matching
  the target table's `cacheId`, and its internal cache-definition relationship
  resolves to the same cache-definition part;
* the target PivotTable has no `conditionalFormats`; and
* the target is a Non-Worksheet PivotTable. The workbook reference collection
  has at least one and fewer than `2^31` entries, while the server-format list
  has at least one and fewer than `2^31` entries.

The cache side of that relationship is also part of the closure. The
cache-definition part reached by the workbook `pivotCaches/pivotCache@r:id`
MUST have all of the following:

* `pivotCacheDefinition/extLst` contains one recognized
  `ext@uri={ABF5C744-AB39-4b91-8756-CFA1BBC848D5}` with the qualified
  `pivotCacheIdVersion` child in
  `http://schemas.microsoft.com/office/spreadsheetml/2010/11/main`. Its required
  `cacheIdSupportedVersion` and `cacheIdCreatedVersion` attributes are
  `unsignedByte` values with no defaults. If the optional logical
  `x14:pivotCacheDefinition@pivotCacheId` is present under the direct
  `ext@uri={725AE2AE-9491-48BE-B2B4-4EB974FC3084}` owner, it must agree with
  the semantic `PivotCacheId` closure. This is an extension element, not an
  attribute on the core cache-definition root; absence must not be replaced
  with a physical relationship ID.
* `cacheSource@type="external"`, using the schema-valid external cache-source
  branch. The F057 `sourceConnection` extension is optional under §2.4.39. If
  `cacheSource/extLst` contains the recognized
  `ext@uri={F057638F-6D5F-4E77-A914-E7F072B9BCA8}`, it has exactly one
  qualified `sourceConnection` whose decoded `name` resolves to exactly one
  workbook `connection@name`. If an explicit standard `connectionId` route is
  also present, it must resolve to that same connection. An explicit numeric
  `connectionId` must resolve even when F057 is absent. Known payloads must be
  direct children of their exact admitted `ext` owner; unknown ext siblings
  remain opaque.

The ABF5 owner row is in [the pivot-cache extension table](<../../../3rdparty/specs/[MS-XLSX]/2%20Structures/2.2%20Extensions.md:3012>); the mandatory external-source requirements are in [the `pivotCaches` global element](<../../../3rdparty/specs/[MS-XLSX]/2%20Structures/2.4%20Global%20Elements.md:455>), and the exact ABF5 grammar is [CT_PivotCacheIdVersion](<../../../3rdparty/specs/[MS-XLSX]/2%20Structures/2.6%20Complex%20Types.md:5690>). The optional F057 mapping is in [the cache-source extension table](<../../../3rdparty/specs/[MS-XLSX]/2%20Structures/2.2%20Extensions.md:3057>). The ECMA `sml.xsd` `CT_CacheSource` choice requires `type` and gives `connectionId` a default of `0`; this default does not turn a missing route into a named source connection. The relationship/content-type constants used for the internal edges are recorded in [`litchi-opc` constants](../../../crates/litchi-opc/src/constants.rs:292).

The workbook `r:id`, the PivotTable part URI, and the cache-definition part
URI are physical relationship diagnostics. The ordinary API uses a semantic
PivotTable selector (for example, workbook PivotTable name or an existing
semantic ordinal) and a logical `PivotCacheId`; it never asks callers for an
OPC relationship ID or part URI. A logical `PivotCacheId` is the semantic
`cacheId`/`cacheID` value and is closed by the matching table `cacheId`,
workbook `pivotCaches/pivotCache@cacheId`, and resolved cache relationship.
It may be exposed as typed semantic metadata, but it must not be confused
with the physical OPC package identity.

The local schema evidence gives `CT_ServerFormat` optional `culture` and
`format` attributes, both `ST_Xstring`, with no defaults. In the checked-in
[ECMA-376 schema archive](<../../../3rdparty/specs/ECMA-376/ECMA-376-4_5th_edition_december_2016.zip>), the nested
`OfficeOpenXML-XMLSchema-Transitional.zip!/sml.xsd` declares those two
optional attributes on `CT_ServerFormat`; it does not declare a default. The
first batch reads the bounded list and permits only conservative scalar
set/replace/clear edits to those optional attributes when the source start tag
is unambiguous. It does not add, remove, or reorder `serverFormat` children,
change `count`, interpret a culture or number-format code, alter a
PivotValueCell, refresh, evaluate, or render. It validates `count` against the
actual child count before exposing the typed list and charges collection
limits before allocation. `pivotValueCellExtra@in` is treated as a zero-based
child index and must be less than the actual list length for typed closure.
That strict-subset rule is a conservative implementation subset, not a claim
that the local prose has been corrected: the prose says “between zero and
count”. A source with `in == count` therefore remains diagnostic/opaque and
cannot be changed until the residual inclusive/exclusive ambiguity is
resolved; it is never silently reinterpreted.

### `cachedUniqueNames` index-anchor reconciliation

The local primary evidence contains a real vocabulary mismatch. MS-XLSX
§2.6.127 names the bound as `CT_Items@count` on the ancestor
`CT_CacheField`, but the checked-in [ECMA-376 schema
archive](<../../../3rdparty/specs/ECMA-376/ECMA-376-4_5th_edition_december_2016.zip>)
(nested `OfficeOpenXML-XMLSchema-Transitional.zip`, `sml.xsd`) defines the
`CT_CacheField` sequence with `sharedItems` (`CT_SharedItems`), `fieldGroup`,
`mpMap`, and `extLst`; `CT_Items` is the type used by a `pivotField` item
collection. The checked-in generated SDK follows that base schema: its
[`CacheField` child model](<../../../3rdparty/Open-XML-SDK/generated/DocumentFormat.OpenXml/DocumentFormat.OpenXml.Generator/DocumentFormat.OpenXml.Generator.OpenXmlGenerator/schemas_openxmlformats_org_spreadsheetml_2006_main.g.cs:8535>) and
[`CacheField` metadata](<../../../3rdparty/Open-XML-SDK/generated/DocumentFormat.OpenXml/DocumentFormat.OpenXml.Generator/DocumentFormat.OpenXml.Generator.OpenXmlGenerator/schemas_openxmlformats_org_spreadsheetml_2006_main.g.cs:8714>) admit `SharedItems`, not `Items`; its
[`SharedItems.Count`](<../../../3rdparty/Open-XML-SDK/generated/DocumentFormat.OpenXml/DocumentFormat.OpenXml.Generator/DocumentFormat.OpenXml.Generator.OpenXmlGenerator/schemas_openxmlformats_org_spreadsheetml_2006_main.g.cs:42166>) is the cache-field item-count attribute.

Until a primary correction or native package establishes whether the MS-XLSX
wording means `sharedItems@count` (or another associated collection), the
design records the bound as **semantically unresolved**. A conforming fixture
may contain the schema-valid `cacheField/sharedItems` host and the extension,
but a fixture must not add `cacheField/items` merely to satisfy §2.6.127, and
tests must not claim that `sharedItems@count` is the normative bound without
evidence. The implementation can validate unsigned indices, uniqueness, and
the 65,535 decoded UTF-16-unit name limit now; index-to-item bounds,
insertion, and deletion remain gated on resolving this anchor.

### Connection and cache identity closure

The first family's external cache-source closure and the model-mode
prerequisite for `cachedUniqueNames` use related but distinct routes. The
external closure requires `cacheSource@type="external"` and the ABF5
cache-ID-version extension; the F057 name route is optional. When
`cacheSource/extLst` contains
`ext@uri={F057638F-6D5F-4E77-A914-E7F072B9BCA8}`, it must contain exactly one
qualified `sourceConnection` child in
`http://schemas.microsoft.com/office/spreadsheetml/2009/9/main` (shown with
the conventional `x14` prefix). Its required `name` is an `ST_Xstring` naming
a workbook `connection@name`, with the MS-XLSX length bound below 65,536
decoded UTF-16 units. The source-connection owner and URI are recorded in [the
MS-XLSX cache-source extension table](<../../../3rdparty/specs/[MS-XLSX]/2%20Structures/2.2%20Extensions.md:3057>), with the `sourceConnection` and connection complex types in [§2.6](<../../../3rdparty/specs/[MS-XLSX]/2%20Structures/2.6%20Complex%20Types.md:2821>). An absent F057 extension is valid for the external cache closure.

For the `cachedUniqueNames` model-mode association, either of two explicit
routes is valid. A F057/name route resolves `sourceConnection@name` to exactly
one workbook connection. If F057 is absent, an explicitly present standard
`cacheSource@connectionId` supplies the connection-only route and resolves its
unsigned logical value against `workbook/connections/connection@id`. The
resolved catalog connection must have
`connection/extLst/ext@uri={DE250136-89BD-433C-8126-D09CA5730AF9}` containing
the 2010/11 model connection metadata and `model="true"`, with its required
type/id constraints. The outer standard `connection@type` MUST be `5`, and
the decoded x15 model `connection@id` MUST be empty when `model="true"`.
The `model` attribute is optional with schema default `false`, so an absent
attribute does not satisfy this closure. `{D79990A0-CA42-45E3-83F4-45C500A0EAA5}` is a different
connection extension and does **not** satisfy the model predicate. The two
connection URI rows are in [the checked-in connection extension table](<../../../3rdparty/specs/[MS-XLSX]/2%20Structures/2.2%20Extensions.md:2671>), and the model constraints are in [CT_Connection](<../../../3rdparty/specs/[MS-XLSX]/2%20Structures/2.6%20Complex%20Types.md:4016>).

If both the F057/name route and an explicit standard `connectionId` are
present, they must resolve to the same catalog connection. A present F057
extension whose name route is missing, duplicated, or unresolved fails typed
closure; it cannot be silently replaced by the numeric route. Conversely, an
explicit `connectionId` route is sufficient when F057 is absent. A missing
`connectionId`, or a schema default of `0` that was not explicitly supplied,
does not by itself establish a named/model association. If neither valid route
exists, the `cachedUniqueNames` payload remains diagnostic/opaque rather than a
normative model-associated view. Do not compare a numeric value with a
connection name, an `r:id`, or an OPC relationship target. The existing
query-table numeric-ID check at
[`connections/package.rs:168`](../../../crates/litchi-xlsx/src/connections/package.rs:168)
is only an implementation analogy for resolving a numeric catalog ID; it is
not PivotTable-extension support or schema evidence. Every supplied route must
still be checked for consistency.

All `ST_Xstring` values in these closures use one documented conservative
decode policy. The XML parser first decodes XML entities; the raw XML input is
validated for XML characters at that boundary. Then
`decode_spreadsheet_text` decodes actual SpreadsheetML `_xHHHH_` escapes with
`_x005F_` escape suppression. An escape-like sequence with non-hex syntax is
preserved as literal text; an actual escape for an unpaired surrogate is
rejected. Escaped control code units such as `_x0001_` remain valid semantic
text and are not rejected merely because the decoded value is not an XML
character. For constraints expressed as “characters <= 65,535”, count decoded
UTF-16 code units (`value.encode_utf16().count()`), not UTF-8 bytes or Unicode
scalar values. This is a conservative implementation bound, not a claim that
the wire lexical encoding of `ST_Xstring` is UTF-16; a supplementary scalar
therefore costs two units. Existing source spans retain their raw lexical form
for a no-op. A changed value is encoded through one XML/SpreadsheetML escape
helper, which escapes XML-illegal controls and protects literal `_xHHHH_`
sequences; an edit is refused if the decoded value exceeds the bound rather
than silently normalizing it. The existing SpreadsheetML decoder is
[`decode_spreadsheet_text`](../../../crates/litchi-xlsx/src/raw/strings.rs:790),
the semantic encoder is
[`encode_spreadsheet_text`](../../../crates/litchi-xlsx/src/raw/strings.rs:846),
and the source-bound XML encoder is
[`escaped_xstring`](../../../crates/litchi-xlsx/src/source_attributes.rs:28).

The newer global elements have exact target namespaces and typed grammar, but
the checked-in MS-XLSX extension table contains no row assigning them an ext
URI or a physical parent. The elements and their schema are still useful
contract evidence:

| Element | Qualified namespace | Semantic owner stated by MS-XLSX | Ext URI / physical owner status | Grammar, defaults, and cardinality |
|---|---|---|---|---|
| `implicitMeasureSupport` | `http://schemas.microsoft.com/office/spreadsheetml/2020/pivotNov2020` | Data connection associated with a pivot cache | **Unresolved in checked-in MS-XLSX.** The local generated SDK metadata corroborates a pivot-cache-definition `extLst/ext` host, but it does not supply the URI. Do not emit this element until a primary mapping source records both values or a documented native compatibility profile is accepted; do not guess an URI from the namespace. | One `xsd:boolean` leaf; no schema default. Preserve `true`/`false`/`1`/`0` lexical form on no-op; canonicalize only on an explicitly supported edit. See §2.4.91 and appendix §5.29. |
| `aggregationInfo` | `http://schemas.microsoft.com/office/spreadsheetml/2023/pivot2023Calculation` | PivotTable data-field item | **Unresolved URI and physical owner.** The semantic candidate is the `dataField` extension, but this is not an ext-URI claim. Promote only with primary mapping evidence or a documented native compatibility profile. | `CT_AggregationInfo` is empty-content with required `aggregationType` (`distinctCount`, `median`, `distinctDuplicates`, `countValuesDuplicated`, or `countRepeatValues`) and required unsigned `sourceField`; no defaults. A present value makes the standard data-field `subtotal` SHOULD be ignored. See §2.4.104, §2.6.256, and §2.7.39. |
| `featureSupportInfo` | `http://schemas.microsoft.com/office/spreadsheetml/2023/pivot2023Calculation` | PivotTable field | **Unresolved URI and physical owner.** The semantic candidate is the `pivotField` extension, but this is not an ext-URI claim. Promote only with primary mapping evidence or a documented native compatibility profile. | `CT_FeatureSupport` is empty-content with required `featureName` (`xsd:string`); no default. A present value means a capable consumer SHOULD ignore the associated field. See §2.4.105 and §2.6.257. |
| `autoRefresh` | `http://schemas.microsoft.com/office/spreadsheetml/2024/pivotAutoRefresh` | Pivot cache, when its source changes | **Unresolved URI and physical owner.** The local generated SDK metadata corroborates a pivot-cache-definition extension host but does not supply the URI. Promote only with primary mapping evidence or a documented native compatibility profile. The metadata remains inert in this project. | One `xsd:boolean` leaf; no schema default. It records a product setting; it does not authorize refresh. See §2.4.106 and appendix §5.41. |
| `pivotAreaReferenceSubtotals` | `http://schemas.microsoft.com/office/spreadsheetml/2023/pivot2023Calculation` | PivotTable area reference | **Unresolved URI and physical owner.** Promote only with primary mapping evidence or a documented native compatibility profile. | `CT_PivotAreaReferenceSubtotals`: `subtotal` of `CT_PivotSubtotalType`, min 1/max unbounded. Each child requires `subtotalType` from `ST_AggregationType`; no defaults. See §2.4.111 and §2.6.266/§2.6.270. |
| `pivotCacheDataSource` | `http://schemas.microsoft.com/office/spreadsheetml/2025/pivotDataSource` | PivotTable data source for a dynamic-array reference or formula | **Unresolved URI and physical owner.** Do not infer that the 2025 namespace itself authorizes a `pivotCacheDefinition` host. Promote only with primary mapping evidence or a documented native compatibility profile. | `CT_PivotCacheDataSource`: optional `xm:f` (at most one), where `xm` is `http://schemas.microsoft.com/office/excel/2006/main`, plus optional `ref` of core `x:ST_Ref`; `f` and `ref` are mutually exclusive. `ref` must name one cell containing a dynamic array. No defaults. See §2.4.112 and §2.6.267/appendix §5.45. |
| `pivotFieldSubtotalLineItems` | `http://schemas.microsoft.com/office/spreadsheetml/2023/pivot2023Calculation` | PivotTable line | **Unresolved URI and physical owner.** The name alone is not enough to choose `pivotField` versus a pivot-area/table host. Promote only with primary mapping evidence or a documented native compatibility profile. | `CT_PivotTableSubtotalLineItems`: `subtotalLineItem` of `CT_PivotItemSubtotal`, min 1/max unbounded. Each item requires `subtotalType` from `ST_AggregationType` and unsigned `itemLocation`; no defaults. See §2.4.113 and §2.6.271/§2.6.269. |
| `pivotFieldSubtotals` | `http://schemas.microsoft.com/office/spreadsheetml/2023/pivot2023Calculation` | PivotTable field | **Unresolved URI and physical owner.** The semantic candidate is the `pivotField` extension, but this is not an ext-URI claim. Promote only with primary mapping evidence or a documented native compatibility profile. | `CT_PivotFieldSubtotals`: `subtotal` of `CT_PivotItemSubtotal`, min 0/max unbounded. Each item requires `subtotalType` and `itemLocation`; no defaults. See §2.4.114 and §2.6.268/§2.6.269. |

The exact namespace/schema references above are in [MS-XLSX §§2.4.91,
2.4.104–2.4.106, and 2.4.111–2.4.114](<../../../3rdparty/specs/[MS-XLSX]/2%20Structures/2.4%20Global%20Elements.md:1093>)
and [the complex-type sections](<../../../3rdparty/specs/[MS-XLSX]/2%20Structures/2.6%20Complex%20Types.md:5286>).
The 2023 enumeration is recorded in
[§2.7.39](<../../../3rdparty/specs/[MS-XLSX]/2%20Structures/2.7%20Simple%20Types.md:1531>).
The absence of a newer URI is itself evidence: a repository-wide search of
the checked-in MS-XLSX text for these element names, namespace suffixes, and
`Ext URI` rows finds only the global/schema definitions and the older rows
listed above. A primary mapping source, or a documented native compatibility
profile with raw owner evidence, must establish the physical owner and URI
before the implementation can move a newer row out of “unresolved”.

## MCE and extension-list contract

The local MS-XLSX introduction requires an `Ignorable` attribute,
`AlternateContent`, or `extLst` when an extension is used
([§2.2.4](<../../../3rdparty/specs/[MS-XLSX]/2%20Structures/2.2%20Extensions.md:2667>)).
The implementation should apply the following owner rules once the physical
parent/URI evidence is available:

1. Treat the standard owner’s `extLst` as an optional, single source span. The
   extension list and every `ext` child remain in document order. A known
   payload is recognized only by the pair `(physical owner, exact URI, QName)`;
   a matching QName under a different URI is opaque.
2. Retain unknown ext siblings, duplicate unknown URIs, original prefixes,
   namespace declarations, comments, processing instructions, and whitespace
   exactly when the owner is unchanged. A duplicate known `(owner, URI, QName)`
   is an ambiguous owner and must fail a changed edit atomically; a read may
   expose it only as an opaque diagnostic.
3. Preserve the `mc:Ignorable` token list and every `mc:Choice`/`mc:Fallback`
   branch. If a target payload is selected only through an
   `mc:AlternateContent` choice, the reader may project the selected branch for
   inspection, but a changed edit must refuse unless the exact selected branch
   and source span can be patched without dropping the other branch. Do not
   synthesize a new `AlternateContent` wrapper from a namespace guess.
4. Keep the transitional/strict main namespace and extension namespace
   lexical choices from the source. New XML may use the caller’s admitted
   conformance dialect only after the owner and namespace declarations have
   been proven; a no-op must retain the original bytes.
5. Run the existing MCE processor only for semantic projection after capturing
   the raw owner source. If MCE processing selects a branch whose source span
   cannot be mapped one-to-one, expose an inert read and return an atomic
   unsupported-source error for edits.

These rules deliberately do not treat an arbitrary future namespace as a
recognized pivot extension. The current pivot reader’s MCE projection must be
reworked behind a source-bound owner before any edit API is added.

## First fully normative family: `pivotTableServerFormats`

`pivotTableServerFormats` is the first independently bounded family for
implementation. Its exact payload owner, workbook non-worksheet reference,
internal relationship topology, ABF5/external-cache closure, logical cache
identity, count/index checks, and `CT_ServerFormat` scalar attributes are all
evidenced above. Native Office positives are still needed
for an Office producer/open-save claim, but their current absence does not
block this schema-bounded implementation.

The ordinary API should be contextual and selector-first, for example (the
selector names are illustrative and should follow the existing public
selector conventions):

```text
workbook.pivot_table(PivotTableSelector::Name("Pivot")).server_formats()
workbook.edit_pivot_table(PivotTableSelector::Name("Pivot"))
    .set_server_format(index, ServerFormatPatch {
        culture: ScalarTextEdit::Set("en-US"),
        format: ScalarTextEdit::Clear,
    })
```

The selector resolves a semantic workbook PivotTable, then its source-bound
definition and the logical `PivotCacheId` closure. It never accepts an OPC
`r:id`, part URI, or relationship target as an ordinary identity. The typed
view contains the bounded ordered server-format list and the optional decoded
`culture`/`format` strings; it does not expose XML nodes, extension wrappers,
or package parts. `ScalarTextEdit` is an explicit tri-state (`Keep`, `Set`, or
`Clear`), so an omitted update, a new value, and removal of an optional XML
attribute cannot be conflated.

The first writer is deliberately scalar-only:

* `set_server_format` may set, replace, or clear either optional `culture` or
  `format` attribute when its owner and start-tag source span are unambiguous.
  Adding or clearing one attribute is a scalar metadata edit: the writer
  inserts/removes only that attribute while preserving neighboring attributes,
  their order, and surrounding source markup. It preserves the original
  `ext` siblings, prefixes, namespace declarations, comments, processing
  instructions, whitespace, and lexical `count`.
* It does not add/remove/reorder `serverFormat` children, change `count`,
  rewrite a whole pivot definition, or normalize a number-format or culture
  value. A changed value uses the single documented XML/SpreadsheetML escape
  helper and the caller-configured decoded-text/work budget. `CT_ServerFormat`
  has no 65,535-character facet; an ambiguous source span or insertion point
  causes the scalar edit to be refused atomically.
* Before publication, the transaction revalidates the workbook
  `pivotTableReferences` relationship closure, all Non-Worksheet criteria,
  logical cache identity, `count == actual children`, the fewer-than-`2^31`
  list bound, and every typed `pivotValueCellExtra@in` reference. It charges
  limits before allocating from any source count. A source with the residual
  `in == count` ambiguity remains diagnostic and cannot be changed.
* The source-bound lifecycle is `Snapshot` → transaction → exact `Patch`:
  no-op returns the original bytes, a changed owner is reopened through the
  bounded reader before publication, stale/signed/encrypted/reused sources
  fail before mutation, and the patch has an exact inverse. If relationship,
  MCE, namespace, limit, or reopen validation fails, both input and output
  remain unchanged. No refresh, cube query, calculation, evaluation, or
  rendering is introduced.

The implementation should retain the raw owner span before the existing MCE
projection, following the source-preserving lifecycle already used by the
data-type-icon owner ([`Snapshot`](../../../crates/litchi-xlsx/src/data_type_icons/package.rs:19),
[`Transaction`](../../../crates/litchi-xlsx/src/data_type_icons/package.rs:275),
[`Patch`](../../../crates/litchi-xlsx/src/data_type_icons/package.rs:452)).

## Diagnostic/partial family: `cachedUniqueNames`

`cachedUniqueNames` has exact namespace, physical owner, URI, per-name bound,
and unsigned/unique index grammar, but it is **diagnostic/partial only** while
the MS-XLSX `CT_Items@count` wording remains inconsistent with schema-valid
`CT_CacheField/sharedItems`. The absence of native positives is a separate
interoperability gate; it does not cure or waive this semantic constraint.
The partial path can read the exact owner and perform an existing-name scalar
edit when all source spans and model connection routes are unambiguous. It
must not claim a normative index-to-item bound or expose insertion, deletion,
or index mutation until the anchor is resolved.

The ordinary API should be contextual and selector-first, for example (the
selector names are illustrative and should follow the existing public
selector conventions):

```text
workbook.pivot_cache(CacheSelector::Id(PivotCacheId(semantic_cache_id)))
    .field(FieldSelector::Ordinal(field_ordinal))
    .cached_unique_names()
workbook.edit_pivot_cache(CacheSelector::Id(PivotCacheId(semantic_cache_id)))
    .field(FieldSelector::Ordinal(field_ordinal))
    .set_cached_unique_name(item_index, mdx_name)
```

The final names must follow the existing XLSX public-facade conventions, but
the ownership is fixed: `CacheSelector::Id(PivotCacheId(...))` is the primary
semantic cache selector and resolves through the workbook cache relationship
to one `pivotCacheDefinition` part. An existing semantic ordinal may remain a
convenience selector, but it is not a substitute for the logical ID. A
semantic field ordinal/name selects one `cacheField`, and the part-local ext
URI selects one `cachedUniqueNames` payload. The physical relationship ID and
part URI remain internal diagnostic closure. A diagnostic/partial view may
expose typed `index`/`name` values and a bounded collection, but must carry the
unresolved-anchor status and must not present the index bound as validated. It
must not expose `quick_xml`, `OpcPackage`, extension wrapper structs, or an
archive part as an ordinary API.

### Batch 2 integration-test API contract

The implementation is in progress. The advanced source-bound test surface is
`litchi_xlsx::pivot::cached_unique_names`; this contract does not claim that
the implementation has passed its production gate. Its public names are
`PivotCacheId`, `CacheSelector`, `FieldSelector`, `CachedUniqueName`,
`DiagnosticStatus`, `Snapshot`, `Transaction`, `Commit`, and `Patch`.

* `PivotCacheId(pub u32)` is the workbook semantic ID.
  `CacheSelector::Id(PivotCacheId(...))` selects that ID; `Position(usize)` is
  the optional source-order convenience selector.
* `FieldSelector::Ordinal(usize)` and `Name(&str)` select a cache field.
  Ambiguous field names must refuse.
* `CachedUniqueName` has public `index: u32` and `name: String` fields.
  `DiagnosticStatus::Unresolved` explicitly reports the unresolved item bound.
* `Snapshot::load(&OpcPackage, impl Into<CacheSelector>, FieldSelector<'_>)`
  returns `Result<Snapshot>`. `entries()` returns `&[CachedUniqueName]`, and
  `diagnostic_status()` returns `DiagnosticStatus`.
* `Transaction::new(&mut OpcPackage, impl Into<CacheSelector>, FieldSelector<'_>)`
  returns `Result<Transaction<'_>>`. Its
  `set_cached_unique_name(index: u32, name: impl AsRef<str>)` returns
  `Result<bool>`: the index is the semantic item index, not the vector ordinal,
  and `false` means the staged value was already equal. The setter validates
  borrowed text against the name and caller limits before allocating its
  staged copy; `&str` and owned `String` inputs both use this contract.
* `Transaction::commit()` returns `Result<Commit>`. `Commit::changed()` returns
  `bool`, `snapshot()` returns `&Snapshot`, and `patch()` returns `&Patch`.
  `Patch::inverse()` returns `Patch`; `apply(&mut OpcPackage)` returns
  `Result<()>` and must validate its source read set before publication.

`CacheSelector`, `FieldSelector`, and `DiagnosticStatus` may be aliases for
the longer names `PivotCacheSelector`, `PivotCacheFieldSelector`, and
`IndexBoundStatus`. The ordinary contextual workbook facade remains required;
this explicit OPC surface is only the advanced layer. The independent test
file is `crates/litchi-xlsx/tests/cached_unique_names.rs`.

The source-bound implementation should provide the same lifecycle already
used by the source-preserving data-type-icon owner
([`Snapshot`](../../../crates/litchi-xlsx/src/data_type_icons/package.rs:19),
[`Transaction`](../../../crates/litchi-xlsx/src/data_type_icons/package.rs:275),
[`Patch`](../../../crates/litchi-xlsx/src/data_type_icons/package.rs:452)):

* `Snapshot` retains the source package lineage, cache part identity, raw
  owner bytes, relationships, conformance dialect, limits, and the parsed
  typed list. An exact no-op returns the same source bytes and shares the
  source snapshot.
* A scalar name edit replaces only the original attribute value span when its
  lexical owner is unambiguous. It keeps ext siblings, prefix choices,
  attribute order, whitespace, and unrelated cache-field markup. It may not
  reorder or regenerate the whole cache definition.
  The required `name` attribute must remain present, but its `ST_Xstring`
  value may be empty: setting `""` writes `name=""`; it does not remove the
  attribute. Missing `name` is invalid; empty `name` is not.
* Insertion, deletion, or index changes are a separate structural operation.
  They are allowed only after the writer proves the `cachedUniqueNames`
  sequence, its parent `ext`, and the resolved cache-field item-count
  dependency closure. The base-schema/MS-XLSX `CT_Items` versus
  `CT_SharedItems` anchor is currently unresolved, so no structural operation
  may assume a literal `cacheField/items` child or silently use
  `sharedItems@count`; return a typed unsupported-structure error until the
  closure is evidenced. Existing-name scalar edits do not change the index or
  item collection and may proceed when the owner span is otherwise
  unambiguous.
* `Commit` validates the prospective collection before publication: either a
  valid F057/name route or an explicit connectionId-only route, the DE250
  model connection (`type=5`, decoded model `id` empty), and consistency when
  both routes are supplied; it also validates unique unsigned indices, the
  65,535 decoded UTF-16-unit name limit, and
  the resolved index bound once its item-count anchor is evidenced, together
  with ext URI/QName, owner span, namespace declarations, and relationship
  closure. A scalar name edit must not invent or rewrite the unresolved item
  collection.
  `Patch` carries source identity and has an exact inverse. Applying a patch to
  stale, signed, encrypted, or reused source state fails before mutation.
* A changed owner is staged and reopened through the ordinary bounded reader
  before package publication. Failure in parsing, source freshness, limits,
  signature policy, or relationship closure leaves the input package and the
  output sink unchanged.

The same owner can later carry `implicitMeasureSupport` and `autoRefresh` only
after primary mapping evidence or an accepted native compatibility profile
supplies their ext URI. Those booleans remain inert metadata; setting
`autoRefresh` never runs a refresh.

## Limits and validation obligations

Every extension parser must charge the caller’s package/XML/work limits before
allocating a collection. Reuse the existing XLSX source-owner scale
(`32 MiB` part input/output, `1 MiB` extension fragment, bounded depth/node
counts) unless a caller supplies a lower limit. The values are implementation
defaults, not a relaxation of the caller’s limits.

Apply these schema and cross-reference checks:

* `pivotTableServerFormats`: check required `count`, actual child count, the
  fewer-than-`2^31` rule, and each `pivotValueCellExtra@in` against the
  server-format collection using the conservative zero-based `< child-count`
  rule. Never allocate from `count` before the configured collection limit is
  charged. `culture` and `format` are opaque strings bounded by caller
  string/XML-work limits; `CT_ServerFormat` supplies no 65,535-character
  facet and the values are not interpreted.
* `cachedUniqueNames`: cap each name at 65,535 decoded UTF-16 units, cap the
  collection and retained extension bytes, reject duplicate unsigned indices,
  and require either a valid F057/name or explicit connectionId-only route
  together with the DE250 model connection closure. When both routes are
  present, they must resolve to the same connection before exposing a
  diagnostic typed payload; an absent/default-only `connectionId` is not an
  association.
  Apply the index-to-item bound only after the `CT_Items` versus
  `CT_SharedItems` semantic anchor is resolved; do not create or require a
  literal `cacheField/items` element. Existing-name scalar edits remain the
  only proposed mutation while that status holds.
* `pivotTableData`: check row and column counts, row `count`, optional row and
  cell indices, required `v`, and default cell type `n`. Independently cap
  logical dimensions and authored row/cell counts. The sparse representation
  charges its authored records and retained source closure; logical dimensions
  never trigger dense allocation or a product budget charge. Any future dense
  materialization must check `rowCount * columnCount` and admit its allocation
  before constructing the matrix. Resolve `cacheId` only through existing package
  relationships and the required OLAP cache extensions. A cached value edit
  is data-preserving metadata; it does not recalculate the PivotTable.
* 2023 subtotal payloads: validate enum values, required attributes, bounded
  unsigned integers, duplicate/ambiguous entries, and `sourceField` or
  `itemLocation` against the selected standard field/item collections once a
  primary mapping or accepted native compatibility profile establishes the
  owner. Do not interpret an aggregation as a request to compute it.
* `featureSupportInfo`: retain the feature string as inert text and enforce a
  bounded string size. “Should be ignored” is a consumer behavior rule, not a
  permission to delete the associated field.
* `implicitMeasureSupport` and `autoRefresh`: accept only XML Schema boolean
  lexical forms on read, retain lexical form for no-op, and avoid changing a
  cache connection or initiating a refresh.
* `pivotCacheDataSource`: enforce `f`/`ref` mutual exclusion, one formula child,
  bounded formula text, and the single-cell `ref` rule. Keep formulas opaque;
  no name resolution, spill calculation, cell reads, refresh, or external
  execution is allowed.

## Staged batches and acceptance gates

| Batch | Scope | Required evidence and boundary |
|---|---|---|
| 0 — evidence gate | Acquire native Office samples for an OLAP/non-worksheet table, a cache field with `cachedUniqueNames`, and each newer namespace. Record package hash, part URI, owner path, exact `ext@uri`, prefix declarations, MCE attributes, and raw XML spans. | No native interoperability claim until the samples are recorded. This gate does not block Batch 1's normative `pivotTableServerFormats` support, which can use the exact local schema/URI evidence plus synthetic source-bound fixtures. Native absence does not block a bounded implementation; it only blocks the corresponding Office producer/open-save claim. If a newer sample has a URI not present in the local spec, retain the sample and its primary citation; never fill the value from a memory or an unrelated extension. |
| 1 — first normative implementation | `pivotTableServerFormats`: typed bounded read plus source-preserving scalar set/replace/clear edits to optional `culture`/`format` attributes. | Synthetic source-bound fixtures cover the exact C510 payload URI, workbook `{983426D0-5260-488c-9760-48F4B6AC55F4}` `pivotTableReferences` owner, internal workbook PivotTable relationship/content type, exactly one incoming workbook reference, no worksheet edge, ABF5 cache-ID-version and external `cacheSource` closure with F057 absent and valid-present cases, semantic `PivotCacheId` closure, alternate prefixes, unknown siblings, duplicate known payloads, MCE branches, count/list/index limits, residual `in == count` diagnostic refusal, no-op bytes, changed bytes, stale/signed-package refusal, exact inverse, reopen, and strict/transitional roots. Structural list edits, refresh, evaluation, rendering, and native Office claims are out of scope. |
| 1b — ordered server-format leaves | `pivotTableServerFormats`: source-preserving insert/push/remove/move/reorder through the ordinary and advanced transactions. | Required count remains equal to the nonempty leaf list. Known direct `pivotValueCellExtra@in` references remap with leaf identity; orphaning removal, opaque/namespaced/wrapped references, unproven MCE ancestry, and the diagnostic boundary refuse structural publication. Unchanged references retain raw lexical spelling. Exact source inverse, stale-source refusal, namespace identity, and caller output/aggregate limits are covered by the retained [47-integration/12-unit isolated gate](xlsx-pivot-server-formats/list-crud-gate/README.md). This extends Batch 1; container/Part lifecycle, other pivot families, refresh, and native interoperability remain open. |
| 2 — diagnostic/partial cache family | `cachedUniqueNames`: bounded diagnostic read plus existing-name source-preserving scalar `name` edits only. | Fixtures cover F057/name, explicit connectionId-only, both-consistent, and mismatch routes against the DE250 model connection; they use semantic cache/field selectors, a schema-valid `cacheField/sharedItems` host, and the unresolved `CT_Items`/`CT_SharedItems` anchor without inventing `cacheField/items`. Validate unsigned/unique indices and decoded UTF-16-unit names, but make no index-to-item bound claim. Insertion, deletion, and index changes remain deferred. Native model-mode OLAP evidence is a separate interoperability gate. |
| 3 — old non-worksheet matrix family | `pivotTableData`, after the shared non-worksheet closure is proven; initially typed bounded read and inert scalar `v`/cell-extra edits only where source spans are unambiguous. | A primary mapping source or an accepted native compatibility profile proves cache-ID/version closure, row/column counts, cell index semantics, and part ownership. Sparse edits separately cap logical dimensions, authored records, retained bytes, and temporary replacement peaks. Any future dense materialization must check and charge the dimension product before allocation. No refresh, cube query, rendering, formula evaluation, or external connection. |
| 4 — cache scalar flags | `implicitMeasureSupport` and `autoRefresh` after each physical owner/URI gate closes. | A primary mapping source or an accepted native compatibility profile proves the physical `ext` owner, URI, prefix, boolean lexical forms, and coexistence with unknown cache extensions. Edits remain inert and source-preserving; `autoRefresh` never contacts a connection or performs a refresh. |
| 5 — 2023 calculation metadata | `aggregationInfo`, `featureSupportInfo`, `pivotAreaReferenceSubtotals`, `pivotFieldSubtotalLineItems`, and `pivotFieldSubtotals`, split by proven physical owner. | For every element, a primary mapping source or an accepted native compatibility profile proves owner, URI, parent/field or line identity, cardinality, duplicates, and cross-reference bounds. Read/write can expose attributes and subtotal lists, but never aggregates, ignores fields by deletion, or recalculates. |
| 6 — 2025 data source | `pivotCacheDataSource` read first; edit only a proven scalar `ref` or formula text after dependency closure is explicit. | A primary mapping source or an accepted native compatibility profile, including dynamic-array and formula samples where applicable, proves owner/URI, `xm:f` prefix and formula lexical preservation, `ref` single-cell behavior, mutual exclusion, and relationship/worksheet dependency handling. Formula and spill evaluation remain out of scope. |
| 7 — matrix and native closure | Update the feature matrix and audit only after batches pass. | Reopen saved output with ordinary bounded readers, compare raw no-op/unchanged owners, run strict/transitional tests and focused fuzz/limit tests, and record native Office open/save evidence. |

Until primary mapping evidence or an accepted native compatibility profile
supplies the newer physical mappings, a safe implementation may retain those
payloads as bounded opaque extension bytes. It must not claim typed support or
generate an ext URI for them. The first independently bounded and reviewable
implementation target is `pivotTableServerFormats`; the
`cachedUniqueNames` path remains diagnostic/partial until its index-count
anchor is resolved.
