# XLSX later PivotTable owner evidence

Status: bounded primary-source review, recorded 2026-09-12. This note is
read-only evidence for the unresolved later pivot extensions. It adds no
production support, changes no feature-matrix row, and makes no native
interoperability claim.

The reviewed targets are `implicitMeasureSupport` (§2.4.91),
`aggregationInfo` (§2.4.104), `featureSupportInfo` (§2.4.105), `autoRefresh`
(§2.4.106), and the 2025/pivot-calculation family
`pivotAreaReferenceSubtotals` (§2.4.111), `pivotCacheDataSource` (§2.4.112),
`pivotFieldSubtotalLineItems` (§2.4.113), and `pivotFieldSubtotals`
(§2.4.114). The existing [newer-owner mapping review](xlsx-pivot-newer-owner-mappings.md:1)
records the wider negative search; this note records the exact ancestry
distinction and the implementation boundary that follows from it.

## Result

No inspected primary source admits a complete mapping for any of the eight
targets. A complete mapping requires the target payload QName, the standard
element whose `extLst` owns the `ext`, and the exact `ext/@uri`. The current
Microsoft extension tables enumerate only older payload registrations. A
namespace, schema type, SDK class name, prefix, or semantic phrase such as
“PivotTable field” does not supply the missing URI.

The two boolean targets have a useful but non-normative ancestry hint in the
vendored Open XML SDK: `implicitMeasureSupport` and `autoRefresh` are listed
as possible children of the SDK's `PivotCacheDefinitionExtension`, which
corresponds to `pivotCacheDefinition/extLst/ext`. The same class requires a
generic token-valued `uri` attribute but contains no target-specific URI
constant. The five 2023 calculation elements and the 2025
`pivotCacheDataSource` class have no parent registration in the checked-in SDK
metadata. These are different findings from “the namespace is unknown”: the
QNames and schema grammars are known, while physical owner/URI identity is not.

## Normative QName and grammar evidence

The checked-in MS-XLSX snapshot is revision 29.1 (2026-05-19). The target
global declarations are in [2.4 Global Elements.md](<../../../3rdparty/specs/[MS-XLSX]/2%20Structures/2.4%20Global%20Elements.md:1093>),
the complex-type descriptions are in [2.6 Complex Types.md](<../../../3rdparty/specs/[MS-XLSX]/2%20Structures/2.6%20Complex%20Types.md:9736>),
and the full schemas are in §§5.29, 5.40, 5.41, and 5.45. The exact local
evidence is:

| Target | Exact payload QName and schema facts | Physical registration result |
|---|---|---|
| `implicitMeasureSupport` (§2.4.91) | `{http://schemas.microsoft.com/office/spreadsheetml/2020/pivotNov2020}implicitMeasureSupport`; `xsd:boolean`. The semantic description concerns whether the data connection for a pivot cache supports client-defined measures. See [global declaration](<../../../3rdparty/specs/[MS-XLSX]/2%20Structures/2.4%20Global%20Elements.md:1093>) and [§5.29 schema](<../../../3rdparty/specs/[MS-XLSX]/5%20Appendix%20A%20-%20Full%20XML%20Schema/5.29%20http---schemas.microsoft.com-office-spreadsheetml-2020-pivotNov2020%20Schema.md:1>). | No `extLst/ext` row or URI in either pivot extension table (§2.2.4.5 or §2.2.4.6). SDK ancestry points to the cache-definition extension only; URI is missing. |
| `aggregationInfo` (§2.4.104) | `{http://schemas.microsoft.com/office/spreadsheetml/2023/pivot2023Calculation}aggregationInfo`; `CT_AggregationInfo`; required `aggregationType` and `sourceField`. The element describes a data-field aggregation and says the associated base `subtotal` should be ignored. See [global declaration](<../../../3rdparty/specs/[MS-XLSX]/2%20Structures/2.4%20Global%20Elements.md:1262>), [type](<../../../3rdparty/specs/[MS-XLSX]/2%20Structures/2.6%20Complex%20Types.md:9736>), and [§5.40 schema](<../../../3rdparty/specs/[MS-XLSX]/5%20Appendix%20A%20-%20Full%20XML%20Schema/5.40%20http---schemas.microsoft.com-office-spreadsheetml-2023-pivot2023Calculation%20Schema.md:1>). | No row in either pivot extension table (§2.2.4.5 or §2.2.4.6). “Data field” is semantic context; it does not authorize reuse of the legacy `dataField` URI. No SDK parent registration. |
| `featureSupportInfo` (§2.4.105) | `{http://schemas.microsoft.com/office/spreadsheetml/2023/pivot2023Calculation}featureSupportInfo`; `CT_FeatureSupport`; required `featureName`. The semantic description concerns versioning information for a PivotTable field. See [global declaration](<../../../3rdparty/specs/[MS-XLSX]/2%20Structures/2.4%20Global%20Elements.md:1274>), [type](<../../../3rdparty/specs/[MS-XLSX]/2%20Structures/2.6%20Complex%20Types.md:9762>), and [§5.40 schema](<../../../3rdparty/specs/[MS-XLSX]/5%20Appendix%20A%20-%20Full%20XML%20Schema/5.40%20http---schemas.microsoft.com-office-spreadsheetml-2023-pivot2023Calculation%20Schema.md:1>). | No row in either pivot extension table (§2.2.4.5 or §2.2.4.6). “PivotTable field” does not establish a `pivotField/ext` owner or URI. No SDK parent registration. |
| `autoRefresh` (§2.4.106) | `{http://schemas.microsoft.com/office/spreadsheetml/2024/pivotAutoRefresh}autoRefresh`; `xsd:boolean`; semantic value says whether a pivot cache refreshes when source data changes. See [global declaration](<../../../3rdparty/specs/[MS-XLSX]/2%20Structures/2.4%20Global%20Elements.md:1286>) and [§5.41 schema](<../../../3rdparty/specs/[MS-XLSX]/5%20Appendix%20A%20-%20Full%20XML%20Schema/5.41%20http---schemas.microsoft.com-office-spreadsheetml-2024-pivotAutoRefresh%20Schema.md:1>). | No `extLst/ext` row or URI in either pivot extension table (§2.2.4.5 or §2.2.4.6). SDK ancestry points to the cache-definition extension only; URI is missing. Do not confuse this 2024 element with the core/external-links `autoRefresh` attribute. |
| `pivotAreaReferenceSubtotals` (§2.4.111) | `{http://schemas.microsoft.com/office/spreadsheetml/2023/pivot2023Calculation}pivotAreaReferenceSubtotals`; `CT_PivotAreaReferenceSubtotals`; direct `subtotal` children are 1..unbounded and each requires `subtotalType`. See [global declaration](<../../../3rdparty/specs/[MS-XLSX]/2%20Structures/2.4%20Global%20Elements.md:1354>), [type](<../../../3rdparty/specs/[MS-XLSX]/2%20Structures/2.6%20Complex%20Types.md:9986>), and [§5.40 schema](<../../../3rdparty/specs/[MS-XLSX]/5%20Appendix%20A%20-%20Full%20XML%20Schema/5.40%20http---schemas.microsoft.com-office-spreadsheetml-2023-pivot2023Calculation%20Schema.md:63>). | No row in either pivot extension table (§2.2.4.5 or §2.2.4.6). The phrase “area reference” does not identify which core element owns the extension list. No SDK parent registration. |
| `pivotCacheDataSource` (§2.4.112) | `{http://schemas.microsoft.com/office/spreadsheetml/2025/pivotDataSource}pivotCacheDataSource`; `CT_PivotCacheDataSource`; optional `xm:f` (0..1) and optional `ref`, mutually exclusive by the prose. `ref` must name one cell containing a dynamic array. See [global declaration](<../../../3rdparty/specs/[MS-XLSX]/2%20Structures/2.4%20Global%20Elements.md:1366>), [type](<../../../3rdparty/specs/[MS-XLSX]/2%20Structures/2.6%20Complex%20Types.md:10012>), and [§5.45 schema](<../../../3rdparty/specs/[MS-XLSX]/5%20Appendix%20A%20-%20Full%20XML%20Schema/5.45%20http---schemas.microsoft.com-office-spreadsheetml-2025-pivotDataSource%20Schema.md:1>). | No row in either pivot extension table (§2.2.4.5 or §2.2.4.6). The word “cache” does not prove a `pivotCacheDefinition` owner. No SDK parent registration. |
| `pivotFieldSubtotalLineItems` (§2.4.113) | `{http://schemas.microsoft.com/office/spreadsheetml/2023/pivot2023Calculation}pivotFieldSubtotalLineItems`; `CT_PivotTableSubtotalLineItems`; direct `subtotalLineItem` children are 1..unbounded and use `CT_PivotItemSubtotal`. See [global declaration](<../../../3rdparty/specs/[MS-XLSX]/2%20Structures/2.4%20Global%20Elements.md:1378>), [type](<../../../3rdparty/specs/[MS-XLSX]/2%20Structures/2.6%20Complex%20Types.md:10118>), and [§5.40 schema](<../../../3rdparty/specs/[MS-XLSX]/5%20Appendix%20A%20-%20Full%20XML%20Schema/5.40%20http---schemas.microsoft.com-office-spreadsheetml-2023-pivot2023Calculation%20Schema.md:81>). | No row in either pivot extension table (§2.2.4.5 or §2.2.4.6). “PivotTable Line” does not distinguish a line, field, area, or table owner. No SDK parent registration. |
| `pivotFieldSubtotals` (§2.4.114) | `{http://schemas.microsoft.com/office/spreadsheetml/2023/pivot2023Calculation}pivotFieldSubtotals`; `CT_PivotFieldSubtotals`; direct `subtotal` children are 0..unbounded and use `CT_PivotItemSubtotal`, whose `subtotalType` and `itemLocation` are required. See [global declaration](<../../../3rdparty/specs/[MS-XLSX]/2%20Structures/2.4%20Global%20Elements.md:1390>), [type](<../../../3rdparty/specs/[MS-XLSX]/2%20Structures/2.6%20Complex%20Types.md:10044>), and [§5.40 schema](<../../../3rdparty/specs/[MS-XLSX]/5%20Appendix%20A%20-%20Full%20XML%20Schema/5.40%20http---schemas.microsoft.com-office-spreadsheetml-2023-pivot2023Calculation%20Schema.md:43>). | No row in either pivot extension table (§2.2.4.5 or §2.2.4.6). The word “field” does not prove the legacy `pivotField` owner or URI. No SDK parent registration. |

The SDK-generated classes corroborate the QName/type/particle facts but are
not a normative URI registry. The 2023 generated file gives exact QNames and
types for `aggregationInfo`, `featureSupportInfo`, `pivotFieldSubtotals`,
`pivotAreaReferenceSubtotals`, and `pivotFieldSubtotalLineItems` at
[lines 25-30, 85-90, 137-142, 203-208, and 269-274](<../../../3rdparty/Open-XML-SDK/generated/DocumentFormat.OpenXml/DocumentFormat.OpenXml.Generator/DocumentFormat.OpenXml.Generator.OpenXmlGenerator/schemas_microsoft_com_office_spreadsheetml_2023_pivot2023Calculation.g.cs:25>),
but the generated file defines no core parent or `uri` value. The 2025 class
similarly defines only the standalone QName/type, `ref`, and `xm:f` particle
at [lines 32-92](<../../../3rdparty/Open-XML-SDK/generated/DocumentFormat.OpenXml/DocumentFormat.OpenXml.Generator/DocumentFormat.OpenXml.Generator.OpenXmlGenerator/schemas_microsoft_com_office_spreadsheetml_2025_pivotDataSource.g.cs:32>).

## Owner-table evidence and exact negative

The current official [MS-XLSX landing page](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/2c5dee00-eff2-4b22-92b6-0738acd4475e)
publishes revision 29.1 dated 2026-05-19. The current PDF is
[`[MS-XLSX].pdf`](https://officeprotocoldoc.z19.web.core.windows.net/files/MS-XLSX/%5BMS-XLSX%5D.pdf),
`v20260519`, SHA-256
`aadce100de55c55d16bd2af2a5ca4133111c833130b2a29365ed01f0f5389388`.
Its §2.2.4.5 table (PDF page 74/437) says the `pivotTableDefinition/extLst`
owner admits only these payload/URI pairs:

| Owner | `ext/@uri` | Payload local name |
|---|---|---|
| `pivotTableDefinition` | `{962EF5D1-5CA2-4C93-8EF4-DBF5C05439D2}` | `pivotTableDefinition` |
| `pivotTableDefinition` | `{44433962-1CF7-4059-B4EE-95C3D5FFCF73}` | `pivotTableData` |
| `pivotTableDefinition` | `{C510F80B-63DE-4267-81D5-13C33094786E}` | `pivotTableServerFormats` |
| `pivotTableDefinition` | `{E67621CE-5B39-4880-91FE-76760E9C1902}` | `pivotTableUISettings` |
| `pivotTableDefinition` | `{747A6164-185A-40DC-8AA5-F01512510D54}` | `pivotTableDefinition16` |
| `pivotField` | `{2946ED86-A175-432A-8AC1-64E0C546D7DE}` | `pivotField` |
| `dataField` | `{E15A36E0-9728-4E99-A89B-3F7291B0FE68}` | `dataField` |
| `pivotHierarchy` | `{F1805F06-0CD304483-9156-8803C3D141DF}` | `pivotHierarchy` |
| `filter` | `{0605FD5F-26C8-4aeb-8148-2DB25E43C511}` | `pivotFilter` |

The same table is available as the official [§2.2.4.5 Pivot Table page](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/aafe627a-ed81-4c50-98ef-78084230b952).
None of the five 2023 payload names appears in it. In particular, the
`pivotField` and `dataField` URIs identify those legacy payload QNames only;
they cannot be assigned to `featureSupportInfo`, `pivotFieldSubtotals`, or
`aggregationInfo`.

The official [§2.2.4.6 Pivot Table Cache Definition page](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/e5333e58-90f1-4f3e-8a80-399ea0a07f78)
(PDF page 76/437) admits only:

| Owner | `ext/@uri` | Payload local name |
|---|---|---|
| `pivotCacheDefinition` | `{725AE2AE-9491-48BE-B2B4-4EB974FC3084}` | `pivotCacheDefinition` |
| `pivotCacheDefinition` | `{5DA0FC9A-693D-419c-AD59-312A39285967}` | `timelinePivotCacheDefinition` |
| `pivotCacheDefinition` | `{ABF5C744-AB39-4b91-8756-CFA1BBC848D5}` | `pivotCacheIdVersion` |
| `cacheField` | `{63CAB8AC-B538-458D-9797-405883B0398D}` | `cacheField` |
| `cacheField` | `{4F2E5C28-24EA-4EB8-9CBF-B6C8F9C3D259}` | `cachedUniqueNames` |
| `cacheHierarchy` | `{8CF416AD-EC4C-4ABA-99F5-12A058AE0983}` | `cacheHierarchy` |
| `cacheHierarchy` | `{B97F6D7D-B522-45F9-BDA1-12C45D357490}` | `cacheHierarchy` |
| `calculatedMember` | `{0C70D0D5-359C-4A49-802D-23BBF952B5CE}` | `calculatedMember` |
| `calculatedMember` | `{57DEB092-E4DC-418E-9C9A-C0C97F8552CB}` | `calculatedMember` |
| `cacheSource` | `{F057638F-6D5F-4E77-A914-E7F072B9BCA8}` | `sourceConnection` |

No target row appears there. The checked-in copy is [2.2 Extensions.md](<../../../3rdparty/specs/[MS-XLSX]/2%20Structures/2.2%20Extensions.md:2827>),
with the pivot-cache table beginning at [line 3010](<../../../3rdparty/specs/[MS-XLSX]/2%20Structures/2.2%20Extensions.md:3010>).
The current PDF revision summary lists only §2.2.2 and §2.6.220 as minor
changes, so it supplies no later pivot-table owner-table update.

The base ECMA archive proves only the generic extension grammar. In
`ECMA-376-4_5th_edition_december_2016.zip`
(outer SHA-256
`bd25da1109f73762356596918bf5ff8b74a1331642dba5f1c1d1dfc6bed34ecd`), the
nested `OfficeOpenXML-XMLSchema-Transitional.zip` entry has SHA-256
`d34187520749998af306faf1b730e568b0ca6d88ad24638a407c0a9bb4ca04fc`, and its
`sml.xsd` entry has SHA-256
`495debc8fa967b77ed37799747b049832f2c95b2ecdb9f19ddcd68f8c9f96ab9`.
That `sml.xsd` defines `CT_Extension/@uri` as `xsd:token` and permits lax
wildcard content. It does not assign product-specific GUIDs to any later
QName. The archive is [vendored here](<../../../3rdparty/specs/ECMA-376/ECMA-376-4_5th_edition_december_2016.zip>).

## SDK ancestry evidence and its limit

The generated main-schema SDK class
[`PivotCacheDefinitionExtension`](<../../../3rdparty/Open-XML-SDK/generated/DocumentFormat.OpenXml/DocumentFormat.OpenXml.Generator/DocumentFormat.OpenXml.Generator.OpenXmlGenerator/schemas_openxmlformats_org_spreadsheetml_2006_main.g.cs:47642>)
is the `x:ext` type under `pivotCacheDefinition/extLst`. It lists
`xxpim:implicitMeasureSupport` and `xlpar:autoRefresh` as possible children,
and its metadata repeats those two QNames. The class requires `uri` and
validates it only as a token at [lines 47700-47740](<../../../3rdparty/Open-XML-SDK/generated/DocumentFormat.OpenXml/DocumentFormat.OpenXml.Generator/DocumentFormat.OpenXml.Generator.OpenXmlGenerator/schemas_openxmlformats_org_spreadsheetml_2006_main.g.cs:47700>).
The matching [ElementChildren.json entry](<../../../3rdparty/Open-XML-SDK/test/DocumentFormat.OpenXml.Framework.Tests/ElementChildren.json:45195>)
confirms the two QNames and their boolean type. There is no value such as a
GUID in either source. The generated core `PivotTableDefinitionExtension`
class lists only its known older payloads at [lines 46590-46602](<../../../3rdparty/Open-XML-SDK/generated/DocumentFormat.OpenXml/DocumentFormat.OpenXml.Generator/DocumentFormat.OpenXml.Generator.OpenXmlGenerator/schemas_openxmlformats_org_spreadsheetml_2006_main.g.cs:46590>);
its generic wildcard does not register any 2023 or 2025 target.

For the 2023 and 2025 target classes, a read of all generated
`ElementChildren.json` parent entries found no parent containing any of these
QNames. The 2023 file defines only the element classes and their own child
particles; the 2025 file defines `pivotCacheDataSource`, its optional `xm:f`,
and `ref`. SDK namespace prefixes (`xxpim`, `xlpar`, `xlpcalc`, `xlpds`, and
`xne`) are serialization aliases, not owner registrations or extension URIs.

## Native fixtures

The durable native baseline is the [pivot native inventory](xlsx-pivot-native-inventory.json:1):
29 explicitly selected archives, 581 total members, 573 XML-like members,
archive-manifest SHA-256
`3eeb82092c9a54776ffa855b1b19d13a5c9709561663ae4fb5d3a07989649cad`, and
`hits: []` for all eight local names and four namespace suffixes. The scan read
XML-like members as bytes in UTF-8, UTF-16LE, and UTF-16BE, so there is no
positive archive/member hash to cite for a target payload.

As a cheap supplemental check, a read-only sweep over the vendored
`3rdparty/` and `test-data/` zip-like fixtures examined 1,269 readable
archives and 20,919 members and found no UTF-8 byte hit for any target token.
Twenty-nine malformed or intentionally damaged archives could not be read;
this one-off sweep has no manifest and is corroboration only. It does not
establish a producer mapping, an Office open/save path, or a URI.

## Actionable boundary

The next actionable owner candidate remains the already proved C444
`pivotTableData` extension: `pivotTableDefinition/extLst/ext` with
`ext/@uri={44433962-1CF7-4059-B4EE-95C3D5FFCF73}`. Its exact owner, URI, QName,
schema particles, and non-worksheet closure are recorded in the
[next-owner design](xlsx-next-pivot-owner-design.md:1). The later targets in
this note should remain opaque until a primary mapping or an accepted native
compatibility profile supplies the missing owner/URI pair.

Once such a mapping exists, the typed boundaries should be:

* `implicitMeasureSupport` and `autoRefresh`: read and inert scalar edit only
  after validating one cache-definition owner, one matching `ext/@uri`,
  namespace/QName, boolean lexical form, and any required cache/connection
  closure. `autoRefresh` never contacts a source or performs a refresh.
* `aggregationInfo`: read/edit required `aggregationType` and `sourceField`
  only after the exact data-field owner and the referenced pivot-field index
  are proved. Do not calculate or rewrite PivotTable values.
* `featureSupportInfo`: read/edit the required `featureName` only after the
  exact field owner and feature semantics are admitted. Preserve the field
  when a producer's “SHOULD be ignored” behavior cannot be executed.
* `pivotFieldSubtotals`, `pivotAreaReferenceSubtotals`, and
  `pivotFieldSubtotalLineItems`: read/edit typed subtotal records only after
  the corresponding field, area-reference, or line owner is proved; validate
  required subtotal attributes and item/field bounds. No aggregation or
  rendered-output promise is implied.
* `pivotCacheDataSource`: read the mutually exclusive `xm:f`/`ref` form only
  after its physical owner and relationship/dependency closure are proved.
  Formula text and the single-cell `ref` may be preserved or edited with
  lexical and dependency checks; dynamic-array spill, formula evaluation,
  cache refresh, and query execution remain outside the surface.

At the package boundary, preserve the complete unknown `ext` span,
`ext/@uri` lexical token, namespace declarations/prefixes, MCE declarations,
and sibling extensions byte-for-byte whenever the physical owner is not
proven. Do not classify these payloads by namespace alone and do not emit a
new URI guessed from a legacy row. A future native profile must record the
archive SHA-256, ZIP entry, owner ancestry, exact URI spelling, payload QName,
MCE admission, and coexistence with unknown siblings before it can authorize
typed emission or editing.
