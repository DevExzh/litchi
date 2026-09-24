# XLSX later PivotTable extension metadata: normative review

Status: read-only specification review, 2026-09-19. This review checks the
current Microsoft publication after the audit's zero-match finding. It does
not add a production owner or authorize emission of any of the eight target
payloads.

## Finding

The current [MS-XLSX publication](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/2c5dee00-eff2-4b22-92b6-0738acd4475e)
is revision 29.1, dated 2026-05-19. Its [Pivot Table owner table](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/aafe627a-ed81-4c50-98ef-78084230b952)
(§2.2.4.5) and [Pivot Table Cache Definition owner table](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/e5333e58-90f1-4f3e-8a80-399ea0a07f78)
(§2.2.4.6) still contain no row for any of the eight target QNames. No
separate owner table or exact `ext/@uri` is supplied by the current
publication. The result from the earlier bounded native inventory is also
unchanged: the [29-archive scan](../xlsx-pivot-native-inventory.json) has no
target hits, so it supplies no producer URI or parent route.

This distinction matters: the element namespaces and schema grammars are
known, but the physical `extLst/ext` owner and URI are not. A semantic phrase
such as “PivotTable field”, an SDK prefix, or a nearby legacy URI is not an
owner mapping. The eight payloads must remain opaque and byte-preserved until
one source supplies both the physical owner and the exact URI.

## QName, grammar, and dependency contract

The following table records the facts that are normative independently of the
missing physical route. “Owner/URI” is intentionally `unresolved`; this is a
contract boundary, not an invitation to infer a route.

| Payload | QName and payload grammar | Defaults, cardinality, and reference closure | Owner / `ext/@uri` |
|---|---|---|---|
| `implicitMeasureSupport` (§2.4.91) | `{http://schemas.microsoft.com/office/spreadsheetml/2020/pivotNov2020}implicitMeasureSupport`; `xsd:boolean`. It describes whether the data connection associated with a PivotCache supports client-defined measures. | The boolean declaration has no schema default. Physical occurrence/cardinality is unresolved with the owner. Its semantic dependency is the PivotCache's data connection; no connection execution follows from this flag. | Unresolved. The generated SDK lists it under `pivotCacheDefinition/extLst/ext`, but that class only requires a generic token `uri` and gives no value. |
| `aggregationInfo` (§2.4.104) | `{http://schemas.microsoft.com/office/spreadsheetml/2023/pivot2023Calculation}aggregationInfo`; `CT_AggregationInfo`, a leaf with required `aggregationType: ST_AggregationType` and `sourceField: unsignedInt`. | No defaults. `sourceField` is a zero-based index into the PivotTable's `CT_PivotFields` collection. If present, the associated data-field item's `subtotal` attribute SHOULD be ignored. Physical occurrence is unresolved. | Unresolved. “Data field” does not authorize the legacy `dataField` route or URI. |
| `featureSupportInfo` (§2.4.105) | `{http://schemas.microsoft.com/office/spreadsheetml/2023/pivot2023Calculation}featureSupportInfo`; `CT_FeatureSupport`, a leaf with required `featureName: string`. | No default. The name carries PivotTable-field versioning information; if the feature is supported, the associated field SHOULD be ignored. Physical occurrence is unresolved. | Unresolved. “PivotTable field” does not authorize the legacy `pivotField` route or URI. |
| `autoRefresh` (§2.4.106) | `{http://schemas.microsoft.com/office/spreadsheetml/2024/pivotAutoRefresh}autoRefresh`; `xsd:boolean`. It describes whether the PivotCache refreshes automatically when source data changes. | The boolean declaration has no schema default. Physical occurrence/cardinality is unresolved with the owner. This is PivotCache metadata; it must not trigger source refresh or query execution. It is distinct from the unrelated external-links `autoRefresh` attribute, whose default is `false`. | Unresolved. The generated SDK lists it under `pivotCacheDefinition/extLst/ext`, but gives no URI value. |
| `pivotAreaReferenceSubtotals` (§2.4.111) | `{http://schemas.microsoft.com/office/spreadsheetml/2023/pivot2023Calculation}pivotAreaReferenceSubtotals`; `CT_PivotAreaReferenceSubtotals`. Its direct `subtotal` sequence uses `CT_PivotSubtotalType`. | `subtotal` is required 1..unbounded; each child requires `subtotalType: ST_AggregationType`. No defaults. The referenced PivotTable area must be resolved once the physical owner is known. | Unresolved. “Area reference” does not identify the owning core element or URI. |
| `pivotCacheDataSource` (§2.4.112) | `{http://schemas.microsoft.com/office/spreadsheetml/2025/pivotDataSource}pivotCacheDataSource`; `CT_PivotCacheDataSource`. The optional child is `xm:f` (`http://schemas.microsoft.com/office/excel/2006/main`, `ST_Formula`, max 1); the optional `ref` attribute is `x:ST_Ref` (`http://schemas.openxmlformats.org/spreadsheetml/2006/main`). | The schema gives both forms as optional and has no default. The prose makes them mutually exclusive: if `xm:f` is present, `ref` MUST NOT be present; if `ref` is present, it MUST identify one cell containing a dynamic array, and `xm:f` MUST NOT be present. Both absent is schema-valid. Formula references may need source-layout adjustment, but no formula evaluation or refresh is implied. | Unresolved. “Cache” does not prove a `pivotCacheDefinition` owner or URI. |
| `pivotFieldSubtotalLineItems` (§2.4.113) | `{http://schemas.microsoft.com/office/spreadsheetml/2023/pivot2023Calculation}pivotFieldSubtotalLineItems`; `CT_PivotTableSubtotalLineItems`. Its direct child is `subtotalLineItem: CT_PivotItemSubtotal`. | `subtotalLineItem` is required 1..unbounded. Each item requires `subtotalType: ST_AggregationType` and `itemLocation: unsignedInt`; the latter identifies a location in the associated collection. Physical collection bounds cannot be checked until the owner route is proved. | Unresolved. “PivotTable line” does not distinguish a line, field, area, or table owner. |
| `pivotFieldSubtotals` (§2.4.114) | `{http://schemas.microsoft.com/office/spreadsheetml/2023/pivot2023Calculation}pivotFieldSubtotals`; `CT_PivotFieldSubtotals`. Its direct child is `subtotal: CT_PivotItemSubtotal`. | `subtotal` is optional 0..unbounded. Each child requires `subtotalType: ST_AggregationType` and `itemLocation: unsignedInt`; the location must be checked against the associated collection after owner resolution. No defaults. | Unresolved. “PivotTable field” does not prove the legacy `pivotField` owner or URI. |

The 2023 `ST_AggregationType` enumeration is exactly
`distinctCount`, `median`, `distinctDuplicates`, `countValuesDuplicated`, and
`countRepeatValues`. It is not the library's general-purpose aggregation
enum. The complete particle and required-attribute grammar is in the
[official 2023 schema](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/bc0e64e7-19ba-49a8-8526-22960eacc4ba); the boolean and 2025
schemas are [5.29](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/a9f2fa9d-070f-4491-80e8-832783f0ada2), [5.41](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/6a531596-7c64-455b-93fa-7a8d7e3fca15), and [5.45](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/04ae0b01-21e2-49a1-8f77-46d3b9180217).

## URI guardrails

The current owner tables do publish adjacent, older mappings. They must not
be reused for a newer QName:

| Physical owner | Existing URI | Existing payload |
|---|---|---|
| `pivotTableDefinition/extLst/ext` | `{44433962-1CF7-4059-B4EE-95C3D5FFCF73}` | `pivotTableData` |
| `pivotTableDefinition/extLst/ext` | `{C510F80B-63DE-4267-81D5-13C33094786E}` | `pivotTableServerFormats` |
| `pivotField/extLst/ext` | `{2946ED86-A175-432A-8AC1-64E0C546D7DE}` | `pivotField` |
| `dataField/extLst/ext` | `{E15A36E0-9728-4E99-A89B-3F7291B0FE68}` | `dataField` |
| `pivotCacheDefinition/extLst/ext` | `{725AE2AE-9491-48BE-B2B4-4EB974FC3084}` | `pivotCacheDefinition` |
| `pivotCacheDefinition/extLst/ext` | `{ABF5C744-AB39-4b91-8756-CFA1BBC848D5}` | `pivotCacheIdVersion` |
| `cacheField/extLst/ext` | `{4F2E5C28-24EA-4EB8-9CBF-B6C8F9C3D259}` | `cachedUniqueNames` |

Consequently, a typed reader may only recognize one of the eight payloads
after it has proved the complete triple `(physical owner, exact URI, payload
QName)`, plus the target's namespace and grammar. An unknown `ext` must retain
its URI spelling, namespace declarations, MCE context, and sibling extensions
byte-for-byte. A typed writer must not create a new extension from the
namespace alone.

## Implementation recommendation

Do not start a typed owner for these eight elements from the current evidence.
The unresolved route is a production-readiness blocker: without it, a reader
cannot distinguish a valid payload from a payload attached to the wrong
PivotTable or PivotCache, and a writer cannot choose a conforming URI. The
correct interim behavior is bounded opaque preservation.

The next substantive pivot family should be a family with a proved owner and
cross-part closure. If more pivot work is desired, complete the remaining
structural and lifecycle surface around `pivotTableData` (§2.4.63): it is a
direct child of `pivotTableDefinition/extLst/ext` with URI
`{44433962-1CF7-4059-B4EE-95C3D5FFCF73}`, requires `rowCount`, `columnCount`,
and `cacheId`, and contains `pivotRow+`, each with `c+`. The specification
also requires its OLAP cache to carry the matching cache-definition and
cache-ID-version extensions. That is a real bounded implementation family
with a checkable workbook → PivotTable → cache relationship graph; it is
preferable to guessing a host for a 2023/2025 global element. The repository's
[C444 design](../xlsx-next-pivot-owner-design.md) records the exact closure and
its intentional non-refresh/non-calculation boundary.

## Source record

The local revision-29.1 snapshot supplies the exact line-level grammar:

* [`2.4 Global Elements.md`](../../../../3rdparty/specs/[MS-XLSX]/2%20Structures/2.4%20Global%20Elements.md) §§2.4.91, 2.4.104–2.4.106, and 2.4.111–2.4.114;
* [`2.6 Complex Types.md`](../../../../3rdparty/specs/[MS-XLSX]/2%20Structures/2.6%20Complex%20Types.md) §§2.6.256–2.6.271;
* [`2.2 Extensions.md`](../../../../3rdparty/specs/[MS-XLSX]/2%20Structures/2.2%20Extensions.md) §§2.2.4.5–2.2.4.6; and
* the generated SDK's [`PivotCacheDefinitionExtension`](../../../../3rdparty/Open-XML-SDK/generated/DocumentFormat.OpenXml/DocumentFormat.OpenXml.Generator/DocumentFormat.OpenXml.Generator.OpenXmlGenerator/schemas_openxmlformats_org_spreadsheetml_2006_main.g.cs) metadata, used only as non-normative ancestry corroboration.

The official pages above were checked after the prior 2026-09-12 mapping
review; the current revision adds no target owner/URI mapping. No production
source or feature matrix was modified for this review.
