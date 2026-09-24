# XLSX newer pivot owner and URI mappings

Status: bounded primary-source evidence, recorded 2026-09-12. This document
covers the eight newer SpreadsheetML pivot elements called out by the pivot
extension design: `implicitMeasureSupport`, `aggregationInfo`,
`featureSupportInfo`, `autoRefresh`, `pivotAreaReferenceSubtotals`,
`pivotCacheDataSource`, `pivotFieldSubtotalLineItems`, and
`pivotFieldSubtotals`.

The result is negative but concrete within the inspected sources: the
v20260108 §2.2.4.5 and §2.2.4.6 extension tables, the target global/type/schema
sections, and the local SDK metadata define no exact physical owner plus
`ext/@uri` for any of these eight elements. The global-element and schema
sections establish namespace, type, semantics, and child cardinality, but they
do not establish an extension-list placement. This is a bounded unresolved
finding, not a claim that no mapping can exist in an uninspected source. A
namespace, an `mc:Ignorable` token, an SDK QName, or a nearby legacy URI is
insufficient evidence to generate one.

## Evidence and bounded method

The primary current source is Microsoft's [MS-XLSX landing
page](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/2c5dee00-eff2-4b22-92b6-0738acd4475e).
Its Published Version table says 5/19/2026, revision 29.1 (minor). The
directly downloaded official PDF is
[`[MS-XLSX]-260108.pdf`](https://officeprotocoldoc.z19.web.core.windows.net/files/MS-XLSX/%5BMS-XLSX%5D-260108.pdf),
whose inspected header says `[MS-XLSX] - v20260108`, `Release: January 8,
2026`; its Revision Summary records `1/8/2026`, revision 29.1, Minor. Its
SHA-256 is
`81aeb7f35ce906eb673eabb05a6dfb5d36ec7a5fb1fb8caf7f82812cceab42cc`.
The landing-page publication date and the PDF release date are therefore
reported separately; they share revision number 29.1 and are not silently
treated as two different protocol revisions. The PDF's §7 Change Tracking
lists only §2.2.2 Formulas and §2.6.220 CT_ThreadedComments2Ext as minor
changes; it lists no pivot owner-table change.
The PDF and the checked-in specification were inspected at these exact
locations:

* §2.2.4.5, Pivot Table, where the specification lists the `pivotTableDefinition`,
  `pivotField`, `dataField`, `pivotHierarchy`, and `filter` extension owners
  and their URIs.
* §2.2.4.6, Pivot Table Cache Definition, where it lists the
  `pivotCacheDefinition`, `cacheField`, `cacheHierarchy`, `calculatedMember`,
  and `cacheSource` extension owners and their URIs.
* §§2.4.91, 2.4.104–2.4.106, and 2.4.111–2.4.114 for the target global
  elements.
* §§2.6.256–2.6.271 and schema §§5.29, 5.40, 5.41, and 5.45 for the target
  types and particles.

The local copies are [2.2 Extensions.md](<../../../3rdparty/specs/[MS-XLSX]/2%20Structures/2.2%20Extensions.md:2827>),
[2.4 Global Elements.md](<../../../3rdparty/specs/[MS-XLSX]/2%20Structures/2.4%20Global%20Elements.md:1093>),
and [2.6 Complex Types.md](<../../../3rdparty/specs/[MS-XLSX]/2%20Structures/2.6%20Complex%20Types.md:9736>).
The bounded search checked the two pivot extension tables, every target
global/type/schema definition, the local Open XML SDK generated metadata, and
the explicit 29-archive native baseline inventory in the
[inventory artifact](xlsx-pivot-native-inventory.json:1). That artifact records
581 archive members (573 `.xml`/`.rels`/`.vml` members), the exact archive
hashes, the sorted glob union, the 12 searched local-name/namespace tokens,
and an empty hit list. The archive scan read every selected XML-like member as
bytes, searching UTF-8, UTF-16LE, and UTF-16BE token encodings; it did not
infer absence from filenames. The native negative therefore
applies only to this listed 29-archive baseline set.

An exact mapping in this document means all of the following are recorded by
one primary source: the standard element whose `extLst` owns the extension,
the exact `ext/@uri` value, and the target payload QName. The semantic owner
descriptions in §2.4 are useful context but are not physical-owner claims.

## Target-by-target result

| Element | Normative namespace, type, and cardinality | Semantic owner stated by MS-XLSX | Physical owner and URI result |
|---|---|---|---|
| `implicitMeasureSupport` | `http://schemas.microsoft.com/office/spreadsheetml/2020/pivotNov2020`; §2.4.91; `xsd:boolean`; global leaf declaration | Data connection associated with a pivot cache | **Unresolved.** It has no row in §2.2.4.6. The SDK exposes it under the cache-definition `extLst/ext` class, but supplies no URI value. |
| `aggregationInfo` | `http://schemas.microsoft.com/office/spreadsheetml/2023/pivot2023Calculation`; §2.4.104; `CT_AggregationInfo`; required `aggregationType` and `sourceField` | PivotTable data-field item | **Unresolved.** It has no row in §2.2.4.5. The nearby `dataField` URI is for the payload QName `dataField`, not `aggregationInfo`. |
| `featureSupportInfo` | `http://schemas.microsoft.com/office/spreadsheetml/2023/pivot2023Calculation`; §2.4.105; `CT_FeatureSupport`; required `featureName` | PivotTable field | **Unresolved.** It has no row in §2.2.4.5. The nearby `pivotField` URI is for the payload QName `pivotField`, not `featureSupportInfo`. |
| `autoRefresh` | `http://schemas.microsoft.com/office/spreadsheetml/2024/pivotAutoRefresh`; §2.4.106; `xsd:boolean`; global leaf declaration | Pivot cache when source data changes | **Unresolved.** It has no row in §2.2.4.6. The SDK exposes it under the cache-definition `extLst/ext` class, but supplies no URI value. It is a 2024 namespace element; there is no normative 2020 `autoRefresh` target in the checked sources. |
| `pivotAreaReferenceSubtotals` | `http://schemas.microsoft.com/office/spreadsheetml/2023/pivot2023Calculation`; §2.4.111; `CT_PivotAreaReferenceSubtotals`; `subtotal` min 1/max unbounded, each requiring `subtotalType` | PivotTable area reference | **Unresolved.** The semantic area-reference description does not identify an extension-list host or URI, and no such row appears in §2.2.4.5. |
| `pivotCacheDataSource` | `http://schemas.microsoft.com/office/spreadsheetml/2025/pivotDataSource`; §2.4.112; `CT_PivotCacheDataSource`; optional `xm:f` max 1 (`xm` is `http://schemas.microsoft.com/office/excel/2006/main`) plus optional `ref` (`x:ST_Ref`) | PivotTable source using a dynamic-array reference or formula | **Unresolved.** No row appears in §2.2.4.6. The name does not authorize inference of a `pivotCacheDefinition` host. |
| `pivotFieldSubtotalLineItems` | `http://schemas.microsoft.com/office/spreadsheetml/2023/pivot2023Calculation`; §2.4.113; `CT_PivotTableSubtotalLineItems`; `subtotalLineItem` min 1/max unbounded, each a `CT_PivotItemSubtotal` | PivotTable line | **Unresolved.** No extension-table row identifies whether the physical host is a field, area reference, or another PivotTable record. |
| `pivotFieldSubtotals` | `http://schemas.microsoft.com/office/spreadsheetml/2023/pivot2023Calculation`; §2.4.114; `CT_PivotFieldSubtotals`; `subtotal` min 0/max unbounded, each a `CT_PivotItemSubtotal` | PivotTable field | **Unresolved.** The semantic field description does not turn the legacy `pivotField` URI into a mapping; §2.2.4.5 has no row for this QName. |

The five 2023 elements share a namespace, but a shared namespace does not
imply a shared physical host or a shared URI. Likewise, the 2025
`pivotCacheDataSource` schema has a cache-related name but does not state that
it is a child of `pivotCacheDefinition`.

## Exact legacy rows that must not be reused

The current [Pivot Table extension table](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/aafe627a-ed81-4c50-98ef-78084230b952)
(§2.2.4.5, last updated 2026-01-08) records these adjacent mappings:

| Physical owner | `ext/@uri` | Payload QName |
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

The [Pivot Table Cache Definition extension table](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/e5333e58-90f1-4f3e-8a80-399ea0a07f78)
(§2.2.4.6) similarly lists only these relevant cache owners:

| Physical owner | `ext/@uri` | Payload QName |
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

These are positive mappings for older payloads only. Reusing the `pivotField`,
`dataField`, or `pivotCacheDefinition` URI for a newer QName would change the
owner/URI identity and is unsupported by the cited source.

## Local SDK corroboration and its limit

The generated SDK is useful for checking ancestry and grammar, but it is not a
normative URI registry:

* [`PivotCacheDefinitionExtension`](<../../../3rdparty/Open-XML-SDK/generated/DocumentFormat.OpenXml/DocumentFormat.OpenXml.Generator/DocumentFormat.OpenXml.Generator.OpenXmlGenerator/schemas_openxmlformats_org_spreadsheetml_2006_main.g.cs:47643>)
  has a required generic `uri` attribute and lists
  `implicitMeasureSupport` and `autoRefresh` as possible children of the
  cache-definition `extLst/ext` class. The generated metadata contains no
  GUID value for either child. The matching
  [`ElementChildren.json`](<../../../3rdparty/Open-XML-SDK/test/DocumentFormat.OpenXml.Framework.Tests/ElementChildren.json:45195>)
  entry confirms those two child QNames.
* The generated 2023 file records the five target QNames and their
  `pivot2023Calculation` types, but has no registration tying them to a core
  `pivotField`, `dataField`, pivot-area, or other parent.
  [`schemas_microsoft_com_office_spreadsheetml_2023_pivot2023Calculation.g.cs`](<../../../3rdparty/Open-XML-SDK/generated/DocumentFormat.OpenXml/DocumentFormat.OpenXml.Generator/DocumentFormat.OpenXml.Generator.OpenXmlGenerator/schemas_microsoft_com_office_spreadsheetml_2023_pivot2023Calculation.g.cs:25>)
  is therefore QName/type evidence only.
* The generated 2025 file likewise defines
  `pivotCacheDataSource` and `CT_PivotCacheDataSource` without a core-parent
  registration. See [`schemas_microsoft_com_office_spreadsheetml_2025_pivotDataSource.g.cs`](<../../../3rdparty/Open-XML-SDK/generated/DocumentFormat.OpenXml/DocumentFormat.OpenXml.Generator/DocumentFormat.OpenXml.Generator.OpenXmlGenerator/schemas_microsoft_com_office_spreadsheetml_2025_pivotDataSource.g.cs:32>).

The local grammar has the expected schema details. In particular, the
[2020 schema](<../../../3rdparty/specs/[MS-XLSX]/5%20Appendix%20A%20-%20Full%20XML%20Schema/5.29%20http---schemas.microsoft.com-office-spreadsheetml-2020-pivotNov2020%20Schema.md:5537>)
and [2024 schema](<../../../3rdparty/specs/[MS-XLSX]/5%20Appendix%20A%20-%20Full%20XML%20Schema/5.41%20http---schemas.microsoft.com-office-spreadsheetml-2024-pivotAutoRefresh%20Schema.md:5864>)
declare the target `implicitMeasureSupport` and `autoRefresh` elements as booleans. The 2020 schema also declares `ignorableAfterVersion` and `dataFieldFutureData`. The [2023 schema](<../../../3rdparty/specs/[MS-XLSX]/5%20Appendix%20A%20-%20Full%20XML%20Schema/5.40%20http---schemas.microsoft.com-office-spreadsheetml-2023-pivot2023Calculation%20Schema.md:5>)
contains the five 2023 elements and their particles. The [2025 schema](<../../../3rdparty/specs/[MS-XLSX]/5%20Appendix%20A%20-%20Full%20XML%20Schema/5.45%20http---schemas.microsoft.com-office-spreadsheetml-2025-pivotDataSource%20Schema.md:5910>)
contains the optional `xm:f` plus `ref` form and no parent declaration.

## Native-fixture and implementation disposition

The [native inventory artifact](xlsx-pivot-native-inventory.json:1) is the
precise negative corpus used here: 29 explicitly selected `.xlsx` baselines
from the OOXML, POI, and Open XML SDK fixture roots, 581 total ZIP members,
and 573 XML-like members. It records SHA-256 values for every archive and
`archive_manifest_sha256=3eeb82092c9a54776ffa855b1b19d13a5c9709561663ae4fb5d3a07989649cad`.
The byte scan searched each XML-like member for all eight target local names
and the four target namespace suffixes in UTF-8, UTF-16LE, and UTF-16BE;
`hits` is empty. This rules out a
positive target payload in that bounded fixture set. It does not rule out a
producer mapping in an unlisted native package or in a separate specification
source.

Because this inspected fixture set contains no target payload, it supplies no
known-producer sample from which to derive all of the required owner, URI,
ancestry, payload, and artifact-hash facts. A native compatibility profile is
not established by this evidence.

Until a later primary mapping or a fully recorded native profile supplies the
missing pair, these elements should remain opaque at the extension boundary:

* retain an existing unknown `ext` and its payload byte-for-byte when the
  physical owner is unchanged;
* do not classify a payload by namespace alone;
* do not emit a new payload or select a URI from a semantic candidate;
* treat a later mapping as requiring both the exact physical owner and exact
  URI, with QName/cardinality evidence checked against the schema.

No production source or existing design/audit file was modified while collecting
this evidence.

## Source index

| Source | Exact version/date and use |
|---|---|
| [MS-XLSX landing page](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/2c5dee00-eff2-4b22-92b6-0738acd4475e) | Published Version table: revision 29.1, 2026-05-19; current version index and PDF link. |
| [Official MS-XLSX PDF](https://officeprotocoldoc.z19.web.core.windows.net/files/MS-XLSX/%5BMS-XLSX%5D-260108.pdf) | Inspected header: v20260108, Release 2026-01-08; Revision Summary: 1/8/2026, revision 29.1, Minor; SHA-256 recorded above. |
| [MS-XLSX change tracking](<../../../3rdparty/specs/[MS-XLSX]/7%20Change%20Tracking/Index.md:1>) | Revision 29.1 local log names only §2.2.2 Formulas and §2.6.220 CT_ThreadedComments2Ext; no pivot extension-table change. |
| [§2.2.4.5 Pivot Table](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/aafe627a-ed81-4c50-98ef-78084230b952) | Last updated 2026-01-08; explicit pivot-table owner/URI rows. |
| [§2.2.4.6 Pivot Table Cache Definition](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/e5333e58-90f1-4f3e-8a80-399ea0a07f78) | Last updated 2025-04-04; explicit cache owner/URI rows and no target newer rows. |
| [§2.4.104 aggregationInfo](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/001858c0-862a-4e80-9ebd-d3321bc9e685) | Last updated 2025-09-16; namespace, semantic owner, and type only. |
| [§2.4.105 featureSupportInfo](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/001e4ef3-7698-4269-b607-beab6f7d12e3) | Last updated 2024-02-20; namespace, semantic owner, and type only. |
| [§2.4.106 autoRefresh](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/3697b315-1ced-4478-a919-b383ac4dbe49) | Last updated 2024-04-16; confirms the 2024 namespace, not 2020. |
| [§2.4.112 pivotCacheDataSource](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/ef16f06d-eede-42f0-a88e-8c906d0017ba) | Last updated 2025-09-16; namespace and dynamic-array/formula type only. |
| [§2.4.114 pivotFieldSubtotals](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/3d8a1b46-0b18-4fa4-9e8e-192d81fbe4f5) | Last updated 2025-09-16; namespace and field semantics only. |
| [Local MS-XLSX front matter](<../../../3rdparty/specs/[MS-XLSX]/Front%20Matter.md:1>) | Checked-in revision history through 29.1 (2026-05-19); local section paths above are the repository's primary snapshot. |
| Local checked-in MS-XLSX sections | Revision 29.1 snapshot: §§2.2.4.5–2.2.4.6, 2.4.91/104–106/111–114, 2.6.256–271, schemas 5.29/5.40/5.41/5.45. |
| [Bounded native pivot inventory](xlsx-pivot-native-inventory.json:1) | 29 archives / 581 members / 573 XML-like members; explicit glob union and byte-scan method; 12 tokens; empty hit list; archive manifest digest recorded in the artifact. |
