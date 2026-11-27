# Word Binary (DOC) Feature Matrix

Feature checklist. ✅ Supported · 🟡 Partial · ❌ Unsupported · N/A Not applicable.

## Auxiliary structures

| Feature | Status | Read | Write |
| --- | --- | --- | --- |
| AutoText and attached glossary (`PlcfGlsy`, `SttbfGlsy`, `SttbGlsyStyle`, `LEGOXTR_V11`) | 🟡 | ✅ | 🟡 |
| Word 2003 XML schema references (`Hplxsdr`, `XSDR`, `TIQ`) | 🟡 | ✅ | ❌ |
| Structured document tags (`SttbfBkmkSdt`, SDT bookmark PLCs, `SDTI`) | 🟡 | ✅ | ❌ |
| Format consistency checker marks (`SttbfBkmkFcc`, `DPCID`) | 🟡 | ✅ | ❌ |
| Text Services Framework tables (`Plcfuim`, `PlfguidUim`, `UIM`) | 🟡 | ✅ | 🟡 |
| Grammar checker cookies (`Plcfcookie` and legacy cookies) | 🟡 | ✅ | 🟡 |
| Language auto-detection state (`Plcflad`, `LadSpls`) | 🟡 | ✅ | 🟡 |
| Table character cache (`PlcfTch`) | 🟡 | ✅ | 🟡 |
| E-mail revision threading (`RmdThreading`) | 🟡 | ✅ | ❌ |
| Repair bookmarks (`SttbfBkmkBPRepairs` and companion PLCs) | 🟡 | ✅ | 🟡 |
| Annotation bookmarks (`SttbfAtnBkmk`) | ✅ | ✅ | ✅ |
| Embedded font descriptions (`SttbTtmbd`) | 🟡 | ✅ | ❌ |
| VBA digital-signature storage (`DigitalSignature` PIDDSI and `StwUser` `Sign`/`SigAgile`/`SigV3`) | 🟡 | ✅ | 🟡 |
| Saved selection (`Selsf`, `fcWss`) | 🟡 | ✅ | 🟡 |
| Paragraph-group properties (`PGPArray`, `PGPInfo`, `PGPOptions`) | 🟡 | ✅ | ❌ |
| Print metadata (`PrDrvr`, `PrEnvPort`, `PrEnvLand`) | 🟡 | ✅ | ❌ |
| Frame/list records (`RgDofr`, `Dofr`) | 🟡 | ✅ | 🟡 |

## Core Word binary document model

| Feature | Status | Read | Write |
| --- | --- | --- | --- |
| `MsoEnvelopeCLSID` / `MsoEnvelope` metadata | ✅ | ✅ | ✅ |

## [MS-DOC] file structure and version

| Feature | Status | Read | Write |
| --- | --- | --- | --- |
| OLE2/CFB container and Word streams | ✅ | ✅ | ✅ |
| Word 97+ FIB and versioned FibRgFcLcb tables | ✅ | ✅ | ✅ |
| Versioned document properties (DopBase through Dop2013) | ✅ | ✅ | ✅ |
| Word 6/95 and pre-Microsoft-Office binary profiles | ❌ | ❌ | ❌ |
| WordDocument, table, Data, and story ranges | ✅ | ✅ | ✅ |
| Custom XML data storage | ✅ | ✅ | ✅ |
| Summary and document-summary property streams | ✅ | ✅ | ✅ |
| Reserved user-defined hyperlink metadata | ✅ | ✅ | ✅ |
| Encryption stream and password-to-open profiles | ✅ | ✅ | ✅ |
| XML signatures storage and signatures stream | ✅ | ✅ | ✅ |
| ObjectPool and embedded OLE/package objects | ✅ | ✅ | ✅ |
| Macros storage and VBA project payload | ✅ | ✅ | ✅ |
| Information Rights Management Data Space and protected content | 🟡 | ✅ | ❌ |
| Unknown streams and storages | 🟡 | ✅ | 🟡 |

## [MS-DOC] PLC, property, and auxiliary-family

| Feature family | Status | Read | Write |
| --- | --- | --- | --- |
| CP, PLC, STTB, SPRM/PRL, and property storage primitives | ✅ | ✅ | ✅ |
| Piece table, compressed/uncompressed Unicode text, FKPs, BTEs, and BinTable | ✅ | ✅ | ✅ |
| Main, footnote, endnote, header, comment, textbox, and header-textbox parts | ✅ | ✅ | ✅ |
| Bookmark PLCs and names | ✅ | ✅ | ✅ |
| Field PLCs and non-Plcfld text-only fields | ✅ | ✅ | ✅ |
| Field evaluation, document navigation, and generated results | 🟡 | 🟡 | 🟡 |
| Tables, nested tables, cell/row marks, merges, and TAP properties | ✅ | ✅ | ✅ |
| Character, paragraph, table, section, and picture SPRM families | ✅ | ✅ | ✅ |
| Stylesheet, list templates, list overrides, names, and numbering | ✅ | ✅ | ✅ |
| Section layout, columns, borders, page/line numbering, and note placement | ✅ | ✅ | ✅ |
| PLRSID, saved-by, associated strings, proofing, grammar, and spelling PLCs | ✅ | ✅ | ✅ |
| Auto-summary (Asumyi and PlcfAsumy) | ✅ | ✅ | ✅ |
| Captions and AutoCaption tables | ✅ | ✅ | ✅ |
| Smart tags and factoids (Plcffactoid, property bags, and related STTBs) | ✅ | ✅ | ✅ |
| Master-document subdocuments (PlcfWKB, WKB, FNPI, SttbFnm) | 🟡 | ✅ | 🟡 |
| Revision marks, author tables, and property revisions | ✅ | ✅ | ✅ |
| Document variables, attached-template and web-export metadata | 🟡 | ✅ | ✅ |
| Mail-merge settings, sources, filters, and recipients (Pms, Pmfs, Rfs, ODSO) | 🟡 | ✅ | ✅ |
| Legacy form fields and FFData | ✅ | ✅ | ✅ |
| ActiveX/OCX control semantics | 🟡 | ✅ | ✅ |
| Document protection settings and range-level protected bookmarks | 🟡 | ✅ | ✅ |
| Document statistics and DOP version metadata | ✅ | ✅ | ✅ |

## Drawings, equations, and payloads

| Feature | Status | Read | Write |
| --- | --- | --- | --- |
| OfficeArt/Escher drawing group and shape extraction | ✅ | ✅ | 🟡 |
| Pictures and BLIP payloads | ✅ | ✅ | ✅ |
| Floating shapes, textboxes, and header drawings | 🟡 | ✅ | 🟡 |
| Embedded OLE and package payloads | ✅ | ✅ | ✅ |
| MathType/MTEF Equation.3 objects | ✅ | ✅ | ✅ |
| Embedded font programs | 🟡 | ✅ | 🟡 |
| Generic rendering and pagination | ❌ | ❌ | ❌ |

## Explicit unsupported or intentionally inert families

| Feature | Status | Read | Write |
| --- | --- | --- | --- |
| Command bars, keymaps, menus, toolbars, and UI customizations | 🟡 | ✅ | ✅ |
| Routing slips and route-slip protection | 🟡 | ✅ | ✅ |
| Macro or control execution | ❌ | ❌ | ❌ |
| External document/database/include resolution | ❌ | ❌ | ❌ |
| Field calculation and generated tables/indexes | 🟡 | 🟡 | 🟡 |
| IRM license evaluation and protected-content access | ❌ | ❌ | ❌ |
| Complete later-version Word binary behavior | 🟡 | 🟡 | 🟡 |
