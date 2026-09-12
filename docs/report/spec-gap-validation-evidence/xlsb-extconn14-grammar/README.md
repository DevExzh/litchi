# MS-XLSB `BrtBeginExtConn14` grammar evidence

**Checked:** 2026-09-12 UTC
**Scope:** evidence only; no production source, fixtures, or commits were changed.

## Result

The named `BrtBeginExtConn14` fields and the complete enclosing record
production are pinned. The official Microsoft **MS-XLSB v20250916** (protocol
revision 21.0) PDF prints the enclosing production. The current **v20251113**
(revision 21.1) PDF
prints the same §2.4.78 text and packet, but replaces the External Data
Connections ABNF with an internal validation-grammar path. Therefore the
status depends on the required release gate:

* **Pinned for the published v20250916 field grammar and envelope:** yes. The
  relevant production is
  `EXTCONN14 = BrtBeginExtConn14 [PCDCALCMEMSEXT] BrtEndExtConn14`, and the
  §2.4.78 packet plus field descriptions define the complete
  `FRTHeader`/`irstCulture`/`irstClientCubeUrn` payload sequence.
* **Independently downloadable v20251113 ABNF source:** no. The latest PDF
  delegates the part-level ABNF to internal files, but the v20251113 §2.4.78
  section is byte-for-byte identical to v20250916 after text extraction. The
  published v20250916 grammar is therefore the pinned version for this
  existing record family.
  The PDF delegates §2.1.7.24 to
  `O25FileFormatDefinitions\\excel12\\cpp\\ValidationABNFs\\Biff12ExternalDataConnectionsGrammar.abnf`,
  and §2.1.8 delegates common rules to
  `O25FileFormatDefinitions\\excel12\\cpp\\Biff12CommonGrammar.abnf`.
  Neither internal file is included in the public PDF or this repository.

This is sufficient to emit and parse the empty calculated-member collection
needed for a Custom Data reference. The internal `PCDCALCMEMSEXT` production
is still needed only if the implementation elects to parse or author a
non-empty OLAP calculated-member collection. That collection is forbidden for
a connection associated with a PivotCache, and it is separate from the
`BrtBeginExtConn14` payload grammar. A latest-release hash can be retained as
an informational gap, but it does not block the v21.0-family field grammar.

## Compact extracted grammar and field order

The following is a compact extraction, not a reproduction of the specification.
The v20250916 External Data Connections grammar gives these relevant rules:

```abnf
FRTEXTCONNECTIONS = [BrtFRTBegin [EXTCONN14] [EXTCONN15] BrtFRTEnd] *FRT
EXTCONN14 = BrtBeginExtConn14 [PCDCALCMEMSEXT] BrtEndExtConn14
```

Thus `PCDCALCMEMSEXT`, when present, is a sequence **after** the
`BrtBeginExtConn14` record and before `BrtEndExtConn14`; it is not an unnamed
tail inside the `BrtBeginExtConn14` payload.

The §2.4.78 packet and field prose identify these named payload fields in this
order:

```text
BrtBeginExtConn14.payload :=
    FRTHeader          // FRTBlank, 4 bytes
    irstCulture        // XLWideString, variable
    irstClientCubeUrn  // XLWideString, variable
```

`FRTBlank.reserved` is a four-byte value that MUST be zero and MUST be
ignored. The binding reader ignores the reserved value on ingress, including
nonzero values, and a source-bound UID rewrite preserves those four bytes
verbatim. This preservation policy does not authorize generating nonzero
reserved values: any future fresh-record author must emit zero. The current
binding helper edits existing records and does not author a new ExtConn14
collection.

`XLWideString` is a four-byte unsigned character count followed by
that many Unicode characters, with `rgchData` occupying `cchCharacters * 2`
bytes. In the repository's little-endian BIFF12 implementation this is a
`u32` count followed by UTF-16LE `u16` code units. The §2.4.78 prose limits
`irstCulture` to fewer than 85 characters and `irstClientCubeUrn` to fewer
than 65,536 characters.

The `...` rows are diagram continuation, not omitted fields. This follows the
MS-XLSB packet-diagram convention, not an inference about this record alone:

* [XLWideString](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-xlsb/5253755b-33cb-4796-835d-caf07bf70ad4)
  shows `cchCharacters`, `rgchData (variable)`, and `...`; its field
  descriptions define only those two fields.
* [BrtModelTable](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-xlsb/be1d97e8-27ea-4d72-b105-695cae940864)
  shows three consecutive variable `XLWideString` fields, with `...` after
  each, and its descriptions define exactly `irstId`, `irstName`, and
  `irstConnection`.
* [ArgDesc](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-xlsb/067ff23f-9caa-45bf-9576-71ffda47cd84)
  shows the same terminal `...` after one variable field and defines only
  `iArgDesc` and `stArgDesc`.
* [DVals](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-xlsb/52633552-72eb-4927-b7b8-b93686000101)
  places `...` at the start of wrapped rows while its prose defines every
  field (`xLeft`, `yTop`, `unused3`, and `idvMac`).

Accordingly, §2.4.78 has no hidden payload field, flag, or future tail between
the three named components. Future records are represented by the separate
outer FRT grammar (`[BrtFRTBegin ... BrtFRTEnd] *FRT`); preserving or refusing
an unsupported outer FRT/alternate-content wrapper is a compatibility policy,
not an unresolved `BrtBeginExtConn14` field grammar.

`irstCulture`, when non-empty, SHOULD be an RFC 3066 language tag. If that
field is absent, the connection uses the server language. A non-empty
`irstClientCubeUrn` MUST equal the `id` attribute of a `datastoreItem` in the
package's Custom Data Properties part. `BrtBeginExtConn15.irstId` is a
separate `XLNullableWideString` data-model identifier and is not a Custom Data
edge.

## Admission constraints

The §2.4.78 prose imposes all of the following facts:

1. The immediately preceding `BrtBeginExtConnection.idbtype` MUST be
   `DBTOLEDB`.
2. If the ExtConn14 calculated-member collection is non-empty, the immediately
   preceding `BrtBeginECDbProps.icmdtype` MUST be `CMDCUBE` (`CmdType` value
   `0x1`).
3. If the external connection is associated with a `PivotCache`, the ExtConn14
   collection MUST be empty.

The `EXTCONN14` optional member in the outer ABNF and these context rules mean
that an empty collection is represented by the begin record followed by the
end record. A scanner must still validate the preceding connection context;
it must not infer a Custom Data edge from `BrtBeginExtConn15` or from an
unresolved FRT/alternate-content wrapper.

## BIFF12 framing implications

The record enumeration assigns kind **1068** to `BrtBeginExtConn14` and kind
**1069** to `BrtEndExtConn14`. The repository raw writer encodes kind 1068 as
the two-byte 7-bit continuation value `AC 08`; this is implementation evidence
for the existing BIFF12 framing, not a substitute for the MS-XLSB specification.

For every XLWideString field that is present, its wire contribution is
`4 + 2 * utf16_code_units`. With both fields present, the complete payload
length is `4 + (4 + 2 * culture_units) + (4 + 2 * urn_units)`, or
`12 + 2 * (culture_units + urn_units)`. The §2.4.78 prose separately says
that if `irstCulture` is not present, server language is used; that semantic
absence does not add an unlisted field or flag. Replacing a UID therefore
requires recomputing the UTF-16 unit count, string bytes, and enclosing BIFF
payload-length varint. The raw writer permits a one-to-four-byte length
varint; the width changes at the usual 7-bit continuation boundaries. Any
source ranges after the record must be based on the resulting complete record
range.

The implementation references used for this inference are
`crates/litchi-xlsb/src/raw/writer.rs:60-85,160-185` and
`crates/litchi-xlsb/src/raw/record.rs:14-17,403-417`. No implementation code
was changed for this evidence task.

## Local evidence first

The checked-in local MS-XLSB material is the v20251113 text. Relevant excerpts
are:

* `3rdparty/specs/[MS-XLSB]/2 Structures/2.4 Records.md:6238-6341` — §2.4.78
  semantics, packet order, `FRTHeader`, `irstCulture`, and
  `irstClientCubeUrn`; `:6478-6479` identifies the following
  `BrtBeginExtConnection`; `:34758-34760` identifies the matching end record.
* `3rdparty/specs/[MS-XLSB]/2 Structures/2.5 Structures.md:194-207` —
  `ArgDesc`'s complete two-field description despite a terminal `...` row.
* `3rdparty/specs/[MS-XLSB]/2 Structures/2.5 Structures.md:25892-25987` —
  `XLWideString`, its four-byte count, and `cchCharacters * 2` data size; its
  own terminal `...` is the closest direct notation analogue.
* `3rdparty/specs/[MS-XLSB]/2 Structures/2.4 Records.md:41342-41445` —
  `BrtModelTable`'s three consecutive variable strings and complete field
  descriptions.
* `3rdparty/specs/[MS-XLSB]/2 Structures/2.1 File Structure.md:803-819` —
  the External Data Connections part and its internal ABNF path.
* `3rdparty/specs/[MS-XLSB]/Front Matter.md:92-93` — revision 21.0 on
  2025-09-16 and revision 21.1 on 2025-11-13.

Local file SHA-256 values are recorded in [`provenance.json`](provenance.json).

## Official Microsoft sources and revision pins

* [MS-XLSB §2.4.78 BrtBeginExtConn14](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-xlsb/985a2de9-9a1b-42c1-a299-d2057caa5e48)
  — field semantics and packet.
* [MS-XLSB §2.1.7.24 External Data Connections](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-xlsb/a4107f74-4883-4f67-8027-9f5a340ccde0)
  — part ownership and the public/internal ABNF boundary.
* [MS-XLSB §2.5.169 XLWideString](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-xlsb/5253755b-33cb-4796-835d-caf07bf70ad4)
  — length-prefixed Unicode structure.
* [MS-XLSB §2.4.715 BrtModelTable](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-xlsb/be1d97e8-27ea-4d72-b105-695cae940864)
  — three variable fields each followed by diagram continuation and no hidden
  fields in the prose descriptions.
* [MS-XLSB §2.5.3 ArgDesc](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-xlsb/067ff23f-9caa-45bf-9576-71ffda47cd84)
  — one variable field followed by diagram continuation.
* [MS-XLSB §2.5.36 DVals](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-xlsb/52633552-72eb-4927-b7b8-b93686000101)
  — wrapped rows whose leading ellipses do not introduce fields.
* [MS-XLSB §2.1.8 Common Productions](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-xlsb/c620f74b-2b48-4ae1-8d07-6033a9038299)
  — FRT wrapper rule and its internal grammar path.
* [MS-XLSB revision index](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-xlsb/acc8aa92-1f02-4167-99f5-84f9f676b95a)
  — published revision 21.1 dated 2025-11-13.
* [MS-XLSB v20250916 PDF](https://officeprotocoldoc.z19.web.core.windows.net/files/MS-XLSB/%5BMS-XLSB%5D-250916.pdf)
  — revision 21.0 artifact that prints the compact External Data Connections
  ABNF used above.
* [MS-XLSB v20251113 PDF](https://officeprotocoldoc.z19.web.core.windows.net/files/MS-XLSB/%5BMS-XLSB%5D-251113.pdf)
  — revision 21.1 artifact; its §2.4.78 extracted text is byte-for-byte equal
  to v20250916, while §2.1.7.24 delegates ABNF to the internal path.

The exact artifact hashes and HTTP metadata are in [`provenance.json`](provenance.json).

## Validation record

The evidence was obtained by searching the local vendored files first, then
retrieving the two official revision-specific PDFs, checking their SHA-256 and
HTTP metadata, extracting text with `pdftotext -layout`, comparing the §2.4.78
section from the section heading through the next section heading, and
checking the official ArgDesc, XLWideString, BrtModelTable, and DVals packet
diagrams against their complete field descriptions.
The extracted section SHA-256 is
`39969f32ef770ef4fad96cb33616d96598953bc2201f2f46fe42088135d4e4a9` for
both v20250916 and v20251113. The downloaded PDFs and text were temporary
research inputs only and are not part of this repository.
