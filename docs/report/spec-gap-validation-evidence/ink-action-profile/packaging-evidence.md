# PPTX InkAction packaging evidence

Review date: 2026-09-11.

Root independently reran the retained scanner and reproduced the receipt
byte-for-byte (SHA-256
`f5734d000672b00b6cc8e925026a554c9c009d0c20a223838ea561272c453fe9`).
The scan is a raw-marker search, not a namespace-aware XML classification;
encoded namespace spellings and skipped/unreadable inputs are outside its
negative result.

The current Microsoft Learn [Ink Extensions clause](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-pptx/0a7df81a-b34d-4846-8e1f-db55b34f696b)
confirms the MCE content-part/fallback shape described below. The
[Ink Content Part clause](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-odrawxml/096dacae-0d2c-4861-bc4d-c8e4c6405ad3)
ties the InkML packaging contract to an InkML root, while the
[InkAction schema](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-odrawxml/eebdf5ec-b1ff-4953-b2bc-6febc0df7f1e)
defines the action vocabulary. These checked references do not establish the
missing action-part packaging fields; that remains an evidence gap, not proof
that no such contract or producer file exists elsewhere.

Status: **bounded negative evidence**. The vendored specifications identify the
PresentationML owner of an InkAction extension, but the inspected package
corpus contains no successfully scanned native or independently validated
package from which to recover the action part's content type, relationship
type, or target URI. This note records that boundary; it does not invent an
OPC contract from the InkML contract.

## What is established

The local `[MS-PPTX]` Ink Extensions clause (`2 Structures/2.2 Extensions.md`,
§2.2.3.1) defines the host-side shape:

- a slide `spTree` or `grpSp` may contain an `mc:AlternateContent`;
- its `mc:Choice` requires both
  `http://schemas.microsoft.com/office/powerpoint/2010/main` and
  `http://schemas.microsoft.com/office/powerpoint/2014/inkAction`;
- the selected `Choice` contains a PresentationML `contentPart`; and
- the `mc:Fallback` contains a PresentationML `pic`.

This is the normative owner boundary. A future PPTX reader must select the MCE
branch by resolved namespace URI, retain the inactive branch and fallback as
source data, and treat the selected `contentPart`, its slide relationship, and
its target as one dependency closure. Prefix spelling is not the identity of a
`Requires` namespace. A duplicate supported choice, an ambiguous selected
branch, an unresolved `r:id`, or an external/missing target must be refused
before exposing a typed action view.

The local `[MS-ODRAWXML]` action section (`2 Structures/2.21
http---schemas.microsoft.com-office-powerpoint-2014-inkAction.md`) defines the
`iact:actions` vocabulary and target namespace. Its schema appendix
(`5 Appendix A - Full XML Schemas/5.19 ...inkAction Schema.md`) defines the
XML root and content model. Neither checked section enumerates an OPC part
content type, a source relationship URI, or a fixed `/ppt/...` target for a
part whose root is `iact:actions`.

The same `[MS-ODRAWXML]` corpus does enumerate an **Ink Content Part** in
`2 Structures/2.1 Part Enumerations.md`, §2.1.4. That contract is specific and
limited to an `inkml:ink` root:

| Part field | Established value for an InkML part |
| --- | --- |
| Content type | `application/inkml+xml` |
| Source relationship | `http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml` |
| Owner | a document/slide/drawing `contentPart` relationship |
| Root | `ink` in `http://www.w3.org/2003/InkML` |

Those values are **not** evidence for an `iact:actions` part. In particular,
`application/inkml+xml`, `customXml`, and `/ppt/ink/inkN.xml` must not be used
for an action writer until a producer fixture or an additional normative
source binds them to the action root.

## Bounded corpus result

The reproducible scanner and its manifest are retained beside this note:

```text
python3 docs/report/spec-gap-validation-evidence/ink-action-profile/scan_pptx_ink_action.py \
  --repo . \
  --output docs/report/spec-gap-validation-evidence/ink-action-profile/packaging-scan.json
```

It scans the explicit `3rdparty/`, `test-data/`, and `docs/` roots for OPC
Presentation package suffixes (`.pptx`, `.pptm`, `.ppsx`, `.potx`, `.ppsm`,
`.potm`). This includes the vendored Open XML SDK, LibreOffice, and Apache POI
corpora under `3rdparty/`, plus the repository's test and evidence fixtures.
Legacy binary `.ppt` files are outside this ZIP/OPC scan.

The same run inspects ZIP/JAR metadata under those roots without extracting
their members. It found no nested Presentation package entry in 174 readable
archives; two intentionally corrupt LibreOffice test ZIPs are recorded as
metadata errors. Thus the result does not silently treat an unexpanded archive
as an inspected action fixture.

The receipt records a SHA-256, path, size, status, and member counts for every
candidate. The scan used a 64 MiB compressed-file bound, a 64 MiB aggregate
uncompressed-package bound, a 16 MiB member bound, and a 10,000-member ZIP
metadata bound; member searches streamed 64 KiB chunks. The recorded result is:

- 1,021 candidates: 902 under `3rdparty/`, 78 under `test-data/`, and 41 under
  `docs/`;
- 1,003 packages were validly scanned, covering 42,578 ZIP members and reading
  42,117 members;
- 18 candidates were malformed or unreadable ZIP inputs (all paths, errors,
  sizes, and hashes are retained in the receipt), so they are excluded from the
  successful-package negative result;
- 176 nested ZIP/JAR archives were metadata-inspected: 174 readable, two
  corrupt, and zero nested Presentation package entries;
- no package or member hit a size bound and no successfully read member
  contained `iact:`, `inkAction`, the action namespace URI, or
  `2014/inkAction`.

This is a negative result for the successfully scanned corpus, not evidence of
Office acceptance. The receipt includes package hashes, but no hash identifies
an action-bearing native or independently validated fixture. The malformed
inputs cannot establish absence or acceptance and must not be used as action
evidence.

The only local package implementation is the separate InkML path in
`crates/litchi-pptx/src/presentation/embedded/ink/`. It deliberately writes
`/ppt/ink/inkN.xml`, `application/inkml+xml`, and a `customXml` relationship for
an `inkml:ink` payload. Its tests in
`crates/litchi-pptx/tests/pptx_ink_annotations.rs` do not contain
`iact:actions`. The generic
`crates/litchi-pptx/src/presentation/embedded/content_parts/` reader is an
opaque `p:contentPart` graph and likewise does not establish the action part
contract.

## Packaging fields that remain unknown

The following fields cannot be filled from current local evidence:

| Field | Current evidence | Safe host policy |
| --- | --- | --- |
| Action part content type | None | Require and validate an explicit `[Content_Types].xml` override/default from a fixture or declared profile. Do not infer it from the root name. |
| Relationship URI | None for `iact:actions` | Require the explicit slide relationship and validate its URI against the accepted profile. Do not reuse `customXml` by analogy. |
| Target path | None | Resolve the relationship target with normal OPC URI rules and retain the exact target part in the snapshot; do not scan `/ppt/ink/` by filename. |
| Target mode | No action fixture | Require an internal target for the typed host; refuse external or unresolved targets. |
| Action root/part co-location | None | Preserve unknown content parts opaquely until a fixture establishes whether the action root is the relationship target or another owner-owned payload. |

The MCE placement is therefore actionable now, while package typing is an
evidence gate. A typed owner may be implemented only after a package fixture
records `[Content_Types].xml`, the owning slide `.rels`, the target member,
the root namespace, and the selected/inactive MCE bytes together with a full
package hash. Until then, a candidate can be classified as an action only when
all of those explicit facts agree; otherwise it remains opaque or returns an
unsupported-profile diagnostic.

## Smallest evidence-completing follow-up

Obtain one native or independently validated PPTX containing the extension and
retain it without normalization. The evidence test should assert, at minimum:

1. a slide or group `AlternateContent` with both required `Choice` URIs, a
   `contentPart`, and a real `pic` fallback;
2. arbitrary namespace prefixes and the producer's `mc:Ignorable` declarations;
3. the owning slide `.rels`, `[Content_Types].xml`, relationship target mode,
   resolved target member, and action root bytes;
4. an `iact:actions` root plus at least one direct/grouped action, property,
   and opaque trace/transform boundary; and
5. controlled mutations for a wrong content type, wrong relationship URI,
   missing/duplicate `r:id`, inactive-only choice, external target, orphaned
   target, duplicate supported choice, and stale-source publication.

That fixture can then define the first bounded PPTX batch: selected-branch
read, source-backed snapshot, add/replace/remove with relationship and content
type closure, and readback of the action profile. It should reuse the shared
`iact:actions` codec while keeping OPC ownership in `litchi-pptx`. It must not
claim action execution, rendering, InkML validation, or native interoperability
until those properties have separate evidence.

## References

- Local `[MS-PPTX]` §2.2.3.1: `3rdparty/specs/[MS-PPTX]/2 Structures/2.2 Extensions.md`.
- Local `[MS-ODRAWXML]` §2.1.4: `3rdparty/specs/[MS-ODRAWXML]/2 Structures/2.1 Part Enumerations.md`.
- Local `[MS-ODRAWXML]` §2.21: `3rdparty/specs/[MS-ODRAWXML]/2 Structures/2.21 http---schemas.microsoft.com-office-powerpoint-2014-inkAction.md`.
- Local action schema: `3rdparty/specs/[MS-ODRAWXML]/5 Appendix A - Full XML Schemas/5.19 http---schemas.microsoft.com-office-powerpoint-2014-inkAction Schema.md`.
- Reproducible corpus scanner: `scan_pptx_ink_action.py` and its package manifest `packaging-scan.json` in this directory.
- [Microsoft Learn: MS-PPTX Ink Extensions](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-pptx/0a7df81a-b34d-4846-8e1f-db55b34f696b).
- [Microsoft Learn: MS-ODRAWXML Ink Content Part](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-odrawxml/096dacae-0d2c-4861-bc4d-c8e4c6405ad3).
