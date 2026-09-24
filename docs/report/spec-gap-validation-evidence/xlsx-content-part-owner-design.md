# SpreadsheetDrawing `contentPart` owner and implementation design

This note closes the design question behind the remaining
`UnknownKind::ContentPart` inventory in `litchi-xlsx`. It is an evidence and
implementation plan. The direct core read slice described below is now
implemented; mutation and extension-profile work remain open.

The word *content part* is overloaded in the OOXML family. The owner that
matters here is the SpreadsheetDrawing owner in a worksheet drawing part. It
has two related, but differently qualified, element profiles:

* the core SpreadsheetML Drawing element, directly under an anchor; and
* the Office 2010 SpreadsheetDrawing extension element, nested in a group.

The DrawingML `a14:contentPart` and PresentationML `p:contentPart` profiles
are separate owners. They must not be used as the relationship or type rule
for the XLSX owner.

## Identity and placement

The following table uses expanded names. A prefix shown in a source file is
only a lexical spelling; the URI is the identity.

| Profile | Expanded element QName | CT/type evidence | Legal SpreadsheetDrawing placement | Relationship facts established locally |
| --- | --- | --- | --- | --- |
| Core, transitional | `{http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing}contentPart` | The local ECMA Part 1 §A.4.5 schema names the core type `CT_Rel`: required `r:id`, with no extension child sequence. The generated SDK's `xdr14:CT_ContentPart` association is a code-generation projection and is not the core grammar. | One object choice item in each `xdr:twoCellAnchor`, `xdr:oneCellAnchor`, and `xdr:absoluteAnchor`. It follows the anchor geometry (`from`/`to`, or `from`/`ext`, or `pos`/`ext`) and precedes required `xdr:clientData`. It is a direct anchored object, has no core group placement, and has no group path of its own. | The transitional explicit relationship type is `http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml` with `TargetMode="Internal"`. The target is arbitrary supported XML; if no explicit MIME exists for that XML format, the normative fallback is `text/xml`. |
| Core, strict | `{http://purl.oclc.org/ooxml/drawingml/spreadsheetDrawing}contentPart` | Same core `CT_Rel` grammar in the strict `xdr` dialect: required `r:id`, no `xdr14` children, and no core group placement. | The same three anchor particles and source order as the transitional core profile, with strict `xdr` and strict relationship attribute namespaces. | The strict explicit relationship type is `http://purl.oclc.org/ooxml/officeDocument/relationships/customXml` with `TargetMode="Internal"`. The target is arbitrary supported XML; `text/xml` is the normative fallback only when no explicit MIME exists. |
| SpreadsheetDrawing group extension, Office 2010 | `{http://schemas.microsoft.com/office/excel/2010/spreadsheetDrawing}contentPart` (often lexically `xdr14:contentPart`) | `[MS-ODRAWXML]` §2.20.1.1 defines the element as `CT_ContentPart`; §2.20.3.2 defines the sequence. The SDK calls this `xdr14:contentPart` and maps its type to `xdr14:CT_ContentPart`. | A repeated object choice under `{xdr}grpSp`, after `nvGrpSpPr` and `grpSpPr`, so it can be nested in a group. The normative integration is an `mc:AlternateContent` branch. If the declaration is `xmlns:xdr14="http://schemas.microsoft.com/office/excel/2010/spreadsheetDrawing"`, the Choice is `mc:Choice Requires="xdr14"` and contains `xdr14:contentPart`; `mc:Fallback` contains `xdr:sp`. `Requires` is a whitespace-separated prefix token list, not a namespace URI. It inherits the outer anchor and group transforms; it has no independent anchor geometry. | `r:id` is required. The local extension profile does not assign its generic relationship type, TargetMode, or target content type. Do not inherit the core `customXml` rule without extension-specific evidence. |

The required relationship attribute has one of these expanded names according
to the package dialect:

* transitional: `{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id`;
* strict: `{http://purl.oclc.org/ooxml/officeDocument/relationships}id`.

The local ECMA schema is the authority for the core grammar. `CT_Rel` is
qualified by the applicable SpreadsheetDrawing schema namespace. In §A.4.5
the core object choice contains `contentPart type="CT_Rel"`, and the three anchor
complex types place that choice before required `clientData`; there is no core
group placement. The generated SDK is still useful for detecting the distinct
expanded names: it records core `xdr:contentPart` in the anchor particles and
`xdr14:contentPart` in the group particle. Its association of the core element
with `xdr14:CT_ContentPart` is a generated projection, not permission to apply
the extension's `cNvPr`, `bwMode`, `nvPr`, or `xfrm` children to core XML.

For the **extension profile only**, the local extension schema gives the child
order:

```text
xdr14:nvContentPartPr?  xdr14:nvPr?  xdr14:xfrm?  xdr14:extLst?
```

The extension CT has required `r:id` and optional `bwMode`. Its non-visual
child has required `xdr14:cNvPr` and optional `xdr14:cNvContentPartPr`, whose
latter type is `a14:CT_NonVisualInkContentPartProperties`. `nvPr` carries
optional SpreadsheetDrawing `macro` and `fPublished` attributes. The element
names above are the extension-schema qualified names; the transform and
extension-list types are DrawingML types. None of these extension children or
attributes is required or implied by a core `xdr:contentPart`. A strict
implementation must verify the corresponding strict extension qualification
before emitting new strict XML.

The primary local evidence is:

* The local ECMA-376 Part 1 PDF in
  `3rdparty/specs/ECMA-376/ECMA-376-1_5th_edition_december_2016.zip` closes
  the core link: §20.5.2.12 (printed p. 3165) requires the package-dialect
  `customXml` relationship and `TargetMode="Internal"`; §15.2.4 defines a
  Content Part as arbitrary supported XML and says `text/xml` shall be used
  when no explicit MIME exists; §A.4.5 XSD contains `EG_ObjectChoices`,
  `CT_Rel`, and the anchor sequences with `clientData` after the
  object choice. The corresponding §B.4.5 RELAX NG names are
  `xdr_EG_ObjectChoices` and `xdr_CT_Rel`. ECMA-376 Part 4 §13.2.4 supplies the transitional
  relationship spelling.
* `[MS-ODRAWXML]` §2.2.2, `2.20.1.1`, and `2.20.3.2` identify the Excel
  extension URI, the group placement, the MCE Choice/Fallback contract, the
  required `r:id`, and the ordered child sequence. The extension relationship
  type remains an open profile question.
* `[MS-XLSX]` §2.4.70 and §2.6.150–151 repeat the Excel target namespace,
  XML-content purpose, required relationship ID, and CT sequence.
* The checked-in SDK schema has the three anchor particles, the group particle,
  and the `xdr14:CT_ContentPart` attributes/children in
  `3rdparty/Open-XML-SDK/data/schemas/schemas_openxmlformats_org_drawingml_2006_spreadsheetDrawing.json`
  and
  `schemas_microsoft_com_office_excel_2010_spreadsheetDrawing.json`.
* The repository's independent base-content-part review records the same
  package-dialect split and internal TargetMode at
  `docs/report/spec-gap-validation-evidence/docx-signatures-ink/ink-host-spec-review.md#L57-L68`.
  That is corroborating evidence for the shared ECMA base profile, not a
  substitution of the Word host for SpreadsheetDrawing.

## Namespaces that are deliberately out of scope for this owner

These expanded names are useful negative boundaries:

| Other profile | Expanded QName and owner | Why it is not this gap |
| --- | --- | --- |
| DrawingML Office 2010 extension | `{http://schemas.microsoft.com/office/drawing/2010/main}contentPart` (`a14:contentPart`), type `CT_GvmlContentPart` | `[MS-ODRAWXML]` §2.3.1.3/§2.3.3.4 places it in DrawingML `grpSp` and `lockedCanvas`. Its relationship profile is out of scope here: the local extension text spells a `customXml` URI without the `/relationships` segment, so this note does not assert either URI as an established rule for that owner. It supplies no relationship evidence for the SpreadsheetDrawing profiles. |
| PresentationML | `{http://schemas.openxmlformats.org/presentationml/2006/main}contentPart` (`p:contentPart`) and its strict counterpart | It is a slide/group owner. The existing PPTX inert inventory is useful API precedent, but its owner, placement, and relationship validation are not XLSX rules. |
| Other `[MS-ODRAWXML]` extension sections | `a14`, `p14`, `w14`, and related extension namespaces | The §2.13–§2.20 range contains several format/host extensions. Only the Excel URI `http://schemas.microsoft.com/office/excel/2010/spreadsheetDrawing` in §2.20 is the nested SpreadsheetDrawing group profile. |

The distinction matters for graph closure. An `a14:contentPart` with a
`customXml` relationship is not evidence that an `xdr:contentPart` may use the
same relation. Conversely, the presence of `a14:cNvContentPartPr` in the
non-visual type does not make the owner an `a14:contentPart` element.

## Package graph closure

The intended owner graph is:

```text
/xl/worksheets/sheetN.xml
  -- worksheet drawing relationship (transitional or strict) -->
/xl/drawings/drawingM.xml  [drawing content type]
  -- drawing .rels, package-dialect customXml relationship ID from core
     SpreadsheetDrawing contentPart -->
content-part target [arbitrary supported XML; explicit MIME preserved, text/xml
                     fallback when no explicit MIME exists]
  -- target .rels, if present --> bounded outbound targets
```

The `r:id` is resolved in the **drawing part's** relationship part, not in the
worksheet relationship part. The worksheet-to-drawing edge is already checked
by `workbook::drawing`: it requires an internal drawing relationship and the
ordinary drawing content type. For a core `xdr:contentPart`, the next edge is
the exact package-dialect `customXml` relationship with `TargetMode="Internal"`.
For the nested Excel extension owner, the edge remains an unresolved
extension-profile rule and must be retained/diagnosed without borrowing the
core rule.

For the known Ink specialization, `[MS-ODRAWXML]` §2.1.4 gives content type
`application/inkml+xml` and transitional source relationship
`http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml`.
It explicitly lists a Worksheet Drawing core `contentPart` and the
SpreadsheetML group extension `contentPart` as valid owners. That is an Ink
profile, not a generic rule for every XML content part. No strict generic
relationship URI or generic target content type should be invented from it.

The generic ECMA Content Part rule allows arbitrary supported XML and says that
`text/xml` is the fallback when an explicit MIME does not exist. That
normative fallback is distinct from a package's actual `[Content_Types].xml`
mapping: preserve and validate an existing declared mapping, and do not invent
an OPC mapping merely because a producer omitted one.

Closure requirements for a future typed owner are:

1. Resolve the worksheet drawing part and its package relationship before
   exposing a content-part selector.
2. Resolve the owner element's required relationship ID against the drawing
   `.rels`; for a core owner require the exact package-dialect `customXml`
   relation and internal TargetMode, and retain target reference and declared
   content type in the semantic handle. A dangling ID, unsupported mode,
   wrong relation, or contradictory catalog entry is a typed validation
   failure, not an empty payload. The extension owner needs a separate relation
   policy before this validation can be enabled for it.
3. Read or retain the target XML under the caller's byte and graph budgets.
   The ECMA generic Content Part rule prohibits implicit or explicit
   relationships to parts defined by ECMA-376; preserve and diagnose any
   producer-defined relationship material encountered rather than assuming a
   supported outbound graph. A target cannot be deleted merely because the
   owner has become unreachable in one semantic projection.
4. Scan all inbound references to the target before `replace`, `remove`, or transfer.
   Shared targets require an explicit disposition (`retain`, detach/retarget,
   or a profile-approved cascade). A local drawing-only scan is insufficient
   to prove exclusivity.
5. Preserve the inactive MCE branch and all unknown child/extension XML. The
   active `Choice` and `Fallback` are part of the source graph even when only
   one branch is semantically selected.

## Pre-implementation XLSX inventory

This section records the baseline that motivated the design. The direct core
read slice now validates `CT_Rel` attributes in both parsers and pairs source
and typed count/order/anchor geometry in the worksheet reader. The physical
relationship ID remains in the source owner; the typed inventory retains the
structural `UnknownKind::ContentPart` classification.

At that baseline, `UnknownKind::ContentPart` was a structural inventory label,
without a content-part owner:

* [`model.rs`](../../../crates/litchi-xlsx/src/drawing/model.rs#L147-L176)
  has `Shape`, `Group`, `Connection`, `ContentPart`, and `Other`, but
  `Unknown` stores only compatibility/complete anchor geometry, an optional
  `cNvPr@descr`, and the enum. It has no relationship handle, element/profile
  identity, source range, group path, MCE branch, target payload, or lifecycle
  owner.
* [`codec.rs`](../../../crates/litchi-xlsx/src/drawing/codec.rs#L488-L612)
  first requires the core transitional or strict SpreadsheetDrawing namespace,
  then matches the local name `contentPart`. This recognizes a direct core
  local name in an anchor; it does not recognize the Excel extension QName in
  a group. It opens `Context::UnknownObject`.
* [`codec.rs`](../../../crates/litchi-xlsx/src/drawing/codec.rs#L465-L478)
  only recognizes `nvSpPr`, `nvGrpSpPr`, and `nvCxnSpPr` in that context. It
  has no path for the extension-only optional `nvContentPartPr` child. When
  that extension child is present, its `cNvPr` is required. This is not a
  missing core child check: core `contentPart` has the childless `CT_Rel` type.
* The picture and chart contexts are the paths that call the relationship
  attribute resolver. The unknown context ignores `r:id` and unknown children.
  Consequently, a direct core `contentPart` with a required `r:id` can be
  accepted as an unknown object with no recorded relationship; a missing or
  malformed content-part relationship can also pass the same path.
* [`codec.rs`](../../../crates/litchi-xlsx/src/drawing/codec.rs#L716-L821)
  constructs `Unknown` only when the unknown object has no tracked image or
  chart relationship. The `(Unknown, _, _)` branch explicitly rejects tracked
  relationships, but the current unknown path never tracks the required
  `contentPart` ID. This is why the result is a lossy structural inventory,
  not semantic validation.
* [`codec.rs`](../../../crates/litchi-xlsx/src/drawing/codec.rs#L851-L895)
  runs MCE processing before the typed parser with default capabilities. The
  typed model receives the processed branch and has no record of
  `AlternateContent`, `Choice/@Requires`, or `Fallback` provenance. Source
  preservation is a separate concern; parsing the selected branch does not
  establish ownership of the inactive branch.
* [`source.rs`](../../../crates/litchi-xlsx/src/drawing/source.rs#L656-L721)
  indexes direct pictures and a global set of relationship-bearing attributes.
  The global set can see an `r:id` lexically, but it has no content-part owner
  or branch association. It cannot answer which content-part element owns an
  edge or whether the target is shared.
* [`workbook/drawing.rs`](../../../crates/litchi-xlsx/src/workbook/drawing.rs#L380-L455)
  resolves the worksheet drawing part, scans it, parses the typed inventory,
  and proves only the source/typed **picture** pairing. There is no
  content-part pairing or target graph closure.

That baseline did not validate a content part or establish that it was safe
to remove. It preserved the original drawing member only through the
surrounding package's unchanged-member behavior. The new read slice validates
the core owner and target on access; safe mutation still requires the later
graph-closure work.

## Proposed public seams

The first public seam should be a read-only, source-order semantic view. The
names below are design names, not an API commitment:

```text
WorksheetDrawing::content_parts() -> bounded iterator of SpreadsheetContentPartView
WorksheetDrawing::content_part(selector) -> SpreadsheetContentPartView
```

`SpreadsheetContentPartView` should contain:

* a stable source-order selector and profile (`CoreAnchor` or
  `GroupExtension`), with a group-path selector for nested objects;
* complete outer anchor geometry for a direct owner, or the outer anchor plus
  ancestor group path/transforms for a nested owner;
* profile-specific authored metadata and source ranges: a core owner has the
  required relationship ID and no extension non-visual child grammar; an
  extension owner may expose `cNvPr`, `macro`, `fPublished`, and `bwMode`, plus
  ranges for its child sequence, group path, and MCE Choice/Fallback branch;
* a relationship projection whose ordinary operations do not require callers
  to pass a physical `r:id`; diagnostics may expose the ID and part URI;
* an opaque, bounded XML payload reference for a generic target, preserving
  target content type, target mode, target reference, and outbound relation
  metadata; and
* an optional shared typed payload view only after the Ink profile's relation,
  content-type, and root/grammar checks succeed.

The payload owner should be layered. `litchi-xlsx` owns the SpreadsheetDrawing
anchor/group placement and its OPC graph wrapper. A neutral typed Ink payload
grammar may live in `litchi-drawingml` when the common grammar is proved. The
DrawingML crate must not own `PackURI`, ZIP members, or relationship catalogs;
that boundary follows ADR 0002 and ADR 0010/0011. A generic opaque payload
owner in the XLSX host may initially retain bytes and relationship metadata
without pretending to know their application semantics.

Lifecycle operations should follow the accepted transaction rules:

* `replace` can be offered first for an existing owner when the resolved edge
  and target are internally consistent. It must update the target bytes and
  dependency closure as one staged transaction and read the result back.
  Replacing a shared target requires package-wide inbound analysis and an
  explicit supported shared-update or clone-and-retarget disposition; it must
  not silently change another owner's payload.
* `remove` removes the owner and its required edge only after inbound/shared
  target analysis. A direct core anchor requires exactly one object choice,
  so removing only `contentPart` is invalid. Remove its containing anchor or
  replace its object choice atomically under a supported disposition; otherwise
  return typed refusal. It must not leave an orphan relationship or silently
  delete a target used by another owner.
* `clear` cannot generically remove the primary payload while keeping the
  owner, because `r:id` is required by both local CT profiles. It needs a
  profile-defined legal empty target or must return typed refusal; deleting
  the required attribute would create an invalid owner.
* `add`, `duplicate`, and cross-sheet transfer should remain implementation-
  staged until package content-type catalog handling, part-name policy, ID
  allocation, and dependency-closure rules are verified. The core relation
  URI and TargetMode are now normative; the nested extension relation remains
  a typed-refusal path until its own profile is closed. A copy of the PPTX
  `p:contentPart` API is not evidence for those rules.

These choices implement ADR 0001's typed refusal and unknown-preservation
rules, ADR 0003's selector-first snapshot/transaction/patch lifecycle, and
ADR 0006's requirement that extensions and MCE owners have explicit
ownership. They also preserve the existing PPTX inert-content precedent from
ADR 0024 without conflating hosts.

## Source ranges, limits, and stale-source rules

The existing source layer has the right primitives but only emits them for its
current picture inventory:

* [`source.rs`](../../../crates/litchi-xlsx/src/drawing/source.rs#L71-L196)
  provides checked half-open `ByteRange`, `ElementRange`, namespace URI,
  lexical prefix, opening-tag end, closing-tag offset, and source identity.
* [`source.rs`](../../../crates/litchi-xlsx/src/drawing/source.rs#L656-L767)
  keeps an immutable source lease and rejects a range against different bytes.
* [`source.rs`](../../../crates/litchi-xlsx/src/drawing/source.rs#L1964-L2097)
  records transitional/strict relationship attributes under bounded limits,
  but currently does not attach each reference to a content-part owner.

The content-part scanner should add an owner record containing:

1. owner element range and exact expanded QName;
2. enclosing anchor range, or group ancestor ranges and child ordinal path;
3. `AlternateContent`, `Choice`, `Fallback`, `Requires`, and selected-branch
   ranges when present;
4. owner `r:id` attribute range and resolved relationship dialect;
5. child element ranges and namespace contexts, including opaque `extLst`; and
6. target part bytes/relationship-member identity when the package graph is
   opened, without copying more than the caller allows.

Every new allocation must remain under finite caller policy. The current
worksheet drawing path already combines the caller's package limits with
bounded drawing defaults: XML bytes are capped at 32 MiB, nodes/events at
1,000,000, depth at 256, direct pictures at 100,000, relationship references
at 4,096 per configured source bucket, and standalone fragments at 16 MiB.
The content-part profile needs additional explicit ceilings for:

* number of content-part owners and nested group paths;
* target payload bytes per part and in total;
* outbound relationships per target and total traversed parts;
* inbound-owner scan count; and
* MCE branch bytes and graph traversal depth.

These must be minima of host defaults and caller `ReadLimits`, and must fail
before unbounded target reads or semantic allocations. A source-bound edit is
stale if any admitted dependency differs from the snapshot: worksheet XML and
worksheet `.rels` used to select the drawing, drawing XML and `.rels`, the
content-types catalog, target bytes and target relationship members, all
recursively traversed target-closure parts, and every relationship member used
by the package-wide inbound scan. Owner and inactive/selected MCE branch bytes
are part of that read set. Newly introduced members or relationships that change
the admitted closure must also conflict or undergo the same final validation. Return the existing
`SourceChanged`/patch-conflict style error rather than applying a textual
best-effort replacement. The inverse patch must retain all changed owner,
relationship, target, and MCE branch bytes; a no-op should share the original
artifact and must not require a reopen.

## Normative and evidence prerequisites still unresolved

The core profile's normative links are closed by the local ECMA evidence. Its
CT is `CT_Rel`, the object choice precedes required `clientData` in
all three anchor kinds, and there is no core group placement. The package
dialect selects exact `customXml` relationship URI and `TargetMode="Internal"`;
the target is arbitrary supported XML with `text/xml` as the normative
fallback when no explicit MIME exists. The following implementation/profile
links remain open:

1. Determine the SpreadsheetDrawing group-extension relationship type,
   TargetMode, target content type, and strict counterpart. `[MS-ODRAWXML]`
   §2.20 and the local `[MS-XLSX]` extension CT require `r:id` but do not
   authorize borrowing the core `customXml` rule or the Ink content type.
2. Verify how the Office 2010 extension URI and its child qualification are
   represented in a strict package. The local generated SDK is transitional
   evidence and cannot settle this write-time question; it also must not be
   read back into the core grammar.
3. Define MCE capability behavior for an Excel extension Choice. The reader
   must retain inactive branches and expose branch provenance even when the
   typed view selects a fallback. The current pre-parser MCE step does not
   provide that provenance.
4. Decide the package content-type catalog policy for read/write. The ECMA
   `text/xml` fallback is a content-part default, not permission to fabricate
   a missing `[Content_Types].xml` mapping or to reject an explicit supported
   XML MIME type.
5. Obtain native or authoritative fixtures containing each relevant case:
   transitional/strict direct anchors, nested group Choice/Fallback, an Ink
   target, a generic XML target if permitted, a shared target, and outbound
   target relationships. Until then, creation and generic relationship edits
   remain typed-refusal paths.

These are explicit open links, not assumptions hidden in an API. In
particular, the fact that §2.20 calls the target XML and the SDK calls the
element `ContentPart` does not establish an OPC relation URI or a safe delete
policy.

## Bounded corpus result

I searched a bounded checked-in fixture corpus, without copying any native
originals:

* roots: `test-data/libreoffice-core`, `test-data/ooxml`, `test-data/poi`, and
  `test-data/office-interop`;
* archive extensions: `.xlsx`, `.xlsm`, `.xltx`, `.xltm`, `.docx`, `.pptx`,
  `.pptm`, `.dotx`, and `.dotm`;
* members examined: 6,981 `.xml` and `.rels` members across 321 successfully
  read archives (15,210,639 archive bytes); matching was by case-insensitive
  XML local name `contentPart`, with 0 archive errors and 0 matching elements;
  the member count and digest independently match the review scan; and
* reproducibility digest for the sorted archive-path/SHA-256 manifest:
  `a593f74c344a5925a79a578d888a803a4ac3491e99df314277ffce3c1ac15cf5`.

This is a bounded negative result, not a claim that no native Office file
contains a content part. The checked-in provenance identifies some fixtures as
LibreOffice/POI or resaved interoperability material; it is not an
all-Microsoft-native corpus. No native manifest entry was found, so there is
no native example hash to report beyond the corpus manifest digest above.

## Staged implementation and verification slice

Steps 1 and 2 are the implemented read-only direct-owner slice, behind the
existing package read limits. Its public entry points are
`Worksheet::{content_part, content_parts}` and
`WorksheetDrawing::{content_part, content_parts}`, selected with
`drawing::ContentPartSelector`. Payload access is deferred through
`WorksheetContentPart::read_payload`, which validates XML and borrows the
package bytes. Outbound relationship metadata is borrowed and bounded; it is
not recursively traversed. Steps 3 onward remain open:

1. Add a source owner index for direct transitional/strict core
   `xdr:contentPart` in all three anchor kinds. Capture exact ranges, the
   required `r:id`, the core's no-child grammar, source-order selector, and
   anchor geometry. Do not require extension `cNvPr`, `bwMode`, or other
   `xdr14` fields, and do not expose editing yet.
2. Resolve the drawing `.rels` edge and return a bounded relationship/payload
   view. For a core owner validate the exact package-dialect `customXml`
   relation, internal TargetMode, required ID, and dangling references. For
   the group extension preserve an unsupported relation as opaque diagnostic
   data until its extension profile is answered.
3. Add graph-closure checks and source-bound `replace`/`remove` only for an
   already-resolved, profile-approved target. Stage the complete candidate,
   validate the changed closure, read back, and expose an inverse patch plus
   stale-source rejection.
4. Add the nested Excel extension profile under group paths, retaining both
   MCE Choice and Fallback ranges. Verify that a group content part cannot be
   misclassified as a direct anchor object.
5. Once the Ink relation/content-type/root proof exists, expose a shared
   typed Ink payload owner from the common DrawingML layer while keeping XLSX
   responsible for package edges and group/anchor placement.
6. Only then consider creation, duplication, and cross-sheet transfer with
   deterministic relationship-ID/part-name allocation and full dependency
   closure.

The verification matrix for those slices should include:

* each transitional and strict anchor kind, missing/duplicate/malformed
  required `r:id`, wrong TargetMode, wrong relation type, dangling target,
  and target content-type mismatch;
* extension group content in an active MCE Choice, fallback-only processing,
  inactive-branch preservation, nested groups, and branch/source-range edits;
* negative nonclassification tests for `a14:contentPart` and `p:contentPart`;
* shared target and outbound relationship retention/removal decisions;
* source identity/stale patch rejection, exact inverse restoration, no-op
  artifact sharing, and caller budget exhaustion before payload allocation; and
* a fixture manifest update whenever an authoritative native example is added.

This sequence closes the owner and graph semantics without silently dropping
the lifecycle requirements in ADR 0003 or treating the current unknown enum as
proof of a valid relationship profile.
