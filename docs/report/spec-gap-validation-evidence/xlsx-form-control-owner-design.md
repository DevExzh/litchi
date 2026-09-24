# XLSX form-control owner and worksheet API design

Status: design handoff only. This file specifies the owner boundary and the
evidence required before implementation status changes. It does not add a
parser, package edge, writer, test, or native-application support claim.

The scope is an existing worksheet form-control property owner and a bounded
scalar edit whose closure may include both its `formControlPr` payload and
mirrored VML `ClientData`. Reads and edits remain inert:
there is no macro invocation, ActiveX instantiation, formula calculation,
rendering, or external-resource access.

## Decision

`litchi_xlsx::form_control` owns the XML payload of one Office 2010 control
properties part. The worksheet/package owner must prove the owner graph before
returning that payload as a worksheet form control.

The admitted part profile is exact:

| Graph value | Required value |
| --- | --- |
| Part content type | `application/vnd.ms-excel.controlproperties+xml` |
| Worksheet source relationship | `http://schemas.openxmlformats.org/officeDocument/2006/relationships/ctrlProp` |
| Root expanded name | `{http://schemas.microsoft.com/office/spreadsheetml/2009/9/main}formControlPr` |
| Root type | `CT_FormControlPr` in MS-XLSX §2.6.65 |
| Worksheet owner | an effective SpreadsheetML `controls/control` with required `shapeId` and `r:id` |
| Target mode | internal; target resolved from the worksheet relationship member |

The first host profile also pins the two MCE capability URIs required by the
fixtures:

```text
worksheet MCE capability: http://schemas.microsoft.com/office/spreadsheetml/2009/9/main
drawing MCE capability:   http://schemas.microsoft.com/office/drawing/2010/main
```

The profile records the expanded URI set, not the source prefixes (`x14` and
`a14`). A snapshot and every patch retain this capability/profile identity;
changing capabilities or mirror/shape rules invalidates the snapshot rather
than silently selecting a different branch.

The first worksheet host profile also pins the worksheet dialect and
relationship namespace. Its worksheet root is exactly
`{http://schemas.openxmlformats.org/spreadsheetml/2006/main}worksheet`; the
`controls`, `control`, `drawing`, and `legacyDrawing` elements are in that
same canonical SpreadsheetML namespace. Their `r:id` attributes and the
corresponding worksheet relationship types use exactly
`http://schemas.openxmlformats.org/officeDocument/2006/relationships`. The
worksheet relationship member itself has the OPC root QName
`{http://schemas.openxmlformats.org/package/2006/relationships}Relationships`
and `Relationship` children; its `Type` values are checked against the
canonical URI table rather than matched by local name. The
strict counterparts are recognized as a separate dialect for dispatch, but
are not admitted to this form-control profile without a native fixture:

| Worksheet graph value | Canonical first profile | Strict counterpart and policy |
| --- | --- | --- |
| worksheet root QName | `{http://schemas.openxmlformats.org/spreadsheetml/2006/main}worksheet` | `{http://purl.oclc.org/ooxml/spreadsheetml/main}worksheet`; dispatch only, no form-control read/write claim |
| `r:id` attribute namespace | `http://schemas.openxmlformats.org/officeDocument/2006/relationships` | `http://purl.oclc.org/ooxml/officeDocument/relationships`; dispatch only, no mixed namespace |
| worksheet relationship-type namespace | `http://schemas.openxmlformats.org/officeDocument/2006/relationships` | `http://purl.oclc.org/ooxml/officeDocument/relationships`; dispatch only, no mixed namespace |

Consequently, a canonical worksheet with a strict `r:id` or relationship
type, or a strict worksheet with a canonical one, is a typed dialect/graph
refusal rather than a prefix-normalization opportunity. The MCE `Requires`
tokens are capability names resolved to the pinned expanded URIs; they do not
change this worksheet namespace policy.

The relationship target is found from the relationship graph and the content
type from `[Content_Types].xml`. A path such as `xl/ctrlProps/ctrlProp1.xml`
is an observed fixture value, never an owner-discovery rule. The strict
relationship substitute is not admitted by this design because the local
MS-XLSX part enumeration supplies the canonical `ctrlProp` URI and the corpus
contains no strict form-control owner.

The selected worksheet member must have the observed content type
`application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml`,
the ECMA `CT_Worksheet` root type, and the canonical root QName in the policy
table. A worksheet that only matches by filename or by a local name in another
namespace is not admitted.

The package layers remain separate. `litchi-opc` owns physical members,
content types, relationship targets, source versions, signatures, and
publication. `litchi-xlsx` owns SpreadsheetML worksheet semantics and the
effective control collection. The leaf codec supplies the already-reviewed
`SourceView`, `SourceProperties`, `Properties`, `ControlSelector`,
`ScalarField`, `ScalarValue`, `Item`, and exact source splices. The worksheet
owner composes those pieces; it must not make the leaf codec resolve package
relationships.

This boundary follows [`docs/GOAL.md`](../../GOAL.md), the accepted ADR index
in [`docs/adr/README.md`](../../adr/README.md), and the package layering in
ADR 0002/0010/0011/0024. The current gap is explicit: the audit lists
`formControlPr` (§2.4.34) as a zero-match XLSX implementation area
([`docs/report/spec-gap-audit.md`](../spec-gap-audit.md)), while
[`crates/litchi-xlsx/tests/form_control_properties.rs`](../../../crates/litchi-xlsx/tests/form_control_properties.rs)
states that its native checks stop at the XML-part boundary. The design below
does not relabel that leaf coverage as worksheet or package support.

The schema inputs are the vendored MS-XLSX files under
[`3rdparty/specs/[MS-XLSX]`](../../../3rdparty/specs/%5BMS-XLSX%5D), plus
`sml.xsd`, `dml-spreadsheetDrawing.xsd`, and
`vml-spreadsheetDrawing.xsd` in the ECMA-376 Transitional schema archive at
[`3rdparty/specs/ECMA-376`](../../../3rdparty/specs/ECMA-376).

The local MS-OI29500 evidence used for the host profile is explicit:

* §2.1.603 (Part 1 §18.3.1.19, `control`) requires Office shape IDs in
  `1..67,098,623`, requires the drawing shape ObjectType to be Picture, and
  limits control names to 32 characters;
* §2.1.605 (Part 1 §18.3.1.21, `controls`) limits controls to 65,535;
* §2.1.604 (Part 1 §18.3.1.20, `controlPr`) documents attributes Excel
  ignores or does not round-trip for form controls;
* §2.1.1335 (Part 1 §20.1.10.55, `ST_ShapeID`) defines the `_x0000_s<num>`
  shape-ID form and the numeric range; and
* §§3.9.3 (`ST_ObjID`) and 3.9.6 (`ST_SpId`) define Office escaped object
  names and the separate numeric VML shape-ID form; and
* §§2.1.1863, 2.1.1868–1891 (Part 4 VML `ClientData` fields) provide the
  field-level VML behavior and limits used to decide whether a mirror mapping
  is proven.

The local text does not by itself provide a complete lexical equivalence for
every x14 attribute and every VML element, especially `editVal`/`VTEdit` and
the absence/default behavior of some `ClientData` children. The review gate
therefore needs the reviewer to provide an exact local MS-OI29500 section and
line citation for any additional mirror rule that should be admitted. Until
that citation and a fixture prove the mapping, the field remains readable but
not writable.

## Schema ownership and ActiveX boundary

The local MS-XLSX material gives three distinct things different owners:

* `formControlPr` (§2.4.34 and `CT_FormControlPr` §2.6.65) is the root of the
  control-properties part. Its scalar attributes and optional `itemLst` and
  `extLst` are the leaf module's model.
* ECMA-376 Transitional `sml.xsd` defines worksheet `CT_Controls` and
  `CT_Control`. A `control` has required `shapeId`, required `r:id`, optional
  `name`, and optional worksheet `controlPr`. Its `controlPr` is
  `CT_ControlPr`, containing an anchor and placement/presentation attributes;
  it is not the x14 `formControlPr` payload.
* ActiveX persistence is a separate relationship edge using one of the exact
  descriptor URIs in the table below, to an ActiveX descriptor
  (`application/vnd.ms-office.activeX+xml`) and possibly an opaque binary
  (`application/vnd.ms-office.activeX`). Those are the existing inert
  `active_x` owner. The presence of worksheet `controlPr`, a drawing macro
  attribute, or VML `FmlaMacro` does not prove an ActiveX edge.

The ActiveX edge is dispatched by its complete URI, never by the local name
`control` or by a target filename:

| Edge | Canonical URI | Strict URI | Dispatch/refusal |
| --- | --- | --- | --- |
| ActiveX descriptor | `http://schemas.openxmlformats.org/officeDocument/2006/relationships/control` | `http://purl.oclc.org/ooxml/officeDocument/relationships/control` | Excluded from form-control selector ordinals; inspection belongs to the inert `active_x` owner, which validates the exact descriptor content type (`application/vnd.ms-office.activeX+xml`) and its closure |
| ActiveX binary companion | `http://schemas.microsoft.com/office/2006/relationships/activeXControlBinary` | no strict URI is admitted by the local evidence | Preserve/dispatch only under the existing ActiveX owner after its descriptor edge is proven; the form-control owner never treats it as `ctrlProp` |

The form-control `ctrlProp` edge remains the canonical URI in the Decision
table. No strict `ctrlProp` URI is guessed: a strict or unknown edge is a
typed unsupported-relationship refusal, while the exact ActiveX URIs above
are inspected by the separate `active_x` API when their descriptor/content-type
closure is valid. The form-control owner does not invoke that loader.
The selected form-control closure must use its admitted dialect;
strict or mixed-dialect form-control ownership is refused. An unrelated
ActiveX control with either exact descriptor URI does not invalidate a
canonical form-control owner on the same worksheet. This permits the
form-control projection to coexist with unrelated ActiveX edges
without treating an ActiveX descriptor as form-control properties or claiming
that form-control inspection validates the separate ActiveX payload.

MS-XLSX §2.1.1 requires a control-properties part to be explicitly related
from a SpreadsheetML control and says that a control with this relationship
must not also have an embedded control-persistence relationship. The host
scanner therefore classifies an effective control per relationship edge:

* canonical `ctrlProp` plus the admitted target is the form-control owner;
* canonical ActiveX `control` plus an ActiveX descriptor is the ActiveX owner;
* a worksheet `controlPr` child may be retained with either owner as inert
  placement metadata; it is not a conversion between owner families; and
* both persistence edges on one control are a typed owner-conflict refusal.

The form-control API never reads the ActiveX descriptor or binary as a
fallback for a missing `ctrlProp` part. It never executes a macro, follows an
external relationship, or decodes VML macro/formula values.

The optional worksheet `controlPr` is retained in the host source range and
is not parsed by the form-control property owner. If it is absent or fails
the `CT_ControlPr` anchor grammar, that is an inert worksheet diagnostic; the
owner must not call the ActiveX loader or discard the valid `ctrlProp` edge.
Scalar publication still requires the independent DrawingML/VML identity
closure below, and any future anchor edit must handle `controlPr` as part of
the full worksheet graph.

The following host relationship/content-type values are observed in the
pinned fixtures and are admitted for the first host profile:

| Worksheet child | Worksheet relationship type | Observed target content type |
| --- | --- | --- |
| `drawing r:id` | `http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing` | `application/vnd.openxmlformats-officedocument.drawing+xml` |
| `legacyDrawing r:id` | `http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing` | `application/vnd.openxmlformats-officedocument.vmlDrawing` (the fixtures use the `.vml` default) |

These drawing and VML values are closure edges, not alternate owners for the
properties payload. Strict drawing/VML relationship substitutions and other
content types require their own fixture evidence; this design does not infer
them from a namespace or filename.

## Native evidence and admitted shape linkage

The bounded corpus is documented in
[`xlsx-form-control-properties/README.md`](xlsx-form-control-properties/README.md)
and pinned by
[`xlsx-form-control-properties/native-corpus.json`](xlsx-form-control-properties/native-corpus.json).
It contains seven LibreOffice test inputs and ten properties parts. These are
real source fixtures, not Office acceptance or rendering evidence.

The package and properties-part hashes are reproduced here so an owner test
can bind itself to the same inputs without guessing from file names:

| Fixture package (bytes) | Package SHA-256 | Properties member: SHA-256 |
| --- | --- | --- |
| `tdf134769.xlsx` (16,892) | `50d4f7d5d0ffe17e2251b21cfb66bd5018e3b1791c7f5c27592a47a9d9e004b1` | `ctrlProp1.xml`: `ee451e92b54ea18058c6346e97fba72b06b7ca68050b6cd334de1efa1d2bd15f` |
| `tdf161365.xlsx` (12,161) | `5c5ec45c1dd5db0a911e74d0e7e6cde5bf713b5ef38fcf84483903ddad583957` | `ctrlProp1.xml`: `be9ec4e18bbc82af16dd7c4867228f37d07bc4eeacbe26c951c5fa69ae373774`; `ctrlProp2.xml`: `9734573ff8601e72482b4f71e93988e9a90c0042121015b56daa42d71612acce` |
| `button-form-control.xlsx` (11,190) | `624321843560b792805592d604e7a645040f19f0ae1f661fec024b9aeec9613f` | `ctrlProp1.xml`: `7c1c0df33df49e6bfaac423ee1d265c62366d59279b13656d21ff468d3232265` |
| `tdf120301_xmlSpaceParsing.xlsx` (11,538) | `73709ecaa9270324eaf3f0c4fae8ec5e733ab639daa7b407f08d42538bac60bf` | `ctrlProp1.xml`: `3ee809532b360df6270a88d5f5f546d380fdce24874e0ff80f97f351226bfde2`; `ctrlProp2.xml`: `27a21638d1c95252a61eb44940f403927f751efbaf286c3676409cc5bde2ca90` |
| `tdf60673.xlsx` (13,249) | `d9abab27b6cb7ae4eda4729473d8d48175f63fa66457e64a8134f9fa59689bc8` | both `ctrlProp1.xml` and `ctrlProp2.xml`: `7c1c0df33df49e6bfaac423ee1d265c62366d59279b13656d21ff468d3232265` |
| `singlecontrol.xlsx` (17,915) | `a13fd1a411a499f4b474bbe553c7950959ab8895131c40687522d14379885cde` | `ctrlProp1.xml`: `87536c4c081860e76ae7636a6e6c6b9b4a3c47282545de174c0106a2babfff85` |
| `checkbox-form-control.xlsx` (11,394) | `8b602b21777bc06ca574f61783594caa5c37f68bdf66c66e58ef598758367038` | `ctrlProp1.xml`: `3ee809532b360df6270a88d5f5f546d380fdce24874e0ff80f97f351226bfde2` |

Every recorded part has the exact content type and root above and one
incoming internal worksheet edge. The observed properties include CheckBox,
Button, and Radio values. `tdf134769.xlsx` contains `fmlaLink="#REF!"`, so
an unrelated scalar edit must preserve that lexical error token. Button
fixtures contain worksheet/VML macro metadata; those values remain inert.

The retained bytes used for the host closure have these member hashes. They
bind the `x14` worksheet ancestry, worksheet relationship edge, `a14`
DrawingML ancestry, and VML mirror observations to the tracked fixtures:

| Fixture | `sheet1.xml` | `sheet1.xml.rels` | `drawing1.xml` | `vmlDrawing1.vml` |
| --- | --- | --- | --- | --- |
| `button-form-control.xlsx` | `dd87ba869ed3e4c8773999520bea4b7566cb291c20cabe72ce27fdf0a6666d95` | `baee0b0705da4d01ad0c621711843652182acd7993aa190608f563776d27ea3f` | `1cdb3305a3add5d0d6e2aeb2a606072adf633ae6dfca652e5f3c2afc28ca9b48` | `0b8636ecf3d6f4c8319383b867aa898b62127196648cac0a5d3a9804ec7a8851` |
| `checkbox-form-control.xlsx` | `c6622daa82ff7de20b73668688fd26b10e073c8380ac12bb5928b74fa1530bcf` | `baee0b0705da4d01ad0c621711843652182acd7993aa190608f563776d27ea3f` | `85f37dc2cd451faf60cd06dfa9f8106846000582a21a137b369893fc3620bd5c` | `9c5f1f992d043039b2423ca4a0758c4fc54d101f246d459b37b6f0c9355cf335` |
| `singlecontrol.xlsx` | `1cfdaf5eea3b7f53f31afb86b572b6ce39453c1cdfffae4192e0e61f61e7398b` | `935fda24f6b7bdafb60d65dac2930d3f51e870bf3c1255a2c92f30cb12f4f4e1` | `8de2bba6b9666cc3a4c3a57eb0d66b8cd9d0800573e379cab8400542dba7e594` | `52ffc160137a8a66c33d2f77838e74f8c329a1584123ca4ff1af9e509a000e9d` |
| `tdf120301_xmlSpaceParsing.xlsx` | `1990ee4b6dfac51abc6d949283ad8e729103fbd86d6a6f62560086a14e085d72` | `8cb65789881543a9132fdec83994b37ed6aec4ecface6acd082fd9e2fb1fa095` | `ba9afa5b91ebd28634b4843939dd238f7bae050827b9dc92af485459133c41cc` | `9c295114582238d3ef807e3c281b2f9b8b59da9ca0e14991610d51a7dc86a218` |
| `tdf134769.xlsx` | `80802877681bee7bd91a1ad010c77fd7021936dc12b1f8312af57bf8b9fc4960` | `ab141e6e4007de7d0c2e11550afee969ac931b8e10054dfbbbe7d37c79c3fd4e` | `677e621380f5d764a11a5314504311bf9c612eb398dc2bdd494e8964980b96f5` | `6018595eb5576f45b27d2dff0a7aee5d351b4c2b930a7e46512fd7e38c54cb1d` |
| `tdf161365.xlsx` | `338a6c6b8d874e380293f0c572aa8f777b03e1f2a960fe0990c6c4cff5c551a0` | `4c3559448931247901ed41157aa39e9fe094a582e8808b275159b53f0ae6f756` | `905882956d8afad49e3891e29cdaf9f3f629f8bbf11ee7e64ee8bd4c45986d6e` | `f8f2a8f52c961907b21a5fa6f0a2187ec77785bd7c98012ad1590ce84c5406f1` |
| `tdf60673.xlsx` | `e0c12e2b769db9ab2b84453bb33a68f1443544665687e89f2931b97840af45c4` | `4c3559448931247901ed41157aa39e9fe094a582e8808b275159b53f0ae6f756` | `9874ec396d162f849f19dc51b2438b3b089b2f77d023400dd6469b7775bd706d` | `180a6b3a395052ce81ea944e6e7e38b1e5b2e849190fa3a1bdf2deca590eaaf1` |

The observed host path in all seven inputs is concrete:

```text
worksheet xl/worksheets/sheet1.xml
  ├─ drawing r:id ─────── worksheet rel ──> xl/drawings/drawing1.xml
  ├─ legacyDrawing r:id ─ worksheet rel ──> xl/drawings/vmlDrawing1.vml
  └─ effective controls/control r:id ────> xl/ctrlProps/ctrlPropN.xml
```

The `controls` collection is inside an outer `mc:AlternateContent` and each
fixture control is inside a nested `mc:AlternateContent`, with
`mc:Choice Requires="x14"`. For example, the source control has the shape
and relationship identity in one element:

```xml
<control xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
         xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
         xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
         shapeId="1026" r:id="rId4" name="Check Box 2">
  <controlPr defaultSize="0" autoFill="0" autoLine="0" autoPict="0">
    <anchor>
      <xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>2</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
      <xdr:to><xdr:col>2</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>3</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
    </anchor>
  </controlPr>
</control>
```

The example includes the required `CT_ControlPr/anchor` because the ECMA
schema makes it mandatory when the optional worksheet `controlPr` is present.
This is the schema-oriented spelling: `CT_ObjectAnchor` references
`xdr:from` and `xdr:to`, so the prefixes are shown explicitly here. A source
may use different legal prefixes, which remain source lexical bytes. The
retained LibreOffice producer profile has a different expanded-name spelling
in every recorded control: its nonempty `controlPr` contains an `anchor`
whose `from` and `to` elements are in the SpreadsheetML default namespace,
with `xdr:col`, `xdr:colOff`, `xdr:row`, and `xdr:rowOff` children. Its
representative source form is:

```xml
<controlPr xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
           xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing">
  <anchor>
    <from>
      <xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff>
      <xdr:row>2</xdr:row><xdr:rowOff>0</xdr:rowOff>
    </from>
    <to>
      <xdr:col>2</xdr:col><xdr:colOff>0</xdr:colOff>
      <xdr:row>3</xdr:row><xdr:rowOff>0</xdr:rowOff>
    </to>
  </anchor>
</controlPr>
```

The owner admits this as a source-preserving `LoSmlAnchorV1` producer
profile (the name is a design identifier), recognizes `{SpreadsheetML}from`
and `{SpreadsheetML}to` in that profile, and preserves its exact lexical
bytes. It does not rewrite those elements to the schema-oriented `xdr` names
or reject an otherwise bounded native source solely because strict
Transitional XSD validation expects `xdr:from`/`xdr:to`. Anchor/geometry edits
remain outside the scalar-edit closure. All seven retained packages have
nonempty anchored `controlPr`; none supports a self-closing claim.

The selected worksheet relationship resolves `rId4` to the canonical
`ctrlProp` edge. `drawing1.xml` contains a DrawingML MCE wrapper of its own.
The following is an identity excerpt: the surrounding `twoCellAnchor`,
required `clientData`, and unrelated shape children are intentionally
abbreviated and remain source bytes:

```xml
<mc:AlternateContent
    xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">
  <mc:Choice xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main"
             xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
             xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
             Requires="a14">
    <xdr:sp>
      <xdr:nvSpPr>
        <xdr:cNvPr id="1026" name="Check Box 2" hidden="1">
          <a:extLst>
            <a:ext uri="{63B3BB69-23CF-44E3-9099-C40C66FF867C}">
              <a14:compatExt spid="_x0000_s1026"/>
            </a:ext>
          </a:extLst>
        </xdr:cNvPr>
      </xdr:nvSpPr>
      <!-- shape properties and clientData remain source bytes -->
    </xdr:sp>
  </mc:Choice>
  <mc:Fallback/>
</mc:AlternateContent>
```

Every sampled DrawingML control shape uses `Choice Requires="a14"`; the
worksheet controls use `Choice Requires="x14"`. The resolver therefore needs
both capabilities and records separate selected ancestry for the worksheet
and drawing projections. It retains every inactive drawing `Fallback` and
ignored `Choice` byte range even when the selected `a14:compatExt/@spid`
matches the control. A direct `xdr:cNvPr` scanner misses this identity.

The local MS-ODRAWXML evidence gives this branch semantic ownership: its
`compatExt` description (§2.3.1.2) says the legacy object must be a form
control or legacy OLE/ActiveX object, and `CT_CompatExt` (§2.3.3.2) requires
an `_x0000_s<num>`-shaped `spid` in the documented range. Its SpreadsheetML
drawing example (§3.6.2) places that `a14:compatExt` below the shape's
`cNvPr`. The selected `a14` ancestry and the inactive bytes are therefore
part of the identity closure, rather than optional drawing decoration.

The concrete shape profile is bounded and explicit:

| Identity component | Admitted source profile |
| --- | --- |
| `control/@shapeId` | decimal unsigned integer in `1..=67,098,623`; no numeric coercion from names |
| worksheet count | at most `65,535` effective `control` elements |
| DrawingML family | observed `xdr:sp` with `xdr:cNvPr`; preserve the family and do not convert it to `xdr:pic` |
| DrawingML ID | `xdr:cNvPr/@id` parses to the exact `shapeId`; disagreement is a graph error |
| DrawingML name | when both are present, `xdr:cNvPr/@name` equals `control/@name` byte-decoded; mismatch refuses; duplicate names remain selector ambiguity |
| DrawingML compatibility ID | selected `a14:compatExt/@spid` normalizes to `_x0000_s<shapeId>`, with the MS-ODRAWXML `1025..=268,435,456` `spid` range, and must agree with the control and VML IDs |
| VML identity | `v:shape/@id` and `o:spid` are each numeric only when they use the exact `_x0000_s<num>` form; if both normalize numerically and disagree, refuse. A nonnumeric VML `id` (the corpus has `Check_x0020_Box_x0020_4`) is an Office-escaped name token, compared to the `cNvPr`/control name without numeric coercion; a canonical numeric ID must still come from `id` or `o:spid` and agree with `shapeId` |
| VML `ClientData/@ObjectType` | fixture values `Checkbox`, `Button`, and `Radio`; `CheckBox` in `formControlPr` maps to `Checkbox`, while other known tokens require an exact profile mapping |

The nonnumeric VML `id` path uses an exact, source-preserving object-name
decoder. After XML entity decoding, it first recognizes the complete numeric
`_x0000_s<num>` form; that path is never passed through the name decoder. For
an object name, the decoder scans left to right as specified by MS-OI29500
§3.9.3: `_xHHHH_` decodes one UTF-16 code unit, while the protected sequence
`_x005F_xHHHH_` decodes to the literal text `_xHHHH_` and does not decode the
second token. A malformed or ambiguous escape-like sequence, an unpaired
surrogate, or a code point that cannot be represented in the XML-decoded name
returns `UnprovenVmlNameEncoding`; it is never guessed as a numeric ID. Other
characters are retained exactly. The resulting code-point string is compared
to the XML-decoded `cNvPr/@name` and worksheet `control/@name` without case,
whitespace, or Unicode normalization. The raw VML spelling is retained and
never re-encoded during an unrelated scalar edit. Thus the corpus value
`Check_x0020_Box_x0020_4` compares as `Check Box 4`; an ambiguous name makes
the graph readable only as opaque source and refuses typed mutation.

MS-OI29500 §2.1.603 supplies the Office `shapeId` range, the 32-character
name limit, and its Office requirement that the drawing shape ObjectType be
Picture. The LibreOffice preservation corpus instead demonstrates `xdr:sp`
shapes with VML `ClientData/@ObjectType` values `Checkbox`, `Button`, and
`Radio`; that observation is a source profile, not an Office writer
conformance claim. A future Office-conformant writer must either prove the
Picture shape profile or refuse; it must not silently rewrite the fixture's
`xdr:sp` family during a scalar property edit.
For this linked DrawingML profile, the usable shape-ID intersection is
therefore `1025..=67,098,623`: the worksheet control bound supplies the upper
limit and the `a14:compatExt` `spid` rule supplies the lower limit.

The shape closure is therefore required for an admitted source-backed
control:

1. select exactly one `a14` DrawingML branch and find exactly one relevant
   `xdr:cNvPr/@id == control/@shapeId`;
2. require the selected `a14:compatExt/@spid` and one canonical VML
   `_x0000_s<num>` identity from `v:shape/@id` or `o:spid` to normalize to
   the same number; if both VML attributes are canonical, they must agree,
   while a noncanonical `v:shape/@id` must pass the name comparison above;
3. follow the worksheet `legacyDrawing` edge and find exactly one VML shape
   with that identity and the expected `ClientData/@ObjectType` mapping; and
4. retain the complete anchors, `clientData`, VML `ClientData`, DrawingML
   macro/extension attributes, VML macro/formula elements, and unrelated
   shapes as opaque source ranges.

If either target is absent, has the wrong content type, has duplicate shape
identities, has a canonical `id`/`o:spid` disagreement, has no canonical VML
identity, has a noncanonical VML name that disagrees, has a missing or
mismatched `a14:compatExt`, or uses an identity representation outside this
profile, the package can expose an opaque inventory diagnostic but a typed
scalar edit must refuse. The design does not infer a VML shape from a name,
cell anchor, member suffix, relationship ordinal, or inactive MCE branch. A
future producer profile may widen this rule only with an exact fixture and
schema evidence.

The ECMA schemas support the structural closure but do not establish the
semantic correspondence by themselves: Transitional `sml.xsd` declares
worksheet `drawing` and `legacyDrawing` references, `CT_Controls`/`CT_Control`
and `CT_ControlPr`; `dml-spreadsheetDrawing.xsd` gives anchors a required
`clientData` and shapes their non-visual `cNvPr` ID; and
`vml-spreadsheetDrawing.xsd` defines legacy `ClientData` fields including
`FmlaMacro`, `FmlaLink`, `Anchor`, and `ObjectType`. The ID, MCE, and object
type correspondence above comes from the pinned fixtures plus the cited
MS-OI29500 profile.

## Source read set and graph validation

The owner starts from a selected workbook sheet, never from an unreferenced
`ctrlProp` filename. A bounded read performs these steps:

1. Open the OPC catalog and content-type index under the caller's
   `ReadLimits`, then resolve the workbook sheet selector. The catalog,
   workbook, workbook relationships, and selected worksheet relationship are
   source inputs; `litchi-opc` keeps physical IDs private to this layer.
2. Read the selected worksheet and its relationship member. Validate that the
   worksheet root is exactly the canonical SpreadsheetML QName and that
   `control`, `controls`, `drawing`, `legacyDrawing`, and their `r:id`
   attributes use the pinned canonical relationship namespace policy. A
   strict worksheet dialect is dispatched/refused as described above; it is
   never mixed into this form-control profile. Then scan the effective
   `controls` collection with the common MCE resolver. The resolver enables
   the `x14` requirement token by its pinned URI, applies the first supported
   `Choice` recursively, records every selected wrapper in the worksheet
   ancestry, and retains exact bytes for ignored choices and fallbacks. It
   does not rewrite the selected payload's source namespaces. A direct-child
   scan is insufficient for the corpus.
3. Enumerate effective `control` elements in source order. A
   `ControlSelector::Position(n)` is the checked zero-based position among
   effective controls admitted to the form-control owner; it is not a shape
   ID, relationship ID, or properties-part ordinal. `Name(value)` is an exact
   name lookup in this same collection. Duplicate names are retained and make
   a name lookup ambiguous. `tdf161365.xlsx` supplies that case with two
   controls named `Check Box 4`.
4. Resolve the selected control's `r:id` in the selected worksheet
   relationship member. Require an internal target, the exact `ctrlProp`
   relationship type, a content-type index entry with the exact properties
   content type, and the exact x14 root expanded name. Missing, external,
   duplicate, wrong-type, or wrong-root edges are graph errors; the loader
   never repairs them by deriving a target URI.
5. Build bounded incoming and identity indexes. A second effective control
   pointing to the same properties part is a shared-owner ambiguity for
   scalar mutation. A second control or drawing/VML shape claiming the same
   identity is also ambiguous. A properties part that is unreferenced is
   preserved as an opaque package member and is not projected as a control.
6. Validate the drawing and VML shape closure described above. Resolve
   DrawingML MCE with the separate `a14` capability, record every selected
   wrapper in the drawing ancestry plus selected and ignored branch ranges,
   and include every DrawingML and VML sidecar
   relationship member if the graph has one. The seven fixtures have no
   drawing relationship member because their observed shapes have no nested
   relationship; that absence is not assumed for another source. A present
   sidecar edge and target are read and fingerprinted as a unit. A present
   sidecar whose relationship grammar or target bytes cannot be bounded and
   retained losslessly is an edit refusal, never a silently dropped member;
   an unrecognized sidecar relationship is not silently ignored.
7. Inspect the OPC relationship member derived from the selected properties
   part's URI. The member's presence or absence, raw bytes, every parsed
   outgoing edge, target mode, target, and relationship type are part of the
   graph fingerprint. A present `ctrlProp` relationship member is retained
   byte-for-byte and included in the read set even when it has no semantic
   edge used by this owner; if it is unreadable, over budget, malformed, or
   cannot be retained losslessly, return a typed
   `UnsupportedControlPropertiesRelationships` refusal. The seven retained
   packages contain no outgoing properties-part `.rels` member, so the corpus
   proves neither a writer nor a semantic interpretation for one. Its absence
   is nevertheless fingerprinted and cannot be manufactured or silently
   ignored.
8. Build the content-type-part inventory before selector projection. Any
   `ctrlProp`-typed part with no incoming effective worksheet `ctrlProp` edge
   receives the deterministic `UnreferencedControlProperties` graph
   diagnostic (ordered by the package owner's canonical internal part key).
   This is an invalid-graph diagnostic; the part remains an opaque source
   member and is excluded from `Position`/`Name`. It is not deleted, guessed
   into an owner, or treated as a typed control.
9. Record a graph fingerprint for every read member, every incoming edge, and
   every present or absent selected-properties `.rels` member. The ordinary
   view exposes semantic name, position, object type, properties,
   and diagnostics. It does not expose package part names or raw `rId` values
   as ordinary selectors.

The minimum source read set for a scalar edit is the catalog/content-type
index, workbook sheet resolution, selected worksheet and worksheet
relationships, all outer and nested worksheet `x14` MCE ancestry/branch
ranges, selected properties part, worksheet drawing and legacy drawing
relationships, all drawing `a14` MCE ancestry/branch ranges, matching
DrawingML and VML payloads, all identity references used for the closure,
every present DrawingML/VML sidecar relationship member, the selected
properties-part `.rels` member (including a fingerprint of its absence), and
the package signature graph. A mirrored scalar write has a two-part write set: one exact
`ctrlProp` overlay and one exact VML `ClientData` overlay. A field with no
proven VML mapping is not writable in this profile. Worksheet, worksheet
relationships, DrawingML, MCE branches, all relationship sidecars including
the properties-part member, and unrelated members are copied from the source
byte-for-byte. They remain in the read set so a
concurrent graph change cannot turn the selector or mirror into a different
control.

## Selector-first worksheet API

The host facade should add small wrappers around the existing leaf names. The
following is the recommended surface; exact error aliases follow the existing
`litchi_xlsx::Result` conventions.

```rust,ignore
impl Worksheet {
    pub fn form_controls(&self) -> Result<FormControlCollection<'_>>;

    pub fn form_control<'a>(
        &self,
        selector: impl Into<crate::form_control::ControlSelector<'a>>,
    ) -> Result<Option<crate::form_control::FormControl<'_>>>;
}

impl SourceWorksheet {
    pub fn form_controls(&self) -> Result<FormControlCollection<'_>>;

    pub fn form_control<'a>(
        &self,
        selector: impl Into<crate::form_control::ControlSelector<'a>>,
    ) -> Result<Option<crate::form_control::FormControl<'_>>>;
}

impl WorksheetEdit<'_> {
    pub fn set_form_control_scalar<'a>(
        &mut self,
        selector: impl Into<crate::form_control::ControlSelector<'a>>,
        field: crate::form_control::ScalarField,
        value: Option<crate::form_control::ScalarValue>,
    ) -> Result<&mut Self>;

    pub fn edit_form_control<'a>(
        &mut self,
        selector: impl Into<crate::form_control::ControlSelector<'a>>,
    ) -> Result<FormControlEdit<'_>>;
}

impl FormControlEdit<'_> {
    pub fn set_scalar(
        &mut self,
        field: crate::form_control::ScalarField,
        value: Option<crate::form_control::ScalarValue>,
    ) -> Result<&mut Self>;
}
```

`FormControlCollection` is an immutable snapshot projection. It may expose an
iterator and `get(selector)` but does not expose physical IDs. The leaf
`form_control::FormControl` currently contains the selector, optional name,
and typed `Properties`; the host may introduce a differently named wrapper if
that avoids future ambiguity, but it must keep the same selector semantics.
`Properties` retains authored presence separately from effective defaults,
known enum values separately from bounded unknown lexical values, and exact
opaque attributes, namespaces, item lists, extension lists, and source
diagnostics.

The first writable host batch changes an existing target's scalar attribute
and its proven VML mirror in one transaction. It uses
`SourceView::replace_scalar(field, value)` (or the equivalent checked source
splice) for the properties member, a companion VML source splice for
`ClientData`, and never serializes a detached `Properties` over an existing
part. A leaf-only properties replacement is a lower-level operation and is
not a valid worksheet edit for a mirrored field. `replace_items`,
`insert_item`, and `remove_item` are available in the leaf API, but worksheet
list edits remain deferred until their ordered VML `ListItem` mirror is
proven. `FormControlDraft` must not gain an `add` behavior just because it can
describe properties: adding a visible control is a graph operation.

An ordinary edit must not add, remove, rename, resize, reposition, or migrate
a control. Such operations would need to update the effective worksheet
control, matching DrawingML anchor, matching VML shape, selected and ignored
MCE branches, relationships, and content-type mappings as one closure.

## Form-control and VML mirror closure

The `formControlPr` part and the legacy VML `ClientData` are two serialized
views of overlapping form-control state. A scalar that has a VML counterpart
cannot be published by changing only the properties part. The source-backed
editor must either produce both exact source splices atomically or return a
typed `UnsupportedMirror` refusal. It must not describe the stale VML value as
inert metadata after changing the x14 value.

The field map for the first profile is:

| `formControlPr` field | VML `ClientData` mirror | Lexical/profile rule |
| --- | --- | --- |
| `objectType` | `@ObjectType` | `CheckBox` maps to the observed VML token `Checkbox`; other values require an exact known token map; disagreement refuses |
| `checked` | `Checked` | x14 tokens `Unchecked`/`Checked`/`Mixed` map to the observed VML decimal tokens `0`/`1`/`2`; other VML integers are unknown and read-only |
| `colored` | `Colored` | VML `ST_TrueFalseBlank` semantic boolean; preserve unchanged lexical form |
| `dropLines` | `DropLines` | checked decimal unsigned integer, with the x14 `0..30000` bound |
| `dropStyle` | `DropStyle` | token spelling is retained; setter requires a known VML token |
| `dx` | `Dx` | checked decimal unsigned integer |
| `firstButton` | `FirstButton` | VML empty element means true; absence/explicit false handling is profile-checked |
| `fmlaGroup` | `FmlaGroup` | same inert formula lexical value; no calculation |
| `fmlaLink` | `FmlaLink` | same inert formula lexical value, including an existing `#REF!` |
| `fmlaRange` | `FmlaRange` | same inert range lexical value; list precedence is preserved |
| `fmlaTxbx` | `FmlaTxbx` | same inert cell/range lexical value |
| `horiz` | `Horiz` | VML boolean mirror |
| `inc` | `Inc` | checked decimal unsigned integer |
| `justLastX` | `JustLastX` | VML boolean mirror |
| `lockText` | `LockText` | VML boolean mirror; absence is not guessed to mean false |
| `max` | `Max` | checked decimal unsigned integer |
| `min` | `Min` | checked decimal unsigned integer |
| `multiSel` | `MultiSel` | exact comma-delimited lexical string; no sorting or normalization |
| `noThreeD` | `NoThreeD` | VML boolean mirror; empty element is a true occurrence |
| `noThreeD2` | `NoThreeD2` | VML boolean mirror |
| `page` | `Page` | checked decimal unsigned integer |
| `sel` | `Sel` | checked decimal selection index |
| `seltype` | `SelType` | token spelling is retained; setter requires a known VML token |
| `textHAlign` | `TextHAlign` | alignment token mapping is profile-checked; no default is inferred from absence |
| `textVAlign` | `TextVAlign` | alignment token mapping is profile-checked; no default is inferred from absence |
| `val` | `Val` | checked decimal unsigned integer |
| `widthMin` | `WidthMin` | checked decimal unsigned integer |
| `editVal` | `VTEdit` | type/value mapping is not established by the local MS-OI29500 text; read-only until the reviewer supplies exact evidence |
| `multiLine` | `MultiLine` | VML boolean mirror |
| `verticalBar` | `VScroll` | semantic name differs; lexical boolean mapping requires the profile |
| `passwordEdit` | `SecretEdit` | semantic name differs; lexical boolean mapping requires the profile |
| `itemLst/item/@val` | ordered `ListItem` elements | pair by source order and exact decoded string; entity/whitespace spelling remains source lexical |

The mirror profile does not invent VML absence defaults. The x14 side uses
the schema's authored presence and declared defaults, while the VML side
records whether the element is absent, present with a recognized lexical
value, present with an unknown value, or present but semantically disagreeing
with the x14 attribute. For example, the corpus has `lockText="1"` in a
properties part while some corresponding VML `ClientData` has no `LockText`
element; that is an `UnprovenMirror` diagnostic, not permission to treat
either side as the winner. The corpus also demonstrates proven pairs for
`objectType`, `checked`, `firstButton`, `fmlaLink`, and `noThreeD` (including
the `#REF!` lexical value and empty VML boolean elements).

Every writable row has a profile entry for its source element/attribute,
semantic decoder, lexical encoder, insertion/removal representation, and
absence rule. A boolean setter therefore cannot merely copy the x14 spelling:
it must emit the admitted VML token or empty-element form and the x14 schema
lexical form as one pair. If the profile has only a reader, or cannot decide
whether false means an absent VML child, that row is read-only even when the
detached leaf codec can edit the x14 attribute.

`objectType` is a type transition in the graph even though it is one scalar
attribute. The first profile may write it only when the exact x14 token, VML
`ObjectType`, shape family, and type-specific mirror set are all admitted;
changing a checkbox into a radio or ActiveX object is a graph migration and
refuses in this owner.

The atomic scalar algorithm is:

1. Read the selected x14 attribute and its corresponding VML occurrence,
   including source ranges, namespace context, MCE ancestry, and lexical
   tokens. A missing or unknown mirror, a disagreement, or an unresolved
   `editVal`/`VTEdit` mapping makes that field read-only.
2. Convert the requested `ScalarValue` through the pinned mirror mapping. The
   field-specific output is checked before allocation: x14 uses its schema
   lexical type, VML uses the mapped `ClientData` element/attribute lexical
   form, and formulas are copied as inert strings. A setter cannot introduce
   `#REF!`; an untouched source `#REF!` is copied byte-for-byte into both
   mirrors.
3. Prepare two source-local overlays, one for the properties part and one for
   the VML part. If an element must be inserted or removed, the operation uses
   a validated source anchor and the profile's canonical representation. No
   one-sided properties overlay is published.
4. Reopen the candidate graph, select the same worksheet and drawing MCE
   branches, compare the x14 and VML semantic mirror values, and verify that
   the DrawingML/VML identity closure and all inactive branches are unchanged.
   Only then publish both overlays atomically.

List-item replacement/insertion/removal follows the same two-part rule. It
must update `itemLst/item/@val` and the ordered VML `ListItem` sequence in one
candidate, preserve each item's exact string lexical bytes when untouched,
and refuse when counts/order/opaque children cannot be paired. Until that
mapping has its own fixtures, the leaf item APIs remain available only below
the host graph owner and host list edits remain unsupported.

## Preservation, MCE, ambiguity, and refusal rules

The source view keeps the complete control-properties bytes and the complete
worksheet host bytes. Scalar splice and no-op behavior must preserve:

* omitted-vs-explicit defaults, attribute order and lexical boolean/integer
  forms;
* arbitrary legal prefixes, namespace declarations, comments, processing
  instructions, whitespace, entity spelling, unknown attributes, unknown enum
  tokens, and both `extLst` subtrees;
* unknown or ignored MCE branch bytes, including the selected branch's full
  outer and nested worksheet `x14` and drawing `a14` ancestry;
* worksheet `controlPr`, DrawingML anchors/shapes and selected/ignored
  `a14:compatExt` branches, VML shapes and `ClientData`, macro-looking values,
  formulas, and unrelated package members; and
* every x14/VML mirror pair that is not the requested field. A requested
  mirrored scalar changes both sides together, so preserving the old VML
  value while changing x14 is never an accepted result; and
* pre-existing invalid-but-preserved lexical values such as `#REF!` when the
  edit does not touch them.

For each `AlternateContent`, the MCE resolver scans `Choice` children in
source order under the pinned caller capability set (`x14` for the worksheet
and `a14` for DrawingML) and selects the first supported `Choice` recursively.
If no `Choice` is supported and exactly one well-formed `Fallback` exists, it
selects that `Fallback` for inventory and records the fallback ancestry. A
fallback-selected control is editable only when its complete owner, shape,
and mirror profile is admitted; the editor never upgrades it to a `Choice`.
Unsupported `Choice` branches remain opaque in either case. No `Fallback`, a
malformed or duplicate `Fallback`, an unsupported wrapper with direct content
beside it, two effective owners, or an edit that would change branch selection
returns a deterministic typed MCE refusal. Ignored branches remain opaque and
are never deleted as “duplicates.” A selected control with an unsupported or
unproven shape or mirror closure may be inventoried as opaque, but cannot be
scalar-edited.

The following conditions refuse a typed read or edit at the earliest layer
that can prove the problem:

* wrong content type, wrong root namespace/local name, malformed XML,
  external target, missing target, duplicate relationship IDs, or a guessed
  relationship/path;
* absent, duplicate, or mismatched DrawingML/VML identity; missing/mismatched
  `a14:compatExt`; shared properties target; duplicate `Name` selector;
  duplicate effective MCE owner; unsupported `Choice` without exactly one
  well-formed `Fallback`; unresolved x14/VML mirror; unreadable or
  unpreservable control-properties `.rels`; or conflicting ActiveX persistence
  edge;
* unsupported schema/profile details that would require rewriting opaque
  markup, including an unproven strict form-control dialect;
* stale source version or any read-set member changed after snapshot;
* signed source mutation. An unchanged source may be copied exactly; a changed
  publication must surface the existing OPC
  `SignedSourceRequiresExplicitPolicy`/signed error and cannot silently strip
  or regenerate signatures; and
* caller cancellation, decompression/relationship/XML limits, output budget,
  or allocation failure.

All refusal paths leave the source and transaction unchanged. A typed view may
carry diagnostics for preserved unknown content, but diagnostics never grant
permission to mutate an ambiguous graph.

## Source-backed transaction and exact inverse

The source-backed implementation should mirror the existing page-break and
hyperlink `SourceBackedEditor` family while allowing a multi-part read set:

```text
SourceBackedFormControlEditor
  ├─ snapshot(sheet selector) -> FormControlSnapshot
  ├─ edit(sheet selector) -> FormControlSourceEdit
  └─ publish_commit_to_stream(writer, commit)

FormControlSourceEdit
  ├─ before: immutable graph/source snapshot
  ├─ staged scalar operations keyed by ControlSelector
  └─ commit() -> FormControlCommit

FormControlPatch
  ├─ exact before/after properties and mirrored VML overlays
  ├─ pinned MCE capability/profile identity
  ├─ source version/fingerprint for every read-set member
  ├─ semantic before/after values and selector diagnostics
  └─ inverse() / apply(source)
```

`FormControlSnapshot` stores an immutable `OwnerProfile` value containing the
profile version, expanded worksheet/drawing MCE capability URIs, admitted
relationship/content-type values, shape-ID normalization rules, mirror-map
version, sidecar inclusion/fingerprint policy (including the selected
properties-part `.rels` member and its absence), the profile ceiling of 65,535
effective controls, and the caller's effective limits/cancellation lineage.
A patch stores the same profile value and a deterministic profile
fingerprint.
`Patch::apply` and `inverse` require exact profile equality; a caller that
changes the `a14` capability, selects a different MCE branch policy, changes
the mirror token map, or widens limits must capture a new snapshot first.
This prevents a patch created under one graph projection from being applied
under another while keeping physical IDs out of ordinary selectors.

The patch's ordinary surface reports semantic selector and before/after
values. Its physical source records remain package-owner data. `apply` checks
the worksheet bytes, selected x14 and a14 MCE ancestry, pinned capability and
profile identity, worksheet relationship member, properties part, mirrored VML
part, content-type entry, DrawingML/VML identity members, every present
sidecar relationship member, selected-properties `.rels` presence/bytes and
edges, incoming edges, and signature state before
cloning a candidate. It reopens the candidate through the same owner scanner
and verifies that the selector still resolves to the same graph identity,
that both mirror values changed together, and that only the requested scalar
semantic state changed. Then it atomically replaces all overlays. A changed
signed source is rejected before output.

`inverse()` swaps the exact source overlays and source fingerprints. Applying
it restores the original properties-part and mirrored VML bytes, all original
lexical forms, inactive branches, and the pinned capability/profile identity;
it does not regenerate semantically equivalent XML. If the source or profile
was changed by another actor, either direction returns a typed patch conflict.
A scalar set to its existing authored and mirrored values produces an empty
patch, shares the original source bytes, and publishes an exact package
no-op.

For an ordinary in-memory `Workbook` transaction, the same readback and
atomicity rules apply. The edit may retain compact source ranges or shared
immutable bytes, but it must not expose package IDs, raw relationship locks,
or archive handles through the public selector API.

## Budgets and allocation plan

The host carries the workbook's caller-supplied OPC `ReadLimits`, source-cache
policy, and explicit execution/cancellation context into every read, edit,
candidate reopen, patch application, and inverse. The leaf policy is lowered
from that aggregate policy and already caps one properties part at 16 MiB,
generated output at 32 MiB, XML depth at 256, XML events at 400,000, items at
65,536, one item value at 1 MiB, and retained opaque bytes at 16 MiB.

The host-specific policy needs finite fields for the additional closure:

| Resource charged before allocation | Required caller cap |
| --- | --- |
| effective controls per selected worksheet | `max_controls` |
| selected/ignored MCE branches and depth | `max_mce_branches`, `max_mce_depth`, `max_mce_bytes` |
| relationship edges and incoming references | `max_relationship_edges` |
| DrawingML and VML bytes retained for identity | `max_drawing_bytes`, `max_vml_bytes` |
| DrawingML/VML sidecar relationship bytes | `max_sidecar_relationship_bytes` |
| selected control-properties `.rels` bytes | `max_control_properties_relationship_bytes` |
| shape identities indexed | `max_shapes` |
| mirrored fields and list-item pairs | `max_mirror_nodes` |
| scalar operations in one edit | `max_operations` |
| changed-part and generated-output bytes | `max_changed_parts`, `max_output_bytes` |

These names are design names, not existing production constants. They must be
implemented as caller-owned lowerable limits, never as ambient global state or
an excuse to raise an OPC hard ceiling. The scanner counts XML events,
relationships, shapes, MCE branches, and source ranges before `try_reserve` or
retaining an owned string. Scalar output length is computed with checked
arithmetic before allocating a replacement buffer. The final package and
temporary buffers are both charged; a source accepted by a part limit cannot
silently exceed the caller's aggregate output budget during publication.
`max_controls` is additionally clamped to the profile's `65,535` ceiling; a
source that exceeds that ceiling is a deterministic graph refusal before
selector projection or allocation.

## Feature boundary and staged delivery

The audit's current `formControlPr` row remains a gap until the focused gate
passes. The batches below keep each claim reviewable:

| Batch | Deliverable | Explicit boundary |
| --- | --- | --- |
| 0 | Fixture/schema admission inventory; exact x14 worksheet and a14 drawing MCE ancestry, shape-ID, selected-properties `.rels` presence/absence, sidecar, and mirror checks for all ten parts | No public owner or writer claim; no fixture is treated as Office acceptance |
| 1 | Source-backed read of effective existing controls, typed leaf properties, mirror diagnostics, and shape closure | Read-only; unsupported/ambiguous graphs and unreferenced properties parts are opaque or refused; no package mutation |
| 2 | Mirror-profile implementation for proven fields, including exact x14/VML lexical mappings and candidate readback | No scalar write is enabled for an unresolved or disagreeing mirror; list items remain read-only until paired |
| 3 | Source-backed scalar edits on existing admitted controls; atomic properties-part plus VML overlay, full closure readback, stale/profile checks, signed refusal, exact no-op and inverse | Worksheet, relationship, DrawingML, MCE, sidecars, and unrelated bytes remain unchanged; no add/remove/rename |
| 4 | `Worksheet`/`SourceWorksheet` views and `WorksheetEdit` selector-first scalar verbs | Public physical IDs remain hidden; only existing controls with proven mirror and shape closure |
| 5 | Host list-item edits using paired x14/VML source splices | Requires separate list semantics/readback tests; still no graph CRUD |
| 6 | Whole-control add/remove/rename/reposition only after native fixture output and complete closure proof | Must update worksheet, selected/ignored x14/a14 MCE branches, DrawingML, VML, sidecars, relationships, and content types atomically; otherwise typed `UnsupportedGraph` |
| 7 | Separate ActiveX inert graph work, if requested | Never merges ActiveX descriptors/binaries or macro execution into the form-control owner |

The supported slice after Batch 4 is therefore: ordinary existing
form-control properties, typed known and unknown values, exact unknown/opaque
preservation, and inert scalar changes whose x14/VML mirrors are both updated
atomically. It does not include control creation/deletion, visual geometry,
DrawingML/VML authoring, unresolved mirror fields, ActiveX
persistence/binary editing, macro or formula execution, linked-cell updates,
recalculation, rendering, external links, native Office acceptance, or
signed-source mutation.

Review evidence request: before admitting a mirror row beyond the fixture
provenance above, the reviewer should cite the exact local MS-OI29500 section
and line that establishes its x14/VML semantic and lexical equivalence,
including absence and list-item behavior. The current local text explicitly
leaves `editVal`/`VTEdit`, unresolved absence rules, and unsupported token
maps read-only; no implementation batch should widen those rows by inference.

## Validation gate

Before the audit row can move from gap to supported, tests must reopen every
part in `native-corpus.json`, prove the worksheet relationship/content-type/root
chain, select worksheet `x14` and drawing `a14` branches, prove each
DrawingML/VML identity including `cNvPr`, `a14:compatExt`, `id`/`o:spid`, and
`ClientData/@ObjectType`, fingerprint every present sidecar relationship
member and the selected-properties `.rels` presence/absence, preserve unknown and inactive bytes, and exercise at least one scalar
edit plus the `#REF!` preservation case. They must assert duplicate-name
ambiguity, shared-target refusal, unreferenced-properties diagnostics,
wrong/external edge refusal, missing/duplicate shape refusal,
`id`/`o:spid` disagreement, unsupported MCE refusal, unresolved mirror
refusal, atomic x14/VML readback, stale profile/read-set conflict, signed
mutation refusal, caller-limit refusal before allocation, exact no-op bytes,
and exact inverse bytes. They must exercise the deterministic MCE rule for an
unsupported `Choice` with one valid `Fallback` (read the fallback and retain
the inactive choice) and refusal for no or malformed fallback. Candidate
packages must pass owner/profile validation and OPC validation. Apply local
MS-XLSX/ECMA XML validation to schema-conforming branches; for the retained
LibreOffice `LoSmlAnchorV1` producer form, validate its expanded names,
required anchor children, and exact byte preservation, and do not reject or
rewrite its SpreadsheetML-default `from`/`to` solely because strict
Transitional XSD expects `xdr:from`/`xdr:to`.

Those checks establish source-backed inert support only. A generated package
has no Microsoft Office or LibreOffice acceptance claim until an independently
recorded application open/edit/save/reopen fixture gate supplies that evidence.
