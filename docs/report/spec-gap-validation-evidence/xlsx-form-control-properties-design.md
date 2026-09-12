# XLSX `formControlPr` control-properties owner design

Status: design and validation plan only. This document does not add a parser,
writer, package edge, or public API to the production tree. It is the handoff
for a future `litchi-xlsx` implementation and its focused tests.

The owner described here is the Office 2010 SpreadsheetML **form-control
properties part**. It is deliberately separate from ActiveX. The existing
`litchi_xlsx::active_x` module owns the worksheet `x:controlPr` child and the
inert ActiveX descriptor, binary, and preview graph. The new owner must never
interpret an `x:controlPr` child as `x14:formControlPr`, and ActiveX support
must not be used as a generic fallback for a `formControlPr` part.

## Decision in one page

The new focused module should be named `litchi_xlsx::form_control` (an
internal `control_properties` implementation module is fine). Its root model
is `form_control::Properties`, not `active_x::ControlProperties`, because the
two values have different XML roots, namespaces, package ownership, and
relationship closure.

The admitted owner profile is exact:

| Property | Admitted value |
|---|---|
| Part content type | `application/vnd.ms-excel.controlproperties+xml` |
| Worksheet relationship type | `http://schemas.openxmlformats.org/officeDocument/2006/relationships/ctrlProp` |
| Root expanded name | `{http://schemas.microsoft.com/office/spreadsheetml/2009/9/main}formControlPr` |
| Complex type | `CT_FormControlPr` |
| Child sequence | `itemLst?`, then `extLst?`, each at most once |
| Part owner | an explicit internal `ctrlProp` edge from a SpreadsheetML `x:control` |
| External targets | refused |
| Strict substitute | none claimed; the Microsoft 2009/9 namespace and relationship URI are the profile |

The implementation should initially support the following complete slice:

* read a direct or effectively selected worksheet `x:control`, resolve its
  `ctrlProp` relationship, parse every `CT_FormControlPr` scalar, and expose
  ordered `itemLst/item` values;
* preserve the raw root attributes, comments, whitespace, unknown foreign
  attributes, and both opaque `extLst` subtrees;
* edit individual scalar attributes and list items by source-range splice;
* publish an atomic source-checked reversible patch after reopening the whole
  changed dependency closure; and
* reserve whole form-control add/remove for a later closure-complete phase.
  The eventual operation must include the matching SpreadsheetDrawing anchor,
  VML shape, selected and ignored MCE references, and every package graph edge;
  until that proof exists, both verbs return typed `UnsupportedGraph` rather
  than deleting or creating only a fragment.

Removing just the `formControlPr` part from an existing `x:control` is not a
valid ordinary operation: `x:control/@r:id` is required and the relationship
is the control's owner edge. The eventual `remove` means removal of the
selected whole effective control object, including its matching DrawingML
anchor, VML shape, selected/ignored MCE references, properties edge, and
required graph closure. A future explicit migration operation could replace a
form-control owner with an ActiveX owner, but this design does not silently
perform that conversion. Until the complete identity/closure proof exists,
`remove` refuses and leaves the source untouched.

## Why this is a gap, and what already exists

The current audit calls `formControlPr` a zero-match XLSX implementation gap.
In
`docs/report/spec-gap-audit.md`, the XLSX “spec areas with no code” list names
“`formControlPr` extension (§2.4.34)” and says that only a generic ActiveX/form
control metadata row exists. The implementation-status report likewise lists
`formControlPr` as a missing per-extension typed-preservation owner. The
repository has no `formControlPr`, `ctrlProp`, `itemLst`, or
`application/vnd.ms-excel.controlproperties+xml` token in the existing XLSX
source.

The existing `active_x` owner is real but different:

* `active_x::model::ControlProperties` represents the SpreadsheetML
  `x:controlPr` child. Its fields include `anchor`, `locked`, `macro_name`,
  `alternate_text`, and a preview relationship ID. It is not the
  `x14:formControlPr` part model.
* `active_x::codec::xml::Controls` parses direct worksheet `<controls>` and
  `<x:control>` elements, enforces unique shape IDs and names, and allows at
  most one `x:controlPr` child. That parser is the useful base for locating
  direct controls, but its relationship loader requires the control edge to
  have the ActiveX `.../relationships/control` role and an ActiveX descriptor
  content type.
* `active_x::package` owns the `control` relationship, ActiveX descriptor,
  `activeXControlBinary` children, and optional image preview. It intentionally
  keeps binaries opaque and does not activate or resolve a CLSID.

The form-control owner must not inherit the ActiveX parser's duplicate-name
rejection: the native corpus contains repeated names (for example, two
`Check Box 4` controls in `tdf161365.xlsx`). Position selection remains
unambiguous; `Name("Check Box 4")` returns a typed ambiguity.

There is no existing `formControlPr` implementation or duplicate code owner to
reconcile. The native corpus does contain real owners (listed below), so the
future module must reject a source where the same effective control is
simultaneously presented as a form-control owner and an ActiveX persistence
owner, rather than allowing whichever loader runs first to win.

The architecture below follows the accepted ADRs: the public layer is
selector-first and free of package IDs; snapshots are immutable and cheap to
share; edits are isolated and source-bound; unknown content is preserved;
unsupported or ambiguous edits are typed refusals; and the OPC package layer
owns relationship and content-type graph changes. Active content, macros,
formulas, external targets, and rendering remain inert.

## Normative evidence

The vendored Microsoft material is the primary source for this owner:

* `[MS-XLSX]/2 Structures/2.1 Part Enumerations.md`, §2.1.1, states the
  control-properties content type, `ctrlProp` source relationship, explicit
  relationship from a SpreadsheetML control, and the `formControlPr` root.
  It also states that when this relationship is present the control MUST NOT
  have an embedded control persistence relationship.
* `[MS-XLSX]/2 Structures/2.4 Global Elements.md`, §2.4.34, assigns
  `formControlPr` to the exact
  `http://schemas.microsoft.com/office/spreadsheetml/2009/9/main` namespace
  and type `CT_FormControlPr`.
* `[MS-XLSX]/2 Structures/2.6 Complex Types.md`, §§2.6.63–2.6.65, defines
  `CT_ListItem`, `CT_ListItems`, and the complete `CT_FormControlPr` sequence
  and attributes.
* `[MS-XLSX]/2 Structures/2.7 Simple Types.md`, §§2.7.14–2.7.18 and
  §§2.7.22–2.7.23, defines the control-type, checked-state, drop-style,
  selection, edit-validation, and text-alignment lexical enumerations.
* The full local schema in
  `[MS-XLSX]/5 Appendix A - Full XML Schema/5.4 ...` is the schema-validation
  gate after the codec is implemented.
* The vendored Open XML SDK corroboration is
  `3rdparty/Open-XML-SDK/data/parts/ControlPropertiesPart.json`: it names the
  same content type, relationship type, root element, `ctrlProps` target
  family, and Office 2010 version. Its typed schema entry maps
  `x14:CT_FormControlPr/x14:formControlPr` to
  `FormControlProperties/ControlPropertiesPart`.

The prose in the Microsoft part enumeration has a typographical space in
`application /vnd.ms-excel.controlproperties+xml`. The canonical OPC content
type has no space, as shown by the SDK metadata and MIME syntax; the codec
must compare the canonical value exactly.

## Bounded native-fixture result

The native corpus search inspected member names in 643 LibreOffice and POI
XLSX archives, then decompressed only the candidate XML and relationship
members. It found seven LibreOffice packages containing ten actual
`formControlPr` parts. The pinned package/part hashes, content types, root
expanded names, and incoming edges are recorded in
[`xlsx-form-control-properties/native-corpus.json`](xlsx-form-control-properties/native-corpus.json),
with the scope and limitations documented in
[`xlsx-form-control-properties/README.md`](xlsx-form-control-properties/README.md).

| Package | Form parts | Observed typed evidence |
|---|---:|---|
| `tdf134769.xlsx` | 1 | CheckBox, `fmlaLink="#REF!"`, `lockText=1`, `noThreeD=1` |
| `tdf161365.xlsx` | 2 | CheckBox with absent/`Checked` checked state |
| `button-form-control.xlsx` | 1 | Button, `lockText=1` |
| `tdf120301_xmlSpaceParsing.xlsx` | 2 | CheckBox and Radio, `firstButton=1` |
| `tdf60673.xlsx` | 2 | two Button properties parts with separate incoming edges |
| `singlecontrol.xlsx` | 1 | CheckBox, `Checked`, nested MCE-selected owner |
| `checkbox-form-control.xlsx` | 1 | CheckBox, `lockText=1`, `noThreeD=1` |

The fixtures show a required existing-owner path that a direct-child-only
scanner would miss. In the native `singlecontrol.xlsx` (and similarly in the
other LibreOffice controls), the worksheet contains an outer
`mc:AlternateContent/mc:Choice Requires="x14"` around `controls`, and an inner
`mc:AlternateContent/mc:Choice Requires="x14"` around the `control`. The
control named `Check Box 2` has `shapeId=1026`, `r:id="rId4"`, and its
relationship is the canonical `officeDocument/2006/relationships/ctrlProp`.
The scanner must run the common MCE branch selector with `x14` capability,
record the selected ancestry, and preserve both ignored and selected branch
bytes. It must not classify the selected native owner as unsupported merely
because it is nested below MCE. A mutation that cannot prove this branch
selection remains a typed refusal.

The native fixtures are read/round-trip evidence inputs, not proof that a new
writer is accepted by Excel or LibreOffice. The implementation gate must
open each fixture, read every owner, perform scalar and item edits where the
source grammar allows them, save, reopen, and run local XSD/package checks.
Native Office acceptance and resave evidence remain separate gates.

## Part and relationship ownership

### Authoritative graph

The owner scanner starts from one worksheet and performs these checks before
exposing a typed `Properties` view:

1. Select a direct or effectively selected `x:control` in a SpreadsheetML
   `<controls>` collection. The scanner applies the common MCE capability
   resolver with the x14 namespace enabled. An outer or nested
   `mc:AlternateContent/mc:Choice Requires="x14"` is therefore part of the
   admitted ancestry, not an opaque wrapper. The public selector is a checked
   zero-based ordinal among controls admitted to this effective form-control
   owner in source order or an exact name in that same collection. `shapeId`,
   relationship IDs, part names, XML offsets, and
   archive member names are diagnostics only. MCE selection follows the
   standard first-supported-choice rule: scan `mc:Choice` children in source
   order, select the first whose `Requires` capabilities are all supported,
   and use `mc:Fallback` only when none is supported. Apply this recursively
   for nested `AlternateContent`; multiple supported choices are not an
   ambiguity because later choices are unselected.
   `Requires` prefixes are resolved through the in-scope MCE declarations;
   `mc:Ignorable` does not by itself make an x14 branch effective. Unknown
   requirements and ignored `Fallback` branches remain opaque.
2. Resolve the control's `r:id` in that worksheet's relationship collection.
   The edge must be internal, have the exact `ctrlProp` relationship type, and
   target a part with the exact control-properties content type.
3. Parse that target as one XML part with the exact x14 root expanded name.
   Prefix spelling, default-vs-prefixed namespace declaration, comments,
   entity spelling, and line endings are lexical source, not the profile
   discriminator.
4. The target part's `formControlPr` grammar has no modeled outgoing
   relationships. An outgoing relationship is therefore a host-profile
   closure refusal, not a claim that the CT schema itself defines it as
   invalid: retain the part and its edges as opaque diagnostic content and do
   not perform a typed edit until the profile has a proven closure policy.
5. Build a bounded incoming-edge index. The selected control edge is the
   owner edge. If another effective control resolves to the same properties part,
   expose a shared-owner diagnostic but refuse scalar or item edits because
   one semantic edit would change multiple selected controls. A whole-control
   removal eventually removes only its proven closure; it retains a shared
   properties target and its content-type mapping until the last inbound edge
   is gone. Until DrawingML/VML identity and MCE references are included in
   that closure, removal is `UnsupportedGraph`.
6. Assign the public position from the effective form-control ordinal in
   source order. Preserve duplicate native control names (the corpus contains
   them); an exact `Name` selector is ambiguous when more than one effective
   control has that name. `shapeId` and relationship IDs are diagnostics and
   are never selector tie-breakers. A shared properties target is a separate
   graph-alias diagnostic and refuses scalar/item mutation.

The worksheet edge is the authoritative reference. An unreferenced part with
the right content type is retained and reported as opaque; it is never
invented as a control owner from a file-name convention. A relationship with
the right type but the wrong content type, wrong target mode, missing target,
or wrong root is a typed graph error for the selected control.

### Coexistence with ActiveX and MCE

The following rules keep the two owner families disjoint:

* An effective control whose `r:id` has the ActiveX
  `.../relationships/control` role belongs to the existing ActiveX owner.
  `form_control` returns an owner-conflict/refusal for that persistence graph
  and does not inspect an `x:controlPr` child as an x14 part.
* An effective control whose `r:id` has the exact `ctrlProp` role belongs to the
  form-control owner. The Microsoft rule forbidding an embedded control
  persistence relationship is enforced before a typed edit is staged. Any
  ActiveX persistence edge, `activeXControlBinary` edge, or contradictory
  descriptor relation on that control produces a typed coexistence error.
* Separate controls in the same worksheet may independently be ActiveX and
  form controls. The prohibition is per control and per owner edge, not a
  ban on a workbook containing both feature families.
* The SpreadsheetML `x:controlPr` child is ordinary control metadata and may
  coexist with a form-properties part. Its presence is not an ActiveX
  persistence edge and is not a conflict by itself; preserve it and include
  its worksheet source range in the read set. The conflict test concerns the
  actual selected control's persistence relationship/descriptor graph.
* MCE branch selection is semantic. A control or `ctrlProp` edge in a selected
  `mc:Choice Requires="x14"` branch is an admitted effective owner when the
  first-supported-choice resolver selects that branch. A control in an
  ignored `Choice` or `Fallback` branch is opaque and is not included in the
  selector ordinal. If selected content produces duplicate effective owners,
  or if a mutation would need to edit an unselected branch, the operation
  refuses with an MCE ambiguity error. The implementation must not create a
  second direct `ctrlProp` edge beside an ignored branch.

## Public semantic API handoff

The names below are the proposed stable handoff for the implementation and
test author. They intentionally avoid package URIs, relationship IDs, shape
IDs, archive handles, and raw XML in ordinary signatures.

### Module and selectors

```rust,ignore
pub mod form_control;

pub use form_control::{
    Checked, ControlSelector, DropStyle, EditValidation, FormControl,
    FormControlDraft, FormControlError, Item, ObjectType, Properties,
    SelectionType, TextHAlign, TextVAlign,
};

#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub enum ControlSelector<'a> {
    Position(usize),
    Name(&'a str),
}

impl<'a> ControlSelector<'a> {
    pub const fn position(index: usize) -> Self { Self::Position(index) }
    pub const fn name(value: &'a str) -> Self { Self::Name(value) }
}
```

`Position(n)` is zero-based among effective controls admitted to this
form-control owner: the capability resolver must select the MCE branch (when
one is present), and the selected control must have the exact internal
`ctrlProp` edge, canonical content type, and typed x14 root. ActiveX controls,
ignored MCE branches, missing or external edges, wrong content types, and
unadmitted or structurally untyped owner parts are excluded from this collection.
An admitted owner containing an opaque extension subtree remains in the
collection and retains that subtree's diagnostic, including a missing extension
URI. The position is not a `shapeId` or relationship ID. `Name` resolves in the same
effective collection; duplicate names return typed ambiguity even if an
implementation could choose the first one. A future selector by semantic
worksheet control role must retain these ambiguity rules.

### Snapshot and read view

The package snapshot owns the source bytes and relationship catalogs. A read
view borrows them; it does not clone a complete worksheet or properties part:

```rust,ignore
pub struct FormControl<'a> {
    pub selector: ControlSelector<'a>,
    pub name: Option<&'a str>,
    pub properties: PropertiesView<'a>,
    pub owner_ancestry: OwnerAncestry<'a>,
    pub ownership: Ownership,
}

pub struct OwnerAncestry<'a> {
    /// Selected `mc:AlternateContent`/`mc:Choice` wrappers, in outer-to-inner
    /// order. Empty means a direct controls collection.
    pub selected_mce: &'a [MceSelection<'a>],
}

pub struct PropertiesView<'a> {
    // Every field retains source presence and decoded typed value.
    // Effective defaults are methods, not loss of omission information.
    pub object_type: Option<KnownOrUnknown<ObjectType<'a>>>,
    pub items: Option<ItemListView<'a>>,
    pub root_extensions: OpaqueExtensionList<'a>,
    // The remaining scalar accessors are field-specific and borrowed.
}

impl Worksheet {
    pub fn form_controls(&self) -> Result<FormControlCollection<'_>>;
    pub fn form_control(
        &self,
        selector: ControlSelector<'_>,
    ) -> Result<Option<FormControl<'_>>>;
}
```

The final model may use a compact generated view rather than the illustrative
field list above, but it must preserve three states for every optional
attribute: absent, present with a recognized value, and present with a
bounded unknown lexical value. A `to_owned()` conversion is explicit and
budgeted. `Properties::default()` means a valid empty `CT_FormControlPr`; it
does not mean “erase the existing part” unless a transaction verb says so.

### Transaction and ordinary worksheet verbs

The ordinary edit surface should be selector-first and chainable, matching
the existing `WorksheetEdit` style:

```rust,ignore
impl<'a> WorksheetEdit<'a> {
    pub fn set_form_control_properties(
        &mut self,
        selector: ControlSelector<'_>,
        value: Properties,
    ) -> Result<&mut Self>;

    pub fn set_form_control_scalar(
        &mut self,
        selector: ControlSelector<'_>,
        field: ScalarField,
        value: ScalarValue,
    ) -> Result<&mut Self>;

    pub fn set_form_control_items(
        &mut self,
        selector: ControlSelector<'_>,
        items: impl IntoIterator<Item = Item>,
    ) -> Result<&mut Self>;

    pub fn insert_form_control_item(
        &mut self,
        selector: ControlSelector<'_>,
        index: usize,
        item: Item,
    ) -> Result<&mut Self>;

    pub fn remove_form_control_item(
        &mut self,
        selector: ControlSelector<'_>,
        index: usize,
    ) -> Result<&mut Self>;

    pub fn clear_form_control_items(
        &mut self,
        selector: ControlSelector<'_>,
    ) -> Result<&mut Self>;

    pub fn add_form_control(
        &mut self,
        draft: FormControlDraft,
    ) -> Result<&mut Self>;

    pub fn remove_form_control(
        &mut self,
        selector: ControlSelector<'_>,
    ) -> Result<&mut Self>;
}
```

`set_form_control_properties` is a complete replacement of the typed scalar
and item state, but it applies staged values onto the current source owner. It
must not serialize a detached value over the source and thereby lose a
namespace declaration, comment, extension, or unknown attribute. The more
focused scalar and item verbs are useful for preservation tests and should be
the first implementation surface.

`add_form_control` is intentionally a graph operation, not a `ctrlProp` XML
writer alone. The eventual implementation must create the worksheet
`<controls>` collection entry, matching SpreadsheetDrawing anchor, matching
VML shape, selected and ignored MCE references, all required relationships,
the `ctrlProp` part, and the exact content-type override as one closure. It
must allocate semantic placement/identity internally and prove every existing
reference before publication. A draft containing only properties is therefore
not sufficient to add a visible form control. Until this full closure is
implemented and tested, the public method returns typed `UnsupportedGraph`
without creating any partial or unreferenced properties part.

`remove_form_control` has the same eventual closure requirement. It must remove
the selected control element, matching DrawingML anchor, matching VML shape,
references in selected and ignored MCE branches, and only the package edges and
parts proven unreachable afterward. Shared targets are retained. No method
removes only `ctrlProp` while leaving a required `x:control/@r:id` dangling,
and no initial implementation may silently narrow removal to the worksheet
control element alone.

At the lower package-owner layer, use source-bound names analogous to the
existing `Snapshot`, `Transaction`, `Commit`, and `Patch` families:

```rust,ignore
pub struct Snapshot { /* source part/range indexes and typed read projection */ }
pub struct Transaction<'a> { /* immutable base + staged scalar/list/graph ops */ }
pub struct Commit { pub snapshot: Workbook, pub patch: Patch, /* diagnostics */ }
pub struct Patch { /* private exact source/readset material + public semantic delta */ }

impl Patch {
    pub fn inverse(&self) -> Result<Self>;
    pub fn apply(&self, workbook: &mut Workbook) -> Result<Commit>;
}
```

The public patch reports a semantic selector and before/after properties. Its
private material retains every changed worksheet XML, worksheet relationship
XML, control-properties part, control-properties relationship target, and
`[Content_Types].xml` source needed for an inverse or graph removal. The
public patch does not expose those physical IDs.

## Exact XML model and grammar

### Root and child sequence

The properties part must have exactly one root expanded name:

```xml
{http://schemas.microsoft.com/office/spreadsheetml/2009/9/main}formControlPr
```

Namespace prefixes are arbitrary and may be inherited from the root
declaration context, but the namespace URI is not inferred from the `x14`
prefix or local name. A root in SpreadsheetML 2006, a strict `purl.oclc.org`
URI, an ActiveX namespace, or an arbitrary foreign URI is not this owner.

The CT sequence is strict:

```text
{x14}formControlPr := {x14}itemLst? {x14}extLst?
{x14}itemLst      := {x14}item* {x14}extLst?
{x14}item         := empty {x14}item with required unqualified @val
{x14}extLst       := zero or more {x}ext children
{x}ext            := optional opaque future payload with optional unqualified @uri
```

Here `x14` is exactly
`http://schemas.microsoft.com/office/spreadsheetml/2009/9/main`, while `x` is
the core SpreadsheetML URI
`http://schemas.openxmlformats.org/spreadsheetml/2006/main`. The four
recognized elements `formControlPr`, `itemLst`, `item`, and `extLst` are all
expanded x14 names because the x14 schema is `elementFormDefault="qualified"`.
The `extLst` type is core `CT_ExtensionList`, whose extension member is the
core `{x}ext` element. The ECMA `CT_Extension` schema permits the unqualified
`uri` attribute to be omitted; SDK-generated typed models commonly require it
when constructing a typed extension. When present, `uri` follows the core
`xsd:token` lexical rule and the child payload is opaque. A missing `uri` is
therefore not by itself a schema-invalid part, but it is an opaque diagnostic
in this typed owner: the reader retains it, exposes no typed extension member,
and emits no generated extension without a URI.

Only one direct `itemLst` and one direct root `extLst` are admitted. Only one
`itemLst/extLst` is admitted. `itemLst` is meaningful only for `List` and
`Drop`; the specification says `fmlaRange` takes precedence when present.
The codec must retain comments, legal whitespace, entity spelling, and
unknown extension payload inside the two `extLst` containers. If an `extLst`
or its member uses a missing/conflicting namespace URI, has a missing `uri`,
or has a non-core `ext` QName, retain the entire list as opaque diagnostic
content and do not expose typed extension members or author a replacement
`ext`. Scalar and item edits may still proceed with that same diagnostic when
the opaque bytes are unchanged. A second recognized owner, an unexpected
recognized child, malformed XML, or text that violates the element-only
grammar is a typed structural error for a changed owner. Unchanged unsupported
package bytes remain preservable through the package's opaque path.

`extLst` is an opaque extension-list value in this slice. The parser validates
its x14 root expanded name, the core `ext` QName and (when present) its
`xsd:token` `uri`, XML well-formedness, namespace declarations, depth, and
bounded source span, then stores its exact source range and namespace context.
A missing `uri` keeps the list in the opaque-diagnostic state described above.
It does not infer arbitrary extension URIs or execute extension payloads. There
is no typed arbitrary-extension authoring API. A scalar/item splice leaves
both extension ranges byte-for-byte unchanged.

### `CT_ListItem` and `CT_ListItems`

`{x14}item/@val` is a required **unqualified** attribute of type `xsd:string`,
not `xsd:token`. Its decoded value therefore retains intentional spaces, line
breaks, and other legal XML string characters. XML entities are decoded only
for the typed view; the
original lexical bytes are retained for a no-op and for edits to unrelated
fields. An item has no typed child content. Unknown foreign attributes or
comments are preserved in source mode and cause a typed item rewrite to
refuse if they cannot be retained by the planned splice.

An `itemLst` view contains an ordered borrowed item range, optional opaque
`itemLst/extLst`, and the source ranges between items. Inserting, deleting, or
replacing one item uses these ranges and preserves the list extension,
comments, whitespace, and root attribute lexical forms. It never sorts items
or regenerates a list from decoded strings.

### `CT_FormControlPr` attributes

The attributes below are unqualified. Namespace-qualified attributes with the
same local names are not aliases. Every XML Schema boolean/unsigned-integer
lexical form is validated before decoding; every bounded integer is checked
for overflow and the Office range. Applicability rules are checked before a
changed candidate is published.

| Attribute | Type and lexical rule | Default/effective value | Applicability or range |
|---|---|---|---|
| `objectType` | `ST_ObjectType` (`xsd:token`) | absent; no default | all control kinds |
| `checked` | `ST_Checked` (`xsd:token`) | absent | `CheckBox`, `Radio`; `Mixed` only `CheckBox` |
| `colored` | `xsd:boolean` | `false` | `Drop` |
| `dropLines` | `xsd:unsignedInt` | `8` | `Drop`; `0..30000` |
| `dropStyle` | `ST_DropStyle` (`xsd:token`) | absent | `Drop` |
| `dx` | `xsd:unsignedInt` | `80` | `List`, `Scroll`, `Spin`, `Drop` |
| `firstButton` | `xsd:boolean` | `false` | `Radio` |
| `fmlaGroup` | `ST_Formula` | absent | `GBox`; cell-reference grammar |
| `fmlaLink` | `ST_Formula` | absent | `CheckBox`, `Radio`, `Scroll`, `Spin`, `Drop`, `List`; cell-reference grammar |
| `fmlaRange` | `ST_Formula` | absent | `List`, `Drop`; cell/range-reference grammar |
| `fmlaTxbx` | `ST_Formula` | absent | `Label`, `EditBox`; cell/range-reference grammar, first cell used |
| `horiz` | `xsd:boolean` | `false` | `Scroll` |
| `inc` | `xsd:unsignedInt` | `1` | `Scroll`, `Spin`; `0..30000` |
| `justLastX` | `xsd:boolean` | `false` | text layout; preserve deprecated Office behavior |
| `lockText` | `xsd:boolean` | `false` | `Button`, `Radio`, `CheckBox`, `Label` |
| `max` | `xsd:unsignedInt` | absent | `Scroll`, `Spin`; `0..30000` |
| `min` | `xsd:unsignedInt` | `0` | `Scroll`, `Spin`; `0..30000` |
| `multiSel` | `xsd:string` | absent | `List` and only when `seltype=multi`; preserve its comma-delimited one-based lexical form |
| `noThreeD` | `xsd:boolean` | `false` | `CheckBox`, `Radio`, `GBox`, `Scroll`, `Drop`, `List`, `Spin` |
| `noThreeD2` | `xsd:boolean` | `false` | `Drop`, `List` |
| `page` | `xsd:unsignedInt` | absent | `Scroll`, `Spin`; `0..30000` |
| `sel` | `xsd:unsignedInt` | absent | `List`, `Drop`; one-based, `0` means none |
| `seltype` | `ST_SelType` (`xsd:token`) | `single` | `List` |
| `textHAlign` | `ST_TextHAlign` (`xsd:string`) | `left` | no object-type restriction in CT_FormControlPr |
| `textVAlign` | `ST_TextVAlign` (`xsd:string`) | `top` | no object-type restriction in CT_FormControlPr |
| `val` | `xsd:unsignedInt` | absent means `0` | `Scroll`, `Spin`, `List`, `Drop` |
| `widthMin` | `xsd:unsignedInt` | absent | `Drop` |
| `editVal` | `ST_EditValidation` (`xsd:token`) | absent means `text` | `EditBox` |
| `multiLine` | `xsd:boolean` | `false` | `EditBox` |
| `verticalBar` | `xsd:boolean` | `false` | `EditBox` |
| `passwordEdit` | `xsd:boolean` | `false` | `EditBox` |

The `0..30000` requirements apply to `dropLines`, `inc`, `max`, `min`, and
`page` as stated in §2.6.65. The implementation should also enforce the
cross-field ordering that a staged scroll/spin value remains meaningful for
the staged `min`/`max`; if the local specification does not define an
ordering, do not invent one for read-only input. `sel` and `multiSel` must be
validated against the staged item count when the relevant typed operation
changes the list. A source with an inapplicable but lexically valid attribute
may be read and preserved with a diagnostic, but a commit that leaves the
candidate semantically invalid must refuse rather than worsen it.

`ST_ObjectType` admits exactly `Button`, `CheckBox`, `Drop`, `GBox`, `Label`,
`List`, `Radio`, `Scroll`, `Spin`, `EditBox`, and `Dialog`. `ST_Checked` admits
`Unchecked`, `Checked`, and `Mixed`. `ST_DropStyle` admits `combo`,
`comboedit`, and `simple`. `ST_SelType` admits `single`, `multi`, and
`extended`. `ST_EditValidation` admits `text`, `integer`, `number`,
`reference`, and `formula`. `ST_TextHAlign` admits `left`, `center`, `right`,
`justify`, and `distributed`; `ST_TextVAlign` admits `top`, `center`, `bottom`,
`justify`, and `distributed`.

The first five enum families are `xsd:token`, so validation uses XML Schema
whitespace collapse before comparing the value. The alignment families are
`xsd:string`; their lexical value is not token-collapsed. A source spelling
with unusual legal string whitespace is therefore retained on no-op and on
unrelated edits, while a typed setter emits one canonical admitted value.
Unknown enum tokens are represented as bounded unknown lexical values for
lossless read/preservation, but cannot be constructed by a typed setter.

`ST_Formula` is not an arbitrary Rust string. Formula fields must be held in a
source-bound checked `FormControlFormula` wrapper that retains the original
dialect and lexical text. The validator must recognize the cell-reference
grammar used by the owner, including A1 and R1C1 references, ranges, defined
names, sheet/workbook prefixes, and external-reference forms, while applying
the field-specific cell/range restriction above. It must retain the lexical
qualifiers rather than resolving a name or opening an external workbook.

The native `tdf134769.xlsx` input contains `fmlaLink="#REF!"`. This existing
error token is read and preserved as a diagnostic; an unrelated scalar or item
edit may proceed when the complete candidate remains package/XML/safety-valid
while retaining the same pre-existing semantic diagnostic and the token is
untouched. A typed setter must never introduce `#REF!` or turn a
previously checked reference into an unvalidated error token. Unsupported
external/name syntax is likewise retained for unchanged sources and refused
for a formula replacement until its grammar is proven. The existing general
`Formula::new` helper's nonempty/XML-character checks are not by themselves
proof of these `ST_Formula` restrictions. No calculation, external-link
resolution, formula refresh, or cached-value regeneration is part of this
owner.

`multiSel` is `xsd:string`, so it is not token-collapsed. The semantic view may
offer a checked interpretation of comma-separated one-based indices, but it
must retain the exact lexical spelling (including legal whitespace) and an
unrelated edit must not rewrite it. A `multiSel` setter validates the staged
list-control context and emits the caller's checked lexical value; it does not
silently sort or normalize indices.

## Defaults, unknown values, and preservation

Presence is part of the model. For example, omitted `dropLines` reads as
`None` with `effective_drop_lines() == 8`; an explicitly stored `dropLines="8"`
reads as `Some(8)` and remains explicitly stored through an unrelated scalar
edit. This distinction is required for exact source preservation and inverse
patches. Typed builders may offer `set_drop_lines(None)` to restore omission,
or `set_effective_drop_lines(8)` only if the caller explicitly accepts the
canonical explicit attribute.

The source owner records, at minimum:

* the complete properties-part byte owner and root `formControlPr` start/end
  range;
* root attribute source ranges, decoded values, raw lexical tokens, and
  namespace context;
* direct `itemLst`, each `item`, `itemLst/extLst`, and root `extLst` ranges;
* comments, processing instructions, legal whitespace, and unknown foreign
  markup that sits between recognized children;
* the worksheet control range and `r:id` source range, including the complete
  selected MCE ancestry and the ignored branch ranges needed for byte
  preservation;
* the worksheet relationship member, target token, target mode, and source
  range; and
* the matching SpreadsheetDrawing XML and relationship ranges, VML
  `legacyDrawing`/shape ranges and relationship members, and every selected or
  ignored MCE branch that references the control identity. Scalar/item edits
  may leave these bytes untouched, but they remain read-set closure evidence;
  add/remove cannot proceed without them.
* the content-type override source range when graph creation/removal is
  staged.

`Properties::write` is therefore a validated detached writer for new parts,
not the default path for an existing owner. For an existing owner, staged
scalar values and item changes are applied onto the current source ranges.
An inherited namespace must be materialized only when a detached fragment is
being written; a direct source splice must retain the original declaration
context. Never rebind or canonicalize a complete source part merely because a
typed scalar changed.

Comments and unknown extension children containing strings that resemble XML
markup are preserved by event-aware source ranges. Byte substring heuristics
must not reject valid comment/CDATA payloads. Conversely, XML declarations or
processing instructions in a detached `formControlPr` fragment must be
validated as actual events and rejected where the part grammar disallows them.

## Transaction, stale source, inverse, and atomic publication

The owner transaction reads a source snapshot and stages operations by
semantic selector. It records a read set containing the complete worksheet XML
(including every selected and ignored MCE branch),
the worksheet relationship member, the selected properties-part bytes and
relationship state, the matching DrawingML part and its relationship member,
the matching VML part/shape and its relationship member, `[Content_Types].xml`
when graph state may change, and all incoming references used by the closure
decision. It records a write set for only the changed ranges/members.

Preflight must finish before any candidate bytes or package parts are
published. The candidate is reopened under the source workbook's retained
`ReadLimits`, then the same selector is read back and compared with the
requested semantic state. The source snapshot remains unchanged.

An exact no-op (including a scalar set to the already stored value or an item
replacement with identical source state) shares the original source bytes,
does not allocate a replacement part, and produces an empty semantic patch.
A real patch contains:

* public before/after `ControlSelector` and typed properties/item changes;
* private source and target worksheet XML bytes/ranges;
* private source and target control-properties XML bytes/ranges;
* private relationship and content-type source records when applicable;
* expected source hashes/lineage for every read-set member; and
* a deterministic diagnostic summary without exposing physical IDs in the
  ordinary API.

`Patch::apply` checks every expected read-set member, not just the properties
part. A changed worksheet control, changed relationship target, changed
content-type mapping, changed target part, or changed incoming edge returns a
typed stale/conflict error. `inverse()` applies the exact accepted target
back to the original source graph and restores the original relationship and
content-type lexical bytes. It does not regenerate a “semantically
equivalent” part and call that an inverse.

The transaction retains the complete source `ReadLimits` policy, including
fields that the current operation does not use directly. Candidate reopen,
patch application, and inverse application use that retained policy (or an
explicitly supplied stricter policy) so a feature-local shortcut cannot make
the same source acceptable under weaker aggregate limits. A caller may not
raise a hard integer, XML-depth, decompression, or graph-safety ceiling by
passing a larger feature-local budget.

If the package carries an observed signature graph, this edit is unavailable
until the caller explicitly enters the accepted unsign or edit-and-resign
operation. Rewriting a control-properties part must not silently strip package
signatures, and a successful source patch is not a cryptographic signature or
trust assertion. The patch's source authorization is exact immutable source
bytes/lineage and read-set state; signature status remains a separate package
diagnostic.

Failed validation, a limit error, an unsupported MCE/ActiveX ambiguity, an
invalid formula, or a graph-closure conflict leaves the source and transaction
unchanged. Add/remove failures return the rejected owned draft when the
ordinary API takes ownership, following the accepted ownership contract.

## Limits and allocation plan

The owner must use the workbook's retained `litchi_opc::ReadLimits`, including
`max_part_bytes`, `max_total_part_bytes`, `max_content_types_bytes`,
`max_content_type_mappings`, relationship-part/XML/event/edge limits,
`max_xml_depth`, `max_xml_attribute_bytes`, and relationship-target byte
limits. The form-control codec must not invent a second unconstrained package
budget.

The implementation should also publish feature-local hard ceilings, aligned
with the existing inert ActiveX safety profile, for values that the OPC
limits do not describe directly:

| Resource | Proposed hard ceiling |
|---|---:|
| One control-properties XML part | 16 MiB |
| One generated properties XML part | 32 MiB |
| Effective form controls per worksheet | 65,535 |
| `item` elements in one list | 65,536 |
| One decoded `item/@val` | 1 MiB of UTF-8/source bytes, subject to the part limit |
| XML nesting | 256 |
| XML event/node budget | 400,000 per changed owner, subject to retained package limits |

These are host safety ceilings, not relaxations of schema rules. A caller may
raise a configurable `ReadLimits` field only within global integer, XML-depth,
and decompression hard ceilings. The implementation must document the final
constant names before coding; a test must exercise each limit and report the
resource, observed value, and limit.

All allocations are charged before construction:

* account source bytes before retaining a properties part, item strings, or
  opaque extension ranges;
* account decoded strings and raw lexical buffers separately where both are
  retained; avoid duplicating an entire part for each item;
* precompute every source-splice output length with checked arithmetic before
  allocating a replacement vector;
* when adding a control, precharge the new properties part, worksheet XML,
  worksheet relationship XML, graph node/edge counts, and content-type
  mapping before building any of them;
* when removing a shared owner, subtract only the edges and members proven
  removable; never budget a whole package copy; and
* use borrowed source ranges and compact range indexes for reads. `Arc` may
  share the package's existing immutable source owner, but `Arc::from(xml)`
  of a complete worksheet or properties part in every scan is not acceptable.

The final output cap applies to the candidate's complete retained package,
not merely a temporary scratch upper bound. Temporary buffers must also be
bounded by the caller's effective cap; a valid 1 MiB source must not be
serialized into an unbounded replacement vector when `max_part_bytes` is 1
byte. A failed preflight must occur before materializing that candidate.

## Validation and test handoff

The production implementation should land with focused tests in a new
`litchi-xlsx` form-control target. The minimum matrix is:

1. **Exact owner admission:** canonical content type and relationship,
   arbitrary x14 prefix/default namespace, escaped namespace values,
   wrong/strict/foreign namespace, wrong content type, external target,
   missing target, malformed root, core-vs-x14 `ext` QName, present/omitted
   `ext@uri` handling (omitted URI is an opaque diagnostic), and an outgoing
   properties-part relation reported as a
   host-profile closure diagnostic rather than a schema claim.
   Open every package in `native-corpus.json` and assert all ten properties
   parts have the exact root/content type and one canonical incoming edge.
2. **Control coexistence:** ActiveX `control` edge, an embedded persistence
   edge, ordinary `x:controlPr` metadata coexistence, duplicate names, shared
   control references to one part, unreferenced `ctrlProp`, and selected versus
   ignored MCE lookalikes. The native nested MCE controls must resolve using
   first-supported `Requires="x14"` selection, while an unselected branch
   remains byte-for-byte opaque. Unknown/opaque branches must survive
   unchanged; ambiguous changes must return refusal.
3. **Typed grammar:** every scalar attribute, all defaults and presence
   states, XML Schema boolean forms, unsigned-integer overflow/0..30000
   limits, each enum and whitespace rule, field applicability, `sel` and
   `multiSel` lexical preservation, item string entities/whitespace, item
   ordering, duplicate `itemLst/extLst`, unknown attrs/elements, comments,
   CDATA, present/omitted core `ext@uri`, and opaque extension children. Cover A1/R1C1,
   defined-name, sheet/workbook-prefix, external-reference, and existing
   `#REF!` formula cases; unchanged `#REF!` survives unrelated edits while a
   setter cannot introduce it.
4. **Source preservation:** exact no-op bytes; scalar splice preserving the
   complete native worksheet MCE ancestry, unknown attrs/comments/line
   endings/prefixes, and unrelated VML/drawing children; item insert/remove
   preserving both extension lists; detached writer namespace completion; and
   malformed untouched owner diagnostics.
5. **Graph CRUD:** source-backed whole-control add/remove only after closure
   proof covers the matching DrawingML anchor, VML shape, selected and ignored
   MCE references, shared target reachability, content-type override
   insertion/removal, self-closing relationship roots, relationship-ID
   allocation, and no orphan part/mapping creation. Before that proof, assert
   `UnsupportedGraph` and exact source preservation rather than testing a
   partial control-only mutation.
6. **Transaction safety:** source changes in worksheet XML, worksheet rels,
   properties XML, content types, and incoming-edge state; exact inverse;
   stale patch refusal; same-transaction overlapping selector conflict;
   atomic source preservation after all errors.
7. **Resource limits:** each retained `ReadLimits` field relevant to the
   changed closure, one-byte-under limits, aggregate item/extension memory,
   XML depth/events, graph node/edge counts, and preallocation refusal before
   constructing an oversized temporary buffer.
8. **Schema and native evidence:** validate generated XML and the complete
   changed package against the local MS-XLSX/ECMA schema. Native Office/LibreOffice
   open, edit/save, and reverse-read evidence must use a real
   `formControlPr` archive and be reported separately from synthetic tests.

Until this matrix exists, the audit row remains ❌. The design itself is not
evidence of support, and the existing ActiveX fixtures must not be relabeled as
form-control properties evidence.

## Implementation order and handoff checklist

1. Add the private owner scanner and exact constants; reuse the common
   worksheet MCE capability resolver for effective control selection, not the
   ActiveX graph loader. Record selected MCE ancestry and preserve ignored
   branches.
2. Add source ranges and borrowed namespace context for the worksheet control,
   relationship member, content-type member, and properties-part root.
3. Implement `CT_FormControlPr` lexical parsing with explicit presence,
   unknown-value, default, and applicability states. Add `itemLst` and two
   opaque extension-list ranges before any writer.
4. Implement scalar and item source splices, preflighted against full retained
   limits, followed by complete candidate reopen/readback.
5. Add the public `Worksheet` snapshot view and selector-first
   `WorksheetEdit` verbs. Keep physical identities in diagnostics only.
6. Keep whole-control graph add/remove as typed `UnsupportedGraph` until the
   matching DrawingML anchor, VML shape, selected/ignored MCE references,
   relationship closure, ActiveX persistence prohibition, and content-type
   ownership have independent tests. Then implement the complete closure in a
   single atomic operation; do not ship a control-only approximation.
7. Add exact reversible patch application and stale/read-set tests.
8. Update `docs/report/spec-gap-audit.md` and the implementation-status matrix
   only after the focused gate, schema validation, and native evidence justify
   changing the ❌ status.

The immediate coder/test handoff is therefore: expose
`form_control::{ControlSelector, Properties, Item, FormControl, FormControlDraft}`
and the `Worksheet`/`WorksheetEdit` methods above; preserve `active_x`'s
existing `ControlProperties` names and behavior; use the pinned native corpus
for existing-owner read/scalar-edit/save/reopen tests; and do not infer a
`formControlPr` owner from ActiveX metadata when the relationship/content-type
graph is absent or ambiguous.
