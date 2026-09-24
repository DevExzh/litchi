# XLSX `formControlPr` ↔ VML `ClientData` mirror evidence

Status: bounded local evidence for the owner-design review. This note does
not add a parser, writer, package edge, or native-application claim. It
records only mappings that are established by the vendored specifications or
by the retained fixtures. Similar element names are not treated as a mapping
when the local material does not establish the meaning.

## Scope and citation key

The x14 side is the `formControlPr` control-properties part. The local
MS-XLSX definition is `CT_FormControlPr` in
`3rdparty/specs/[MS-XLSX]/2 Structures/2.6 Complex Types.md`, §2.6.65
(lines 2895–3092). Its child/list shape is `CT_ListItem`/`CT_ListItems`,
§§2.6.63–2.6.64 (lines 2848–2893). The root and namespace are
`formControlPr`, §2.4.34 in
`3rdparty/specs/[MS-XLSX]/2 Structures/2.4 Global Elements.md`, lines
395–405.

The x14 lexical enumerations are in
`3rdparty/specs/[MS-XLSX]/2 Structures/2.7 Simple Types.md`:

* `ST_ObjectType`, §2.7.14, lines 544–598;
* `ST_Checked`, §2.7.15, lines 600–630;
* `ST_DropStyle`, §2.7.16, lines 632–662;
* `ST_SelType`, §2.7.17, lines 664–694;
* `ST_EditValidation`, §2.7.18, lines 696–732; and
* `ST_TextHAlign`/`ST_TextVAlign`, §§2.7.22–2.7.23, lines 827–901.

The VML side is the Transitional ECMA-376 material retained locally:

* `3rdparty/specs/ECMA-376/ECMA-376-4_5th_edition_december_2016.zip`, PDF
  member `Ecma Office Open XML Part 4 - Transitional Migration Features.pdf`,
  §§19.4.2.11–19.4.2.68 (printed pages 836–858), and §19.4.3.2 (printed
  page 859); and
* the nested member
  `OfficeOpenXML-XMLSchema-Transitional.zip/vml-spreadsheetDrawing.xsd`.
  The extracted schema has `CT_ClientData` and its element/type declarations
  at lines 9–81, `ST_ObjectType` at lines 85–107. Its imported
  `shared-commonSimpleTypes.xsd` defines `ST_TrueFalseBlank` at lines 90–100.

The PDF section citations below are the authoritative local prose for VML
meaning, omitted-value behavior, and permitted values. The nested XSD is the
lexical/content-model citation. The local MS-OI29500 file is a product
variation source, not a general bridge between x14 attributes and VML
elements:
`3rdparty/specs/[MS-OI29500]/2 Conformance Statements/2.1 Normative
Variations.md`.

Relevant exact MS-OI29500 entries are:

* §2.1.603 (`control`), lines 8971–8979: Office shape-ID and name limits;
* §2.1.1863 (`ClientData`), lines 30946–30954: unsupported `Movie`, `LineA`,
  and `RectA` object types;
* §2.1.1865 (`Colored`), lines 30962–30966: Excel colors the dropdown arrow;
* §§2.1.1868–2.1.1870, lines 30980–30996: Excel bounds `DropLines` and
  describes `DropStyle`/`Dx` usage;
* §2.1.1871 (`FmlaGroup`), lines 30998–31002: Excel may move the value to the
  first radio button's `FmlaLink`;
* §§2.1.1875–2.1.1876, lines 31030–31036: Excel narrows `Horiz` use and
  bounds `Inc`;
* §2.1.1878 (`ListItem`), lines 31044–31048: Excel ignores `ListItem` when
  `FmlaRange` is present;
* §§2.1.1879–2.1.1880 and §2.1.1883, lines 31050–31060 and 31074–31078:
  Excel bounds `Max`, `Min`, and `Page`;
* §2.1.1882 (`NoThreeD`), lines 31068–31072: Excel also uses it for spinners;
* §2.1.1887 (`TextVAlign`), lines 31098–31102: Excel treats it as vertical
  alignment; and
* §2.1.1889 (`Val`), lines 31110–31118: Excel bounds it and uses it for more
  list-like controls; and
* §2.1.1891 (`WidthMin`), lines 31126–31130: Excel uses it for dropdowns
  only.

The local MS-OI29500 text has no `editVal`, `VTEdit`, `checked`,
`firstButton`, or `lockText` bridge statement. Those mappings below come from
the exact ECMA VML sections and MS-XLSX attribute meanings, not from a name
comparison.

## Direct conclusions from the local specifications

The two specifications define parallel concepts but do not say that every
same-named field must be mirrored. A row is *spec-proven* below only when the
field meaning, the VML value/lexical rule, and the control family line up. A
present-value proof does not prove an absence/default rule.

`ST_TrueFalseBlank` is exactly the lexical set `t`, `f`, `true`, `false`, the
empty string, `True`, and `False` (nested VML schema, lines 90–100). For the
VML fields whose prose says “if specified without a value, it is assumed to
be true,” an empty element is therefore a proven true occurrence. That prose
does not, by itself, establish that absence is false; absence is recorded
separately in the matrix.

These are the only VML Boolean lexicals admitted by the cited schema. The
type is a restricted `xsd:string`: do not add numeric `0`/`1`, `on`/`off`,
whitespace variants, or case-folded spellings. Numeric `0`/`1` belong to
integer-valued VML fields such as `Checked`, not to VML Boolean elements. The
x14 attributes use `xsd:boolean` where the CT schema says so; that does not
authorize copying an x14 Boolean lexical directly into a VML element. VML
string tokens are likewise case-sensitive prose values: `DropStyle` is
`Combo`, `ComboEdit`, or `Simple` (§19.4.2.22); `SelType` is `Single`,
`Multi`, or `Extend` (§19.4.2.58); `TextHAlign` is `Left`, `Justify`,
`Center`, `Right`, or `Distributed` (§19.4.2.60); and `TextVAlign` is
`Top`, `Justify`, `Center`, `Bottom`, or `Distributed` (§19.4.2.61). The
nested VML schema types those fields as `xsd:string`, and MS-OI29500
§2.1.1869 notes that Office permits arbitrary `DropStyle` strings at schema
level; unknown strings therefore remain read-only rather than being
case-folded or normalized. The nested schema types numeric fields such as
`Checked`, `VTEdit`, `Min`, and `Val` as `xsd:integer` (lines 39, 54, and
60–62); their permitted values and Office bounds are separate from Boolean
lexical handling.

`objectType` has no normative cross-part bridge in the local material.
MS-XLSX `ST_ObjectType` (§2.7.14, lines 550–566) lists `Button`, `CheckBox`,
`Drop`, `GBox`, `Label`, `List`, `Radio`, `Scroll`, `Spin`, `EditBox`, and
`Dialog` with x14 form-control meanings. ECMA VML `ST_ObjectType`
§19.4.3.2 (printed page 859), together with the nested XSD (lines 85–107),
lists `Button`, `Checkbox`, `Dialog`, `Drop`, `Edit`, `GBox`, `Label`,
`List`, `Radio`, `Scroll`, and `Spin` among other VML objects. The exact-token
overlap (`Button`, `Dialog`, `Drop`, `GBox`, `Label`, `List`, `Radio`,
`Scroll`, `Spin`) and the matching prose meanings are candidate evidence, but
neither specification says that `formControlPr/@objectType` must equal
`ClientData/@ObjectType`. ECMA §19.4.2.7 (printed page 834) confirms VML
attached-text meanings for `Button`, `Checkbox`, `Dialog`, `Edit`, `GBox`,
`Label`, and `Radio`, but still supplies no x14 bridge. `CheckBox ↔ Checkbox`,
`Button ↔ Button`, and `Radio ↔ Radio` remain the only admitted pairs because
the retained fixtures exercise them; `EditBox ↔ Edit` and every other
additional pair remain unresolved/read-only. MS-OI29500's local object-type
entries only reject unsupported VML `Movie`, `LineA`, and `RectA` values
(§2.1.1863, lines 30946–30954); they do not add the missing bridge.
The x14 enum is based on `xsd:token` (schema fragment lines 568–598), while
the VML enum is based on `xsd:string`; the differing `CheckBox`/`Checkbox`
spelling is therefore a fixture-backed rule, not a case-conversion rule.

The `checked` mapping is exact. MS-XLSX `ST_Checked` gives `Unchecked`,
`Checked`, and `Mixed`, with their checkbox/radio meanings (§2.7.15, lines
606–624). ECMA VML §19.4.2.11 (printed page 836) gives VML `Checked` values
`0` = unchecked/unselected, `1` = checked/selected, and `2` = mixed, and
declares the content an XML Schema integer. Thus the proven present-value
map is `Unchecked ↔ 0`, `Checked ↔ 1`, and `Mixed ↔ 2`. Other integer
values are outside the VML permitted-value table; alternate schema-valid
integer spellings also have no local lexical-normalization rule, so preserve
them as read-only rather than coercing them.

The `editVal`/`VTEdit` mapping is also exact in the local material, despite
the owner-design note having left it unresolved. MS-XLSX §2.7.18 (lines
702–726) defines `text`, `integer`, `number`, `reference`, and `formula`; ECMA
VML §19.4.2.67 (printed page 857) defines `0` = Text, `1` = Integer, `2` =
Number, `3` = Reference, and `4` = Formula. Both sections state that the
omitted value means Text. The proven map is therefore:
`text ↔ 0`, `integer ↔ 1`, `number ↔ 2`, `reference ↔ 3`, and `formula ↔ 4`,
with absence/effective Text on both sides. This is a specification proof;
the retained corpus has no edit-box fixture, so the host profile still needs
an end-to-end fixture before enabling a writer row.

`lockText` has a deliberate default mismatch. MS-XLSX `CT_FormControlPr`
declares `lockText` optional with default `false` (lines 3054–3059). ECMA
VML §19.4.2.38 (printed page 847) says that omitted `LockText` means the
object text is locked, and an empty `LockText` is true. Consequently, VML
absence cannot be treated as x14 false. The mismatch is specifically x14
effective false paired with VML absence (effective true); authored x14
`lockText="1"` paired with VML absence is an effective-true match. An
explicit VML false lexical value must remain explicit if the x14 effective
value is false. The retained fixtures show both cases: several shapes have
x14 `lockText="1"` with no VML `LockText`, while both `tdf161365.xlsx`
shapes have x14 lockText absent and `<LockText>False</LockText>`; this is
source variation, not permission to normalize absence.

`firstButton` and `noThreeD` each have an x14 default of false in the
`CT_FormControlPr` schema (lines 3042 and 3066–3068). ECMA VML §19.4.2.24
(printed page 842) and §19.4.2.45 (printed page 849) say an empty element is
true; the other `ST_TrueFalseBlank` lexicals provide explicit false values.
Neither section gives a VML omitted-value default. The matrix therefore
admits authored true/empty and explicit false pairs, but leaves absence to a
future explicit profile rule. `NoThreeD` is documented for checkboxes, radio
buttons, group boxes, and scroll bars; MS-OI29500 §2.1.1882 also records
spinner use. The x14 `noThreeD` attribute additionally names dropdowns and
list boxes, where VML has `NoThreeD2`; the local material does not state which
of those two VML fields wins for an x14 dropdown/list value. The fixtures show
the true pairs, including `firstButton="1"` with empty `FirstButton` and
`noThreeD="1"` with empty `NoThreeD`.

For formulas, MS-XLSX `fmlaLink`, `fmlaRange`, `fmlaGroup`, and `fmlaTxbx`
are `ST_Formula` attributes and the relevant fields MUST be cell references
or ranges (§2.6.65, lines 2935–2941). ECMA VML §§19.4.2.25–19.4.2.30
(printed pages 842–844) describe `FmlaGroup`, `FmlaLink`, `FmlaRange`, and
`FmlaTxbx` as string cell/formula references. The value can be carried as an
inert source string when it is valid for both fields. The fixture
`tdf134769.xlsx` contains the pre-existing `fmlaLink="#REF!"` and matching
VML `FmlaLink>#REF!</FmlaLink>`. That source-error token is retained
unchanged; the fixture proves byte preservation only, and does not prove that
the token is an acceptable new value under the MS-XLSX “MUST be a cell
reference” rule.
MS-OI29500 §2.1.1871 also warns that Office may relocate `FmlaGroup` to the
first radio button's `FmlaLink`, so a host writer must not treat that field as
a detached one-control scalar without a graph policy.

## Effective-default review

The following are effective-value comparisons from the local prose/schema.
They do not erase authored presence: a source-preserving writer still needs a
field-specific insertion/removal rule.

| x14 field ↔ VML field | x14 when absent | VML when absent | Result |
| --- | --- | --- | --- |
| `horiz` ↔ `Horiz` | schema default `false` (vertical) (`CT_FormControlPr`, line 3052) | vertical (§19.4.2.32, printed page 845) | **Effective default matches** for the scroll-bar overlap |
| `min` ↔ `Min` | schema default `0` (line 3062) | `0` (§19.4.2.41, printed page 848) | **Effective default matches** |
| `seltype` ↔ `SelType` | schema default `single` (line 3074) | `Single` (§19.4.2.58, printed page 854) | **Effective default matches**; lexical case differs by field |
| `textHAlign` ↔ `TextHAlign` | schema default `left` (line 3076) | `Left` (§19.4.2.60, printed page 855) | **Effective default matches**; lexical case differs by field |
| `textVAlign` ↔ `TextVAlign` | schema default `top` (line 3078) | `Top` (§19.4.2.61, printed page 855) | **Effective default matches**; lexical case differs by field |
| `val` ↔ `Val` | prose assumes `0` when omitted (§2.6.65, line 2991) | `0` (§19.4.2.63, printed page 856) | **Effective default matches** in the documented overlap; MS-OI29500 extends VML use to combo/list scroll bars (§2.1.1889) |
| `verticalBar` ↔ `VScroll` | schema default `false` (no vertical scroll) (line 3088) | no vertical scroll (§19.4.2.66, printed page 857) | **Effective default matches** for edit-control semantics |
| `lockText` ↔ `LockText` | schema default `false` (text unlocked) (line 3058) | text locked (§19.4.2.38, printed page 847) | **Default mismatch only for x14 effective false + VML absence**; authored x14 true + VML absence is an effective-true match |

## Retained fixture token evidence

The retained package files are under
`crates/litchi-xlsx/tests/fixtures/form_control_properties/`; their source
identity and part hashes are pinned in
`docs/report/spec-gap-validation-evidence/xlsx-form-control-properties/native-corpus.json`.
The focused test reads all seven retained packages and ten parts, checks
content type/relationship metadata, and requires an exact no-op write at
`crates/litchi-xlsx/tests/form_control_properties.rs:112-160`. Its selected
attribute assertions are at lines 166–185; the `#REF!` preservation check is
at lines 188–205. VML excerpts below are from each package's
`xl/drawings/vmlDrawing1.vml`, matched to the control shape; property excerpts
are from the named `xl/ctrlProps/ctrlPropN.xml` member.

| Fixture and member | `formControlPr` tokens | Matching VML `ClientData` tokens | Evidence use |
| --- | --- | --- | --- |
| `singlecontrol.xlsx!xl/ctrlProps/ctrlProp1.xml` | `objectType="CheckBox" checked="Checked" lockText="1" noThreeD="1"` | `ObjectType="Checkbox"`, `<Checked>1</Checked>`, `<NoThreeD/>`; no `LockText` | exact `CheckBox ↔ Checkbox`, `Checked ↔ 1`, and effective-true `1 ↔ omitted` lockText observations |
| `tdf120301_xmlSpaceParsing.xlsx!xl/ctrlProps/ctrlProp2.xml` | `objectType="Radio" firstButton="1" lockText="1" noThreeD="1"` | `ObjectType="Radio"`, `<FirstButton/>`, `<NoThreeD/>`; no `LockText` | exact Radio mapping, empty true occurrences, and effective-true `1 ↔ omitted` lockText |
| `tdf134769.xlsx!xl/ctrlProps/ctrlProp1.xml` | `objectType="CheckBox" fmlaLink="#REF!" lockText="1" noThreeD="1"` | `ObjectType="Checkbox"`, `<FmlaLink>#REF!</FmlaLink>`, `<NoThreeD/>`; no `LockText` | same inert formula bytes; source-error token is preservation-only; authored `1` and VML omission are effective-true lockText matches |
| `tdf161365.xlsx!xl/ctrlProps/ctrlProp1.xml` | `objectType="CheckBox" noThreeD="1"` (no `checked`, no `lockText`) | `ObjectType="Checkbox"`, `<LockText>False</LockText>`, `<NoThreeD/>`; no `Checked` | x14 lockText omission (effective false) paired with explicit VML false; observed omission/default distinction |
| `tdf161365.xlsx!xl/ctrlProps/ctrlProp2.xml` | `objectType="CheckBox" checked="Checked" noThreeD="1"` (no `lockText`) | `ObjectType="Checkbox"`, `<Checked>1</Checked>`, `<LockText>False</LockText>`, `<NoThreeD/>` | explicit VML false for x14 effective false plus the checked/noThreeD pairs |
| `button-form-control.xlsx!xl/ctrlProps/ctrlProp1.xml` | `objectType="Button" lockText="1"` | `ObjectType="Button"`; no `LockText` | exact Button mapping; authored `1` and VML omission are effective-true lockText matches |
| `tdf60673.xlsx!xl/ctrlProps/ctrlProp1.xml` and `xl/ctrlProps/ctrlProp2.xml` | both `objectType="Button" lockText="1"` | both `ObjectType="Button"`; neither has `LockText` | repeated owner instances confirm effective-true `1 ↔ omitted` lockText; no new lexical rule |

No retained properties part contains `itemLst`, `editVal`, `dropLines`,
`dropStyle`, `dx`, `fmlaGroup`, `fmlaRange`, `fmlaTxbx`, `horiz`, `inc`,
`justLastX`, `max`, `min`, `multiSel`, `noThreeD2`, `page`, `sel`, `seltype`,
`textHAlign`, `textVAlign`, `val`, `widthMin`, `multiLine`, `verticalBar`, or
`passwordEdit`. Those rows below are specification evidence only until a
retained package exercises both sides.

## Mapping and writer-gate matrix

The status column is deliberately split from the spec-proof column:

* **Proven** means the local specs establish the present-value semantic and
  lexical mapping. A writer still needs both authored occurrences and the
  stated absence/default rule.
* **Read-only** means the value may be parsed and preserved, but a host
  scalar writer must refuse until the listed presence/default, fixture, or
  list rule is proven.
* **Unresolved** means the local material does not establish the mapping;
  same spelling or similar names are insufficient.

“Proven present” is scoped to the overlapping control family described by the
two cited fields. A broader x14 applicability list does not authorize a VML
field outside its own documented family; those family-selection cases remain
profile gates.

| x14 `formControlPr` field | VML `ClientData` field | Proven local mapping | Absence/default and profile status | Fixture/profile token |
| --- | --- | --- | --- | --- |
| `objectType` | `@ObjectType` | `CheckBox ↔ Checkbox`, `Button ↔ Button`, `Radio ↔ Radio` are fixture-backed; ECMA §19.4.3.2 gives VML meanings and MS-XLSX §2.7.14 gives x14 meanings | Exact-token overlaps (`Dialog`, `Drop`, `GBox`, `Label`, `List`, `Scroll`, `Spin`) are candidate evidence only: no local cross-part bridge requires equality. Do not infer `EditBox ↔ Edit` or case mappings | **Proven only for the three fixture pairs; all additional pairs unresolved/read-only** |
| `checked` | `Checked` | `Unchecked ↔ 0`, `Checked ↔ 1`, `Mixed ↔ 2`; ECMA §19.4.2.11 explicitly supplies all three VML meanings | Neither side supplies an omitted-value default for this pair; absence is not equivalent to `Unchecked` without a profile rule | **Proven present; absence read-only**; `Checked`/`1` in `singlecontrol`, `tdf161365` |
| `colored` | `Colored` | Boolean semantic match; ECMA §19.4.2.14 says the empty element is true | x14 default false; VML section gives no omitted default | **Proven present; absence read-only** |
| `dropLines` | `DropLines` | Direct decimal count; x14 bounds 0–30000 (§2.6.65); VML §19.4.2.21 uses the same dropdown count | x14 default is 8; VML omission is one line, so defaults differ | **Proven present; absence read-only** |
| `dropStyle` | `DropStyle` | Exact prose tokens `combo ↔ Combo`, `comboedit ↔ ComboEdit`, `simple ↔ Simple`; both tables give the same three meanings (§2.7.16; ECMA §19.4.2.22) | No admitted omitted-value default; Office permits arbitrary `DropStyle` strings at schema level (§2.1.1869), so unknown strings stay read-only and case is preserved | **Proven present for the three tokens; unknown/absence read-only** |
| `dx` | `Dx` | Direct nonnegative width/count value for the permitted control families; ECMA §19.4.2.23 | x14 default 80; VML omission has no local default | **Proven present; absence read-only** |
| `firstButton` | `FirstButton` | An authored x14 Boolean ↔ authored VML element; empty/recognized true lexical means true and recognized false lexical means false; ECMA §19.4.2.24 says empty means true | x14 default false; VML omission has no local default | **Proven authored value; absence read-only**; only `1`/empty occurs in `tdf120301` |
| `fmlaGroup` | `FmlaGroup` | Both are inert group-box linked-cell strings; ECMA §19.4.2.25 describes the same group-box link | MS-OI29500 §2.1.1871 permits Office relocation into the first radio's `FmlaLink`; cross-control rewrite is not proven | **Read-only host field** until relocation policy is explicit |
| `fmlaLink` | `FmlaLink` | Valid cell-reference string with the same linked-cell meaning; ECMA §19.4.2.26 | No default; x14 requires a cell reference. Preserve pre-existing `#REF!`, but do not use it as a new-value validity example | **Proven valid present; source-error/presence changes read-only**; `#REF!` pair in `tdf134769` |
| `fmlaRange` | `FmlaRange` | Valid source-range string; ECMA §19.4.2.29 and MS-OI29500 §2.1.1873 cover dropdown/list use | No default; `FmlaRange` takes precedence over x14 `itemLst` and VML `ListItem` (§2.6.64; §2.1.1878) | **Proven present; list/absence read-only** |
| `fmlaTxbx` | `FmlaTxbx` | Valid cell/range string; ECMA §19.4.2.30 and MS-OI29500 §2.1.1874 | No default; cell/range syntax must remain inert and source-preserving | **Proven present; fixture needed before writer** |
| `horiz` | `Horiz` | Boolean; ECMA §19.4.2.32 says empty means true and omission means vertical | x14 default false (vertical), matching VML omission | **Proven including absence/default** |
| `inc` | `Inc` | Direct decimal increment, with x14/MS-OI29500 0–30000 bound | x14 default 1; VML omission is 0 (§19.4.2.33) | **Proven present; absence read-only** |
| `justLastX` | `JustLastX` | Boolean semantic match; ECMA §19.4.2.34 says empty means true | x14 default false; VML omission has no local default | **Proven present; absence read-only** |
| `lockText` | `LockText` | Explicit boolean semantic match; ECMA §19.4.2.38 supplies empty=true and explicit false is schema-valid | x14 default false; VML omission means true. Label a mismatch only for x14 effective false + VML absence; authored x14 true + VML absence is an effective-true match | **Proven explicit; effective-false/absence transition is read-only** |
| `max` | `Max` | Direct scroll/spin maximum; x14/MS-OI29500 bounds 0–30000; ECMA §19.4.2.40 | x14 has no default; VML omission is computed from the last visible item | **Proven present; absence read-only** |
| `min` | `Min` | Direct nonnegative scroll/spin minimum; x14/MS-OI29500 bounds 0–30000; ECMA §19.4.2.41 | Both x14 and VML omission are 0 | **Proven including absence/default** |
| `multiSel` | `MultiSel` | Both are comma-delimited selected-item indices; x14 §2.6.65 and ECMA §19.4.2.44 | No default; preserve exact commas/whitespace and do not normalize or sort | **Proven present semantic; list edit read-only** |
| `noThreeD` | `NoThreeD` | Boolean semantic match for checkbox, radio, group-box, and scroll-bar families; ECMA §19.4.2.45 says empty means true; MS-OI29500 §2.1.1882 also records spinner use | x14 default false; x14 also names dropdown/list controls, but local material does not choose `NoThreeD` versus `NoThreeD2` for them; VML omission has no local default | **Proven present only in the overlapping families; field-choice/absence read-only**; `1`/empty in five retained shapes |
| `noThreeD2` | `NoThreeD2` | Boolean semantic match for dropdown/list controls; ECMA §19.4.2.46 says empty means true | x14 default false; x14 also has `noThreeD` for these families and local material does not state precedence; VML omission has no local default | **Proven present when the profile selects this field; field-choice/absence read-only** |
| `page` | `Page` | Direct decimal page increment; x14/MS-OI29500 bounds 0–30000; ECMA §19.4.2.47 | No default in either cited field description | **Proven present; absence read-only** |
| `sel` | `Sel` | Direct one-based selected-item index for the list-box overlap; x14 value 0 means none and ECMA §19.4.2.57 says VML 0 means none | VML omission means 0; x14 also allows dropdowns, but the local VML section documents `Sel` for list boxes only; x14 absence is not silently converted | **Proven list-box present/zero; dropdown/absence read-only** |
| `seltype` | `SelType` | Exact prose tokens `single ↔ Single`, `multi ↔ Multi`, `extended ↔ Extend`; tables give the same selection meanings (§2.7.17; ECMA §19.4.2.58) | x14 default `single`; VML omission defaults `Single`; do not case-fold arbitrary strings | **Proven including absence/default for listed tokens** |
| `textHAlign` | `TextHAlign` | Exact prose token mapping `left ↔ Left`, `center ↔ Center`, `right ↔ Right`, `justify ↔ Justify`, `distributed ↔ Distributed`; §2.7.22 and ECMA §19.4.2.60 list the same meanings for the attached-text overlap | x14 default `left`; VML omission defaults `Left`; do not case-fold arbitrary strings | **Proven including absence/default for listed tokens in the attached-text overlap** |
| `textVAlign` | `TextVAlign` | Exact prose token mapping `top ↔ Top`, `center ↔ Center`, `bottom ↔ Bottom`, `justify ↔ Justify`, `distributed ↔ Distributed`; §2.7.23 and ECMA §19.4.2.61 list the same meanings for the attached-text overlap | x14 default `top`; VML omission defaults `Top`; do not case-fold arbitrary strings | **Proven including absence/default for listed tokens in the attached-text overlap** |
| `val` | `Val` | Direct scroll/spin position; MS-OI29500 §2.1.1889 extends VML use to combo/list scroll bars; x14 also allows list boxes/dropdowns | Both cited fields omit to 0; preserve one-based positions and 0, while refusing an unsupported control-family transition | **Proven including absence/default in the documented overlap; family transition read-only** |
| `widthMin` | `WidthMin` | Direct minimum dropdown width; ECMA §19.4.2.68 | No default in either cited field description | **Proven present; absence read-only** |
| `editVal` | `VTEdit` | `text ↔ 0`, `integer ↔ 1`, `number ↔ 2`, `reference ↔ 3`, `formula ↔ 4`; MS-XLSX §2.7.18 and ECMA §19.4.2.67 | Both omissions mean Text | **Spec-proven including absence/default; fixture still required** |
| `multiLine` | `MultiLine` | Boolean semantic match; ECMA §19.4.2.43 says empty means true | x14 default false; VML omission has no local default | **Proven present; absence read-only** |
| `verticalBar` | `VScroll` | Boolean semantic match under the different names; ECMA §19.4.2.66 says empty=true and omission means no vertical scroll | x14 default false, matching VML omission | **Proven including absence/default** |
| `passwordEdit` | `SecretEdit` | Boolean semantic match under the different names; ECMA §19.4.2.56 says empty means true | x14 default false; VML omission has no local default | **Proven present; absence read-only** |
| `itemLst/item/@val` | ordered `ListItem` | Both carry list-item strings; x14 CT_ListItems is ordered and VML §19.4.2.36 is a persisted non-linked list item | No local rule pairs counts/order or says how to handle `FmlaRange` precedence during a two-part edit; MS-OI29500 §2.1.1878 says VML `ListItem` is ignored when `FmlaRange` is present | **Read-only/unresolved host list mirror** |

## Implementation consequences

1. A candidate scalar writer may use a row only after it has both source
   occurrences, a field-specific lexical encoder, and an absence/default rule
   from the matrix. The x14 source presence and the VML source presence remain
   separate facts; effective values do not authorize deletion or insertion by
   themselves.
2. The strongest first-batch rows are `checked` when both elements are
   present, `horiz`, `min`, `seltype`, `textHAlign`, `textVAlign`, `val`,
   `editVal`, and `verticalBar`, subject to a retained fixture/profile for
   each control family. `editVal` is no longer blocked by missing
   specification semantics; it is blocked only by the current corpus lacking
   an edit-control fixture.
3. `lockText` needs an explicit profile branch for VML absence because its
   effective default is true while x14 `lockText` defaults false. Compare
   effective values before labeling a mismatch: authored x14 true plus VML
   absence is an effective-true match. A writer changing an effective false
   value must emit an explicit VML false token (`f`, `false`, or `False`)
   selected by the lexical profile; it must not use `0`/`1` or remove
   `LockText`.
4. `firstButton` and the applicable `noThreeD`/`noThreeD2` field may compare
   authored true to an empty VML element and authored false to a recognized
   false lexical value. They must refuse an absent transition until the
   profile defines the VML absence rule, and refuse a dropdown/list field
   choice until the profile distinguishes `NoThreeD` from `NoThreeD2`. The
   retained empty-element observations are not evidence that absence means
   false.
5. `fmlaGroup`, `itemLst`/`ListItem`, and any operation involving
   `FmlaRange` need graph-level precedence/relocation handling. A detached
   source splice can leave the other control or list representation stale.
6. Preserve a pre-existing `#REF!` byte sequence during an unrelated edit, as
   the focused test does, but do not accept a new `#REF!` `fmlaLink` value
   merely because the corpus preserves one; validate new values against the
   MS-XLSX cell-reference requirement. No formula is evaluated.
7. Do not widen `objectType` by spelling or case. The retained profile proves
   only `CheckBox ↔ Checkbox`, `Button ↔ Button`, and `Radio ↔ Radio`; a new
   token needs an exact spec/fixture pair.

This evidence supports a source-backed, inert mirror reader and a narrowly
profiled writer. It does not establish Office acceptance, rendering, macro
execution, ActiveX persistence, list CRUD, or whole-control graph mutation.
