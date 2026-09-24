# ODS data-style vocabulary gap design

Status: bounded design and evidence only. This note changes no production API,
implementation, feature matrix, or audit row. The owned implementation family
is `litchi-ods`; the ODF formula crate is not involved. The design is based on
the checked-in ODF resource and the current shared worktree as of 2026-09-12.

## Verified boundary

The ODS audit still has a real gap for the named vocabulary. The audit lists
`number:fraction`, `number:scientific-number`, `number:embedded-text`, and
`number:transliteration-*` as unsupported in
[`docs/report/spec-gap-audit.md`](../spec-gap-audit.md). The current feature
matrix describes the adjacent capability accurately: the unified document root
can add, same-family replace, and conservatively remove a closed automatic
style graph, and can resolve a limited effective cell-style projection. It does
not expose a public `styles.xml` graph editor.

The current owner is concentrated in
[`crates/litchi-ods/src/advanced.rs`](../../../crates/litchi-ods/src/advanced.rs)
and the unified transaction methods in
[`crates/litchi-ods/src/document.rs`](../../../crates/litchi-ods/src/document.rs):

* `NumberStyleNode` can only emit a `number:number-style` containing the
  existing decimal `number:number`, with mandatory bounded
  `decimal-places` and `min-integer-digits` fields plus simple prefix/suffix
  text.
* `DataStyleNode` only authors date, time, currency, percentage, and boolean
  roots. Its currency and percentage number children have no embedded-text
  representation. The `TextStyleNode` already present is a
  `style:style family="text"` run style, not the ODF `number:text-style` data
  style.
* There is no data-style definition parser. Cell-style projection retains a
  `style:data-style-name` string and supported cell/text properties, but does
  not resolve a data-style body or its attributes.
* `automatic_style_kind` classifies a `number:number-style` by its outer
  element. It does not inspect whether the body is decimal, scientific, or a
  fraction. Consequently an existing unsupported number-style can pass the
  current same-family check and be replaced by a canonical decimal style; that
  operation can discard the unsupported child. The new implementation must
  classify the body before replacement and refuse when it cannot preserve it.
* `put_style_graph` and `replace_style_graph` currently splice
  `content.xml`'s `office:automatic-styles`. Existing `styles.xml` definitions
  are inspected and preserved but are not ordinary style-graph edit targets.
  `set_cell_style`, worksheet edits, and rich-cell style integration already
  provide the ordinary reference side of the transaction.

At the checked-in baseline inspected for this note, production Rust has no
typed support for the four named child/attribute tokens, and the existing
style transaction tests cover the closed date/time/currency/percentage/
boolean families only. That absence is baseline evidence, not an acceptance
invariant: once the pending extension lands, the source occurrence check must
change with it. The persistent safety requirement is that the legacy
replacement path gain a guard which refuses an unsupported number-style body
instead of treating every outer `number:number-style` as a decimal style. The
retained native LibreOffice corpus does contain the missing vocabulary,
however; exact tracked files and member hashes are recorded below. The audit
row therefore remains actionable rather than stale.

## Normative basis

The checked-in
[OpenDocument 1.4 archive](../../../3rdparty/specs/OpenDocument-v1.4-os.zip)
has SHA-256
`9867665f9702b365076c2c6557b23c8c938959b443f6f50712fdb2d0dfb8aac4`.
The design uses these local entries, rather than inferred behavior:

* `schemas/OpenDocument-v1.4-schema.rng` defines `number-number-style`,
  `any-number`, `number-number`, `number-scientific-number`,
  `number-fraction`, `number-embedded-text`, `number-text-style`, and the
  `common-data-style-attlist`.
* `part3-schema/OpenDocument-v1.4-os-part3-schema.html` sections 16.29.1–
  16.29.7 and 16.29.26 describe the data-style particles. Sections 19.343,
  19.345, 19.346, 19.347, 19.349–19.358, and 19.365–19.368 define the
  attributes and their defaults.

For reproducibility, the archive member hashes are
`schemas/OpenDocument-v1.4-schema.rng` →
`4034ec6be29205d5fc1ee5f42468ac6ef824287b3aba6d9289032af4fafbda7f` and
`part3-schema/OpenDocument-v1.4-os-part3-schema.html` →
`43fb603f9f54f030db7082518aff6f136297d182abb20124a859d271a7969a15`.

The schema is the source of particle cardinality. In particular, the prose
describes one `number:embedded-text` child under `number:number`, while the
v1.4 RNG uses `zeroOrMore`; the implementation must accept and preserve every
bounded child permitted by the RNG. The prose rules still apply: a data style
must not put two `number:text` elements in sequence and may contain at most one
`number:fill-character`.

The implementation boundary follows
[ADR 0001](../../adr/0001-priorities-and-api-layers.md) (typed ordinary APIs
and typed refusal), [ADR 0002](../../adr/0002-crate-topology.md) and
[ADR 0023](../../adr/0023-odf-family-crate-split.md) (`litchi-ods` owns the
grammar), [ADR 0003](../../adr/0003-snapshots-edits-and-patches.md) and
[ADR 0004](../../adr/0004-semantic-api-design.md) (immutable snapshots,
selector-first atomic edits), [ADR 0005](../../adr/0005-io-memory-and-performance.md)
and [ADR 0006](../../adr/0006-validation-security-and-compatibility.md)
(bounded borrowed parsing, preservation, and fail-closed mutation), and
[ADR 0007](../../adr/0007-office-object-models.md) (inert data-bearing models).

## Smallest complete implementation batch

The complete batch is one typed data-style projection and one source-checked
automatic-style edit extension. It must close all four named tokens together;
adding only a fraction or only a scientific variant would leave the same
number-style grammar and preservation hazard in place.

### Semantic value shape

Use one shared metadata value for every data-style root. Optional fields retain
whether an attribute was omitted. The effective accessors apply the ODF defaults
without writing an omitted attribute back into the source:

```rust
pub struct DataStyleAttributes {
    pub display_name: Option<String>,       // style:display-name
    pub language: Option<String>,           // number:language
    pub country: Option<String>,            // number:country
    pub script: Option<String>,             // number:script
    pub rfc_language_tag: Option<String>,   // number:rfc-language-tag
    pub title: Option<String>,              // number:title
    pub volatile: Option<bool>,             // style:volatile
    pub transliteration: Transliteration,
}

pub struct Transliteration {
    pub format: Option<String>,             // one Unicode Nd digit whose value is 1
    pub language: Option<String>,           // schema countryCode in v1.4 RNG
    pub country: Option<String>,            // schema countryCode
    pub style: Option<TransliterationStyle>,
}

pub enum TransliterationStyle { Short, Medium, Long }
```

`format` remains a string in the semantic model because the ODF attribute type
is `string` and its prose constraint is a particular Unicode decimal digit
character, not necessarily the ASCII spelling. Validation accepts exactly one
Unicode `Nd` character with numeric value 1. The two language fields follow the
RNG (`countryCode`) even though the prose calls the language attribute an RFC
5646 language code; they are not silently normalized.

The existing `DataStyleNode` enum and `NumberStyleNode` struct are public and
are constructed directly by callers. Adding fields to their variants would be
a source-breaking change, so the extension must be additive. Keep those
closed-family types and their existing `StyleGraph` path unchanged. Add
extended number/data-style node types and an extended graph/builder for the
new vocabulary; do not claim that the old direct constructors preserve source
compatibility after a field or variant change. The extended types own one
shared `DataStyleAttributes` value rather than duplicating transliteration
fields.

A data-style catalog must recognize `number:text-style` as a metadata host,
but its body has a separate opaque-body type in this batch. This is distinct
from the existing `TextStyleNode`: a text data style is needed because
§19.365–368 explicitly lists it as a host for all four transliteration
attributes. The first batch exposes its validated metadata and an
attribute-preserving patch, while full `number:text-style` content semantics
remain outside the batch. Existing date/time/currency/percentage/boolean
roots with bodies outside the closed projection receive the same opaque-body
catalog treatment; they are not silently coerced into the legacy enum.

The effective defaults are:

| Attribute | ODF type and effective behavior |
|---|---|
| `number:transliteration-format` | Schema `string`; absent means ASCII Latin-Indic digits and causes the other transliteration attributes to be ignored. Its effective value is the ASCII digit `1`. |
| `number:transliteration-language`, `number:transliteration-country` | Schema `countryCode` in the RNG. If no language/country locale combination is supplied, use the data-style locale; this crate records the values but does not implement locale lookup. |
| `number:transliteration-style` | `short`, `medium`, or `long`; effective default `short`. Semantics are locale- and implementation-dependent. |
| `number:decimal-places` | Schema `integer`; if absent on `number:number` or `number:scientific-number`, use the default table-cell style's value. |
| `number:min-decimal-places` | Schema `integer`; it must not exceed `decimal-places`. For decimal `number:number`, absent means `0` when `decimal-replacement` is empty and otherwise `decimal-places`; for scientific numbers it means `decimal-places`. |
| `number:display-factor` | Schema `double`; effective default `1`. It scales by division before display. |
| `number:grouping` | Schema `boolean`; effective default `false`. |
| `number:exponent-interval` | Schema `positiveInteger`; effective default `1`. |
| `number:forced-exponent-sign` | Schema `boolean`; effective default `true`. |
| `number:max-denominator-value` | Schema `positiveInteger`; absent means no maximum and is ignored when `denominator-value` is present. |
| `number:min-integer-digits` | Schema `integer`; absent on a fraction means no integer portion. |
| `number:denominator-value` | Schema `integer`; absent selects an appropriate denominator. |
| `number:min-numerator-digits`, `number:min-denominator-digits`, `number:min-exponent-digits` | Schema `integer`; no additional ODF default is specified. |
| `number:position` | Schema `integer`; position is 1-based from the right of the integer portion, and text is inserted before that digit. Authoring must reject a non-positive position because it has no valid 1-based display position. |

The table-cell-style default is context-dependent. A catalog projection must
retain an omitted `decimal-places` as `Inherited`, not invent a number. A
cell-aware resolver returns `Explicit(value)`, `Inherited(value)` when the
default table-cell style is found, or `Unresolved` when no cell/default-style
context was supplied. The same state is used for a scientific style's
`min-decimal-places` when it inherits `decimal-places`. A style-level accessor
therefore reports the declared `Option<i64>` and a separate resolver result;
it never reports `0` or another fabricated concrete default.

The decimal/scientific/fraction fields need an explicit omitted state. A
recommended model, using bounded signed integers for schema `integer`, is:

```rust
pub struct ExtendedNumberStyleNode {
    pub name: String,
    pub attributes: DataStyleAttributes,
    pub leading: Option<NumberStyleAffix>,
    pub format: Option<NumberFormat>,
    pub trailing: Option<NumberStyleAffix>,
}

pub struct NumberStyleAffix {
    pub text: Option<String>,
    pub fill_character: Option<String>,
    pub text_after_fill: Option<String>,
}

pub enum NumberFormat {
    Decimal(DecimalNumber),
    Scientific(ScientificNumber),
    Fraction(FractionNumber),
}

pub struct DecimalNumber {
    pub decimal_places: Option<i64>,
    pub min_decimal_places: Option<i64>,
    pub min_integer_digits: Option<i64>,
    pub grouping: Option<bool>,
    pub decimal_replacement: Option<String>,
    pub display_factor: Option<OdfDouble>,
    pub embedded_text: Vec<EmbeddedText>,
}

pub struct ScientificNumber {
    pub decimal_places: Option<i64>,
    pub min_decimal_places: Option<i64>,
    pub min_integer_digits: Option<i64>,
    pub grouping: Option<bool>,
    pub min_exponent_digits: Option<i64>,
    pub exponent_interval: Option<u64>,
    pub forced_exponent_sign: Option<bool>,
}

pub struct FractionNumber {
    pub min_numerator_digits: Option<i64>,
    pub min_denominator_digits: Option<i64>,
    pub denominator_value: Option<i64>,
    pub max_denominator_value: Option<u64>,
    pub min_integer_digits: Option<i64>,
    pub grouping: Option<bool>,
}

pub struct EmbeddedText {
    pub position: i64,
    pub text: String,
}

pub struct ExtendedDataStyleNode {
    pub name: String,
    pub attributes: DataStyleAttributes,
    pub body: ExtendedDataStyleBody,
}

pub enum ExtendedDataStyleBody {
    Date,
    Time { decimal_places: Option<i64> },
    Currency { symbol: String, number: DecimalNumber },
    Percentage { number: DecimalNumber },
    Boolean,
}

pub struct OpaqueDataStyleNode {
    pub name: String,
    pub family: DataStyleFamily,
    pub attributes: DataStyleAttributes,
    pub body: OpaqueDataStyleBody,
}

pub enum OpaqueDataStyleBody { Preserved }

pub enum DataStyleFamily {
    Number, Date, Time, Currency, Percentage, Boolean, Text,
}

pub struct StyleGraphExtension {
    pub number_styles: Vec<ExtendedNumberStyleNode>,
    pub data_styles: Vec<ExtendedDataStyleNode>,
}

pub enum DataStyleCatalogEntry {
    Number(ExtendedNumberStyleNode),
    Existing(ExtendedDataStyleNode),
    Opaque(OpaqueDataStyleNode),
    Text(TextDataStyleCatalogEntry),
}

pub enum DecimalPlacesResolution {
    Explicit(i64),
    Inherited(i64),
    Unresolved,
}

pub struct TextDataStyleCatalogEntry {
    pub owner: DataStyleOwner,
    pub name: String,
    pub attributes: DataStyleAttributes,
    pub body: OpaqueDataStyleBody,
}
```

`StyleGraphExtension` is the additive graph containing
`ExtendedNumberStyleNode`/`ExtendedDataStyleNode`; it is passed to new
`put_extended_style_graph` and `replace_extended_style_graph` methods. The
legacy `StyleGraph`, `NumberStyleNode`, and `DataStyleNode` fields remain
available for existing direct struct/enum construction. An opaque catalog
entry is not a raw XML escape hatch: callers can inspect its name, family, and
metadata or request a metadata patch, but cannot access or regenerate its
body. `ExtendedDataStyleNode` derives its family from its single
`ExtendedDataStyleBody` variant; there is no second family discriminator that
could contradict the emitted element.

The `Extended*` names above are illustrative design names, not a request to
ship duplicate compatibility layers. Implementation review should settle one
concise `data_style` module and one additive type family before coding, while
leaving the legacy public types intact.

The `DataStyleCatalogEntry` variants are typed only for the explicitly
supported body grammar. `number:text-style` is returned as the separate
`TextDataStyleCatalogEntry` opaque-body value in this batch, as are recognized
existing-family roots containing unsupported children or extension
attributes. The catalog therefore reports that a style exists without
pretending that a renderer or a lossless semantic replacement is available.

`OdfDouble` is a validated `double` value that retains its source lexical form
until a deliberate canonical rewrite. If the crate does not already have that
wrapper, add one rather than reducing a valid XML `double` to an untracked
`f64`. Likewise, `i64`/`u64` are bounded API representations of XML Schema
`integer`/`positiveInteger`; a valid source value outside those bounds remains
preservable but is not a typed rewrite target and returns a limit/unsupported
projection result. This is a safety boundary, not a new ODF restriction.

`DecimalNumber` is shared by `number:number` under `number:number-style`,
`number:currency-style`, and `number:percentage-style`, so embedded text is
not accidentally limited to plain number styles. The existing currency,
percentage, and simple decimal constructors should become convenience
constructors over this structure. Existing field-shaped constructors can be
mapped into the additive extended types by builders; keeping duplicate scalar
fields in one extended value would make it possible for the emitted XML and
semantic value to disagree.

### Exact particles and supported rewrite envelope

The parser and serializer must implement these particles:

* `number:number-style`: required common data-style attributes; optional
  style text properties; optional leading `number:text-with-fillchar`; optional
  one `any-number`; optional trailing `number:text-with-fillchar`; then zero
  or more `style:map` elements. `any-number` is exactly one of decimal
  `number:number`, `number:scientific-number`, or `number:fraction`.
* `number:number`: its optional decimal/display/grouping/integer attributes and
  zero or more bounded `number:embedded-text` children.
* `number:scientific-number`: decimal-place/min-decimal-place,
  min-integer/grouping, min-exponent, exponent-interval, and forced-sign
  attributes; no children.
* `number:fraction`: min-numerator, min-denominator, denominator-value,
  max-denominator, min-integer, and grouping attributes; no children.
* `number:embedded-text`: required `number:position` and character data only.
* `number:text-style`: common data-style attributes, optional text properties,
  optional leading text/fill, zero or more `number:text-content` plus optional
  text/fill sequences, and zero or more style maps. The body is retained and
  can be selected for an attribute-only patch, but is not projected as a
  rendering model in this batch.

The general data-style constraints are checked across the whole root: no two
`number:text` elements in sequence and at most one `number:fill-character`.
The typed graph authoring path supports the affix shape above; style maps and
style text properties are retained as source-owned markup and cause a full
semantic replacement to refuse rather than disappear. An attribute-only
patch can still update transliteration fields on such a style because it
splices only the selected attributes after validating the original owner.

### Selectors, builders, and ordinary edits

Keep the existing dependency-checked `StyleGraph` as the compatibility
transaction primitive. Add an additive `StyleGraphExtension` and checked
builders instead of adding fields to the public legacy structs or asking
callers to hand-assemble all new vectors. The extension builder should expose
`decimal`, `scientific`, and `fraction` constructors, setters for every field
above, and `build() -> Result<StyleGraphExtension>` that checks names,
particles, cross-field constraints, and graph dependencies.
The ergonomic pattern is conceptually:

```rust
let graph = StyleGraphExtension::builder()
    .number_style(
        ExtendedNumberStyleBuilder::fraction("Fraction")
            .min_numerator_digits(2)
            .max_denominator_value(999),
    )
    .build()?;

let mut edit = snapshot.edit();
edit.put_extended_style_graph(&graph)?;
// FractionCell is an existing table-cell style that already names Fraction.
edit.set_cell_style("Sheet1", 0, 0, "FractionCell")?;
let committed = edit.commit()?;
```

Names remain part of the ordinary selector because ODF style references are
names, but a name alone is not enough for source editing. The same name can be
present in `content.xml` automatic styles and `styles.xml`, and a producer can
reuse names across data-style families. Use an owner and family in every
mutation selector:

```rust
pub enum DataStyleOwner { ContentAutomatic, CommonStyles, StylesAutomatic }

pub struct DataStyleSelector<'a> {
    pub owner: DataStyleOwner,
    pub family: DataStyleFamily,
    pub name: &'a str,
}
```

`ContentAutomatic` selects direct `office:automatic-styles` in `content.xml`
and is the only mutable owner in this batch. `CommonStyles` selects direct
`office:styles` in `styles.xml`; `StylesAutomatic` selects the separate direct
`office:automatic-styles` container in that same member. Both styles-member
owners are valid for inspection and typed catalog lookup and refuse every
mutation. Their scopes remain distinct even when names coincide. A selector that omits
owner or family must first resolve to exactly one catalog entry; otherwise it
returns an ambiguity error instead of selecting an owner or family by
precedence. `DataStyleSelector::automatic(name, family)`,
`DataStyleSelector::common(name, family)`, and
`DataStyleSelector::styles_automatic(name, family)` are the concise constructors.
The styles-member automatic scope is not an implicit fallback for a reference
originating in `content.xml`.

A `DataStyleAttributePatch` with `Keep`, `Set`, and `Clear` operations should
allow metadata edits on number/date/time/currency/percentage/boolean and text
data-style roots, including a `number:text-style` whose body is otherwise
opaque. The patch must locate exactly one style in the selected
`ContentAutomatic` owner and preserve its child/body bytes. It must refuse a
`CommonStyles` or `StylesAutomatic` target explicitly, even when the source
body is fully typed. `Set("")` retains a present empty attribute when its
scalar type permits that value; `Clear` removes the attribute. Setting an
existing attribute to its successfully decoded semantic value preserves that
attribute’s exact source spelling, including character references. The whole
patch is an exact source no-op only when every operation is a semantic no-op.
Malformed or unsupported lexical values must not be normalized through this
equality rule.

The existing ordinary integration stays source-checked and atomic:

1. `Snapshot::edit()` stages the graph and cell-style reference in one clone.
2. `put_extended_style_graph` adds the new automatic definitions to the
   existing `content.xml` automatic-style owner. It validates all dependencies
   before changing the candidate. The old `put_style_graph` remains available
   for the closed compatibility graph.
3. `set_cell_style`, worksheet cell edits, rich-cell style references, and the
   mutable facade continue to reference the resulting style name. They do not
   need a second number-format path.
4. `replace_extended_style_graph` first parses the selected style's actual
   body and replaces only a matching semantic family. A decimal replacement
   cannot silently replace a scientific or fraction body. The old
   `replace_style_graph` keeps its existing contract for legacy graph nodes;
   it must refuse a selected source node whose body is outside that legacy
   contract rather than treating every `number:number-style` as decimal.
5. `remove_automatic_styles` keeps its package-wide retained-reference check.
   Removing an unreferenced unsupported style is intentional deletion; it is
   not a claim that the style body is semantically understood.

`effective_cell_style` should retain its current `data_style` name for source
compatibility and add a definition accessor that returns the typed data-style
projection when the selected body is supported. Resolution must search the
direct `content.xml` automatic styles with the current precedence, then direct
`office:styles` definitions in `styles.xml`, preserving the latter as a
read-only source owner. `StylesAutomatic` requires explicit owner selection
and is excluded from this implicit content-cell fallback. A style with
an unsupported body can still be opened and preserved, but its body accessor
and semantic replacement return a typed unsupported result. When a decimal
default depends on the cell's table-cell style, the accessor returns the
explicit/inherited/unresolved resolution state described above.

## Preservation, limits, and failures

The implementation follows the accepted ODF ADR rules: immutable snapshots,
source-checked edits, exact no-op publication, bounded work, and fail-closed
typed mutation.

The current limits are owner-specific. `advanced::scan` counts up to its
`MAX_ELEMENTS = 1_048_576` for `content.xml`, but has no generic nesting-depth
cap. `flat.rs`'s `MAX_XML_DEPTH = 1_024` applies only to flat FODS parsing.
`worksheet::validation::MAX_CONTENT_XML_BYTES = 256 MiB` bounds worksheet
content paths; it does not automatically bound `styles.xml`. The data-style
scanner must therefore introduce and apply the same explicit owner budget to
each scanned XML owner rather than assuming that a content limit covers the
styles member.

* Opening and an untouched commit must retain the original package bytes and
  every unselected XML member. In particular, unknown attributes, foreign
  namespace children, prefix spelling, attribute order, style-map conditions,
  text properties, and lexical numeric spellings remain intact.
* A new extended graph may use canonical compact XML in the selected
  `content.xml` automatic-style owner. A replacement may canonicalize exactly
  the selected style element only after proving that every source child and
  attribute is in the supported envelope. Unselected styles, XML owners, ZIP
  members, and opaque source bodies stay byte-identical. A selected supported
  style is deliberately allowed to change lexical spelling under that
  canonical replacement. Exact no-op publication and opaque-body or
  attribute-only edits retain the original lexical bytes. Otherwise return a
  typed unsupported/refusal error and leave the candidate unchanged.
  `styles.xml` is explicitly read-only in this batch, including for metadata
  patches.
* Keep the existing `4_096` node bound for an authored
  `StyleGraphExtension` (it is an edit-graph bound, not a source-catalog
  bound). The source catalog uses separate named budgets:
  `MAX_STYLE_CATALOG_ENTRIES = 65_536` recognized data-style entries per XML
  owner, `MAX_STYLE_BODY_ELEMENTS = 4_096` child/particle elements in one
  data-style body, `MAX_STYLE_TEXT_BYTES = 65_536` bytes in each text or
  metadata scalar (the existing `invalid_style_text` bound),
  `MAX_STYLE_METADATA_BYTES = 64 KiB` of aggregate common metadata for one
  style, and `MAX_STYLE_BODY_BYTES = 1 MiB` of aggregate body payload for one
  style. These source budgets are distinct from the edit-graph bound and
  prevent an RNG `zeroOrMore` from causing an unbounded allocation.
* Define owner scanner hard ceilings of
  `MAX_STYLE_OWNER_BYTES = 256 MiB`,
  `MAX_STYLE_OWNER_ELEMENTS = 1_048_576`, and
  `MAX_STYLE_OWNER_DEPTH = 1_024` for each scanned `content.xml` or
  `styles.xml` owner. All counters are checked before reservation. For a
  packaged `document::Snapshot`, the effective byte ceiling is
  `min(MAX_STYLE_OWNER_BYTES, limits.max_package_bytes())`; the actual member
  length must also fit that ceiling. The current `document::Limits` has no
  separate owner-byte or depth field, so the 1,024 depth and 1,048,576 element
  ceilings always apply; if owner-specific caller limits are added later, the
  effective values are the minimum of hard and caller limits. For a flat
  `FlatSnapshot`, the caller's `FlatLimits.input_bytes` is the additional byte
  minimum and the existing flat depth/element ceilings apply. Output uses the
  corresponding caller output/package ceiling. This makes explicit that the
  existing content limit does not automatically protect `styles.xml`.
* Continue to enforce the existing default package/resource limits where they
  are lower than these scanner ceilings, XML-depth/scan limits, and
  style-name token validation. A limit failure is typed and atomic; the
  implementation must not silently widen package admission to reach the hard
  owner ceiling.
* Present numeric attributes are parsed as their schema type. Non-positive
  `exponent-interval` or `max-denominator-value`, invalid booleans, malformed
  integer/double lexemes, a `min-decimal-places` greater than
  `decimal-places`, invalid transliteration format/style, duplicate names
  within one owner/family, duplicate attributes, and bad particles are typed
  invalid-format failures. A name reused across distinct owners or families is
  retained in the catalog and requires the explicit owner/family selector
  above; an omitted selector that cannot resolve uniquely is an ambiguity
  failure.
  A schema-valid embedded-text integer of zero or less is retained as an
  opaque source body for no-op preservation; typed projection and authoring
  refuse it because the prose display position is 1-based. A valid
  integer/double beyond the bounded public scalar is a limit/unsupported
  projection, never a truncation.
* Missing fields are not filled in during a read projection. Effective helper
  methods report only the ODF defaults listed above. In particular, absent
  transliteration format means ASCII digits and ignores the other
  transliteration values, while absent transliteration style reports `short`.
* XML/package parse failures, output growth, source conflicts, stale patches,
  signature/encryption policy refusal, and unsupported owned markup use the
  existing typed error categories. No failure after staging may publish a
  partially changed package.
* This batch does not add a number-formatting engine. Locale resolution,
  digit transliteration, scientific/fraction rendering, formula evaluation,
  cached-value recalculation, and display-factor application remain inert
  metadata. The typed model describes and preserves the ODF pattern; it does
  not promise a rendered string.

## Corpus and fixture evidence

The native corpus is tracked and does exercise the missing vocabulary. The
whole-file hashes and relevant XML-member hashes below were computed from the
working-tree files; the files are clean tracked entries (`git ls-files`):

| Fixture | Whole-file SHA-256 | Relevant member | Member SHA-256 | Observed vocabulary |
|---|---|---|---|---|
| `test-data/libreoffice-core/sc/qa/unit/data/ods/formats.ods` | `7bc8c21dac01ab99502e19403d8d70047cbf3ef128f9f2d21b0cd64409225634` | `styles.xml` | `8bf93a5b08d7f0dc3c8e2efdfee21142b2a2d9da996a1bb57f2cdfec704f5cc8` | two `number:fraction` and six `number:scientific-number` children |
| same | same | `content.xml` | `aa94157cefd0f16640b1b555ce4c5d11136c905732c55466603fe4a8afbb11d0` | one `number:scientific-number` child |
| `test-data/libreoffice-core/sc/qa/unit/data/functions/financial/fods/yielddisc.fods` | `b0335da945928c043c0266ee8f6446f11d81156d80f01c1decd7b45e30938933` | flat XML | whole-file hash above | three `number:scientific-number` and two `number:embedded-text` children |

`formats.ods` is a useful native packaged-ODS read fixture, but its two fraction
children also contain LibreOffice's `loext:max-numerator-digits`; that
attribute is outside the ODF 1.4 `number:fraction` grammar, which defines
`max-denominator-value` but not a core max-numerator attribute. The six
scientific children in `styles.xml` and the one in `content.xml` use the core
`number:*` attributes. `yielddisc.fods` contains core
`number:decimal-places`, `number:min-integer-digits`,
`number:min-exponent-digits`, and `number:position`, but uses
`loext:min-decimal-places`, `loext:exponent-interval`, and
`loext:forced-exponent-sign` on its scientific children. Those `loext:*`
attributes are evidence of producer extensions, not normative ODF aliases;
the implementation must preserve them and refuse a canonical semantic
replacement that would discard them. They are separate from the core
scientific/fraction/embedded-text support described here.

The current `crates/litchi-ods/tests/fixtures` directory still has no compact
fixture for these cases. Add deterministic package/XML fixtures alongside the
native corpus tests for:

1. a core `number:number-style` fraction with every core fraction attribute;
2. a core scientific style with each optional attribute and omitted-default
   cases;
3. decimal number children with zero, one, and multiple embedded-text nodes,
   including positions greater than the minimum integer digits;
4. transliteration metadata on number, date, time, currency, percentage,
   boolean, and `number:text-style` roots;
5. valid affix/fill arrangements and invalid adjacent text or multiple-fill
   arrangements; and
6. unknown attributes/foreign children, the native `loext:*` cases, and
   style-map/text-property cases to prove no-op preservation and changed-write
   refusal.

The native files establish producer evidence and read/preservation targets;
they do not establish that LibreOffice will retain every extension after a
round trip. `yielddisc.fods` is a flat XML document and cannot be passed to the
packaged ODS `document::Snapshot`/`Spreadsheet` reopen path. A flat-specific
test may use `FlatSnapshot`/`FlatSpreadsheet`; a package-owner test must use a
deterministically extracted minimal package fixture containing that flat
`content.xml`, with the extraction method and resulting member hash recorded.
Any future native save probe must record producer/version, member/file hashes,
and the observed normalization separately.

## Verification and performance plan

Unit tests should cover every field and default in the tables, canonical
serialization of decimal/scientific/fraction bodies, the shared decimal
embedded-text path in currency and percentage styles, metadata validation,
affix particle rules, source lexical preservation, unresolved/inherited
decimal defaults, and extension-graph dependency failures. Add malformed XML
cases for wrong namespaces, duplicate attrs, wrong child order, invalid
Unicode transliteration format (including a non-ASCII `Nd` digit whose numeric
value is 1), integer overflow, invalid positive values, and excess
embedded-text/body/catalog work. A non-positive `number:position` case must
prove that open/no-op preserves it as opaque while typed projection and
authoring refuse it.

Document integration tests should exercise:

* add a fraction/scientific style extension graph, assign it through an
  existing cell style, commit, reopen, and inspect the typed definition;
* replace each new body while proving unrelated content, styles, manifest,
  and opaque package members are byte-identical;
* patch transliteration attributes on every supported host, including an
  opaque `number:text-style` and opaque existing-family body, while retaining
  their body bytes;
* reject replacement of a source style containing unsupported maps/properties
  without changing the candidate;
* preserve, explicitly owner-select, and resolve data styles from both direct
  style containers in `styles.xml`, keep both owners read-only, and reject an
  owner-ambiguous name selector;
* reopen `formats.ods` through the packaged ODS path, inspect `yielddisc.fods`
  through `FlatSnapshot` or an explicitly recorded extracted-package fixture,
  classify core and `loext:*` attributes separately, and prove that
  extension-bearing nodes are preserved and refused for canonical replacement;
* preserve exact no-op bytes and inverse/source-checked patch behavior;
* keep `set_cell_style`, worksheet edits, mutable-facade edits, and atomic
  rich-cell style edits working with the new names; and
* reopen every generated fixture through the normal package validator.

Use the existing ODS XML/package fuzz and malformed-input gates with the new
fixtures. No benchmark or Cargo run is claimed by this design note. Before
implementation is accepted, measure style-catalog projection and graph
put/replace on small and large catalogs under the existing limits. The target
is lazy parsing on a style query, borrowed attribute/text slices where safe,
allocation proportional to the selected graph/catalog under the separate
source budgets, and no ZIP rebuild for an exact no-op. Replacement should scan
only the selected automatic-style owner plus the existing reference-closure
checks; it must not parse all worksheet cells merely to inspect one style.

## Remaining boundaries after this batch

The batch closes the audited number/fraction, scientific-number,
embedded-text, and transliteration-attribute gap for the owned automatic-style
path and provides metadata patching for all ODF transliteration hosts in that
owner. It does not claim the following:

* full date/time data-style vocabulary (era, weekday, quarter, textual month,
  calendar, reorder, and related elements), full currency-symbol metadata, or
  a rendering/locale service;
* semantic `number:text-style` body editing, style-map condition evaluation,
  style:text-properties editing, or arbitrary foreign/extension markup
  rewriting;
* graph authoring or metadata mutation for either read-only `styles.xml`
  owner (`CommonStyles` or `StylesAutomatic`); those definitions remain
  source-preserved and explicitly resolvable where typed projection is available;
* semantic authoring of LibreOffice `loext:*` attributes such as
  `loext:max-numerator-digits`, `loext:min-decimal-places`,
  `loext:exponent-interval`, or `loext:forced-exponent-sign`; they remain
  opaque producer markup;
* page, master-page, layout, conditional-style calculation, formulas,
  recalculation, pagination, or rendering; or
* native producer round-trip interoperability until a producer/versioned save
  probe is added, and no packaged ODS reopen claim for the flat `yielddisc.fods`
  source itself; and
* a unified package transaction over flat FODS input. Flat support remains the
  existing `FlatSnapshot`/`FlatSpreadsheet` owner or requires an explicitly
  extracted package fixture.

These boundaries keep the implementation within the existing ODS family and
ordinary transaction architecture while ensuring that every named audit token
has a typed, bounded, preservation-safe path.
