# OTH body metadata fidelity design

Status: bounded design only. This document changes no production API or
implementation and does not change the OTH audit. It describes the remaining
metadata projection work after the existing body-structure slice.

## Verified boundary

The audit priority for “OTH body-structure projection” is stale as an
implementation request. Commit b97c344f0 already supplies lazy projections
for tables, annotations, tracked changes, indexes, sections, notes, frames,
and ruby. The current source has no TableProperties, ChangeTracking, or
IndexSource type, so the remaining gap is metadata fidelity inside those
already admitted roots, not root discovery or a new OTH mutation engine.

The current focused gate is:

~~~text
cargo test --locked -p litchi-oth --tests --offline
4 unit + 10 body_structures + 33 semantic_api + 9 structure_edits = 56 passed
~~~

The current crates/litchi-oth tree has no production diff. The source of truth
for the existing boundary is crates/litchi-oth/docs/FEATURE_MATRIX.md: typed
body structures are a bounded read projection, while rich or multi-block
structure bodies, table replacement/removal, formula/reference rewriting,
index generation, layout, rendering, and change accept/reject remain outside
the ordinary API.

## Normative basis and API rules

The relevant local ODF 1.4 schema is in the
[OpenDocument 1.4 archive](../../../3rdparty/specs/OpenDocument-v1.4-os.zip)
(SHA-256 `9867665f9702b365076c2c6557b23c8c938959b443f6f50712fdb2d0dfb8aac4`),
entry `schemas/OpenDocument-v1.4-schema.rng`:

- table-table-attlist, table-table-column-attlist,
  table-table-row-attlist, table-table-cell-attlist, and
  table-table-cell-attlist-extra define table metadata and spans;
- table-table-source-attlist and table-linked-source-attlist define an
  inert linked table source;
- office-annotation-attlist defines the existing annotation scalar fields
  plus drawing placement attributes;
- text-table-of-content-source-attlist,
  text-illustration-index-source-attrs,
  text-object-index-source-attrs, text-user-index-source-attr,
  text-alphabetical-index-source-attrs, and text-bibliography-source define
  index source rules;
- text-tracked-changes-attr, text-changed-region-attr, and
  office-change-info define tracking state and change identity/metadata.

The design follows ADR 0001 (typed ordinary APIs, no raw source generic), ADR
0003 (immutable snapshots, selector-first edits, source-checked no-ops), ADR
0005 (borrowed parsing, fallible reservation, hierarchical bounds and no
unsupported performance claims), ADR 0006 (preserve unknown markup and fail
closed for unsafe typed edits), ADR 0007 (data-bearing table/index/review
models without rendering), and ADR 0023 (family-owned grammar and native
evidence per family).

## Remaining public projection

The first complete implementation batch should be read-only metadata fidelity.
It must not add table metadata setters, index generation, or review-change
accept/reject. Existing plain structure text/lifecycle edits stay unchanged.

The Rust blocks below describe the owned value shape and accessor contract;
model fields stay private and are populated only by `pub(crate)` projection
constructors, as in the existing OTH models. Optional fields mean that the
schema attribute was absent. A present attribute with an invalid lexical value
is a typed projection error; it is never silently converted to an absent value.

### Tables

Keep the existing Table::name, Table::style_name, repeated physical rows and
columns, covered cells, logical widths, formula text, cached values, and
paragraph projection. Add the following owned semantic values:

~~~rust
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
#[non_exhaustive]
pub enum Visibility {
    Visible,
    Collapse,
    Filter,
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum TableSourceMode {
    CopyAll,
    CopyResultsOnly,
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum TableSourceActuate {
    OnRequest,
}

#[derive(Clone, Debug, Eq, PartialEq)]
pub struct TableProperties {
    template_name: Option<String>,
    use_first_row_styles: Option<bool>,
    use_last_row_styles: Option<bool>,
    use_first_column_styles: Option<bool>,
    use_last_column_styles: Option<bool>,
    use_banding_rows_styles: Option<bool>,
    use_banding_columns_styles: Option<bool>,
    protected: Option<bool>,
    protection_key: Option<String>,
    protection_key_digest_algorithm: Option<String>,
    print: Option<bool>,
    print_ranges: Option<String>,
    xml_id: Option<String>,
    is_sub_table: Option<bool>,
}

#[derive(Clone, Debug, Eq, PartialEq)]
pub struct TableSource {
    mode: Option<TableSourceMode>,
    table_name: Option<String>,
    href: String,
    actuate: Option<TableSourceActuate>,
    filter_name: Option<String>,
    filter_options: Option<String>,
    refresh_delay: Option<litchi_odf_common::datatype::DurationValue>,
}

#[derive(Clone, Debug, Eq, PartialEq)]
pub struct InContentMeta {
    about: String,
    property: String,
    datatype: Option<String>,
    content: Option<String>,
}

#[derive(Clone, Debug, Eq, PartialEq)]
#[non_exhaustive]
pub enum CellValue {
    Float { lexical: String },
    Percentage { lexical: String },
    Currency {
        lexical: String,
        currency: Option<String>,
    },
    Date { lexical: String },
    Time {
        value: litchi_odf_common::datatype::DurationValue,
    },
    Boolean { lexical: String, value: bool },
    String { value: Option<String> },
    Error { value: Option<String> },
}
~~~

The invariant struct constructors remain private projection constructors. The
`CellValue` enum is a public read-only value; callers inspect its public
variants rather than mutating a cached projection. Ordinary accessors are:

~~~rust
impl Table {
    pub fn properties(&self) -> &TableProperties;
    pub fn source(&self) -> Option<&TableSource>;
}
impl TableProperties {
    pub fn template_name(&self) -> Option<&str>;
    pub fn use_first_row_styles(&self) -> Option<bool>;
    pub fn use_last_row_styles(&self) -> Option<bool>;
    pub fn use_first_column_styles(&self) -> Option<bool>;
    pub fn use_last_column_styles(&self) -> Option<bool>;
    pub fn use_banding_rows_styles(&self) -> Option<bool>;
    pub fn use_banding_columns_styles(&self) -> Option<bool>;
    pub fn protected(&self) -> Option<bool>;
    pub fn protection_key(&self) -> Option<&str>;
    pub fn protection_key_digest_algorithm(&self) -> Option<&str>;
    pub fn print(&self) -> Option<bool>;
    pub fn print_ranges(&self) -> Option<&str>;
    pub fn xml_id(&self) -> Option<&str>;
    pub fn is_sub_table(&self) -> Option<bool>;
}
impl TableSource {
    pub fn mode(&self) -> Option<TableSourceMode>;
    pub fn table_name(&self) -> Option<&str>;
    pub fn href(&self) -> &str;
    pub fn actuate(&self) -> Option<TableSourceActuate>;
    pub fn filter_name(&self) -> Option<&str>;
    pub fn filter_options(&self) -> Option<&str>;
    pub fn refresh_delay(&self) -> Option<&litchi_odf_common::datatype::DurationValue>;
}
impl InContentMeta {
    pub fn about(&self) -> &str;
    pub fn property(&self) -> &str;
    pub fn datatype(&self) -> Option<&str>;
    pub fn content(&self) -> Option<&str>;
}
impl CellValue {
    pub fn lexical(&self) -> Option<&str>;
    pub fn boolean_value(&self) -> Option<bool>;
    pub fn currency(&self) -> Option<&str>;
    pub fn duration(&self) -> Option<&litchi_odf_common::datatype::DurationValue>;
    pub fn string_value(&self) -> Option<&str>;
    pub fn error_value(&self) -> Option<&str>;
}
impl Column {
    pub fn visibility(&self) -> Option<Visibility>;
    pub fn xml_id(&self) -> Option<&str>;
    pub fn declared_repeat_count(&self) -> Option<usize>;
}
impl Row {
    pub fn default_cell_style_name(&self) -> Option<&str>;
    pub fn visibility(&self) -> Option<Visibility>;
    pub fn xml_id(&self) -> Option<&str>;
    pub fn declared_repeat_count(&self) -> Option<usize>;
}
impl Cell {
    pub fn content_validation_name(&self) -> Option<&str>;
    pub fn protect(&self) -> Option<bool>;
    pub fn protected(&self) -> Option<bool>;
    pub fn xml_id(&self) -> Option<&str>;
    pub fn in_content_meta(&self) -> Option<&InContentMeta>;
    pub fn typed_value(&self) -> Option<&CellValue>;
    pub fn declared_repeat_count(&self) -> Option<usize>;
    pub fn declared_columns_spanned(&self) -> Option<usize>;
    pub fn declared_rows_spanned(&self) -> Option<usize>;
    pub fn matrix_columns_spanned(&self) -> Option<usize>;
    pub fn matrix_rows_spanned(&self) -> Option<usize>;
}
~~~

repeat_count, columns_spanned, and rows_spanned retain their current
effective-value behavior: an absent schema attribute reads as one. The new
declared_* accessors distinguish absence from an explicit value. All positive
integers are checked in 1..=MAX_REPEAT before being retained.

`CellValue` is the lossless projection of the schema's
`common-value-and-type-attlist`. Float and percentage values retain the
`office:value` `double` lexical form; currency retains that lexical form and
the optional schema-string `office:currency` without ISO-code normalization;
date retains `office:date-value` after
`dateOrDateTime` validation; and time retains the full `DurationValue`.
Boolean retains both its validated logical value and its lexical spelling.
String and error retain the optional `office:string-value` without replacing it
with the separately projected visible cell text. The existing `value_type()`
and `value()` accessors remain compatibility views; `typed_value()` is the
only API that claims complete value-family coverage. No float/date/time value
is converted to a native numeric or clock type during projection.

table:table-source is metadata only. Its xlink:type is the fixed literal
`"simple"`; required xlink:href is validated against the schema's `anyIRI`
datatype, returned as decoded lexical data, and never opened, refreshed, or
resolved.
table:refresh-delay uses the existing lossless ODF DurationValue; its lexical
form remains available through DurationValue::as_str(). The complete retained
duration value, including its component strings, is charged against the
projection budget. `table:print-ranges` is the schema's
`cellRangeAddressList`, whose RNG definition is an unconstrained `xsd:string`
with a space-separated-cell-range description; retain the exact decoded string
and do not claim a stricter range parser or reject a value merely because that
documentation-level pattern is not understood. The
`table:protection-key-digest-algorithm` value is also validated as `anyIRI`
while remaining a lexical string.

`table:table-cell` may also carry the required `xhtml:about` and
`xhtml:property` pair plus optional `xhtml:datatype` and `xhtml:content` from
`common-in-content-meta-attlist`. `InContentMeta` retains those decoded lexical
values. The scanner validates `xhtml:about` as `URIorSafeCURIE`, and
`xhtml:property`/`xhtml:datatype` as schema `CURIEs`/`CURIE`; it does not
interpret RDFa or follow a URI. `table:table-column-group/@display` and
`table:table-row-group/@display` remain source-preserved in this first batch:
the current projection intentionally flattens groups into columns and rows, so
exposing a group flag without a group boundary would be misleading. Table
templates, titles/descriptions, named expressions, and other non-attribute
children follow the same source-only policy.

### Annotations

The existing Annotation accessors already cover the ordinary non-rendering
scalar projection: name, creator, date, date_string, initials, display, and
flattened inert text. No new annotation scalar API belongs in this batch.

office:annotation also permits text:list body blocks and drawing
caption/position/size/style attributes. The list body is rich or multi-block
content, and the drawing attributes are placement/rendering data. Both are
explicitly outside the current OTH ordinary slice. They must remain in the
source authority and be preserved byte-for-byte by existing source-bound
operations. A future annotation-body projection would need an explicit
AnnotationBlock model and its own edit/readback contract; this design does not
silently flatten it into a new mutation API.

### Indexes

Keep Index::kind, name, source, and cached body. Add root common metadata and
a typed scalar source configuration. The source configuration is read-only;
templates and cached page/layout entries are not regenerated.

~~~rust
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum IndexScope {
    Document,
    Chapter,
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum CaptionSequenceFormat {
    Text,
    CategoryAndValue,
    Caption,
}

#[derive(Clone, Debug, Eq, PartialEq)]
pub enum IndexSourceOptions {
    TableOfContents {
        outline_level: Option<usize>,
        use_outline_level: Option<bool>,
        use_index_marks: Option<bool>,
        use_index_source_styles: Option<bool>,
    },
    Illustration {
        use_caption: Option<bool>,
        caption_sequence_name: Option<String>,
        caption_sequence_format: Option<CaptionSequenceFormat>,
    },
    Object {
        use_spreadsheet_objects: Option<bool>,
        use_math_objects: Option<bool>,
        use_draw_objects: Option<bool>,
        use_chart_objects: Option<bool>,
        use_other_objects: Option<bool>,
    },
    User {
        use_index_marks: Option<bool>,
        use_index_source_styles: Option<bool>,
        use_graphics: Option<bool>,
        use_tables: Option<bool>,
        use_floating_frames: Option<bool>,
        use_objects: Option<bool>,
        copy_outline_levels: Option<bool>,
        index_name: String,
    },
    Alphabetical {
        ignore_case: Option<bool>,
        main_entry_style_name: Option<String>,
        alphabetical_separators: Option<bool>,
        combine_entries: Option<bool>,
        combine_entries_with_dash: Option<bool>,
        combine_entries_with_pp: Option<bool>,
        use_keys_as_entries: Option<bool>,
        capitalize_entries: Option<bool>,
        comma_separated: Option<bool>,
        language: Option<String>,
        country: Option<String>,
        script: Option<String>,
        rfc_language_tag: Option<String>,
        sort_algorithm: Option<String>,
    },
    Bibliography,
}

#[derive(Clone, Debug, Eq, PartialEq)]
pub struct IndexSource {
    scope: Option<IndexScope>,
    relative_tab_stop_position: Option<bool>,
    options: IndexSourceOptions,
}
~~~

The table-index source uses the schema's illustration-source attributes and
therefore maps to IndexSourceOptions::Illustration. text:index-name is
required for a user-index source; a present source with no required name is a
typed projection error. Known closed tokens and booleans are rejected on
invalid lexical values rather than mapped to a guessed default.

Add these accessors while retaining the current flattened source text for
compatibility and diagnostics:

~~~rust
impl Index {
    pub fn style_name(&self) -> Option<&str>;
    pub fn protected(&self) -> Option<bool>;
    pub fn protection_key(&self) -> Option<&str>;
    pub fn protection_key_digest_algorithm(&self) -> Option<&str>;
    pub fn xml_id(&self) -> Option<&str>;
    pub fn source_options(&self) -> Option<&IndexSource>;
}
impl IndexSource {
    pub fn scope(&self) -> Option<IndexScope>;
    pub fn relative_tab_stop_position(&self) -> Option<bool>;
    pub fn options(&self) -> &IndexSourceOptions;
}
~~~

For a schema-admitted index root, `text:name` and the matching source element
are required even though the compatibility `Index::name()` and `Index::source()`
accessors remain optional. The new typed projection reports a missing required
value instead of manufacturing an empty name or source. `IndexSourceOptions`
variants expose only attributes belonging to their corresponding schema source;
templates and cached entries stay outside the value model.

is_protected() remains the effective convenience value (false when the
optional source attribute is absent); protected() reports source presence.
Index title/entry templates and bibliography configuration remain opaque source
content in this batch. That is a deliberate rendering/generation boundary, not
a promise that page numbers or entries can be recomputed.

### Change tracking

The current Change::id() merges three different ODF identities. Keep it as a
compatibility convenience, but add exact identity and change-info models:

~~~rust
#[derive(Clone, Debug, Eq, PartialEq)]
pub struct ChangeTracking {
    track_changes: Option<bool>,
}

#[derive(Clone, Debug, Eq, PartialEq)]
pub struct ChangeInfo {
    creator: String,
    date: String,
    paragraphs: Vec<Paragraph>,
}

impl ChangeTracking {
    pub fn track_changes(&self) -> Option<bool>;
}

impl ChangeInfo {
    pub fn creator(&self) -> &str;
    pub fn date(&self) -> &str;
    pub fn paragraphs(&self) -> &[Paragraph];
}

impl TextBody {
    pub fn change_tracking(&self) -> Result<Option<&ChangeTracking>>;
}

impl Change {
    pub fn xml_id(&self) -> Option<&str>;
    pub fn region_id(&self) -> Option<&str>;
    pub fn marker_change_id(&self) -> Option<&str>;
    pub fn info(&self) -> Option<&ChangeInfo>;
}
~~~

`BodyStructures` gains one optional `ChangeTracking` slot alongside its
existing change vector; `TextBody::change_tracking()` reads that declaration
without inferring state from the presence of individual marks.

text:changed-region/@xml:id is the required XML ID. Optional/deprecated
@text:id remains a separate NCName only when present and must equal the
region's xml:id; consumers give xml:id precedence. Start/end/change markers
expose @text:change-id through marker_change_id. id() may return
xml_id.or(region_id).or(marker_change_id) for source compatibility, but new
code must use the exact accessors. A text:tracked-changes declaration is
represented by Some(ChangeTracking) even when text:track-changes is absent;
absence of the declaration is None and makes otherwise inert change marks read
as the current empty change list. `office:change-info` has required
`dc:creator` and `dc:date` children and zero or more `text:p` children in ODF
1.4. `ChangeInfo` therefore stores creator/date as string/date-time lexical
values, respectively, and exposes the repeated paragraphs; `date` is validated
with the common ODF date-time grammar but its lexical spelling is retained
without canonicalization;
missing required children are a typed projection error. It has no
`meta:date-string` or creator-initials fields: those belong to annotations, not
change-info. `Change::info()` is `None` for standalone start/end markers and
`Some` for a changed region after its required change-info validates. ChangeInfo
never invokes a clock, actor, resolver, or accept/reject operation.

The change scanner validates the direct schema shape before projection. A
`text:changed-region` must have exactly one direct `text:insertion`,
`text:deletion`, or `text:format-change`; that child must have exactly one
direct `office:change-info`, whose ordered children are one `dc:creator`, one
`dc:date`, then zero or more `text:p` elements. Deletion content follows the
change-info child; insertion and format-change do not acquire an invented
payload. Standalone `text:change-start`, `text:change-end`, and `text:change`
markers expose only their `text:change-id`. The marker value is validated as
an XML `IDREF` and must resolve to the required `xml:id` of a projected
changed region. An optional `text:id` is never required for marker resolution.

## Scanner, ownership, and retained budget

Template::text_body() continues to hold a lifetime-free Snapshot handle; the
public models remain owned values cached behind the snapshot's Arc. The scanner
in crates/litchi-oth/src/codec/structure.rs may borrow content.xml while
reading quick_xml events, but it must not expose 'a source generics in ordinary
APIs or detach a raw XML fragment from its source provenance.

The same borrowed pass must apply the schema lexical rules to every admitted
metadata field: `anyIRI` for linked-source hrefs and digest algorithm values;
`URIorSafeCURIE` for `xhtml:about`; `CURIEs` for `xhtml:property`; `CURIE` for
`xhtml:datatype`; XML `ID` for every admitted `xml:id`; optional/deprecated
changed-region `text:id` as an `NCName` that must equal its `xml:id` when
present; and XML `IDREF` for marker `text:change-id`. Keep a bounded registry
of all `xml:id` values encountered in the content part and reject a duplicate
before publishing any projection. Optional `text:id` values must not create a
second identity and are never required. Unresolved marker references are a
typed error, not an inert string fallback.

The implementation should extend the existing Node projection rather than add
a second DOM or a second source copy:

1. Resolve the existing namespace-qualified attributes while the event is
   borrowed.
2. Precharge each raw Start/Empty event's qualified-name, namespace, and
   attribute/value bytes against the temporary-node budget before calling
   quick_xml decoded-and-normalized-value helpers or constructing decoded
   strings. Only then decode and validate a scalar into a temporary value.
3. Precharge the new model's strings and vector slots before cloning them into
   Table, Index, or Change values. The existing BodyStructures::account total
   must include every new retained string, including nested source metadata,
   `InContentMeta`, change-info paragraphs, and all retained DurationValue
   lexical/component bytes.
4. Use try_reserve/try_reserve_exact for metadata vectors. Repeated rows,
   columns, and cells remain compact counts; no repetition expansion is
   permitted.
5. Parse text:track-changes in the existing bounded declaration scan. A second
   whole-document scan is acceptable for the first implementation, but it must
   remain within the existing input ceiling and must not retain a second XML
   copy.

The scanner has two separate 16 MiB accounting domains. The temporary-node
budget covers decoded element/attribute names, namespace URIs, attribute
values, and `Node` child/vector slots while the owned scan tree is alive. The
retained-model budget covers the strings, `DurationValue` components,
paragraphs, and vector slots kept in the cached `BodyStructures`. A temporary
copy being dropped does not refund or consume the retained-model budget, and a
model value is never constructed before its checked retained charge and vector
reservation succeed. `measure_*` helpers should perform those checked sums
before `project_*` calls `to_owned`; publication of a partial model is
forbidden on either cap.

The inherited safety ceilings remain:

| Resource | Ceiling | Required behavior |
| --- | ---: | --- |
| content.xml bytes | 64 MiB default/hard compact XML ceiling | refuse before projection |
| XML depth | 256 default structure depth | refuse before deeper state growth |
| projected structure nodes | 1,000,000 | checked increment before node ownership |
| temporary decoded Node text/metadata | 16 MiB aggregate | precharge before temporary String/Vec construction |
| retained structure text/metadata | 16 MiB aggregate | precharge before cached-model clone |
| repeated rows/columns/cells and checked dimensions | 1,000,000 | preserve compact count; refuse overflow |

The OTH facade currently does not expose from_bytes_with_limits or
open_with_limits. This batch must not claim caller-selectable semantic limits.
Threading SourcePackageLimits and a caller-owned structure budget through OTH
opening is a separate API task; until then, the fixed ceilings above are the
public safety contract. The implementation must nevertheless keep the seam
explicit: the projection entry point should accept an internal borrowed budget
whose fields are the minimum of any future caller request and the hard ceilings
above. Each `account`, node, depth, repetition, and `try_reserve` operation
must consult that budget before constructing the owned value. A future public
`Template::from_bytes_with_limits(bytes, limits)` may thread the same budget
through package opening, but this metadata batch does not invent that facade
API or expose a raw source lifetime. Codec-level tests may use a tight internal
budget to prove one-under/one-over refusal; public tests must rely on the fixed
contract until the facade limit type exists.

## Preservation and no-op policy

Metadata projection never rewrites content.xml. Original XML remains the
authority for namespace spelling, attribute order, lexical values, comments,
CDATA, processing instructions, foreign descendants, index templates, and
annotation list/drawing content. The typed projection owns only bounded,
checked value copies and existing paragraph projections; it never synthesizes
replacement XML from those values.

The first batch adds no metadata setters. An exact Template::edit() no-op must
continue to share the original snapshot/source bytes. Existing text-only
section, note, annotation, lifecycle, and one-cell-table edits keep their
source-range preconditions, preserve unselected bytes, reopen the candidate,
and produce an exact inverse. If a selected structure contains unsupported rich
content, metadata access remains readable but a future edit must return a
typed refusal; it must never flatten or silently drop the content.

## Selectors and mutation boundary

Read access remains source order:

~~~rust
let tables = template.text_body()?.tables()?;
let indexes = template.text_body()?.indexes()?;
let changes = template.text_body()?.changes()?;
~~~

No Index or Change selector is added because index generation and
change accept/reject are outside this batch. Existing StructureSelector::Table
and StructureSelector::Annotation continue to address current bounded text
edits by checked source-order Position; optional names and IDs are metadata,
not globally unique mutation selectors. A later metadata edit would require a
new source span/read-set design before adding a setter.

## Regression matrix

The implementation batch should add
crates/litchi-oth/tests/body_structure_metadata.rs and extend the existing
56-test gate only where the assertion is about a new field.

| Area | Schema field/profile | Expected public result | Regression |
| --- | --- | --- | --- |
| Table | table:template-name, six style flags | Table::properties() preserves Option<String>/Option<bool> | all absent, explicit false, explicit true |
| Table | protection key, digest IRI, table:protected | optional lexical values and presence | entity-escaped IRI, invalid boolean refusal |
| Table | table:print, table:print-ranges, xml:id, is-sub-table | optional typed scalars | `cellRangeAddressList` remains an exact RNG string; escaped and arbitrary lexical strings survive; duplicate XML IDs refuse |
| Table | column/row visibility | Visibility enum | visible/collapse/filter plus invalid token refusal |
| Table | normal and matrix spans | declared normal/matrix spans | absent default one versus explicit values; checked dimensions |
| Table | repeat declarations | declared optional count plus current effective count | absent versus explicit one; zero/overflow refusal |
| Table | cell validation/protect/protected/IDs/spans | optional accessors | both protection flags remain distinct; absent spans remain distinguishable from one |
| Table | office:value-type plus float/percentage/currency/date/time/boolean/string/error values | CellValue | every required companion attr is validated and lexical values are retained without numeric/date coercion |
| Table | xhtml:about/property/datatype/content | InContentMeta | URIorSafeCURIE/CURIEs/CURIE lexical negatives refuse; no URI/RDF interpretation |
| Table | table:table-source and linked-source attrs | inert TableSource | no external fetch; malformed xlink pair refuses typed projection |
| Table | column/row-group display and table child templates | source-only preservation | flattened groups and cached/template children survive exact no-op |
| Annotation | office/DC/meta scalar fields | existing accessors unchanged | all fields, escaped values, explicit false |
| Annotation | text:list and draw placement | source-preserved but no new typed mutation | read/no-op keeps unknown list and placement bytes unchanged |
| Index | common section metadata | style, name, optional protection/ID accessors | absent versus explicit false; digest lexical value |
| Index | TOC source attrs | TableOfContents options | outline level, scope, booleans, invalid token refusal |
| Index | illustration/table source attrs | Illustration options | caption name/format and closed format enum |
| Index | object source attrs | Object options | all five object booleans and absent state |
| Index | user source attrs | User options | required index-name, optional flags, missing-name refusal |
| Index | alphabetical source attrs | Alphabetical options | language/script/tag/sort lexical values and booleans |
| Index | templates/cached body | existing source/body plus raw preservation | no generation, no page/layout rewrite, byte-exact no-op |
| Changes | tracked declaration | change_tracking() | absent declaration, present without flag, explicit true/false |
| Changes | changed-region IDs | separate xml_id and optional region_id | required XML ID; optional/deprecated NCName text:id must equal xml_id; duplicate XML ID refusal and xml_id precedence |
| Changes | marker target | marker_change_id | IDREF lexical validation and resolution to required region xml:id; optional text:id is never required; start/end/change markers remain separate kinds |
| Changes | office:change-info | ChangeInfo | direct insertion/deletion/format-change shape; creator/date then repeated `text:p`; required-child and ordering refusal |
| Changes | malformed known fields | typed error, no partial model publication | invalid booleans, invalid ID/NCName/IDREF/anyIRI/CURIE values, duplicate IDs |
| Ownership | namespace aliases/foreign wrappers | only admitted office:body/office:text roots project | same-named foreign roots remain inert |
| Limits | temporary-node versus retained-model caps | typed limit error | exact and one-under/one-over 16 MiB checks before decoded String/Vec construction and cached-model clone |
| Preservation | unknown attr/child/comment/CDATA/PI | source bytes unchanged | no-op and existing inverse remain exact |

The test document should include every admitted index family in one source and at
least one nested/grouped table. It should assert both TextBody readback and
Template::content_xml() byte preservation. It must not use a raw source range or
package ID as an ordinary selector.

## Native evidence boundary

The checked-in OTH fixture
crates/litchi-oth/tests/fixtures/libreoffice-desktop-html.oth is 4,640 bytes
(SHA-256
a3a880bde11afb96cfe35027c57cf7e7cf81c58d9cad617519ad2a4fcd852a33). Its
content.xml contains sequence declarations and an empty paragraph but no
table, annotation, index, or tracked-change root. It can verify the existing
native package envelope and the absence case only.

The following local ODF fixtures are useful schema/oracle material but are ODT
packages and must not be reported as native OTH evidence:

| Fixture | SHA-256 | Use |
| --- | --- | --- |
| test-data/odf/corpus/writer-table.odt | 267768eee0718f2df8faa26ae6e817d381e02a7b5aabbbc00b770058b266b15d | table metadata and grouped rows |
| test-data/odf/corpus/writer-table-of-contents.odt | 93d34dfc06e18b1d44b55503317a6eb2c600185896428f7d42b57e8dfe6ca457 | TOC source/template oracle |
| test-data/odf/odt/table-cell-column-span.odt | bd6bc6d4ce5ef0f11faa4fe188a82c4ec6a884628c8d9e521c845a623f51d96 | covered and spanned-cell oracle |

No checked-in native OTH structure fixture currently exists. A future native
gate may add a LibreOffice Writer/Web-generated .oth only after confirming that
the producer actually emits this MIME family and recording the package hash
plus exact content.xml roots. Until then, metadata tests are normative
synthetic XML tests and native evidence remains limited to the existing
empty-structure fixture.

## Coding split and acceptance gate

1. src/model/table.rs: add table properties, source, visibility, declared
   presence, `InContentMeta`, `CellValue`, and cell/row/column metadata
   accessors.
2. src/model/index.rs: add common index metadata and typed scalar IndexSource
   options; keep templates and cached bodies source-backed.
3. src/model/change.rs: add tracking declaration, exact IDs, direct
   kind-specific ChangeInfo, and compatibility-preserving accessors.
4. src/codec/structure.rs: decode schema-qualified attributes, validate
   RDFa/IRI/ID lexical forms and identity closure, enforce closed
   tokens/positive integers, separate temporary/retained budgets, precharge
   retained values, and populate cached BodyStructures without a new DOM or
   public source lifetime.
5. src/package/snapshot.rs, src/facade/mod.rs, and src/lib.rs: cache and
   re-export the new read accessors only.
6. Add the regression matrix tests, run the existing 56-test gate, strict
   Clippy/rustdoc/format checks, and record native evidence limits. Update
   FEATURE_MATRIX.md only after code and tests land; do not revive the stale
   audit priority as a new root-projection task.

Completion means every field in the table/index/change metadata slice has a
typed read accessor, optional presence is testable, unknown source bytes remain
unchanged, limit failures are atomic, and existing source-bound edits keep
their exact no-op and inverse behavior. It does not mean OTH rendering,
pagination, external refresh, index generation, rich annotation authoring, or
change acceptance/rejection.
