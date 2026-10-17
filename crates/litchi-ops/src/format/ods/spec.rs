//! ODS workbook vocabulary, property tables and capability gaps (record
//! 0277).
//!
//! The adapter shares the workbook family's element kinds and path tree
//! (`format::workbook::spec`). The leaf owns typed worksheets, named
//! ranges and expressions, cell annotations and merge spans, which this
//! table projects; the rest of the vocabulary is listed as explicit gaps
//! with their follow-ups, so a query never manufactures a node the leaf
//! does not model.

use crate::error::Code;
use crate::format::workbook::spec as workbook;
use crate::node::spec::{
    Access as ReadAccess, ElementSpec, FormatSpec, GapSpec, PropSpec, Subject, check_format_spec,
};
use crate::registry::{Access, Alias, Constraint, Entry, Table};
use crate::value::{Tokens, ValueType};
pub(crate) use workbook::{Kind, WorkbookKey};

/// The ODF value types of a cell, as the `type` key reports them.
pub(crate) static CELL_TYPE: Tokens = Tokens::new(
    "ODS cell type",
    &[
        "Number",
        "String",
        "Boolean",
        "Date",
        "Time",
        "Currency",
        "Percentage",
        "Unknown",
    ],
    &[],
);

/// The workbook root: the `meta.xml` projection the leaf reads, in the
/// shared workbook key order (`title` .. `modified`).
pub(crate) static WORKBOOK_PROPS: &[PropSpec] = &[
    workbook::WORKBOOK_PROPS[0],
    workbook::WORKBOOK_PROPS[1],
    workbook::WORKBOOK_PROPS[2],
    workbook::WORKBOOK_PROPS[3],
    workbook::WORKBOOK_PROPS[4],
    workbook::WORKBOOK_PROPS[5],
    workbook::WORKBOOK_PROPS[6],
    workbook::WORKBOOK_PROPS[7],
    workbook::WORKBOOK_PROPS[8],
    workbook::WORKBOOK_PROPS[9],
];

/// A worksheet: its `table:name`.
pub(crate) static SHEET_PROPS: &[PropSpec] = &[PropSpec::new("name", ValueType::Text)];

/// The cell keys, in [`CELL_PROPS`] order.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum CellKey {
    Type,
    Formula,
    CachedValue,
    Empty,
    Merge,
    MergeAnchor,
    Link,
}

impl CellKey {
    const ALL: [Self; 7] = [
        Self::Type,
        Self::Formula,
        Self::CachedValue,
        Self::Empty,
        Self::Merge,
        Self::MergeAnchor,
        Self::Link,
    ];

    pub(crate) fn at(index: usize) -> Option<Self> {
        Self::ALL.get(index).copied()
    }
}

/// A cell: its ODF value type, its inert OpenFormula token (stored
/// verbatim, `of:=SUM([.A1:.A3])`), the formula's cached display, the blank
/// fact, its merge span and its first hyperlink. The direct cell style name
/// is the node view's `style` member.
pub(crate) static CELL_PROPS: &[PropSpec] = &[
    PropSpec::new("type", ValueType::Token(&CELL_TYPE)),
    PropSpec::new("formula", ValueType::Text),
    PropSpec::new("cachedValue", ValueType::Text),
    PropSpec::new("empty", ValueType::Bool),
    PropSpec::new("merge", ValueType::Text),
    PropSpec::new("mergeAnchor", ValueType::Bool),
    PropSpec::new("link", ValueType::Text),
];

/// A range: whether it is exactly one merged span.
pub(crate) static RANGE_PROPS: &[PropSpec] = &[PropSpec::new("merge", ValueType::Bool)];

/// A named range or named expression, in [`NAMED_RANGE_PROPS`] order.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum NamedRangeKey {
    Name,
    Ref,
    Scope,
}

impl NamedRangeKey {
    const ALL: [Self; 3] = [Self::Name, Self::Ref, Self::Scope];

    pub(crate) fn at(index: usize) -> Option<Self> {
        Self::ALL.get(index).copied()
    }
}

/// A named range (`table:named-range`) or expression
/// (`table:named-expression`): its name, ODF reference or expression, and
/// scope (`workbook` or the owning sheet's name).
pub(crate) static NAMED_RANGE_PROPS: &[PropSpec] = &[
    PropSpec::new("name", ValueType::Text),
    PropSpec::new("ref", ValueType::Text),
    PropSpec::new("scope", ValueType::Text),
];

/// A cell annotation, in [`COMMENT_PROPS`] order.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum CommentKey {
    Ref,
    Author,
    Text,
    Date,
    AnchoredTo,
}

impl CommentKey {
    const ALL: [Self; 5] = [
        Self::Ref,
        Self::Author,
        Self::Text,
        Self::Date,
        Self::AnchoredTo,
    ];

    pub(crate) fn at(index: usize) -> Option<Self> {
        Self::ALL.get(index).copied()
    }
}

/// A cell annotation (`office:annotation`): its anchor cell, `dc:creator`,
/// plain text, `dc:date` and the path of its cell (OfficeCLI's comment keys
/// plus the date).
pub(crate) static COMMENT_PROPS: &[PropSpec] = &[
    PropSpec::new("ref", ValueType::Text),
    PropSpec::new("author", ValueType::Text),
    PropSpec::new("text", ValueType::Text),
    PropSpec::new("date", ValueType::Text),
    PropSpec::new("anchoredTo", ValueType::Text),
];

/// A merged span: its A1 range.
pub(crate) static MERGE_PROPS: &[PropSpec] = &[PropSpec::new("ref", ValueType::Text)];

/// A write-only key: admitted by `set`, never read or queried.
const fn write_only(name: &'static str, ty: ValueType) -> PropSpec {
    let mut spec = PropSpec::new(name, ty);
    spec.access = ReadAccess {
        get: false,
        set: false,
        add: false,
    };
    spec
}

/// The cell assignments the unified ODS writer stages.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum CellSetter {
    /// A scalar value; `null` clears the value and keeps the direct style.
    Value,
    /// `clear=true`: the value, formula and links go; the direct style stays.
    Clear,
    /// An inert OpenFormula token; `null` removes the formula and keeps the
    /// cached value.
    Formula,
    /// A direct cell style reference by name; `null` removes it.
    Style,
}

/// Cell assignments. `text` names `value`, as in the other workbook
/// registries (record 0269).
pub(crate) const CELL_WRITE_TABLE: Table<CellSetter, 4, 1> = Table::new(
    [
        Entry {
            spec: write_only("value", ValueType::Scalar),
            access: Access::Set(CellSetter::Value),
            constraint: Constraint::None,
            clear: true,
        },
        Entry {
            spec: write_only("clear", ValueType::Bool),
            access: Access::Set(CellSetter::Clear),
            constraint: Constraint::None,
            clear: false,
        },
        Entry {
            spec: PropSpec::new("formula", ValueType::Text),
            access: Access::Set(CellSetter::Formula),
            constraint: Constraint::NotEmpty,
            clear: true,
        },
        Entry {
            spec: write_only("style", ValueType::Text),
            access: Access::Set(CellSetter::Style),
            constraint: Constraint::NotEmpty,
            clear: true,
        },
    ],
    [Alias {
        name: "text",
        entry: 0,
    }],
);

pub(crate) const CELL_WRITE_PROPS: &[PropSpec] = &CELL_WRITE_TABLE.readable();

/// The range merge assignment.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum MergeSetter {
    Merge,
}

/// `merge=true` spans the range from its first cell; `merge=false` (or
/// `null`) removes a span anchored there.
pub(crate) const MERGE_WRITE_TABLE: Table<MergeSetter, 1, 0> = Table::new(
    [Entry {
        spec: PropSpec::new("merge", ValueType::Bool),
        access: Access::Set(MergeSetter::Merge),
        constraint: Constraint::None,
        clear: true,
    }],
    [],
);

pub(crate) const MERGE_WRITE_PROPS: &[PropSpec] = &MERGE_WRITE_TABLE.readable();

/// Worksheet assignments.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum SheetSetter {
    Name,
}

pub(crate) const SHEET_WRITE_TABLE: Table<SheetSetter, 1, 0> = Table::new(
    [Entry {
        spec: PropSpec::new("name", ValueType::Text),
        access: Access::SetAndAdd(SheetSetter::Name),
        constraint: Constraint::NotEmpty,
        clear: false,
    }],
    [],
);

pub(crate) const SHEET_WRITE_PROPS: &[PropSpec] = &SHEET_WRITE_TABLE.readable();

/// Named-range assignments; `add` takes `name` and `ref` (required) and an
/// optional sheet `scope`.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum NameSetter {
    Name,
    Ref,
    Scope,
}

pub(crate) const NAME_WRITE_TABLE: Table<NameSetter, 3, 0> = Table::new(
    [
        Entry {
            spec: PropSpec::new("name", ValueType::Text),
            access: Access::SetAndAdd(NameSetter::Name),
            constraint: Constraint::NotEmpty,
            clear: false,
        },
        Entry {
            spec: PropSpec::new("ref", ValueType::Text),
            access: Access::SetAndAdd(NameSetter::Ref),
            constraint: Constraint::NotEmpty,
            clear: false,
        },
        Entry {
            spec: PropSpec::new("scope", ValueType::Text),
            access: Access::Add(NameSetter::Scope),
            constraint: Constraint::NotEmpty,
            clear: false,
        },
    ],
    [],
);

pub(crate) const NAME_WRITE_PROPS: &[PropSpec] = &NAME_WRITE_TABLE.readable();

/// Annotation assignments; `add` under a cell or a sheet (with `ref`).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum CommentSetter {
    Ref,
    Author,
    Text,
    Date,
}

pub(crate) const COMMENT_WRITE_TABLE: Table<CommentSetter, 4, 0> = Table::new(
    [
        Entry {
            spec: PropSpec::new("ref", ValueType::Text),
            access: Access::Add(CommentSetter::Ref),
            constraint: Constraint::NotEmpty,
            clear: false,
        },
        Entry {
            spec: PropSpec::new("author", ValueType::Text),
            access: Access::SetAndAdd(CommentSetter::Author),
            constraint: Constraint::None,
            clear: true,
        },
        Entry {
            spec: PropSpec::new("text", ValueType::Text),
            access: Access::SetAndAdd(CommentSetter::Text),
            constraint: Constraint::None,
            clear: false,
        },
        Entry {
            spec: PropSpec::new("date", ValueType::Text),
            access: Access::SetAndAdd(CommentSetter::Date),
            constraint: Constraint::NotEmpty,
            clear: true,
        },
    ],
    [],
);

pub(crate) const COMMENT_WRITE_PROPS: &[PropSpec] = &COMMENT_WRITE_TABLE.readable();

const fn same_name(left: &str, right: &str) -> bool {
    let left = left.as_bytes();
    let right = right.as_bytes();
    if left.len() != right.len() {
        return false;
    }
    let mut index = 0;
    while index < left.len() {
        if left[index] != right[index] {
            return false;
        }
        index += 1;
    }
    true
}

/// A read table with the set and add flags of a write table merged in.
const fn merged<const N: usize>(read: &[PropSpec], write: &[PropSpec]) -> [PropSpec; N] {
    let mut props = [PropSpec::new("", ValueType::Text); N];
    let mut index = 0;
    while index < N {
        let mut prop = read[index];
        let mut write_index = 0;
        while write_index < write.len() {
            if same_name(prop.name, write[write_index].name) {
                prop.access.set = write[write_index].access.set;
                prop.access.add = write[write_index].access.add;
                prop.admission = write[write_index].admission;
            }
            write_index += 1;
        }
        props[index] = prop;
        index += 1;
    }
    props
}

/// The cell capability projection: readable keys with their write flags,
/// then the write-only `value`, `clear` and `style`.
const fn cell_capability_array() -> [PropSpec; CELL_PROPS.len() + 3] {
    let mut props = [PropSpec::new("", ValueType::Text); CELL_PROPS.len() + 3];
    let read: [PropSpec; CELL_PROPS.len()] = merged(CELL_PROPS, CELL_WRITE_PROPS);
    let mut index = 0;
    while index < read.len() {
        props[index] = read[index];
        index += 1;
    }
    props[CELL_PROPS.len()] = CELL_WRITE_PROPS[0];
    props[CELL_PROPS.len() + 1] = CELL_WRITE_PROPS[1];
    props[CELL_PROPS.len() + 2] = CELL_WRITE_PROPS[3];
    props
}

static CELL_CAPABILITY_PROPS: [PropSpec; CELL_PROPS.len() + 3] = cell_capability_array();

/// The range capability projection: `merge`, then the cell setters a range
/// broadcasts.
const fn range_capability_array() -> [PropSpec; RANGE_PROPS.len() + CELL_WRITE_PROPS.len()] {
    let mut props =
        [PropSpec::new("", ValueType::Text); RANGE_PROPS.len() + CELL_WRITE_PROPS.len()];
    let read: [PropSpec; RANGE_PROPS.len()] = merged(RANGE_PROPS, MERGE_WRITE_PROPS);
    props[0] = read[0];
    let mut index = 0;
    while index < CELL_WRITE_PROPS.len() {
        let mut prop = CELL_WRITE_PROPS[index];
        prop.access.get = false;
        props[RANGE_PROPS.len() + index] = prop;
        index += 1;
    }
    props
}

static RANGE_CAPABILITY_PROPS: [PropSpec; RANGE_PROPS.len() + CELL_WRITE_PROPS.len()] =
    range_capability_array();
static SHEET_CAPABILITY_PROPS: [PropSpec; 1] = merged(SHEET_PROPS, SHEET_WRITE_PROPS);
static NAME_CAPABILITY_PROPS: [PropSpec; 3] = merged(NAMED_RANGE_PROPS, NAME_WRITE_PROPS);
static COMMENT_CAPABILITY_PROPS: [PropSpec; 5] = merged(COMMENT_PROPS, COMMENT_WRITE_PROPS);

/// The capability projection of one element's properties, when it differs
/// from the read table.
pub(crate) fn capability_props(element: &str) -> Option<&'static [PropSpec]> {
    match element {
        "cell" => Some(&CELL_CAPABILITY_PROPS),
        "range" => Some(&RANGE_CAPABILITY_PROPS),
        "sheet" => Some(&SHEET_CAPABILITY_PROPS),
        "namedrange" => Some(&NAME_CAPABILITY_PROPS),
        "comment" => Some(&COMMENT_CAPABILITY_PROPS),
        _ => None,
    }
}

const fn gap(element: &'static str, follow_up: &'static str, reason: &'static str) -> GapSpec {
    GapSpec {
        subject: Subject::Element(element),
        code: Code::UnsupportedCapability,
        follow_up,
        reason,
    }
}

const fn property_gap(
    element: &'static str,
    key: &'static str,
    follow_up: &'static str,
    reason: &'static str,
) -> GapSpec {
    GapSpec {
        subject: Subject::Property { element, key },
        code: Code::UnsupportedCapability,
        follow_up,
        reason,
    }
}

const FOLLOW_UP: &str = "ODS adapter follow-up (record 0277)";

static TABLE_GAP: GapSpec = gap(
    "table",
    FOLLOW_UP,
    "ODS database ranges are read by the leaf but not projected into unified table nodes yet",
);
static CHART_GAP: GapSpec = gap(
    "chart",
    FOLLOW_UP,
    "ODS chart objects are read by the leaf but not projected into unified nodes yet",
);
static SERIES_GAP: GapSpec = gap(
    "series",
    FOLLOW_UP,
    "ODS chart series are not projected into unified nodes yet",
);
static AXIS_GAP: GapSpec = gap(
    "axis",
    FOLLOW_UP,
    "ODS chart axes are not projected into unified nodes yet",
);
static PIVOT_GAP: GapSpec = gap(
    "pivottable",
    FOLLOW_UP,
    "ODS data pilot tables are read by the leaf but not projected into unified nodes yet",
);
static SLICER_GAP: GapSpec = gap(
    "slicer",
    "none: OpenDocument has no slicer element",
    "OpenDocument spreadsheets have no slicer concept",
);
static SPARKLINE_GAP: GapSpec = gap(
    "sparkline",
    FOLLOW_UP,
    "ODS sparkline groups are read by the leaf but not projected into unified nodes yet",
);
static OLE_GAP: GapSpec = gap(
    "ole",
    FOLLOW_UP,
    "ODS embedded objects are read by the leaf but not projected into unified nodes yet",
);
static RUN_GAP: GapSpec = gap(
    "run",
    FOLLOW_UP,
    "ODS rich cell text is flattened into the cell's displayed text",
);

const fn with_props(
    element: &'static ElementSpec<Kind>,
    props: &'static [PropSpec],
) -> ElementSpec<Kind> {
    ElementSpec {
        kind: element.kind,
        name: element.name,
        aliases: element.aliases,
        parents: element.parents,
        children: element.children,
        identities: element.identities,
        paths: element.paths,
        props,
        gap: None,
    }
}

const fn with_gap(element: &'static ElementSpec<Kind>, gap: &'static GapSpec) -> ElementSpec<Kind> {
    ElementSpec {
        kind: element.kind,
        name: element.name,
        aliases: element.aliases,
        parents: element.parents,
        children: element.children,
        identities: element.identities,
        paths: element.paths,
        props: &[],
        gap: Some(gap),
    }
}

static ELEMENTS: [ElementSpec<Kind>; 26] = [
    with_props(&workbook::SPEC.elements[0], WORKBOOK_PROPS),
    with_props(&workbook::SPEC.elements[1], SHEET_PROPS),
    with_props(&workbook::SPEC.elements[2], CELL_PROPS),
    with_props(&workbook::SPEC.elements[3], RANGE_PROPS),
    with_props(&workbook::SPEC.elements[4], &[]),
    with_props(&workbook::SPEC.elements[5], &[]),
    with_gap(&workbook::SPEC.elements[6], &TABLE_GAP),
    with_props(&workbook::SPEC.elements[7], NAMED_RANGE_PROPS),
    with_props(&workbook::SPEC.elements[8], COMMENT_PROPS),
    with_gap(&workbook::SPEC.elements[9], &CHART_GAP),
    with_gap(&workbook::SPEC.elements[10], &SERIES_GAP),
    with_gap(&workbook::SPEC.elements[11], &AXIS_GAP),
    with_gap(&workbook::SPEC.elements[12], &PIVOT_GAP),
    with_gap(&workbook::SPEC.elements[13], &SLICER_GAP),
    with_gap(&workbook::SPEC.elements[14], &SPARKLINE_GAP),
    with_props(
        &workbook::SPEC.elements[15],
        super::features::VALIDATION_PROPS,
    ),
    with_props(&workbook::SPEC.elements[16], super::features::CF_PROPS),
    with_props(
        &workbook::SPEC.elements[17],
        super::features::AUTOFILTER_PROPS,
    ),
    with_props(
        &workbook::SPEC.elements[18],
        super::features::HYPERLINK_PROPS,
    ),
    with_props(&workbook::SPEC.elements[19], MERGE_PROPS),
    with_props(&workbook::SPEC.elements[20], super::features::PICTURE_PROPS),
    with_props(&workbook::SPEC.elements[21], super::features::SHAPE_PROPS),
    with_gap(&workbook::SPEC.elements[22], &OLE_GAP),
    with_props(
        &workbook::SPEC.elements[23],
        super::features::ROW_BREAK_PROPS,
    ),
    with_props(
        &workbook::SPEC.elements[24],
        super::features::COL_BREAK_PROPS,
    ),
    with_gap(&workbook::SPEC.elements[25], &RUN_GAP),
];

pub(crate) static SPEC: FormatSpec<Kind> = FormatSpec {
    root: workbook::SPEC.root,
    elements: &ELEMENTS,
    gaps: &[
        // OfficeCLI's heuristic table node (record 0339).
        gap(
            "detectedtable",
            "WP11 (record 0339 F5): header-row table detection in queries",
            "OfficeCLI's detectedtable is a header-row block found heuristically in unstructured cells and reported by table queries; Litchi's queries return declared tables only",
        ),
        GapSpec {
            subject: Subject::Operation {
                verb: "query",
                element: "row",
                context: "header-column-predicates",
            },
            code: Code::UnsupportedCapability,
            follow_up: FOLLOW_UP,
            reason: "ODS does not yet bind row predicate names to header columns",
        },
        GapSpec {
            subject: Subject::Operation {
                verb: "query",
                element: "cell",
                context: "styled-blank-cells",
            },
            code: Code::UnsupportedCapability,
            follow_up: FOLLOW_UP,
            reason: "queries enumerate content cells (a value, text, formula, link or merge span); styled blank padding runs are addressable by coordinate only",
        },
        property_gap(
            "workbook",
            "workbook.date1904",
            FOLLOW_UP,
            "the ODS null date lives in settings.xml, which is not projected yet",
        ),
        property_gap(
            "workbook",
            "activeTab",
            FOLLOW_UP,
            "the ODS active table lives in settings.xml, which is not projected yet",
        ),
        property_gap(
            "sheet",
            "hidden",
            FOLLOW_UP,
            "ODS sheet visibility lives in the table's automatic style, which is not resolved yet",
        ),
        property_gap(
            "sheet",
            "visibility",
            FOLLOW_UP,
            "ODS sheet visibility lives in the table's automatic style, which is not resolved yet",
        ),
        property_gap(
            "cell",
            "font.bold",
            FOLLOW_UP,
            "ODS cell formatting lives in named and automatic styles, which node views do not resolve yet; set `style` to apply a named style",
        ),
        property_gap(
            "cell",
            "numberformat",
            FOLLOW_UP,
            "ODS number formats are data styles referenced by the cell style, which node views do not resolve yet",
        ),
        property_gap(
            "cell",
            "computedValue",
            FOLLOW_UP,
            "ODS formulas use OpenFormula syntax; `File::evaluate` covers the Excel-syntax workbook formats only",
        ),
        property_gap(
            "row",
            "height",
            FOLLOW_UP,
            "ODS row heights live in row styles, which are not resolved yet",
        ),
        property_gap(
            "row",
            "hidden",
            FOLLOW_UP,
            "ODS row visibility is not modeled by the leaf worksheet projection",
        ),
        property_gap(
            "column",
            "width",
            FOLLOW_UP,
            "ODS column widths live in column styles, which are not resolved yet",
        ),
        property_gap(
            "column",
            "hidden",
            FOLLOW_UP,
            "ODS column visibility is not modeled by the leaf worksheet projection",
        ),
    ],
};

check_format_spec!(SPEC);

/// Elements whose properties the unified ODS writer sets (the worksheet
/// features of record 0327 included).
pub(crate) const SETTABLE: &[&str] = &[
    "cell",
    "range",
    "sheet",
    "namedrange",
    "comment",
    "hyperlink",
    "picture",
    "shape",
    "cf",
    "dataValidation",
    "autofilter",
    "rowbreak",
    "colbreak",
];
/// Elements the unified ODS writer adds.
pub(crate) const ADDABLE: &[&str] = &[
    "sheet",
    "row",
    "namedrange",
    "comment",
    "hyperlink",
    "picture",
    "shape",
    "cf",
    "dataValidation",
    "autofilter",
    "rowbreak",
    "colbreak",
];
/// Elements the unified ODS writer removes.
pub(crate) const REMOVABLE: &[&str] = &[
    "cell",
    "range",
    "sheet",
    "row",
    "namedrange",
    "comment",
    "mergeCell",
    "hyperlink",
    "picture",
    "shape",
    "cf",
    "dataValidation",
    "autofilter",
    "rowbreak",
    "colbreak",
];

const fn operation_gap(
    verb: &'static str,
    element: &'static str,
    context: &'static str,
    reason: &'static str,
) -> GapSpec {
    GapSpec {
        subject: Subject::Operation {
            verb,
            element,
            context,
        },
        code: Code::UnsupportedCapability,
        follow_up: FOLLOW_UP,
        reason,
    }
}

/// Write families that remain element- or property-level gaps. The
/// integrated verbs are listed by [`SETTABLE`], [`ADDABLE`] and
/// [`REMOVABLE`]; these rows keep the rest visible.
pub(crate) static WRITE_GAPS: [GapSpec; 12] = [
    operation_gap(
        "get",
        "metadata",
        "path-read",
        "`set(\"/metadata\", ..)` writes meta.xml; the path is set-only, and reads go through File::metadata() and the workbook root",
    ),
    operation_gap(
        "set",
        "metadata",
        "unowned-fields",
        "the ODF metadata family (record 0288) writes title, subject, description, author, lastModifiedBy, language and keywords; created, modified, revision, template, statistics and custom.<name> are refused",
    ),
    operation_gap(
        "set",
        "cell",
        "formatting-properties",
        "cell formatting is applied by naming an existing cell style (`style`); direct font, fill, border, alignment and number-format keys need an automatic-style author in the unified layer",
    ),
    operation_gap(
        "set",
        "cell",
        "formula-evaluation",
        "an authored OpenFormula is stored inert with the cell's cached value (or none); the unified layer does not evaluate ODS formulas",
    ),
    operation_gap(
        "set",
        "sheet",
        "producer-table-markup",
        "a rename copies the table fragment under its new name and drops the original; the leaf refuses sheets with formulas, ranges, drawings, forms, protection or other dependency owners, and the writer refuses names a formula or named range still uses",
    ),
    operation_gap(
        "move",
        "sheet",
        "native-dependency-closure",
        "the leaf moves only dependency-free worksheets: formulas, named/database/validation/print ranges, settings, scripts and tracked changes are refused; an added sheet placed before the end is moved the same way",
    ),
    operation_gap(
        "add",
        "row",
        "logical-row-shift",
        "rows are inserted, removed, moved and swapped through the leaf's logical row editor, which refuses formulas, merges, ranges, validations, drawings, annotations and other coordinate owners on the worksheet",
    ),
    operation_gap(
        "add",
        "col",
        "native-editor",
        "inserting or removing a column that shifts cells needs a logical column editor the leaf does not have",
    ),
    operation_gap(
        "set",
        "row",
        "layout",
        "row height and visibility, and column width and visibility, live in automatic styles the unified writer does not author yet",
    ),
    operation_gap(
        "set",
        "namedrange",
        "scope",
        "a named range's scope is fixed once added; named expressions are retargeted through `ref` and are not added",
    ),
    operation_gap(
        "set",
        "workbook",
        "flat-odf",
        "flat single-XML spreadsheets (.fods) are edited through the leaf's packaged projection and republished element by element (record 0292); an edit that adds a package member or changes a part's root has no flat form and is refused, as is a projection the packaged reader refuses",
    ),
    operation_gap(
        "batch",
        "cell",
        "coalescing",
        "consecutive direct cell value, text, clear and formula sets share one leaf edit; style, merge, range and query sets and every other verb close the run",
    ),
];

const _: () = {
    let mut index = 0;
    while index < WRITE_GAPS.len() {
        assert!(crate::node::spec::gap_valid(&WRITE_GAPS[index]));
        index += 1;
    }
};

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn capability_projections_keep_the_read_order() {
        let cell = capability_props("cell").unwrap_or_default();
        assert_eq!(cell[0].name, "type");
        assert!(
            cell.iter()
                .any(|prop| prop.name == "value" && prop.access.set)
        );
        assert!(
            cell.iter()
                .any(|prop| prop.name == "formula" && prop.access.get && prop.access.set)
        );
        let range = capability_props("range").unwrap_or_default();
        assert!(range[0].access.set && range[0].access.get);
    }
}
