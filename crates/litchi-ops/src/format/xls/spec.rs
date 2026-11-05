//! XLS workbook vocabulary and capability table.
//!
//! The legacy BIFF owner currently exposes the workbook, ordinary worksheet
//! tabs, sparse cells, rows, columns and ranges.  The remaining shared
//! workbook vocabulary is kept in the table as explicit capability gaps so a
//! query never silently manufactures a node for a feature the native owner
//! does not project.

use crate::error::Code;
use crate::format::workbook::spec as workbook;
use crate::node::spec::{
    Access as ReadAccess, ElementSpec, FormatSpec, GapSpec, PropSpec, Subject, check_format_spec,
};
use crate::registry::{Access, Alias, Constraint, Entry, Table};
use crate::value::ValueType;
pub(crate) use workbook::{
    CELL_PROPS as SHARED_CELL_PROPS, ColumnKey, Kind, NAMED_RANGE_PROPS, NamedRangeKey, RowKey,
    SHEET_PROPS, SheetKey, WORKBOOK_PROPS as SHARED_WORKBOOK_PROPS, WorkbookKey,
};

/// The cell properties projected by the native BIFF reader: the type, the
/// rendered formula text (without its leading `=`) and its cached result, the
/// blank fact, and the cell XF's font, fill, border, alignment, protection
/// and number format.  Formatting keys report only non-default values.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum CellKey {
    Type,
    Formula,
    CachedValue,
    Empty,
    Bold,
    Italic,
    Strike,
    Underline,
    FontColor,
    FontSize,
    FontName,
    Fill,
    FillPattern,
    FillBg,
    BorderLeft,
    BorderRight,
    BorderTop,
    BorderBottom,
    WrapText,
    Horizontal,
    Vertical,
    Indent,
    ShrinkToFit,
    Locked,
    FormulaHidden,
    NumFmtId,
    NumberFormat,
}

/// Each [`CellKey`] with its shared workbook property name, in table order.
const CELL_KEYS: [(CellKey, &str); 27] = [
    (CellKey::Type, "type"),
    (CellKey::Formula, "formula"),
    (CellKey::CachedValue, "cachedValue"),
    (CellKey::Empty, "empty"),
    (CellKey::Bold, "font.bold"),
    (CellKey::Italic, "font.italic"),
    (CellKey::Strike, "strike"),
    (CellKey::Underline, "underline"),
    (CellKey::FontColor, "font.color"),
    (CellKey::FontSize, "font.size"),
    (CellKey::FontName, "font.name"),
    (CellKey::Fill, "fill"),
    (CellKey::FillPattern, "fillPattern"),
    (CellKey::FillBg, "fillBg"),
    (CellKey::BorderLeft, "border.left"),
    (CellKey::BorderRight, "border.right"),
    (CellKey::BorderTop, "border.top"),
    (CellKey::BorderBottom, "border.bottom"),
    (CellKey::WrapText, "alignment.wrapText"),
    (CellKey::Horizontal, "alignment.horizontal"),
    (CellKey::Vertical, "alignment.vertical"),
    (CellKey::Indent, "alignment.indent"),
    (CellKey::ShrinkToFit, "alignment.shrinkToFit"),
    (CellKey::Locked, "protection.locked"),
    (CellKey::FormulaHidden, "protection.hidden"),
    (CellKey::NumFmtId, "numFmtId"),
    (CellKey::NumberFormat, "numberformat"),
];

impl CellKey {
    pub(crate) fn at(index: usize) -> Option<Self> {
        CELL_KEYS.get(index).map(|(key, _)| *key)
    }
}

/// The shared workbook cell property spec named `name`.
const fn shared_cell_prop(name: &str) -> PropSpec {
    let mut index = 0;
    while index < SHARED_CELL_PROPS.len() {
        if same_name(SHARED_CELL_PROPS[index].name, name) {
            return SHARED_CELL_PROPS[index];
        }
        index += 1;
    }
    panic!("XLS cell key is not a shared workbook cell property")
}

const fn cell_props() -> [PropSpec; CELL_KEYS.len()] {
    let mut props = [PropSpec::new("", ValueType::Text); CELL_KEYS.len()];
    let mut index = 0;
    while index < CELL_KEYS.len() {
        props[index] = shared_cell_prop(CELL_KEYS[index].1);
        index += 1;
    }
    props
}

pub(crate) static CELL_PROPS: &[PropSpec] = &cell_props();

/// The scalar cell payloads admitted by the initial unified XLS writer.
///
/// `value` accepts the common scalar union and treats `null` as a clear for
/// non-formula cells. The explicit `clear` flag accepts only `true`; its
/// `null` form is deliberately refused by the registry rather than being
/// confused with a value clear. Existing formula cells are a separate
/// operation gap until formula-token and formula-cache writes are integrated.
/// Admission also depends on the native storage and dependency closure:
/// packed `MulBlank` cells and non-adjacent `ExtSST` records remain refused.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum CellSetter {
    Value,
    Clear,
    Formula,
    Horizontal,
    Vertical,
    WrapText,
    Indent,
    ShrinkToFit,
    BorderLeft,
    BorderRight,
    BorderTop,
    BorderBottom,
    FillPattern,
    Fill,
    FillBg,
}

/// One XF-backed cell formatting setter.
const fn format_entry(name: &'static str, setter: CellSetter, clear: bool) -> Entry<CellSetter> {
    Entry {
        spec: shared_cell_prop(name),
        access: Access::Set(setter),
        constraint: Constraint::None,
        clear,
    }
}

const fn write_only(name: &'static str, ty: ValueType) -> PropSpec {
    let mut spec = PropSpec::new(name, ty);
    spec.access = ReadAccess {
        get: false,
        set: false,
        add: false,
    };
    spec
}

pub(crate) const CELL_WRITE_TABLE: Table<CellSetter, 15, 1> = Table::new(
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
            spec: SHARED_CELL_PROPS[1],
            access: Access::Set(CellSetter::Formula),
            constraint: Constraint::None,
            clear: false,
        },
        format_entry("alignment.horizontal", CellSetter::Horizontal, false),
        format_entry("alignment.vertical", CellSetter::Vertical, false),
        format_entry("alignment.wrapText", CellSetter::WrapText, false),
        Entry {
            spec: shared_cell_prop("alignment.indent"),
            access: Access::Set(CellSetter::Indent),
            constraint: Constraint::Int { min: 0, max: 15 },
            clear: false,
        },
        format_entry("alignment.shrinkToFit", CellSetter::ShrinkToFit, false),
        format_entry("border.left", CellSetter::BorderLeft, true),
        format_entry("border.right", CellSetter::BorderRight, true),
        format_entry("border.top", CellSetter::BorderTop, true),
        format_entry("border.bottom", CellSetter::BorderBottom, true),
        format_entry("fillPattern", CellSetter::FillPattern, false),
        format_entry("fill", CellSetter::Fill, false),
        format_entry("fillBg", CellSetter::FillBg, false),
    ],
    // `text` names `value`, as in the XLSX cell registry (record 0269).
    [Alias {
        name: "text",
        entry: 0,
    }],
);

/// Read-schema projection of the cell write registry.  The native node keeps
/// using [`CELL_PROPS`]; this slice is for mutation admission and capability
/// introspection only.
pub(crate) const CELL_WRITE_PROPS: &[PropSpec] = &CELL_WRITE_TABLE.readable();

/// Sheet properties supported by the initial unified XLS writer.  Legacy XLS
/// can rename a worksheet while preserving its catalog identity; visibility
/// changes remain a separate capability gap.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum SheetSetter {
    Name,
}

pub(crate) const SHEET_WRITE_TABLE: Table<SheetSetter, 1, 0> = Table::new(
    [Entry {
        spec: PropSpec::new("name", ValueType::Text),
        access: Access::SetAndAdd(SheetSetter::Name),
        constraint: Constraint::None,
        clear: false,
    }],
    [],
);

pub(crate) const SHEET_WRITE_PROPS: &[PropSpec] = &SHEET_WRITE_TABLE.readable();

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

/// Capability projection for cells: readable native properties followed by
/// the write-only scalar payload keys.  Keeping this separate from
/// [`CELL_PROPS`] prevents a write-only key from becoming queryable or being
/// emitted by the read node view.
const fn cell_capability_array() -> [PropSpec; CELL_PROPS.len() + 2] {
    // `value` and `clear` are write-only; `formula` is both read and set.
    let mut props = [PropSpec::new("", ValueType::Text); CELL_PROPS.len() + 2];
    let mut index = 0;
    while index < CELL_PROPS.len() {
        let mut prop = CELL_PROPS[index];
        let mut write_index = 0;
        while write_index < CELL_WRITE_PROPS.len() {
            let write = CELL_WRITE_PROPS[write_index];
            if same_name(prop.name, write.name) {
                prop.access.set = write.access.set;
                prop.access.add = write.access.add;
                prop.admission = write.admission;
            }
            write_index += 1;
        }
        props[index] = prop;
        index += 1;
    }
    props[CELL_PROPS.len()] = CELL_WRITE_PROPS[0];
    props[CELL_PROPS.len() + 1] = CELL_WRITE_PROPS[1];
    props
}

pub(crate) static CELL_CAPABILITY_PROPS: &[PropSpec] = &cell_capability_array();

/// Capability projection for sheets: only `name` gains a set flag; `hidden`
/// and `visibility` remain read-only until the native tab-state policy lands.
const fn sheet_capability_array() -> [PropSpec; SHEET_PROPS.len()] {
    let mut props = [PropSpec::new("", ValueType::Text); SHEET_PROPS.len()];
    let mut index = 0;
    while index < SHEET_PROPS.len() {
        let mut prop = SHEET_PROPS[index];
        let mut write_index = 0;
        while write_index < SHEET_WRITE_PROPS.len() {
            let write = SHEET_WRITE_PROPS[write_index];
            if same_name(prop.name, write.name) {
                prop.access.set = write.access.set;
                prop.access.add = write.access.add;
                prop.admission = write.admission;
            }
            write_index += 1;
        }
        props[index] = prop;
        index += 1;
    }
    props
}

pub(crate) static SHEET_CAPABILITY_PROPS: &[PropSpec] = &sheet_capability_array();

/// The source-backed owner does not project the workbook window's active tab.
pub(crate) static WORKBOOK_PROPS: &[PropSpec] = &[
    SHARED_WORKBOOK_PROPS[0],
    SHARED_WORKBOOK_PROPS[1],
    SHARED_WORKBOOK_PROPS[2],
    SHARED_WORKBOOK_PROPS[3],
    SHARED_WORKBOOK_PROPS[4],
    SHARED_WORKBOOK_PROPS[5],
    SHARED_WORKBOOK_PROPS[6],
    SHARED_WORKBOOK_PROPS[7],
    SHARED_WORKBOOK_PROPS[8],
    SHARED_WORKBOOK_PROPS[9],
    SHARED_WORKBOOK_PROPS[10],
];

/// Row and column layout read through the eager reader's `Row` and `ColInfo`
/// projection (a source-backed workbook materializes it on first use).
pub(crate) use workbook::{COLUMN_PROPS, ROW_PROPS};

/// Row assignments: height (points, written as twips) and hidden.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum RowSetter {
    Height,
    Hidden,
}

pub(crate) const ROW_WRITE_TABLE: Table<RowSetter, 2, 0> = Table::new(
    [
        Entry {
            spec: ROW_PROPS[0],
            access: Access::Set(RowSetter::Height),
            constraint: Constraint::Length {
                min: crate::value::Length::from_points(0),
                max: crate::value::Length::from_points(409),
            },
            clear: false,
        },
        Entry {
            spec: ROW_PROPS[1],
            access: Access::Set(RowSetter::Hidden),
            constraint: Constraint::None,
            clear: false,
        },
    ],
    [],
);

pub(crate) const ROW_WRITE_PROPS: &[PropSpec] = &ROW_WRITE_TABLE.readable();

/// Column assignments: width (characters, written as 1/256 character) and
/// hidden.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum ColumnSetter {
    Width,
    Hidden,
}

pub(crate) const COLUMN_WRITE_TABLE: Table<ColumnSetter, 2, 0> = Table::new(
    [
        Entry {
            spec: COLUMN_PROPS[0],
            access: Access::Set(ColumnSetter::Width),
            constraint: Constraint::Number {
                min: 0.0,
                max: 255.0,
            },
            clear: false,
        },
        Entry {
            spec: COLUMN_PROPS[1],
            access: Access::Set(ColumnSetter::Hidden),
            constraint: Constraint::None,
            clear: false,
        },
    ],
    [],
);

pub(crate) const COLUMN_WRITE_PROPS: &[PropSpec] = &COLUMN_WRITE_TABLE.readable();

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
            }
            write_index += 1;
        }
        props[index] = prop;
        index += 1;
    }
    props
}

static ROW_CAPABILITY_PROPS: [PropSpec; 5] = merged(ROW_PROPS, ROW_WRITE_PROPS);
static COLUMN_CAPABILITY_PROPS: [PropSpec; 5] = merged(COLUMN_PROPS, COLUMN_WRITE_PROPS);

/// The capability projection of one element's properties, when it differs
/// from the read table.
pub(crate) fn capability_props(element: &str) -> Option<&'static [PropSpec]> {
    match element {
        "cell" => Some(CELL_CAPABILITY_PROPS),
        "sheet" => Some(SHEET_CAPABILITY_PROPS),
        "row" => Some(&ROW_CAPABILITY_PROPS),
        "col" => Some(&COLUMN_CAPABILITY_PROPS),
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

static ACTIVE_TAB_GAP: GapSpec = property_gap(
    "workbook",
    "activeTab",
    "M1-extended (R6b)",
    "the legacy XLS Package read bridge does not project the workbook window selection",
);
static EVALUATED_GAP: GapSpec = property_gap(
    "cell",
    "evaluated",
    "M1-extended (R6b)",
    "the XLS adapter has no formula evaluator or stable cached-result admission flag",
);
static ARRAY_FORMULA_GAP: GapSpec = property_gap(
    "cell",
    "arrayformula",
    "M1-extended (R6b)",
    "array-formula metadata is not projected by the Package read bridge",
);
static ARRAY_REF_GAP: GapSpec = property_gap(
    "cell",
    "arrayref",
    "M1-extended (R6b)",
    "array-formula ranges are not projected by the Package read bridge",
);

static TABLE_GAP: GapSpec = gap(
    "table",
    "M1-extended (R6b)",
    "the legacy XLS Package read bridge does not project ListObject records into unified nodes yet",
);
static CHART_GAP: GapSpec = gap(
    "chart",
    "M1-extended (R6b)",
    "legacy XLS chart metadata is not projected into unified nodes yet",
);
static SERIES_GAP: GapSpec = gap(
    "series",
    "M1-extended (R6b)",
    "legacy XLS chart series are not projected into unified nodes yet",
);
static AXIS_GAP: GapSpec = gap(
    "axis",
    "M1-extended (R6b)",
    "legacy XLS chart axes are not projected into unified nodes yet",
);
static PIVOT_GAP: GapSpec = gap(
    "pivottable",
    "M1-extended (R6b)",
    "legacy XLS pivot tables are not projected into unified nodes yet",
);
static SLICER_GAP: GapSpec = gap(
    "slicer",
    "M1-extended (R6b)",
    "legacy XLS slicers are not projected into unified nodes yet",
);
static SPARKLINE_GAP: GapSpec = gap(
    "sparkline",
    "M1-extended (R6b)",
    "legacy XLS sparkline groups are not projected into unified nodes yet",
);
static MERGE_GAP: GapSpec = gap(
    "mergeCell",
    "M1-extended (R6b)",
    "legacy XLS merged ranges are not projected into unified nodes yet",
);
static OLE_GAP: GapSpec = gap(
    "ole",
    "M1-extended (R6b)",
    "legacy XLS OLE objects are not projected into unified nodes yet",
);
static RUN_GAP: GapSpec = gap(
    "run",
    "M1-extended (R6b)",
    "legacy XLS rich-text runs are flattened by the native cell projection",
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
    with_props(&workbook::SPEC.elements[3], &[]),
    with_props(&workbook::SPEC.elements[4], ROW_PROPS),
    with_props(&workbook::SPEC.elements[5], COLUMN_PROPS),
    with_gap(&workbook::SPEC.elements[6], &TABLE_GAP),
    with_props(&workbook::SPEC.elements[7], NAMED_RANGE_PROPS),
    with_props(&workbook::SPEC.elements[8], super::features::COMMENT_PROPS),
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
    with_gap(&workbook::SPEC.elements[19], &MERGE_GAP),
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
            "WP10 (record 0339 F5): header-row table detection in queries",
            "OfficeCLI's detectedtable is a header-row block found heuristically in unstructured cells and reported by table queries; Litchi's queries return declared tables only",
        ),
        GapSpec {
            subject: Subject::Operation {
                verb: "query",
                element: "row",
                context: "header-column-predicates",
            },
            code: Code::UnsupportedCapability,
            follow_up: "M1-extended (R6b)",
            reason: "legacy XLS does not yet bind row predicate names to header columns",
        },
        GapSpec {
            subject: Subject::Operation {
                verb: "get",
                element: "cell",
                context: "non-worksheet-tab",
            },
            code: Code::UnsupportedCapability,
            follow_up: "M1-extended (R6b)",
            reason: "chart and macro tabs retain catalog positions but do not expose ordinary worksheet cells",
        },
        GapSpec {
            subject: Subject::Operation {
                verb: "view",
                element: "workbook",
                context: "non-worksheet-content",
            },
            code: Code::UnsupportedCapability,
            follow_up: "M1-extended (R6b)",
            reason: "text, outline and cell statistics cover ordinary worksheets; chart and macro payloads remain unprojected",
        },
        ACTIVE_TAB_GAP,
        EVALUATED_GAP,
        ARRAY_FORMULA_GAP,
        ARRAY_REF_GAP,
    ],
};

check_format_spec!(SPEC);

/// Elements whose properties the unified XLS writer sets (the worksheet
/// features of record 0327 included).
pub(crate) const SETTABLE: &[&str] = &[
    "cell",
    "sheet",
    "row",
    "col",
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
/// Elements the unified XLS writer adds: worksheet tabs, and empty rows and
/// columns inserted under a sheet.
pub(crate) const ADDABLE: &[&str] = &[
    "sheet",
    "row",
    "col",
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
/// Elements the unified XLS writer removes.
pub(crate) const REMOVABLE: &[&str] = &[
    "cell",
    "sheet",
    "row",
    "col",
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

const fn operation_gap(
    verb: &'static str,
    element: &'static str,
    context: &'static str,
    follow_up: &'static str,
    reason: &'static str,
) -> GapSpec {
    GapSpec {
        subject: Subject::Operation {
            verb,
            element,
            context,
        },
        code: Code::UnsupportedCapability,
        follow_up,
        reason,
    }
}

/// Write families that remain element- or property-level gaps.  The integrated
/// verbs are listed by [`SETTABLE`], [`ADDABLE`] and [`REMOVABLE`]; these rows
/// keep the remaining families visible so a format-level report never implies
/// more than the adapter implements.
pub(crate) static WRITE_GAPS: [GapSpec; 11] = [
    operation_gap(
        "get",
        "metadata",
        "path-read",
        "File::metadata(); /metadata node reads with the shared metadata vocabulary",
        "`set(\"/metadata\", ..)` writes the OLE document properties (root-metadata keys and custom.<name>); the path is set-only, and reads go through File::metadata()",
    ),
    operation_gap(
        "remove",
        "cell",
        "native-dependency-closure",
        "M2 XLS dependency-complete cell removal",
        "formula cells and unrepresented range or drawing dependencies remain refused; removing the last rich or opaque string identity requires a semantic inverse that retains its metadata",
    ),
    operation_gap(
        "set",
        "cell",
        "non-scalar-properties",
        "M2 XLS font, number-format and protection setters",
        "alignment, border line styles and palette fill colors are written through a reused or authored XF; font, border colors, number format, protection, rotation and reading order are read-only because a new XF must keep an existing font/number-format/protection prefix, and formatting is set in its own call rather than together with value, clear or formula",
    ),
    operation_gap(
        "set",
        "cell",
        "existing-formula-cells",
        "M4 XLS dependency-aware formula replacement",
        "an existing Formula record (possibly a shared, array or table owner, or a string-cache group) is not replaced by a value, clear or new formula; authored formulas use the BIFF8 tokenizer, which refuses 3-D, external and name references, and carry an empty cached result with always-recalculate set",
    ),
    operation_gap(
        "set",
        "cell",
        "native-dependency-closure",
        "M2 XLS dependency-complete scalar replacement",
        "admitted scalar storage-family replacements preserve the XF and use native row-block and SST rewrites; packed MulBlank cells, unsupported formula ownership, non-adjacent ExtSST records and removal of the last rich or opaque string identity remain typed refusals until their dependency closures and semantic inverses are represented",
    ),
    operation_gap(
        "set",
        "sheet",
        "tab-state",
        "M2 XLS worksheet visibility setters",
        "the XLS writer renames worksheets but does not change hidden or very-hidden tab state; added tabs are visible",
    ),
    operation_gap(
        "remove",
        "sheet",
        "native-dependency-closure",
        "M2 XLS dependency-closing tab removal",
        "one ordinary worksheet per call; tabs that an internal 3-D reference range, a sheet-scoped name, a pivot data connection, a drawing or a pivot view still uses are refused, as are the last tab, the last visible tab, chart/macro tabs and macro-bearing or structure-protected workbooks",
    ),
    operation_gap(
        "move",
        "sheet",
        "native-dependency-closure",
        "M2 XLS 3-D range rewriting",
        "tab order changes that would make an internal 3-D reference range noncontiguous, or that a protected or shared workbook structure forbids, are refused; cells, rows and columns do not move",
    ),
    operation_gap(
        "set",
        "row",
        "outline-and-defaults",
        "M2 XLS outline and default-layout setters",
        "row height and hidden and column width and hidden are written; outline level, collapsed state, clearing a custom size and worksheet default sizes are read-only",
    ),
    operation_gap(
        "set",
        "namedrange",
        "native-editor",
        "M2 XLS defined-name editor",
        "defined names are read from the workbook's Lbl records; adding, changing and removing them needs a native Lbl/ExternSheet editor the leaf does not have",
    ),
    operation_gap(
        "add",
        "row",
        "native-dependency-closure",
        "M2 XLS formula and name rewriting for axis shifts",
        "row and column insertion requires an explicit position and takes no properties; workbooks whose formulas, names, tables or array/shared formulas the native axis shift cannot rewrite are refused, and deletion is limited to empty rows and columns",
    ),
];

const _: () = {
    let mut index = 0;
    while index < WRITE_GAPS.len() {
        assert!(crate::node::spec::gap_valid(&WRITE_GAPS[index]));
        index += 1;
    }
};
