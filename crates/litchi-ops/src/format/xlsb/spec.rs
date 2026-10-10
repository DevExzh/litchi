//! The XLSB specialization of the shared workbook capability vocabulary.
//!
//! XLSB and XLSX use the same semantic element and property keys.  Their
//! leaves do not have the same deferred parts, however, so the capability
//! report keeps an XLSB-local table with the shared elements and properties
//! and replaces the XLSX-specific gap explanations.  Keeping the table local
//! also gives the XLSB resolver a place to add a binary-only gap without
//! changing the XLSX report.

use crate::error::Code;
use crate::format::workbook::spec as workbook;
use crate::node::spec::{ElementSpec, FormatSpec, GapSpec, PropSpec, Subject, check_format_spec};
use crate::registry::{Access, Constraint, Entry, Table};
use crate::value::ValueType;

pub(crate) use workbook::{
    COLUMN_PROPS, ColumnKey, Kind, ROW_PROPS, RowKey, SheetKey, TABLE_PROPS, TableKey,
    WORKBOOK_PROPS, WorkbookKey,
};

/// The XLSB sheet setter vocabulary. The first unified XLSB write exposes
/// only a worksheet name; hidden/visibility and every other sheet property
/// remain read-only or capability gaps.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum SheetSetter {
    Name,
}

pub(crate) const SHEET_WRITE_TABLE: Table<SheetSetter, 1, 0> = Table::new(
    [Entry {
        spec: PropSpec::new("name", ValueType::Text),
        access: Access::Set(SheetSetter::Name),
        constraint: Constraint::None,
        clear: false,
    }],
    [],
);

/// Read-schema projection of the XLSB sheet write registry.
pub(crate) const SHEET_WRITE_PROPS: &[PropSpec] = &SHEET_WRITE_TABLE.readable();

/// XLSB's cell payload setter vocabulary.
///
/// The binary cell owner accepts the common scalar union and an explicit
/// `clear: true`.  A null `value` is admitted through the registry's clear
/// path and has the same blank-cell meaning.  Formula authoring remains
/// deliberately absent from this table: `formula` is readable in the
/// capability schema, but its setter is still a format-local gap.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum CellSetter {
    Value,
    Clear,
}

pub(crate) const CELL_WRITE_TABLE: Table<CellSetter, 2, 0> = Table::new(
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
    ],
    [],
);

const fn write_only(name: &'static str, ty: ValueType) -> PropSpec {
    let mut spec = PropSpec::new(name, ty);
    spec.access.get = false;
    spec
}

pub(crate) const CELL_WRITE_PROPS: &[PropSpec] = &CELL_WRITE_TABLE.readable();

/// Capability projection for cell writes.  The read projection remains the
/// local `CELL_PROPS` slice; the two write-only payload entries are appended
/// so the capability report exposes their registry types and set flags.
const fn cell_capability_array() -> [PropSpec; CELL_PROPS.len() + 2] {
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

const fn same_name(left: &str, right: &str) -> bool {
    let left = left.as_bytes();
    let right = right.as_bytes();
    if left.len() != right.len() {
        return false;
    }
    let mut index = 0;
    while index < left.len() {
        let mut a = left[index];
        let mut b = right[index];
        if a.is_ascii_uppercase() {
            a = a.to_ascii_lowercase();
        }
        if b.is_ascii_uppercase() {
            b = b.to_ascii_lowercase();
        }
        if a != b {
            return false;
        }
        index += 1;
    }
    true
}

const fn sheet_props_with_writes() -> [PropSpec; workbook::SHEET_PROPS.len()] {
    let mut props = [PropSpec::new("", ValueType::Text); workbook::SHEET_PROPS.len()];
    let mut index = 0;
    while index < workbook::SHEET_PROPS.len() {
        let mut prop = workbook::SHEET_PROPS[index];
        let mut write_index = 0;
        while write_index < SHEET_WRITE_PROPS.len() {
            let write = SHEET_WRITE_PROPS[write_index];
            if same_name(prop.name, write.name) {
                prop.access.set = write.access.set;
                prop.access.add = write.access.add;
            }
            write_index += 1;
        }
        props[index] = prop;
        index += 1;
    }
    props
}

/// XLSB's local sheet table, with only `name` marked settable.
pub(crate) static SHEET_PROPS: &[PropSpec] = &sheet_props_with_writes();

/// XLSB's node exposes only this subset of the shared workbook cell
/// vocabulary.  Keeping the local key order identical to the filtered table
/// makes `Props::get(index)` safe after unsupported properties are removed.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum CellKey {
    Type,
    Formula,
    CachedValue,
    Evaluated,
    ArrayFormula,
    ArrayRef,
    Empty,
    WrapText,
    Horizontal,
    Vertical,
    TextRotation,
    Indent,
    ShrinkToFit,
    ReadingOrder,
    Merge,
    MergeAnchor,
}

impl CellKey {
    const ALL: &'static [Self] = &[
        Self::Type,
        Self::Formula,
        Self::CachedValue,
        Self::Evaluated,
        Self::ArrayFormula,
        Self::ArrayRef,
        Self::Empty,
        Self::WrapText,
        Self::Horizontal,
        Self::Vertical,
        Self::TextRotation,
        Self::Indent,
        Self::ShrinkToFit,
        Self::ReadingOrder,
        Self::Merge,
        Self::MergeAnchor,
    ];

    pub(crate) fn at(index: usize) -> Option<Self> {
        Self::ALL.get(index).copied()
    }
}

pub(crate) static CELL_PROPS: &[PropSpec] = &[
    workbook::CELL_PROPS[0],
    workbook::CELL_PROPS[1],
    workbook::CELL_PROPS[2],
    workbook::CELL_PROPS[3],
    workbook::CELL_PROPS[4],
    workbook::CELL_PROPS[5],
    workbook::CELL_PROPS[6],
    workbook::CELL_PROPS[31],
    workbook::CELL_PROPS[32],
    workbook::CELL_PROPS[33],
    workbook::CELL_PROPS[34],
    workbook::CELL_PROPS[35],
    workbook::CELL_PROPS[36],
    workbook::CELL_PROPS[37],
    workbook::CELL_PROPS[43],
    workbook::CELL_PROPS[44],
];

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum NamedRangeKey {
    Name,
}

impl NamedRangeKey {
    const ALL: &'static [Self] = &[Self::Name];

    pub(crate) fn at(index: usize) -> Option<Self> {
        Self::ALL.get(index).copied()
    }
}

pub(crate) static NAMED_RANGE_PROPS: &[PropSpec] = &[workbook::NAMED_RANGE_PROPS[0]];

const fn element_gap(
    element: &'static str,
    follow_up: &'static str,
    reason: &'static str,
) -> GapSpec {
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

// The binary leaf already retains these parts, but the first XLSB adapter
// does not project them into ordinary nodes.  Keep the reasons format-local;
// in particular, do not inherit the XLSX follow-ups from the shared table.
static COMMENT_GAP: GapSpec = element_gap(
    "comment",
    "M1-extended (R4x)",
    "the XLSB worksheet keeps comments in the leaf part, but no unified comment node reads them yet",
);
static CHART_GAP: GapSpec = element_gap(
    "chart",
    "M1-extended (R4x)",
    "the XLSB workbook keeps drawing and chart parts, but the unified node has no chart projection yet",
);
static SERIES_GAP: GapSpec = element_gap(
    "series",
    "M1-extended (R4x)",
    "XLSB chart series are not projected into unified nodes yet",
);
static AXIS_GAP: GapSpec = element_gap(
    "axis",
    "M1-extended (R4x)",
    "XLSB chart axes are not projected into unified nodes yet",
);
static PIVOT_GAP: GapSpec = element_gap(
    "pivottable",
    "M1-extended (R4x)",
    "XLSB pivot tables are retained by the leaf but not projected into unified nodes yet",
);
static SLICER_GAP: GapSpec = element_gap(
    "slicer",
    "M1-extended (R4x)",
    "XLSB slicers are retained by the leaf but not projected into unified nodes yet",
);
static SPARKLINE_GAP: GapSpec = element_gap(
    "sparkline",
    "M1-extended (R4x)",
    "XLSB sparkline groups are retained by the leaf but not projected into unified nodes yet",
);
static VALIDATION_GAP: GapSpec = element_gap(
    "dataValidation",
    "M1-extended (R4x)",
    "XLSB data validations are retained by the leaf but not projected into unified nodes yet",
);
static CF_GAP: GapSpec = element_gap(
    "conditionalFormatting",
    "M1-extended (R4x)",
    "XLSB conditional formats are retained by the leaf but not projected into unified nodes yet",
);
static AUTOFILTER_GAP: GapSpec = element_gap(
    "autofilter",
    "M1-extended (R4x)",
    "XLSB auto-filters are retained by the leaf but not projected into unified nodes yet",
);
static HYPERLINK_GAP: GapSpec = element_gap(
    "hyperlink",
    "M1-extended (R4x)",
    "XLSB hyperlinks are retained by the leaf but their targets are not projected into nodes yet",
);
static MERGE_GAP: GapSpec = element_gap(
    "mergeCell",
    "M1-extended (R4x)",
    "XLSB merged ranges are retained by the leaf but are not listed as nodes yet",
);
static PICTURE_GAP: GapSpec = element_gap(
    "picture",
    "M1-extended (R4x)",
    "XLSB drawings are retained by the leaf but are not projected into nodes yet",
);
static SHAPE_GAP: GapSpec = element_gap(
    "shape",
    "M1-extended (R4x)",
    "XLSB drawing shapes are retained by the leaf but are not projected into nodes yet",
);
static OLE_GAP: GapSpec = element_gap(
    "ole",
    "M1-extended (R4x)",
    "XLSB OLE objects are retained by the leaf but are not projected into nodes yet",
);
static ROW_BREAK_GAP: GapSpec = element_gap(
    "rowbreak",
    "M1-extended (R4x)",
    "XLSB row page breaks are not projected into unified nodes yet",
);
static COL_BREAK_GAP: GapSpec = element_gap(
    "colbreak",
    "M1-extended (R4x)",
    "XLSB column page breaks are not projected into unified nodes yet",
);
static RUN_GAP: GapSpec = element_gap(
    "run",
    "M1-extended (R4x)",
    "the XLSB cell model exposes a cell value, not a unified rich-text run node",
);

const fn with_props(
    element: &'static ElementSpec<Kind>,
    props: &'static [PropSpec],
    gap: Option<&'static GapSpec>,
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
        gap,
    }
}

const fn with_gap(
    element: &'static ElementSpec<Kind>,
    gap: Option<&'static GapSpec>,
) -> ElementSpec<Kind> {
    with_props(element, element.props, gap)
}

// `Kind` is shared with the workbook resolver and its discriminants are the
// indices of the common table.  Rebuild the array in that order so the XLSB
// resolver can use the local `SPEC` without duplicating the vocabulary.
static ELEMENTS: [ElementSpec<Kind>; 26] = [
    with_gap(&workbook::SPEC.elements[0], None),
    with_props(&workbook::SPEC.elements[1], SHEET_PROPS, None),
    with_props(&workbook::SPEC.elements[2], CELL_CAPABILITY_PROPS, None),
    with_props(&workbook::SPEC.elements[3], CELL_WRITE_PROPS, None),
    with_gap(&workbook::SPEC.elements[4], None),
    with_gap(&workbook::SPEC.elements[5], None),
    with_gap(&workbook::SPEC.elements[6], None),
    with_props(&workbook::SPEC.elements[7], NAMED_RANGE_PROPS, None),
    with_gap(&workbook::SPEC.elements[8], Some(&COMMENT_GAP)),
    with_gap(&workbook::SPEC.elements[9], Some(&CHART_GAP)),
    with_gap(&workbook::SPEC.elements[10], Some(&SERIES_GAP)),
    with_gap(&workbook::SPEC.elements[11], Some(&AXIS_GAP)),
    with_gap(&workbook::SPEC.elements[12], Some(&PIVOT_GAP)),
    with_gap(&workbook::SPEC.elements[13], Some(&SLICER_GAP)),
    with_gap(&workbook::SPEC.elements[14], Some(&SPARKLINE_GAP)),
    with_gap(&workbook::SPEC.elements[15], Some(&VALIDATION_GAP)),
    with_gap(&workbook::SPEC.elements[16], Some(&CF_GAP)),
    with_gap(&workbook::SPEC.elements[17], Some(&AUTOFILTER_GAP)),
    with_gap(&workbook::SPEC.elements[18], Some(&HYPERLINK_GAP)),
    with_gap(&workbook::SPEC.elements[19], Some(&MERGE_GAP)),
    with_gap(&workbook::SPEC.elements[20], Some(&PICTURE_GAP)),
    with_gap(&workbook::SPEC.elements[21], Some(&SHAPE_GAP)),
    with_gap(&workbook::SPEC.elements[22], Some(&OLE_GAP)),
    with_gap(&workbook::SPEC.elements[23], Some(&ROW_BREAK_GAP)),
    with_gap(&workbook::SPEC.elements[24], Some(&COL_BREAK_GAP)),
    with_gap(&workbook::SPEC.elements[25], Some(&RUN_GAP)),
];

// These are the workbook properties whose shared baseline deliberately does
// not claim a read.  The values are the same gaps as XLSX's table, but the
// reasons describe the XLSB leaf and its deferred binary parts.
static GAPS: &[GapSpec] = &[
    property_gap(
        "workbook",
        "calc.mode",
        "M1-extended (R4x)",
        "XLSB calculation properties are retained by the leaf but not projected into the workbook node",
    ),
    property_gap(
        "workbook",
        "firstSheet",
        "M1-extended (R4x)",
        "the XLSB workbook node does not expose the authored first-sheet setting yet",
    ),
    property_gap(
        "sheet",
        "freeze",
        "M1-extended (R4x)",
        "XLSB sheet views are retained by the leaf but freeze panes are not projected yet",
    ),
    property_gap(
        "sheet",
        "zoom",
        "M1-extended (R4x)",
        "XLSB sheet views are retained by the leaf but zoom is not projected yet",
    ),
    property_gap(
        "sheet",
        "autoFilter",
        "M1-extended (R4x)",
        "XLSB auto-filter metadata is retained by the leaf but not projected into a sheet node yet",
    ),
    property_gap(
        "sheet",
        "tabColor",
        "M1-extended (R4x)",
        "XLSB sheet tab color is not projected into the sheet node yet",
    ),
    property_gap(
        "sheet",
        "protect",
        "M1-extended (R4x)",
        "XLSB sheet protection metadata is retained by the leaf but not projected yet",
    ),
    property_gap(
        "sheet",
        "printArea",
        "M1-extended (R4x)",
        "XLSB print settings are not projected into the sheet node yet",
    ),
    property_gap(
        "cell",
        "computedValue",
        "M4 (formula evaluation)",
        "the full XLSB grid exposes formula text and cached results; the source cached-value fallback has neither formula text nor formatting metadata, and there is no evaluator for an uncached result",
    ),
    property_gap(
        "cell",
        "font.bold",
        "M1-extended (R4d)",
        "the XLSB node does not project font records into cell properties",
    ),
    property_gap(
        "cell",
        "font.italic",
        "M1-extended (R4d)",
        "the XLSB node does not project font records into cell properties",
    ),
    property_gap(
        "cell",
        "strike",
        "M1-extended (R4d)",
        "the XLSB node does not project font records into cell properties",
    ),
    property_gap(
        "cell",
        "underline",
        "M1-extended (R4d)",
        "the XLSB node does not project font records into cell properties",
    ),
    property_gap(
        "cell",
        "font.color",
        "M1-extended (R4d)",
        "the XLSB node does not project font records into cell properties",
    ),
    property_gap(
        "cell",
        "superscript",
        "M1-extended (R4d)",
        "the XLSB node does not project font records into cell properties",
    ),
    property_gap(
        "cell",
        "subscript",
        "M1-extended (R4d)",
        "the XLSB node does not project font records into cell properties",
    ),
    property_gap(
        "cell",
        "font.size",
        "M1-extended (R4d)",
        "the XLSB node does not project font records into cell properties",
    ),
    property_gap(
        "cell",
        "font.name",
        "M1-extended (R4d)",
        "the XLSB node does not project font records into cell properties",
    ),
    property_gap(
        "cell",
        "font.outline",
        "M1-extended (R4d)",
        "the XLSB node does not project font records into cell properties",
    ),
    property_gap(
        "cell",
        "font.shadow",
        "M1-extended (R4d)",
        "the XLSB node does not project font records into cell properties",
    ),
    property_gap(
        "cell",
        "font.condense",
        "M1-extended (R4d)",
        "the XLSB node does not project font records into cell properties",
    ),
    property_gap(
        "cell",
        "font.extend",
        "M1-extended (R4d)",
        "the XLSB node does not project font records into cell properties",
    ),
    property_gap(
        "cell",
        "font.charset",
        "M1-extended (R4d)",
        "the XLSB node does not project font records into cell properties",
    ),
    property_gap(
        "cell",
        "font.family",
        "M1-extended (R4d)",
        "the XLSB node does not project font records into cell properties",
    ),
    property_gap(
        "cell",
        "alignment.relativeIndent",
        "M1-extended (R4d)",
        "the XLSB node does not project this alignment attribute into cell properties",
    ),
    property_gap(
        "cell",
        "alignment.justifyLastLine",
        "M1-extended (R4d)",
        "the XLSB node does not project this alignment attribute into cell properties",
    ),
    property_gap(
        "cell",
        "font.scheme",
        "M1-extended (R4d)",
        "the XLSB node does not project font records into cell properties",
    ),
    property_gap(
        "cell",
        "fill",
        "M1-extended (R4d)",
        "the XLSB node does not project fill records into cell properties",
    ),
    property_gap(
        "cell",
        "fillPattern",
        "M1-extended (R4d)",
        "the XLSB node does not project fill records into cell properties",
    ),
    property_gap(
        "cell",
        "fillBg",
        "M1-extended (R4d)",
        "the XLSB node does not project fill records into cell properties",
    ),
    property_gap(
        "cell",
        "border.left",
        "M1-extended (R4d)",
        "the XLSB node does not project border records into cell properties",
    ),
    property_gap(
        "cell",
        "border.left.color",
        "M1-extended (R4d)",
        "the XLSB node does not project border records into cell properties",
    ),
    property_gap(
        "cell",
        "border.right",
        "M1-extended (R4d)",
        "the XLSB node does not project border records into cell properties",
    ),
    property_gap(
        "cell",
        "border.right.color",
        "M1-extended (R4d)",
        "the XLSB node does not project border records into cell properties",
    ),
    property_gap(
        "cell",
        "border.top",
        "M1-extended (R4d)",
        "the XLSB node does not project border records into cell properties",
    ),
    property_gap(
        "cell",
        "border.top.color",
        "M1-extended (R4d)",
        "the XLSB node does not project border records into cell properties",
    ),
    property_gap(
        "cell",
        "border.bottom",
        "M1-extended (R4d)",
        "the XLSB node does not project border records into cell properties",
    ),
    property_gap(
        "cell",
        "border.bottom.color",
        "M1-extended (R4d)",
        "the XLSB node does not project border records into cell properties",
    ),
    property_gap(
        "cell",
        "border.diagonal",
        "M1-extended (R4d)",
        "the XLSB node does not project border records into cell properties",
    ),
    property_gap(
        "cell",
        "border.diagonal.color",
        "M1-extended (R4d)",
        "the XLSB node does not project border records into cell properties",
    ),
    property_gap(
        "cell",
        "border.diagonalUp",
        "M1-extended (R4d)",
        "the XLSB node does not project border records into cell properties",
    ),
    property_gap(
        "cell",
        "border.diagonalDown",
        "M1-extended (R4d)",
        "the XLSB node does not project border records into cell properties",
    ),
    property_gap(
        "cell",
        "protection.locked",
        "M1-extended (R4d)",
        "the XLSB node does not project cell protection metadata",
    ),
    property_gap(
        "cell",
        "protection.hidden",
        "M1-extended (R4d)",
        "the XLSB node does not project cell protection metadata",
    ),
    property_gap(
        "cell",
        "quotePrefix",
        "M1-extended (R4d)",
        "the XLSB node does not project quote-prefix metadata into cell properties",
    ),
    property_gap(
        "cell",
        "numFmtId",
        "M1-extended (R4d)",
        "native number-format IDs remain in the XLSB style table and are not admitted as unified properties",
    ),
    property_gap(
        "cell",
        "numberformat",
        "M1-extended (R4d)",
        "the XLSB node does not resolve number-format codes into cell properties",
    ),
    property_gap(
        "cell",
        "link",
        "M1-extended (R4x)",
        "XLSB hyperlink targets are retained by the leaf but not resolved into cell properties yet",
    ),
    property_gap(
        "cell",
        "tooltip",
        "M1-extended (R4x)",
        "XLSB hyperlink tooltips are not resolved into cell properties yet",
    ),
    property_gap(
        "cell",
        "display",
        "M1-extended (R4x)",
        "XLSB cells expose stored values, not a separately formatted display string",
    ),
    property_gap(
        "cell",
        "phonetic",
        "M1-extended (R4x)",
        "XLSB phonetic runs are retained by the leaf but not projected into cell properties yet",
    ),
    property_gap(
        "cell",
        "richtext",
        "M1-extended (R4x)",
        "the XLSB cell model exposes a flattened value and no unified rich-text run node",
    ),
    property_gap(
        "cell",
        "alt",
        "M1-extended (R4x)",
        "XLSB in-cell pictures are not read by the unified adapter",
    ),
    property_gap(
        "namedrange",
        "ref",
        "M1-extended (R4d)",
        "the XLSB defined-name catalog exposes only the authored name",
    ),
    property_gap(
        "namedrange",
        "scope",
        "M1-extended (R4d)",
        "the XLSB defined-name catalog exposes only the authored name",
    ),
    property_gap(
        "namedrange",
        "comment",
        "M1-extended (R4d)",
        "the XLSB defined-name catalog exposes only the authored name",
    ),
    property_gap(
        "namedrange",
        "volatile",
        "M1-extended (R4d)",
        "the XLSB defined-name catalog exposes only the authored name",
    ),
    property_gap(
        "namedrange",
        "hidden",
        "M1-extended (R4d)",
        "the XLSB defined-name catalog exposes only the authored name",
    ),
    property_gap(
        "column",
        "numberformat",
        "M1-extended (R4x)",
        "XLSB column styles are not projected into column nodes yet",
    ),
    property_gap(
        "range",
        "merge",
        "M1-extended (R4x)",
        "XLSB ranges do not report whether they are merged yet",
    ),
];

/// XLSB capability table: shared workbook structure, with binary-leaf gap
/// metadata and property follow-ups.
pub(crate) static SPEC: FormatSpec<Kind> = FormatSpec {
    root: workbook::SPEC.root,
    elements: &ELEMENTS,
    gaps: GAPS,
};

check_format_spec!(SPEC);
