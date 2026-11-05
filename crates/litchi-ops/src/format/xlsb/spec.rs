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
use crate::registry::{Access, Alias, Constraint, Entry, Table};
use crate::value::{ColorContext, Length, LengthSpec, Tokens, ValueType};

pub(crate) use workbook::{
    COLUMN_PROPS, ColumnKey, Kind, ROW_PROPS, RowKey, SheetKey, TABLE_PROPS, TableKey,
    WORKBOOK_PROPS, WorkbookKey,
};

/// The XLSB sheet setter vocabulary, shared by coercion and capabilities.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum SheetSetter {
    Name,
    Hidden,
    Visibility,
}

pub(crate) const SHEET_WRITE_TABLE: Table<SheetSetter, 3, 0> = Table::new(
    [
        Entry {
            spec: PropSpec::new("name", ValueType::Text),
            access: Access::SetAndAdd(SheetSetter::Name),
            constraint: Constraint::None,
            clear: false,
        },
        Entry {
            spec: PropSpec::new("hidden", ValueType::Bool),
            access: Access::SetAndAdd(SheetSetter::Hidden),
            constraint: Constraint::None,
            clear: false,
        },
        Entry {
            spec: PropSpec::new("visibility", ValueType::Token(&workbook::VISIBILITY)),
            access: Access::SetAndAdd(SheetSetter::Visibility),
            constraint: Constraint::None,
            clear: false,
        },
    ],
    [],
);

/// Read-schema projection of the XLSB sheet write registry.
pub(crate) const SHEET_WRITE_PROPS: &[PropSpec] = &SHEET_WRITE_TABLE.readable();

/// XLSB's cell setter vocabulary (records 0168, 0261).
///
/// Payloads: the common scalar union (`value`), an explicit `clear: true`
/// (a null `value` is admitted through the registry's clear path and has the
/// same blank-cell meaning) and `formula` text, compiled by the leaf for
/// references within the sheet and stored for recalculation on open.
///
/// Formatting: font flags, size, name and RGB color; a solid RGB `fill`
/// (null removes the fill); the four outer border sides with RGB colors
/// (`border.all` sets all four); the number-format code; and the alignment
/// fields. Each change is a partial edit of the cell's current format, so
/// unnamed attributes keep their exact records.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum CellSetter {
    Value,
    Formula,
    Clear,
    Bold,
    Italic,
    Strike,
    Underline,
    FontName,
    FontSize,
    FontColor,
    Fill,
    BorderAll,
    BorderAllColor,
    BorderLeft,
    BorderLeftColor,
    BorderRight,
    BorderRightColor,
    BorderTop,
    BorderTopColor,
    BorderBottom,
    BorderBottomColor,
    NumberFormat,
    Horizontal,
    Vertical,
    WrapText,
    TextRotation,
    Indent,
    ShrinkToFit,
    ReadingOrder,
}

/// XLSB underline values: the binary font record keeps a single underline.
static XLSB_UNDERLINE: Tokens = Tokens::new(
    "underline",
    &["single", "none"],
    &[("false", 1), ("true", 0)],
);

/// XLSB border lines (`Blxf` styles 1–13) and `none`, which removes a line.
static XLSB_BORDER_LINE: Tokens = Tokens::new(
    "border style",
    &[
        "thin",
        "medium",
        "dashed",
        "dotted",
        "thick",
        "double",
        "hair",
        "mediumDashed",
        "dashDot",
        "mediumDashDot",
        "dashDotDot",
        "mediumDashDotDot",
        "slantDashDot",
        "none",
    ],
    &[],
);

const COLOR: ValueType = ValueType::Color(ColorContext::Spreadsheet);

const fn set(
    name: &'static str,
    ty: ValueType,
    setter: CellSetter,
    clear: bool,
) -> Entry<CellSetter> {
    Entry {
        spec: PropSpec::new(name, ty),
        access: Access::Set(setter),
        constraint: Constraint::None,
        clear,
    }
}

const fn constrained(
    name: &'static str,
    ty: ValueType,
    setter: CellSetter,
    constraint: Constraint,
) -> Entry<CellSetter> {
    Entry {
        spec: PropSpec::new(name, ty),
        access: Access::Set(setter),
        constraint,
        clear: false,
    }
}

const fn write_only_entry(
    name: &'static str,
    ty: ValueType,
    setter: CellSetter,
    clear: bool,
) -> Entry<CellSetter> {
    Entry {
        spec: write_only(name, ty),
        access: Access::Set(setter),
        constraint: Constraint::None,
        clear,
    }
}

pub(crate) const CELL_WRITE_TABLE: Table<CellSetter, 29, 25> = Table::new(
    [
        write_only_entry("value", ValueType::Scalar, CellSetter::Value, true),
        constrained(
            "formula",
            ValueType::Text,
            CellSetter::Formula,
            Constraint::NotEmpty,
        ),
        write_only_entry("clear", ValueType::Bool, CellSetter::Clear, false),
        set("font.bold", ValueType::Bool, CellSetter::Bold, false),
        set("font.italic", ValueType::Bool, CellSetter::Italic, false),
        set("strike", ValueType::Bool, CellSetter::Strike, false),
        set(
            "underline",
            ValueType::Token(&XLSB_UNDERLINE),
            CellSetter::Underline,
            false,
        ),
        constrained(
            "font.name",
            ValueType::Text,
            CellSetter::FontName,
            Constraint::NotEmpty,
        ),
        constrained(
            "font.size",
            ValueType::Length(LengthSpec::FONT_SIZE),
            CellSetter::FontSize,
            Constraint::Length {
                min: Length::from_points(1),
                max: Length::from_points(409),
            },
        ),
        set("font.color", COLOR, CellSetter::FontColor, true),
        set("fill", COLOR, CellSetter::Fill, true),
        write_only_entry(
            "border.all",
            ValueType::Token(&XLSB_BORDER_LINE),
            CellSetter::BorderAll,
            true,
        ),
        write_only_entry("border.all.color", COLOR, CellSetter::BorderAllColor, false),
        set(
            "border.left",
            ValueType::Token(&XLSB_BORDER_LINE),
            CellSetter::BorderLeft,
            true,
        ),
        set(
            "border.left.color",
            COLOR,
            CellSetter::BorderLeftColor,
            false,
        ),
        set(
            "border.right",
            ValueType::Token(&XLSB_BORDER_LINE),
            CellSetter::BorderRight,
            true,
        ),
        set(
            "border.right.color",
            COLOR,
            CellSetter::BorderRightColor,
            false,
        ),
        set(
            "border.top",
            ValueType::Token(&XLSB_BORDER_LINE),
            CellSetter::BorderTop,
            true,
        ),
        set("border.top.color", COLOR, CellSetter::BorderTopColor, false),
        set(
            "border.bottom",
            ValueType::Token(&XLSB_BORDER_LINE),
            CellSetter::BorderBottom,
            true,
        ),
        set(
            "border.bottom.color",
            COLOR,
            CellSetter::BorderBottomColor,
            false,
        ),
        constrained(
            "numberformat",
            ValueType::Text,
            CellSetter::NumberFormat,
            Constraint::NotEmpty,
        ),
        set(
            "alignment.horizontal",
            ValueType::Token(&workbook::HORIZONTAL),
            CellSetter::Horizontal,
            false,
        ),
        set(
            "alignment.vertical",
            ValueType::Token(&workbook::VERTICAL),
            CellSetter::Vertical,
            false,
        ),
        set(
            "alignment.wrapText",
            ValueType::Bool,
            CellSetter::WrapText,
            false,
        ),
        constrained(
            "alignment.textRotation",
            ValueType::Int,
            CellSetter::TextRotation,
            Constraint::Int { min: 0, max: 255 },
        ),
        constrained(
            "alignment.indent",
            ValueType::Int,
            CellSetter::Indent,
            Constraint::Int { min: 0, max: 250 },
        ),
        set(
            "alignment.shrinkToFit",
            ValueType::Bool,
            CellSetter::ShrinkToFit,
            false,
        ),
        set(
            "alignment.readingOrder",
            ValueType::Token(&workbook::READING_ORDER),
            CellSetter::ReadingOrder,
            false,
        ),
    ],
    [
        // `text` names `value`, as in the XLSX cell registry (record 0269).
        Alias {
            name: "text",
            entry: 0,
        },
        Alias {
            name: "bold",
            entry: 3,
        },
        Alias {
            name: "italic",
            entry: 4,
        },
        Alias {
            name: "font.strike",
            entry: 5,
        },
        Alias {
            name: "strikethrough",
            entry: 5,
        },
        Alias {
            name: "font.underline",
            entry: 6,
        },
        Alias {
            name: "font",
            entry: 7,
        },
        Alias {
            name: "fontname",
            entry: 7,
        },
        Alias {
            name: "size",
            entry: 8,
        },
        Alias {
            name: "fontsize",
            entry: 8,
        },
        Alias {
            name: "color",
            entry: 9,
        },
        Alias {
            name: "bgcolor",
            entry: 10,
        },
        Alias {
            name: "bg",
            entry: 10,
        },
        Alias {
            name: "border",
            entry: 11,
        },
        Alias {
            name: "border.color",
            entry: 12,
        },
        Alias {
            name: "numfmt",
            entry: 21,
        },
        Alias {
            name: "format",
            entry: 21,
        },
        Alias {
            name: "halign",
            entry: 22,
        },
        Alias {
            name: "align",
            entry: 22,
        },
        Alias {
            name: "valign",
            entry: 23,
        },
        Alias {
            name: "wrap",
            entry: 24,
        },
        Alias {
            name: "wraptext",
            entry: 24,
        },
        Alias {
            name: "rotation",
            entry: 25,
        },
        Alias {
            name: "indent",
            entry: 26,
        },
        Alias {
            name: "readingorder",
            entry: 28,
        },
    ],
);

const fn write_only(name: &'static str, ty: ValueType) -> PropSpec {
    let mut spec = PropSpec::new(name, ty);
    spec.access.get = false;
    spec
}

pub(crate) const CELL_WRITE_PROPS: &[PropSpec] = &CELL_WRITE_TABLE.readable();

/// How many write properties have no read property of the same name: the
/// write-only payloads (`value`, `clear`, `border.all`, `border.all.color`).
const fn write_only_count() -> usize {
    let mut count = 0;
    let mut write_index = 0;
    while write_index < CELL_WRITE_PROPS.len() {
        let mut found = false;
        let mut index = 0;
        while index < CELL_PROPS.len() {
            if same_name(CELL_PROPS[index].name, CELL_WRITE_PROPS[write_index].name) {
                found = true;
            }
            index += 1;
        }
        if !found {
            count += 1;
        }
        write_index += 1;
    }
    count
}

const WRITE_ONLY_COUNT: usize = write_only_count();

/// Capability projection for cell writes.  The read projection remains the
/// local `CELL_PROPS` slice with matching registry properties upgraded to
/// writable; the write-only entries are appended so the capability report
/// exposes their registry types and set flags.
const fn cell_capability_array() -> [PropSpec; CELL_PROPS.len() + WRITE_ONLY_COUNT] {
    let mut props = [PropSpec::new("", ValueType::Text); CELL_PROPS.len() + WRITE_ONLY_COUNT];
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
    let mut next = CELL_PROPS.len();
    let mut write_index = 0;
    while write_index < CELL_WRITE_PROPS.len() {
        let write = CELL_WRITE_PROPS[write_index];
        let mut found = false;
        let mut read = 0;
        while read < CELL_PROPS.len() {
            if same_name(CELL_PROPS[read].name, write.name) {
                found = true;
            }
            read += 1;
        }
        if !found {
            props[next] = write;
            next += 1;
        }
        write_index += 1;
    }
    props
}

pub(crate) static CELL_CAPABILITY_PROPS: &[PropSpec] = &cell_capability_array();

/// The properties a range target accepts: the cell setters (applied to
/// every cell of the range) and `merge`.
pub(crate) static RANGE_PROPS: &[PropSpec] = &range_props();

/// The index of `merge` in [`RANGE_PROPS`].
pub(crate) const RANGE_MERGE_INDEX: usize = CELL_WRITE_PROPS.len();

const fn range_props() -> [PropSpec; CELL_WRITE_PROPS.len() + 1] {
    let mut props = [PropSpec::new("", ValueType::Text); CELL_WRITE_PROPS.len() + 1];
    let mut index = 0;
    while index < CELL_WRITE_PROPS.len() {
        // A range reads no per-cell property; its setters are write-only.
        let mut prop = CELL_WRITE_PROPS[index];
        prop.access.get = false;
        props[index] = prop;
        index += 1;
    }
    props[CELL_WRITE_PROPS.len()] = MERGE_WRITE_PROPS[0];
    props
}

/// The XLSB merge setter: `merge` on a range target, `true` to merge the
/// range exactly and `false` or null to remove exactly that merge.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum MergeSetter {
    Merge,
}

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
                prop.admission = write.admission;
            }
            write_index += 1;
        }
        props[index] = prop;
        index += 1;
    }
    props
}

/// XLSB's local sheet table with registry-derived setter access.
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
    Bold,
    Italic,
    Strike,
    Underline,
    FontColor,
    FontSize,
    FontName,
    Fill,
    BorderLeft,
    BorderLeftColor,
    BorderRight,
    BorderRightColor,
    BorderTop,
    BorderTopColor,
    BorderBottom,
    BorderBottomColor,
    WrapText,
    Horizontal,
    Vertical,
    TextRotation,
    Indent,
    ShrinkToFit,
    ReadingOrder,
    NumberFormat,
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
        Self::Bold,
        Self::Italic,
        Self::Strike,
        Self::Underline,
        Self::FontColor,
        Self::FontSize,
        Self::FontName,
        Self::Fill,
        Self::BorderLeft,
        Self::BorderLeftColor,
        Self::BorderRight,
        Self::BorderRightColor,
        Self::BorderTop,
        Self::BorderTopColor,
        Self::BorderBottom,
        Self::BorderBottomColor,
        Self::WrapText,
        Self::Horizontal,
        Self::Vertical,
        Self::TextRotation,
        Self::Indent,
        Self::ShrinkToFit,
        Self::ReadingOrder,
        Self::NumberFormat,
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
    workbook::CELL_PROPS[7],
    workbook::CELL_PROPS[8],
    workbook::CELL_PROPS[9],
    // The binary font keeps a single underline; advertise what it stores.
    PropSpec::new("underline", ValueType::Token(&XLSB_UNDERLINE)),
    workbook::CELL_PROPS[11],
    workbook::CELL_PROPS[14],
    workbook::CELL_PROPS[15],
    workbook::CELL_PROPS[16],
    workbook::CELL_PROPS[19],
    workbook::CELL_PROPS[20],
    workbook::CELL_PROPS[21],
    workbook::CELL_PROPS[22],
    workbook::CELL_PROPS[23],
    workbook::CELL_PROPS[24],
    workbook::CELL_PROPS[25],
    workbook::CELL_PROPS[26],
    workbook::CELL_PROPS[31],
    workbook::CELL_PROPS[32],
    workbook::CELL_PROPS[33],
    workbook::CELL_PROPS[34],
    workbook::CELL_PROPS[35],
    workbook::CELL_PROPS[36],
    workbook::CELL_PROPS[37],
    workbook::CELL_PROPS[42],
    workbook::CELL_PROPS[43],
    workbook::CELL_PROPS[44],
];

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum NamedRangeKey {
    Name,
    Ref,
    Scope,
}

impl NamedRangeKey {
    const ALL: &'static [Self] = &[Self::Name, Self::Ref, Self::Scope];

    pub(crate) fn at(index: usize) -> Option<Self> {
        Self::ALL.get(index).copied()
    }
}

/// The XLSB defined-name setters (record 0261): `add` takes `name`, `ref`
/// and an optional `scope` (a sheet name; absent is workbook scope); `set`
/// replaces `ref`.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum NamedRangeSetter {
    Name,
    Ref,
    Scope,
}

pub(crate) const NAMED_RANGE_WRITE_TABLE: Table<NamedRangeSetter, 3, 0> = Table::new(
    [
        Entry {
            spec: PropSpec::new("name", ValueType::Text),
            access: Access::Add(NamedRangeSetter::Name),
            constraint: Constraint::NotEmpty,
            clear: false,
        },
        Entry {
            spec: PropSpec::new("ref", ValueType::Text),
            access: Access::SetAndAdd(NamedRangeSetter::Ref),
            constraint: Constraint::NotEmpty,
            clear: false,
        },
        Entry {
            spec: PropSpec::new("scope", ValueType::Text),
            access: Access::Add(NamedRangeSetter::Scope),
            constraint: Constraint::NotEmpty,
            clear: false,
        },
    ],
    [],
);

const NAMED_RANGE_WRITE_PROPS: &[PropSpec] = &NAMED_RANGE_WRITE_TABLE.readable();

const fn named_range_props() -> [PropSpec; 3] {
    let mut props = [
        workbook::NAMED_RANGE_PROPS[0],
        workbook::NAMED_RANGE_PROPS[1],
        workbook::NAMED_RANGE_PROPS[2],
    ];
    let mut index = 0;
    while index < props.len() {
        props[index].access.set = NAMED_RANGE_WRITE_PROPS[index].access.set;
        props[index].access.add = NAMED_RANGE_WRITE_PROPS[index].access.add;
        index += 1;
    }
    props
}

pub(crate) static NAMED_RANGE_PROPS: &[PropSpec] = &named_range_props();

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

// The binary leaf already retains these parts, but the XLSB adapter does
// not project them into ordinary nodes yet (record 0322 projects the
// worksheet features).  Keep the reasons format-local; in particular, do not
// inherit the XLSX follow-ups from the shared table.
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
static MERGE_GAP: GapSpec = element_gap(
    "mergeCell",
    "M1-extended (R4x)",
    "XLSB merged ranges are retained by the leaf but are not listed as nodes yet",
);
static OLE_GAP: GapSpec = element_gap(
    "ole",
    "M1-extended (R4x)",
    "XLSB OLE objects are retained by the leaf but are not projected into nodes yet",
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

/// A chart axis resolves by its role, like the XLSX axis (record 0322).
const XLSB_AXIS: ElementSpec<Kind> = ElementSpec {
    props: super::chart::AXIS_PROPS,
    identities: &["role"],
    gap: None,
    paths: &[
        "/<sheet>/chart[N]/axis[@role=<role>]",
        "/<sheet>/chart[N]/axis[N]",
    ],
    ..workbook::SPEC.elements[11]
};

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
    with_props(&workbook::SPEC.elements[3], RANGE_PROPS, None),
    with_gap(&workbook::SPEC.elements[4], None),
    with_gap(&workbook::SPEC.elements[5], None),
    with_gap(&workbook::SPEC.elements[6], None),
    with_props(&workbook::SPEC.elements[7], NAMED_RANGE_PROPS, None),
    with_props(
        &workbook::SPEC.elements[8],
        super::features::COMMENT_PROPS,
        None,
    ),
    with_props(&workbook::SPEC.elements[9], super::chart::CHART_PROPS, None),
    with_props(
        &workbook::SPEC.elements[10],
        super::chart::SERIES_PROPS,
        None,
    ),
    XLSB_AXIS,
    with_gap(&workbook::SPEC.elements[12], Some(&PIVOT_GAP)),
    with_gap(&workbook::SPEC.elements[13], Some(&SLICER_GAP)),
    with_gap(&workbook::SPEC.elements[14], Some(&SPARKLINE_GAP)),
    with_props(
        &workbook::SPEC.elements[15],
        super::features::VALIDATION_PROPS,
        None,
    ),
    with_props(
        &workbook::SPEC.elements[16],
        super::features::CF_PROPS,
        None,
    ),
    with_props(
        &workbook::SPEC.elements[17],
        super::features::AUTOFILTER_PROPS,
        None,
    ),
    with_props(
        &workbook::SPEC.elements[18],
        super::features::HYPERLINK_PROPS,
        None,
    ),
    with_gap(&workbook::SPEC.elements[19], Some(&MERGE_GAP)),
    with_props(
        &workbook::SPEC.elements[20],
        super::features::PICTURE_PROPS,
        None,
    ),
    with_props(
        &workbook::SPEC.elements[21],
        super::features::SHAPE_PROPS,
        None,
    ),
    with_gap(&workbook::SPEC.elements[22], Some(&OLE_GAP)),
    with_props(
        &workbook::SPEC.elements[23],
        super::features::ROW_BREAK_PROPS,
        None,
    ),
    with_props(
        &workbook::SPEC.elements[24],
        super::features::COL_BREAK_PROPS,
        None,
    ),
    with_gap(&workbook::SPEC.elements[25], Some(&RUN_GAP)),
];

// These are the workbook properties whose shared baseline deliberately does
// not claim a read.  The values are the same gaps as XLSX's table, but the
// reasons describe the XLSB leaf and its deferred binary parts.
static GAPS: &[GapSpec] = &[
    // OfficeCLI's heuristic table node (record 0339).
    element_gap(
        "detectedtable",
        "0339-F5 (deferred-P1): header-row table detection in queries",
        "OfficeCLI's detectedtable is a header-row block found heuristically in unstructured cells and reported by table queries; Litchi's queries return declared tables only",
    ),
    // Defined names are added, retargeted and removed (record 0261), but
    // only the last name can go: removing an earlier one would renumber the
    // name index every formula's name references use (record 0330).
    GapSpec {
        subject: Subject::Operation {
            verb: "remove",
            element: "namedrange",
            context: "earlier-name",
        },
        code: Code::UnsupportedCapability,
        follow_up: "XLSB defined-name index renumbering (record 0330)",
        reason: "only the workbook's last defined name can be removed; an earlier one would renumber every formula's name references",
    },
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
        "M4 follow-up: project computed values in cell views",
        "node views do not evaluate; `File::evaluate` and `File::recalculate` compute formula results from the full grid (record 0266); the source cached-value fallback has no formula text",
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
        "comment",
        "M1-extended (R4d)",
        "the XLSB defined-name projection reads name, ref and scope; BrtName comments are not projected yet",
    ),
    property_gap(
        "namedrange",
        "volatile",
        "M1-extended (R4d)",
        "the XLSB defined-name projection reads name, ref and scope; macro and volatile flags are not projected yet",
    ),
    property_gap(
        "namedrange",
        "hidden",
        "M1-extended (R4d)",
        "the XLSB defined-name projection reads name, ref and scope; the hidden flag is not projected yet",
    ),
    property_gap(
        "column",
        "numberformat",
        "M1-extended (R4x)",
        "XLSB column styles are not projected into column nodes yet",
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
