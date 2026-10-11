//! XLS workbook vocabulary and capability table.
//!
//! The legacy BIFF owner currently exposes the workbook, ordinary worksheet
//! tabs, sparse cells, rows, columns and ranges.  The remaining shared
//! workbook vocabulary is kept in the table as explicit capability gaps so a
//! query never silently manufactures a node for a feature the native owner
//! does not project.

use crate::error::Code;
use crate::format::workbook::spec as workbook;
use crate::node::spec::{ElementSpec, FormatSpec, GapSpec, PropSpec, Subject, check_format_spec};
pub(crate) use workbook::{
    CELL_PROPS as SHARED_CELL_PROPS, ColumnKey, Kind, RowKey, SHEET_PROPS, SheetKey,
    WORKBOOK_PROPS as SHARED_WORKBOOK_PROPS, WorkbookKey,
};

/// The cell properties projected by the native BIFF reader. Formula text is
/// intentionally absent: the common Package projection retains values and a
/// formula bit, while source-backed cells do not retain formula tokens.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum CellKey {
    Type,
    Empty,
}

impl CellKey {
    const ALL: &'static [Self] = &[Self::Type, Self::Empty];

    pub(crate) fn at(index: usize) -> Option<Self> {
        Self::ALL.get(index).copied()
    }
}

pub(crate) static CELL_PROPS: &[PropSpec] = &[SHARED_CELL_PROPS[0], SHARED_CELL_PROPS[6]];

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

/// BIFF row and column layout is kept by the native reader, but the
/// source-backed Package bridge currently has no immutable layout accessor.
/// Keep these tables empty until both owner routes can answer them.
pub(crate) static ROW_PROPS: &[PropSpec] = &[];
pub(crate) static COLUMN_PROPS: &[PropSpec] = &[];

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
static FORMULA_GAP: GapSpec = property_gap(
    "cell",
    "formula",
    "M1-extended (R6b)",
    "the Package cell projection retains formula classification and cached scalar values but not formula token text",
);
static CACHED_VALUE_GAP: GapSpec = property_gap(
    "cell",
    "cachedValue",
    "M1-extended (R6b)",
    "the legacy source-backed cell projection does not expose a separate cached-value property",
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
static NAMED_RANGE_GAP: GapSpec = gap(
    "namedrange",
    "M1-extended (R6b)",
    "the legacy XLS Package read bridge does not project defined names into unified nodes yet",
);
static COMMENT_GAP: GapSpec = gap(
    "comment",
    "M1-extended (R6b)",
    "legacy XLS comments are retained by the leaf but are not projected into unified nodes yet",
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
static VALIDATION_GAP: GapSpec = gap(
    "dataValidation",
    "M1-extended (R6b)",
    "legacy XLS data validations are not projected into unified nodes yet",
);
static CF_GAP: GapSpec = gap(
    "conditionalFormatting",
    "M1-extended (R6b)",
    "legacy XLS conditional formats are not projected into unified nodes yet",
);
static AUTOFILTER_GAP: GapSpec = gap(
    "autofilter",
    "M1-extended (R6b)",
    "legacy XLS auto-filters are not projected into unified nodes yet",
);
static HYPERLINK_GAP: GapSpec = gap(
    "hyperlink",
    "M1-extended (R6b)",
    "legacy XLS hyperlinks are not projected into unified nodes yet",
);
static MERGE_GAP: GapSpec = gap(
    "mergeCell",
    "M1-extended (R6b)",
    "legacy XLS merged ranges are not projected into unified nodes yet",
);
static PICTURE_GAP: GapSpec = gap(
    "picture",
    "M1-extended (R6b)",
    "legacy XLS drawings are not projected into unified nodes yet",
);
static SHAPE_GAP: GapSpec = gap(
    "shape",
    "M1-extended (R6b)",
    "legacy XLS drawing shapes are not projected into unified nodes yet",
);
static OLE_GAP: GapSpec = gap(
    "ole",
    "M1-extended (R6b)",
    "legacy XLS OLE objects are not projected into unified nodes yet",
);
static ROW_BREAK_GAP: GapSpec = gap(
    "rowbreak",
    "M1-extended (R6b)",
    "legacy XLS row page breaks are not projected into unified nodes yet",
);
static COL_BREAK_GAP: GapSpec = gap(
    "colbreak",
    "M1-extended (R6b)",
    "legacy XLS column page breaks are not projected into unified nodes yet",
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
    with_props(&workbook::SPEC.elements[4], &[]),
    with_props(&workbook::SPEC.elements[5], COLUMN_PROPS),
    with_gap(&workbook::SPEC.elements[6], &TABLE_GAP),
    with_gap(&workbook::SPEC.elements[7], &NAMED_RANGE_GAP),
    with_gap(&workbook::SPEC.elements[8], &COMMENT_GAP),
    with_gap(&workbook::SPEC.elements[9], &CHART_GAP),
    with_gap(&workbook::SPEC.elements[10], &SERIES_GAP),
    with_gap(&workbook::SPEC.elements[11], &AXIS_GAP),
    with_gap(&workbook::SPEC.elements[12], &PIVOT_GAP),
    with_gap(&workbook::SPEC.elements[13], &SLICER_GAP),
    with_gap(&workbook::SPEC.elements[14], &SPARKLINE_GAP),
    with_gap(&workbook::SPEC.elements[15], &VALIDATION_GAP),
    with_gap(&workbook::SPEC.elements[16], &CF_GAP),
    with_gap(&workbook::SPEC.elements[17], &AUTOFILTER_GAP),
    with_gap(&workbook::SPEC.elements[18], &HYPERLINK_GAP),
    with_gap(&workbook::SPEC.elements[19], &MERGE_GAP),
    with_gap(&workbook::SPEC.elements[20], &PICTURE_GAP),
    with_gap(&workbook::SPEC.elements[21], &SHAPE_GAP),
    with_gap(&workbook::SPEC.elements[22], &OLE_GAP),
    with_gap(&workbook::SPEC.elements[23], &ROW_BREAK_GAP),
    with_gap(&workbook::SPEC.elements[24], &COL_BREAK_GAP),
    with_gap(&workbook::SPEC.elements[25], &RUN_GAP),
];

pub(crate) static SPEC: FormatSpec<Kind> = FormatSpec {
    root: workbook::SPEC.root,
    elements: &ELEMENTS,
    gaps: &[
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
        FORMULA_GAP,
        CACHED_VALUE_GAP,
        EVALUATED_GAP,
        ARRAY_FORMULA_GAP,
        ARRAY_REF_GAP,
        property_gap(
            "row",
            "height",
            "M1-extended (R6b)",
            "BIFF ROW layout is not projected",
        ),
        property_gap(
            "row",
            "hidden",
            "M1-extended (R6b)",
            "BIFF ROW layout is not projected",
        ),
        property_gap(
            "row",
            "outlineLevel",
            "M1-extended (R6b)",
            "BIFF ROW layout is not projected",
        ),
        property_gap(
            "row",
            "collapsed",
            "M1-extended (R6b)",
            "BIFF ROW layout is not projected",
        ),
        property_gap(
            "row",
            "customHeight",
            "M1-extended (R6b)",
            "BIFF ROW layout is not projected",
        ),
        property_gap(
            "column",
            "width",
            "M1-extended (R6b)",
            "BIFF COLINFO layout is not projected",
        ),
        property_gap(
            "column",
            "hidden",
            "M1-extended (R6b)",
            "BIFF COLINFO layout is not projected",
        ),
        property_gap(
            "column",
            "customWidth",
            "M1-extended (R6b)",
            "BIFF COLINFO layout is not projected",
        ),
        property_gap(
            "column",
            "outlineLevel",
            "M1-extended (R6b)",
            "BIFF COLINFO layout is not projected",
        ),
        property_gap(
            "column",
            "collapsed",
            "M1-extended (R6b)",
            "BIFF COLINFO layout is not projected",
        ),
    ],
};

check_format_spec!(SPEC);
