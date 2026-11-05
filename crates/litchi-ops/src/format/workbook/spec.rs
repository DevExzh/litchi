//! Shared workbook element kinds and baseline spec table (record 0010 §4.2, §4.7,
//! §7.2): the one table path resolution, the query compiler, the node view
//! and `capabilities()` read.
//!
//! Every canonical name of the workbook vocabulary is a [`Kind`], so the
//! resolver never falls back to an unlisted `unsupported_capability`: the
//! elements Litchi reads (workbook, sheet, cell, range, row, column, table,
//! defined name) carry their property tables, and the others carry the
//! [`GapSpec`] that names their follow-up.
//!
//! Property keys are OfficeCLI's (`schemas/help/xlsx/*.json`, and the order
//! `ExcelHandler.Helpers.Node.cs` `CellToNode`, `ExcelHandler.Query.cs` and
//! `TableToNode` emit them). [`props!`] declares each table together with a
//! key enum whose discriminant is the key's index, so the node view reads a
//! key with one `match` and no string comparison.

// ODS reads the kinds, the path tree and the workbook keys; the shared
// SpreadsheetML token tables serve the other workbook adapters only.
#![cfg_attr(
    not(any(feature = "xls", feature = "xlsx", feature = "xlsb")),
    allow(dead_code, reason = "an ODS-only build reads the kinds and path tree")
)]

use crate::error::Code;
use crate::node::spec::{ElementSpec, FormatSpec, GapSpec, PropSpec, Subject, check_format_spec};
use crate::path::navigate::ElementKind;
use crate::path::vocabulary::{self, Vocabulary};
use crate::value::{ColorContext, LengthContext, LengthSpec, LengthStyle, Tokens, ValueType};

/// The element kinds of an XLSX workbook: the workbook vocabulary's
/// canonical names, in spec-table order.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[repr(u8)]
pub(crate) enum Kind {
    /// The workbook (`/`).
    Workbook,
    /// A sheet of any kind (`/Sheet1`, `/sheet[N]`).
    Sheet,
    /// A worksheet cell (`/Sheet1/A1`).
    Cell,
    /// A worksheet range (`/Sheet1/A1:C5`).
    Range,
    /// A worksheet row by number (`/Sheet1/row[3]`).
    Row,
    /// A worksheet column (`/Sheet1/col[B]`).
    Column,
    /// A worksheet table, a ListObject (`/Sheet1/table[1]`).
    Table,
    /// A defined name (`/namedrange[1]`).
    NamedRange,
    /// A legacy cell comment (gap).
    Comment,
    /// A chart (gap).
    Chart,
    /// A chart series (gap).
    Series,
    /// A chart axis (gap).
    Axis,
    /// A pivot table (gap).
    PivotTable,
    /// A slicer (gap).
    Slicer,
    /// A sparkline group (gap).
    Sparkline,
    /// A data validation (gap).
    DataValidation,
    /// A conditional format (gap).
    Cf,
    /// The worksheet auto-filter (gap).
    Autofilter,
    /// A hyperlink (gap).
    Hyperlink,
    /// A merged range (gap).
    MergeCell,
    /// A picture (gap).
    Picture,
    /// A drawing shape (gap).
    Shape,
    /// An embedded OLE object (gap).
    Ole,
    /// A manual row page break (gap).
    RowBreak,
    /// A manual column page break (gap).
    ColBreak,
    /// A rich-text run of a cell (gap).
    Run,
}

impl Kind {
    /// Every kind, by ordinal.
    const ALL: [Self; 26] = [
        Self::Workbook,
        Self::Sheet,
        Self::Cell,
        Self::Range,
        Self::Row,
        Self::Column,
        Self::Table,
        Self::NamedRange,
        Self::Comment,
        Self::Chart,
        Self::Series,
        Self::Axis,
        Self::PivotTable,
        Self::Slicer,
        Self::Sparkline,
        Self::DataValidation,
        Self::Cf,
        Self::Autofilter,
        Self::Hyperlink,
        Self::MergeCell,
        Self::Picture,
        Self::Shape,
        Self::Ole,
        Self::RowBreak,
        Self::ColBreak,
        Self::Run,
    ];

    /// The kind's spec entry.
    #[cfg(feature = "xlsx")]
    pub(crate) fn spec(self) -> &'static ElementSpec<Self> {
        &SPEC.elements[self as usize]
    }

    /// The canonical vocabulary name (`col`, `cf`, `dataValidation`).
    const fn canonical(self) -> &'static str {
        match self {
            Self::Workbook => "workbook",
            Self::Sheet => "sheet",
            Self::Cell => "cell",
            Self::Range => "range",
            Self::Row => "row",
            Self::Column => "col",
            Self::Table => "table",
            Self::NamedRange => "namedrange",
            Self::Comment => "comment",
            Self::Chart => "chart",
            Self::Series => "series",
            Self::Axis => "axis",
            Self::PivotTable => "pivottable",
            Self::Slicer => "slicer",
            Self::Sparkline => "sparkline",
            Self::DataValidation => "dataValidation",
            Self::Cf => "cf",
            Self::Autofilter => "autofilter",
            Self::Hyperlink => "hyperlink",
            Self::MergeCell => "mergeCell",
            Self::Picture => "picture",
            Self::Shape => "shape",
            Self::Ole => "ole",
            Self::RowBreak => "rowbreak",
            Self::ColBreak => "colbreak",
            Self::Run => "run",
        }
    }
}

impl ElementKind for Kind {
    const VOCABULARY: &'static Vocabulary = &vocabulary::WORKBOOK;
    const ROOT: Self = Self::Workbook;

    fn from_name(name: &str) -> Option<Self> {
        Self::ALL.into_iter().find(|kind| kind.canonical() == name)
    }

    fn name(self) -> &'static str {
        self.canonical()
    }

    fn ordinal(self) -> u8 {
        self as u8
    }
}

// The kinds are the workbook vocabulary's elements, one for one.
const _: () = assert!(Kind::ALL.len() == vocabulary::WORKBOOK.elements.len());

/// Declares one property table with its key enum: the enum's discriminant
/// is the key's index in the table.
macro_rules! props {
    (
        $(#[$meta:meta])*
        $key:ident in $table:ident {
            $($variant:ident => $name:literal: $ty:expr $(; $policy:ident)?,)+
        }
    ) => {
        $(#[$meta])*
        #[derive(Debug, Clone, Copy, PartialEq, Eq)]
        #[repr(u8)]
        pub(crate) enum $key {
            $(
                #[doc = concat!("`", $name, "`")]
                $variant,
            )+
        }

        $(#[$meta])*
        impl $key {
            const ALL: &'static [Self] = &[$(Self::$variant,)+];

            /// The key at `index` of its property table.
            pub(crate) fn at(index: usize) -> Option<Self> {
                Self::ALL.get(index).copied()
            }
        }

        pub(crate) const $table: &[PropSpec] = &[$(PropSpec::new($name, $ty)$(.$policy())?,)+];
    };
}

static CELL_TYPE: Tokens = Tokens::new(
    "cell type",
    &[
        "Number",
        "String",
        "SharedString",
        "Boolean",
        "Error",
        "InlineString",
        "Date",
        "Unknown",
    ],
    &[],
);
static UNDERLINE: Tokens = Tokens::new(
    "underline",
    &["single", "double", "singleAccounting", "doubleAccounting"],
    &[],
);
pub(crate) static FONT_SCHEME: Tokens =
    Tokens::new("font scheme", &["none", "major", "minor"], &[]);
pub(crate) static FILL_PATTERN: Tokens = Tokens::new(
    "fill pattern",
    &[
        "none",
        "solid",
        "mediumGray",
        "darkGray",
        "lightGray",
        "darkHorizontal",
        "darkVertical",
        "darkDown",
        "darkUp",
        "darkGrid",
        "darkTrellis",
        "lightHorizontal",
        "lightVertical",
        "lightDown",
        "lightUp",
        "lightGrid",
        "lightTrellis",
        "gray125",
        "gray0625",
    ],
    &[("gray25", 4), ("gray50", 2), ("gray75", 3)],
);
static BORDER_LINE: Tokens = Tokens::new(
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
    ],
    &[],
);
pub(crate) static HORIZONTAL: Tokens = Tokens::new(
    "horizontal alignment",
    &[
        "general",
        "left",
        "center",
        "right",
        "fill",
        "justify",
        "centerContinuous",
        "distributed",
    ],
    &[("centeracrossselection", 6)],
);
pub(crate) static VERTICAL: Tokens = Tokens::new(
    "vertical alignment",
    &["top", "center", "bottom", "justify", "distributed"],
    &[],
);
pub(crate) static READING_ORDER: Tokens = Tokens::new(
    "reading order",
    &["context", "ltr", "rtl"],
    &[
        ("0", 0),
        ("1", 1),
        ("2", 2),
        ("contextdependent", 0),
        ("lefttoright", 1),
        ("righttoleft", 2),
    ],
);
pub(crate) static VISIBILITY: Tokens =
    Tokens::new("visibility", &["visible", "hidden", "veryHidden"], &[]);

/// Spreadsheet colours: ARGB or a theme slot with its tint.
const COLOR: ValueType = ValueType::Color(ColorContext::Spreadsheet);

/// Row heights: points, a bare number read as points (OfficeCLI writes
/// `"15pt"`).
const ROW_HEIGHT: ValueType = ValueType::Length(LengthSpec::new(
    LengthContext::LineWidth,
    LengthStyle::Points,
));

props! {
    /// The workbook root (`ExcelHandler.Query.cs` `Get("/")` and
    /// `PopulateWorkbookSettings`).
    WorkbookKey in WORKBOOK_PROPS {
        Title => "title": ValueType::Text,
        Author => "author": ValueType::Text,
        Subject => "subject": ValueType::Text,
        Keywords => "keywords": ValueType::List,
        Description => "description": ValueType::Text,
        Category => "category": ValueType::Text,
        LastModifiedBy => "lastModifiedBy": ValueType::Text,
        RevisionNumber => "revisionNumber": ValueType::Text,
        Created => "created": ValueType::Text,
        Modified => "modified": ValueType::Text,
        Date1904 => "workbook.date1904": ValueType::Bool,
        ActiveTab => "activeTab": ValueType::Int,
        ExtendedTemplate => "extended.template": ValueType::Text,
        ExtendedManager => "extended.manager": ValueType::Text,
        ExtendedCompany => "extended.company": ValueType::Text,
        ExtendedApplication => "extended.application": ValueType::Text,
        ExtendedPages => "extended.pages": ValueType::Int,
        ExtendedWords => "extended.words": ValueType::Int,
        ExtendedCharacters => "extended.characters": ValueType::Int,
        ExtendedTotalTime => "extended.totalTime": ValueType::Int,
    }
}

props! {
    /// A sheet (`schemas/help/xlsx/sheet.json`).
    SheetKey in SHEET_PROPS {
        Name => "name": ValueType::Text,
        Hidden => "hidden": ValueType::Bool,
        Visibility => "visibility": ValueType::Token(&VISIBILITY),
    }
}

props! {
    /// A cell, in `CellToNode` order.
    #[cfg(feature = "xlsx")]
    CellKey in CELL_PROPS {
        Type => "type": ValueType::Token(&CELL_TYPE),
        Formula => "formula": ValueType::Text,
        CachedValue => "cachedValue": ValueType::Text,
        Evaluated => "evaluated": ValueType::Bool,
        ArrayFormula => "arrayformula": ValueType::Bool,
        ArrayRef => "arrayref": ValueType::Text,
        Empty => "empty": ValueType::Bool,
        Bold => "font.bold": ValueType::Bool,
        Italic => "font.italic": ValueType::Bool,
        Strike => "strike": ValueType::Bool,
        Underline => "underline": ValueType::Token(&UNDERLINE),
        FontColor => "font.color": COLOR,
        Superscript => "superscript": ValueType::Bool,
        Subscript => "subscript": ValueType::Bool,
        FontSize => "font.size": ValueType::Length(LengthSpec::FONT_SIZE),
        FontName => "font.name": ValueType::Text,
        Fill => "fill": COLOR,
        FillPattern => "fillPattern": ValueType::Token(&FILL_PATTERN),
        FillBg => "fillBg": COLOR,
        BorderLeft => "border.left": ValueType::Token(&BORDER_LINE),
        BorderLeftColor => "border.left.color": COLOR,
        BorderRight => "border.right": ValueType::Token(&BORDER_LINE),
        BorderRightColor => "border.right.color": COLOR,
        BorderTop => "border.top": ValueType::Token(&BORDER_LINE),
        BorderTopColor => "border.top.color": COLOR,
        BorderBottom => "border.bottom": ValueType::Token(&BORDER_LINE),
        BorderBottomColor => "border.bottom.color": COLOR,
        BorderDiagonal => "border.diagonal": ValueType::Token(&BORDER_LINE),
        BorderDiagonalColor => "border.diagonal.color": COLOR,
        BorderDiagonalUp => "border.diagonalUp": ValueType::Bool,
        BorderDiagonalDown => "border.diagonalDown": ValueType::Bool,
        WrapText => "alignment.wrapText": ValueType::Bool,
        Horizontal => "alignment.horizontal": ValueType::Token(&HORIZONTAL),
        Vertical => "alignment.vertical": ValueType::Token(&VERTICAL),
        TextRotation => "alignment.textRotation": ValueType::Int,
        Indent => "alignment.indent": ValueType::Int,
        ShrinkToFit => "alignment.shrinkToFit": ValueType::Bool,
        ReadingOrder => "alignment.readingOrder": ValueType::Token(&READING_ORDER),
        Locked => "protection.locked": ValueType::Bool,
        FormulaHidden => "protection.hidden": ValueType::Bool,
        QuotePrefix => "quotePrefix": ValueType::Bool,
        NumFmtId => "numFmtId": ValueType::Int; raw,
        NumberFormat => "numberformat": ValueType::Text,
        Merge => "merge": ValueType::Text,
        MergeAnchor => "mergeAnchor": ValueType::Bool,
        FontOutline => "font.outline": ValueType::Bool,
        FontShadow => "font.shadow": ValueType::Bool,
        FontCondense => "font.condense": ValueType::Bool,
        FontExtend => "font.extend": ValueType::Bool,
        FontCharset => "font.charset": ValueType::Int,
        FontFamily => "font.family": ValueType::Int,
        FontScheme => "font.scheme": ValueType::Token(&FONT_SCHEME),
        RelativeIndent => "alignment.relativeIndent": ValueType::Int,
        JustifyLastLine => "alignment.justifyLastLine": ValueType::Bool,
    }
}

props! {
    /// A row (`GetSheetChildNodes` and the row query; `ROW_PROPERTIES` of the
    /// query compiler).
    RowKey in ROW_PROPS {
        Height => "height": ROW_HEIGHT,
        Hidden => "hidden": ValueType::Bool,
        OutlineLevel => "outlineLevel": ValueType::Int,
        Collapsed => "collapsed": ValueType::Bool,
        CustomHeight => "customHeight": ValueType::Bool,
    }
}

props! {
    /// A column (`Get` of `/Sheet1/col[B]` and the column query).
    ColumnKey in COLUMN_PROPS {
        Width => "width": ValueType::Number,
        Hidden => "hidden": ValueType::Bool,
        CustomWidth => "customWidth": ValueType::Bool,
        OutlineLevel => "outlineLevel": ValueType::Int,
        Collapsed => "collapsed": ValueType::Bool,
    }
}

props! {
    /// A table, in `TableToNode` order.
    #[cfg(any(feature = "xlsx", feature = "xlsb"))]
    TableKey in TABLE_PROPS {
        Name => "name": ValueType::Text,
        DisplayName => "displayName": ValueType::Text,
        Ref => "ref": ValueType::Text,
        Style => "style": ValueType::Text,
        BandedRows => "bandedRows": ValueType::Bool,
        BandedCols => "bandedCols": ValueType::Bool,
        FirstCol => "firstCol": ValueType::Bool,
        LastCol => "lastCol": ValueType::Bool,
        HeaderRow => "headerRow": ValueType::Bool,
        TotalRow => "totalRow": ValueType::Bool,
        Columns => "columns": ValueType::Text,
    }
}

props! {
    /// A defined name (`/namedrange[N]`).
    #[cfg(any(feature = "xlsx", feature = "xls"))]
    NamedRangeKey in NAMED_RANGE_PROPS {
        Name => "name": ValueType::Text,
        Ref => "ref": ValueType::Text,
        Scope => "scope": ValueType::Text,
        Comment => "comment": ValueType::Text,
        Volatile => "volatile": ValueType::Bool,
    }
}

/// An element gap whose leaf model is planned (record 0010 §4.11).
const fn leaf_gap(element: &'static str, follow_up: &'static str, reason: &'static str) -> GapSpec {
    GapSpec {
        subject: Subject::Element(element),
        code: Code::UnsupportedCapability,
        follow_up,
        reason,
    }
}

/// A property gap of an element Litchi reads.
const fn prop_gap(
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

static COMMENT_GAP: GapSpec = leaf_gap(
    "comment",
    "L-XLSX-R5",
    "the XLSX leaf has no comments accessor on either owner yet",
);
static CHART_GAP: GapSpec = leaf_gap(
    "chart",
    "L-XLSX-R7",
    "the XLSX leaf reads chart presence only, not chart parts",
);
static SERIES_GAP: GapSpec = leaf_gap("series", "L-XLSX-R7", "chart series need the chart reader");
static AXIS_GAP: GapSpec = leaf_gap("axis", "L-XLSX-R7", "chart axes need the chart reader");
static PIVOT_GAP: GapSpec = leaf_gap(
    "pivottable",
    "M1-extended (R4x)",
    "pivot tables are not projected into nodes yet",
);
static SLICER_GAP: GapSpec = leaf_gap(
    "slicer",
    "M1-extended (R4x)",
    "slicers are not projected into nodes yet",
);
static SPARKLINE_GAP: GapSpec = leaf_gap(
    "sparkline",
    "M1-extended (R4x)",
    "sparkline groups are not projected into nodes yet",
);
static VALIDATION_GAP: GapSpec = leaf_gap(
    "dataValidation",
    "M1-extended (R4x)",
    "data validations are read by the leaf (L-XLSX-R2) but not projected into nodes yet",
);
static CF_GAP: GapSpec = leaf_gap(
    "conditionalFormatting",
    "M1-extended (R4x)",
    "conditional formats are read by the leaf (L-XLSX-R2) but not projected into nodes yet",
);
static AUTOFILTER_GAP: GapSpec = leaf_gap(
    "autofilter",
    "M1-extended (R4x)",
    "the auto-filter is read by the leaf (L-XLSX-R2) but not projected into a node yet",
);
static HYPERLINK_GAP: GapSpec = leaf_gap(
    "hyperlink",
    "M1-extended (R4x)",
    "hyperlinks are read by the leaf (L-XLSX-R2) but their targets are not resolved into nodes yet",
);
static MERGE_GAP: GapSpec = leaf_gap(
    "mergeCell",
    "M1-extended (R4x)",
    "merged ranges are not listed as nodes yet; cells report them as `merge`",
);
static PICTURE_GAP: GapSpec = leaf_gap(
    "picture",
    "L-XLSX-R7",
    "drawings are not readable on the source-backed owner yet",
);
static SHAPE_GAP: GapSpec = leaf_gap(
    "shape",
    "L-XLSX-R7",
    "drawings are not readable on the source-backed owner yet",
);
static OLE_GAP: GapSpec = leaf_gap(
    "ole",
    "M1-extended (R4x)",
    "OLE objects are not projected into nodes yet",
);
static ROW_BREAK_GAP: GapSpec = leaf_gap(
    "rowbreak",
    "M1-extended (R4x)",
    "page breaks are not readable on the source-backed owner yet",
);
static COL_BREAK_GAP: GapSpec = leaf_gap(
    "colbreak",
    "M1-extended (R4x)",
    "page breaks are not readable on the source-backed owner yet",
);
static RUN_GAP: GapSpec = leaf_gap(
    "run",
    "M1-extended (R4x)",
    "the XLSX cell model flattens rich-text runs into plain text",
);

/// One spec entry; the fields every entry spells out.
const fn element(
    kind: Kind,
    name: &'static str,
    aliases: &'static [&'static str],
    parents: &'static [Kind],
    children: &'static [Kind],
    identities: &'static [&'static str],
    paths: &'static [&'static str],
    props: &'static [PropSpec],
) -> ElementSpec<Kind> {
    ElementSpec {
        kind,
        name,
        aliases,
        parents,
        children,
        identities,
        paths,
        props,
        gap: None,
    }
}

/// A spec entry for an element Litchi cannot read yet.
const fn gap(
    kind: Kind,
    name: &'static str,
    aliases: &'static [&'static str],
    parents: &'static [Kind],
    children: &'static [Kind],
    paths: &'static [&'static str],
    gap: &'static GapSpec,
) -> ElementSpec<Kind> {
    ElementSpec {
        kind,
        name,
        aliases,
        parents,
        children,
        identities: &[],
        paths,
        props: &[],
        gap: Some(gap),
    }
}

/// The XLSX spec table.
///
/// The child lists are the query walker's tree, which is the path tree:
/// cells, rows, columns and tables sit directly under their worksheet
/// (`/Sheet1/A1`, `/Sheet1/row[3]`), so a cell query never enters a row and
/// a match's path is written from the walk. The node view lists a sheet's
/// rows and a row's cells, as OfficeCLI does, from its own children
/// (`node::Project`), independently of this table.
pub(crate) static SPEC: FormatSpec<Kind> = {
    use Kind::{
        Autofilter, Axis, Cell, Cf, Chart, ColBreak, Column, Comment, DataValidation, Hyperlink,
        MergeCell, NamedRange, Ole, Picture, PivotTable, Range, Row, RowBreak, Run, Series, Shape,
        Sheet, Slicer, Sparkline, Table, Workbook,
    };
    const SHEET: &[Kind] = &[Sheet];
    FormatSpec {
        root: Workbook,
        elements: &[
            element(
                Workbook,
                "workbook",
                &["workbook"],
                &[],
                &[Sheet, NamedRange],
                &[],
                &["/"],
                WORKBOOK_PROPS,
            ),
            element(
                Sheet,
                "sheet",
                &["sheet"],
                &[Workbook],
                &[
                    Column,
                    Row,
                    Cell,
                    Table,
                    Comment,
                    Chart,
                    PivotTable,
                    Slicer,
                    Sparkline,
                    DataValidation,
                    Cf,
                    Autofilter,
                    Hyperlink,
                    MergeCell,
                    Picture,
                    Shape,
                    Ole,
                    RowBreak,
                    ColBreak,
                ],
                &[],
                &["/<sheet>", "/sheet[N]", "<sheet>!<A1>"],
                SHEET_PROPS,
            ),
            element(
                Cell,
                "cell",
                &["cell"],
                &[Sheet, Range],
                &[Run],
                &[],
                &["/<sheet>/<A1>"],
                CELL_PROPS,
            ),
            element(
                Range,
                "range",
                &["range"],
                SHEET,
                &[Cell],
                &[],
                &["/<sheet>/<A1>:<A1>"],
                &[],
            ),
            element(
                Row,
                "row",
                &["row"],
                SHEET,
                &[],
                &[],
                &["/<sheet>/row[N]", "<sheet>!N:N"],
                ROW_PROPS,
            ),
            element(
                Column,
                "column",
                &["col", "column"],
                SHEET,
                &[],
                &[],
                &["/<sheet>/col[<letters>]", "/<sheet>/col[N]", "<sheet>!B:B"],
                COLUMN_PROPS,
            ),
            element(
                Table,
                "table",
                &["listobject", "table"],
                &[Sheet, Workbook],
                &[],
                &["name"],
                &[
                    "/<sheet>/table[N]",
                    "/<sheet>/table[@name=<name>]",
                    "/table[@name=<name>]",
                ],
                TABLE_PROPS,
            ),
            element(
                NamedRange,
                "namedrange",
                &["definedname", "namedrange"],
                &[Workbook],
                &[],
                &["name"],
                &[
                    "/namedrange[N]",
                    "/namedrange[<name>]",
                    "/namedrange[@name=<name>]",
                ],
                NAMED_RANGE_PROPS,
            ),
            gap(
                Comment,
                "comment",
                &["comment", "note"],
                SHEET,
                &[],
                &["/<sheet>/comment[N]"],
                &COMMENT_GAP,
            ),
            gap(
                Chart,
                "chart",
                &["chart"],
                SHEET,
                &[Series, Axis],
                &["/<sheet>/chart[N]"],
                &CHART_GAP,
            ),
            gap(
                Series,
                "series",
                &["series"],
                &[Chart],
                &[],
                &["/<sheet>/chart[N]/series[K]"],
                &SERIES_GAP,
            ),
            gap(
                Axis,
                "axis",
                &["axis"],
                &[Chart],
                &[],
                &["/<sheet>/chart[N]/axis[@role=<role>]"],
                &AXIS_GAP,
            ),
            gap(
                PivotTable,
                "pivottable",
                &["pivot", "pivottable"],
                SHEET,
                &[],
                &["/<sheet>/pivottable[N]"],
                &PIVOT_GAP,
            ),
            gap(
                Slicer,
                "slicer",
                &["slicer"],
                SHEET,
                &[],
                &["/<sheet>/slicer[N]"],
                &SLICER_GAP,
            ),
            gap(
                Sparkline,
                "sparkline",
                &["sparkline"],
                SHEET,
                &[],
                &["/<sheet>/sparkline[N]"],
                &SPARKLINE_GAP,
            ),
            gap(
                DataValidation,
                "dataValidation",
                &["datavalidation", "validation"],
                SHEET,
                &[],
                &["/<sheet>/dataValidation[N]"],
                &VALIDATION_GAP,
            ),
            gap(
                Cf,
                "conditionalFormatting",
                &["cf", "conditionalformatting"],
                SHEET,
                &[],
                &["/<sheet>/cf[N]"],
                &CF_GAP,
            ),
            gap(
                Autofilter,
                "autofilter",
                &["autofilter"],
                SHEET,
                &[],
                &["/<sheet>/autofilter"],
                &AUTOFILTER_GAP,
            ),
            gap(
                Hyperlink,
                "hyperlink",
                &["hyperlink"],
                SHEET,
                &[],
                &["/<sheet>/hyperlink[N]"],
                &HYPERLINK_GAP,
            ),
            gap(
                MergeCell,
                "mergeCell",
                &["merge", "mergecell", "mergedrange"],
                SHEET,
                &[],
                &["/<sheet>/mergeCell[N]"],
                &MERGE_GAP,
            ),
            gap(
                Picture,
                "picture",
                &["image", "media", "picture"],
                SHEET,
                &[],
                &["/<sheet>/picture[N]"],
                &PICTURE_GAP,
            ),
            gap(
                Shape,
                "shape",
                &["shape"],
                SHEET,
                &[],
                &["/<sheet>/shape[N]"],
                &SHAPE_GAP,
            ),
            gap(
                Ole,
                "ole",
                &["embed", "object", "ole", "oleobject"],
                SHEET,
                &[],
                &["/<sheet>/ole[N]"],
                &OLE_GAP,
            ),
            gap(
                RowBreak,
                "rowbreak",
                &["rowbreak"],
                SHEET,
                &[],
                &["/<sheet>/rowbreak[N]"],
                &ROW_BREAK_GAP,
            ),
            gap(
                ColBreak,
                "colbreak",
                &["colbreak"],
                SHEET,
                &[],
                &["/<sheet>/colbreak[N]"],
                &COL_BREAK_GAP,
            ),
            gap(
                Run,
                "run",
                &["run"],
                &[Cell],
                &[],
                &["/<sheet>/<A1>/run[N]"],
                &RUN_GAP,
            ),
        ],
        gaps: &[
            prop_gap(
                "workbook",
                "calc.mode",
                "M1-extended (R4x)",
                "calculation properties are read by the eager owner only",
            ),
            prop_gap(
                "workbook",
                "firstSheet",
                "M1-extended (R4x)",
                "the workbook views are not retained by the source-backed catalog",
            ),
            prop_gap(
                "sheet",
                "freeze",
                "M1-extended (R4x)",
                "sheet views are read by the leaf (L-XLSX-R2) but not projected yet",
            ),
            prop_gap(
                "sheet",
                "zoom",
                "M1-extended (R4x)",
                "sheet views are read by the leaf (L-XLSX-R2) but not projected yet",
            ),
            prop_gap(
                "sheet",
                "autoFilter",
                "M1-extended (R4x)",
                "the auto-filter is read by the leaf (L-XLSX-R2) but not projected yet",
            ),
            prop_gap(
                "sheet",
                "tabColor",
                "M1-extended (R4x)",
                "sheet properties are not readable on the source-backed owner yet",
            ),
            prop_gap(
                "sheet",
                "protect",
                "M1-extended (R4x)",
                "sheet protection is not readable on the source-backed owner yet",
            ),
            prop_gap(
                "sheet",
                "printArea",
                "M1-extended (R4x)",
                "print settings are not projected yet",
            ),
            prop_gap(
                "cell",
                "computedValue",
                "M4 follow-up: project computed values in cell views",
                "node views do not evaluate; `File::evaluate` and `File::recalculate` compute formula results (record 0266)",
            ),
            prop_gap(
                "cell",
                "link",
                "M1-extended (R4x)",
                "hyperlink targets are not resolved into cell properties yet",
            ),
            prop_gap(
                "cell",
                "tooltip",
                "M1-extended (R4x)",
                "hyperlink targets are not resolved into cell properties yet",
            ),
            prop_gap(
                "cell",
                "display",
                "M1-extended (R4x)",
                "hyperlink targets are not resolved into cell properties yet",
            ),
            prop_gap(
                "cell",
                "phonetic",
                "M1-extended (R4x)",
                "the XLSX cell model does not keep phonetic runs",
            ),
            prop_gap(
                "cell",
                "richtext",
                "M1-extended (R4x)",
                "the XLSX cell model flattens rich-text runs into plain text",
            ),
            prop_gap(
                "cell",
                "alt",
                "M1-extended (R4x)",
                "in-cell pictures are not read",
            ),
            prop_gap(
                "column",
                "numberformat",
                "M1-extended (R4x)",
                "column styles are not readable on the source-backed owner yet",
            ),
            prop_gap(
                "range",
                "merge",
                "M1-extended (R4x)",
                "ranges do not report whether they are merged yet",
            ),
        ],
    }
};

check_format_spec!(SPEC);

#[cfg(test)]
mod tests {
    use super::{Kind, SPEC};
    use crate::path::navigate::ElementKind;
    use crate::path::vocabulary;

    #[test]
    fn every_vocabulary_element_is_a_kind_at_its_ordinal() {
        for rule in vocabulary::WORKBOOK.elements {
            let kind = Kind::from_name(rule.name).expect("a kind per vocabulary element");
            assert_eq!(kind.name(), rule.name);
            assert_eq!(SPEC.elements[usize::from(kind.ordinal())].kind, kind);
        }
    }

    #[test]
    fn every_alias_names_its_canonical_element() {
        for element in SPEC.elements {
            for alias in element.aliases {
                if element.kind == Kind::Workbook {
                    continue;
                }
                assert_eq!(
                    vocabulary::WORKBOOK.canonical(alias),
                    Some(element.kind.name()),
                    "{alias}"
                );
            }
        }
    }
}
