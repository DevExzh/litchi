//! Apple Numbers workbooks (`.numbers`): sheets, their tables, the tables'
//! stored cells and merged ranges, read from `litchi_numbers::Package` and
//! written through its focused transactions (records 0296 and 0297).
//!
//! A Numbers sheet holds tables, and cells belong to a table: a cell's path
//! is `/sheet[N]/table[M]/A1`, or by name `/Sheet 1/Table 1/A1`; a merged
//! range is `/sheet[N]/table[M]/mergeCell[K]`. The tree holds the cells the
//! table stores; `set` also reaches an empty cell inside the table's grid.
//!
//! | Verb | Element | Leaf transaction |
//! |---|---|---|
//! | `set` | `cell` `value`/`formula`/`cachedValue` | `edit_table_cells` (one per table) |
//! | `set` | `cell` `numberFormat` | `edit_table_cell_{number,percentage}_format` |
//! | `set` | `cell` `comment` (`null` clears) | `set_table_cell_comment`, `clear_table_cell_comment` (record 0318) |
//! | `add` | `comment` under a cell (`text`) | `set_table_cell_comment` |
//! | `remove` | `<cell>/comment` | `clear_table_cell_comment` |
//! | `set` | `sheet`/`table` `name` | `edit_names` (one for every rename) |
//! | `add` | `mergeCell` | `edit_table_merges` |
//! | `remove` | `cell` (clears) / `mergeCell` | `edit_table_cells` / `edit_table_merges` |
//! | `move` | `sheet` | `edit_sheet_order` |
//! | `move` | `table` into another sheet | `edit_table_relocation` |
//! | `move` | `table` within its sheet | `edit_table_order` |
//! | `add`/`remove` | `sheet` | `edit_sheet_lifecycle` (+ `edit_sheet_order` for a position) |
//! | `add`/`remove` | `table` | `edit_table_lifecycle` (+ `edit_table_order` for a position) |
//! | `add`/`remove` | `row`/`col` of a table | `edit_table_topology` (record 0309) |
//! | `swap` | two `sheet`s | two `edit_sheet_order` moves (record 0347) |
//! | `swap` | two `table`s of one sheet | two `edit_table_order` moves |
//! | `swap` | two `table`s of two sheets | two `edit_table_relocation`s, then two `edit_table_order` moves (two tables of one name pass through a temporary `edit_names` rename) |
//!
//! Rows and columns are addressed by position, `/sheet[N]/table[M]/row[K]`
//! and `col[K]`; the tree does not list them (a table's `rowCount` and
//! `columnCount` give their extent).

pub(crate) mod coalesce;

use std::collections::BTreeMap;
use std::fmt;

use litchi_core::selector::{CellRef, Path, Predicate, Step};
use litchi_core::{FileFormat, Position};
use litchi_numbers::cell::Value as CellValue;
use litchi_numbers::cell::data_format::{
    DecimalPlaces, NegativeStyle, Number, Percentage, ThousandsSeparator,
};
use litchi_numbers::cell::text_style::CharacterProperties;
use litchi_numbers::formula::{CachedValue, Expression};
use litchi_numbers::sheet::chart::font::{
    Property as ChartFontProperty, PropertyKind as ChartFontPropertyKind,
};
use litchi_numbers::table::cells::{Change, Edit, Input};
use litchi_numbers::table::merge::Region;
use litchi_numbers::table::topology::{ColumnDeletion, ColumnInsertion, RowDeletion, RowInsertion};
use litchi_numbers::table::{Area, At};
use litchi_numbers::{Package, SheetSelector, TableSelector};

use super::tree::{count, text};
use super::write::{
    Chain, Commit, LeafPatch, bound, destination, element_refusal, finish, key_refusal,
    leaf_refusal, node_at, refusal, resolve,
};
use super::{
    Flavor, Iwork, LeafError, Node, Owner, TreeBuilder, clearable_writable, element, writable,
};
use crate::Result;
use crate::error::Code;
use crate::new_node::NewNode;
use crate::node::spec::{Access, FormatSpec, GapSpec, PropSpec, Subject};
use crate::node::{PathStep, Value};
use crate::path::{IntoPath, PathArg};
use crate::props::{PropValue, Props};
use crate::target::TargetArg;
use crate::transaction::{LimitKind, SemanticLimits};
use crate::value::{Scalar, ValueType};

kinds! {
    /// The element kinds of a Numbers workbook.
    Kind: &crate::path::vocabulary::NUMBERS, root Workbook {
        Workbook => "workbook",
        Sheet => "sheet",
        Table => "table",
        Cell => "cell",
        MergeCell => "mergeCell",
        Row => "row",
        Column => "col",
        Chart => "chart",
    }
}

/// A property `set` writes but the tree does not read back (the leaf reads
/// it per cell only).
const fn set_only(name: &'static str, ty: ValueType) -> PropSpec {
    let mut spec = PropSpec::new(name, ty);
    spec.access = Access {
        get: false,
        set: true,
        add: false,
    };
    spec
}

const WORKBOOK_PROPS: &[PropSpec] = &[PropSpec::new("sheetCount", ValueType::Int)];

/// A property `set` writes and `add` accepts.
const fn addable_writable(name: &'static str, ty: ValueType) -> PropSpec {
    let mut spec = writable(name, ty);
    spec.access.add = true;
    spec
}

/// A property the tree reads and `add` accepts.
const fn addable_read(name: &'static str, ty: ValueType) -> PropSpec {
    let mut spec = PropSpec::new(name, ty);
    spec.access.add = true;
    spec
}

pub(crate) const SHEET_PROPS: &[PropSpec] = &[
    addable_writable("name", ValueType::Text),
    PropSpec::new("tableCount", ValueType::Int),
];

pub(crate) const TABLE_PROPS: &[PropSpec] = &[
    addable_writable("name", ValueType::Text),
    addable_read("rowCount", ValueType::Int),
    addable_read("columnCount", ValueType::Int),
    PropSpec::new("mergeCount", ValueType::Int),
];

pub(crate) const CELL_PROPS: &[PropSpec] = &[
    PropSpec::new("type", ValueType::Text),
    writable("value", ValueType::Scalar),
    writable("formula", ValueType::Text),
    set_only("cachedValue", ValueType::Scalar),
    clearable_writable("comment", ValueType::Text),
    PropSpec::new("merge", ValueType::Text),
    PropSpec::new("mergeAnchor", ValueType::Bool),
    set_only("numberFormat", ValueType::Text),
    set_only("font.bold", ValueType::Bool),
    set_only("font.italic", ValueType::Bool),
    set_only("font.underline", ValueType::Bool),
    set_only("font.strike", ValueType::Bool),
    set_only("font.size", FONT_SIZE),
    set_only("font.name", ValueType::Text),
    set_only("font.color", FONT_COLOR),
];

const CHART_PROPS: &[PropSpec] = &[clearable_writable("font.size", FONT_SIZE)];

/// A font size in points.
const FONT_SIZE: ValueType = ValueType::Length(crate::value::LengthSpec::FONT_SIZE);
/// A font color.
const FONT_COLOR: ValueType = ValueType::Color(crate::value::ColorContext::Spreadsheet);

pub(crate) const MERGE_PROPS: &[PropSpec] = &[{
    let mut spec = PropSpec::new("ref", ValueType::Text);
    spec.access.add = true;
    spec
}];

const FOLLOW_UP: &str = "record 0297 follow-ups (IWORK_API_REQUESTS.md)";

static GAPS: [GapSpec; 1] = [GapSpec {
    subject: Subject::Operation {
        verb: "set",
        element: "cell",
        context: "cross-table formula references",
    },
    code: Code::UnsupportedCapability,
    follow_up: FOLLOW_UP,
    reason: "formula text is parsed with local A1 references and the authorable function subset; references to other tables need source-bound table handles",
}];

static OPERATION_GAPS: [GapSpec; 5] = [
    GapSpec {
        subject: Subject::Operation {
            verb: "add",
            element: "table",
            context: "a workbook without a table",
        },
        code: Code::UnsupportedCapability,
        follow_up: "record 0301 follow-ups (IWORK_API_REQUESTS.md request 5)",
        reason: "a new table copies the structure (styles, sizes, storage layout) of the workbook's first table; a workbook with no table has no template yet",
    },
    GapSpec {
        subject: Subject::Operation {
            verb: "add",
            element: "row",
            context: "formulas the coordinate shift refuses, merges",
        },
        code: Code::UnsupportedCapability,
        follow_up: "record 0329 follow-ups (IWORK_API_REQUESTS.md request 6)",
        reason: "rows and columns move every formula reference into the table, the table's own formulas and other tables' (record 0318); a removed row or column a surviving formula reads turns into #REF! (ranges shrink), and a shared formula whose cells diverge is split (record 0329); refused: a chart or other non-table owner reading the table, volatile, spanning, error or spill dependency state, a removal whose cells the cell writer cannot clear under a range that reads them, merged cells, filters, categories, pivots and (for columns) sort rules, and removing a hidden or bordered row or column",
    },
    GapSpec {
        subject: Subject::Operation {
            verb: "set",
            element: "row",
            context: "height and visibility",
        },
        code: Code::UnsupportedCapability,
        follow_up: "record 0309 follow-ups",
        reason: "rows and columns are structural only here: their sizes and hidden states are not wired to the leaf's dimension transaction",
    },
    GapSpec {
        subject: Subject::Operation {
            verb: "move",
            element: "row",
            context: "rows and columns",
        },
        code: Code::UnsupportedCapability,
        follow_up: "record 0347 F1 (a row and column move in `table::topology`)",
        reason: "the leaf topology transaction inserts and removes rows and columns but does not reorder them; reordering must carry storage, headers, UIDs, borders, merges and every formula coordinate that reads the table",
    },
    GapSpec {
        subject: Subject::Operation {
            verb: "swap",
            element: "row",
            context: "rows and columns",
        },
        code: Code::UnsupportedCapability,
        follow_up: "record 0347 F1 (a row and column move in `table::topology`)",
        reason: "a swap is two moves, and the leaf topology transaction does not reorder rows or columns; sheets and tables swap",
    },
];

/// The Numbers spec table.
pub(crate) static SPEC: FormatSpec<Kind> = FormatSpec {
    root: Kind::Workbook,
    elements: &[
        element(
            Kind::Workbook,
            "workbook",
            &["workbook"],
            &[],
            &[Kind::Sheet],
            &["/"],
            WORKBOOK_PROPS,
        ),
        element(
            Kind::Sheet,
            "sheet",
            &["sheet"],
            &[Kind::Workbook],
            &[Kind::Table, Kind::Chart],
            &["/sheet[N]", "/<sheet name>"],
            SHEET_PROPS,
        ),
        element(
            Kind::Table,
            "table",
            &["table"],
            &[Kind::Sheet],
            &[Kind::Cell, Kind::MergeCell, Kind::Row, Kind::Column],
            &["/sheet[N]/table[N]", "/<sheet name>/<table name>"],
            TABLE_PROPS,
        ),
        element(
            Kind::Cell,
            "cell",
            &["cell"],
            &[Kind::Table],
            &[],
            &["/sheet[N]/table[N]/A1"],
            CELL_PROPS,
        ),
        element(
            Kind::MergeCell,
            "mergeCell",
            &["mergecell"],
            &[Kind::Table],
            &[],
            &["/sheet[N]/table[N]/mergeCell[N]"],
            MERGE_PROPS,
        ),
        element(
            Kind::Row,
            "row",
            &["row"],
            &[Kind::Table],
            &[],
            &["/sheet[N]/table[N]/row[N]"],
            &[],
        ),
        element(
            Kind::Column,
            "col",
            &["col", "column"],
            &[Kind::Table],
            &[],
            &["/sheet[N]/table[N]/col[N]"],
            &[],
        ),
        element(
            Kind::Chart,
            "chart",
            &["chart"],
            &[Kind::Sheet],
            &[],
            &["/sheet[N]/chart[N]"],
            CHART_PROPS,
        ),
    ],
    gaps: &GAPS,
};

crate::node::spec::check_format_spec!(SPEC);

/// The Numbers flavor.
pub(crate) struct Numbers;

impl Flavor for Numbers {
    type Kind = Kind;
    const FORMAT: FileFormat = FileFormat::Numbers;
    const SPEC: &'static FormatSpec<Kind> = &SPEC;
    const TOP: &'static [Kind] = &[Kind::Workbook];
}

/// The elements each write verb accepts.
pub(crate) static WRITES: super::Writes = super::Writes {
    settable: &["cell", "sheet", "table", "chart"],
    addable: &["col", "mergeCell", "row", "sheet", "table"],
    removable: &["cell", "col", "mergeCell", "row", "sheet", "table"],
    movable: &["sheet", "table"],
    gaps: &OPERATION_GAPS,
};

// --------------------------------------------------------------- leaf errors

/// A failed Numbers transaction other than the table-cell transaction.
#[derive(Debug)]
pub enum EditError {
    /// A sheet or table rename.
    Names(litchi_numbers::names::Error),
    /// A sheet move.
    Order(litchi_numbers::sheet::order::Error),
    /// A table relocation into another sheet.
    Relocation(litchi_numbers::table::relocation::Error),
    /// A merge or unmerge.
    Merges(litchi_numbers::table::merge::transaction::Error),
    /// A decimal display format.
    NumberFormat(litchi_numbers::cell::data_format::number::transaction::Error),
    /// A percentage display format.
    PercentageFormat(litchi_numbers::cell::data_format::percentage::transaction::Error),
    /// Formula text that does not parse into an authorable expression.
    Formula(litchi_numbers::formula::Error),
    /// A fresh workbook.
    Create(litchi_numbers::create::Error),
    /// A sheet addition or removal.
    SheetLifecycle(litchi_numbers::sheet::lifecycle::Error),
    /// A table addition or removal.
    TableLifecycle(litchi_numbers::table::lifecycle::Error),
    /// A table move within its sheet.
    TableOrder(litchi_numbers::table::order::Error),
    /// A row or column insertion or removal.
    Topology(litchi_numbers::table::topology::Error),
    /// Cell character properties.
    TextStyle(litchi_numbers::cell::text_style::Error),
    /// A sheet-chart point-size transaction (D7-A).
    ChartFont(litchi_numbers::sheet::chart::font::Error),
    /// A cell comment (record 0318).
    Comment(litchi_numbers::cell::comment::Error),
}

impl fmt::Display for EditError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Names(error) => error.fmt(formatter),
            Self::Order(error) => error.fmt(formatter),
            Self::Relocation(error) => error.fmt(formatter),
            Self::Merges(error) => error.fmt(formatter),
            Self::NumberFormat(error) => error.fmt(formatter),
            Self::PercentageFormat(error) => error.fmt(formatter),
            Self::Formula(error) => error.fmt(formatter),
            Self::Create(error) => error.fmt(formatter),
            Self::SheetLifecycle(error) => error.fmt(formatter),
            Self::TableLifecycle(error) => error.fmt(formatter),
            Self::TableOrder(error) => error.fmt(formatter),
            Self::Topology(error) => error.fmt(formatter),
            Self::TextStyle(error) => error.fmt(formatter),
            Self::ChartFont(error) => error.fmt(formatter),
            Self::Comment(error) => error.fmt(formatter),
        }
    }
}

impl EditError {
    pub(crate) fn source_error(&self) -> Option<&(dyn std::error::Error + 'static)> {
        Some(match self {
            Self::Names(error) => error,
            Self::Order(error) => error,
            Self::Relocation(error) => error,
            Self::Merges(error) => error,
            Self::NumberFormat(error) => error,
            Self::PercentageFormat(error) => error,
            Self::Formula(error) => error,
            Self::Create(error) => error,
            Self::SheetLifecycle(error) => error,
            Self::TableLifecycle(error) => error,
            Self::TableOrder(error) => error,
            Self::Topology(error) => error,
            Self::TextStyle(error) => error,
            Self::ChartFont(error) => error,
            Self::Comment(error) => error,
        })
    }

    /// The unified code: a refusal reports its kind's code (record 0351
    /// §2.4, record 0361), a configured budget `limit_exceeded`, a malformed
    /// source or a patch for another base `invalid_input`, the rest
    /// `internal_error`.
    pub(crate) fn code(&self) -> Code {
        macro_rules! family {
            ($error:expr, $module:path) => {{
                use $module as E;
                match $error {
                    E::Refused(refusal) => crate::error::refusal_code(refusal.kind()),
                    E::Limit { .. } => Code::LimitExceeded,
                    E::InvalidSource { .. } | E::PatchConflict => Code::InvalidInput,
                    _ => Code::InternalError,
                }
            }};
        }
        match self {
            Self::Names(error) => family!(error, litchi_numbers::names::Error),
            Self::Order(error) => family!(error, litchi_numbers::sheet::order::Error),
            Self::Relocation(error) => family!(error, litchi_numbers::table::relocation::Error),
            Self::Merges(error) => family!(error, litchi_numbers::table::merge::transaction::Error),
            Self::NumberFormat(error) => {
                family!(
                    error,
                    litchi_numbers::cell::data_format::number::transaction::Error
                )
            },
            Self::PercentageFormat(error) => {
                family!(
                    error,
                    litchi_numbers::cell::data_format::percentage::transaction::Error
                )
            },
            Self::Formula(error) => match error {
                litchi_numbers::formula::Error::Limit { .. } => Code::LimitExceeded,
                litchi_numbers::formula::Error::UnsupportedFunction => Code::UnsupportedCapability,
                _ => Code::InvalidValue,
            },
            Self::Create(error) => match error {
                litchi_numbers::create::Error::MissingIdentity
                | litchi_numbers::create::Error::IdentityNotDistinct => Code::InvalidArgument,
                _ => Code::InternalError,
            },
            // Removing the final sheet or table axis is a caller argument
            // error, as for PPTX `SlideRemovalRefusal::FinalSlide`.
            Self::SheetLifecycle(litchi_numbers::sheet::lifecycle::Error::Refused(
                litchi_numbers::sheet::lifecycle::Refusal::LastSheet,
            ))
            | Self::Topology(litchi_numbers::table::topology::Error::Refused(
                litchi_numbers::table::topology::Refusal::LastAxis,
            )) => Code::InvalidArgument,
            Self::SheetLifecycle(error) => family!(error, litchi_numbers::sheet::lifecycle::Error),
            Self::TableLifecycle(error) => family!(error, litchi_numbers::table::lifecycle::Error),
            Self::TableOrder(error) => family!(error, litchi_numbers::table::order::Error),
            Self::Topology(error) => match error {
                // A cell clear inside a topology edit reports the cell
                // transaction's own code.
                litchi_numbers::table::topology::Error::Refused(
                    litchi_numbers::table::topology::Refusal::CellClear { error },
                ) => cells_error(*error).code(),
                _ => family!(error, litchi_numbers::table::topology::Error),
            },
            Self::TextStyle(error) => family!(error, litchi_numbers::cell::text_style::Error),
            Self::ChartFont(error) => family!(error, litchi_numbers::sheet::chart::font::Error),
            Self::Comment(error) => family!(error, litchi_numbers::cell::comment::Error),
        }
    }
}

/// The reversible proof of one Numbers transaction other than cells.
pub(crate) enum Patch {
    /// Renames.
    Names(litchi_numbers::names::Patch),
    /// A sheet move.
    Order(litchi_numbers::sheet::order::Patch),
    /// A table relocation.
    Relocation(litchi_numbers::table::relocation::Patch),
    /// Merges and unmerges.
    Merges(litchi_numbers::table::merge::transaction::Patch),
    /// A decimal display format.
    NumberFormat(litchi_numbers::cell::data_format::number::transaction::Patch),
    /// A percentage display format.
    PercentageFormat(litchi_numbers::cell::data_format::percentage::transaction::Patch),
    /// A sheet addition or removal.
    SheetLifecycle(litchi_numbers::sheet::lifecycle::Patch),
    /// A table addition or removal.
    TableLifecycle(litchi_numbers::table::lifecycle::Patch),
    /// A table move within its sheet.
    TableOrder(litchi_numbers::table::order::Patch),
    /// Row and column insertions and removals.
    Topology(litchi_numbers::table::topology::Patch),
    /// Cell character properties.
    TextStyle(litchi_numbers::cell::text_style::Patch),
    /// A sheet-chart point-size set (D7-A).
    ChartFontSize(litchi_numbers::sheet::chart::font::Patch),
    /// A cell comment set or cleared.
    Comment(litchi_numbers::cell::comment::Patch),
}

impl Patch {
    pub(crate) fn is_noop(&self) -> bool {
        match self {
            Self::Names(patch) => patch.is_noop(),
            Self::Order(patch) => patch.is_noop(),
            Self::Relocation(patch) => patch.is_noop(),
            Self::Merges(patch) => patch.is_noop(),
            Self::NumberFormat(patch) => patch.is_noop(),
            Self::PercentageFormat(patch) => patch.is_noop(),
            Self::SheetLifecycle(patch) => patch.is_noop(),
            Self::TableLifecycle(patch) => patch.is_noop(),
            Self::TableOrder(patch) => patch.is_noop(),
            Self::Topology(patch) => patch.is_noop(),
            Self::TextStyle(patch) => patch.is_noop(),
            Self::ChartFontSize(patch) => patch.is_noop(),
            Self::Comment(patch) => patch.is_noop(),
        }
    }
}

fn edit_error(error: EditError) -> crate::Error {
    leaf_refusal(LeafError::NumbersEdit(error), FileFormat::Numbers)
}

fn cells_error(error: litchi_numbers::table::cells::Error) -> crate::Error {
    leaf_refusal(LeafError::NumbersCells(error), FileFormat::Numbers)
}

// --------------------------------------------------------------------- reads

/// The `type` token, typed value and display text of a stored cell value.
fn cell_value(value: &CellValue) -> (&'static str, Option<Value<'static>>, Option<String>) {
    match value {
        CellValue::Empty => ("empty", None, None),
        CellValue::Text(text) => (
            "text",
            Some(Value::Text(text.clone().into())),
            Some(text.clone()),
        ),
        CellValue::Number(number) => number_value("number", number.get()),
        CellValue::Boolean(value) => (
            "boolean",
            Some(Value::Bool(*value)),
            Some(if *value { "TRUE" } else { "FALSE" }.to_owned()),
        ),
        CellValue::Date(seconds) => number_value("date", seconds.get()),
        CellValue::Duration(seconds) => number_value("duration", seconds.get()),
        CellValue::Formula(formula) => ("formula", None, Some(formula.clone())),
        CellValue::Error(error) => (
            "error",
            Some(Value::Text(error.clone().into())),
            Some(error.clone()),
        ),
    }
}

fn number_value(
    kind: &'static str,
    number: f64,
) -> (&'static str, Option<Value<'static>>, Option<String>) {
    let value = Value::number(number);
    let display = value.as_ref().map(ToString::to_string);
    (kind, value, display)
}

/// The A1 range of a merged region (`B2:C3`).
fn region_ref(region: Region) -> String {
    let start = region.start();
    let end = region.end();
    format!("{start}:{end}")
}

/// Build the eager tree of `package`: the workbook, its sheets and their
/// tables, reading each table's merged regions. A table's cells and merged
/// ranges are a lazily built part ([`Cells`], record 0305).
///
/// # Errors
///
/// The leaf's typed merge-reader error when a table's merge storage cannot
/// be read.
pub(crate) fn build(package: &Package) -> Result<(super::Tree, Box<[super::Part]>)> {
    let merges_error = |error| edit_error(EditError::Merges(error));
    let mut tree = TreeBuilder::new(SPEC.elements, Numbers::SINGLETONS);
    let sheets = package.sheets();
    tree.root().prop("sheetCount", count(sheets.len())).done();
    let mut parts = Vec::new();
    for (sheet_index, sheet) in sheets.iter().enumerate() {
        let sheet_item = tree
            .push(super::tree::ROOT, Kind::Sheet)
            .prop("name", text(Some(sheet.name())))
            .prop("tableCount", count(sheet.table_count()))
            .label(Some(sheet.name()))
            .leaf([sheet_index, 0, 0, 0])
            .done();
        if let Ok(drawables) = package.sheet_drawable_list(SheetSelector::index(sheet_index)) {
            for drawable in drawables
                .iter()
                .filter(|drawable| drawable.kind() == litchi_numbers::sheet::drawable::Kind::Chart)
            {
                let chart = drawable.kind_selector().as_index();
                tree.push(sheet_item, Kind::Chart)
                    .leaf([sheet_index, chart, 0, 0])
                    .done();
            }
        }
        for (table_index, table) in sheet.tables().enumerate() {
            let regions = package
                .table_merges(
                    SheetSelector::index(sheet_index),
                    TableSelector::index(table_index),
                )
                .map_err(merges_error)?;
            let table_item = tree
                .push(sheet_item, Kind::Table)
                .prop("name", text(Some(table.name())))
                .prop("rowCount", count(table.row_count() as usize))
                .prop("columnCount", count(table.column_count() as usize))
                .prop("mergeCount", count(regions.len()))
                .label(Some(table.name()))
                .leaf([sheet_index, table_index, 0, 0])
                .part(u32::try_from(parts.len() + 1).unwrap_or(u32::MAX))
                .done();
            parts.push(super::Part::Table(Cells::new(
                [sheet_index, table_index],
                tree.path_of(table_item),
                table.cell_count(),
                regions,
                [table.row_count(), table.column_count()],
            )));
        }
    }
    Ok((tree.build(), parts.into_boxed_slice()))
}

/// Build only one sheet's chart structural projection used by the canonical
/// chart path fast route. The ordinary workbook tree remains unchanged and
/// unrelated sheet drawable inventories are never inspected.
pub(crate) fn build_charts(package: &Package, sheet_index: usize) -> Result<super::Tree> {
    let mut tree = TreeBuilder::new(SPEC.elements, Numbers::SINGLETONS);
    let sheets = package.sheets();
    tree.root().prop("sheetCount", count(sheets.len())).done();
    let Some(sheet) = sheets.get(sheet_index) else {
        return Ok(tree.build());
    };
    let sheet_item = tree
        .push_step(
            super::tree::ROOT,
            Kind::Sheet,
            PathStep::at("sheet", Position::new(sheet_index)),
        )
        .prop("name", text(Some(sheet.name())))
        .prop("tableCount", count(sheet.table_count()))
        .label(Some(sheet.name()))
        .leaf([sheet_index, 0, 0, 0])
        .done();
    let drawables = package
        .sheet_drawable_list(SheetSelector::index(sheet_index))
        .map_err(|error| leaf_refusal(LeafError::NumbersDrawables(error), FileFormat::Numbers))?;
    for drawable in drawables
        .iter()
        .filter(|drawable| drawable.kind() == litchi_numbers::sheet::drawable::Kind::Chart)
    {
        tree.push(sheet_item, Kind::Chart)
            .leaf([sheet_index, drawable.kind_selector().as_index(), 0, 0])
            .done();
    }
    Ok(tree.build())
}

/// The cells, merged ranges, rows and columns of one table, each tree item
/// materialized the first time a read reaches it (record 0305): slot `i`
/// below the cell count is the table's `i`-th stored cell in row-major
/// order, the slots past it its merged ranges, then its rows and its
/// columns (record 0309; `add` and `remove` address them as `row[N]` and
/// `col[N]`, and their items are built together on the first read of
/// one). A cell step looks its cell up in the leaf's sorted cell list, so
/// `get` of one cell builds one item.
pub(crate) struct Cells {
    sheet: usize,
    table: usize,
    /// The table's path.
    path: Box<[PathStep<'static>]>,
    regions: Box<[Region]>,
    cells: Box<[std::sync::OnceLock<Box<super::tree::Item>>]>,
    merges: Box<[std::sync::OnceLock<Box<super::tree::Item>>]>,
    /// The row and column counts.
    axes: [u32; 2],
    /// The row items, then the column items.
    axis_items: std::sync::OnceLock<Box<[super::tree::Item]>>,
    /// The table's cell comments, read on the first cell item (record
    /// 0318): `(position, text)` in row-major order.
    comments: std::sync::OnceLock<Box<[(At, String)]>>,
}

impl Cells {
    fn new(
        [sheet, table]: [usize; 2],
        path: Box<[PathStep<'static>]>,
        cells: usize,
        regions: Vec<Region>,
        axes: [u32; 2],
    ) -> Self {
        Self {
            sheet,
            table,
            path,
            cells: (0..cells).map(|_| std::sync::OnceLock::new()).collect(),
            merges: regions.iter().map(|_| std::sync::OnceLock::new()).collect(),
            regions: regions.into_boxed_slice(),
            axes,
            axis_items: std::sync::OnceLock::new(),
            comments: std::sync::OnceLock::new(),
        }
    }

    /// The comment text of the cell at `position`, if any. A table whose
    /// comment list is empty costs one model and list decode; a comment
    /// graph the leaf cannot read lists no comment.
    fn comment(&self, package: &Package, position: At) -> Option<&str> {
        let comments = self.comments.get_or_init(|| {
            package
                .table_cell_comments(
                    SheetSelector::index(self.sheet),
                    TableSelector::index(self.table),
                )
                .map(|comments| {
                    comments
                        .into_iter()
                        .map(|(at, comment)| (at, comment.text().to_owned()))
                        .collect()
                })
                .unwrap_or_default()
        });
        comments
            .binary_search_by(|(at, _)| at.cmp(&position))
            .ok()
            .map(|index| comments[index].1.as_str())
    }

    /// The first row slot.
    fn axis_start(&self) -> u32 {
        u32::try_from(self.cells.len() + self.merges.len()).unwrap_or(u32::MAX)
    }

    fn cell_count(&self) -> u32 {
        u32::try_from(self.cells.len()).unwrap_or(u32::MAX)
    }

    #[cfg(test)]
    pub(crate) fn state(&self) -> String {
        let built = self
            .cells
            .iter()
            .chain(self.merges.iter())
            .filter(|item| item.get().is_some())
            .count();
        format!("cells {built}/{}", self.cells.len() + self.merges.len())
    }

    /// The slot count: the stored cells, the merged ranges, the rows and
    /// the columns.
    pub(crate) fn len(&self) -> u32 {
        let [rows, columns] = self.axes;
        self.axis_start()
            .saturating_add(rows)
            .saturating_add(columns)
    }

    /// The kind ordinal of slot `slot`.
    pub(crate) fn kind(&self, slot: u32) -> u8 {
        let start = self.axis_start();
        if slot < self.cell_count() {
            Kind::Cell as u8
        } else if slot < start {
            Kind::MergeCell as u8
        } else if slot - start < self.axes[0] {
            Kind::Row as u8
        } else {
            Kind::Column as u8
        }
    }

    /// The position of slot `slot` among the table's children of its kind.
    pub(crate) fn ordinal(&self, slot: u32) -> u32 {
        let cells = self.cell_count();
        let start = self.axis_start();
        if slot < cells {
            slot
        } else if slot < start {
            slot - cells
        } else if slot - start < self.axes[0] {
            slot - start
        } else {
            slot - start - self.axes[0]
        }
    }

    /// The leaf table.
    fn table<'p>(&self, package: &'p Package) -> Option<&'p litchi_numbers::Table> {
        package.sheets().get(self.sheet)?.tables().nth(self.table)
    }

    fn package(file: &Iwork) -> Option<&Package> {
        match file.owner() {
            Owner::Numbers(package) => Some(package),
            #[allow(unreachable_patterns, reason = "the owner arms depend on features")]
            _ => None,
        }
    }

    /// The item of slot `slot`, materialized on the first call.
    pub(crate) fn item<'f>(&'f self, file: &'f Iwork, slot: u32) -> &'f super::tree::Item {
        let cells = self.cell_count();
        let start = self.axis_start();
        if slot < cells {
            self.cells[slot as usize].get_or_init(|| Box::new(self.cell(file, slot)))
        } else if slot < start {
            let merge = (slot - cells) as usize;
            self.merges[merge].get_or_init(|| Box::new(self.merge(merge)))
        } else {
            &self.axis_items.get_or_init(|| self.axis_items())[(slot - start) as usize]
        }
    }

    /// The row items, then the column items.
    fn axis_items(&self) -> Box<[super::tree::Item]> {
        let [rows, columns] = self.axes;
        let axis = |kind: Kind, index: u32| {
            let mut path = Vec::with_capacity(self.path.len() + 1);
            path.extend_from_slice(&self.path);
            path.push(PathStep::at(
                crate::path::navigate::ElementKind::name(kind),
                Position::new(index as usize),
            ));
            let mut item = super::tree::Item::detached(
                &SPEC.elements[kind as usize],
                path.into_boxed_slice(),
                index,
            );
            let (row, column) = if matches!(kind, Kind::Row) {
                (index as usize, 0)
            } else {
                (0, index as usize)
            };
            item.draft::<Kind>()
                .leaf([self.sheet, self.table, row, column])
                .done();
            item
        };
        (0..rows)
            .map(|row| axis(Kind::Row, row))
            .chain((0..columns).map(|column| axis(Kind::Column, column)))
            .collect()
    }

    fn cell(&self, file: &Iwork, slot: u32) -> super::tree::Item {
        let spec = &SPEC.elements[Kind::Cell as usize];
        let cell = Self::package(file)
            .and_then(|package| self.table(package))
            .and_then(|table| table.iter_cells().nth(slot as usize));
        let Some(cell) = cell else {
            // The handle's package is immutable: its table lists every
            // slot `Cells::new` counted.
            return super::tree::Item::detached(spec, self.path.clone(), slot);
        };
        let position = cell.position();
        let comment = Self::package(file).and_then(|package| self.comment(package, position));
        let mut path = Vec::with_capacity(self.path.len() + 1);
        path.extend_from_slice(&self.path);
        path.push(PathStep::cell(position.column(), position.row()));
        let mut item = super::tree::Item::detached(spec, path.into_boxed_slice(), slot);
        let (kind, value, display) = cell_value(cell.value());
        let formula = match cell.value() {
            CellValue::Formula(formula) => text(Some(formula)),
            _ => None,
        };
        let region = self.regions.iter().find(|region| {
            let (row, column) = (position.row(), position.column());
            row >= region.row()
                && row < region.row() + region.row_count()
                && column >= region.column()
                && column < region.column() + region.column_count()
        });
        let mut draft = item
            .draft::<Kind>()
            .prop("type", text(Some(kind)))
            .prop("value", value)
            .prop("formula", formula)
            .prop("comment", text(comment))
            .prop("merge", region.map(|region| region_ref(*region).into()))
            .prop(
                "mergeAnchor",
                region.map(|region| Value::Bool(region.start() == position)),
            )
            .leaf([
                self.sheet,
                self.table,
                position.row() as usize,
                position.column() as usize,
            ]);
        if let Some(display) = display {
            draft = draft.text(display);
        }
        draft.done();
        item
    }

    fn merge(&self, merge: usize) -> super::tree::Item {
        let spec = &SPEC.elements[Kind::MergeCell as usize];
        let region = self.regions[merge];
        let mut path = Vec::with_capacity(self.path.len() + 1);
        path.extend_from_slice(&self.path);
        path.push(PathStep::at(
            crate::path::navigate::ElementKind::name(Kind::MergeCell),
            Position::new(merge),
        ));
        let ordinal = u32::try_from(merge).unwrap_or(u32::MAX);
        let mut item = super::tree::Item::detached(spec, path.into_boxed_slice(), ordinal);
        item.draft::<Kind>()
            .prop("ref", text(Some(&region_ref(region))))
            .leaf([
                self.sheet,
                self.table,
                region.row() as usize,
                region.column() as usize,
            ])
            .done();
        item
    }

    /// The slot of the stored cell at `row`, `column` when `kind` is the
    /// cell kind (`Some(None)`: the table stores no such cell); `None` for
    /// another kind, which the caller scans for.
    #[allow(
        clippy::option_option,
        reason = "`None` asks the caller to scan; `Some(None)` is a miss"
    )]
    pub(crate) fn locate(
        &self,
        file: &Iwork,
        kind: u8,
        row: u32,
        column: u32,
    ) -> Option<Option<u32>> {
        if kind != Kind::Cell as u8 {
            return None;
        }
        let table = Self::package(file).and_then(|package| self.table(package))?;
        let wanted = At::new(row, column);
        // The leaf keeps a table's stored cells sorted by position, without
        // duplicates (its own lookups binary-search them).
        let (mut low, mut high) = (0_usize, table.cell_count());
        while low < high {
            let middle = low + (high - low) / 2;
            let position = table.iter_cells().nth(middle)?.position();
            match position.cmp(&wanted) {
                std::cmp::Ordering::Less => low = middle + 1,
                std::cmp::Ordering::Greater => high = middle,
                std::cmp::Ordering::Equal => {
                    return Some(u32::try_from(middle).ok());
                },
            }
        }
        Some(None)
    }
}

// -------------------------------------------------------------------- targets

/// One cell a write addresses: sheet, table and zero-based coordinates.
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord)]
pub(crate) struct Address {
    pub(crate) sheet: usize,
    pub(crate) table: usize,
    pub(crate) row: u32,
    pub(crate) column: u32,
}

impl Address {
    const fn position(self) -> At {
        At::new(self.row, self.column)
    }
}

/// The nodes `target` resolves to, or the empty in-grid cell its path names
/// (`Err` keeps the resolution error when the path names no such cell).
enum Targets<'f> {
    Nodes(Vec<Node<'f>>),
    EmptyCell(Address),
}

fn targets<'f>(
    file: &'f Iwork,
    target: &TargetArg<'_>,
    limits: &SemanticLimits,
) -> Result<Targets<'f>> {
    match resolve::<Numbers>(file, target, limits.max_expanded_targets()) {
        Ok(nodes) => Ok(Targets::Nodes(nodes)),
        Err(error) if error.code() == Code::PathNotFound => {
            let Some(path) = target.path() else {
                return Err(error);
            };
            empty_cell(file, path.path())
                .map_or(Err(error), |address| Ok(Targets::EmptyCell(address)))
        },
        Err(error) => Err(error),
    }
}

/// The address of a cell node.
fn cell_address(node: &Node<'_>) -> Address {
    let [sheet, table, row, column] = node.item().leaf();
    Address {
        sheet: sheet as usize,
        table: table as usize,
        row,
        column,
    }
}

/// The empty cell `path` names: a table path and an A1 step inside the
/// table's grid.
pub(crate) fn empty_cell(file: &Iwork, path: &Path<'_>) -> Option<Address> {
    let (last, parent) = path.steps().split_last()?;
    if !last.predicates().is_empty() {
        return None;
    }
    let cell = CellRef::parse(last.name())?;
    let parent = Path::from_steps(parent.iter().cloned());
    let parent = parent.into_path().ok()?;
    let table = node_at::<Numbers>(file, &parent).ok()?;
    if table.item().type_name() != "table" {
        return None;
    }
    let [sheet, table_index, _, _] = table.item().leaf();
    let (rows, columns) = table_extent(&table)?;
    (cell.row() < rows && cell.column() < columns).then_some(Address {
        sheet: sheet as usize,
        table: table_index as usize,
        row: cell.row(),
        column: cell.column(),
    })
}

pub(crate) fn table_extent(table: &Node<'_>) -> Option<(u32, u32)> {
    let read = |key| {
        table.item().value(key).and_then(|value| match value {
            Value::Int(value) => u32::try_from(*value).ok(),
            _ => None,
        })
    };
    Some((read("rowCount")?, read("columnCount")?))
}

// ---------------------------------------------------------------------- set

/// One cell assignment a write stages.
#[derive(Debug, Clone)]
pub(crate) enum CellWrite {
    /// A scalar value; `None` clears the cell.
    Value(Option<Input>),
    /// A formula.
    Formula(Expression),
    /// A formula with an explicitly authored typed display cache.
    FormulaCached {
        expression: Expression,
        cached: CachedValue,
    },
}

impl CellWrite {
    pub(crate) fn change(&self, position: At) -> Change {
        match self {
            Self::Value(Some(input)) => Change::set(position, input.clone()),
            Self::Value(None) => Change::clear(position),
            Self::Formula(expression) => Change::set_formula(position, expression.clone()),
            Self::FormulaCached { expression, cached } => {
                Change::set_formula_cached(position, expression.clone(), cached.clone())
            },
        }
    }
}

/// The leaf input of a cell `value`; `None` clears the cell.
pub(crate) fn value_input(value: &PropValue<'_>) -> Result<Option<Input>> {
    let staging = |error| cells_error(error);
    let non_finite = || {
        refusal(Code::InvalidValue, "a cell number must be finite")
            .with_property("value".to_owned())
    };
    value
        .with_coerced(ValueType::Scalar, |scalar| {
            Ok(match scalar {
                None => None,
                Some(Scalar::Text(text)) => Some(Input::text(text).map_err(staging)?),
                Some(Scalar::Bool(value)) => Some(Input::boolean(value)),
                Some(Scalar::Int(value)) => {
                    #[allow(
                        clippy::cast_precision_loss,
                        reason = "Numbers stores cell numbers as doubles"
                    )]
                    let number = value as f64;
                    Some(Input::number(number).map_err(|_| non_finite())?)
                },
                Some(Scalar::Number(value)) => {
                    Some(Input::number(value).map_err(|_| non_finite())?)
                },
                Some(other) => {
                    return Err(refusal(
                        Code::InvalidValue,
                        format!("a Numbers cell value cannot be `{other}`"),
                    ));
                },
            })
        })
        .map_err(|error| error.with_property("value"))
}

/// The expression of a `formula` value; `null` clears the cell.
fn formula_expression(value: &PropValue<'_>) -> Result<Option<Expression>> {
    let text = super::write::text(value, "formula")?;
    match text {
        None => Ok(None),
        Some(text) => Expression::parse(&text).map(Some).map_err(|error| {
            edit_error(EditError::Formula(error)).with_property("formula".to_owned())
        }),
    }
}

/// The expression of a `formula` value; `null` clears the cell.
pub(crate) fn formula_write(value: &PropValue<'_>) -> Result<CellWrite> {
    match formula_expression(value)? {
        None => Ok(CellWrite::Value(None)),
        Some(expression) => Ok(CellWrite::Formula(expression)),
    }
}

/// Convert a unified scalar into the bounded typed cache accepted by the
/// Numbers formula author.  Date and duration caches deliberately have no
/// unified scalar representation, so they cannot be guessed or flattened.
fn cached_value(value: &PropValue<'_>) -> Result<CachedValue> {
    value
        .with_coerced(ValueType::Scalar, |scalar| match scalar {
            Some(Scalar::Text(text)) => {
                CachedValue::text(text).map_err(|error| edit_error(EditError::Formula(error)))
            },
            Some(Scalar::Bool(value)) => Ok(CachedValue::boolean(value)),
            Some(Scalar::Int(value)) => {
                #[allow(
                    clippy::cast_precision_loss,
                    reason = "Numbers stores cached numeric results as doubles"
                )]
                let value = value as f64;
                CachedValue::number(value).map_err(|error| edit_error(EditError::Formula(error)))
            },
            Some(Scalar::Number(value)) => {
                CachedValue::number(value).map_err(|error| edit_error(EditError::Formula(error)))
            },
            None => Err(refusal(
                Code::InvalidValue,
                "`cachedValue` must be a finite number, boolean or text",
            )),
            Some(other) => Err(refusal(
                Code::InvalidValue,
                format!("a Numbers cached formula result cannot be `{other}`"),
            )),
        })
        .map_err(|error| error.with_property("cachedValue".to_owned()))
}

/// Parse the one cached-formula assignment in a property list before any
/// cell changes are staged.  The last formula/cache key wins, matching the
/// ordered property-list contract while preserving the relative order of
/// unrelated assignments.
fn cached_formula_write(props: &Props<'_>) -> Result<Option<(usize, CellWrite)>> {
    if !props
        .iter()
        .any(|(key, _)| key.eq_ignore_ascii_case("cachedValue"))
    {
        return Ok(None);
    }
    let pair_index = props
        .iter()
        .enumerate()
        .filter(|(_, (key, _))| {
            key.eq_ignore_ascii_case("formula") || key.eq_ignore_ascii_case("cachedValue")
        })
        .map(|(index, _)| index)
        .fold(None, |_, index| Some(index))
        .expect("cachedValue key found while building cached formula");
    let Some(formula) = props.get("formula") else {
        return Err(refusal(
            Code::InvalidArgument,
            "`cachedValue` requires a non-null `formula` in the same cell assignment",
        )
        .with_property("cachedValue".to_owned()));
    };
    let Some(expression) = formula_expression(formula)? else {
        return Err(refusal(
            Code::InvalidArgument,
            "`cachedValue` requires a non-null `formula` in the same cell assignment",
        )
        .with_property("cachedValue".to_owned()));
    };
    let cached = cached_value(
        props
            .get("cachedValue")
            .expect("cachedValue key found while building cached formula"),
    )?;
    Ok(Some((
        pair_index,
        CellWrite::FormulaCached { expression, cached },
    )))
}

/// The cell assignment of one `value` or `formula` key; `None` for keys of
/// the other cell families.
pub(crate) fn cell_write(key: &str, value: &PropValue<'_>) -> Result<Option<CellWrite>> {
    if key.eq_ignore_ascii_case("value") || key.eq_ignore_ascii_case("text") {
        return value_input(value).map(|input| Some(CellWrite::Value(input)));
    }
    if key.eq_ignore_ascii_case("formula") {
        return formula_write(value).map(Some);
    }
    Ok(None)
}

/// A decimal or percentage display format parsed from a format code.
enum DisplayFormat {
    Number(Number),
    Percentage(Percentage),
    /// `null`, `General`: back to the inherited/automatic format.
    Reset,
}

/// Parse `0`, `0.00`, `#,##0.0`, `0%`, `0.0%`, `General`.
fn display_format(value: &PropValue<'_>) -> Result<DisplayFormat> {
    let Some(code) = super::write::text(value, "numberFormat")? else {
        return Ok(DisplayFormat::Reset);
    };
    let invalid = || {
        refusal(
            Code::InvalidValue,
            "`numberFormat` takes `General`, a decimal code (`0`, `0.00`, `#,##0.00`) or a percentage code (`0%`, `0.0%`)",
        )
        .with_property("numberFormat".to_owned())
    };
    let code = code.trim();
    if code.eq_ignore_ascii_case("general") {
        return Ok(DisplayFormat::Reset);
    }
    let (body, percent) = match code.strip_suffix('%') {
        Some(body) => (body, true),
        None => (code, false),
    };
    let (body, thousands) = match body.strip_prefix("#,##") {
        Some(rest) => (rest, ThousandsSeparator::Shown),
        None => (body, ThousandsSeparator::Hidden),
    };
    let places = match body {
        "0" => 0,
        _ => {
            let fraction = body.strip_prefix("0.").ok_or_else(invalid)?;
            if fraction.is_empty() || !fraction.bytes().all(|byte| byte == b'0') {
                return Err(invalid());
            }
            u8::try_from(fraction.len()).map_err(|_| invalid())?
        },
    };
    let places = DecimalPlaces::fixed(places).map_err(|_| invalid())?;
    Ok(if percent {
        DisplayFormat::Percentage(Percentage::new(places, NegativeStyle::MinusSign, thousands))
    } else {
        DisplayFormat::Number(Number::new(places, NegativeStyle::MinusSign, thousands))
    })
}

/// The typed refusal of a property the spec declares as a gap.
fn property_gap(element: &str, key: &str) -> Option<crate::Error> {
    GAPS.iter().find_map(|gap| match gap.subject {
        Subject::Property {
            element: gap_element,
            key: gap_key,
        } if gap_element == element && gap_key.eq_ignore_ascii_case(key) => Some(
            refusal(gap.code, gap.reason)
                .with_property(key.to_owned())
                .with_suggestion(gap.follow_up),
        ),
        _ => None,
    })
}

fn chart_font_size(
    value: &PropValue<'_>,
) -> Result<litchi_numbers::sheet::chart::font::TextPointSize> {
    let points = super::write::chart_font_points(value)?;
    litchi_numbers::sheet::chart::font::TextPointSize::from_points(points).map_err(|_| {
        refusal(
            Code::InvalidValue,
            "a chart font size must be positive and finite",
        )
        .with_property("font.size".to_owned())
    })
}

/// The writes of one `set` call, grouped by leaf transaction.
#[derive(Default)]
struct Staged {
    renames_sheets: Vec<(usize, String)>,
    renames_tables: Vec<(usize, usize, String)>,
    cells: BTreeMap<(usize, usize), Vec<Change>>,
    formats: Vec<(Address, DisplayFormat)>,
    /// Cell character properties, per table (record 0309).
    text_styles: BTreeMap<(usize, usize), Vec<(At, CharacterProperties)>>,
    /// Cell comments, in call order; `None` clears (record 0318).
    comments: Vec<(Address, Option<String>)>,
    /// Chart point-size writes, in call order (D7-A).
    chart_sizes: Vec<(
        usize,
        usize,
        Option<litchi_numbers::sheet::chart::font::TextPointSize>,
    )>,
}

/// Set properties on every target. Renames share one names transaction,
/// cell values and formulas one cell transaction per table.
pub(crate) fn set(
    file: &Iwork,
    package: &Package,
    target: &TargetArg<'_>,
    props: &Props<'_>,
    limits: &SemanticLimits,
) -> Result<Commit> {
    let targets = targets(file, target, limits)?;
    let target_count = match &targets {
        Targets::EmptyCell(_) => 1,
        Targets::Nodes(nodes) => nodes.len(),
    };
    let admitted_operations = target_count.checked_mul(props.len()).unwrap_or(usize::MAX);
    bound(
        LimitKind::Operations,
        admitted_operations,
        limits.max_operations(),
    )?;
    let cached_formula = match &targets {
        Targets::EmptyCell(_) => cached_formula_write(props)?,
        Targets::Nodes(nodes) if nodes.iter().any(|node| node.item().type_name() == "cell") => {
            cached_formula_write(props)?
        },
        Targets::Nodes(_) => None,
    };
    let cached_formula_index = cached_formula.as_ref().map(|(index, _)| *index);
    let cached_formula_write = cached_formula.as_ref().map(|(_, write)| write);
    let mut staged = Staged::default();
    let mut operations = 0_usize;
    let stage_cell =
        |staged: &mut Staged, address: Address, index: usize, key: &str, value: &PropValue<'_>| {
            if let Some(write) = cached_formula_write
                && (key.eq_ignore_ascii_case("formula") || key.eq_ignore_ascii_case("cachedValue"))
            {
                if Some(index) == cached_formula_index {
                    staged
                        .cells
                        .entry((address.sheet, address.table))
                        .or_default()
                        .push(write.change(address.position()));
                }
                return Ok(());
            }
            if let Some(write) = cell_write(key, value)? {
                staged
                    .cells
                    .entry((address.sheet, address.table))
                    .or_default()
                    .push(write.change(address.position()));
            } else if key.eq_ignore_ascii_case("numberformat") {
                staged.formats.push((address, display_format(value)?));
            } else if key.eq_ignore_ascii_case("comment") {
                staged
                    .comments
                    .push((address, super::write::text(value, "comment")?));
            } else if key.len() > 5 && key[..5].eq_ignore_ascii_case("font.") {
                let cells = staged
                    .text_styles
                    .entry((address.sheet, address.table))
                    .or_default();
                let index = match cells
                    .iter()
                    .position(|(cell, _)| *cell == address.position())
                {
                    Some(index) => index,
                    None => {
                        cells.push((address.position(), CharacterProperties::new()));
                        cells.len() - 1
                    },
                };
                if !super::write::font_property(&mut cells[index].1, key, value, FONT_COLOR)? {
                    return Err(property_gap("cell", key)
                        .unwrap_or_else(|| super::write::unwritable("cell", key)));
                }
            } else {
                return Err(property_gap("cell", key)
                    .unwrap_or_else(|| super::write::unwritable("cell", key)));
            }
            Ok(())
        };
    match targets {
        Targets::EmptyCell(address) => {
            for (index, (key, value)) in props.iter().enumerate() {
                operations += 1;
                bound(LimitKind::Operations, operations, limits.max_operations())?;
                stage_cell(&mut staged, address, index, key.as_str(), value)?;
            }
        },
        Targets::Nodes(nodes) => {
            for node in &nodes {
                for (index, (key, value)) in props.iter().enumerate() {
                    operations += 1;
                    bound(LimitKind::Operations, operations, limits.max_operations())?;
                    let key = key.as_str();
                    let [sheet, table, _, _] = node.item().leaf().map(|value| value as usize);
                    match node.item().type_name() {
                        "cell" => stage_cell(&mut staged, cell_address(node), index, key, value)?,
                        "chart" if key.eq_ignore_ascii_case("font.size") => {
                            staged.chart_sizes.push((
                                sheet,
                                node.item().leaf()[1] as usize,
                                (!value.is_null())
                                    .then(|| chart_font_size(value))
                                    .transpose()?,
                            ));
                        },
                        "sheet" if key.eq_ignore_ascii_case("name") => {
                            staged.renames_sheets.push((
                                sheet,
                                super::write::text(value, "name")?.unwrap_or_default(),
                            ))
                        },
                        "table" if key.eq_ignore_ascii_case("name") => {
                            staged.renames_tables.push((
                                sheet,
                                table,
                                super::write::text(value, "name")?.unwrap_or_default(),
                            ))
                        },
                        "sheet" | "table" | "workbook" | "mergeCell" => {
                            return Err(key_refusal(node, key));
                        },
                        "chart" => return Err(key_refusal(node, key)),
                        _ => return Err(element_refusal("set", node)),
                    }
                }
            }
        },
    }
    let mut chain = Chain::new(package.clone());
    commit_staged(&mut chain, staged)?;
    finish(
        file,
        chain.changed,
        Owner::Numbers(chain.package),
        chain.patches,
        operations,
        limits,
    )
}

fn push(chain: &mut Chain<Package>, package: Package, patch: Patch) {
    let changed = !patch.is_noop();
    chain.push(package, LeafPatch::Numbers(patch), changed);
}

/// Commit the staged writes in a fixed order: renames, cells, display
/// formats. Every later step addresses by position, which the
/// earlier ones keep.
fn commit_staged(chain: &mut Chain<Package>, staged: Staged) -> Result<()> {
    if !staged.renames_sheets.is_empty() || !staged.renames_tables.is_empty() {
        let names = |error| edit_error(EditError::Names(error));
        let mut edit = chain.package.edit_names();
        for (sheet, name) in &staged.renames_sheets {
            edit = edit
                .rename_sheet(SheetSelector::index(*sheet), name)
                .map_err(names)?;
        }
        for (sheet, table, name) in &staged.renames_tables {
            edit = edit
                .rename_table(
                    SheetSelector::index(*sheet),
                    TableSelector::index(*table),
                    name,
                )
                .map_err(names)?;
        }
        let commit = edit.commit().map_err(names)?;
        let patch = commit.patch().clone();
        push(chain, commit.into_package(), Patch::Names(patch));
    }
    commit_cells(chain, staged.cells)?;
    for (address, text) in staged.comments {
        commit_comment(chain, address, text.as_deref())?;
    }
    for ((sheet, table), cells) in staged.text_styles {
        let error = |error| edit_error(EditError::TextStyle(error));
        let mut edit = chain
            .package
            .edit_table_cell_text_style(SheetSelector::index(sheet), TableSelector::index(table))
            .map_err(error)?;
        for (position, properties) in cells {
            edit = edit.set(position, properties).map_err(error)?;
        }
        let commit = edit.commit().map_err(error)?;
        let (next, patch) = commit.into_parts();
        push(chain, next, Patch::TextStyle(patch));
    }
    for (sheet, chart, size) in staged.chart_sizes {
        let error = |error| edit_error(EditError::ChartFont(error));
        let mut edit = chain
            .package
            .edit_sheet_chart_font(
                SheetSelector::index(sheet),
                litchi_numbers::ChartSelector::index(chart),
            )
            .map_err(error)?;
        match size {
            Some(size) => {
                edit.set(ChartFontProperty::Size(size)).map_err(error)?;
            },
            None => {
                edit.reset(ChartFontPropertyKind::Size).map_err(error)?;
            },
        }
        let commit = edit.commit().map_err(error)?;
        let patch = commit.patch().clone();
        push(chain, commit.into_package(), Patch::ChartFontSize(patch));
    }
    for (address, format) in staged.formats {
        let sheet = SheetSelector::index(address.sheet);
        let table = TableSelector::index(address.table);
        match format {
            DisplayFormat::Percentage(value) => {
                let error = |error| edit_error(EditError::PercentageFormat(error));
                let commit = chain
                    .package
                    .edit_table_cell_percentage_format(sheet, table, address.position())
                    .map_err(error)?
                    .set(value)
                    .commit()
                    .map_err(error)?;
                let patch = commit.patch().clone();
                push(chain, commit.into_package(), Patch::PercentageFormat(patch));
            },
            DisplayFormat::Number(value) => {
                let error = |error| edit_error(EditError::NumberFormat(error));
                let commit = chain
                    .package
                    .edit_table_cell_number_format(sheet, table, address.position())
                    .map_err(error)?
                    .set(value)
                    .commit()
                    .map_err(error)?;
                let patch = commit.patch().clone();
                push(chain, commit.into_package(), Patch::NumberFormat(patch));
            },
            DisplayFormat::Reset => {
                let error = |error| edit_error(EditError::NumberFormat(error));
                let commit = chain
                    .package
                    .edit_table_cell_number_format(sheet, table, address.position())
                    .map_err(error)?
                    .clear()
                    .commit()
                    .map_err(error)?;
                let patch = commit.patch().clone();
                push(chain, commit.into_package(), Patch::NumberFormat(patch));
            },
        }
    }
    Ok(())
}

/// Set (`Some`) or clear (`None`) one cell comment.
fn commit_comment(chain: &mut Chain<Package>, address: Address, text: Option<&str>) -> Result<()> {
    let error = |error| edit_error(EditError::Comment(error));
    let sheet = SheetSelector::index(address.sheet);
    let table = TableSelector::index(address.table);
    let commit = match text {
        Some(text) => chain
            .package
            .set_table_cell_comment(sheet, table, address.position(), text),
        None => chain
            .package
            .clear_table_cell_comment(sheet, table, address.position()),
    }
    .map_err(|failure| error(failure).with_property("comment".to_owned()))?;
    let patch = commit.patch().clone();
    push(chain, commit.into_package(), Patch::Comment(patch));
    Ok(())
}

/// Stage each table's changes on one cell transaction, tables in document
/// order.
pub(crate) fn commit_cells(
    chain: &mut Chain<Package>,
    cells: BTreeMap<(usize, usize), Vec<Change>>,
) -> Result<()> {
    for ((sheet, table), changes) in cells {
        let mut edit: Edit<'_> = chain
            .package
            .edit_table_cells(SheetSelector::index(sheet), TableSelector::index(table))
            .map_err(cells_error)?;
        edit = edit.extend(changes).map_err(cells_error)?;
        let commit = edit.commit().map_err(cells_error)?;
        let patch = commit.patch().clone();
        let changed = !patch.is_noop();
        chain.push(
            commit.into_package(),
            LeafPatch::NumbersCells(patch),
            changed,
        );
    }
    Ok(())
}

// ------------------------------------------------------------- add / remove

/// Add a merged range (`mergeCell` with `ref`) to a table.
pub(crate) fn add(
    file: &Iwork,
    package: &Package,
    parent: &PathArg<'_>,
    node: &NewNode<'_>,
    limits: &SemanticLimits,
) -> Result<(Commit, Path<'static>)> {
    let kind = node.kind();
    if kind.eq_ignore_ascii_case("sheet") {
        return add_sheet(file, package, parent, node, limits);
    }
    if kind.eq_ignore_ascii_case("table") {
        return add_table(file, package, parent, node, limits);
    }
    if let Some(row_axis) = axis_kind(kind) {
        return add_axis(file, package, parent, node, row_axis, limits);
    }
    if kind.eq_ignore_ascii_case("comment") {
        return add_comment(file, package, parent, node, limits);
    }
    if !(kind.eq_ignore_ascii_case("mergeCell")
        || kind.eq_ignore_ascii_case("merge")
        || kind.eq_ignore_ascii_case("mergedRange"))
    {
        return Err(add_gap(kind));
    }
    let table = node_at::<Numbers>(file, parent)?;
    if table.item().type_name() != "table" {
        return Err(refusal(
            Code::InvalidArgument,
            "a merged range is added under a table (`/sheet[N]/table[N]`)",
        ));
    }
    let mut range = None;
    for (key, value) in node.props().iter() {
        if key.as_str().eq_ignore_ascii_case("ref") {
            range = super::write::text(value, "ref")?;
        } else {
            return Err(super::write::unwritable("mergeCell", key.as_str()));
        }
    }
    let range = range.ok_or_else(|| {
        refusal(
            Code::InvalidArgument,
            "a merged range needs `ref` (`B2:C3`)",
        )
        .with_property("ref".to_owned())
    })?;
    let region = parse_region(&range)?;
    bound(LimitKind::Operations, 1, limits.max_operations())?;
    let [sheet, table_index, _, _] = table.item().leaf().map(|value| value as usize);
    let merges = |error| edit_error(EditError::Merges(error));
    let mut edit = package
        .edit_table_merges(
            SheetSelector::index(sheet),
            TableSelector::index(table_index),
        )
        .map_err(merges)?;
    edit.merge(region).map_err(merges)?;
    let commit = edit.commit().map_err(merges)?;
    let ordinal = commit
        .package()
        .table_merges(
            SheetSelector::index(sheet),
            TableSelector::index(table_index),
        )
        .map_err(merges)?
        .iter()
        .position(|candidate| *candidate == region)
        .unwrap_or_default();
    let patch = commit.patch().clone();
    let mut chain = Chain::new(package.clone());
    push(&mut chain, commit.into_package(), Patch::Merges(patch));
    let commit = finish(
        file,
        chain.changed,
        Owner::Numbers(chain.package),
        chain.patches,
        1,
        limits,
    )?;
    let path = Path::from_steps([
        Step::new("sheet").with(Predicate::Index(Position::new(sheet))),
        Step::new("table").with(Predicate::Index(Position::new(table_index))),
        Step::new("mergeCell").with(Predicate::Index(Position::new(ordinal))),
    ]);
    Ok((commit, path))
}

/// The first `<prefix> N` (N from 1) that `taken` does not hold.
fn fresh_name(prefix: &str, taken: &[&str]) -> String {
    (1..)
        .map(|index| format!("{prefix} {index}"))
        .find(|name| !taken.contains(&name.as_str()))
        .unwrap_or_else(|| prefix.to_owned())
}

/// A non-negative count property of a new table.
fn count_prop(value: &PropValue<'_>, key: &str) -> Result<u32> {
    value
        .with_coerced(ValueType::Int, |scalar| match scalar {
            Some(Scalar::Int(value)) => u32::try_from(value)
                .map_err(|_| refusal(Code::InvalidValue, "a table size must be a positive count")),
            _ => Err(refusal(Code::InvalidValue, "a table size must be a count")),
        })
        .map_err(|error| error.with_property(key.to_owned()))
}

/// The zero-based final position `placement` names among `count` siblings
/// once the new one is appended, or `None` to keep it last.
fn add_position(
    count: usize,
    placement: Option<&crate::Placement<'_>>,
    anchor: impl Fn(&PathArg<'_>) -> Result<usize>,
) -> Result<Option<usize>> {
    use crate::Placement;
    let Some(placement) = placement else {
        return Ok(None);
    };
    let last = count;
    let to = match placement {
        Placement::End | Placement::Into(_, None) => last,
        Placement::Start => 0,
        Placement::Position(position) | Placement::Into(_, Some(position)) => position.get(),
        Placement::Before(path) => anchor(path)?,
        Placement::After(path) => anchor(path)? + 1,
    };
    if to > last {
        return Err(refusal(
            Code::OutOfRange,
            "the insertion position is outside the parent's children",
        ));
    }
    Ok((to != last).then_some(to))
}

/// Add an empty sheet (`name`, default `Sheet N`), appended or placed.
fn add_sheet(
    file: &Iwork,
    package: &Package,
    parent: &PathArg<'_>,
    node: &NewNode<'_>,
    limits: &SemanticLimits,
) -> Result<(Commit, Path<'static>)> {
    let root = node_at::<Numbers>(file, parent)?;
    if root.item().type_name() != "workbook" {
        return Err(refusal(
            Code::InvalidArgument,
            "a Numbers sheet is added under the workbook (`/`)",
        ));
    }
    let mut name = None;
    for (key, value) in node.props().iter() {
        if key.as_str().eq_ignore_ascii_case("name") {
            name = super::write::text(value, "name")?;
        } else {
            return Err(super::write::unwritable("sheet", key.as_str()));
        }
    }
    let taken: Vec<&str> = package.sheets().iter().map(|sheet| sheet.name()).collect();
    let name = name.unwrap_or_else(|| fresh_name("Sheet", &taken));
    let count = package.sheets().len();
    let position = add_position(count, node.placement(), |anchor| {
        let anchor = node_at::<Numbers>(file, anchor)?;
        if anchor.item().type_name() == "sheet" {
            Ok(anchor.item().leaf()[0] as usize)
        } else {
            Err(refusal(
                Code::InvalidArgument,
                "a sheet is placed before or after another sheet",
            ))
        }
    })?;
    bound(LimitKind::Operations, 1, limits.max_operations())?;
    let error = |error| edit_error(EditError::SheetLifecycle(error));
    let commit = package
        .edit_sheet_lifecycle()
        .add(&name)
        .map_err(error)?
        .commit()
        .map_err(error)?;
    let mut chain = Chain::new(package.clone());
    let (next, patch) = commit.into_parts();
    push(&mut chain, next, Patch::SheetLifecycle(patch));
    let index = position.unwrap_or(count);
    if let Some(to) = position {
        let error = |error| edit_error(EditError::Order(error));
        let commit = chain
            .package
            .edit_sheet_order()
            .move_sheet(SheetSelector::index(count), to)
            .map_err(error)?
            .commit()
            .map_err(error)?;
        let patch = commit.patch().clone();
        push(&mut chain, commit.into_package(), Patch::Order(patch));
    }
    let commit = finish(
        file,
        chain.changed,
        Owner::Numbers(chain.package),
        chain.patches,
        1,
        limits,
    )?;
    let path = Path::from_steps([Step::new("sheet").with(Predicate::Index(Position::new(index)))]);
    Ok((commit, path))
}

/// Add an empty table (`name`, `rowCount`, `columnCount`; defaults
/// `Table N`, 10 × 5) to a sheet, appended or placed among its tables.
fn add_table(
    file: &Iwork,
    package: &Package,
    parent: &PathArg<'_>,
    node: &NewNode<'_>,
    limits: &SemanticLimits,
) -> Result<(Commit, Path<'static>)> {
    let sheet_node = node_at::<Numbers>(file, parent)?;
    if sheet_node.item().type_name() != "sheet" {
        return Err(refusal(
            Code::InvalidArgument,
            "a Numbers table is added under a sheet (`/sheet[N]`)",
        ));
    }
    let sheet = sheet_node.item().leaf()[0] as usize;
    let (mut name, mut rows, mut columns) = (None, 10_u32, 5_u32);
    for (key, value) in node.props().iter() {
        let key = key.as_str();
        if key.eq_ignore_ascii_case("name") {
            name = super::write::text(value, "name")?;
        } else if key.eq_ignore_ascii_case("rowCount") || key.eq_ignore_ascii_case("rows") {
            rows = count_prop(value, "rowCount")?;
        } else if key.eq_ignore_ascii_case("columnCount") || key.eq_ignore_ascii_case("columns") {
            columns = count_prop(value, "columnCount")?;
        } else {
            return Err(super::write::unwritable("table", key));
        }
    }
    let leaf_sheet = package
        .sheets()
        .get(sheet)
        .ok_or_else(|| refusal(Code::PathNotFound, "the sheet is no longer in the workbook"))?;
    let taken: Vec<&str> = leaf_sheet.tables().map(|table| table.name()).collect();
    let name = name.unwrap_or_else(|| fresh_name("Table", &taken));
    let count = leaf_sheet.table_count();
    let position = add_position(count, node.placement(), |anchor| {
        let anchor = node_at::<Numbers>(file, anchor)?;
        let [anchor_sheet, anchor_table, _, _] = anchor.item().leaf().map(|value| value as usize);
        if anchor.item().type_name() == "table" && anchor_sheet == sheet {
            Ok(anchor_table)
        } else {
            Err(refusal(
                Code::InvalidArgument,
                "a table is placed before or after another table of its sheet",
            ))
        }
    })?;
    bound(LimitKind::Operations, 1, limits.max_operations())?;
    let error = |error| edit_error(EditError::TableLifecycle(error));
    let commit = package
        .edit_table_lifecycle(SheetSelector::index(sheet))
        .map_err(error)?
        .add(&name, rows, columns)
        .map_err(error)?
        .commit()
        .map_err(error)?;
    let mut chain = Chain::new(package.clone());
    let (next, patch) = commit.into_parts();
    push(&mut chain, next, Patch::TableLifecycle(patch));
    let index = position.unwrap_or(count);
    if let Some(to) = position {
        let error = |error| edit_error(EditError::TableOrder(error));
        let commit = chain
            .package
            .edit_table_order(SheetSelector::index(sheet))
            .map_err(error)?
            .move_table(TableSelector::index(count), to)
            .map_err(error)?
            .commit()
            .map_err(error)?;
        let (next, patch) = commit.into_parts();
        push(&mut chain, next, Patch::TableOrder(patch));
    }
    let commit = finish(
        file,
        chain.changed,
        Owner::Numbers(chain.package),
        chain.patches,
        1,
        limits,
    )?;
    let path = Path::from_steps([
        Step::new("sheet").with(Predicate::Index(Position::new(sheet))),
        Step::new("table").with(Predicate::Index(Position::new(index))),
    ]);
    Ok((commit, path))
}

fn parse_region(range: &str) -> Result<Region> {
    let invalid = || {
        refusal(
            Code::InvalidValue,
            "`ref` must be an A1 range of at least two cells (`B2:C3`)",
        )
        .with_property("ref".to_owned())
    };
    let range = Area::from_a1(range).map_err(|_| invalid())?;
    let start = range.start();
    let end = range.end();
    Region::new(
        start.row(),
        start.column(),
        // `Area` is half-open: its end is exclusive.
        end.row().saturating_sub(start.row()),
        end.column().saturating_sub(start.column()),
    )
    .map_err(|_| invalid())
}

/// The typed refusal of an `add` the leaf cannot perform.
fn add_gap(_kind: &str) -> crate::Error {
    refusal(
        Code::UnsupportedCapability,
        "a Numbers workbook adds sheets, tables, rows, columns (`row`, `col`), merged ranges (`mergeCell`) and cell comments (`comment` under a cell)",
    )
}

/// The cell `path` names, stored or empty in the grid.
fn cell_at(file: &Iwork, path: &PathArg<'_>) -> Result<Address> {
    match node_at::<Numbers>(file, path) {
        Ok(node) if node.item().type_name() == "cell" => Ok(cell_address(&node)),
        Ok(_) => Err(refusal(
            Code::InvalidArgument,
            "a cell comment goes on a cell (`/sheet[N]/table[N]/A1`)",
        )),
        Err(error) if error.code() == Code::PathNotFound => {
            empty_cell(file, path.path()).ok_or(error)
        },
        Err(error) => Err(error),
    }
}

/// Add a comment (`text`) to a cell that has none (record 0318).
fn add_comment(
    file: &Iwork,
    package: &Package,
    parent: &PathArg<'_>,
    node: &NewNode<'_>,
    limits: &SemanticLimits,
) -> Result<(Commit, Path<'static>)> {
    let address = cell_at(file, parent)?;
    let mut text = None;
    for (key, value) in node.props().iter() {
        if key.as_str().eq_ignore_ascii_case("text") || key.as_str().eq_ignore_ascii_case("comment")
        {
            text = super::write::text(value, "text")?;
        } else {
            return Err(super::write::unwritable("comment", key.as_str()));
        }
    }
    let text = text.ok_or_else(|| {
        refusal(Code::InvalidArgument, "a cell comment needs `text`")
            .with_property("text".to_owned())
    })?;
    bound(LimitKind::Operations, 1, limits.max_operations())?;
    let existing = package
        .table_cell_comment(
            SheetSelector::index(address.sheet),
            TableSelector::index(address.table),
            address.position(),
        )
        .map_err(|error| edit_error(EditError::Comment(error)))?;
    if existing.is_some() {
        return Err(refusal(
            Code::InvalidArgument,
            "the cell already has a comment; set its `comment` to replace it",
        ));
    }
    let mut chain = Chain::new(package.clone());
    commit_comment(&mut chain, address, Some(&text))?;
    let commit = finish(
        file,
        chain.changed,
        Owner::Numbers(chain.package),
        chain.patches,
        1,
        limits,
    )?;
    Ok((commit, address_path(address)))
}

/// The canonical path of a cell.
fn address_path(address: Address) -> Path<'static> {
    let mut column = address.column + 1;
    let mut letters = Vec::new();
    while column > 0 {
        let digit = (column - 1) % 26;
        letters.push(char::from(b'A' + u8::try_from(digit).unwrap_or(0)));
        column = (column - 1) / 26;
    }
    let name: String = letters
        .into_iter()
        .rev()
        .chain((address.row + 1).to_string().chars())
        .collect();
    Path::from_steps([
        Step::new("sheet").with(Predicate::Index(Position::new(address.sheet))),
        Step::new("table").with(Predicate::Index(Position::new(address.table))),
        Step::new(name),
    ])
}

/// The cell whose comment `path` names: `<cell>/comment`.
fn comment_target(file: &Iwork, path: &Path<'_>) -> Option<Address> {
    let (last, parent) = path.steps().split_last()?;
    if !last.name().eq_ignore_ascii_case("comment")
        || !(last.predicates().is_empty()
            || matches!(last.predicates(), [Predicate::Index(position)] if position.get() == 0))
    {
        return None;
    }
    let parent = Path::from_steps(parent.iter().cloned());
    let parent = parent.into_path().ok()?;
    cell_at(file, &parent).ok()
}

// ------------------------------------------------------------ rows / columns

/// `Some(true)` for a row element name, `Some(false)` for a column one.
fn axis_kind(kind: &str) -> Option<bool> {
    if kind.eq_ignore_ascii_case("row") {
        Some(true)
    } else if kind.eq_ignore_ascii_case("col") || kind.eq_ignore_ascii_case("column") {
        Some(false)
    } else {
        None
    }
}

/// A table's size and section counts, for mapping a physical row or column
/// to its section.
#[derive(Debug, Clone, Copy)]
struct Sections {
    rows: usize,
    columns: usize,
    header_rows: usize,
    header_columns: usize,
    footer_rows: usize,
}

fn sections(package: &Package, sheet: usize, table: usize) -> Result<Sections> {
    let error = |error| edit_error(EditError::Topology(error));
    let leaf = package
        .sheets()
        .get(sheet)
        .and_then(|sheet| sheet.tables().nth(table))
        .ok_or_else(|| refusal(Code::PathNotFound, "the table is no longer in the workbook"))?;
    let settings = package
        .table_header_settings(SheetSelector::index(sheet), TableSelector::index(table))
        .map_err(|_| error(litchi_numbers::table::topology::Error::InvalidSource))?;
    Ok(Sections {
        rows: leaf.row_count() as usize,
        columns: leaf.column_count() as usize,
        header_rows: settings.header_row_count(),
        header_columns: settings.header_column_count(),
        footer_rows: settings.footer_row_count(),
    })
}

/// The section-relative insertion before physical row `at` (`at` may equal
/// the row count): a row inside the header rows grows them, one past the
/// body joins the footer rows, every other one joins the body.
fn row_insertion(sections: Sections, at: usize) -> RowInsertion {
    let body_end = sections.rows - sections.footer_rows;
    if at < sections.header_rows {
        RowInsertion::header(at)
    } else if at <= body_end {
        RowInsertion::body(at - sections.header_rows)
    } else {
        RowInsertion::footer(at - body_end)
    }
}

fn row_deletion(sections: Sections, at: usize) -> RowDeletion {
    let body_end = sections.rows - sections.footer_rows;
    if at < sections.header_rows {
        RowDeletion::header(at)
    } else if at < body_end {
        RowDeletion::body(at - sections.header_rows)
    } else {
        RowDeletion::footer(at - body_end)
    }
}

fn column_insertion(sections: Sections, at: usize) -> ColumnInsertion {
    if at < sections.header_columns {
        ColumnInsertion::header(at)
    } else {
        ColumnInsertion::body(at - sections.header_columns)
    }
}

fn column_deletion(sections: Sections, at: usize) -> ColumnDeletion {
    if at < sections.header_columns {
        ColumnDeletion::header(at)
    } else {
        ColumnDeletion::body(at - sections.header_columns)
    }
}

/// The table and zero-based index a `row[K]`/`col[K]` path names, with the
/// axis (`true` for rows).
fn axis_target(file: &Iwork, path: &Path<'_>) -> Option<(usize, usize, bool, usize)> {
    let (last, parent) = path.steps().split_last()?;
    let row_axis = axis_kind(last.name())?;
    let [Predicate::Index(position)] = last.predicates() else {
        return None;
    };
    let parent = Path::from_steps(parent.iter().cloned());
    let parent = parent.into_path().ok()?;
    let table = node_at::<Numbers>(file, &parent).ok()?;
    if table.item().type_name() != "table" {
        return None;
    }
    let [sheet, table_index, _, _] = table.item().leaf().map(|value| value as usize);
    Some((sheet, table_index, row_axis, position.get()))
}

/// Insert a blank row or column into a table, appended to its body or
/// placed (`Start`, `End`, `Position(N)`, before or after another row or
/// column of the table).
fn add_axis(
    file: &Iwork,
    package: &Package,
    parent: &PathArg<'_>,
    node: &NewNode<'_>,
    row_axis: bool,
    limits: &SemanticLimits,
) -> Result<(Commit, Path<'static>)> {
    let noun = if row_axis { "row" } else { "column" };
    let table = node_at::<Numbers>(file, parent)?;
    if table.item().type_name() != "table" {
        return Err(refusal(
            Code::InvalidArgument,
            format!("a Numbers {noun} is added under a table (`/sheet[N]/table[N]`)"),
        ));
    }
    if let Some((key, _)) = node.props().iter().next() {
        return Err(super::write::unwritable(noun, key.as_str()));
    }
    let [sheet, table_index, _, _] = table.item().leaf().map(|value| value as usize);
    let sections = sections(package, sheet, table_index)?;
    let count = if row_axis {
        sections.rows
    } else {
        sections.columns
    };
    let anchor = |path: &PathArg<'_>| -> Result<usize> {
        match axis_target(file, path.path()) {
            Some((anchor_sheet, anchor_table, anchor_row, index))
                if (anchor_sheet, anchor_table, anchor_row) == (sheet, table_index, row_axis)
                    && index < count =>
            {
                Ok(index)
            },
            _ => Err(refusal(
                Code::InvalidArgument,
                format!("a {noun} is placed before or after another {noun} of its table"),
            )),
        }
    };
    let at = match add_position(count, node.placement(), anchor)? {
        Some(at) => at,
        // Appended: the end of the body (before any footer rows).
        None if row_axis => sections.rows - sections.footer_rows,
        None => sections.columns,
    };
    bound(LimitKind::Operations, 1, limits.max_operations())?;
    let error = |error| edit_error(EditError::Topology(error));
    let edit = package
        .edit_table_topology(
            SheetSelector::index(sheet),
            TableSelector::index(table_index),
        )
        .map_err(error)?;
    let edit = if row_axis {
        edit.insert_row(row_insertion(sections, at))
    } else {
        edit.insert_column(column_insertion(sections, at))
    }
    .map_err(error)?;
    let commit = edit.commit().map_err(error)?;
    let mut chain = Chain::new(package.clone());
    let (next, patch) = commit.into_parts();
    push(&mut chain, next, Patch::Topology(patch));
    let commit = finish(
        file,
        chain.changed,
        Owner::Numbers(chain.package),
        chain.patches,
        1,
        limits,
    )?;
    let path = Path::from_steps([
        Step::new("sheet").with(Predicate::Index(Position::new(sheet))),
        Step::new("table").with(Predicate::Index(Position::new(table_index))),
        Step::new(if row_axis { "row" } else { "col" }).with(Predicate::Index(Position::new(at))),
    ]);
    Ok((commit, path))
}

/// Remove one row or column of a table, with its cells.
fn remove_axis(
    file: &Iwork,
    package: &Package,
    (sheet, table, row_axis, at): (usize, usize, bool, usize),
    limits: &SemanticLimits,
) -> Result<Commit> {
    let sections = sections(package, sheet, table)?;
    let count = if row_axis {
        sections.rows
    } else {
        sections.columns
    };
    if at >= count {
        return Err(refusal(
            Code::PathNotFound,
            "the table has no such row or column",
        ));
    }
    bound(LimitKind::Operations, 1, limits.max_operations())?;
    let error = |error| edit_error(EditError::Topology(error));
    let edit = package
        .edit_table_topology(SheetSelector::index(sheet), TableSelector::index(table))
        .map_err(error)?;
    let edit = if row_axis {
        edit.remove_row(row_deletion(sections, at))
    } else {
        edit.remove_column(column_deletion(sections, at))
    }
    .map_err(error)?;
    let commit = edit.commit().map_err(error)?;
    let mut chain = Chain::new(package.clone());
    let (next, patch) = commit.into_parts();
    push(&mut chain, next, Patch::Topology(patch));
    finish(
        file,
        chain.changed,
        Owner::Numbers(chain.package),
        chain.patches,
        1,
        limits,
    )
}

/// Remove targets: a cell is cleared, a merged range unmerged.
pub(crate) fn remove(
    file: &Iwork,
    package: &Package,
    target: &TargetArg<'_>,
    limits: &SemanticLimits,
) -> Result<Commit> {
    if let Some(axis) = target
        .path()
        .and_then(|path| axis_target(file, path.path()))
    {
        return remove_axis(file, package, axis, limits);
    }
    if let Some(address) = target
        .path()
        .and_then(|path| comment_target(file, path.path()))
    {
        bound(LimitKind::Operations, 1, limits.max_operations())?;
        let mut chain = Chain::new(package.clone());
        commit_comment(&mut chain, address, None)?;
        return finish(
            file,
            chain.changed,
            Owner::Numbers(chain.package),
            chain.patches,
            1,
            limits,
        );
    }
    let nodes = match targets(file, target, limits)? {
        Targets::EmptyCell(_) => {
            // An empty in-grid cell has nothing to remove.
            return finish(
                file,
                false,
                Owner::Numbers(package.clone()),
                Vec::new(),
                0,
                limits,
            );
        },
        Targets::Nodes(nodes) => nodes,
    };
    let mut cells: BTreeMap<(usize, usize), Vec<Change>> = BTreeMap::new();
    let mut merges: BTreeMap<(usize, usize), Vec<Region>> = BTreeMap::new();
    let mut tables: Vec<(usize, usize)> = Vec::new();
    let mut sheets: Vec<usize> = Vec::new();
    let mut operations = 0_usize;
    for node in &nodes {
        operations += 1;
        bound(LimitKind::Operations, operations, limits.max_operations())?;
        let [sheet, table, _, _] = node.item().leaf().map(|value| value as usize);
        match node.item().type_name() {
            "cell" => {
                let address = cell_address(node);
                cells
                    .entry((sheet, table))
                    .or_default()
                    .push(Change::clear(address.position()));
            },
            "mergeCell" => {
                let range = match node.value("ref")? {
                    Some(Value::Text(range)) => range.to_string(),
                    _ => return Err(element_refusal("remove", node)),
                };
                merges
                    .entry((sheet, table))
                    .or_default()
                    .push(parse_region(&range)?);
            },
            "sheet" => sheets.push(sheet),
            "table" => tables.push((sheet, table)),
            _ => return Err(element_refusal("remove", node)),
        }
    }
    let mut chain = Chain::new(package.clone());
    for ((sheet, table), regions) in merges {
        let error = |error| edit_error(EditError::Merges(error));
        let mut edit = chain
            .package
            .edit_table_merges(SheetSelector::index(sheet), TableSelector::index(table))
            .map_err(error)?;
        for region in regions {
            edit.unmerge(region).map_err(error)?;
        }
        let commit = edit.commit().map_err(error)?;
        let patch = commit.patch().clone();
        push(&mut chain, commit.into_package(), Patch::Merges(patch));
    }
    commit_cells(&mut chain, cells)?;
    // Whole tables, then whole sheets, each from the last so the earlier
    // positions hold. A table of a removed sheet goes with the sheet.
    sheets.sort_unstable();
    sheets.dedup();
    tables.retain(|(sheet, _)| !sheets.contains(sheet));
    tables.sort_unstable();
    tables.dedup();
    for (sheet, table) in tables.iter().rev() {
        let error = |error| edit_error(EditError::TableLifecycle(error));
        let commit = chain
            .package
            .edit_table_lifecycle(SheetSelector::index(*sheet))
            .map_err(error)?
            .remove(TableSelector::index(*table))
            .map_err(error)?
            .commit()
            .map_err(error)?;
        let (next, patch) = commit.into_parts();
        push(&mut chain, next, Patch::TableLifecycle(patch));
    }
    for sheet in sheets.iter().rev() {
        let error = |error| edit_error(EditError::SheetLifecycle(error));
        let commit = chain
            .package
            .edit_sheet_lifecycle()
            .remove(SheetSelector::index(*sheet))
            .map_err(error)?
            .commit()
            .map_err(error)?;
        let (next, patch) = commit.into_parts();
        push(&mut chain, next, Patch::SheetLifecycle(patch));
    }
    finish(
        file,
        chain.changed,
        Owner::Numbers(chain.package),
        chain.patches,
        operations,
        limits,
    )
}

// ---------------------------------------------------------------------- move

/// Move a sheet among the workbook's sheets, or a table into another sheet.
pub(crate) fn move_to(
    file: &Iwork,
    package: &Package,
    path: &PathArg<'_>,
    placement: &crate::Placement<'_>,
    limits: &SemanticLimits,
) -> Result<Commit> {
    let node = node_at::<Numbers>(file, path)?;
    let mut chain = Chain::new(package.clone());
    match node.item().type_name() {
        "sheet" => {
            if matches!(placement, crate::Placement::Into(..)) {
                return Err(refusal(
                    Code::UnsupportedCapability,
                    "a Numbers sheet moves among the workbook's sheets",
                ));
            }
            let from = node.item().leaf()[0] as usize;
            let sheets = file.tree().item(super::tree::ROOT).children().len();
            let to = destination(from, sheets, placement, |anchor| {
                let anchor = node_at::<Numbers>(file, anchor)?;
                if anchor.item().type_name() == "sheet" {
                    Ok(anchor.item().leaf()[0] as usize)
                } else {
                    Err(refusal(
                        Code::InvalidArgument,
                        "a sheet moves before or after another sheet",
                    ))
                }
            })?;
            move_sheet(&mut chain, from, to)?;
        },
        "table" => {
            let crate::Placement::Into(destination_path, position) = placement else {
                // Within its sheet.
                let [sheet, table, _, _] = node.item().leaf().map(|value| value as usize);
                let count = package
                    .sheets()
                    .get(sheet)
                    .map_or(0, |sheet| sheet.table_count());
                let to = destination(table, count, placement, |anchor| {
                    let anchor = node_at::<Numbers>(file, anchor)?;
                    let [anchor_sheet, anchor_table, _, _] =
                        anchor.item().leaf().map(|value| value as usize);
                    if anchor.item().type_name() == "table" && anchor_sheet == sheet {
                        Ok(anchor_table)
                    } else {
                        Err(refusal(
                            Code::InvalidArgument,
                            "a table moves before or after another table of its sheet; use `Placement::Into` for another sheet",
                        ))
                    }
                })?;
                order_table(&mut chain, sheet, table, to)?;
                bound(LimitKind::Operations, 1, limits.max_operations())?;
                return finish(
                    file,
                    chain.changed,
                    Owner::Numbers(chain.package),
                    chain.patches,
                    1,
                    limits,
                );
            };
            let destination = node_at::<Numbers>(file, destination_path)?;
            if destination.item().type_name() != "sheet" {
                return Err(refusal(Code::InvalidArgument, "a table moves into a sheet"));
            }
            let [sheet, table, _, _] = node.item().leaf().map(|value| value as usize);
            let to = destination.item().leaf()[0] as usize;
            let appended = destination
                .item()
                .children()
                .iter()
                .filter(|child| file.tree().item(**child).type_name() == "table")
                .count();
            if position.is_some_and(|position| position.get() != appended) {
                return Err(refusal(
                    Code::UnsupportedCapability,
                    "a relocated Numbers table is appended after the destination sheet's tables",
                ));
            }
            relocate_table(&mut chain, sheet, table, to)?;
        },
        _ => return Err(element_refusal("move", &node)),
    }
    bound(LimitKind::Operations, 1, limits.max_operations())?;
    finish(
        file,
        chain.changed,
        Owner::Numbers(chain.package),
        chain.patches,
        1,
        limits,
    )
}

// ---------------------------------------------------------------------- swap

/// Exchange two sheets, or two tables (of one sheet or of two sheets).
///
/// The exchange equals two moves applied one at a time: the first element
/// takes the second's place, then the second takes the first's. Every step
/// is one leaf transaction on the previous step's package, so the swap
/// publishes all of them or none. Rows and columns are refused: the leaf's
/// topology transaction inserts and removes them but does not reorder them.
pub(crate) fn swap(
    file: &Iwork,
    package: &Package,
    left: &PathArg<'_>,
    right: &PathArg<'_>,
    limits: &SemanticLimits,
) -> Result<Commit> {
    if axis_target(file, left.path()).is_some() || axis_target(file, right.path()).is_some() {
        return Err(refusal(
            Code::UnsupportedCapability,
            "Numbers rows and columns are inserted and removed, not reordered; the leaf topology transaction has no row or column move (follow-up: record 0347 F1)",
        ));
    }
    let left = node_at::<Numbers>(file, left)?;
    let right = node_at::<Numbers>(file, right)?;
    let (left_kind, right_kind) = (left.item().type_name(), right.item().type_name());
    let [left_sheet, left_table, _, _] = left.item().leaf().map(|value| value as usize);
    let [right_sheet, right_table, _, _] = right.item().leaf().map(|value| value as usize);
    let mut chain = Chain::new(package.clone());
    match (left_kind, right_kind) {
        ("sheet", "sheet") => {
            let (first, second) = (left_sheet.min(right_sheet), left_sheet.max(right_sheet));
            if first != second {
                // The first sheet takes the second's place, which shifts the
                // second back by one; that one then takes the first's place.
                move_sheet(&mut chain, first, second)?;
                move_sheet(&mut chain, second - 1, first)?;
            }
        },
        ("table", "table") if left_sheet == right_sheet => {
            let (first, second) = (left_table.min(right_table), left_table.max(right_table));
            if first != second {
                order_table(&mut chain, left_sheet, first, second)?;
                order_table(&mut chain, left_sheet, second - 1, first)?;
            }
        },
        ("table", "table") => {
            // Two sheets: each table is relocated into the other's sheet,
            // where the leaf appends it, and then takes the other's place.
            // Table names are unique within a sheet, so two tables of one
            // name (`Table 1` in both) pass through a temporary name.
            let shared = table_name(package, left_sheet, left_table)
                .filter(|name| table_name(package, right_sheet, right_table) == Some(*name));
            if let Some(name) = shared {
                let taken = [left_sheet, right_sheet]
                    .iter()
                    .filter_map(|sheet| package.sheets().get(*sheet))
                    .flat_map(|sheet| sheet.tables().map(litchi_numbers::Table::name))
                    .collect::<Vec<_>>();
                let temporary = fresh_name(name, &taken);
                rename_table(&mut chain, left_sheet, left_table, &temporary)?;
            }
            relocate_table(&mut chain, left_sheet, left_table, right_sheet)?;
            relocate_table(&mut chain, right_sheet, right_table, left_sheet)?;
            let left_count = table_count(&chain.package, left_sheet);
            let right_count = table_count(&chain.package, right_sheet);
            order_table(&mut chain, left_sheet, left_count - 1, left_table)?;
            order_table(&mut chain, right_sheet, right_count - 1, right_table)?;
            if let Some(name) = shared {
                rename_table(&mut chain, right_sheet, right_table, name)?;
            }
        },
        ("sheet" | "table", "sheet" | "table") => {
            return Err(refusal(
                Code::InvalidArgument,
                "a Numbers swap exchanges two sheets or two tables",
            ));
        },
        ("sheet" | "table", _) => return Err(element_refusal("swap", &right)),
        _ => return Err(element_refusal("swap", &left)),
    }
    bound(LimitKind::Operations, 1, limits.max_operations())?;
    finish(
        file,
        chain.changed,
        Owner::Numbers(chain.package),
        chain.patches,
        1,
        limits,
    )
}

fn table_name(package: &Package, sheet: usize, table: usize) -> Option<&str> {
    package
        .sheets()
        .get(sheet)
        .and_then(|sheet| sheet.tables().nth(table))
        .map(litchi_numbers::Table::name)
}

/// Commit one table rename on the chain's package.
fn rename_table(chain: &mut Chain<Package>, sheet: usize, table: usize, name: &str) -> Result<()> {
    let error = |error| edit_error(EditError::Names(error));
    let commit = chain
        .package
        .edit_names()
        .rename_table(
            SheetSelector::index(sheet),
            TableSelector::index(table),
            name,
        )
        .map_err(error)?
        .commit()
        .map_err(error)?;
    let patch = commit.patch().clone();
    push(chain, commit.into_package(), Patch::Names(patch));
    Ok(())
}

fn table_count(package: &Package, sheet: usize) -> usize {
    package
        .sheets()
        .get(sheet)
        .map_or(0, |sheet| sheet.table_count())
}

/// Commit one sheet move on the chain's package.
fn move_sheet(chain: &mut Chain<Package>, from: usize, to: usize) -> Result<()> {
    let error = |error| edit_error(EditError::Order(error));
    let commit = chain
        .package
        .edit_sheet_order()
        .move_sheet(SheetSelector::index(from), to)
        .map_err(error)?
        .commit()
        .map_err(error)?;
    let patch = commit.patch().clone();
    push(chain, commit.into_package(), Patch::Order(patch));
    Ok(())
}

/// Commit one table move within its sheet on the chain's package.
fn order_table(chain: &mut Chain<Package>, sheet: usize, from: usize, to: usize) -> Result<()> {
    let error = |error| edit_error(EditError::TableOrder(error));
    let commit = chain
        .package
        .edit_table_order(SheetSelector::index(sheet))
        .map_err(error)?
        .move_table(TableSelector::index(from), to)
        .map_err(error)?
        .commit()
        .map_err(error)?;
    let (next, patch) = commit.into_parts();
    push(chain, next, Patch::TableOrder(patch));
    Ok(())
}

/// Commit one table relocation (appended to `to`'s tables) on the chain's
/// package.
fn relocate_table(chain: &mut Chain<Package>, sheet: usize, table: usize, to: usize) -> Result<()> {
    let error = |error| edit_error(EditError::Relocation(error));
    let commit = chain
        .package
        .edit_table_relocation(
            SheetSelector::index(sheet),
            TableSelector::index(table),
            SheetSelector::index(to),
        )
        .map_err(error)?
        .commit()
        .map_err(error)?;
    let patch = commit.patch().clone();
    push(chain, commit.into_package(), Patch::Relocation(patch));
    Ok(())
}

// -------------------------------------------------------------------- create

/// A fresh workbook: the leaf's minimal `Package::new()` (one sheet, one
/// empty 10 × 5 table, a fixed document identity).
pub(crate) fn create(options: &crate::CreateOptions) -> Result<Package> {
    if options.flavor != crate::Flavor::Document {
        return Err(refusal(
            Code::UnsupportedCapability,
            "a Numbers template is not authored; create an ordinary workbook",
        ));
    }
    if options.metadata.is_some() {
        return Err(refusal(
            Code::UnsupportedCapability,
            "iWork metadata writes are not supported (ADR 0034 §7); create the workbook without initial metadata",
        ));
    }
    Ok(Package::new())
}
