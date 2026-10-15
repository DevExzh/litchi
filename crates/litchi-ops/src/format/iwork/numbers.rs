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
//! | `set` | `cell` `value`/`formula` | `edit_table_cells` (one per table) |
//! | `set` | `cell` `numberFormat` | `edit_table_cell_{number,percentage}_format` |
//! | `set` | `sheet`/`table` `name` | `edit_names` (one for every rename) |
//! | `add` | `mergeCell` | `edit_table_merges` |
//! | `remove` | `cell` (clears) / `mergeCell` | `edit_table_cells` / `edit_table_merges` |
//! | `move` | `sheet` | `edit_sheet_order` |
//! | `move` | `table` into another sheet | `edit_table_relocation` |
//!
//! Sheet, table, row and column lifecycle stays a typed gap (requests 4–6).

pub(crate) mod coalesce;

use std::collections::BTreeMap;
use std::fmt;

use litchi_core::selector::{CellRef, Path, Predicate, Step};
use litchi_core::{FileFormat, Position};
use litchi_numbers::cell::Value as CellValue;
use litchi_numbers::cell::data_format::{
    DecimalPlaces, NegativeStyle, Number, Percentage, ThousandsSeparator,
};
use litchi_numbers::formula::Expression;
use litchi_numbers::table::cells::{Change, Edit, Input};
use litchi_numbers::table::merge::Region;
use litchi_numbers::table::{CellPosition, CellRange};
use litchi_numbers::{Package, SheetSelector, TableSelector};

use super::tree::{count, text};
use super::write::{
    Chain, Commit, LeafPatch, bound, destination, element_refusal, finish, key_refusal,
    leaf_refusal, node_at, refusal, resolve,
};
use super::{Flavor, Iwork, LeafError, Node, Owner, TreeBuilder, element, writable};
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

pub(crate) const SHEET_PROPS: &[PropSpec] = &[
    writable("name", ValueType::Text),
    PropSpec::new("tableCount", ValueType::Int),
];

pub(crate) const TABLE_PROPS: &[PropSpec] = &[
    writable("name", ValueType::Text),
    PropSpec::new("rowCount", ValueType::Int),
    PropSpec::new("columnCount", ValueType::Int),
    PropSpec::new("mergeCount", ValueType::Int),
];

pub(crate) const CELL_PROPS: &[PropSpec] = &[
    PropSpec::new("type", ValueType::Text),
    writable("value", ValueType::Scalar),
    writable("formula", ValueType::Text),
    PropSpec::new("merge", ValueType::Text),
    PropSpec::new("mergeAnchor", ValueType::Bool),
    set_only("numberFormat", ValueType::Text),
];

pub(crate) const MERGE_PROPS: &[PropSpec] = &[{
    let mut spec = PropSpec::new("ref", ValueType::Text);
    spec.access.add = true;
    spec
}];

const FOLLOW_UP: &str = "record 0297 follow-ups (IWORK_API_REQUESTS.md)";

static GAPS: [GapSpec; 4] = [
    GapSpec {
        subject: Subject::Element("row"),
        code: Code::UnsupportedCapability,
        follow_up: "IWORK_API_REQUESTS.md request 6",
        reason: "rows and columns are table coordinates here; the leaf has no package transaction that inserts or deletes them (the migration host's editor does, outside the supported crates)",
    },
    GapSpec {
        subject: Subject::Property {
            element: "cell",
            key: "font.bold",
        },
        code: Code::UnsupportedCapability,
        follow_up: "IWORK_API_REQUESTS.md request 7",
        reason: "the Numbers leaf has no cell text-style transaction; `numberFormat` covers decimal and percentage display formats only",
    },
    GapSpec {
        subject: Subject::Property {
            element: "cell",
            key: "comment",
        },
        code: Code::UnsupportedCapability,
        follow_up: FOLLOW_UP,
        reason: "the leaf's cell-comment transaction refuses every natively authored corpus package (`InvalidSource`/`UnsupportedDependency`: author storage and comment-list ownership outside its admitted scope), and comments are resolved per cell, so neither reads nor writes are claimed",
    },
    GapSpec {
        subject: Subject::Operation {
            verb: "set",
            element: "cell",
            context: "cross-table formula references",
        },
        code: Code::UnsupportedCapability,
        follow_up: FOLLOW_UP,
        reason: "formula text is parsed with local A1 references and the authorable function subset; references to other tables need source-bound table handles",
    },
];

static OPERATION_GAPS: [GapSpec; 5] = [
    GapSpec {
        subject: Subject::Operation {
            verb: "add",
            element: "sheet",
            context: "a new sheet",
        },
        code: Code::UnsupportedCapability,
        follow_up: "IWORK_API_REQUESTS.md request 4",
        reason: "the Numbers leaf has no sheet creation transaction",
    },
    GapSpec {
        subject: Subject::Operation {
            verb: "remove",
            element: "sheet",
            context: "a sheet and its graph",
        },
        code: Code::UnsupportedCapability,
        follow_up: "IWORK_API_REQUESTS.md request 4",
        reason: "the Numbers leaf has no sheet deletion transaction",
    },
    GapSpec {
        subject: Subject::Operation {
            verb: "add",
            element: "table",
            context: "a new table",
        },
        code: Code::UnsupportedCapability,
        follow_up: "IWORK_API_REQUESTS.md request 5",
        reason: "the Numbers leaf has no table creation transaction",
    },
    GapSpec {
        subject: Subject::Operation {
            verb: "remove",
            element: "table",
            context: "a table and its storage",
        },
        code: Code::UnsupportedCapability,
        follow_up: "IWORK_API_REQUESTS.md request 5",
        reason: "the Numbers leaf has no table deletion transaction",
    },
    GapSpec {
        subject: Subject::Operation {
            verb: "move",
            element: "table",
            context: "a position within its sheet",
        },
        code: Code::UnsupportedCapability,
        follow_up: FOLLOW_UP,
        reason: "the leaf relocates a table into another sheet (appended there); it has no transaction that reorders tables within one sheet",
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
            &[Kind::Table],
            &["/sheet[N]", "/<sheet name>"],
            SHEET_PROPS,
        ),
        element(
            Kind::Table,
            "table",
            &["table"],
            &[Kind::Sheet],
            &[Kind::Cell, Kind::MergeCell],
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
    settable: &["cell", "sheet", "table"],
    addable: &["mergeCell"],
    removable: &["cell", "mergeCell"],
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
        })
    }

    /// The unified code: selectors that match nothing are `path_not_found`,
    /// refused values `invalid_value`, locked tables `document_protected`,
    /// native shapes the leaf does not rewrite `unsupported_capability`.
    pub(crate) fn code(&self) -> Code {
        macro_rules! common {
            ($error:expr, $module:path, { $($arm:tt)* }) => {{
                use $module as E;
                match $error {
                    E::SheetNotFound => Code::PathNotFound,
                    E::LimitExceeded { .. } => Code::LimitExceeded,
                    E::UnsupportedSource { .. } => Code::UnsupportedCapability,
                    E::InvalidSource { .. } | E::PatchConflict => Code::InvalidInput,
                    $($arm)*
                    _ => Code::InternalError,
                }
            }};
        }
        match self {
            Self::Names(error) => common!(error, litchi_numbers::names::Error, {
                E::TableNotFound => Code::PathNotFound,
                E::InvalidName { .. } => Code::InvalidValue,
                E::DuplicateName { .. } | E::DuplicateTarget { .. } => Code::InvalidValue,
                E::TableLocked { .. } => Code::DocumentProtected,
                E::UnsupportedDependency { .. } => Code::UnsupportedCapability,
            }),
            Self::Order(error) => common!(error, litchi_numbers::sheet::order::Error, {
                E::DestinationOutOfRange { .. } => Code::OutOfRange,
            }),
            Self::Relocation(error) => common!(error, litchi_numbers::table::relocation::Error, {
                E::TableNotFound => Code::PathNotFound,
                E::TableLocked { .. } => Code::DocumentProtected,
            }),
            Self::Merges(error) => {
                common!(error, litchi_numbers::table::merge::transaction::Error, {
                    E::TableNotFound => Code::PathNotFound,
                    E::AmbiguousSelector => Code::InvalidPath,
                    E::OverlappingRegion | E::InvalidRegion => Code::InvalidValue,
                })
            },
            Self::NumberFormat(error) => common!(
                error,
                litchi_numbers::cell::data_format::number::transaction::Error,
                {
                    E::TableNotFound => Code::PathNotFound,
                    E::CellNotFound => Code::PathNotFound,
                    E::TableLocked { .. } => Code::DocumentProtected,
                    E::WrongFormatFamily { .. } | E::UnsupportedDependency { .. } => {
                        Code::UnsupportedCapability
                    },
                }
            ),
            Self::PercentageFormat(error) => common!(
                error,
                litchi_numbers::cell::data_format::percentage::transaction::Error,
                {
                    E::TableNotFound => Code::PathNotFound,
                    E::CellNotFound => Code::PathNotFound,
                    E::TableLocked { .. } => Code::DocumentProtected,
                    E::WrongFormatFamily { .. } | E::UnsupportedDependency { .. } => {
                        Code::UnsupportedCapability
                    },
                }
            ),
            Self::Formula(error) => match error {
                litchi_numbers::formula::Error::LimitExceeded { .. } => Code::LimitExceeded,
                litchi_numbers::formula::Error::UnsupportedFunction => Code::UnsupportedCapability,
                _ => Code::InvalidValue,
            },
            Self::Create(error) => match error {
                litchi_numbers::create::Error::MissingIdentity
                | litchi_numbers::create::Error::IdentityNotDistinct => Code::InvalidArgument,
                _ => Code::InternalError,
            },
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

/// Build the tree of `package`, reading each table's merged regions.
///
/// # Errors
///
/// The leaf's typed merge-reader error when a table's merge storage cannot
/// be read.
pub(crate) fn build(package: &Package) -> Result<super::Tree> {
    let merges_error = |error| edit_error(EditError::Merges(error));
    let mut tree = TreeBuilder::new(SPEC.elements, Numbers::SINGLETONS);
    let sheets = package.sheets();
    tree.root().prop("sheetCount", count(sheets.len())).done();
    for (sheet_index, sheet) in sheets.iter().enumerate() {
        let sheet_item = tree
            .push(super::tree::ROOT, Kind::Sheet)
            .prop("name", text(Some(sheet.name())))
            .prop("tableCount", count(sheet.table_count()))
            .label(Some(sheet.name()))
            .leaf([sheet_index, 0, 0, 0])
            .done();
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
                .done();
            for cell in table.iter_cells() {
                let position = cell.position();
                let (kind, value, display) = cell_value(cell.value());
                let formula = match cell.value() {
                    CellValue::Formula(formula) => text(Some(formula)),
                    _ => None,
                };
                let region = regions.iter().find(|region| {
                    let (row, column) = (position.row(), position.column());
                    row >= region.row()
                        && row < region.row() + region.row_count()
                        && column >= region.column()
                        && column < region.column() + region.column_count()
                });
                let mut draft = tree
                    .push_step(
                        table_item,
                        Kind::Cell,
                        PathStep::cell(position.column(), position.row()),
                    )
                    .prop("type", text(Some(kind)))
                    .prop("value", value)
                    .prop("formula", formula)
                    .prop("merge", region.map(|region| region_ref(*region).into()))
                    .prop(
                        "mergeAnchor",
                        region.map(|region| Value::Bool(region.start() == position)),
                    )
                    .leaf([
                        sheet_index,
                        table_index,
                        position.row() as usize,
                        position.column() as usize,
                    ]);
                if let Some(display) = display {
                    draft = draft.text(display);
                }
                draft.done();
            }
            for region in &regions {
                tree.push(table_item, Kind::MergeCell)
                    .prop("ref", text(Some(&region_ref(*region))))
                    .leaf([
                        sheet_index,
                        table_index,
                        region.row() as usize,
                        region.column() as usize,
                    ])
                    .done();
            }
        }
    }
    Ok(tree.build())
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
    const fn position(self) -> CellPosition {
        CellPosition::new(self.row, self.column)
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
        table.value(key).and_then(|value| match value {
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
}

impl CellWrite {
    pub(crate) fn change(&self, position: CellPosition) -> Change {
        match self {
            Self::Value(Some(input)) => Change::set(position, input.clone()),
            Self::Value(None) => Change::clear(position),
            Self::Formula(expression) => Change::set_formula(position, expression.clone()),
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
pub(crate) fn formula_write(value: &PropValue<'_>) -> Result<CellWrite> {
    let text = super::write::text(value, "formula")?;
    match text {
        None => Ok(CellWrite::Value(None)),
        Some(text) => Expression::parse(&text)
            .map(CellWrite::Formula)
            .map_err(|error| {
                edit_error(EditError::Formula(error)).with_property("formula".to_owned())
            }),
    }
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

/// The writes of one `set` call, grouped by leaf transaction.
#[derive(Default)]
struct Staged {
    renames_sheets: Vec<(usize, String)>,
    renames_tables: Vec<(usize, usize, String)>,
    cells: BTreeMap<(usize, usize), Vec<Change>>,
    formats: Vec<(Address, DisplayFormat)>,
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
    let mut staged = Staged::default();
    let mut operations = 0_usize;
    let stage_cell =
        |staged: &mut Staged, address: Address, key: &str, value: &PropValue<'_>| -> Result<()> {
            if let Some(write) = cell_write(key, value)? {
                staged
                    .cells
                    .entry((address.sheet, address.table))
                    .or_default()
                    .push(write.change(address.position()));
            } else if key.eq_ignore_ascii_case("numberformat") {
                staged.formats.push((address, display_format(value)?));
            } else {
                return Err(property_gap("cell", key)
                    .unwrap_or_else(|| super::write::unwritable("cell", key)));
            }
            Ok(())
        };
    match targets(file, target, limits)? {
        Targets::EmptyCell(address) => {
            for (key, value) in props.iter() {
                operations += 1;
                bound(LimitKind::Operations, operations, limits.max_operations())?;
                stage_cell(&mut staged, address, key.as_str(), value)?;
            }
        },
        Targets::Nodes(nodes) => {
            for node in &nodes {
                for (key, value) in props.iter() {
                    operations += 1;
                    bound(LimitKind::Operations, operations, limits.max_operations())?;
                    let key = key.as_str();
                    let [sheet, table, _, _] = node.item().leaf().map(|value| value as usize);
                    match node.item().type_name() {
                        "cell" => stage_cell(&mut staged, cell_address(node), key, value)?,
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

fn parse_region(range: &str) -> Result<Region> {
    let invalid = || {
        refusal(
            Code::InvalidValue,
            "`ref` must be an A1 range of at least two cells (`B2:C3`)",
        )
        .with_property("ref".to_owned())
    };
    let range = CellRange::from_a1(range).map_err(|_| invalid())?;
    let start = range.start();
    let end = range.end();
    Region::new(
        start.row(),
        start.column(),
        // `CellRange` is half-open: its end is exclusive.
        end.row().saturating_sub(start.row()),
        end.column().saturating_sub(start.column()),
    )
    .map_err(|_| invalid())
}

/// The typed refusal of an `add` the leaf cannot perform.
fn add_gap(kind: &str) -> crate::Error {
    let message = if kind.eq_ignore_ascii_case("sheet") {
        "adding a Numbers sheet needs a leaf sheet-creation transaction (IWORK_API_REQUESTS.md request 4)"
    } else if kind.eq_ignore_ascii_case("table") {
        "adding a Numbers table needs a leaf table-creation transaction (IWORK_API_REQUESTS.md request 5)"
    } else if ["row", "col", "column"]
        .iter()
        .any(|name| kind.eq_ignore_ascii_case(name))
    {
        "inserting Numbers rows or columns needs a leaf topology transaction (IWORK_API_REQUESTS.md request 6)"
    } else {
        "a Numbers workbook adds merged ranges (`mergeCell`) only"
    };
    refusal(Code::UnsupportedCapability, message)
}

/// Remove targets: a cell is cleared, a merged range unmerged.
pub(crate) fn remove(
    file: &Iwork,
    package: &Package,
    target: &TargetArg<'_>,
    limits: &SemanticLimits,
) -> Result<Commit> {
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
                let range = match node.value("ref") {
                    Some(Value::Text(range)) => range.to_string(),
                    _ => return Err(element_refusal("remove", node)),
                };
                merges
                    .entry((sheet, table))
                    .or_default()
                    .push(parse_region(&range)?);
            },
            "sheet" => {
                return Err(refusal(
                    Code::UnsupportedCapability,
                    "removing a Numbers sheet needs a leaf sheet-deletion transaction (IWORK_API_REQUESTS.md request 4)",
                ));
            },
            "table" => {
                return Err(refusal(
                    Code::UnsupportedCapability,
                    "removing a Numbers table needs a leaf table-deletion transaction (IWORK_API_REQUESTS.md request 5)",
                ));
            },
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
            let error = |error| edit_error(EditError::Order(error));
            let commit = package
                .edit_sheet_order()
                .move_sheet(SheetSelector::index(from), to)
                .map_err(error)?
                .commit()
                .map_err(error)?;
            let patch = commit.patch().clone();
            push(&mut chain, commit.into_package(), Patch::Order(patch));
        },
        "table" => {
            let crate::Placement::Into(destination_path, position) = placement else {
                return Err(refusal(
                    Code::UnsupportedCapability,
                    "a Numbers table moves into another sheet (`Placement::Into`), where it is appended; the leaf does not reorder tables within a sheet",
                ));
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
            let error = |error| edit_error(EditError::Relocation(error));
            let commit = package
                .edit_table_relocation(
                    SheetSelector::index(sheet),
                    TableSelector::index(table),
                    SheetSelector::index(to),
                )
                .map_err(error)?
                .commit()
                .map_err(error)?;
            let patch = commit.patch().clone();
            push(&mut chain, commit.into_package(), Patch::Relocation(patch));
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
