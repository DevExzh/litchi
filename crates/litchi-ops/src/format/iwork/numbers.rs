//! Apple Numbers workbooks (`.numbers`): sheets, their tables and the
//! tables' stored cells, read from `litchi_numbers::Package` and written
//! through its table-cell transaction (record 0296).
//!
//! A Numbers sheet holds tables, and cells belong to a table: a cell's path
//! is `/sheet[N]/table[M]/A1`, or by name `/Sheet 1/Table 1/A1`. The tree
//! holds the cells the table stores; `set` also reaches an empty cell inside
//! the table's grid.

use std::collections::BTreeMap;

use litchi_core::FileFormat;
use litchi_core::selector::{CellRef, Path};
use litchi_numbers::cell::Value as CellValue;
use litchi_numbers::table::CellPosition;
use litchi_numbers::table::cells::{Edit, Input};
use litchi_numbers::{Package, SheetSelector, TableSelector};

use super::tree::{count, text};
use super::write::{
    Chain, Commit, LeafPatch, bound, element_refusal, finish, key_refusal, leaf_refusal, node_at,
    refusal, resolve,
};
use super::{Flavor, Iwork, LeafError, Owner, TreeBuilder, element, writable};
use crate::Result;
use crate::error::Code;
use crate::node::spec::{FormatSpec, GapSpec, PropSpec, Subject};
use crate::node::{PathStep, Value};
use crate::path::IntoPath;
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
    }
}

const WORKBOOK_PROPS: &[PropSpec] = &[PropSpec::new("sheetCount", ValueType::Int)];

pub(crate) const SHEET_PROPS: &[PropSpec] = &[
    PropSpec::new("name", ValueType::Text),
    PropSpec::new("tableCount", ValueType::Int),
];

pub(crate) const TABLE_PROPS: &[PropSpec] = &[
    PropSpec::new("name", ValueType::Text),
    PropSpec::new("rowCount", ValueType::Int),
    PropSpec::new("columnCount", ValueType::Int),
];

pub(crate) const CELL_PROPS: &[PropSpec] = &[
    PropSpec::new("type", ValueType::Text),
    writable("value", ValueType::Scalar),
    PropSpec::new("formula", ValueType::Text),
];

static GAPS: [GapSpec; 3] = [
    GapSpec {
        subject: Subject::Element("row"),
        code: Code::UnsupportedCapability,
        follow_up: "0299 Numbers rows and columns (IWORK_API_REQUESTS.md request 6)",
        reason: "rows and columns are table coordinates here; the leaf's row and column insertion and deletion are not package transactions yet",
    },
    GapSpec {
        subject: Subject::Property {
            element: "cell",
            key: "font.bold",
        },
        code: Code::UnsupportedCapability,
        follow_up: "IWORK_API_REQUESTS.md request 7",
        reason: "the Numbers leaf has no cell text-style transaction",
    },
    GapSpec {
        subject: Subject::Operation {
            verb: "set",
            element: "cell",
            context: "a formula",
        },
        code: Code::UnsupportedCapability,
        follow_up: "0299 Numbers formula writes",
        reason: "the leaf's formula authoring is not wired; `value` writes text, numbers and booleans",
    },
];

static OPERATION_GAPS: [GapSpec; 3] = [
    GapSpec {
        subject: Subject::Operation {
            verb: "add",
            element: "sheet",
            context: "a new sheet",
        },
        code: Code::UnsupportedCapability,
        follow_up: "IWORK_API_REQUESTS.md request 4",
        reason: "the Numbers leaf has no sheet lifecycle transaction",
    },
    GapSpec {
        subject: Subject::Operation {
            verb: "add",
            element: "table",
            context: "a new table",
        },
        code: Code::UnsupportedCapability,
        follow_up: "IWORK_API_REQUESTS.md request 5",
        reason: "the Numbers leaf has no table lifecycle transaction",
    },
    GapSpec {
        subject: Subject::Operation {
            verb: "create",
            element: "workbook",
            context: "a workbook from scratch",
        },
        code: Code::UnsupportedCapability,
        follow_up: "IWORK_API_REQUESTS.md request 1",
        reason: "the Numbers leaf has no from-scratch package builder",
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
            &[Kind::Cell],
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

/// The elements `set` writes.
pub(crate) static WRITES: super::Writes = super::Writes {
    settable: &["cell"],
    addable: &[],
    removable: &[],
    movable: &[],
    gaps: &OPERATION_GAPS,
};

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

/// Build the tree of `package`.
pub(crate) fn build(package: &Package) -> super::Tree {
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
            let table_item = tree
                .push(sheet_item, Kind::Table)
                .prop("name", text(Some(table.name())))
                .prop("rowCount", count(table.row_count() as usize))
                .prop("columnCount", count(table.column_count() as usize))
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
                let mut draft = tree
                    .push_step(
                        table_item,
                        Kind::Cell,
                        PathStep::cell(position.column(), position.row()),
                    )
                    .prop("type", text(Some(kind)))
                    .prop("value", value)
                    .prop("formula", formula)
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
        }
    }
    tree.build()
}

/// One cell a set addresses: sheet, table and zero-based coordinates.
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord)]
struct Address {
    sheet: usize,
    table: usize,
    row: u32,
    column: u32,
}

/// The cells `target` addresses. A path to an empty cell inside a table's
/// grid addresses that cell although the tree stores no item for it.
fn addresses(
    file: &Iwork,
    target: &TargetArg<'_>,
    limits: &SemanticLimits,
) -> Result<Vec<Address>> {
    let resolved = resolve::<Numbers>(file, target, limits.max_expanded_targets());
    let nodes = match resolved {
        Ok(nodes) => nodes,
        Err(error) if error.code() == Code::PathNotFound => {
            let Some(path) = target.path() else {
                return Err(error);
            };
            return empty_cell(file, path.path())
                .map(|address| vec![address])
                .ok_or(error);
        },
        Err(error) => return Err(error),
    };
    let mut cells = Vec::with_capacity(nodes.len());
    for node in &nodes {
        if node.item().type_name() != "cell" {
            return Err(element_refusal("set", node));
        }
        let [sheet, table, row, column] = node.item().leaf();
        cells.push(Address {
            sheet: sheet as usize,
            table: table as usize,
            row,
            column,
        });
    }
    Ok(cells)
}

/// The empty cell `path` names: a table path and an A1 step inside the
/// table's grid.
fn empty_cell(file: &Iwork, path: &Path<'_>) -> Option<Address> {
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
    let rows = table.value("rowCount").and_then(|value| match value {
        Value::Int(rows) => u32::try_from(*rows).ok(),
        _ => None,
    })?;
    let columns = table.value("columnCount").and_then(|value| match value {
        Value::Int(columns) => u32::try_from(*columns).ok(),
        _ => None,
    })?;
    (cell.row() < rows && cell.column() < columns).then_some(Address {
        sheet: sheet as usize,
        table: table_index as usize,
        row: cell.row(),
        column: cell.column(),
    })
}

/// The leaf input of a cell `value`; `None` clears the cell.
fn input(value: &PropValue<'_>) -> Result<Option<Input>> {
    let staging = |error| leaf_refusal(LeafError::NumbersCells(error), FileFormat::Numbers);
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

/// Set cell values. Cells of one table stage on one leaf transaction; the
/// tables commit in document order.
pub(crate) fn set(
    file: &Iwork,
    package: &Package,
    target: &TargetArg<'_>,
    props: &Props<'_>,
    limits: &SemanticLimits,
) -> Result<Commit> {
    let cells = addresses(file, target, limits)?;
    let mut inputs = Vec::with_capacity(props.len());
    for (key, value) in props.iter() {
        if key.as_str() != "value" {
            let first = resolve::<Numbers>(file, target, 1).ok();
            return Err(match first.as_ref().and_then(|nodes| nodes.first()) {
                Some(node) => key_refusal(node, key.as_str()),
                None => super::write::unwritable("cell", key.as_str()),
            });
        }
        inputs.push(input(value)?);
    }
    let mut tables: BTreeMap<(usize, usize), Vec<(CellPosition, Option<Input>)>> = BTreeMap::new();
    let mut operations = 0_usize;
    for cell in &cells {
        for input in &inputs {
            operations += 1;
            bound(LimitKind::Operations, operations, limits.max_operations())?;
            tables
                .entry((cell.sheet, cell.table))
                .or_default()
                .push((CellPosition::new(cell.row, cell.column), input.clone()));
        }
    }
    let error = |error| leaf_refusal(LeafError::NumbersCells(error), FileFormat::Numbers);
    let mut chain = Chain::new(package.clone());
    for ((sheet, table), changes) in tables {
        let mut edit: Edit<'_> = chain
            .package
            .edit_table_cells(SheetSelector::index(sheet), TableSelector::index(table))
            .map_err(error)?;
        for (position, input) in changes {
            edit = match input {
                Some(input) => edit.set(position, input),
                None => edit.clear(position),
            }
            .map_err(error)?;
        }
        let commit = edit.commit().map_err(error)?;
        let patch = commit.patch().clone();
        let changed = !patch.is_noop();
        chain.push(
            commit.into_package(),
            LeafPatch::NumbersCells(patch),
            changed,
        );
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
