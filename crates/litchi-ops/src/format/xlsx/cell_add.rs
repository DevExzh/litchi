//! `add` of an XLSX worksheet cell and of a rich-text run (record 0338).
//!
//! As OfficeCLI's `AddCell`, the parent picks the cell:
//!
//! - a sheet (`/Sheet1`): `ref` (alias `address`) names the cell, else the
//!   first empty cell of row 1;
//! - a row (`/Sheet1/row[5]`): `ref`, else the row's first empty cell;
//! - a cell (`/Sheet1/B2`): that cell.
//!
//! The other properties are the cell `set` properties (`value`, `formula`,
//! `type`, `runs` and the formatting keys), written through the ordinary
//! cell write in one native edit. Unlike OfficeCLI, `add` never overwrites:
//! a cell that already holds a value or formula is refused, and `set`
//! changes it. A `ref` that disagrees with a cell or row parent is refused
//! instead of warned about.

use litchi_core::selector::{CellRef, Path, Predicate, Step};
use litchi_core::{FileFormat, Position};
use litchi_xlsx::Address;

use super::node::{Grid, Item, Node, Sheet, SheetHandle};
use super::write::Commit;
use super::{Nav, Xlsx};
use crate::error::{Code, Error, Result};
use crate::new_node::NewNode;
use crate::placement::Placement;
use crate::props::{PropValue, Props};
use crate::target::IntoTarget;
use crate::transaction::SemanticLimits;

/// The XLSX grid's last column, zero-based (`XFD`).
const LAST_COLUMN: u32 = (1 << 14) - 1;

impl Xlsx {
    /// Add one cell under `parent`.
    pub(super) fn add_cell(
        &self,
        parent: &Node<'_>,
        node: &NewNode<'_>,
        limits: &SemanticLimits,
    ) -> Result<(Commit, Path<'static>)> {
        let sheet = parent.sheet_position();
        let (reference, props) = split_ref(node.props())?;
        let from_parent = match parent.item() {
            Item::Cell(_) => parent.write_cell_target().map(|(_, at)| at),
            Item::Sheet(_) | Item::Row(_) => None,
            _ => {
                return Err(Error::new(
                    Code::UnsupportedType,
                    "an XLSX cell is added under a sheet, a row or a cell",
                )
                .with_format(FileFormat::Xlsx));
            },
        };
        let row = match parent.item() {
            Item::Row(row) => Some(row.index),
            _ => None,
        };
        let at = match (reference, from_parent) {
            (Some(at), Some(parent_at)) if at != parent_at => {
                return Err(disagree("`ref` names another cell than the parent path"));
            },
            (Some(at), _) => {
                if row.is_some_and(|row| row != at.row().get()) {
                    return Err(disagree(
                        "`ref` names a cell in another row than the parent",
                    ));
                }
                at
            },
            (None, Some(at)) => at,
            (None, None) => first_empty(parent, row.unwrap_or(0))?,
        };
        if props.is_empty() {
            return Err(Error::new(
                Code::MissingProperty,
                "a cell add needs a value, a formula, runs or formatting",
            )
            .with_property("value")
            .with_format(FileFormat::Xlsx));
        }
        if occupied(parent, at)? {
            return Err(Error::new(
                Code::InvalidArgument,
                "the cell already holds a value; add never overwrites it",
            )
            .with_format(FileFormat::Xlsx)
            .with_suggestion("set the cell's properties instead"));
        }
        let path = cell_path(sheet, at);
        let commit = self.set_with_limits(path.clone().into_target()?, props, limits)?;
        Ok((commit, path))
    }

    /// Add one run under the cell `parent`.
    pub(super) fn add_run_node(
        &self,
        parent: &Node<'_>,
        node: &NewNode<'_>,
        limits: &SemanticLimits,
    ) -> Result<(Commit, Path<'static>)> {
        let index = match node.placement() {
            None | Some(Placement::End) => None,
            Some(Placement::Start) => Some(0),
            Some(Placement::Position(position)) => Some(position.get()),
            Some(Placement::Before(path) | Placement::After(path)) => {
                let anchor = crate::path::resolve::get(&Nav::new(self), path)?;
                let Item::Run(run) = anchor.item() else {
                    return Err(Error::new(
                        Code::UnsupportedType,
                        "a run is placed before or after another run of the cell",
                    )
                    .with_format(FileFormat::Xlsx));
                };
                let index = run.index as usize;
                Some(if matches!(node.placement(), Some(Placement::After(_))) {
                    index + 1
                } else {
                    index
                })
            },
            Some(Placement::Into(..)) => {
                return Err(Error::new(
                    Code::InvalidArgument,
                    "a run is added under its cell path, not into another parent",
                )
                .with_format(FileFormat::Xlsx));
            },
        };
        let (commit, placed) = self.add_run(parent, node.props(), index, limits)?;
        let Some((sheet, at)) = parent.write_cell_target() else {
            return Err(Error::new(Code::InternalError, "a run parent is a cell"));
        };
        let mut path = cell_path(sheet, at);
        path.push(Step::new("run").with(Predicate::Index(Position::new(placed))))
            .map_err(|error| Error::new(Code::InvalidPath, error.to_string()))?;
        Ok((commit, path))
    }
}

fn disagree(message: &'static str) -> Error {
    Error::new(Code::InvalidArgument, message)
        .with_property("ref")
        .with_format(FileFormat::Xlsx)
}

/// `ref` (or `address`) and the other properties.
fn split_ref<'a>(props: &Props<'a>) -> Result<(Option<Address>, Props<'a>)> {
    let mut reference = None;
    let mut rest = Props::new();
    for (key, value) in props.iter() {
        if key.as_str().eq_ignore_ascii_case("ref") || key.as_str().eq_ignore_ascii_case("address")
        {
            if reference.is_some() {
                return Err(Error::new(
                    Code::InvalidArgument,
                    "the cell address was given more than once",
                )
                .with_property("ref")
                .with_format(FileFormat::Xlsx));
            }
            let PropValue::Text(text) = value else {
                return Err(bad_ref());
            };
            let cell = CellRef::parse(text.trim()).ok_or_else(bad_ref)?;
            reference = Some(Address::at(cell.row(), cell.column()).map_err(|_| bad_ref())?);
            continue;
        }
        rest.push(key.clone(), value.clone());
    }
    Ok((reference, rest))
}

fn bad_ref() -> Error {
    Error::new(Code::InvalidValue, "ref is one A1 cell reference, e.g. B2")
        .with_property("ref")
        .with_format(FileFormat::Xlsx)
}

/// The first column of `row` without a stored record (OfficeCLI's
/// auto-assign).
fn first_empty(parent: &Node<'_>, row: u32) -> Result<Address> {
    let grid = parent_grid(parent)?;
    for column in 0..=LAST_COLUMN {
        let at = Address::at(row, column).map_err(super::children::coordinate)?;
        if grid.record(at)?.is_none() {
            return Ok(at);
        }
    }
    Err(Error::new(Code::OutOfRange, "the row has no empty cell").with_format(FileFormat::Xlsx))
}

/// Whether the cell at `at` holds a value or formula.
fn occupied(parent: &Node<'_>, at: Address) -> Result<bool> {
    let grid = parent_grid(parent)?;
    Ok(grid.record(at)?.is_some_and(super::node::has_content))
}

/// The grid of the parent's worksheet.
fn parent_grid(parent: &Node<'_>) -> Result<Grid> {
    Ok(match parent.item() {
        Item::Sheet(sheet) => sheet.grid()?.clone(),
        Item::Row(row) => row.grid.clone(),
        _ => {
            let file = parent.file();
            let position = parent.sheet_position();
            let handle = SheetHandle::at(file, position)?.ok_or_else(|| {
                Error::new(
                    Code::OutOfRange,
                    "the XLSX worksheet position is out of range",
                )
                .with_format(FileFormat::Xlsx)
            })?;
            Sheet::new(file, position, handle).grid()?.clone()
        },
    })
}

fn cell_path(sheet: usize, at: Address) -> Path<'static> {
    Path::from_steps([
        Step::new("sheet").with(Predicate::Index(Position::new(sheet))),
        Step::new(at.a1()),
    ])
}
