//! Implementation of the public [`crate::table::topology`] transactions:
//! inserting and removing table rows and columns (`IWORK_API_REQUESTS.md`
//! request 6, record 0309).
//!
//! The structural rewrite is the shared table-topology engine
//! (`litchi_iwa_table::topology`, lifted from this module in record 0412e so
//! Pages and Keynote run it too); this module keeps the Numbers selectors,
//! the comment and cell clears of a removed axis, and the verification.
//!
//! An axis edit moves every cell after the edited index. One edit rewrites,
//! in one staged working set over the exact source:
//!
//! - the stored rows of every storage tile (a row edit renumbers the rows
//!   after the index; a column edit splices each row's current and pre-BNC
//!   cell-offset tables, so cell payloads keep their bytes);
//! - the row or column header buckets (sizes, hidden states and styles move
//!   with their row or column);
//! - the row/column UID map (an inserted axis gets a fresh identity above
//!   every existing one, as the host editor allocates it);
//! - the stroke sidecar's border layers (layers on the edited axis move,
//!   perpendicular runs shift, split or shorten);
//! - the table model's row or column count and, for a header or footer
//!   section edit, its section count;
//! - the table formula owner's table extents (the drawable's layout cache
//!   is left for the application to rebuild, as the host editor leaves it);
//! - every formula coordinate of the calculation engine that addresses the
//!   table (the engine's formula shift, record 0318): the table's own formulas and their
//!   dependency records, and other tables' formulas that read it.
//!
//! A removed row or column first has its comments cleared through the
//! comment transaction and its stored cells cleared through the
//! [`crate::table::cells`] transaction (value, style and list references are
//! released there, and the caches of formulas that read them refreshed), so
//! the structural pass only drops empty storage. Where the cell writer
//! cannot clear them, the structural pass drops the cells itself, with
//! their strings, formulas and dependency records (record 0329). A
//! surviving formula that reads a removed cell reads `#REF!` there and
//! caches an error result; a range that loses a row or column shrinks; a
//! shared formula whose cells diverge is split into one entry per formula.
//!
//! Ported from the host editor's `row_insert`, `column_insert`,
//! `table_delete` and `formula_dependency_shift` phases. Refused with a
//! typed [`Unsupported`] reason so a failed edit never publishes: a chart or
//! other non-table owner reading the table, volatile, spanning, error or
//! spill dependency state, a stale-cache range over cells the cell writer
//! cannot clear, a removed range among several of one tract, shared
//! formulas that diverge in a segmented list, merged cells, filters,
//! categories, pivots, column edits under sort rules, removing a hidden row
//! or column, and removing a row or column that carries a border layer.

#![allow(
    clippy::arbitrary_source_item_ordering,
    reason = "the focused transaction keeps its public types beside the staged rewrite"
)]

pub(super) use litchi_iwa_table::structure::wire;

pub(super) use litchi_numbers_wire::table_topology::{extents as calc, storage};

use std::fmt;

use litchi_iwa_table::topology::{self as engine, Operation, Step};
use thiserror::Error as ThisError;

#[cfg(test)]
use self::wire::single_bytes;
use self::wire::single_reference;
use super::Package;
use super::structure::{self, DOCUMENT_SHEETS_FIELD, FaultResult, Published, Working, references};
use crate::selector::{SheetSelector, TableSelector};
use crate::table::At;
use crate::table::topology::{ColumnDeletion, ColumnInsertion, RowDeletion, RowInsertion};

pub use litchi_iwa_table::topology::Unsupported;
pub(super) use litchi_iwa_table::topology::{Geometry, Target};

structure::structural_family!("table-topology");

/// `TST.Tile`.
pub(super) const TILE_MESSAGE_TYPE: u32 = 6_002;

/// A typed, content-redacted table-topology failure.
#[derive(Debug, Clone, Copy, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Error {
    /// A well-formed request this route refuses; [`Refusal::kind`] classifies it
    /// (record 0351 §2.4, record 0361).
    #[error(transparent)]
    Refused(Refusal),
    /// The rooted source is malformed or ambiguous.
    #[error("rooted table-topology source is invalid")]
    InvalidSource,
    /// A finite transaction limit was exceeded.
    #[error("{kind} limit exceeded: observed {observed}, maximum {maximum}")]
    Limit {
        /// Resource category.
        kind: LimitKind,
        /// Observed amount.
        observed: u64,
        /// Configured maximum.
        maximum: u64,
    },
    /// A bounded allocation failed.
    #[error("could not allocate {amount} table-topology items or bytes")]
    Allocation {
        /// Requested amount.
        amount: usize,
    },
    /// The reopened candidate did not reproduce the requested state.
    #[error("table-topology candidate verification failed")]
    Verification,
    /// The patch was applied to an artifact other than its exact source.
    #[error("table-topology patch conflicts with this package")]
    PatchConflict,
}

/// A well-formed request this family's route refuses, classified by
/// [`Refusal::kind`] (record 0351 §2.4, record 0361).
#[derive(Debug, Clone, Copy, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Refusal {
    /// The sheet selector did not resolve one sheet.
    #[error("sheet selector did not resolve")]
    SheetNotFound,
    /// The table selector did not resolve one table of the sheet.
    #[error("table selector did not resolve")]
    TableNotFound,
    /// The section-relative index is outside its section.
    #[error("index {index} is outside a section of {length}")]
    PositionOutOfRange {
        /// Requested section-relative index.
        index: usize,
        /// The section's length (an insertion may equal it).
        length: usize,
    },
    /// A header or footer section would exceed Numbers' five rows or
    /// columns.
    #[error("a header or footer section holds at most five rows or columns")]
    SectionFull,
    /// The edit would remove the table's last row or column, or a body row
    /// or column its header and footer sections need.
    #[error("the table must keep at least one row and one column")]
    LastAxis,
    /// The edit touches table content this transaction does not rewrite.
    #[error("table topology edits are not supported for {reason}")]
    Unsupported {
        /// What the table holds.
        reason: Unsupported,
    },
    /// Clearing the removed row's or column's stored cells failed.
    #[error("clearing the removed cells failed: {error:?}")]
    CellClear {
        /// The cell transaction's error.
        error: crate::table::cells::Error,
    },
    /// Commit was requested before staging an operation.
    #[error("no table-topology operation is staged")]
    NoStagedOperation,
    /// A surviving object depends on the edited graph.
    #[error("a surviving object depends on the table's graph")]
    UnsupportedDependency,
    /// The source cannot publish an exact changed artifact.
    #[error("source representation does not support table topology changes")]
    UnsupportedSource,
}

impl Refusal {
    /// The shared refusal kind litchi-ops classifies this reason by.
    #[must_use]
    pub const fn kind(&self) -> litchi_core::edit::RefusalKind {
        use litchi_core::edit::RefusalKind;
        match self {
            Self::LastAxis { .. } => RefusalKind::UnsupportedTarget,
            Self::SheetNotFound { .. } | Self::TableNotFound { .. } => RefusalKind::NotFound,
            Self::PositionOutOfRange { .. } => RefusalKind::InvalidValue,
            Self::SectionFull { .. }
            | Self::Unsupported { .. }
            | Self::CellClear { .. }
            | Self::UnsupportedDependency { .. }
            | Self::UnsupportedSource { .. } => RefusalKind::UnsupportedContent,
            Self::NoStagedOperation { .. } => RefusalKind::UnsupportedOperation,
        }
    }
}

impl From<Refusal> for Error {
    fn from(refusal: Refusal) -> Self {
        Self::Refused(refusal)
    }
}

const fn refuse(reason: Unsupported) -> Error {
    Error::Refused(Refusal::Unsupported { reason })
}

impl From<engine::Error> for Error {
    fn from(error: engine::Error) -> Self {
        use engine::{Error as E, Refusal as R};
        match error {
            E::Refused(R::PositionOutOfRange { index, length }) => {
                Self::Refused(Refusal::PositionOutOfRange { index, length })
            },
            E::Refused(R::SectionFull) => Self::Refused(Refusal::SectionFull),
            E::Refused(R::LastAxis) => Self::Refused(Refusal::LastAxis),
            E::Refused(R::Unsupported { reason }) => refuse(reason),
            E::Fault(fault) => Self::from(fault),
            E::Allocation { amount } => Self::Allocation { amount },
            _ => Self::InvalidSource,
        }
    }
}

/// One immutable edit that stages row and column insertions and removals
/// in one table, applied in order.
pub struct Edit<'a> {
    source: &'a Package,
    sheet: usize,
    table: usize,
    geometry: Geometry,
    operations: Vec<Operation>,
}

impl fmt::Debug for Edit<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Edit")
            .field("sheet", &self.sheet)
            .field("table", &self.table)
            .field("operations", &self.operations)
            .finish_non_exhaustive()
    }
}

impl Edit<'_> {
    fn stage(mut self, operation: Operation) -> Result<Self, Error> {
        let step = engine::resolve(operation, self.geometry)?;
        self.operations
            .try_reserve(1)
            .map_err(|_error| Error::Allocation { amount: 1 })?;
        self.operations.push(operation);
        self.geometry = step.after();
        Ok(self)
    }

    /// Stage a blank row inserted at a section-relative position. An index
    /// equal to the section's length appends to it; a header or footer
    /// insertion grows that section.
    ///
    /// # Errors
    ///
    /// [`Refusal::PositionOutOfRange`], [`Refusal::SectionFull`].
    pub fn insert_row(self, insertion: RowInsertion) -> Result<Self, Error> {
        self.stage(Operation::InsertRow(insertion))
    }

    /// Stage a blank column inserted at a section-relative position.
    ///
    /// # Errors
    ///
    /// [`Refusal::PositionOutOfRange`], [`Refusal::SectionFull`].
    pub fn insert_column(self, insertion: ColumnInsertion) -> Result<Self, Error> {
        self.stage(Operation::InsertColumn(insertion))
    }

    /// Stage the removal of the row at a section-relative position, with
    /// its cells.
    ///
    /// # Errors
    ///
    /// [`Refusal::PositionOutOfRange`], [`Refusal::LastAxis`].
    pub fn remove_row(self, deletion: RowDeletion) -> Result<Self, Error> {
        self.stage(Operation::RemoveRow(deletion))
    }

    /// Stage the removal of the column at a section-relative position,
    /// with its cells.
    ///
    /// # Errors
    ///
    /// [`Refusal::PositionOutOfRange`], [`Refusal::LastAxis`].
    pub fn remove_column(self, deletion: ColumnDeletion) -> Result<Self, Error> {
        self.stage(Operation::RemoveColumn(deletion))
    }

    /// Validate and atomically publish the staged operations, in order.
    ///
    /// # Errors
    ///
    /// [`Refusal::NoStagedOperation`] for an empty edit;
    /// [`Refusal::Unsupported`] for table content this transaction does not
    /// rewrite; [`Refusal::CellClear`] when a removed cell cannot be cleared;
    /// source, resource and verification errors. Failure never publishes a
    /// partial package.
    pub fn commit(self) -> Result<Commit, Error> {
        if self.operations.is_empty() {
            return Err(Error::Refused(Refusal::NoStagedOperation));
        }
        let mut current: Option<Package> = None;
        let mut touched = 0_usize;
        let mut previews = 0_usize;
        for operation in &self.operations {
            let base = current.as_ref().unwrap_or(self.source);
            let published = run(base, self.sheet, self.table, *operation)?;
            touched = touched.saturating_add(published.touched_components);
            previews = previews.saturating_add(published.deleted_previews);
            current = Some(published.package);
        }
        let package = current.ok_or(Error::Refused(Refusal::NoStagedOperation))?;
        Commit::new(
            self.source,
            Published {
                package,
                touched_components: touched,
                deleted_previews: previews,
            },
        )
    }
}

/// The table drawable and model of one table.
pub(super) fn locate(working: &Working<'_>, sheet: usize, table: usize) -> Result<Target, Error> {
    let document = working.object(1)?;
    let document_index = structure::document_message(document)?;
    let sheets = references::identifiers(
        &document.messages[document_index].data,
        &[DOCUMENT_SHEETS_FIELD],
    )?;
    let sheet = *sheets
        .get(sheet)
        .ok_or(Error::Refused(Refusal::SheetNotFound))?;
    let sheet_object = working.object(sheet)?;
    let (sheet_index, sheet_path) = structure::sheet_message(sheet_object)?;
    let mut tables = Vec::new();
    for drawable in references::identifiers(&sheet_object.messages[sheet_index].data, sheet_path)? {
        if structure::table_info_message(working.object(drawable)?).is_ok() {
            tables.push(drawable);
        }
    }
    let info = *tables
        .get(table)
        .ok_or(Error::Refused(Refusal::TableNotFound))?;
    let info_object = working.object(info)?;
    let info_index = structure::table_info_message(info_object)?;
    let model =
        single_reference(&info_object.messages[info_index].data, 2)?.ok_or(Error::InvalidSource)?;
    structure::unique_message(working.object(model)?, &[6_001])?;
    Ok(Target { info, model })
}

/// The unique message of `identifier` with one of `types`.
pub(super) fn message<'w>(
    working: &'w Working<'_>,
    identifier: u64,
    types: &[u32],
) -> FaultResult<(usize, &'w [u8])> {
    let object = working.object(identifier)?;
    let index = structure::unique_message(object, types)?;
    Ok((index, &object.messages[index].data))
}

/// A table model payload's geometry.
pub(super) fn geometry(model: &[u8]) -> Result<Geometry, Error> {
    Ok(engine::geometry(model)?)
}

/// The storage tiles of a data store: `(key, tile object)`, and the rows
/// per tile.
pub(super) fn tiles(store: &[u8]) -> Result<(Vec<(u32, u64)>, u64), Error> {
    Ok(engine::tiles(store)?)
}

fn read_geometry(source: &Package, sheet: usize, table: usize) -> Result<Geometry, Error> {
    let working = Working::new(source)?;
    let target = locate(&working, sheet, table)?;
    let (_, model) = message(&working, target.model, &[6_001])?;
    geometry(model)
}

/// Clear the comments on a removed row or column through the comment
/// transaction (record 0329); `None` when the axis holds none.
fn clear_comments(
    source: &Package,
    sheet: usize,
    table: usize,
    row_axis: bool,
    index: u32,
) -> Result<Option<Package>, Error> {
    let refused = |_error| refuse(Unsupported::Comments);
    let comments = source
        .table_cell_comments(SheetSelector::index(sheet), TableSelector::index(table))
        .map_err(refused)?;
    let mut current: Option<Package> = None;
    for (position, _) in comments {
        let on_axis = if row_axis {
            position.row() == index
        } else {
            position.column() == index
        };
        if !on_axis {
            continue;
        }
        let base = current.as_ref().unwrap_or(source);
        let commit = base
            .clear_table_cell_comment_for_axis_removal(
                SheetSelector::index(sheet),
                TableSelector::index(table),
                position,
            )
            .map_err(refused)?;
        current = Some(commit.into_package());
    }
    Ok(current)
}

/// Clear the stored cells of a removed row or column through the cell
/// transaction, so value, style and list references are released.
fn clear_axis(
    source: &Package,
    sheet: usize,
    table: usize,
    cells: &[At],
) -> Result<Package, Error> {
    let cell_error = |error| Error::Refused(Refusal::CellClear { error });
    let mut edit = source
        .edit_table_cells(SheetSelector::index(sheet), TableSelector::index(table))
        .map_err(cell_error)?;
    for position in cells {
        edit = edit.clear(*position).map_err(cell_error)?;
    }
    Ok(edit.commit().map_err(cell_error)?.into_package())
}

/// The table's stored cells, by position.
fn table_cells(
    package: &Package,
    sheet: usize,
    table: usize,
) -> Result<Vec<(At, crate::cell::Value)>, Error> {
    let table = package
        .workbook_ref()
        .sheets()
        .get(sheet)
        .and_then(|sheet| sheet.tables().nth(table))
        .ok_or(Error::Verification)?;
    let mut output = Vec::new();
    output
        .try_reserve(table.cell_count())
        .map_err(|_error| Error::Allocation {
            amount: table.cell_count(),
        })?;
    for cell in table.iter_cells() {
        output.push((cell.position(), cell.value().clone()));
    }
    Ok(output)
}

/// Run one operation: check, clear a removed axis, rewrite, publish and
/// verify. The structural rewrite is the shared engine
/// (`litchi_iwa_table::topology`, record 0412e); the comment and cell
/// clears are this package's own transactions.
fn run(
    source: &Package,
    sheet: usize,
    table: usize,
    operation: Operation,
) -> Result<Published, Error> {
    let before = read_geometry(source, sheet, table)?;
    let step = engine::resolve(operation, before)?;
    let target = locate(&Working::new(source)?, sheet, table)?;
    // Check everything before clearing anything.
    engine::check(source, &target, step, before)?;
    let (cells, inert) = match step.removed() {
        Some(_) => (
            engine::removed_cells(source, &target, step, before)?,
            engine::table_formula_inert(source, &target)?,
        ),
        None => (Vec::new(), true),
    };
    // Comments on the removed axis go through the comment transaction
    // first, which removes their storage graphs (record 0329).
    let uncommented = match step.removed() {
        Some(index) if !cells.is_empty() => {
            clear_comments(source, sheet, table, step.row_axis(), index)?
        },
        _ => None,
    };
    let commented = uncommented.as_ref().unwrap_or(source);
    // The cell writer clears the removed cells and refreshes the caches of
    // the formulas that read them; where it cannot, the structural pass
    // drops them itself (record 0329).
    // A table no formula touches drops the removed cells in the structural
    // pass, as the host's axis removal does (record 0370b): the previews
    // stay, and a dropped cell's string key is handed back.
    let (cleared, drop_cells) = if cells.is_empty() {
        (None, false)
    } else if inert {
        (None, true)
    } else {
        match clear_axis(commented, sheet, table, &cells) {
            Ok(package) => (Some(package), false),
            Err(
                error @ Error::Refused(Refusal::CellClear {
                    error: crate::table::cells::Error::Allocation { .. },
                }),
            ) => return Err(error),
            Err(Error::Refused(Refusal::CellClear { .. })) => (None, true),
            Err(other) => return Err(other),
        }
    };
    let base = cleared.as_ref().unwrap_or(commented);
    let target = locate(&Working::new(base)?, sheet, table)?;
    let mut options = engine::Publish::default();
    options.drop_cells = drop_cells;
    let published = engine::publish(base, &target, step, before, options)?;
    verify(source, &published.package, sheet, table, step)?;
    Ok(published)
}

/// The candidate has the new size and every surviving cell at its shifted
/// position with its value.
fn verify(
    source: &Package,
    candidate: &Package,
    sheet: usize,
    table: usize,
    step: Step,
) -> Result<(), Error> {
    let candidate_table = candidate
        .workbook_ref()
        .sheets()
        .get(sheet)
        .and_then(|sheet| sheet.tables().nth(table))
        .ok_or(Error::Verification)?;
    if (candidate_table.row_count(), candidate_table.column_count())
        != (step.after().rows, step.after().columns)
    {
        return Err(Error::Verification);
    }
    if read_geometry(candidate, sheet, table)? != step.after() {
        return Err(Error::Verification);
    }
    let mut expected = Vec::new();
    for (position, value) in table_cells(source, sheet, table)? {
        let (row, column) = if step.row_axis() {
            match step.shifted(position.row())? {
                Some(row) => (row, position.column()),
                None => continue,
            }
        } else {
            match step.shifted(position.column())? {
                Some(column) => (position.row(), column),
                None => continue,
            }
        };
        expected
            .try_reserve(1)
            .map_err(|_error| Error::Allocation { amount: 1 })?;
        expected.push((At::new(row, column), value));
    }
    let mut actual = table_cells(candidate, sheet, table)?;
    let key = |cell: &(At, crate::cell::Value)| (cell.0.row(), cell.0.column());
    expected.sort_by_key(key);
    actual.sort_by_key(key);
    // A formula's text follows the shifted coordinates: only its presence is
    // compared here.
    for cell in expected.iter_mut().chain(actual.iter_mut()) {
        if let crate::cell::Value::Formula(text) = &mut cell.1 {
            text.clear();
        }
    }
    if expected != actual {
        return Err(Error::Verification);
    }
    Ok(())
}

pub(super) fn table_position(
    source: &Package,
    sheet: usize,
    selector: TableSelector<'_>,
) -> Result<usize, Error> {
    let sheet = source
        .workbook_ref()
        .sheets()
        .get(sheet)
        .ok_or(Error::Refused(Refusal::SheetNotFound))?;
    match selector {
        TableSelector::Position(position) => {
            let position = position.get();
            if position >= sheet.table_count() {
                return Err(Error::Refused(Refusal::TableNotFound));
            }
            Ok(position)
        },
        TableSelector::Name(name) => {
            let mut matches = sheet
                .tables()
                .enumerate()
                .filter(|(_, table)| table.name() == name);
            let (position, _) = matches
                .next()
                .ok_or(Error::Refused(Refusal::TableNotFound))?;
            if matches.next().is_some() {
                return Err(Error::InvalidSource);
            }
            Ok(position)
        },
    }
}

impl Package {
    /// Start an immutable row and column insertion and removal edit on one
    /// table.
    ///
    /// ```no_run
    /// use litchi_numbers::Package;
    /// use litchi_numbers::table::topology::{RowDeletion, RowInsertion};
    ///
    /// # fn main() -> Result<(), Box<dyn std::error::Error>> {
    /// let package = Package::open("input.numbers")?;
    /// let commit = package
    ///     .edit_table_topology(0, 0)?
    ///     .insert_row(RowInsertion::body(0))?
    ///     .remove_row(RowDeletion::body(3))?
    ///     .commit()?;
    /// let restored = commit
    ///     .package()
    ///     .apply_table_topology(&commit.patch().inverse())?;
    /// # let _ = restored;
    /// # Ok(())
    /// # }
    /// ```
    ///
    /// # Errors
    ///
    /// [`Refusal::SheetNotFound`], [`Refusal::TableNotFound`]; source errors
    /// for a semantic-only or malformed source.
    pub fn edit_table_topology<'sheet, 'table>(
        &self,
        sheet: impl Into<SheetSelector<'sheet>>,
        table: impl Into<TableSelector<'table>>,
    ) -> Result<Edit<'_>, Error> {
        let sheet = self
            .workbook_ref()
            .sheet(sheet)
            .map_err(|_error| Error::InvalidSource)?
            .ok_or(Error::Refused(Refusal::SheetNotFound))?
            .index();
        let table = table_position(self, sheet, table.into())?;
        let geometry = read_geometry(self, sheet, table)?;
        Ok(Edit {
            source: self,
            sheet,
            table,
            geometry,
            operations: Vec::new(),
        })
    }

    /// Apply a table-topology patch to its exact source package.
    ///
    /// # Errors
    ///
    /// [`Error::PatchConflict`] when this package is not the patch's exact
    /// source.
    pub fn apply_table_topology(&self, patch: &Patch) -> Result<Commit, Error> {
        apply_patch(self, patch)
    }
}

#[cfg(test)]
mod tests {
    use litchi_numbers_wire::{BncCell, CachedScalar};

    use super::{TILE_MESSAGE_TYPE, locate, message, tiles};
    use crate::formula::Expression;
    use crate::package::structure::Working;
    use crate::table::cells::Input;
    use crate::table::topology::RowDeletion;
    use crate::{At, Package};

    /// The stored BNC cell at `position` of the first table.
    fn stored(package: &Package, position: At) -> BncCell {
        let working = Working::new(package).expect("working set");
        let target = locate(&working, 0, 0).expect("table");
        let (_, model) = message(&working, target.model, &[6_001]).expect("model");
        let store = super::single_bytes(model, 4)
            .expect("store")
            .expect("store");
        let (tiles, size) = tiles(store).expect("tiles");
        let key = u32::try_from(u64::from(position.row()) / size).expect("key");
        let local = u32::try_from(u64::from(position.row()) % size).expect("row");
        let tile = tiles
            .iter()
            .find(|(candidate, _)| *candidate == key)
            .expect("tile")
            .1;
        let (_, data) = message(&working, tile, &[TILE_MESSAGE_TYPE]).expect("tile payload");
        let limits = crate::cell::wire::limits_for_source(data);
        let bytes =
            litchi_numbers_wire::table_write::tile_cell(data, local, position.column(), &limits)
                .expect("cell")
                .expect("stored");
        BncCell::parse(&bytes, &crate::cell::wire::limits_for_source(&bytes)).expect("cell")
    }

    /// Record 0329: a formula whose reference is removed reads `#REF!` and
    /// caches an error result.
    #[test]
    fn deleted_references_cache_an_error() {
        let base = Package::new();
        let mut edit = base.edit_table_cells(0, 0).expect("edit");
        edit = edit
            .set_a1("B2", Input::number(4.0).expect("number"))
            .expect("value")
            .set_a1("B3", Input::number(5.0).expect("number"))
            .expect("value");
        edit = edit
            .set_formula_a1("E3", Expression::parse("B2+B3").expect("formula"))
            .expect("formula");
        let package = edit.commit().expect("cells").into_package();
        let before = stored(&package, At::new(2, 4));
        assert!(matches!(
            before.cached_scalar(),
            Ok(Some(CachedScalar::Number(_)))
        ));
        let removed = package
            .edit_table_topology(0, 0)
            .expect("topology")
            .remove_row(RowDeletion::body(0))
            .expect("row")
            .commit()
            .expect("commit");
        let after = stored(removed.package(), At::new(1, 4));
        assert!(matches!(
            after.cached_scalar(),
            Ok(Some(CachedScalar::Unsupported(8)))
        ));
        assert!(matches!(
            after.stored_value(),
            litchi_numbers_wire::StoredValue::Formula(_)
        ));
    }
}
