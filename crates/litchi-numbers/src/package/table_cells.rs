//! Selector-first semantic Numbers table-cell reads.

use std::fmt;

use crate::{
    SheetSelector, TableSelector,
    table::{
        CellPosition, CellRange, Dimensions, Table,
        cells::{State, Storage},
    },
};

use super::Package;

/// A content-free semantic location used by cell-read diagnostics.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum Path {
    /// The complete Numbers package.
    Package,
    /// One table at checked semantic positions.
    Table { sheet: u32, table: u32 },
}

/// A finite resource governed by a dense cell read.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum LimitKind {
    /// Cell changes staged by one atomic batch.
    Updates,
    /// Elements retained by a dense result.
    RetainedElements,
    /// UTF-8 bytes owned by stored text, formula, and error results.
    OwnedValueBytes,
    /// Strict protobuf fields inspected by the transaction.
    WireFields,
    /// Strict protobuf traversal work charged by the transaction.
    WireWork,
    /// Native objects inspected or retained by the transaction.
    Objects,
    /// Native object references inspected or retained by the transaction.
    References,
    /// Formula nodes and dependency work inspected by the transaction.
    FormulaWork,
    /// Bytes retained by the transaction plan or exact artifacts.
    RetainedBytes,
    /// Peak temporary bytes required before publication.
    PeakScratchBytes,
    /// Bytes in the candidate package artifact.
    OutputBytes,
    /// Work required to reopen and verify the candidate.
    ReopenWork,
    /// Aggregate deterministic transaction work.
    TransactionWork,
}

/// A modeled native dependency required by a changed cell batch.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum DependencyKind {
    /// Sparse tile, row, or header storage.
    CellStorage,
    /// Plain string-list ownership and refcounts.
    SharedString,
    /// Rich-text payload, storage, style, or copy-on-write ownership.
    RichText,
    /// Formula or formula-error list ownership.
    Formula,
    /// Formula dependency and cached-result ownership.
    FormulaCache,
    /// Cell comment metadata preserved by a clear.
    Comment,
    /// Calculation-engine header-name indexes.
    HeaderNameIndex,
    /// Merged-cell ownership.
    Merge,
    /// Pivot-table ownership or derived state.
    Pivot,
    /// Category-table ownership or derived state.
    Category,
    /// Spill or array-formula ownership.
    Spill,
    /// Hidden-row or hidden-column formula ownership.
    HiddenState,
    /// Conditional-style formula ownership.
    ConditionalStyle,
}

/// Failure from a selector-first semantic table-cell operation.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum Error {
    InvalidAddress,
    SheetNotFound,
    TableNotFound,
    AmbiguousSource {
        path: Path,
    },
    OutOfBounds {
        position: CellPosition,
        dimensions: Dimensions,
    },
    LimitExceeded {
        kind: LimitKind,
        observed: u64,
        maximum: u64,
        path: Path,
    },
    Allocation {
        kind: LimitKind,
        amount: usize,
    },
    DuplicatePosition {
        position: CellPosition,
    },
    InvalidSource {
        path: Path,
    },
    TableLocked {
        path: Path,
    },
    UnsupportedSource {
        path: Path,
    },
    UnsupportedDependency {
        path: Path,
        kind: DependencyKind,
    },
    Verification {
        path: Path,
    },
    PatchConflict,
}

impl fmt::Display for Error {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::InvalidAddress => formatter.write_str("invalid Numbers cell address"),
            Self::SheetNotFound => formatter.write_str("the Numbers sheet selector did not match"),
            Self::TableNotFound => formatter.write_str("the Numbers table selector did not match"),
            Self::AmbiguousSource { .. } => {
                formatter.write_str("the Numbers table-cell selector is ambiguous")
            },
            Self::OutOfBounds { .. } => {
                formatter.write_str("the Numbers cell coordinate is outside the selected table")
            },
            Self::LimitExceeded {
                observed, maximum, ..
            } => write!(
                formatter,
                "Numbers table-cell operation limit exceeded: observed {observed}, maximum {maximum}"
            ),
            Self::Allocation { kind, amount } => write!(
                formatter,
                "could not allocate {amount} units of {kind:?} for the Numbers table-cell operation"
            ),
            Self::DuplicatePosition { .. } => {
                formatter.write_str("the Numbers cell batch contains a duplicate coordinate")
            },
            Self::InvalidSource { .. } => {
                formatter.write_str("the Numbers table-cell source is malformed")
            },
            Self::TableLocked { .. } => formatter.write_str("the selected Numbers table is locked"),
            Self::UnsupportedSource { .. } => {
                formatter.write_str("the Numbers table-cell source is not safely writable")
            },
            Self::UnsupportedDependency { .. } => {
                formatter.write_str("the Numbers table-cell change has an unsupported dependency")
            },
            Self::Verification { .. } => {
                formatter.write_str("the Numbers table-cell operation failed verification")
            },
            Self::PatchConflict => {
                formatter.write_str("the Numbers table-cell patch source conflicts")
            },
        }
    }
}

impl std::error::Error for Error {}

impl Package {
    /// Read one presence-preserving cell from a semantically selected table.
    ///
    /// # Errors
    ///
    /// Returns a typed error for a missing or ambiguous selector, a coordinate
    /// outside the selected table, or a malformed semantic source position.
    pub fn table_cell<'sheet, 'table>(
        &self,
        sheet: impl Into<SheetSelector<'sheet>>,
        table: impl Into<TableSelector<'table>>,
        position: CellPosition,
    ) -> Result<State, Error> {
        let selected = resolve_table(self, sheet.into(), table.into())?;
        ensure_position(selected.table.dimensions(), position, selected.path)?;
        state_at(selected.table, position, selected.path)
    }

    /// Read a bounded dense row-major range from a selected table.
    ///
    /// Missing cells remain [`Storage::Missing`], while an explicitly stored
    /// [`crate::cell::Value::Empty`] remains [`Storage::Stored`].
    ///
    /// # Errors
    ///
    /// Returns a typed selector, bounds, resource, or allocation error before
    /// publishing a partial result.
    pub fn table_cells<'sheet, 'table>(
        &self,
        sheet: impl Into<SheetSelector<'sheet>>,
        table: impl Into<TableSelector<'table>>,
        range: CellRange,
    ) -> Result<Vec<State>, Error> {
        let selected = resolve_table(self, sheet.into(), table.into())?;
        ensure_range(selected.table.dimensions(), range, selected.path)?;
        let requested = range.area().ok_or(Error::LimitExceeded {
            kind: LimitKind::RetainedElements,
            observed: u64::MAX,
            maximum: usize_to_u64(self.state.options.semantic().max_materialized_cells()),
            path: selected.path,
        })?;
        let maximum = self.state.options.semantic().max_materialized_cells();
        if requested > maximum {
            return Err(Error::LimitExceeded {
                kind: LimitKind::RetainedElements,
                observed: usize_to_u64(requested),
                maximum: usize_to_u64(maximum),
                path: selected.path,
            });
        }
        if requested == 0 {
            return Ok(Vec::new());
        }

        let text_maximum = self.state.options.semantic().max_output_text_bytes();
        let selected_text_bytes =
            selected_text_bytes(selected.table, range).ok_or(Error::LimitExceeded {
                kind: LimitKind::OwnedValueBytes,
                observed: u64::MAX,
                maximum: usize_to_u64(text_maximum),
                path: selected.path,
            })?;
        if selected_text_bytes > text_maximum {
            return Err(Error::LimitExceeded {
                kind: LimitKind::OwnedValueBytes,
                observed: usize_to_u64(selected_text_bytes),
                maximum: usize_to_u64(text_maximum),
                path: selected.path,
            });
        }

        let mut states = Vec::new();
        states
            .try_reserve_exact(requested)
            .map_err(|_error| Error::Allocation {
                kind: LimitKind::RetainedElements,
                amount: requested,
            })?;
        let (start, end) = range.bounds();
        let mut stored = selected
            .table
            .cells(range)
            .map_err(|_error| Error::Verification {
                path: selected.path,
            })?
            .peekable();
        for row in start.row()..end.row() {
            for column in start.column()..end.column() {
                let position = CellPosition::new(row, column);
                while stored.peek().is_some_and(|cell| cell.position() < position) {
                    stored.next();
                }
                let storage = match stored.peek() {
                    Some(cell) if cell.position() == position => {
                        Storage::Stored(try_clone_value(cell.value(), selected.path)?)
                    },
                    Some(_) | None => Storage::Missing,
                };
                states.push(State::new(position, storage));
            }
        }
        if states.len() != requested {
            return Err(Error::Verification {
                path: selected.path,
            });
        }
        Ok(states)
    }
}

#[derive(Clone, Copy)]
pub(super) struct SelectedTable<'package> {
    pub(super) table: &'package Table,
    pub(super) path: Path,
}

pub(super) fn resolve_table<'package>(
    source: &'package Package,
    sheet_selector: SheetSelector<'_>,
    table_selector: TableSelector<'_>,
) -> Result<SelectedTable<'package>, Error> {
    let (sheet_position, sheet) = match sheet_selector {
        SheetSelector::Index(position) => source
            .state
            .document
            .sheets()
            .get(position)
            .map(|sheet| (position, sheet))
            .ok_or(Error::SheetNotFound)?,
        SheetSelector::Name(name) => {
            let mut matches = source
                .state
                .document
                .sheets()
                .iter()
                .enumerate()
                .filter(|(_position, sheet)| sheet.name() == name);
            let first = matches.next().ok_or(Error::SheetNotFound)?;
            if matches.next().is_some() {
                return Err(Error::AmbiguousSource {
                    path: Path::Package,
                });
            }
            first
        },
    };
    let compact_sheet = compact_position(sheet_position, Path::Package)?;

    let (table_position, table) = match table_selector {
        TableSelector::Index(position) => sheet
            .tables()
            .nth(position)
            .map(|table| (position, table))
            .ok_or(Error::TableNotFound)?,
        TableSelector::Name(name) => {
            let mut matches = sheet
                .tables()
                .enumerate()
                .filter(|(_position, table)| table.name() == name);
            let first = matches.next().ok_or(Error::TableNotFound)?;
            if matches.next().is_some() {
                return Err(Error::AmbiguousSource {
                    path: Path::Table {
                        sheet: compact_sheet,
                        table: compact_position(first.0, Path::Package)?,
                    },
                });
            }
            first
        },
    };
    let compact_table = compact_position(table_position, Path::Package)?;
    Ok(SelectedTable {
        table,
        path: Path::Table {
            sheet: compact_sheet,
            table: compact_table,
        },
    })
}

fn state_at(table: &Table, position: CellPosition, path: Path) -> Result<State, Error> {
    let storage = match table.get(position) {
        Some(value) => Storage::Stored(try_clone_value(value, path)?),
        None => Storage::Missing,
    };
    Ok(State::new(position, storage))
}

fn selected_text_bytes(table: &Table, range: CellRange) -> Option<usize> {
    table.cells(range).ok()?.try_fold(0_usize, |total, cell| {
        total.checked_add(owned_text_bytes(cell.value()))
    })
}

const fn owned_text_bytes(value: &crate::cell::Value) -> usize {
    match value {
        crate::cell::Value::Text(text)
        | crate::cell::Value::Formula(text)
        | crate::cell::Value::Error(text) => text.len(),
        crate::cell::Value::Empty
        | crate::cell::Value::Number(_)
        | crate::cell::Value::Boolean(_)
        | crate::cell::Value::Date(_)
        | crate::cell::Value::Duration(_) => 0,
    }
}

fn try_clone_value(value: &crate::cell::Value, _path: Path) -> Result<crate::cell::Value, Error> {
    use crate::cell::Value;

    Ok(match value {
        Value::Empty => Value::Empty,
        Value::Text(text) => Value::Text(try_clone_text(text)?),
        Value::Number(number) => Value::Number(*number),
        Value::Boolean(boolean) => Value::Boolean(*boolean),
        Value::Date(date) => Value::Date(*date),
        Value::Duration(duration) => Value::Duration(*duration),
        Value::Formula(formula) => Value::Formula(try_clone_text(formula)?),
        Value::Error(error) => Value::Error(try_clone_text(error)?),
    })
}

fn try_clone_text(source: &str) -> Result<String, Error> {
    let mut target = String::new();
    target
        .try_reserve_exact(source.len())
        .map_err(|_error| Error::Allocation {
            kind: LimitKind::OwnedValueBytes,
            amount: source.len(),
        })?;
    target.push_str(source);
    Ok(target)
}

fn ensure_position(
    dimensions: Dimensions,
    position: CellPosition,
    _path: Path,
) -> Result<(), Error> {
    if position.row() >= dimensions.rows() || position.column() >= dimensions.columns() {
        return Err(Error::OutOfBounds {
            position,
            dimensions,
        });
    }
    Ok(())
}

fn ensure_range(dimensions: Dimensions, range: CellRange, _path: Path) -> Result<(), Error> {
    let (_start, end) = range.bounds();
    if end.row() > dimensions.rows() || end.column() > dimensions.columns() {
        return Err(Error::OutOfBounds {
            position: end,
            dimensions,
        });
    }
    Ok(())
}

fn compact_position(position: usize, path: Path) -> Result<u32, Error> {
    u32::try_from(position).map_err(|_error| Error::LimitExceeded {
        kind: LimitKind::RetainedElements,
        observed: usize_to_u64(position),
        maximum: u64::from(u32::MAX),
        path,
    })
}

fn usize_to_u64(value: usize) -> u64 {
    u64::try_from(value).unwrap_or(u64::MAX)
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::cell::Value;

    #[test]
    fn package_projection_preserves_stored_empty() {
        let position = CellPosition::new(0, 0);
        let mut builder = Table::builder("Presence", Dimensions::new(1, 2));
        builder.set(position, Value::Empty).expect("in bounds");
        let table = builder.finish().expect("unique coordinate");

        let stored = state_at(&table, position, Path::Package).expect("stored value");
        let missing =
            state_at(&table, CellPosition::new(0, 1), Path::Package).expect("missing value");

        assert!(matches!(stored.storage(), Storage::Stored(Value::Empty)));
        assert!(matches!(missing.storage(), Storage::Missing));
    }
}
