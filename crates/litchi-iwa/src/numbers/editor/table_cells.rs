//! Transactional collection and one-pass application of table-cell batches.

use std::collections::HashSet;

use super::*;

pub(crate) struct TableCellBatch {
    updates: Vec<TableCellUpdate>,
    coordinates: Vec<(usize, usize)>,
}

impl TableCellBatch {
    pub(crate) fn collect(updates: impl IntoIterator<Item = TableCellUpdate>) -> Result<Self> {
        let updates = updates.into_iter().collect::<Vec<_>>();
        let mut seen = HashSet::with_capacity(updates.len());
        let mut coordinates = Vec::with_capacity(updates.len());
        for update in &updates {
            if !seen.insert((update.row, update.column)) {
                return Err(Error::ParseError(format!(
                    "Table cell batch repeats coordinate ({}, {})",
                    update.row, update.column
                )));
            }
            validate_value(&update.value)?;
            coordinates.push((update.row, update.column));
        }
        Ok(Self {
            updates,
            coordinates,
        })
    }

    pub(crate) fn is_empty(&self) -> bool {
        self.updates.is_empty()
    }

    pub(crate) fn len(&self) -> usize {
        self.updates.len()
    }

    pub(crate) fn apply_numbers(self, package: &mut IWorkPackage, table_id: u64) -> Result<usize> {
        let Self {
            updates,
            coordinates,
        } = self;
        let count = updates.len();
        model::set_cells_in_package(package, table_id, updates)?;
        formula_cache::refresh_formula_caches_after_cell_writes(package, table_id, &coordinates)?;
        Ok(count)
    }

    pub(crate) fn apply_attached(self, package: &mut IWorkPackage, table_id: u64) -> Result<usize> {
        let Self {
            updates,
            coordinates,
        } = self;
        let count = updates.len();
        model::set_attached_cells_in_package(package, table_id, updates)?;
        formula_cache::refresh_formula_caches_after_cell_writes(package, table_id, &coordinates)?;
        Ok(count)
    }
}

fn validate_value(value: &CellValue) -> Result<()> {
    match value {
        CellValue::Number(number) if !number.is_finite() => Err(Error::ParseError(
            "Table cell batches cannot store non-finite numbers".to_owned(),
        )),
        CellValue::Date(date) if !date.is_finite() => Err(Error::ParseError(
            "Table cell batches cannot store non-finite dates".to_owned(),
        )),
        CellValue::Duration(duration) if !duration.is_finite() => Err(Error::ParseError(
            "Table cell batches cannot store non-finite durations".to_owned(),
        )),
        CellValue::Formula(_) | CellValue::Error(_) => Err(Error::ParseError(
            "Formula and error cell writes require referenced-table construction".to_owned(),
        )),
        _ => Ok(()),
    }
}
