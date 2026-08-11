//! Archive application of validated Numbers table-cell batches.

use super::*;

use litchi_numbers::CellPosition;
use litchi_numbers::table::edit::{Batch, Budget, Error as BatchError};

const MAX_TABLE_CELL_BATCH: usize = 1 << 20;

pub(crate) struct TableCellBatch {
    batch: Batch,
}

impl TableCellBatch {
    pub(crate) fn collect(updates: impl IntoIterator<Item = TableCellUpdate>) -> Result<Self> {
        let batch =
            Batch::collect(updates, Budget::new(MAX_TABLE_CELL_BATCH)).map_err(map_batch_error)?;
        Ok(Self { batch })
    }

    pub(crate) fn is_empty(&self) -> bool {
        self.batch.is_empty()
    }

    pub(crate) fn len(&self) -> usize {
        self.batch.len()
    }

    pub(crate) fn apply_attached(self, package: &mut IWorkPackage, table_id: u64) -> Result<usize> {
        let (updates, coordinates) = self.batch.into_parts();
        let coordinates = into_formula_cache_coordinates(coordinates)?;
        let count = updates.len();
        model::set_attached_cells_in_package(package, table_id, updates.into_vec())?;
        formula_cache::refresh_formula_caches_after_cell_writes(package, table_id, &coordinates)?;
        Ok(count)
    }
}

/// Converts compact semantic coordinates to the legacy tuple shape consumed
/// by the private formula-cache implementation. Keeping this allocation at
/// the archive boundary prevents platform-sized coordinates from widening the
/// owned semantic batch.
fn into_formula_cache_coordinates(
    coordinates: Box<[CellPosition]>,
) -> Result<Box<[(usize, usize)]>> {
    coordinates
        .into_vec()
        .into_iter()
        .map(|position| {
            let row = usize::try_from(position.row()).map_err(|_| {
                Error::ParseError("Numbers row exceeds the platform index range".to_owned())
            })?;
            let column = usize::try_from(position.column()).map_err(|_| {
                Error::ParseError("Numbers column exceeds the platform index range".to_owned())
            })?;
            Ok((row, column))
        })
        .collect::<Result<Vec<_>>>()
        .map(Vec::into_boxed_slice)
}

fn map_batch_error(error: BatchError) -> Error {
    match error {
        BatchError::Allocation { resource, amount } => {
            Error::IwaCommon(litchi_iwa_common::Error::Allocation { resource, amount })
        },
        other => Error::ParseError(other.to_string()),
    }
}
