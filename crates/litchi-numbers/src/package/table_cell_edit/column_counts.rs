//! Column-header cell counts (record 0376c).
//!
//! Native files count the stored cells of each column in the column-header
//! bucket as well as each row's in the row-header buckets. The row counts
//! follow the tiles' final rows; the column counts follow each cell's
//! transition (a cell stored where none was adds one, a removed cell takes
//! one away), through the shared writer's `table_write::header_adjust_many`,
//! which creates a missing header, drops a bare header left counting none
//! and, as the host does, ignores a removal the bucket never counted.

use litchi_numbers_wire::{Limits, table_write};

use super::{
    budget, resolve,
    tile::{CellTransition, CellValue},
    usize_u64,
};
use crate::package::table_cells::{DependencyKind, Error, LimitKind, Path, Refusal};

/// Per-column cell-count deltas, sorted by column, in storage reserved up
/// front.
#[derive(Debug, Default)]
pub(super) struct Deltas(Vec<(u32, i64)>);

impl Deltas {
    /// Room for the deltas of `columns` columns.
    pub(super) fn with_capacity(columns: usize) -> Result<Self, Error> {
        let mut deltas = Vec::new();
        deltas
            .try_reserve_exact(columns)
            .map_err(|_error| Error::Allocation {
                kind: LimitKind::RetainedElements,
                amount: columns,
            })?;
        Ok(Self(deltas))
    }

    /// Record one cell transition.
    pub(super) fn record(&mut self, transition: &CellTransition, path: Path) -> Result<(), Error> {
        let delta = i64::from(transition.after != CellValue::Missing)
            - i64::from(transition.before != CellValue::Missing);
        if delta == 0 {
            return Ok(());
        }
        match self
            .0
            .binary_search_by_key(&transition.column, |(column, _)| *column)
        {
            Ok(index) => {
                let entry = &mut self.0[index].1;
                *entry = entry
                    .checked_add(delta)
                    .ok_or(Error::InvalidSource { path })?;
            },
            Err(_) if self.0.len() == self.0.capacity() => {
                return Err(Error::Verification { path });
            },
            Err(index) => self.0.insert(index, (transition.column, delta)),
        }
        Ok(())
    }

    /// Whether a column count changes.
    pub(super) fn changes(&self) -> bool {
        self.0.iter().any(|(_, delta)| *delta != 0)
    }

    fn output_bound(&self, bucket: usize, path: Path) -> Result<usize, Error> {
        self.0
            .len()
            .checked_mul(table_write::HEADER_BYTES)
            .and_then(|growth| growth.checked_add(bucket))
            .ok_or(Error::InvalidSource { path })
    }

    /// The envelope of rewriting a `bucket`-byte column-header bucket: the
    /// bucket, each of its headers and each changed header are scanned once,
    /// and the output is one allocation.
    pub(super) fn envelope(&self, bucket: usize, path: Path) -> Result<budget::Usage, Error> {
        let output = self.output_bound(bucket, path)?;
        let scanned = bucket.checked_mul(3).ok_or(Error::InvalidSource { path })?;
        let work = scanned
            .checked_add(output)
            .ok_or(Error::InvalidSource { path })?;
        Ok(budget::Usage {
            wire_bytes: usize_u64(
                bucket
                    .checked_add(output)
                    .ok_or(Error::InvalidSource { path })?,
            ),
            wire_fields: usize_u64(scanned),
            wire_work: usize_u64(work),
            header_reads: usize_u64(bucket / 2),
            header_writes: usize_u64(self.0.len()),
            peak_scratch_bytes: usize_u64(output),
            allocation_events: 1,
            transaction_work: usize_u64(work),
            ..budget::Usage::default()
        })
    }

    /// The column-header bucket `bucket` of `storage` with the deltas
    /// applied, or `None` when no column count changes.
    pub(super) fn apply(
        &self,
        bucket: &[u8],
        storage: &resolve::StorageRoutes,
        path: Path,
    ) -> Result<Option<Vec<u8>>, Error> {
        if !self.changes() {
            return Ok(None);
        }
        let refused = Error::Refused(Refusal::UnsupportedDependency {
            path,
            kind: DependencyKind::CellStorage,
        });
        if storage.row_headers.contains(&storage.column_headers) {
            return Err(refused);
        }
        let bound = self.output_bound(bucket.len(), path)?;
        let limits = Limits::DEFAULT
            .with_input_bytes(bucket.len().max(1))
            .and_then(|limits| limits.with_output_bytes(bound.max(1)))
            .map_err(|_error| Error::InvalidSource { path })?;
        let output = table_write::header_adjust_many(bucket, &self.0, &limits).map_err(
            |error| match error {
                table_write::WriteError::Unsupported => refused,
                table_write::WriteError::Allocation { requested } => Error::Allocation {
                    kind: LimitKind::RetainedBytes,
                    amount: requested,
                },
                _ => Error::InvalidSource { path },
            },
        )?;
        Ok((output != bucket).then_some(output))
    }
}
