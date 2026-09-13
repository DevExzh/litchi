//! Private, bounded lookup metadata for worksheet formula references.
//!
//! The source worksheet model deliberately retains ODF repetition as physical
//! row and cell runs.  This index follows that representation: one row
//! descriptor is retained for every physical row run and one cumulative cell
//! end is retained for every physical cell run.  A lookup therefore performs
//! binary searches over physical runs without expanding a repeated grid or
//! cloning a sheet, cell, or cell string.

use crate::worksheet::{Cell, Sheet};
use litchi_core::{ExecutionContext, ExecutionError, Reservation, Resource};
use std::cmp::Ordering;
use std::{collections::TryReserveError, mem::size_of};

const ONE_WORK_UNIT: u64 = 1;
const NAME_COMPARE_CHUNK_BYTES: usize = 4 * 1024;

type BuildResult<T> = Result<T, super::Error>;

/// Borrowed physical-run lookup metadata used by the worksheet formula
/// adapter.
///
/// The index borrows the supplied sheet slice for its entire lifetime.  It
/// stores source-order sheet descriptors and a name-sorted view containing
/// only borrowed `str` slices and source positions.  The source slice remains
/// the authority for returned [`Cell`] references.
pub(super) struct Index<'source> {
    source: &'source [Sheet],
    names: Vec<NameEntry<'source>>,
    sheets: Vec<SheetEntry>,
    rows: Vec<RowEntry>,
    cell_ends: Vec<usize>,
    /// Keeps the exact metadata-capacity charge alive for this index.
    reservation: Reservation,
}

#[derive(Clone, Copy, Debug)]
struct NameEntry<'source> {
    name: &'source str,
    sheet: usize,
}

#[derive(Clone, Copy, Debug)]
struct SheetEntry {
    rows_start: usize,
    rows_len: usize,
}

#[derive(Clone, Copy, Debug)]
struct RowEntry {
    logical_end: usize,
    cell_ends_start: usize,
    cell_ends_len: usize,
}

impl<'source> Index<'source> {
    /// Build bounded metadata over borrowed worksheet runs.
    ///
    /// Construction performs a checked sizing pass before reserving any index
    /// vector.  It then scans the physical runs again to populate the exact
    /// capacities.  The second pass is intentional: all destination vectors
    /// are fallibly reserved and budgeted before their first growth.  Repeated
    /// rows and cells contribute one metadata entry each, never one entry per
    /// logical coordinate.
    ///
    /// # Errors
    ///
    /// Returns a duplicate-name or coordinate-overflow error when the source
    /// cannot be indexed. Allocation and execution-policy failures retain
    /// their typed diagnostics. A failed construction drops its in-progress
    /// reservation before returning.
    pub(super) fn new(source: &'source [Sheet], context: &ExecutionContext) -> BuildResult<Self> {
        context.check().map_err(super::Error::from)?;

        let counts = sizing_pass(source, context)?;
        let required_bytes = metadata_bytes(counts)?;
        let reservation = context
            .reserve(Resource::Memory, u64_len(required_bytes)?)
            .map_err(super::Error::from)?;

        let mut names = Vec::new();
        let mut sheets = Vec::new();
        let mut rows = Vec::new();
        let mut cell_ends = Vec::new();

        reserve_vec(
            context,
            &mut names,
            counts.sheets,
            "ODS formula sheet-name index",
        )?;
        reserve_vec(
            context,
            &mut sheets,
            counts.sheets,
            "ODS formula sheet index",
        )?;
        reserve_vec(context, &mut rows, counts.rows, "ODS formula row index")?;
        reserve_vec(
            context,
            &mut cell_ends,
            counts.cells,
            "ODS formula cell-run index",
        )?;

        populate(
            source,
            context,
            &mut names,
            &mut sheets,
            &mut rows,
            &mut cell_ends,
        )?;

        // Sorting only reorders borrowed metadata. A fallible heap sort keeps
        // cancellation and byte work checks inside every common-prefix scan;
        // no standard sort can hide an attacker-sized string comparison
        // between those checks.
        sort_names(&mut names, context)?;
        context.check().map_err(super::Error::from)?;

        for pair in names.windows(2) {
            if compare_name_text(context, pair[0].name, pair[1].name)? == Ordering::Equal {
                return Err(super::Error::DuplicateSheetName);
            }
        }

        Ok(Self {
            source,
            names,
            sheets,
            rows,
            cell_ends,
            reservation,
        })
    }

    /// Return the stable zero-based source position for an exact sheet name.
    ///
    /// Name comparisons are charged in bounded byte chunks, so a long common
    /// prefix cannot make one lookup uninterruptible. The supplied execution
    /// context controls cancellation and work charging for this lookup.
    pub(super) fn sheet_index(
        &self,
        name: &str,
        execution: &ExecutionContext,
    ) -> Result<Option<usize>, ExecutionError> {
        execution.check()?;
        let mut low = 0usize;
        let mut high = self.names.len();
        while low < high {
            let midpoint = low + (high - low) / 2;
            let entry = &self.names[midpoint];
            match compare_name_text(execution, entry.name, name)? {
                Ordering::Less => low = midpoint + 1,
                Ordering::Equal => return Ok(Some(entry.sheet)),
                Ordering::Greater => high = midpoint,
            }
        }
        Ok(None)
    }

    /// Look up a physical cell run by zero-based source sheet, row, and
    /// column coordinates.
    ///
    /// Repeated row and cell runs are resolved through binary searches over
    /// their physical descriptors.  A coordinate outside the represented
    /// logical extent, including an empty row or empty cell range, returns
    /// `None`. The supplied execution context controls cancellation and work
    /// charging for this lookup.
    pub(super) fn cell(
        &self,
        sheet: usize,
        row: usize,
        column: usize,
        execution: &ExecutionContext,
    ) -> Result<Option<&'source Cell>, ExecutionError> {
        execution.check()?;
        let Some(source_sheet) = self.source.get(sheet) else {
            return Ok(None);
        };
        let Some(sheet_entry) = self.sheets.get(sheet) else {
            return Ok(None);
        };
        let Some(rows_end) = sheet_entry.rows_start.checked_add(sheet_entry.rows_len) else {
            return Ok(None);
        };
        let Some(row_entries) = self.rows.get(sheet_entry.rows_start..rows_end) else {
            return Ok(None);
        };
        let Some(physical_row) = lower_bound_row(row_entries, row, execution)? else {
            return Ok(None);
        };
        let Some(row_entry) = row_entries.get(physical_row) else {
            return Ok(None);
        };

        let Some(cells_end) = row_entry
            .cell_ends_start
            .checked_add(row_entry.cell_ends_len)
        else {
            return Ok(None);
        };
        let Some(cell_entries) = self.cell_ends.get(row_entry.cell_ends_start..cells_end) else {
            return Ok(None);
        };
        let Some(physical_cell) = lower_bound_cell(cell_entries, column, execution)? else {
            return Ok(None);
        };
        Ok(source_sheet
            .rows
            .get(physical_row)
            .and_then(|row| row.cells.get(physical_cell)))
    }

    /// Return the number of source-order worksheets covered by this index.
    pub(super) fn sheet_count(
        &self,
        execution: &ExecutionContext,
    ) -> Result<usize, ExecutionError> {
        execution.check()?;
        execution.consume(Resource::Work, ONE_WORK_UNIT)?;
        Ok(self.source.len())
    }

    /// Return a borrowed source sheet name by stable zero-based position.
    pub(super) fn sheet_name_at(
        &self,
        sheet: usize,
        execution: &ExecutionContext,
    ) -> Result<Option<&'source str>, ExecutionError> {
        execution.check()?;
        execution.consume(Resource::Work, ONE_WORK_UNIT)?;
        Ok(self.source.get(sheet).map(|sheet| sheet.name.as_str()))
    }

    /// Return the metadata capacity charge retained by this index.
    #[must_use]
    pub(super) fn reserved_bytes(&self) -> u64 {
        self.reservation.amount()
    }
}

#[derive(Clone, Copy, Debug)]
struct Sizing {
    sheets: usize,
    rows: usize,
    cells: usize,
}

fn sizing_pass(source: &[Sheet], context: &ExecutionContext) -> BuildResult<Sizing> {
    let mut rows = 0usize;
    let mut cells = 0usize;

    for sheet in source {
        consume_work(context, ONE_WORK_UNIT)?;
        let mut logical_row_end = 0usize;
        for row in &sheet.rows {
            consume_work(context, ONE_WORK_UNIT)?;
            logical_row_end = logical_row_end
                .checked_add(row.repeat())
                .ok_or(super::Error::CoordinateOverflow)?;
            rows = rows
                .checked_add(1)
                .ok_or(super::Error::CoordinateOverflow)?;

            let mut logical_column_end = 0usize;
            for cell in &row.cells {
                consume_work(context, ONE_WORK_UNIT)?;
                logical_column_end = logical_column_end
                    .checked_add(cell.repeat())
                    .ok_or(super::Error::CoordinateOverflow)?;
                cells = cells
                    .checked_add(1)
                    .ok_or(super::Error::CoordinateOverflow)?;
            }
        }
    }

    Ok(Sizing {
        sheets: source.len(),
        rows,
        cells,
    })
}

fn populate<'source>(
    source: &'source [Sheet],
    context: &ExecutionContext,
    names: &mut Vec<NameEntry<'source>>,
    sheets: &mut Vec<SheetEntry>,
    rows: &mut Vec<RowEntry>,
    cell_ends: &mut Vec<usize>,
) -> BuildResult<()> {
    for (sheet_index, sheet) in source.iter().enumerate() {
        consume_work(context, ONE_WORK_UNIT)?;
        names.push(NameEntry {
            name: sheet.name.as_str(),
            sheet: sheet_index,
        });
        let rows_start = rows.len();
        let mut logical_row_end = 0usize;

        for row in &sheet.rows {
            consume_work(context, ONE_WORK_UNIT)?;
            logical_row_end = logical_row_end
                .checked_add(row.repeat())
                .ok_or(super::Error::CoordinateOverflow)?;

            let cell_ends_start = cell_ends.len();
            let mut logical_column_end = 0usize;
            for cell in &row.cells {
                consume_work(context, ONE_WORK_UNIT)?;
                logical_column_end = logical_column_end
                    .checked_add(cell.repeat())
                    .ok_or(super::Error::CoordinateOverflow)?;
                cell_ends.push(logical_column_end);
            }
            let cell_ends_len = cell_ends
                .len()
                .checked_sub(cell_ends_start)
                .ok_or(super::Error::CoordinateOverflow)?;
            rows.push(RowEntry {
                logical_end: logical_row_end,
                cell_ends_start,
                cell_ends_len,
            });
        }

        let rows_len = rows
            .len()
            .checked_sub(rows_start)
            .ok_or(super::Error::CoordinateOverflow)?;
        sheets.push(SheetEntry {
            rows_start,
            rows_len,
        });
    }
    Ok(())
}

fn metadata_bytes(counts: Sizing) -> BuildResult<usize> {
    let mut bytes = 0usize;
    bytes = checked_bytes(bytes, counts.sheets, size_of::<NameEntry<'static>>())?;
    bytes = checked_bytes(bytes, counts.sheets, size_of::<SheetEntry>())?;
    bytes = checked_bytes(bytes, counts.rows, size_of::<RowEntry>())?;
    checked_bytes(bytes, counts.cells, size_of::<usize>())
}

fn checked_bytes(current: usize, count: usize, element_size: usize) -> BuildResult<usize> {
    let bytes = count
        .checked_mul(element_size)
        .ok_or(super::Error::CoordinateOverflow)?;
    current
        .checked_add(bytes)
        .ok_or(super::Error::CoordinateOverflow)
}

fn reserve_vec<T>(
    context: &ExecutionContext,
    vector: &mut Vec<T>,
    capacity: usize,
    resource: &'static str,
) -> BuildResult<()> {
    context.check().map_err(super::Error::from)?;
    vector
        .try_reserve_exact(capacity)
        .map_err(|source| allocation(resource, source))
}

fn sort_names(names: &mut [NameEntry<'_>], context: &ExecutionContext) -> BuildResult<()> {
    context.check().map_err(super::Error::from)?;
    if names.len() < 2 {
        return Ok(());
    }

    // Heapify and drain the max heap in place.  This has deterministic
    // O(n log n) comparisons and does not allocate a staging vector.
    let len = names.len();
    let mut start = len / 2;
    while start > 0 {
        start -= 1;
        sift_down(names, start, len, context)?;
    }

    let mut end = len;
    while end > 1 {
        end -= 1;
        names.swap(0, end);
        sift_down(names, 0, end, context)?;
    }
    Ok(())
}

fn sift_down(
    names: &mut [NameEntry<'_>],
    mut root: usize,
    end: usize,
    context: &ExecutionContext,
) -> BuildResult<()> {
    loop {
        context.check().map_err(super::Error::from)?;
        let left = root
            .checked_mul(2)
            .and_then(|value| value.checked_add(1))
            .ok_or(super::Error::CoordinateOverflow)?;
        if left >= end {
            return Ok(());
        }
        let right = left
            .checked_add(1)
            .ok_or(super::Error::CoordinateOverflow)?;
        let mut candidate = left;
        if right < end
            && compare_name_entries(context, &names[candidate], &names[right])? == Ordering::Less
        {
            candidate = right;
        }
        if compare_name_entries(context, &names[root], &names[candidate])? != Ordering::Less {
            return Ok(());
        }
        names.swap(root, candidate);
        root = candidate;
    }
}

fn compare_name_entries(
    context: &ExecutionContext,
    left: &NameEntry<'_>,
    right: &NameEntry<'_>,
) -> BuildResult<Ordering> {
    let ordering = compare_name_text(context, left.name, right.name)?;
    if ordering == Ordering::Equal {
        Ok(left.sheet.cmp(&right.sheet))
    } else {
        Ok(ordering)
    }
}

fn compare_name_text(
    context: &ExecutionContext,
    left: &str,
    right: &str,
) -> Result<Ordering, ExecutionError> {
    let left = left.as_bytes();
    let right = right.as_bytes();
    let common = left.len().min(right.len());
    let mut offset = 0usize;
    while offset < common {
        let remaining = common - offset;
        let step = remaining.min(NAME_COMPARE_CHUNK_BYTES);
        let end = offset + step;
        context.consume(Resource::Work, step as u64)?;
        let ordering = left[offset..end].cmp(&right[offset..end]);
        if ordering != Ordering::Equal {
            return Ok(ordering);
        }
        offset = end;
    }
    // A comparison of two empty strings still reaches a checked cancellation
    // point, and the length comparison accounts for the common-prefix check.
    context.consume(Resource::Work, ONE_WORK_UNIT)?;
    Ok(left.len().cmp(&right.len()))
}

fn lower_bound_row(
    rows: &[RowEntry],
    row: usize,
    context: &ExecutionContext,
) -> Result<Option<usize>, ExecutionError> {
    let mut low = 0usize;
    let mut high = rows.len();
    while low < high {
        consume_lookup_work(context)?;
        let midpoint = low + (high - low) / 2;
        if rows[midpoint].logical_end <= row {
            low = midpoint + 1;
        } else {
            high = midpoint;
        }
    }
    if low < rows.len() {
        Ok(Some(low))
    } else {
        Ok(None)
    }
}

fn lower_bound_cell(
    ends: &[usize],
    column: usize,
    context: &ExecutionContext,
) -> Result<Option<usize>, ExecutionError> {
    let mut low = 0usize;
    let mut high = ends.len();
    while low < high {
        consume_lookup_work(context)?;
        let midpoint = low + (high - low) / 2;
        if ends[midpoint] <= column {
            low = midpoint + 1;
        } else {
            high = midpoint;
        }
    }
    if low < ends.len() {
        Ok(Some(low))
    } else {
        Ok(None)
    }
}

fn consume_lookup_work(context: &ExecutionContext) -> Result<(), ExecutionError> {
    context.consume(Resource::Work, ONE_WORK_UNIT)
}

fn consume_work(context: &ExecutionContext, amount: u64) -> BuildResult<()> {
    if amount == 0 {
        return Ok(());
    }
    context
        .consume(Resource::Work, amount)
        .map_err(super::Error::from)
}

fn u64_len(value: usize) -> BuildResult<u64> {
    u64::try_from(value).map_err(|_| super::Error::CoordinateOverflow)
}

fn allocation(resource: &'static str, source: TryReserveError) -> super::Error {
    super::Error::Allocation { resource, source }
}

#[cfg(test)]
mod tests {
    #![allow(
        clippy::expect_used,
        clippy::unwrap_used,
        reason = "index fixtures use panic-on-failure assertions"
    )]

    use super::*;
    use crate::worksheet::formula::Error;
    use crate::worksheet::{CellValue, Row};
    use litchi_core::{Budget, CancellationSource, ExecutionLimits, Limits};
    use std::{
        num::{NonZeroU64, NonZeroUsize},
        ptr,
    };

    fn context(memory: u64, work: u64) -> (Budget, CancellationSource, ExecutionContext) {
        let budget = Budget::root(
            "worksheet-formula-index-test",
            Limits::new(memory, u64::MAX, u64::MAX, u64::MAX, u64::MAX, work),
        );
        let (source, token) = CancellationSource::pair();
        let limits = ExecutionLimits::new(
            NonZeroUsize::MIN,
            NonZeroUsize::MIN,
            NonZeroU64::new(1).unwrap(),
            0,
        )
        .unwrap();
        (
            budget.clone(),
            source,
            ExecutionContext::new(budget, token, limits),
        )
    }

    fn repeated_sheets() -> Vec<Sheet> {
        let mut repeated = Sheet::new("Repeated").unwrap();
        let mut first = Row::repeated(2).unwrap();
        first
            .push_cell(Cell::repeated(CellValue::Text("left".to_string()), "left", 3).unwrap())
            .unwrap();
        first
            .push_cell(Cell::new(CellValue::Text("right".to_string()), "right"))
            .unwrap();
        repeated.rows.push(first);
        let mut tail = Row::new();
        tail.push_cell(Cell::new(CellValue::Text("tail".to_string()), "tail"))
            .unwrap();
        repeated.rows.push(tail);

        let mut direct = Sheet::new("Direct").unwrap();
        let mut row = Row::new();
        row.push_cell(Cell::new(CellValue::Text("tail".to_string()), "tail"))
            .unwrap();
        direct.rows.push(row);
        vec![repeated, direct]
    }

    #[test]
    fn repeated_runs_lookup_without_expansion_and_preserve_cell_identity() {
        let sheets = repeated_sheets();
        let (budget, _source, execution) = context(1024 * 1024, 1_000_000);
        let index = Index::new(&sheets, &execution).unwrap();

        assert_eq!(index.sheet_index("Repeated", &execution).unwrap(), Some(0));
        assert_eq!(index.sheet_index("Direct", &execution).unwrap(), Some(1));
        assert_eq!(index.sheet_index("missing", &execution).unwrap(), None);
        assert_eq!(index.sheet_count(&execution).unwrap(), 2);
        assert_eq!(
            index.sheet_name_at(0, &execution).unwrap(),
            Some("Repeated")
        );
        assert_eq!(index.sheet_name_at(1, &execution).unwrap(), Some("Direct"));
        assert_eq!(index.sheet_name_at(2, &execution).unwrap(), None);

        let first = &sheets[0].rows[0].cells[0];
        for row in [0, 1] {
            let found = index.cell(0, row, 0, &execution).unwrap().unwrap();
            assert!(ptr::eq(found, first));
            assert!(ptr::eq(
                index.cell(0, row, 2, &execution).unwrap().unwrap(),
                first
            ));
        }
        assert!(ptr::eq(
            index.cell(0, 0, 3, &execution).unwrap().unwrap(),
            &sheets[0].rows[0].cells[1]
        ));
        assert!(ptr::eq(
            index.cell(0, 2, 0, &execution).unwrap().unwrap(),
            &sheets[0].rows[1].cells[0]
        ));
        assert_eq!(budget.used(Resource::Memory), index.reserved_bytes());

        drop(index);
        assert_eq!(budget.used(Resource::Memory), 0);
    }

    #[test]
    fn missing_empty_and_high_coordinates_return_none() {
        let sheets = repeated_sheets();
        let (_budget, _source, execution) = context(1024 * 1024, 1_000_000);
        let index = Index::new(&sheets, &execution).unwrap();

        assert_eq!(index.cell(0, 3, 0, &execution).unwrap(), None);
        assert_eq!(index.cell(0, 0, 4, &execution).unwrap(), None);
        assert_eq!(
            index.cell(0, usize::MAX, usize::MAX, &execution).unwrap(),
            None
        );
        assert_eq!(index.cell(sheets.len(), 0, 0, &execution).unwrap(), None);
    }

    #[test]
    fn duplicate_names_are_rejected_and_memory_is_released() {
        let first = Sheet::new("Duplicate").unwrap();
        let second = Sheet::new("Duplicate").unwrap();
        let sheets = vec![first, second];
        let (budget, _source, execution) = context(1024 * 1024, 1_000_000);

        let result = Index::new(&sheets, &execution);
        assert!(matches!(result, Err(Error::DuplicateSheetName)));
        assert_eq!(budget.used(Resource::Memory), 0);
    }

    #[test]
    fn row_and_column_extents_reject_checked_overflow() {
        let mut row_overflow = Sheet::new("Rows").unwrap();
        row_overflow.rows = vec![Row::repeated(usize::MAX).unwrap(), Row::new()];
        let sheets = vec![row_overflow];
        let (_budget, _source, execution) = context(1024 * 1024, 1_000_000);
        assert!(matches!(
            Index::new(&sheets, &execution),
            Err(Error::CoordinateOverflow)
        ));

        let mut column_overflow = Sheet::new("Columns").unwrap();
        let mut row = Row::new();
        row.cells = vec![
            Cell::repeated(CellValue::Empty, "", usize::MAX).unwrap(),
            Cell::empty(),
        ];
        column_overflow.rows = vec![row];
        let sheets = vec![column_overflow];
        let (_budget, _source, execution) = context(1024 * 1024, 1_000_000);
        assert!(matches!(
            Index::new(&sheets, &execution),
            Err(Error::CoordinateOverflow)
        ));
    }

    #[test]
    fn memory_budget_and_cancellation_fail_without_reservation_leaks() {
        let sheets = repeated_sheets();
        let (budget, _source, execution) = context(0, 1_000_000);
        assert!(matches!(
            Index::new(&sheets, &execution),
            Err(Error::ResourceLimit(limit)) if limit.resource == Resource::Memory
        ));
        assert_eq!(budget.used(Resource::Memory), 0);

        let (budget, source, execution) = context(1024 * 1024, 1_000_000);
        source.cancel();
        assert!(matches!(
            Index::new(&sheets, &execution),
            Err(Error::Execution(ExecutionError::Cancelled))
        ));
        assert_eq!(budget.used(Resource::Memory), 0);
    }

    #[test]
    fn work_budget_is_charged_for_physical_scan_and_releases_memory_on_failure() {
        let sheets = repeated_sheets();
        let (budget, _source, execution) = context(1024 * 1024, 1);
        assert!(matches!(
            Index::new(&sheets, &execution),
            Err(Error::ResourceLimit(limit)) if limit.resource == Resource::Work
        ));
        assert!(budget.used(Resource::Work) <= 1);
        assert_eq!(budget.used(Resource::Memory), 0);
    }

    #[test]
    fn name_index_borrows_source_string_storage() {
        let sheets = repeated_sheets();
        let (_budget, _source, execution) = context(1024 * 1024, 1_000_000);
        let index = Index::new(&sheets, &execution).unwrap();
        let source_name = sheets[0].name.as_str();
        let indexed_name = index.names[1].name;
        assert!(ptr::eq(source_name.as_ptr(), indexed_name.as_ptr()));
        assert_eq!(source_name.len(), indexed_name.len());
    }

    #[test]
    fn caller_execution_policy_checks_each_lookup() {
        let sheets = repeated_sheets();
        let (_budget, source, execution) = context(1024 * 1024, 1_000_000);
        let index = Index::new(&sheets, &execution).unwrap();

        source.cancel();
        assert!(matches!(
            index.sheet_index("Repeated", &execution),
            Err(ExecutionError::Cancelled)
        ));
        assert!(matches!(
            index.cell(0, 0, 0, &execution),
            Err(ExecutionError::Cancelled)
        ));
        assert!(matches!(
            index.sheet_count(&execution),
            Err(ExecutionError::Cancelled)
        ));
        assert!(matches!(
            index.sheet_name_at(0, &execution),
            Err(ExecutionError::Cancelled)
        ));
    }

    #[test]
    fn long_common_prefix_work_is_bounded_and_releases_reservation() {
        let mut name = "p".repeat(NAME_COMPARE_CHUNK_BYTES * 2 + 1);
        name.push('A');
        let sheets = vec![Sheet::new(name.clone()).unwrap()];
        let (budget, source, execution) = context(1024 * 1024, 2);
        let index = Index::new(&sheets, &execution).unwrap();

        let result = index.sheet_index(&name, &execution);
        assert!(matches!(
            result,
            Err(ExecutionError::ResourceLimit(limit))
                if limit.resource == Resource::Work
        ));
        assert_eq!(budget.used(Resource::Memory), index.reserved_bytes());

        source.cancel();
        assert!(matches!(
            index.sheet_index(&name, &execution),
            Err(ExecutionError::Cancelled)
        ));

        drop(index);
        assert_eq!(budget.used(Resource::Memory), 0);
    }

    #[test]
    fn constructor_long_common_prefix_stops_before_unbounded_compare() {
        let prefix = "q".repeat(NAME_COMPARE_CHUNK_BYTES * 2 + 1);
        let first_name = format!("{prefix}A");
        let second_name = format!("{prefix}B");
        let sheets = vec![
            Sheet::new(first_name).unwrap(),
            Sheet::new(second_name).unwrap(),
        ];
        let (budget, _source, execution) = context(1024 * 1024, 8);

        let result = Index::new(&sheets, &execution);
        assert!(matches!(
            result,
            Err(Error::ResourceLimit(limit)) if limit.resource == Resource::Work
        ));
        assert_eq!(budget.used(Resource::Memory), 0);
    }
}
