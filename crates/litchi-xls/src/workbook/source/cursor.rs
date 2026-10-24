//! A streaming cursor over the stored cells of a source-backed worksheet
//! (record 0267).
//!
//! [`SourceBackedWorksheet::visit_cells`] scans the whole worksheet substream
//! and reports its records in stream order; a reader that wants the sorted
//! grid (one cell per coordinate, row-major) collects and sorts them, so a
//! caller that wants only the first few cells (a query's `first()` or
//! `limit(n)`) pays for the whole sheet.
//!
//! [`SourceBackedWorksheet::cell_cursor`] first runs one validating scan that
//! decodes no value: every record takes the checks the visitor's scan takes
//! (the measure path of change 0636, which runs every check the parse runs),
//! and the scan notes whether the stored positions strictly ascend in
//! row-major order. A refused sheet is refused there, before any cell. A
//! sheet whose records do not ascend (a position stored out of order or
//! twice) gets no cursor: its sorted grid is not its stream order. Otherwise
//! the cursor runs the visitor's scan in batches: each batch scans from where
//! the last one stopped until it holds [`BATCH`] cells and stops at the next
//! record boundary that leaves no string-valued `FORMULA` waiting for its
//! `STRING`. A resumed scan keeps counting the worksheet's scanned bytes and
//! records from the sheet start, so the limits apply as to one scan, and
//! each cell is the one the visitor reports, so the cursor returns exactly
//! the sorted grid's cells in order.
//!
//! What the validating scan does not decode is a cell's value: a shared
//! string index the string table cannot resolve, or a `STRING` result
//! record, is refused when the batch holding it is scanned. A batch whose
//! scan is refused returns none of its cells, only the refusal.
//!
//! The cursor retains at most one batch of cells and the scan position;
//! every scan runs the source and cancellation fences of a scan.

use std::collections::VecDeque;
use std::sync::Arc;

use super::{
    CellOccurrence, CellRecord, CellSink, Result, ScanContext, ScanResume, SharedStringResolver,
    SourceBackedCell, SourceBackedWorksheet, SourceInner, VisitCells, cell_xf_index,
    scan_worksheet, scan_worksheet_from,
};
use crate::records::MeasuredCell;

/// Cells one batch scans for.
const BATCH: usize = 64;

/// A cursor over a worksheet's stored cells, in row-major order.
///
/// Returned by [`SourceBackedWorksheet::cell_cursor`] for a worksheet whose
/// stored positions strictly ascend. Each item is the cell
/// [`SourceBackedWorksheet::visit_cells`] reports at that point of its scan;
/// see the module documentation for the refusals a batch can still report.
///
/// ```no_run
/// use litchi_xls::SourceBackedWorkbook;
/// let workbook = SourceBackedWorkbook::open("book.xls")?;
/// let sheet = workbook.worksheet_by_index(0)?.expect("worksheet");
/// if let Some(mut cursor) = sheet.cell_cursor()? {
///     let first = cursor.next().transpose()?;
///     # let _ = first;
/// }
/// # Ok::<(), litchi_xls::SourceBackedError>(())
/// ```
pub struct CellCursor {
    owner: Arc<SourceInner>,
    /// The projection owner, retained separately from the effective scan
    /// owner so a deferred publication or package handoff keeps checking its
    /// dynamic source while batches are read from the materialized target.
    source_owner: Arc<SourceInner>,
    sheet_index: usize,
    /// Whether a formula of the worksheet renders no text without the
    /// workbook's name and external-sheet tables.
    needs_formula_context: bool,
    /// Where the next batch starts; `None` before the first batch.
    resume: Option<ScanResume>,
    batch: VecDeque<SourceBackedCell>,
    done: bool,
}

impl SourceBackedWorksheet {
    /// A cursor over this worksheet's stored cells, in row-major order,
    /// which decodes only as many cells as the caller reads (a batch at a
    /// time), after one scan that validates the sheet without decoding
    /// values.
    ///
    /// `None` when the stored positions do not strictly ascend in row-major
    /// order (a position stored out of order or twice): the sorted grid is
    /// then not the stream order, and the caller sorts the cells
    /// [`Self::visit_cells`] reports.
    ///
    /// # Errors
    ///
    /// The framing, limit, parse and freshness errors of
    /// [`Self::visit_cells`], from the validating scan.
    pub fn cell_cursor(&self) -> Result<Option<CellCursor>> {
        let owner = self.effective_owner()?;
        let kept = owner
            .cursor_checks
            .get(self.sheet_index)
            .and_then(std::sync::OnceLock::get)
            .copied();
        let check = match kept {
            // The snapshot an earlier cursor validated: only the freshness
            // fence remains.
            Some(check) => {
                self.owner.ensure_current()?;
                check
            },
            None => {
                let mut order = Order {
                    last: None,
                    ascending: true,
                    needs_formula_context: false,
                };
                let scan = scan_worksheet(&owner, self.sheet_index, None, &mut order);
                if self.owner.needs_source_fence() {
                    self.owner.ensure_current()?;
                }
                scan?;
                let check = Check {
                    ascending: order.ascending,
                    needs_formula_context: order.needs_formula_context,
                };
                if let Some(slot) = owner.cursor_checks.get(self.sheet_index) {
                    let _kept = slot.set(check);
                }
                check
            },
        };
        Ok(check.ascending.then(|| CellCursor {
            owner,
            source_owner: Arc::clone(&self.owner),
            sheet_index: self.sheet_index,
            needs_formula_context: check.needs_formula_context,
            resume: None,
            batch: VecDeque::new(),
            done: false,
        }))
    }
}

impl CellCursor {
    /// Whether a formula of this worksheet renders no text without the
    /// workbook's name and external-sheet tables: its cell's
    /// [`SourceBackedCell::formula`] is `None` although it is a formula, and
    /// a reader that wants the text reads it from the eager worksheet.
    #[must_use]
    pub const fn needs_formula_context(&self) -> bool {
        self.needs_formula_context
    }

    /// Scan the next batch.
    fn refill(&mut self) -> Result<()> {
        if self.source_owner.needs_source_fence() {
            self.source_owner.ensure_current()?;
        }
        let count = std::cell::Cell::new(0_usize);
        let batch = &mut self.batch;
        let mut sink = Batch {
            visit: VisitCells {
                visitor: |cell: SourceBackedCell| {
                    batch.push_back(cell);
                    count.set(count.get().saturating_add(1));
                    Ok(())
                },
            },
            count: &count,
        };
        let scan = scan_worksheet_from(&self.owner, self.sheet_index, None, &mut sink, self.resume);
        if self.source_owner.needs_source_fence() {
            self.source_owner.ensure_current()?;
        }
        let resume = scan?;
        match resume {
            Some(resume) => self.resume = Some(resume),
            None => self.done = true,
        }
        Ok(())
    }
}

impl Iterator for CellCursor {
    type Item = Result<SourceBackedCell>;

    fn next(&mut self) -> Option<Self::Item> {
        loop {
            if let Some(cell) = self.batch.pop_front() {
                return Some(Ok(cell));
            }
            if self.done {
                return None;
            }
            if let Err(error) = self.refill() {
                self.done = true;
                self.batch.clear();
                return Some(Err(error));
            }
        }
    }
}

impl std::iter::FusedIterator for CellCursor {}

impl std::fmt::Debug for CellCursor {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("CellCursor")
            .field("sheet_index", &self.sheet_index)
            .field("buffered", &self.batch.len())
            .field("done", &self.done)
            .finish_non_exhaustive()
    }
}

/// The whole-sheet visitor's sink, stopping once it holds a batch.
struct Batch<'c, F> {
    visit: VisitCells<F>,
    /// Cells this batch holds.
    count: &'c std::cell::Cell<usize>,
}

impl<F> CellSink for Batch<'_, F>
where
    F: FnMut(SourceBackedCell) -> Result<()>,
{
    fn accept(
        &mut self,
        record: &CellRecord,
        occurrence: CellOccurrence,
        scan: &ScanContext<'_>,
        strings: &mut SharedStringResolver<'_>,
    ) -> Result<()> {
        self.visit.accept(record, occurrence, scan, strings)
    }

    fn defers_string_formula(&self, record: &CellRecord) -> bool {
        self.visit.defers_string_formula(record)
    }

    fn wants(&self, row: u16, col: u16) -> bool {
        self.visit.wants(row, col)
    }

    fn stop(&self) -> bool {
        self.count.get() >= BATCH
    }
}

/// What the validating scan found about one worksheet.
#[derive(Debug, Clone, Copy)]
pub(super) struct Check {
    ascending: bool,
    needs_formula_context: bool,
}

/// The validating scan's sink: it decodes no value (every record but a
/// formula takes the measure path), and notes whether the positions strictly
/// ascend and whether a formula needs the workbook's name tables to render.
struct Order {
    last: Option<(u16, u16)>,
    ascending: bool,
    needs_formula_context: bool,
}

impl Order {
    fn note(&mut self, row: u16, col: u16) {
        if self.last.is_some_and(|last| (row, col) <= last) {
            self.ascending = false;
        }
        self.last = Some((row, col));
    }
}

impl CellSink for Order {
    fn accept(
        &mut self,
        record: &CellRecord,
        _occurrence: CellOccurrence,
        scan: &ScanContext<'_>,
        _strings: &mut SharedStringResolver<'_>,
    ) -> Result<()> {
        // Reached for formulas, and for records the scan parses whatever the
        // sink wants.
        scan.owner
            .check_cell_xf(scan.formatting, cell_xf_index(record), record.row());
        if matches!(record, CellRecord::Formula { .. })
            && crate::cell::Cell::from_record_with_formula_context(record, None, None, None)
                .formula()
                .is_none()
        {
            self.needs_formula_context = true;
        }
        self.note(record.row(), record.col());
        Ok(())
    }

    fn defers_string_formula(&self, _record: &CellRecord) -> bool {
        false
    }

    fn wants(&self, _row: u16, _col: u16) -> bool {
        false
    }

    fn wants_formula(&self, _row: u16, _col: u16) -> bool {
        true
    }

    fn accept_measured(
        &mut self,
        measured: &MeasuredCell,
        _occurrence: CellOccurrence,
        scan: &ScanContext<'_>,
    ) -> Result<()> {
        scan.owner
            .check_cell_xf(scan.formatting, measured.xf_index, measured.row);
        self.note(measured.row, measured.col);
        Ok(())
    }
}

#[cfg(test)]
mod tests {
    use std::path::{Path, PathBuf};

    use super::super::{SourceBackedCell, SourceBackedWorkbook};

    fn corpus(directory: &Path, out: &mut Vec<PathBuf>) {
        let Ok(entries) = std::fs::read_dir(directory) else {
            return;
        };
        for entry in entries.flatten() {
            let path = entry.path();
            if path.is_dir() {
                corpus(&path, out);
            } else if path
                .extension()
                .is_some_and(|extension| extension.eq_ignore_ascii_case("xls"))
            {
                out.push(path);
            }
        }
    }

    /// Every cell the whole-sheet visitor reports, then its result.
    fn visited(sheet: &super::SourceBackedWorksheet) -> (Vec<SourceBackedCell>, Option<String>) {
        let mut cells = Vec::new();
        let result = sheet.visit_cells(|cell| {
            cells.push(cell);
            Ok(())
        });
        (cells, result.err().map(|error| error.to_string()))
    }

    /// Every cell the cursor returns, then the error that ended it; `None`
    /// when the sheet's positions do not ascend.
    fn streamed(
        sheet: &super::SourceBackedWorksheet,
    ) -> Option<(Vec<SourceBackedCell>, Option<String>)> {
        let mut cells = Vec::new();
        let cursor = match sheet.cell_cursor() {
            Ok(cursor) => cursor?,
            Err(error) => return Some((cells, Some(error.to_string()))),
        };
        for cell in cursor {
            match cell {
                Ok(cell) => cells.push(cell),
                Err(error) => return Some((cells, Some(error.to_string()))),
            }
        }
        Some((cells, None))
    }

    #[test]
    fn the_cursor_returns_what_the_visitor_reports_on_the_corpus() {
        let root = Path::new(env!("CARGO_MANIFEST_DIR")).join("../../test-data");
        let mut files = Vec::new();
        corpus(&root, &mut files);
        files.sort();
        let (mut sheets, mut cells, mut batched, mut unordered) = (0, 0, 0, 0);
        for path in files {
            let Ok(workbook) = SourceBackedWorkbook::open(&path) else {
                continue;
            };
            let Ok(worksheets) = workbook.worksheets() else {
                continue;
            };
            for sheet in worksheets {
                let expected = visited(&sheet);
                let Some(actual) = streamed(&sheet) else {
                    // No cursor: the positions do not ascend.
                    assert!(expected.1.is_none(), "{}", path.display());
                    assert!(
                        !expected.0.windows(2).all(|pair| {
                            (pair[0].row(), pair[0].column()) < (pair[1].row(), pair[1].column())
                        }),
                        "{}",
                        path.display()
                    );
                    unordered += 1;
                    continue;
                };
                // The same result: the visitor's cells in order, or the
                // same refusal. A refusal the validating scan raises comes
                // before any cell.
                assert_eq!(actual.1, expected.1, "{}", path.display());
                if expected.1.is_none() {
                    assert_eq!(actual.0, expected.0, "{}", path.display());
                } else {
                    assert!(expected.0.starts_with(&actual.0), "{}", path.display());
                }
                sheets += 1;
                cells += expected.0.len();
                batched += usize::from(expected.0.len() > super::BATCH);
            }
        }
        // The corpus exercises resumed scans. Its out-of-order sheets were
        // embedded chart caches read as cells, which record 0316 skips; the
        // unordered case is pinned by `a_sheet_whose_positions_descend_has_no_cursor`.
        assert!(
            sheets > 100 && cells > 10_000 && batched > 10,
            "{sheets} {cells} {batched} {unordered}"
        );
    }

    #[test]
    fn a_sheet_whose_positions_descend_has_no_cursor() {
        use std::io::Cursor;

        let mut builder = crate::writer::Builder::new();
        let sheet = builder.add_worksheet("Data").unwrap();
        builder.write_number(sheet, 0, 0, 1.0).unwrap();
        builder.write_number(sheet, 1, 0, 2.0).unwrap();
        let bytes = builder.build().unwrap().to_bytes().unwrap();
        let mut stream =
            litchi_cfb::OleFile::open(Cursor::new(&bytes), litchi_cfb::OleFileLimits::default())
                .unwrap()
                .open_stream(&["Workbook"])
                .unwrap();
        // Swap the two Number records, which have the same length.
        let mut numbers = Vec::new();
        let mut offset = 0;
        while offset + 4 <= stream.len() {
            let kind = u16::from_le_bytes([stream[offset], stream[offset + 1]]);
            let length = usize::from(u16::from_le_bytes([stream[offset + 2], stream[offset + 3]]));
            if kind == 0x0203 {
                numbers.push(offset);
            }
            offset += 4 + length;
        }
        let [first, second] = numbers[..] else {
            panic!("two Number records");
        };
        let a = stream[first..first + 18].to_vec();
        let b = stream[second..second + 18].to_vec();
        stream[first..first + 18].copy_from_slice(&b);
        stream[second..second + 18].copy_from_slice(&a);
        let mut writer = litchi_cfb::OleWriter::new();
        writer.create_stream(&["Workbook"], &stream).unwrap();
        let mut package = Vec::new();
        writer.write_to(&mut Cursor::new(&mut package)).unwrap();

        let workbook = SourceBackedWorkbook::from_read_at(std::sync::Arc::new(
            litchi_core::OwnedSource::new(package),
        ))
        .unwrap();
        let sheet = workbook.worksheet_by_index(0).unwrap().unwrap();
        assert!(streamed(&sheet).is_none());
        assert_eq!(visited(&sheet).0.len(), 2);
    }

    #[test]
    fn a_cursor_that_stops_early_scans_one_batch() {
        let mut builder = crate::writer::Builder::new();
        let sheet = builder.add_worksheet("Data").unwrap();
        for row in 0..10_000_u32 {
            builder
                .write_number(sheet, row, 0, f64::from(row) + 0.5)
                .unwrap();
        }
        let bytes = builder.build().unwrap().to_bytes().unwrap();
        let workbook = SourceBackedWorkbook::from_read_at(std::sync::Arc::new(
            litchi_core::OwnedSource::new(bytes),
        ))
        .unwrap();
        let sheet = workbook.worksheet_by_index(0).unwrap().unwrap();
        let mut cursor = sheet.cell_cursor().unwrap().unwrap();
        let first = cursor.next().unwrap().unwrap();
        assert_eq!((first.row(), first.column()), (0, 0));
        assert!(cursor.batch.len() < 2 * super::BATCH);
        assert_eq!(cursor.count(), 9_999);
    }
}
