//! A streaming cursor over the stored cells of a source-backed worksheet
//! (record 0267).
//!
//! [`SourceBackedWorksheet::grid`] decodes every cell record of the
//! worksheet stream before the first cell can be read, so a caller that
//! stops after a few cells (a query's `first()` or `limit(n)`) pays for the
//! whole sheet. The cursor decodes the stream's cells in batches with the
//! reader the grid uses ([`CellsReader`]), under the same shared strings,
//! formula context, style count and leniency. A batch of N cells reads the
//! stream again from its start and stops after the cells already returned
//! plus N; the first batch holds [`FIRST_BATCH`] cells and each later one
//! [`GROWTH`] times more, so the decoding repeated for a longer prefix stays
//! a constant fraction of what was read, and once [`MAX_STREAMED`] cells
//! were returned the cursor reads the complete grid and continues after the
//! last coordinate it returned.
//!
//! # What the cursor checks before the first cell
//!
//! The cursor is created only after every refusal the grid's construction
//! raises before it decodes cells (the formula context, unsupported
//! worksheet adjuncts, sparkline groups), after the worksheet's scenarios and
//! comments parse as the grid parses them, and after one walk of the record
//! headers shows that the stored cells strictly ascend in row-major order.
//! The grid keeps one cell per coordinate, sorted; a worksheet whose records
//! do not ascend (a cell stored out of order or twice) is read through the
//! complete grid from its first cell. A batch whose decoding records a
//! repair is not used either: the cursor then reads the complete grid. So
//! the cursor returns exactly the grid's cells, in order, on every worksheet
//! whose cell records the grid decodes without refusal.
//!
//! What a batch cannot see is the stream after it: a cell record the grid
//! refuses after the cells read is refused when a later batch reaches it.
//!
//! # Resources
//!
//! The cursor holds the worksheet's verified decoded part, as the grid's
//! construction does, and one batch's cells in a worksheet snapshot. Every
//! read that consults the source runs the package's freshness fences.

use std::collections::VecDeque;
use std::io::Cursor;
use std::sync::Arc;

use super::{MaterializationInputs, SourceBackedWorksheet, postflight_package, preflight_package};
use crate::leniency::ToleranceReport;
use crate::package::cells_reader::CellsReader;
use crate::package::error::{Error, Result};
use crate::package::records::Stream;
use crate::raw::{Records, kind};
use crate::sheet::{Grid, Kind, Worksheet};

/// Cells in the first batch.
const FIRST_BATCH: usize = 64;

/// How many times more cells each later batch holds.
const GROWTH: usize = 8;

/// Cells the cursor returns from batches before it reads the complete grid.
const MAX_STREAMED: usize = FIRST_BATCH * (1 + GROWTH + GROWTH * GROWTH);

/// A cursor over a worksheet's stored cells, in row-major order.
///
/// Returned by [`SourceBackedWorksheet::cell_cursor`]. Each item is a stored
/// cell's zero-based `(row, column)`; [`Self::grid`] holds it: a snapshot of
/// the current batch's cells while the cursor decodes batches, then the
/// worksheet's complete grid ([`Self::is_complete`]). The cells are those of
/// [`SourceBackedWorksheet::grid`], in its order (see the module
/// documentation for refusals after the cells read).
///
/// ```no_run
/// # fn main() -> Result<(), Box<dyn std::error::Error>> {
/// use litchi_xlsb::SourceBackedWorkbook;
/// let workbook = SourceBackedWorkbook::from_reader(std::fs::File::open("book.xlsb")?)?;
/// let sheet = workbook.worksheet_by_index(0)?.expect("worksheet");
/// let mut cursor = sheet.cell_cursor()?;
/// if let Some(at) = cursor.next().transpose()? {
///     let cell = cursor.grid().expect("its grid").cell_at(at)?;
///     # let _ = cell;
/// }
/// # Ok(())
/// # }
/// ```
pub struct CellCursor {
    sheet: SourceBackedWorksheet,
    stage: Stage,
    last: Option<(u32, u32)>,
}

enum Stage {
    Streaming(Box<Streaming>),
    /// The complete grid, walked after the last cell returned.
    Complete {
        grid: Grid,
    },
    /// The complete grid, not read yet.
    Pending,
    Done,
}

/// The first batch of a validated worksheet, kept for later cursors.
pub(super) struct FirstBatch {
    grid: Grid,
    cells: Vec<(u32, u32)>,
    finished: bool,
}

impl std::fmt::Debug for FirstBatch {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("FirstBatch")
            .field("cells", &self.cells.len())
            .field("finished", &self.finished)
            .finish_non_exhaustive()
    }
}

struct Streaming {
    /// Read when a batch is decoded; a cursor that starts from the kept
    /// first batch reads it only to decode a second one.
    inputs: Option<MaterializationInputs>,
    /// Cells returned from earlier batches.
    read: usize,
    /// Cells the next batch decodes.
    size: usize,
    /// The current batch's snapshot and its cells not returned yet.
    grid: Option<Grid>,
    batch: VecDeque<(u32, u32)>,
    /// The stream has no cell after the current batch.
    finished: bool,
}

impl SourceBackedWorksheet {
    /// A cursor over this worksheet's stored cells in row-major order, which
    /// decodes only as many cells as the caller reads; see [`CellCursor`].
    ///
    /// On a worksheet whose grid is resident in the workbook's cache, or
    /// whose cell records do not strictly ascend, the cursor reads the
    /// complete grid.
    ///
    /// # Errors
    ///
    /// The refusals [`Self::grid`] raises before it decodes cells, and the
    /// refusals of the worksheet's scenarios and comments; the source and
    /// cancellation errors of every source-backed read.
    pub fn cell_cursor(&self) -> Result<CellCursor> {
        let _ = preflight_package(&self.inner.package)?;
        let stage = (|| {
            if !matches!(self.metadata().kind, Kind::Worksheet) {
                return Err(Error::UnsupportedFeature(
                    "source-backed XLSB cell cursors require a worksheet tab",
                ));
            }
            if self
                .inner
                .worksheet_cache
                .get(self.catalog_position, false)
                .is_some()
            {
                return Ok(Stage::Pending);
            }
            // A worksheet an earlier cursor validated starts from its first
            // batch: the source is the immutable snapshot that cursor read.
            if let Some(first) = self.metadata().first_batch.get() {
                return Ok(Stage::Streaming(Box::new(Streaming {
                    inputs: None,
                    read: first.cells.len(),
                    size: FIRST_BATCH * GROWTH,
                    grid: Some(first.grid.clone()),
                    batch: first.cells.iter().copied().collect(),
                    finished: first.finished,
                })));
            }
            let partname = self.physical_partname()?.clone();
            let inputs = self.materialization_inputs(partname)?;
            if !cells_ascend(inputs.data.as_bytes(), self.inner.raw_limits)? {
                return Ok(Stage::Pending);
            }
            let _ = crate::package::scenarios::parse_worksheet(
                inputs.data.as_bytes(),
                self.inner.raw_limits,
            )?;
            if let Some(partname) = &inputs.comments_part {
                let _ = self.read_comments(partname, &mut ToleranceReport::default())?;
            }
            Ok(Stage::Streaming(Box::new(Streaming {
                inputs: Some(inputs),
                read: 0,
                size: FIRST_BATCH,
                grid: None,
                batch: VecDeque::new(),
                finished: false,
            })))
        })();
        let fence = postflight_package(&self.inner.package);
        fence?;
        Ok(CellCursor {
            sheet: self.clone(),
            stage: stage?,
            last: None,
        })
    }
}

impl CellCursor {
    /// The grid that holds the cells returned so far: the current batch's
    /// snapshot or the worksheet's complete grid. `None` before the first
    /// cell and once the cursor ended.
    #[must_use]
    pub fn grid(&self) -> Option<&Grid> {
        match &self.stage {
            Stage::Streaming(streaming) => streaming.grid.as_ref(),
            Stage::Complete { grid, .. } => Some(grid),
            Stage::Pending | Stage::Done => None,
        }
    }

    /// Whether [`Self::grid`] is the worksheet's complete grid.
    #[must_use]
    pub fn is_complete(&self) -> bool {
        matches!(self.stage, Stage::Complete { .. })
    }

    fn step(&mut self) -> Result<Option<(u32, u32)>> {
        loop {
            match &mut self.stage {
                Stage::Done => return Ok(None),
                Stage::Pending => self.stage = self.complete()?,
                Stage::Complete { grid } => return Ok(grid.worksheet().cell_after(self.last)),
                Stage::Streaming(streaming) => {
                    if let Some(at) = streaming.batch.pop_front() {
                        return Ok(Some(at));
                    }
                    if streaming.finished {
                        self.stage = Stage::Done;
                        return Ok(None);
                    }
                    if streaming.read >= MAX_STREAMED {
                        self.stage = Stage::Pending;
                        continue;
                    }
                    if !self.sheet.refill(streaming, self.last)? {
                        self.stage = Stage::Pending;
                    }
                },
            }
        }
    }

    /// The complete grid; the walk continues after the last cell returned.
    fn complete(&self) -> Result<Stage> {
        Ok(Stage::Complete {
            grid: self.sheet.grid()?,
        })
    }
}

impl SourceBackedWorksheet {
    /// Decode the next batch into `streaming`; `false` when it cannot be
    /// used (a repair, or cells that do not continue after `last`), and the
    /// cursor reads the complete grid instead.
    fn refill(&self, streaming: &mut Streaming, last: Option<(u32, u32)>) -> Result<bool> {
        let _ = preflight_package(&self.inner.package)?;
        if streaming.inputs.is_none() {
            let partname = self.physical_partname()?.clone();
            streaming.inputs = Some(self.materialization_inputs(partname)?);
        }
        let batch: Result<Option<(Worksheet, VecDeque<(u32, u32)>, bool)>> = (|| {
            let Some(inputs) = &streaming.inputs else {
                return Ok(None);
            };
            let context = inputs.formula_context.for_sheet(inputs.workbook_position);
            let mut reader = CellsReader::new(
                Stream::new(Cursor::new(inputs.data.as_bytes()), self.inner.raw_limits),
                inputs.shared_strings.as_slice(),
                &context,
                inputs.cell_xf_count,
                self.inner.raw_limits,
                self.inner.leniency,
            )?;
            for _ in 0..streaming.read {
                if reader.next_cell()?.is_none() {
                    return Ok(None);
                }
            }
            let mut worksheet = Worksheet::new(inputs.name.clone());
            let mut cells = VecDeque::new();
            let mut previous = last;
            let mut finished = false;
            while cells.len() < streaming.size {
                let Some(cell) = reader.next_cell()? else {
                    finished = true;
                    break;
                };
                let at = {
                    use litchi_core::sheet::Cell as _;
                    (cell.row(), cell.column())
                };
                if previous.is_some_and(|previous| at <= previous) {
                    return Ok(None);
                }
                previous = Some(at);
                cells.push_back(at);
                worksheet.add_cell(cell);
            }
            if !reader.tolerance.is_clean() {
                return Ok(None);
            }
            Ok(Some((worksheet, cells, finished)))
        })();
        let fence = postflight_package(&self.inner.package);
        fence?;
        let Some((worksheet, cells, finished)) = batch? else {
            return Ok(false);
        };
        let grid = Grid::from_arc(Arc::new(worksheet));
        if streaming.read == 0 {
            let _kept = self.metadata().first_batch.set(Arc::new(FirstBatch {
                grid: grid.clone(),
                cells: cells.iter().copied().collect(),
                finished,
            }));
        }
        streaming.read += cells.len();
        streaming.size = streaming.size.saturating_mul(GROWTH);
        streaming.grid = Some(grid);
        streaming.batch = cells;
        streaming.finished = finished;
        Ok(true)
    }
}

impl Iterator for CellCursor {
    type Item = Result<(u32, u32)>;

    fn next(&mut self) -> Option<Self::Item> {
        match self.step() {
            Ok(Some(at)) => {
                self.last = Some(at);
                Some(Ok(at))
            },
            Ok(None) => None,
            Err(error) => {
                self.stage = Stage::Done;
                Some(Err(error))
            },
        }
    }
}

impl std::iter::FusedIterator for CellCursor {}

impl std::fmt::Debug for CellCursor {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("CellCursor")
            .field("complete", &self.is_complete())
            .finish_non_exhaustive()
    }
}

/// Whether the worksheet stream's cell records strictly ascend in row-major
/// order: each cell record's column, under the row of the last row header
/// before it, comes after the previous cell's position.
///
/// The walk reads record headers and the leading row or column of a
/// payload. It takes every record of a cell kind as a cell, including one
/// the reader skips, so the reader's cells are a subsequence of the walked
/// ones and ascend whenever they do. A row header or cell payload too short
/// to hold its number gives `false`.
fn cells_ascend(data: &[u8], raw_limits: crate::raw::Limits) -> Result<bool> {
    let mut row = 0_u32;
    let mut last: Option<(u32, u32)> = None;
    for record in Records::new(data, raw_limits) {
        let record = record?;
        let kind = record.kind();
        let is_cell = (kind::CELL_BLANK.get()..=kind::FMLA_ERROR.get()).contains(&kind.get())
            || kind == kind::CELL_R_STRING;
        if kind != kind::ROW_HDR && !is_cell {
            continue;
        }
        let Some(head) = record.payload().get(..4) else {
            return Ok(false);
        };
        let number = u32::from_le_bytes([head[0], head[1], head[2], head[3]]);
        if kind == kind::ROW_HDR {
            row = number;
            continue;
        }
        let at = (row, number);
        if last.is_some_and(|last| at <= last) {
            return Ok(false);
        }
        last = Some(at);
    }
    Ok(true)
}

#[cfg(test)]
mod tests {
    use std::io::Cursor;
    use std::path::{Path, PathBuf};

    use litchi_core::sheet::{Cell as _, CellValue};

    use super::super::SourceBackedWorkbook;

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
                .is_some_and(|extension| extension.eq_ignore_ascii_case("xlsb"))
            {
                out.push(path);
            }
        }
    }

    type Listed = Vec<((u32, u32), CellValue)>;

    /// The complete grid's cells, in order.
    fn complete(sheet: &super::SourceBackedWorksheet) -> Option<Listed> {
        let grid = sheet.grid().ok()?;
        Some(
            grid.cells_in(crate::sheet::Rect::ALL)
                .unwrap()
                .map(|cell| ((cell.row(), cell.column()), cell.value().clone()))
                .collect(),
        )
    }

    /// The cursor's cells, each read from the grid it names.
    fn streamed(sheet: &super::SourceBackedWorksheet, path: &Path) -> Option<Listed> {
        let mut cursor = sheet.cell_cursor().ok()?;
        let mut out = Vec::new();
        while let Some(at) = cursor.next() {
            let at = at.unwrap_or_else(|error| panic!("{}: {error:?}", path.display()));
            let cell = cursor.grid().unwrap().cell_at(at).unwrap().unwrap();
            out.push((at, cell.value().clone()));
        }
        Some(out)
    }

    #[test]
    fn the_cursor_returns_the_grids_cells_on_the_corpus() {
        let root = Path::new(env!("CARGO_MANIFEST_DIR")).join("../../test-data");
        let mut files = Vec::new();
        corpus(&root, &mut files);
        files.sort();
        let mut sheets = 0;
        for path in files {
            let Ok(bytes) = std::fs::read(&path) else {
                continue;
            };
            let Ok(workbook) = SourceBackedWorkbook::from_reader(Cursor::new(bytes)) else {
                continue;
            };
            let Ok(worksheets) = workbook.worksheets() else {
                continue;
            };
            for sheet in worksheets {
                let streamed = streamed(&sheet, &path);
                let complete = complete(&sheet);
                match streamed {
                    // A cursor refused before its first cell: so is the grid.
                    None => assert!(complete.is_none(), "{}", path.display()),
                    Some(streamed) => {
                        assert_eq!(Some(streamed), complete, "{}", path.display());
                        sheets += 1;
                    },
                }
            }
        }
        assert!(sheets >= 10, "{sheets}");
    }

    fn generated(rows: u32, order: impl IntoIterator<Item = u32>) -> SourceBackedWorkbook {
        let mut sheet = crate::builder::Worksheet::new("Data");
        for row in order.into_iter().take(rows as usize) {
            sheet.set_cell(row, 0, f64::from(row) + 0.5);
            sheet.set_cell(row, 1, format!("text {row}"));
            sheet.set_cell(row, 2, row % 2 == 0);
        }
        let mut builder = crate::Builder::new();
        builder.add_worksheet(sheet);
        let bytes = builder.build().unwrap().to_bytes().unwrap();
        SourceBackedWorkbook::from_reader(Cursor::new(bytes)).unwrap()
    }

    #[test]
    fn batches_then_the_complete_grid_return_the_grids_cells() {
        let workbook = generated(5_000, 0..5_000);
        let sheet = workbook.worksheet_by_index(0).unwrap().unwrap();
        let mut cursor = sheet.cell_cursor().unwrap();
        let first = cursor.next().unwrap().unwrap();
        assert_eq!(first, (0, 0));
        assert!(!cursor.is_complete());
        assert!(cursor.grid().unwrap().stored_cell_count() <= super::FIRST_BATCH);
        let mut stages = vec![false];
        while let Some(at) = cursor.next() {
            at.unwrap();
            if stages.last() != Some(&cursor.is_complete()) {
                stages.push(cursor.is_complete());
            }
        }
        // Batches, then the complete grid.
        assert_eq!(stages, [false, true]);
        let fresh = generated(5_000, 0..5_000);
        let sheet = fresh.worksheet_by_index(0).unwrap().unwrap();
        assert_eq!(
            streamed(&sheet, Path::new("generated workbook")),
            complete(&sheet)
        );
        // Once the grid is resident, the cursor reads it.
        let mut cursor = sheet.cell_cursor().unwrap();
        cursor.next().unwrap().unwrap();
        assert!(cursor.is_complete());
    }

    #[test]
    fn a_later_cursor_starts_from_the_kept_first_batch() {
        let workbook = generated(5_000, 0..5_000);
        let sheet = workbook.worksheet_by_index(0).unwrap().unwrap();
        let first = sheet.cell_cursor().unwrap().next().unwrap().unwrap();
        let before = workbook.cache_diagnostics();
        let again = workbook.worksheet_by_index(0).unwrap().unwrap();
        let mut cursor = again.cell_cursor().unwrap();
        assert_eq!(cursor.next().unwrap().unwrap(), first);
        let after = workbook.cache_diagnostics();
        // No part was read for it.
        assert_eq!(
            (after.hits, after.cold_loads),
            (before.hits, before.cold_loads)
        );
        // Reading on decodes the stream again, with the same cells.
        let rest: Vec<_> = cursor.map(Result::unwrap).collect();
        let fresh = generated(5_000, 0..5_000);
        let sheet = fresh.worksheet_by_index(0).unwrap().unwrap();
        let expected: Vec<_> = complete(&sheet)
            .unwrap()
            .into_iter()
            .map(|(at, _)| at)
            .collect();
        assert_eq!([vec![first], rest].concat(), expected);
    }

    fn records(records: &[(crate::raw::Kind, &[u8])]) -> Vec<u8> {
        let mut out = Vec::new();
        let mut encoder =
            crate::raw::Encoder::bounded(&mut out, crate::raw::Limits::DEFAULT).unwrap();
        for (kind, payload) in records {
            encoder.write_record(*kind, payload).unwrap();
        }
        drop(encoder);
        out
    }

    #[test]
    fn the_order_walk_reads_rows_and_columns() {
        use crate::raw::kind;
        let row = |number: u32| number.to_le_bytes().repeat(4);
        let cell = |column: u32| [column.to_le_bytes(), [0; 4]].concat();
        let (r1, r2, a, b) = (row(1), row(2), cell(0), cell(1));
        let ascending = records(&[
            (kind::ROW_HDR, &r1),
            (kind::CELL_BLANK, &a),
            (kind::CELL_RK, &b),
            (kind::ROW_HDR, &r2),
            (kind::FMLA_NUM, &a),
        ]);
        assert!(super::cells_ascend(&ascending, crate::raw::Limits::DEFAULT).unwrap());
        let backwards = records(&[
            (kind::ROW_HDR, &r2),
            (kind::CELL_BLANK, &a),
            (kind::ROW_HDR, &r1),
            (kind::CELL_BLANK, &b),
        ]);
        assert!(!super::cells_ascend(&backwards, crate::raw::Limits::DEFAULT).unwrap());
        let twice = records(&[
            (kind::ROW_HDR, &r1),
            (kind::CELL_BLANK, &b),
            (kind::CELL_R_STRING, &b),
        ]);
        assert!(!super::cells_ascend(&twice, crate::raw::Limits::DEFAULT).unwrap());
        let short = records(&[(kind::ROW_HDR, &r1[..2])]);
        assert!(!super::cells_ascend(&short, crate::raw::Limits::DEFAULT).unwrap());
    }
}
