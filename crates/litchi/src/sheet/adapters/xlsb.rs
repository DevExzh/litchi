//! XLSB-to-umbrella spreadsheet bridge.
//!
//! The standalone XLSB owner exposes a source-backed worksheet catalog whose
//! worksheet streams are loaded on demand. This adapter keeps that deferred
//! ownership at the facade boundary while implementing the older dynamic
//! spreadsheet traits.

use crate::xlsb;
use litchi_core::ReadAt;
use litchi_core::sheet::{
    Cell as CoreCell, CellIterator, CellValue, Result as SheetResult, RowIterator, WorkbookTrait,
    Worksheet as CoreWorksheet, WorksheetIterator,
};
use litchi_opc::{OpcError, ReadLimits};
use std::borrow::Cow;
use std::io::{self, Read, Seek, SeekFrom};
use std::sync::{Arc, Mutex, OnceLock};

type BoxError = Box<dyn std::error::Error + Send + Sync>;
type XlsbError = xlsb::package::PackageError;

fn boxed_error(message: impl Into<String>) -> BoxError {
    Box::new(litchi_core::Error::Other(message.into()))
}

fn boxed_xlsb_error(error: XlsbError) -> BoxError {
    match error {
        XlsbError::Opc(OpcError::SourceChanged { expected, actual }) => {
            Box::new(litchi_core::Error::SourceChanged {
                expected,
                observed: actual,
            })
        },
        error => Box::new(error),
    }
}

/// A checked `Read + Seek` view over an immutable positional source.
///
/// The eager XLSB compatibility loader requires a conventional stream, while
/// source-backed owners intentionally expose only positional reads. Keeping
/// this bridge local ensures the fallback cannot accidentally share a cursor
/// with another operation.
struct ReadAtReader {
    source: Arc<dyn ReadAt>,
    position: u64,
}

impl ReadAtReader {
    fn new(source: Arc<dyn ReadAt>) -> Self {
        Self {
            source,
            position: 0,
        }
    }
}

impl Read for ReadAtReader {
    fn read(&mut self, output: &mut [u8]) -> io::Result<usize> {
        if output.is_empty() {
            return Ok(0);
        }
        let read = self.source.read_at(self.position, output)?;
        if read > output.len() {
            return Err(io::Error::new(
                io::ErrorKind::InvalidData,
                "positional source reported more bytes than requested",
            ));
        }
        let next = self
            .position
            .checked_add(u64::try_from(read).map_err(|_| {
                io::Error::new(io::ErrorKind::InvalidData, "read count does not fit u64")
            })?)
            .ok_or_else(|| {
                io::Error::new(io::ErrorKind::InvalidInput, "stream position overflow")
            })?;
        self.position = next;
        Ok(read)
    }
}

impl Seek for ReadAtReader {
    fn seek(&mut self, position: SeekFrom) -> io::Result<u64> {
        let next = match position {
            SeekFrom::Start(offset) => offset,
            SeekFrom::Current(offset) => checked_seek(self.position, offset)?,
            SeekFrom::End(offset) => checked_seek(self.source.len()?, offset)?,
        };
        self.position = next;
        Ok(next)
    }
}

fn checked_seek(base: u64, offset: i64) -> io::Result<u64> {
    if offset >= 0 {
        base.checked_add(u64::try_from(offset).map_err(|_| {
            io::Error::new(io::ErrorKind::InvalidInput, "seek offset does not fit u64")
        })?)
        .ok_or_else(|| io::Error::new(io::ErrorKind::InvalidInput, "seek position overflow"))
    } else {
        base.checked_sub(offset.unsigned_abs()).ok_or_else(|| {
            io::Error::new(io::ErrorKind::InvalidInput, "seek before start of stream")
        })
    }
}

/// Internal dynamic-trait view over a source-backed XLSB workbook.
pub(crate) struct Workbook {
    workbook: xlsb::SourceBackedWorkbook,
    source: Arc<dyn ReadAt>,
    limits: ReadLimits,
    names: Box<[String]>,
    date1904: bool,
    eager: OnceLock<xlsb::Workbook>,
    eager_init: Mutex<()>,
    worksheets: Box<[OnceLock<Arc<xlsb::Worksheet>>]>,
    worksheet_init: Box<[Mutex<()>]>,
}

impl std::fmt::Debug for Workbook {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("Workbook")
            .field("worksheet_count", &self.names.len())
            .field("date1904", &self.date1904)
            .finish()
    }
}

impl Workbook {
    /// Create an adapter from the source owner and the exact source used to
    /// build it. The date-system value is retained by the facade so no eager
    /// compatibility parse is needed for ordinary catalog queries.
    pub(crate) fn from_source_backed(
        workbook: xlsb::SourceBackedWorkbook,
        source: Arc<dyn ReadAt>,
        limits: ReadLimits,
    ) -> SheetResult<Self> {
        let names = workbook
            .worksheet_names()
            .map_err(boxed_xlsb_error)?
            .into_boxed_slice();
        let catalog_count = workbook.sheet_count().map_err(boxed_xlsb_error)?;
        let date1904 = workbook.is_1904_date_system().map_err(boxed_xlsb_error)?;
        let worksheets = std::iter::repeat_with(OnceLock::new)
            .take(catalog_count)
            .collect::<Vec<_>>()
            .into_boxed_slice();
        let worksheet_init = std::iter::repeat_with(|| Mutex::new(()))
            .take(catalog_count)
            .collect::<Vec<_>>()
            .into_boxed_slice();
        Ok(Self {
            workbook,
            source,
            limits,
            names,
            date1904,
            eager: OnceLock::new(),
            eager_init: Mutex::new(()),
            worksheets,
            worksheet_init,
        })
    }

    pub(crate) fn ensure_source_current(&self) -> SheetResult<()> {
        self.workbook
            .source_version()
            .map(|_| ())
            .map_err(boxed_xlsb_error)
    }

    fn eager_workbook(&self) -> SheetResult<&xlsb::Workbook> {
        self.ensure_source_current()?;
        if self.eager.get().is_none() {
            let _guard = self
                .eager_init
                .lock()
                .map_err(|_| boxed_error("XLSB eager initialization lock was poisoned"))?;
            if self.eager.get().is_none() {
                let reader = ReadAtReader::new(Arc::clone(&self.source));
                let eager = xlsb::Workbook::new_with_limits(reader, self.limits)
                    .map_err(boxed_xlsb_error)?;
                self.ensure_source_current()?;
                self.eager.set(eager).map_err(|_| {
                    boxed_error("XLSB eager compatibility workbook was already published")
                })?;
            }
        }
        self.eager
            .get()
            .ok_or_else(|| boxed_error("XLSB eager compatibility workbook was not published"))
    }

    fn eager_worksheet(
        &self,
        index: usize,
        workbook: &xlsb::Workbook,
    ) -> SheetResult<Arc<xlsb::Worksheet>> {
        let source = self
            .worksheets
            .get(index)
            .ok_or_else(|| boxed_error("XLSB eager worksheet cache position is out of bounds"))?;
        let init = self
            .worksheet_init
            .get(index)
            .ok_or_else(|| boxed_error("XLSB eager worksheet lock position is out of bounds"))?;
        if let Some(worksheet) = source.get() {
            return Ok(Arc::clone(worksheet));
        }
        let _guard = init
            .lock()
            .map_err(|_| boxed_error("XLSB worksheet initialization lock was poisoned"))?;
        if let Some(worksheet) = source.get() {
            return Ok(Arc::clone(worksheet));
        }
        let worksheet = Arc::new(workbook.worksheet(index).map_err(boxed_xlsb_error)?);
        self.ensure_source_current()?;
        source.set(Arc::clone(&worksheet)).map_err(|error| {
            drop(error);
            boxed_error("XLSB eager worksheet was already published")
        })?;
        Ok(worksheet)
    }

    fn materialize<'a, T>(
        &'a self,
        handle: &SourceBackedWorksheet<'a>,
        operation: impl FnOnce(&dyn CoreWorksheet) -> SheetResult<T>,
    ) -> SheetResult<T> {
        self.ensure_source_current()?;
        let source = self
            .worksheets
            .get(handle.catalog_position)
            .ok_or_else(|| boxed_error("XLSB source worksheet cache position is out of bounds"))?;
        let init = self
            .worksheet_init
            .get(handle.catalog_position)
            .ok_or_else(|| boxed_error("XLSB source worksheet lock position is out of bounds"))?;
        let result = match source.get() {
            Some(worksheet) => operation(worksheet.as_ref()),
            None => {
                let guard = init
                    .lock()
                    .map_err(|_| boxed_error("XLSB worksheet initialization lock was poisoned"))?;
                if let Some(worksheet) = source.get() {
                    operation(worksheet.as_ref())
                } else {
                    match handle.handle.materialize() {
                        Ok(worksheet) => {
                            self.ensure_source_current()?;
                            source.set(Arc::new(worksheet)).map_err(|_| {
                                boxed_error("XLSB source worksheet was already published")
                            })?;
                            operation(
                                source
                                    .get()
                                    .ok_or_else(|| {
                                        boxed_error("XLSB source worksheet was not published")
                                    })?
                                    .as_ref(),
                            )
                        },
                        Err(XlsbError::UnsupportedFeature(_)) => {
                            drop(guard);
                            let workbook = self.eager_workbook()?;
                            let worksheet =
                                self.eager_worksheet(handle.catalog_position, workbook)?;
                            operation(worksheet.as_ref())
                        },
                        Err(error) => Err(boxed_xlsb_error(error)),
                    }
                }
            },
        };
        self.ensure_source_current()?;
        result
    }

    pub(crate) fn text(&self) -> SheetResult<String> {
        match self.workbook.text() {
            Ok(text) => Ok(text),
            Err(XlsbError::UnsupportedFeature(_)) => self.eager_text(),
            Err(error) => Err(boxed_xlsb_error(error)),
        }
    }

    fn eager_text(&self) -> SheetResult<String> {
        let workbook = self.eager_workbook()?;
        let source_worksheets = self.workbook.worksheets().map_err(boxed_xlsb_error)?;
        let mut output = String::new();
        for source_worksheet in source_worksheets {
            self.ensure_source_current()?;
            let catalog_position = source_worksheet
                .workbook_position()
                .map_err(boxed_xlsb_error)?;
            let worksheet = self.eager_worksheet(catalog_position, workbook)?;
            let mut rows = worksheet.rows();
            while let Some(row) = rows.next() {
                let row = row?;
                for (column, cell) in row.iter().enumerate() {
                    if column != 0 {
                        output.push('\t');
                    }
                    append_eager_cell_text(&mut output, cell);
                }
                output.push('\n');
            }
        }
        self.ensure_source_current()?;
        Ok(output)
    }

    fn worksheet(&self, index: usize) -> SheetResult<Box<dyn CoreWorksheet + '_>> {
        let handle = self
            .workbook
            .worksheet_by_index(index)
            .map_err(boxed_xlsb_error)?
            .ok_or_else(|| boxed_error(format!("XLSB worksheet index {index} is out of bounds")))?;
        let catalog_position = handle.workbook_position().map_err(boxed_xlsb_error)?;
        Ok(Box::new(SourceBackedWorksheet::new(
            self,
            catalog_position,
            handle,
        )?))
    }

    fn worksheet_named(&self, name: &str) -> SheetResult<Box<dyn CoreWorksheet + '_>> {
        let handle = self
            .workbook
            .worksheet_by_name(name)
            .map_err(boxed_xlsb_error)?
            .ok_or_else(|| boxed_error(format!("XLSB worksheet '{name}' was not found")))?;
        let catalog_position = handle.workbook_position().map_err(boxed_xlsb_error)?;
        Ok(Box::new(SourceBackedWorksheet::new(
            self,
            catalog_position,
            handle,
        )?))
    }
}

fn append_eager_cell_text(output: &mut String, value: &CellValue) {
    match value {
        CellValue::Empty => {},
        CellValue::Bool(value) => output.push_str(if *value { "TRUE" } else { "FALSE" }),
        CellValue::Int(value) => output.push_str(&value.to_string()),
        CellValue::Float(value) | CellValue::DateTime(value) => output.push_str(&value.to_string()),
        CellValue::String(value) | CellValue::Error(value) => output.push_str(value),
        CellValue::Formula {
            formula,
            cached_value,
            ..
        } => match cached_value.as_deref() {
            Some(value) if !matches!(value, CellValue::Empty) => {
                append_eager_cell_text(output, value)
            },
            _ => {
                output.push('=');
                output.push_str(formula);
            },
        },
    }
}

impl WorkbookTrait for Workbook {
    fn active_sheet_index(&self) -> usize {
        0
    }

    fn active_worksheet(&self) -> SheetResult<Box<dyn CoreWorksheet + '_>> {
        self.worksheet(0)
    }

    fn is_1904_date_system(&self) -> bool {
        self.date1904
    }

    fn worksheet_by_index(&self, index: usize) -> SheetResult<Box<dyn CoreWorksheet + '_>> {
        self.worksheet(index)
    }

    fn worksheet_by_name(&self, name: &str) -> SheetResult<Box<dyn CoreWorksheet + '_>> {
        self.worksheet_named(name)
    }

    fn worksheet_count(&self) -> usize {
        self.names.len()
    }

    fn worksheet_names(&self) -> &[String] {
        &self.names
    }

    fn worksheets(&self) -> Box<dyn WorksheetIterator<'_> + '_> {
        Box::new(Worksheets {
            workbook: self,
            index: 0,
        })
    }
}

struct Worksheets<'a> {
    workbook: &'a Workbook,
    index: usize,
}

impl<'a> WorksheetIterator<'a> for Worksheets<'a> {
    fn next(&mut self) -> Option<SheetResult<Box<dyn CoreWorksheet + 'a>>> {
        if self.index >= self.workbook.worksheet_count() {
            return None;
        }
        let index = self.index;
        self.index += 1;
        Some(self.workbook.worksheet(index))
    }
}

struct SourceBackedWorksheet<'a> {
    workbook: &'a Workbook,
    handle: xlsb::SourceBackedWorksheet,
    catalog_position: usize,
    name: String,
}

impl<'a> SourceBackedWorksheet<'a> {
    fn new(
        workbook: &'a Workbook,
        catalog_position: usize,
        handle: xlsb::SourceBackedWorksheet,
    ) -> SheetResult<Self> {
        let name = handle.name().map_err(boxed_xlsb_error)?.to_owned();
        Ok(Self {
            workbook,
            handle,
            catalog_position,
            name,
        })
    }

    fn value_at(&self, row: u32, column: u32) -> SheetResult<CellValue> {
        self.workbook.materialize(self, |worksheet| {
            Ok(worksheet
                .cell(row.saturating_sub(1), column.saturating_sub(1))?
                .value()
                .clone())
        })
    }

    fn row_values(&self, row: usize) -> SheetResult<Vec<CellValue>> {
        self.workbook
            .materialize(self, |worksheet| Ok(worksheet.row(row)?.into_owned()))
    }

    fn copy_cell(cell: &dyn CoreCell) -> SheetResult<FacadeCell> {
        let row = cell
            .row()
            .checked_add(1)
            .ok_or_else(|| boxed_error("XLSB cell row overflow"))?;
        let column = cell
            .column()
            .checked_add(1)
            .ok_or_else(|| boxed_error("XLSB cell column overflow"))?;
        Ok(FacadeCell {
            row,
            column,
            value: cell.value().clone(),
        })
    }

    fn dimensions_inner(&self) -> SheetResult<Option<(u32, u32, u32, u32)>> {
        self.workbook.materialize(self, |worksheet| {
            let Some((min_row, min_column, max_row, max_column)) = worksheet.dimensions() else {
                return Ok(None);
            };
            let min_row = min_row
                .checked_add(1)
                .ok_or_else(|| boxed_error("XLSB worksheet minimum row overflow"))?;
            let min_column = min_column
                .checked_add(1)
                .ok_or_else(|| boxed_error("XLSB worksheet minimum column overflow"))?;
            let max_row = max_row
                .checked_add(1)
                .ok_or_else(|| boxed_error("XLSB worksheet maximum row overflow"))?;
            let max_column = max_column
                .checked_add(1)
                .ok_or_else(|| boxed_error("XLSB worksheet maximum column overflow"))?;
            Ok(Some((min_row, min_column, max_row, max_column)))
        })
    }

    fn stored_cells(&self) -> SheetResult<Vec<FacadeCell>> {
        self.workbook.materialize(self, |worksheet| {
            let mut cells = Vec::new();
            let mut iterator = worksheet.cells();
            while let Some(cell) = iterator.next() {
                cells.push(Self::copy_cell(cell?.as_ref())?);
            }
            Ok(cells)
        })
    }
}

impl CoreWorksheet for SourceBackedWorksheet<'_> {
    fn name(&self) -> &str {
        &self.name
    }

    fn row_count(&self) -> usize {
        self.workbook
            .materialize(self, |worksheet| Ok(worksheet.row_count()))
            .unwrap_or(0)
    }

    fn column_count(&self) -> usize {
        self.workbook
            .materialize(self, |worksheet| Ok(worksheet.column_count()))
            .unwrap_or(0)
    }

    fn dimensions(&self) -> Option<(u32, u32, u32, u32)> {
        self.dimensions_inner().ok().flatten()
    }

    fn cell(&self, row: u32, column: u32) -> SheetResult<Box<dyn CoreCell + '_>> {
        let cell = self.workbook.materialize(self, |worksheet| {
            let cell = worksheet.cell(row.saturating_sub(1), column.saturating_sub(1))?;
            Self::copy_cell(cell.as_ref())
        })?;
        Ok(Box::new(cell))
    }

    fn cell_by_coordinate(&self, coordinate: &str) -> SheetResult<Box<dyn CoreCell + '_>> {
        let cell = self.workbook.materialize(self, |worksheet| {
            let cell = worksheet.cell_by_coordinate(coordinate)?;
            Self::copy_cell(cell.as_ref())
        })?;
        Ok(Box::new(cell))
    }

    fn cell_value(&self, row: u32, column: u32) -> SheetResult<Cow<'_, CellValue>> {
        Ok(Cow::Owned(self.value_at(row, column)?))
    }

    fn cells(&self) -> Box<dyn CellIterator<'_> + '_> {
        match self.stored_cells() {
            Ok(cells) => Box::new(Cells {
                cells: cells.into_iter(),
                error: None,
            }),
            Err(error) => Box::new(Cells {
                cells: Vec::new().into_iter(),
                error: Some(error),
            }),
        }
    }

    fn row(&self, row_idx: usize) -> SheetResult<Cow<'_, [CellValue]>> {
        Ok(Cow::Owned(self.row_values(row_idx)?))
    }

    fn rows(&self) -> Box<dyn RowIterator<'_> + '_> {
        let (end, error) = match self
            .workbook
            .materialize(self, |worksheet| Ok(worksheet.row_count()))
        {
            Ok(end) => (end, None),
            Err(error) => (0, Some(error)),
        };
        Box::new(Rows {
            worksheet: self,
            index: 0,
            end,
            error,
        })
    }
}

struct Cells {
    cells: std::vec::IntoIter<FacadeCell>,
    error: Option<BoxError>,
}

impl<'a> CellIterator<'a> for Cells {
    fn next(&mut self) -> Option<SheetResult<Box<dyn CoreCell + 'a>>> {
        if let Some(error) = self.error.take() {
            return Some(Err(error));
        }
        self.cells
            .next()
            .map(|cell| Ok(Box::new(cell) as Box<dyn CoreCell + 'a>))
    }
}

struct Rows<'a, 'workbook> {
    worksheet: &'a SourceBackedWorksheet<'workbook>,
    index: usize,
    end: usize,
    error: Option<BoxError>,
}

impl<'a, 'workbook> RowIterator<'a> for Rows<'a, 'workbook> {
    fn next(&mut self) -> Option<SheetResult<Cow<'a, [CellValue]>>> {
        if let Some(error) = self.error.take() {
            return Some(Err(error));
        }
        if self.index >= self.end {
            return None;
        }
        let index = self.index;
        self.index += 1;
        Some(self.worksheet.row_values(index).map(Cow::Owned))
    }
}

#[derive(Debug, Clone)]
struct FacadeCell {
    row: u32,
    column: u32,
    value: CellValue,
}

impl CoreCell for FacadeCell {
    fn row(&self) -> u32 {
        self.row
    }

    fn column(&self) -> u32 {
        self.column
    }

    fn coordinate(&self) -> String {
        coordinate(self.row, self.column)
    }

    fn value(&self) -> &CellValue {
        &self.value
    }

    fn is_formula(&self) -> bool {
        matches!(self.value, CellValue::Formula { .. })
    }
}

fn coordinate(row: u32, column: u32) -> String {
    let mut column = column;
    let mut letters = String::new();
    while column > 0 {
        column -= 1;
        let letter = char::from_u32(u32::from(b'A') + (column % 26)).unwrap_or('?');
        letters.push(letter);
        column /= 26;
    }
    let letters: String = letters.chars().rev().collect();
    format!("{letters}{row}")
}

#[cfg(test)]
mod tests {
    use super::Workbook;
    use litchi_core::sheet::{CellValue, WorkbookTrait};
    use litchi_core::{ReadAt, SourceVersion};
    use litchi_opc::ReadLimits;
    use std::io;
    use std::sync::atomic::{AtomicUsize, Ordering};
    use std::sync::{Arc, Barrier};
    use std::thread;

    #[derive(Clone)]
    struct CountingSource {
        bytes: Arc<Vec<u8>>,
        reads: Arc<AtomicUsize>,
    }

    impl CountingSource {
        fn new(bytes: Vec<u8>) -> Self {
            Self {
                bytes: Arc::new(bytes),
                reads: Arc::new(AtomicUsize::new(0)),
            }
        }

        fn reset_reads(&self) {
            self.reads.store(0, Ordering::SeqCst);
        }

        fn read_count(&self) -> usize {
            self.reads.load(Ordering::SeqCst)
        }
    }

    impl ReadAt for CountingSource {
        fn len(&self) -> io::Result<u64> {
            Ok(self.bytes.len() as u64)
        }

        fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
            self.reads.fetch_add(1, Ordering::SeqCst);
            if output.is_empty() {
                return Ok(0);
            }
            let offset = usize::try_from(offset)
                .map_err(|_| io::Error::new(io::ErrorKind::InvalidInput, "offset overflow"))?;
            if offset >= self.bytes.len() {
                return Ok(0);
            }
            let count = output.len().min(self.bytes.len() - offset);
            output[..count].copy_from_slice(&self.bytes[offset..offset + count]);
            Ok(count)
        }

        fn version(&self) -> io::Result<SourceVersion> {
            Ok(SourceVersion::new(0x584c_5342_4641_4345, 0))
        }
    }

    fn fixture() -> Vec<u8> {
        std::fs::read(concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../../test-data/ooxml/xlsb/Simple.xlsb"
        ))
        .unwrap()
    }

    fn adapter(source: &Arc<CountingSource>) -> Workbook {
        let read_at: Arc<dyn ReadAt> = source.clone();
        let workbook =
            crate::xlsb::SourceBackedWorkbook::from_read_at(Arc::clone(&read_at)).unwrap();
        Workbook::from_source_backed(workbook, read_at, ReadLimits::default()).unwrap()
    }

    #[test]
    fn eager_initialization_is_shared_across_barrier_synchronized_calls() {
        let baseline_source = Arc::new(CountingSource::new(fixture()));
        let baseline = adapter(&baseline_source);
        baseline_source.reset_reads();
        let baseline_count = baseline.eager_workbook().unwrap().worksheet_count();
        let baseline_reads = baseline_source.read_count();

        let concurrent_source = Arc::new(CountingSource::new(fixture()));
        let concurrent = Arc::new(adapter(&concurrent_source));
        concurrent_source.reset_reads();
        let barrier = Arc::new(Barrier::new(3));
        let workers = (0..2)
            .map(|_| {
                let workbook = Arc::clone(&concurrent);
                let barrier = Arc::clone(&barrier);
                thread::spawn(move || {
                    barrier.wait();
                    workbook.eager_workbook().unwrap().worksheet_count()
                })
            })
            .collect::<Vec<_>>();
        barrier.wait();
        let results = workers
            .into_iter()
            .map(|worker| worker.join().unwrap())
            .collect::<Vec<_>>();

        assert_eq!(results, vec![baseline_count, baseline_count]);
        assert_eq!(concurrent_source.read_count(), baseline_reads);
    }

    #[test]
    fn ordinary_cell_materialization_is_shared_across_worksheet_handles() {
        let baseline_source = Arc::new(CountingSource::new(fixture()));
        let baseline = adapter(&baseline_source);
        baseline_source.reset_reads();
        let baseline_value = WorkbookTrait::worksheet_by_index(&baseline, 0)
            .unwrap()
            .cell_value(1, 1)
            .unwrap()
            .into_owned();
        let baseline_reads = baseline_source.read_count();

        let concurrent_source = Arc::new(CountingSource::new(fixture()));
        let concurrent = Arc::new(adapter(&concurrent_source));
        concurrent_source.reset_reads();
        let barrier = Arc::new(Barrier::new(3));
        let workers = (0..2)
            .map(|_| {
                let workbook = Arc::clone(&concurrent);
                let barrier = Arc::clone(&barrier);
                thread::spawn(move || {
                    barrier.wait();
                    WorkbookTrait::worksheet_by_index(workbook.as_ref(), 0)
                        .unwrap()
                        .cell_value(1, 1)
                        .unwrap()
                        .into_owned()
                })
            })
            .collect::<Vec<_>>();
        barrier.wait();
        let results = workers
            .into_iter()
            .map(|worker| worker.join().unwrap())
            .collect::<Vec<CellValue>>();

        assert_eq!(results, vec![baseline_value.clone(), baseline_value]);
        assert_eq!(concurrent_source.read_count(), baseline_reads);
    }
}
