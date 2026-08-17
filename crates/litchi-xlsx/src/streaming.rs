//! Restricted, bounded sequential creation of SpreadsheetML workbooks.
//!
//! This module intentionally exposes a narrow authoring surface.  It creates
//! one ordinary worksheet, uses inline scalar values, and keeps the active row
//! in one reusable buffer.  It does not attempt to be a second general
//! workbook model: formulas, shared strings, styles, drawings, protection,
//! macros, encryption, and signing are outside this writer's contract.

use std::io::{self, Write};
use std::sync::{
    Arc, Mutex,
    atomic::{AtomicU64, Ordering},
};

use itoa::Buffer as IntegerBuffer;
use litchi_core::{ExecutionContext, Reservation, Resource};
use litchi_opc::phys_pkg::{PartWriter, PhysPkgWriter};
use litchi_opc::{OpcError, PackURI};
use ryu::Buffer as FloatBuffer;

use crate::cell::ErrorValue;
use crate::error::{Error, Result, invalid};

const MAX_ROWS: u32 = 1_048_576;
const MAX_COLUMNS: u32 = 16_384;
const MAX_CELL_CHARACTERS: u64 = 32_767;
const MIN_OUTPUT_BYTES: u64 = 22;

const CONTENT_TYPES_URI: &str = "/[Content_Types].xml";
const PACKAGE_RELS_URI: &str = "/_rels/.rels";
const WORKBOOK_URI: &str = "/xl/workbook.xml";
const WORKBOOK_RELS_URI: &str = "/xl/_rels/workbook.xml.rels";
const STYLES_URI: &str = "/xl/styles.xml";
const SHEET_URI: &str = "/xl/worksheets/sheet1.xml";

const CONTENT_TYPES_XML: &[u8] = concat!(
    r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#,
    r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">"#,
    r#"<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>"#,
    r#"<Default Extension="xml" ContentType="application/xml"/>"#,
    r#"<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>"#,
    r#"<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>"#,
    r#"<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>"#,
    r#"</Types>"#,
).as_bytes();

const PACKAGE_RELS_XML: &[u8] = concat!(
    r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#,
    r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">"#,
    r#"<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>"#,
    r#"</Relationships>"#,
).as_bytes();

const WORKBOOK_XML: &[u8] = concat!(
    r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#,
    r#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" "#,
    r#"xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">"#,
    r#"<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>"#,
    r#"</workbook>"#,
)
.as_bytes();

const WORKBOOK_RELS_XML: &[u8] = concat!(
    r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#,
    r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">"#,
    r#"<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>"#,
    r#"<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>"#,
    r#"</Relationships>"#,
).as_bytes();

const STYLES_XML: &[u8] = concat!(
    r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#,
    r#"<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">"#,
    r#"<fonts count="1"><font/></fonts>"#,
    r#"<fills count="2"><fill><patternFill patternType="none"/></fill>"#,
    r#"<fill><patternFill patternType="gray125"/></fill></fills>"#,
    r#"<borders count="1"><border/></borders>"#,
    r#"<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>"#,
    r#"<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>"#,
    r#"<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>"#,
    r#"</styleSheet>"#,
).as_bytes();

const SHEET_PREFIX: &[u8] = concat!(
    r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#,
    r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>"#,
)
.as_bytes();
const SHEET_SUFFIX: &[u8] = b"</sheetData></worksheet>";

/// Finite semantic limits for [`StreamingWorkbookWriter`].
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct StreamingWorkbookLimits {
    /// Largest accepted worksheet row number.
    pub max_row: u32,
    /// Largest number of cells accepted over the complete worksheet.
    pub max_cells: u64,
    /// Largest UTF-8 byte length of one text value.
    pub max_cell_text_bytes: u64,
    /// Capacity of the reusable row scratch buffer.
    pub max_row_bytes: u64,
    /// Largest complete `sheet1.xml` payload, including its XML envelope.
    pub max_sheet_xml_bytes: u64,
    /// Largest number of bytes accepted by the complete output sink.
    pub max_output_bytes: u64,
}

impl StreamingWorkbookLimits {
    /// Creates explicit finite semantic limits.
    #[must_use]
    pub const fn new(
        max_row: u32,
        max_cells: u64,
        max_cell_text_bytes: u64,
        max_row_bytes: u64,
        max_sheet_xml_bytes: u64,
        max_output_bytes: u64,
    ) -> Self {
        Self {
            max_row,
            max_cells,
            max_cell_text_bytes,
            max_row_bytes,
            max_sheet_xml_bytes,
            max_output_bytes,
        }
    }
}

impl Default for StreamingWorkbookLimits {
    fn default() -> Self {
        Self {
            max_row: MAX_ROWS,
            max_cells: 10_000_000,
            max_cell_text_bytes: 131_068,
            max_row_bytes: 1024 * 1024,
            max_sheet_xml_bytes: 256 * 1024 * 1024,
            max_output_bytes: 256 * 1024 * 1024,
        }
    }
}

/// A scalar value accepted by the bounded streaming XLSX authoring API.
#[derive(Debug, Clone)]
#[non_exhaustive]
pub enum StreamingCellValue<'a> {
    /// Inline UTF-8 text. Shared strings are never created.
    Text(&'a str),
    /// A SpreadsheetML boolean (`1` or `0`).
    Bool(bool),
    /// A finite IEEE-754 number rendered by the deterministic `ryu` formatter.
    Number(f64),
    /// One of the recognized SpreadsheetML error values.
    Error(ErrorValue),
    /// An explicit blank cell record.
    Blank,
}

/// One worksheet cell supplied to [`StreamingWorkbookWriter::write_row`].
#[derive(Debug, Clone)]
pub struct StreamingCell<'a> {
    /// One-based Excel column number (`A` is `1`, `XFD` is `16_384`).
    pub column: u32,
    /// Scalar payload written for this cell.
    pub value: StreamingCellValue<'a>,
}

impl<'a> StreamingCell<'a> {
    /// Construct one scalar cell.
    #[must_use]
    pub const fn new(column: u32, value: StreamingCellValue<'a>) -> Self {
        Self { column, value }
    }
}

/// A sequential, bounded writer for one ordinary XLSX worksheet.
///
/// The writer owns the caller's non-seek sink.  Static package members are
/// published first; `xl/worksheets/sheet1.xml` is then opened as the final
/// member and accepts rows in strictly increasing order.  The active row is
/// assembled into one bounded reusable scratch buffer before it is published,
/// so an invalid cell or ordering error cannot leave a half-written row.
/// The sheet name is fixed to `Sheet1` in this v1 API; atomic filesystem path
/// replacement is likewise the caller's responsibility because this writer
/// accepts only an owned `Write` sink.
/// The execution context charges one writer object and six finalized package
/// parts, plus one row and one object per emitted cell.  A nonempty row's
/// object charge is reserved until its bytes are accepted.  Work is charged
/// for package setup, every row attempt, and every cell encoding; rejected row
/// attempts therefore still consume Work but do not consume row Objects.
/// Dropping the writer before [`Self::finish`] abandons the package and can
/// leave the owned sink with an intentionally incomplete archive.
pub struct StreamingWorkbookWriter<W: Write> {
    part: Option<PartWriter<BudgetedOutput<W>>>,
    context: ExecutionContext,
    execution_state: Arc<ExecutionFailureState>,
    limits: StreamingWorkbookLimits,
    row_scratch: Vec<u8>,
    sheet_xml_bytes: u64,
    _row_reservation: Reservation,
    last_row: Option<u32>,
    cells: u64,
    output_counter: Arc<AtomicU64>,
    poisoned: bool,
    poison_message: Option<String>,
}

impl<W: Write> StreamingWorkbookWriter<W> {
    /// Create a writer and publish the five static package members.
    ///
    /// The sink is never required to implement `Seek`.  The explicit context
    /// is checked and charged before any output is attempted.
    pub fn new(
        writer: W,
        context: ExecutionContext,
        limits: StreamingWorkbookLimits,
    ) -> Result<Self> {
        validate_limits(&limits)?;
        context.check().map_err(map_execution)?;
        let scratch_bytes = usize::try_from(limits.max_row_bytes)
            .map_err(|_| invalid("streaming XLSX row scratch exceeds usize"))?;
        let row_reservation = context
            .reserve(Resource::Memory, limits.max_row_bytes)
            .map_err(map_execution)?;
        let mut row_scratch = Vec::new();
        row_scratch
            .try_reserve_exact(scratch_bytes)
            .map_err(|error| crate::error::allocation("streaming XLSX row scratch", error))?;
        context
            .consume(Resource::Objects, 1)
            .and_then(|_| context.consume(Resource::Work, 1))
            .map_err(map_execution)?;

        let output_counter = Arc::new(AtomicU64::new(0));
        let execution_state = Arc::new(ExecutionFailureState::default());
        let output = BudgetedOutput::new(
            writer,
            context.clone(),
            limits.max_output_bytes,
            Arc::clone(&output_counter),
            Arc::clone(&execution_state),
        );
        let mut physical = PhysPkgWriter::with_writer(output);
        for (name, bytes) in [
            (CONTENT_TYPES_URI, CONTENT_TYPES_XML),
            (PACKAGE_RELS_URI, PACKAGE_RELS_XML),
            (WORKBOOK_URI, WORKBOOK_XML),
            (WORKBOOK_RELS_URI, WORKBOOK_RELS_XML),
            (STYLES_URI, STYLES_XML),
        ] {
            let uri = PackURI::new(name).map_err(|error| invalid(error.to_string()))?;
            let mut member = physical.start_part(&uri).map_err(|error| {
                package_error_with_progress_or_execution(
                    error,
                    output_counter.load(Ordering::Acquire),
                    &execution_state,
                )
            })?;
            if let Err(write_error) = member.write_all(bytes) {
                let finish_error = member.finish().err();
                return Err(map_member_write_error(
                    write_error,
                    finish_error,
                    output_counter.load(Ordering::Acquire),
                    &execution_state,
                ));
            }
            physical = member.finish().map_err(|error| {
                package_error_with_progress_or_execution(
                    error,
                    output_counter.load(Ordering::Acquire),
                    &execution_state,
                )
            })?;
            context.consume(Resource::Objects, 1).map_err(|error| {
                execution_failure(error, output_counter.load(Ordering::Acquire))
            })?;
        }
        let sheet_uri = PackURI::new(SHEET_URI).map_err(|error| invalid(error.to_string()))?;
        let mut part = physical.start_part(&sheet_uri).map_err(|error| {
            package_error_with_progress_or_execution(
                error,
                output_counter.load(Ordering::Acquire),
                &execution_state,
            )
        })?;
        if let Err(write_error) = part.write_all(SHEET_PREFIX) {
            let finish_error = part.finish().err();
            return Err(map_member_write_error(
                write_error,
                finish_error,
                output_counter.load(Ordering::Acquire),
                &execution_state,
            ));
        }

        Ok(Self {
            part: Some(part),
            context,
            execution_state,
            limits,
            row_scratch,
            sheet_xml_bytes: u64::try_from(SHEET_PREFIX.len()).unwrap_or(u64::MAX),
            _row_reservation: row_reservation,
            last_row: None,
            cells: 0,
            output_counter,
            poisoned: false,
            poison_message: None,
        })
    }

    /// Number of accepted worksheet cells.
    #[must_use]
    pub const fn cell_count(&self) -> u64 {
        self.cells
    }

    /// Number of worksheet XML bytes already committed, excluding the final
    /// closing envelope.
    #[must_use]
    pub const fn worksheet_xml_bytes(&self) -> u64 {
        self.sheet_xml_bytes
    }

    /// Semantic limits retained by this writer.
    #[must_use]
    pub const fn limits(&self) -> StreamingWorkbookLimits {
        self.limits
    }

    /// Number of bytes accepted by the owned output sink so far.
    #[must_use]
    pub fn output_bytes(&self) -> u64 {
        self.output_counter.load(Ordering::Acquire)
    }

    /// Whether a sink, cancellation, or finalization failure permanently
    /// invalidated this writer.
    #[must_use]
    pub const fn is_poisoned(&self) -> bool {
        self.poisoned
    }

    /// Write one sparse worksheet row.
    ///
    /// Rows may be skipped.  An empty iterator is a deliberate no-op and is
    /// still considered a consumed row position; this makes dropping an empty
    /// row deterministic and prevents a later call from going backwards.
    pub fn write_row<'a, I>(&mut self, row: u32, cells: I) -> Result<()>
    where
        I: IntoIterator<Item = StreamingCell<'a>>,
    {
        self.ensure_usable()?;
        self.context
            .consume(Resource::Work, 1)
            .map_err(|error| self.poison_execution(error))?;
        if row == 0 || row > self.limits.max_row || row > MAX_ROWS {
            return Err(invalid(format!(
                "streaming XLSX row {row} is outside the configured domain"
            )));
        }
        if self.last_row.is_some_and(|last| row <= last) {
            return Err(invalid(format!(
                "streaming XLSX rows must strictly increase: {row}"
            )));
        }
        self.context
            .check()
            .map_err(|error| self.poison_execution(error))?;
        self.row_scratch.clear();
        let mut row_number = IntegerBuffer::new();
        let row_number = row_number.format(row).as_bytes();
        push_bytes(&mut self.row_scratch, self.limits.max_row_bytes, b"<row r=")?;
        push_quoted_bytes(&mut self.row_scratch, self.limits.max_row_bytes, row_number)?;
        push_bytes(&mut self.row_scratch, self.limits.max_row_bytes, b">")?;

        let mut last_column = None;
        let mut row_cells = 0_u64;
        for cell in cells {
            self.context
                .check()
                .map_err(|error| self.poison_execution(error))?;
            if cell.column == 0 || cell.column > MAX_COLUMNS {
                return Err(invalid(format!(
                    "streaming XLSX column {} is outside the configured domain",
                    cell.column
                )));
            }
            if last_column.is_some_and(|last| cell.column <= last) {
                return Err(invalid(format!(
                    "streaming XLSX cells must strictly increase in row {row}: column {}",
                    cell.column
                )));
            }
            if self.cells.saturating_add(row_cells).saturating_add(1) > self.limits.max_cells {
                return Err(invalid("streaming XLSX cell limit exceeded"));
            }
            self.context
                .consume(Resource::Work, 1)
                .map_err(|error| self.poison_execution(error))?;
            append_cell(
                &mut self.row_scratch,
                self.limits.max_row_bytes,
                row_number,
                &cell,
                self.limits.max_cell_text_bytes,
            )?;
            last_column = Some(cell.column);
            row_cells += 1;
        }
        push_bytes(&mut self.row_scratch, self.limits.max_row_bytes, b"</row>")?;
        let row_bytes = u64::try_from(self.row_scratch.len()).unwrap_or(u64::MAX);
        if row_cells == 0 {
            self.last_row = Some(row);
            return Ok(());
        }
        let object_reservation = self
            .context
            .reserve(Resource::Objects, row_cells.saturating_add(1))
            .map_err(|error| self.poison_execution(error))?;
        {
            let next_sheet_bytes = self
                .sheet_xml_bytes
                .saturating_add(row_bytes)
                .saturating_add(u64::try_from(SHEET_SUFFIX.len()).unwrap_or(u64::MAX));
            if next_sheet_bytes > self.limits.max_sheet_xml_bytes {
                return Err(invalid("streaming XLSX worksheet XML limit exceeded"));
            }
            let part = self
                .part
                .as_mut()
                .ok_or_else(|| invalid("streaming XLSX part is unavailable"))?;
            if let Err(error) = part.write_all(&self.row_scratch) {
                return Err(self.poison_io(error));
            }
            if !object_reservation.commit(row_cells.saturating_add(1)) {
                return Err(invalid("streaming XLSX row object reservation underflow"));
            }
            self.cells = self.cells.saturating_add(row_cells);
            self.sheet_xml_bytes = self.sheet_xml_bytes.saturating_add(row_bytes);
        }
        self.last_row = Some(row);
        Ok(())
    }

    /// Finalize the worksheet, central directory, and caller-owned sink.
    pub fn finish(mut self) -> Result<W> {
        if self.poisoned {
            return Err(self.poison_error());
        }
        if let Err(error) = self.context.check() {
            return Err(self.poison_execution(error));
        }
        let Some(mut part) = self.part.take() else {
            return Err(invalid("streaming XLSX part is unavailable"));
        };
        if let Err(error) = part.write_all(SHEET_SUFFIX) {
            self.poisoned = true;
            let output = self.output_bytes();
            let mapped = if let Some(execution) = execution_from_io_error(&error) {
                execution_failure(execution, output)
            } else {
                match part.finish() {
                    Err(finish_error) => package_error_with_progress_or_execution(
                        finish_error,
                        output,
                        &self.execution_state,
                    ),
                    Ok(_) => incomplete_io_error_or_execution(error, output, &self.execution_state),
                }
            };
            self.poison_message = Some(mapped.to_string());
            return Err(mapped);
        }
        let physical = match part.finish() {
            Ok(physical) => physical,
            Err(error) => {
                self.poisoned = true;
                let mapped = package_error_with_progress_or_execution(
                    error,
                    self.output_bytes(),
                    &self.execution_state,
                );
                self.poison_message = Some(mapped.to_string());
                return Err(mapped);
            },
        };
        if let Err(error) = self.context.consume(Resource::Objects, 1) {
            self.poisoned = true;
            let mapped = execution_failure(error, self.output_bytes());
            self.poison_message = Some(mapped.to_string());
            return Err(mapped);
        }
        let output = match physical.finish_into_inner() {
            Ok(output) => output,
            Err(error) => {
                self.poisoned = true;
                let mapped = package_error_with_progress_or_execution(
                    error,
                    self.output_bytes(),
                    &self.execution_state,
                );
                self.poison_message = Some(mapped.to_string());
                return Err(mapped);
            },
        };
        output.into_inner().map_err(|error| {
            incomplete_io_error_or_execution(error, self.output_bytes(), &self.execution_state)
        })
    }

    fn ensure_usable(&self) -> Result<()> {
        if self.poisoned {
            return Err(self.poison_error());
        }
        Ok(())
    }

    fn poison_execution(&mut self, error: litchi_core::ExecutionError) -> Error {
        self.poisoned = true;
        self.poison_message = Some(error.to_string());
        execution_failure(error, self.output_bytes())
    }

    fn poison_io(&mut self, error: io::Error) -> Error {
        self.poisoned = true;
        let output = self.output_bytes();
        if let Some(execution) = execution_from_io_error(&error) {
            self.part.take();
            let mapped = execution_failure(execution, output);
            self.poison_message = Some(mapped.to_string());
            return mapped;
        }
        let mapped = if let Some(part) = self.part.take() {
            match part.finish() {
                Err(error) => {
                    package_error_with_progress_or_execution(error, output, &self.execution_state)
                },
                Ok(_) => incomplete_io_error_or_execution(error, output, &self.execution_state),
            }
        } else {
            incomplete_io_error_or_execution(error, output, &self.execution_state)
        };
        self.poison_message = Some(format!("{mapped}; accepted output bytes: {output}"));
        mapped
    }

    fn poison_error(&self) -> Error {
        invalid(
            self.poison_message
                .as_deref()
                .unwrap_or("streaming XLSX writer is poisoned"),
        )
    }
}

/// Convenience alias for the scalar value enum.
pub type StreamingValue<'a> = StreamingCellValue<'a>;

#[derive(Default)]
struct ExecutionFailureState {
    error: Mutex<Option<litchi_core::ExecutionError>>,
}

impl ExecutionFailureState {
    fn record(&self, error: litchi_core::ExecutionError) {
        let mut recorded = self
            .error
            .lock()
            .unwrap_or_else(|poisoned| poisoned.into_inner());
        if recorded.is_none() {
            *recorded = Some(error);
        }
    }

    fn get(&self) -> Option<litchi_core::ExecutionError> {
        self.error
            .lock()
            .unwrap_or_else(|poisoned| poisoned.into_inner())
            .clone()
    }
}

struct BudgetedOutput<W> {
    writer: W,
    context: ExecutionContext,
    maximum: u64,
    accepted: u64,
    counter: Arc<AtomicU64>,
    execution_state: Arc<ExecutionFailureState>,
}

impl<W> BudgetedOutput<W> {
    fn new(
        writer: W,
        context: ExecutionContext,
        maximum: u64,
        counter: Arc<AtomicU64>,
        execution_state: Arc<ExecutionFailureState>,
    ) -> Self {
        Self {
            writer,
            context,
            maximum,
            accepted: 0,
            counter,
            execution_state,
        }
    }

    fn into_inner(self) -> io::Result<W> {
        Ok(self.writer)
    }
}

impl<W: Write> Write for BudgetedOutput<W> {
    fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
        if bytes.is_empty() {
            return Ok(0);
        }
        if let Err(error) = self.context.check() {
            self.execution_state.record(error.clone());
            return Err(execution_io_error(error));
        }
        let remaining = self.maximum.saturating_sub(self.accepted);
        if remaining == 0 {
            return Err(io::Error::other(format!(
                "streaming XLSX output limit exceeded: {} >= {}",
                self.accepted, self.maximum
            )));
        }
        let requested = u64::try_from(bytes.len())
            .unwrap_or(u64::MAX)
            .min(remaining);
        let request_len = usize::try_from(requested).unwrap_or(bytes.len());
        let reservation = self
            .context
            .reserve(Resource::OutputBytes, requested)
            .map_err(|error| {
                self.execution_state.record(error.clone());
                execution_io_error(error)
            })?;
        let written = match self.writer.write(&bytes[..request_len]) {
            Ok(written) => written,
            Err(error) => {
                drop(reservation);
                return Err(error);
            },
        };
        let written_u64 = u64::try_from(written).unwrap_or(u64::MAX);
        if !reservation.commit(written_u64) {
            return Err(io::Error::other(
                "streaming XLSX sink returned more bytes than requested",
            ));
        }
        self.accepted = self.accepted.saturating_add(written_u64);
        self.counter.store(self.accepted, Ordering::Release);
        Ok(written)
    }

    fn flush(&mut self) -> io::Result<()> {
        if let Err(error) = self.context.check() {
            self.execution_state.record(error.clone());
            return Err(execution_io_error(error));
        }
        let result = self.writer.flush();
        if result.is_ok() {
            if let Err(error) = self.context.check() {
                self.execution_state.record(error.clone());
                return Err(execution_io_error(error));
            }
        }
        result
    }
}

fn validate_limits(limits: &StreamingWorkbookLimits) -> Result<()> {
    if limits.max_row == 0 || limits.max_row > MAX_ROWS {
        return Err(invalid(
            "streaming XLSX max_row is outside the Excel row domain",
        ));
    }
    if limits.max_cells == 0 {
        return Err(invalid("streaming XLSX max_cells must be positive"));
    }
    if limits.max_cell_text_bytes == 0 {
        return Err(invalid(
            "streaming XLSX max_cell_text_bytes must be positive",
        ));
    }
    if limits.max_row_bytes < 64 {
        return Err(invalid("streaming XLSX max_row_bytes is too small"));
    }
    let minimum_sheet_bytes =
        u64::try_from(SHEET_PREFIX.len() + SHEET_SUFFIX.len()).unwrap_or(u64::MAX);
    if limits.max_sheet_xml_bytes < minimum_sheet_bytes {
        return Err(invalid("streaming XLSX max_sheet_xml_bytes is too small"));
    }
    if limits.max_output_bytes < MIN_OUTPUT_BYTES {
        return Err(invalid(
            "streaming XLSX max_output_bytes is too small for ZIP",
        ));
    }
    Ok(())
}

fn map_execution(error: litchi_core::ExecutionError) -> Error {
    Error::Package(OpcError::Execution(error))
}

#[derive(Debug)]
struct ExecutionMarker(litchi_core::ExecutionError);

impl std::fmt::Display for ExecutionMarker {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        self.0.fmt(formatter)
    }
}

impl std::error::Error for ExecutionMarker {}

fn execution_io_error(error: litchi_core::ExecutionError) -> io::Error {
    io::Error::other(ExecutionMarker(error))
}

fn execution_from_io_error(error: &io::Error) -> Option<litchi_core::ExecutionError> {
    error
        .get_ref()
        .and_then(|source| source.downcast_ref::<ExecutionMarker>())
        .map(|marker| marker.0.clone())
}

fn execution_failure(error: litchi_core::ExecutionError, output: u64) -> Error {
    let source = OpcError::Execution(error);
    if output == 0 {
        Error::Package(source)
    } else {
        Error::Package(OpcError::IncompleteOutput {
            written: output,
            source: Box::new(source),
        })
    }
}

fn package_error_with_progress(error: OpcError, output: u64) -> Error {
    if output == 0 || matches!(&error, OpcError::IncompleteOutput { .. }) {
        Error::Package(error)
    } else {
        Error::Package(OpcError::IncompleteOutput {
            written: output,
            source: Box::new(error),
        })
    }
}

fn package_error_with_progress_or_execution(
    error: OpcError,
    output: u64,
    execution_state: &ExecutionFailureState,
) -> Error {
    execution_state.get().map_or_else(
        || package_error_with_progress(error, output),
        |error| execution_failure(error, output),
    )
}

fn map_member_write_error(
    write_error: io::Error,
    finish_error: Option<OpcError>,
    output: u64,
    execution_state: &ExecutionFailureState,
) -> Error {
    if let Some(execution) = execution_from_io_error(&write_error) {
        return execution_failure(execution, output);
    }
    if let Some(execution) = execution_state.get() {
        return execution_failure(execution, output);
    }
    finish_error.map_or_else(
        || incomplete_io_error(write_error, output),
        |error| package_error_with_progress_or_execution(error, output, execution_state),
    )
}

fn incomplete_io_error(error: io::Error, output: u64) -> Error {
    let source = OpcError::IoError(error);
    if output == 0 {
        Error::Package(source)
    } else {
        Error::Package(OpcError::IncompleteOutput {
            written: output,
            source: Box::new(source),
        })
    }
}

fn incomplete_io_error_or_execution(
    error: io::Error,
    output: u64,
    execution_state: &ExecutionFailureState,
) -> Error {
    execution_state.get().map_or_else(
        || incomplete_io_error(error, output),
        |error| execution_failure(error, output),
    )
}

fn push_bytes(buffer: &mut Vec<u8>, maximum: u64, bytes: &[u8]) -> Result<()> {
    let next = u64::try_from(buffer.len())
        .unwrap_or(u64::MAX)
        .saturating_add(u64::try_from(bytes.len()).unwrap_or(u64::MAX));
    if next > maximum {
        return Err(invalid("streaming XLSX row scratch limit exceeded"));
    }
    buffer.extend_from_slice(bytes);
    Ok(())
}

fn push_quoted_bytes(buffer: &mut Vec<u8>, maximum: u64, value: &[u8]) -> Result<()> {
    push_bytes(buffer, maximum, b"\"")?;
    push_bytes(buffer, maximum, value)?;
    push_bytes(buffer, maximum, b"\"")
}

fn append_cell(
    buffer: &mut Vec<u8>,
    maximum: u64,
    row_number: &[u8],
    cell: &StreamingCell<'_>,
    max_text_bytes: u64,
) -> Result<()> {
    push_bytes(buffer, maximum, b"<c r=\"")?;
    append_column(buffer, maximum, cell.column)?;
    push_bytes(buffer, maximum, row_number)?;
    match &cell.value {
        StreamingCellValue::Blank => push_bytes(buffer, maximum, b"\"/>")?,
        StreamingCellValue::Bool(value) => {
            push_bytes(buffer, maximum, b"\" t=\"b\"><v>")?;
            push_bytes(buffer, maximum, if *value { b"1" } else { b"0" })?;
            push_bytes(buffer, maximum, b"</v></c>")?;
        },
        StreamingCellValue::Number(value) => {
            if !value.is_finite() {
                return Err(invalid("streaming XLSX numbers must be finite"));
            }
            push_bytes(buffer, maximum, b"\"><v>")?;
            let mut number = FloatBuffer::new();
            push_bytes(buffer, maximum, number.format(*value).as_bytes())?;
            push_bytes(buffer, maximum, b"</v></c>")?;
        },
        StreamingCellValue::Error(value) => {
            if matches!(value, ErrorValue::Unknown(_)) {
                return Err(invalid("streaming XLSX error value is not recognized"));
            }
            push_bytes(buffer, maximum, b"\" t=\"e\"><v>")?;
            push_bytes(buffer, maximum, value.as_str().as_bytes())?;
            push_bytes(buffer, maximum, b"</v></c>")?;
        },
        StreamingCellValue::Text(value) => {
            let text_bytes = u64::try_from(value.len()).unwrap_or(u64::MAX);
            if text_bytes > max_text_bytes
                || (text_bytes > MAX_CELL_CHARACTERS
                    && u64::try_from(value.chars().count()).unwrap_or(u64::MAX)
                        > MAX_CELL_CHARACTERS)
            {
                return Err(invalid(
                    "streaming XLSX text value exceeds its finite limit",
                ));
            }
            push_bytes(
                buffer,
                maximum,
                b"\" t=\"inlineStr\"><is><t xml:space=\"preserve\">",
            )?;
            push_escaped(buffer, maximum, value)?;
            push_bytes(buffer, maximum, b"</t></is></c>")?;
        },
    }
    Ok(())
}

fn append_column(buffer: &mut Vec<u8>, maximum: u64, mut column: u32) -> Result<()> {
    let mut letters = [0_u8; 3];
    let mut index = letters.len();
    while column != 0 {
        index -= 1;
        letters[index] = b'A' + ((column - 1) % 26) as u8;
        column = (column - 1) / 26;
    }
    push_bytes(buffer, maximum, &letters[index..])
}

fn push_escaped(buffer: &mut Vec<u8>, maximum: u64, value: &str) -> Result<()> {
    let mut run_start = 0;
    for (offset, character) in value.char_indices() {
        if !is_xml_character(character) {
            if run_start < offset {
                push_bytes(buffer, maximum, &value.as_bytes()[run_start..offset])?;
            }
            return Err(invalid(format!(
                "streaming XLSX text contains XML-invalid character U+{:04X}",
                character as u32
            )));
        }
        let escaped = match character {
            '&' => "&amp;",
            '<' => "&lt;",
            '>' => "&gt;",
            '"' => "&quot;",
            '\'' => "&apos;",
            _ => continue,
        };
        if run_start < offset {
            push_bytes(buffer, maximum, &value.as_bytes()[run_start..offset])?;
        }
        push_bytes(buffer, maximum, escaped.as_bytes())?;
        run_start = offset.saturating_add(character.len_utf8());
    }
    if run_start < value.len() {
        push_bytes(buffer, maximum, &value.as_bytes()[run_start..])?;
    }
    Ok(())
}

const fn is_xml_character(character: char) -> bool {
    let value = character as u32;
    matches!(value, 0x9 | 0xA | 0xD | 0x20..=0xD7FF | 0xE000..=0xFFFD | 0x10000..=0x10FFFF)
}

#[cfg(test)]
mod tests {
    use super::*;
    use litchi_core::{Budget, CancellationSource, ExecutionLimits, Limits};
    use std::num::{NonZeroU64, NonZeroUsize};
    use std::sync::atomic::{AtomicBool, AtomicUsize, Ordering as AtomicOrdering};

    fn context(output: u64) -> ExecutionContext {
        context_pair(output).1
    }

    fn context_with_objects_and_work(objects: u64, work: u64) -> ExecutionContext {
        let budget = Budget::root(
            "streaming-xlsx-test",
            Limits::new(
                8 * 1024 * 1024,
                u64::MAX,
                16 * 1024 * 1024,
                objects,
                u64::MAX,
                work,
            ),
        );
        let (_source, token) = CancellationSource::pair();
        let limits = ExecutionLimits::new(
            NonZeroUsize::new(1).unwrap(),
            NonZeroUsize::new(1).unwrap(),
            NonZeroU64::new(8 * 1024 * 1024).unwrap(),
            1,
        )
        .unwrap();
        ExecutionContext::new(budget, token, limits)
    }

    fn context_pair(output: u64) -> (CancellationSource, ExecutionContext) {
        let budget = Budget::root(
            "streaming-xlsx-test",
            Limits::new(
                8 * 1024 * 1024,
                u64::MAX,
                output,
                10_000,
                u64::MAX,
                10_000_000,
            ),
        );
        let (source, token) = CancellationSource::pair();
        let limits = ExecutionLimits::new(
            NonZeroUsize::new(1).unwrap(),
            NonZeroUsize::new(1).unwrap(),
            NonZeroU64::new(8 * 1024 * 1024).unwrap(),
            1,
        )
        .unwrap();
        (source, ExecutionContext::new(budget, token, limits))
    }

    fn writer() -> StreamingWorkbookWriter<Vec<u8>> {
        StreamingWorkbookWriter::new(
            Vec::new(),
            context(16 * 1024 * 1024),
            StreamingWorkbookLimits::default(),
        )
        .unwrap()
    }

    #[test]
    fn scalar_rows_reopen_without_dimension_or_shared_strings() {
        let mut value_writer = writer();
        value_writer
            .write_row(
                1,
                [
                    StreamingCell::new(1, StreamingCellValue::Text("a&<")),
                    StreamingCell::new(2, StreamingCellValue::Bool(true)),
                    StreamingCell::new(3, StreamingCellValue::Number(1.25)),
                    StreamingCell::new(4, StreamingCellValue::Error(ErrorValue::DivZero)),
                    StreamingCell::new(5, StreamingCellValue::Blank),
                ],
            )
            .unwrap();
        let bytes = value_writer.finish().unwrap();
        let package = litchi_opc::phys_pkg::PhysPkgReader::new(&bytes).unwrap();
        assert_eq!(
            package.member_names().unwrap(),
            vec![
                "[Content_Types].xml",
                "_rels/.rels",
                "xl/workbook.xml",
                "xl/_rels/workbook.xml.rels",
                "xl/styles.xml",
                "xl/worksheets/sheet1.xml",
            ]
        );
        assert!(
            !bytes
                .windows(b"sharedStrings".len())
                .any(|window| window == b"sharedStrings")
        );
        assert!(
            !bytes
                .windows(b"dimension".len())
                .any(|window| window == b"dimension")
        );
        let workbook = crate::Workbook::from_bytes(bytes).unwrap();
        let sheet = workbook.sheet("Sheet1").unwrap().unwrap();
        assert!(matches!(
            sheet.cell("A1").unwrap().stored(),
            Some(crate::Cell::Value(crate::Value::Text(text))) if text.as_str() == "a&<"
        ));
    }

    #[test]
    fn empty_workbook_has_compact_canonical_six_part_topology() {
        let bytes = writer().finish().unwrap();
        let package = litchi_opc::phys_pkg::PhysPkgReader::new(&bytes).unwrap();
        assert_eq!(package.len(), 6);
        for (uri, expected) in [
            (CONTENT_TYPES_URI, CONTENT_TYPES_XML),
            (PACKAGE_RELS_URI, PACKAGE_RELS_XML),
            (WORKBOOK_URI, WORKBOOK_XML),
            (WORKBOOK_RELS_URI, WORKBOOK_RELS_XML),
            (STYLES_URI, STYLES_XML),
        ] {
            let uri = PackURI::new(uri).unwrap();
            assert_eq!(package.blob_for(&uri).unwrap(), expected);
        }
        let mut expected_sheet = Vec::with_capacity(SHEET_PREFIX.len() + SHEET_SUFFIX.len());
        expected_sheet.extend_from_slice(SHEET_PREFIX);
        expected_sheet.extend_from_slice(SHEET_SUFFIX);
        let sheet_uri = PackURI::new(SHEET_URI).unwrap();
        assert_eq!(package.blob_for(&sheet_uri).unwrap(), expected_sheet);
        assert!(
            !bytes
                .windows(b"dimension".len())
                .any(|window| window == b"dimension")
        );
        assert!(
            !bytes
                .windows(b"sharedStrings".len())
                .any(|window| window == b"sharedStrings")
        );
        let workbook = crate::Workbook::from_bytes(bytes).unwrap();
        let sheet = workbook.sheet("Sheet1").unwrap().unwrap();
        assert!(sheet.cell("A1").unwrap().is_missing());
    }

    #[test]
    fn rows_and_cells_are_strict_and_empty_rows_are_dropped() {
        let mut writer = writer();
        writer.write_row(2, std::iter::empty()).unwrap();
        assert!(writer.write_row(2, std::iter::empty()).is_err());
        assert!(
            writer
                .write_row(
                    3,
                    [
                        StreamingCell::new(2, StreamingCellValue::Blank),
                        StreamingCell::new(1, StreamingCellValue::Blank)
                    ]
                )
                .is_err()
        );
        writer
            .write_row(4, [StreamingCell::new(2, StreamingCellValue::Blank)])
            .unwrap();
        let bytes = writer.finish().unwrap();
        assert!(
            !bytes
                .windows(b"r=\"2\"".len())
                .any(|window| window == b"r=\"2\"")
        );
    }

    #[test]
    fn text_escaping_rejects_xml_controls() {
        let mut writer = writer();
        assert!(
            writer
                .write_row(
                    1,
                    [StreamingCell::new(1, StreamingCellValue::Text("bad\u{1}"))]
                )
                .is_err()
        );
    }

    fn push_escaped_scalar_reference(
        buffer: &mut Vec<u8>,
        maximum: u64,
        value: &str,
    ) -> Result<()> {
        for character in value.chars() {
            if !is_xml_character(character) {
                return Err(invalid(format!(
                    "streaming XLSX text contains XML-invalid character U+{:04X}",
                    character as u32
                )));
            }
            let escaped = match character {
                '&' => "&amp;",
                '<' => "&lt;",
                '>' => "&gt;",
                '"' => "&quot;",
                '\'' => "&apos;",
                character => {
                    let mut bytes = [0_u8; 4];
                    push_bytes(
                        buffer,
                        maximum,
                        character.encode_utf8(&mut bytes).as_bytes(),
                    )?;
                    continue;
                },
            };
            push_bytes(buffer, maximum, escaped.as_bytes())?;
        }
        Ok(())
    }

    #[test]
    fn batched_text_escaping_matches_scalar_error_order_and_bytes() {
        let values = [
            "",
            "ordinary ASCII run",
            "&<>\"'",
            "before&between<after",
            "\t\n\r",
            "café 東京 🍋",
            "ordinary run before\u{1}invalid",
            "escaped&ampersand then\u{1}invalid",
            "\u{1}invalid first",
        ];
        for value in values {
            for maximum in 0..=128 {
                let mut scalar = b"prefix".to_vec();
                let mut batched = scalar.clone();
                let scalar_result = push_escaped_scalar_reference(&mut scalar, maximum, value)
                    .map_err(|error| error.to_string());
                let batched_result =
                    push_escaped(&mut batched, maximum, value).map_err(|error| error.to_string());
                assert_eq!(
                    batched_result, scalar_result,
                    "value={value:?}, max={maximum}"
                );
                if batched_result.is_ok() {
                    assert_eq!(batched, scalar, "value={value:?}, max={maximum}");
                }
            }
        }
    }

    #[test]
    fn rejected_text_row_can_be_retried_with_exact_scalar_values() {
        let mut writer = writer();
        assert!(
            writer
                .write_row(
                    1,
                    [StreamingCell::new(
                        1,
                        StreamingCellValue::Text("ordinary run before\u{1}invalid"),
                    )],
                )
                .is_err()
        );
        writer
            .write_row(
                1,
                [
                    StreamingCell::new(1, StreamingCellValue::Text("&<>\"' café 東京 🍋")),
                    StreamingCell::new(26, StreamingCellValue::Number(-0.0)),
                    StreamingCell::new(27, StreamingCellValue::Bool(false)),
                    StreamingCell::new(702, StreamingCellValue::Error(ErrorValue::Null)),
                    StreamingCell::new(703, StreamingCellValue::Blank),
                ],
            )
            .unwrap();
        let bytes = writer.finish().unwrap();
        let workbook = crate::Workbook::from_bytes(bytes).unwrap();
        let sheet = workbook.sheet("Sheet1").unwrap().unwrap();
        assert!(matches!(
            sheet.cell("A1").unwrap().stored(),
            Some(crate::Cell::Value(crate::Value::Text(text)))
                if text.as_str() == "&<>\"' café 東京 🍋"
        ));
        assert!(matches!(
            sheet.cell("Z1").unwrap().stored(),
            Some(crate::Cell::Value(crate::Value::Number(value)))
                if value.as_f64() == Some(-0.0)
        ));
        assert!(matches!(
            sheet.cell("AA1").unwrap().stored(),
            Some(crate::Cell::Value(crate::Value::Bool(false)))
        ));
        assert!(matches!(
            sheet.cell("ZZ1").unwrap().stored(),
            Some(crate::Cell::Value(crate::Value::Error(ErrorValue::Null)))
        ));
        assert!(matches!(
            sheet.cell("AAA1").unwrap().stored(),
            Some(crate::Cell::Empty)
        ));
    }

    #[test]
    fn cell_limit_is_exact_and_one_over() {
        let limits = StreamingWorkbookLimits {
            max_cells: 1,
            ..StreamingWorkbookLimits::default()
        };
        let mut exact =
            StreamingWorkbookWriter::new(Vec::new(), context(16 * 1024 * 1024), limits).unwrap();
        exact
            .write_row(1, [StreamingCell::new(1, StreamingCellValue::Blank)])
            .unwrap();
        assert_eq!(exact.cell_count(), 1);
        assert!(exact.finish().is_ok());

        let mut over =
            StreamingWorkbookWriter::new(Vec::new(), context(16 * 1024 * 1024), limits).unwrap();
        assert!(
            over.write_row(
                1,
                [
                    StreamingCell::new(1, StreamingCellValue::Blank),
                    StreamingCell::new(2, StreamingCellValue::Blank),
                ],
            )
            .is_err()
        );
        over.write_row(1, [StreamingCell::new(1, StreamingCellValue::Blank)])
            .unwrap();
    }

    #[test]
    fn nonfinite_and_unknown_errors_are_rejected() {
        let mut number = writer();
        assert!(
            number
                .write_row(
                    1,
                    [StreamingCell::new(1, StreamingCellValue::Number(f64::NAN))]
                )
                .is_err()
        );

        let mut error = writer();
        assert!(
            error
                .write_row(
                    1,
                    [StreamingCell::new(
                        1,
                        StreamingCellValue::Error(ErrorValue::Unknown("#X!".into())),
                    )],
                )
                .is_err()
        );
    }

    #[test]
    fn row_limit_is_exact_and_one_over() {
        let limits = StreamingWorkbookLimits {
            max_row: 2,
            ..StreamingWorkbookLimits::default()
        };
        let mut writer =
            StreamingWorkbookWriter::new(Vec::new(), context(16 * 1024 * 1024), limits).unwrap();
        writer.write_row(2, std::iter::empty()).unwrap();
        assert!(writer.write_row(3, std::iter::empty()).is_err());
    }

    #[test]
    fn cell_text_and_row_scratch_limits_are_exact_and_one_over() {
        let text_limits = StreamingWorkbookLimits {
            max_cell_text_bytes: 3,
            ..StreamingWorkbookLimits::default()
        };
        let mut exact_text =
            StreamingWorkbookWriter::new(Vec::new(), context(16 * 1024 * 1024), text_limits)
                .unwrap();
        exact_text
            .write_row(1, [StreamingCell::new(1, StreamingCellValue::Text("abc"))])
            .unwrap();
        assert!(exact_text.finish().is_ok());

        let mut over_text =
            StreamingWorkbookWriter::new(Vec::new(), context(16 * 1024 * 1024), text_limits)
                .unwrap();
        assert!(
            over_text
                .write_row(1, [StreamingCell::new(1, StreamingCellValue::Text("abcd"))])
                .is_err()
        );

        let row_text = "x".repeat(64);
        let mut probe = writer();
        probe
            .write_row(
                1,
                [StreamingCell::new(
                    1,
                    StreamingCellValue::Text(row_text.as_str()),
                )],
            )
            .unwrap();
        let row_bytes = probe.worksheet_xml_bytes() - u64::try_from(SHEET_PREFIX.len()).unwrap();
        let row_limits = StreamingWorkbookLimits {
            max_row_bytes: row_bytes,
            ..StreamingWorkbookLimits::default()
        };
        let mut exact_row =
            StreamingWorkbookWriter::new(Vec::new(), context(16 * 1024 * 1024), row_limits)
                .unwrap();
        exact_row
            .write_row(
                1,
                [StreamingCell::new(
                    1,
                    StreamingCellValue::Text(row_text.as_str()),
                )],
            )
            .unwrap();
        assert!(exact_row.finish().is_ok());

        let row_limits = StreamingWorkbookLimits {
            max_row_bytes: row_bytes.saturating_sub(1),
            ..StreamingWorkbookLimits::default()
        };
        let mut over_row =
            StreamingWorkbookWriter::new(Vec::new(), context(16 * 1024 * 1024), row_limits)
                .unwrap();
        assert!(
            over_row
                .write_row(
                    1,
                    [StreamingCell::new(
                        1,
                        StreamingCellValue::Text(row_text.as_str())
                    )],
                )
                .is_err()
        );
        over_row
            .write_row(1, [StreamingCell::new(1, StreamingCellValue::Blank)])
            .unwrap();
    }

    #[test]
    fn maximum_character_text_and_unicode_byte_boundaries_are_checked() {
        let ascii = "x".repeat(32_767);
        let mut exact = writer();
        exact
            .write_row(1, [StreamingCell::new(1, StreamingCellValue::Text(&ascii))])
            .unwrap();
        assert!(exact.finish().is_ok());

        let too_many_ascii = "x".repeat(32_768);
        let mut over = writer();
        assert!(
            over.write_row(
                1,
                [StreamingCell::new(
                    1,
                    StreamingCellValue::Text(&too_many_ascii)
                )],
            )
            .is_err()
        );
        over.write_row(1, [StreamingCell::new(1, StreamingCellValue::Blank)])
            .unwrap();

        let unicode_limits = StreamingWorkbookLimits {
            max_cell_text_bytes: 65_536,
            ..StreamingWorkbookLimits::default()
        };
        let unicode = "é".repeat(32_767);
        let mut unicode_exact =
            StreamingWorkbookWriter::new(Vec::new(), context(16 * 1024 * 1024), unicode_limits)
                .unwrap();
        assert_eq!(unicode.len(), 65_534);
        unicode_exact
            .write_row(
                1,
                [StreamingCell::new(1, StreamingCellValue::Text(&unicode))],
            )
            .unwrap();
        assert!(unicode_exact.finish().is_ok());

        let unicode_over = "é".repeat(32_768);
        let mut unicode_rejected =
            StreamingWorkbookWriter::new(Vec::new(), context(16 * 1024 * 1024), unicode_limits)
                .unwrap();
        assert_eq!(unicode_over.len(), 65_536);
        assert!(
            unicode_rejected
                .write_row(
                    1,
                    [StreamingCell::new(
                        1,
                        StreamingCellValue::Text(&unicode_over)
                    )],
                )
                .is_err()
        );
    }

    #[test]
    fn worksheet_xml_limit_is_exact_and_one_over() {
        let mut probe = writer();
        probe
            .write_row(1, [StreamingCell::new(1, StreamingCellValue::Blank)])
            .unwrap();
        let exact = probe
            .worksheet_xml_bytes()
            .saturating_add(u64::try_from(SHEET_SUFFIX.len()).unwrap());

        let limits = StreamingWorkbookLimits {
            max_sheet_xml_bytes: exact,
            ..StreamingWorkbookLimits::default()
        };
        let mut exact_writer =
            StreamingWorkbookWriter::new(Vec::new(), context(16 * 1024 * 1024), limits).unwrap();
        exact_writer
            .write_row(1, [StreamingCell::new(1, StreamingCellValue::Blank)])
            .unwrap();
        assert!(exact_writer.finish().is_ok());

        let limits = StreamingWorkbookLimits {
            max_sheet_xml_bytes: exact.saturating_sub(1),
            ..StreamingWorkbookLimits::default()
        };
        let mut over_writer =
            StreamingWorkbookWriter::new(Vec::new(), context(16 * 1024 * 1024), limits).unwrap();
        assert!(
            over_writer
                .write_row(1, [StreamingCell::new(1, StreamingCellValue::Blank)])
                .is_err()
        );
        over_writer.write_row(1, std::iter::empty()).unwrap();
    }

    #[test]
    fn coordinate_boundaries_are_checked() {
        let mut row_writer = writer();
        row_writer
            .write_row(
                MAX_ROWS,
                [StreamingCell::new(MAX_COLUMNS, StreamingCellValue::Blank)],
            )
            .unwrap();
        assert!(
            row_writer
                .write_row(MAX_ROWS + 1, std::iter::empty())
                .is_err()
        );

        let mut column_writer = writer();
        assert!(
            column_writer
                .write_row(
                    1,
                    [StreamingCell::new(
                        MAX_COLUMNS + 1,
                        StreamingCellValue::Blank
                    )]
                )
                .is_err()
        );
    }

    #[test]
    fn cancellation_before_finish_reports_partial_output() {
        let (source, context) = context_pair(16 * 1024 * 1024);
        let writer =
            StreamingWorkbookWriter::new(Vec::new(), context, StreamingWorkbookLimits::default())
                .unwrap();
        assert!(writer.output_bytes() > 0);
        source.cancel();
        let error = writer.finish().unwrap_err();
        assert!(matches!(
            error,
            Error::Package(OpcError::IncompleteOutput {
                source,
                ..
            }) if matches!(*source, OpcError::Execution(litchi_core::ExecutionError::Cancelled))
        ));
    }

    #[test]
    fn cancellation_during_row_poison_is_typed() {
        let (source, context) = context_pair(16 * 1024 * 1024);
        let mut writer =
            StreamingWorkbookWriter::new(Vec::new(), context, StreamingWorkbookLimits::default())
                .unwrap();
        source.cancel();
        let error = writer
            .write_row(1, [StreamingCell::new(1, StreamingCellValue::Blank)])
            .unwrap_err();
        assert!(matches!(
            error,
            Error::Package(OpcError::IncompleteOutput {
                source,
                ..
            }) if matches!(*source, OpcError::Execution(litchi_core::ExecutionError::Cancelled))
        ));
        assert!(writer.is_poisoned());
    }

    #[test]
    fn output_limit_is_enforced_before_finalization() {
        let probe = writer().finish().unwrap();
        let output = u64::try_from(probe.len()).unwrap();
        let limits = StreamingWorkbookLimits {
            max_output_bytes: output,
            ..StreamingWorkbookLimits::default()
        };
        let mut exact =
            StreamingWorkbookWriter::new(Vec::new(), context(16 * 1024 * 1024), limits).unwrap();
        exact.write_row(1, std::iter::empty()).unwrap();
        assert!(exact.finish().is_ok());

        let limits = StreamingWorkbookLimits {
            max_output_bytes: output.saturating_sub(1),
            ..StreamingWorkbookLimits::default()
        };
        let over =
            StreamingWorkbookWriter::new(Vec::new(), context(16 * 1024 * 1024), limits).unwrap();
        let error = over.finish().unwrap_err();
        assert!(matches!(
            error,
            Error::Package(OpcError::IncompleteOutput { written, .. })
                if written == output.saturating_sub(1)
        ));
    }

    #[test]
    fn object_charge_covers_writer_parts_rows_and_cells() {
        let mut exact = StreamingWorkbookWriter::new(
            Vec::new(),
            context_with_objects_and_work(9, 8),
            StreamingWorkbookLimits::default(),
        )
        .unwrap();
        exact
            .write_row(1, [StreamingCell::new(1, StreamingCellValue::Blank)])
            .unwrap();
        assert!(exact.finish().is_ok());

        let mut one_under = StreamingWorkbookWriter::new(
            Vec::new(),
            context_with_objects_and_work(7, 8),
            StreamingWorkbookLimits::default(),
        )
        .unwrap();
        assert!(
            one_under
                .write_row(1, [StreamingCell::new(1, StreamingCellValue::Blank)])
                .is_err()
        );
    }

    #[test]
    fn work_charge_counts_rejected_row_attempts() {
        let mut writer = StreamingWorkbookWriter::new(
            Vec::new(),
            context_with_objects_and_work(7, 2),
            StreamingWorkbookLimits::default(),
        )
        .unwrap();
        assert!(writer.write_row(0, std::iter::empty()).is_err());
        assert!(writer.write_row(1, std::iter::empty()).is_err());
    }

    #[test]
    fn work_charge_is_exact_for_setup_row_and_cell() {
        let mut exact = StreamingWorkbookWriter::new(
            Vec::new(),
            context_with_objects_and_work(9, 3),
            StreamingWorkbookLimits::default(),
        )
        .unwrap();
        exact
            .write_row(1, [StreamingCell::new(1, StreamingCellValue::Blank)])
            .unwrap();
        assert!(exact.finish().is_ok());

        let mut one_under = StreamingWorkbookWriter::new(
            Vec::new(),
            context_with_objects_and_work(9, 2),
            StreamingWorkbookLimits::default(),
        )
        .unwrap();
        assert!(
            one_under
                .write_row(1, [StreamingCell::new(1, StreamingCellValue::Blank)])
                .is_err()
        );
    }

    struct NonSeekSink(Vec<u8>);
    impl Write for NonSeekSink {
        fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
            self.0.extend_from_slice(bytes);
            Ok(bytes.len())
        }
        fn flush(&mut self) -> io::Result<()> {
            Ok(())
        }
    }

    #[test]
    fn non_seek_sink_and_determinism() {
        let mut first = StreamingWorkbookWriter::new(
            NonSeekSink(Vec::new()),
            context(16 * 1024 * 1024),
            StreamingWorkbookLimits::default(),
        )
        .unwrap();
        first
            .write_row(1, [StreamingCell::new(1, StreamingCellValue::Number(2.0))])
            .unwrap();
        let first = first.finish().unwrap().0;
        let mut second = StreamingWorkbookWriter::new(
            NonSeekSink(Vec::new()),
            context(16 * 1024 * 1024),
            StreamingWorkbookLimits::default(),
        )
        .unwrap();
        second
            .write_row(1, [StreamingCell::new(1, StreamingCellValue::Number(2.0))])
            .unwrap();
        let second = second.finish().unwrap().0;
        assert_eq!(first, second);
    }

    struct FailingSink {
        writes: AtomicUsize,
    }
    impl Write for FailingSink {
        fn write(&mut self, _bytes: &[u8]) -> io::Result<usize> {
            self.writes.fetch_add(1, AtomicOrdering::Relaxed);
            Err(io::Error::other("sink failure"))
        }
        fn flush(&mut self) -> io::Result<()> {
            Ok(())
        }
    }

    #[test]
    fn sink_failure_poison_is_observable() {
        let result = StreamingWorkbookWriter::new(
            FailingSink {
                writes: AtomicUsize::new(0),
            },
            context(16 * 1024 * 1024),
            StreamingWorkbookLimits::default(),
        );
        assert!(result.is_err());
    }

    struct FailAfterSink {
        accepted: Arc<AtomicUsize>,
        limit: usize,
    }

    impl Write for FailAfterSink {
        fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
            let accepted = self.accepted.load(AtomicOrdering::Relaxed);
            if accepted >= self.limit {
                return Err(io::Error::other("injected failure"));
            }
            let amount = (self.limit - accepted).min(bytes.len());
            self.accepted.fetch_add(amount, AtomicOrdering::Relaxed);
            Ok(amount)
        }

        fn flush(&mut self) -> io::Result<()> {
            Ok(())
        }
    }

    #[test]
    fn static_member_failure_reports_exact_partial_output() {
        let accepted = Arc::new(AtomicUsize::new(0));
        let result = StreamingWorkbookWriter::new(
            FailAfterSink {
                accepted: Arc::clone(&accepted),
                limit: 1,
            },
            context(16 * 1024 * 1024),
            StreamingWorkbookLimits::default(),
        );
        let error = match result {
            Ok(_) => panic!("a sink that fails after one byte must reject construction"),
            Err(error) => error,
        };
        assert!(matches!(
            error,
            Error::Package(OpcError::IncompleteOutput { written: 1, .. })
        ));
        assert_eq!(accepted.load(AtomicOrdering::Relaxed), 1);
    }

    #[test]
    fn central_directory_failure_reports_partial_output() {
        let complete = writer().finish().unwrap();
        let limit = complete.len().saturating_sub(1);
        let accepted = Arc::new(AtomicUsize::new(0));
        let writer = StreamingWorkbookWriter::new(
            FailAfterSink {
                accepted: Arc::clone(&accepted),
                limit,
            },
            context(16 * 1024 * 1024),
            StreamingWorkbookLimits::default(),
        )
        .unwrap();
        let error = match writer.finish() {
            Ok(_) => panic!("a sink that fails before the final byte must reject finish"),
            Err(error) => error,
        };
        assert!(matches!(
            error,
            Error::Package(OpcError::IncompleteOutput { written, .. })
                if written == u64::try_from(limit).unwrap()
        ));
        assert_eq!(accepted.load(AtomicOrdering::Relaxed), limit);
    }

    #[derive(Debug)]
    struct CancelAfterFirstFinalizationWriteSink {
        accepted: Arc<AtomicUsize>,
        source: CancellationSource,
        armed: Arc<AtomicBool>,
        cancelled: bool,
    }

    impl Write for CancelAfterFirstFinalizationWriteSink {
        fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
            self.accepted
                .fetch_add(bytes.len(), AtomicOrdering::Relaxed);
            if self.armed.load(AtomicOrdering::Acquire) && !self.cancelled {
                self.cancelled = true;
                self.source.cancel();
            }
            Ok(bytes.len())
        }

        fn flush(&mut self) -> io::Result<()> {
            Ok(())
        }
    }

    #[test]
    fn cancellation_during_member_finalization_preserves_typed_progress() {
        let (source, context) = context_pair(16 * 1024 * 1024);
        let accepted = Arc::new(AtomicUsize::new(0));
        let armed = Arc::new(AtomicBool::new(false));
        let mut writer = StreamingWorkbookWriter::new(
            CancelAfterFirstFinalizationWriteSink {
                accepted: Arc::clone(&accepted),
                source: source.clone(),
                armed: Arc::clone(&armed),
                cancelled: false,
            },
            context,
            StreamingWorkbookLimits::default(),
        )
        .unwrap();
        writer
            .write_row(1, [StreamingCell::new(1, StreamingCellValue::Blank)])
            .unwrap();
        armed.store(true, AtomicOrdering::Release);
        let error = writer.finish().unwrap_err();
        assert!(matches!(
            error,
            Error::Package(OpcError::IncompleteOutput { written, source })
                if written == u64::try_from(accepted.load(AtomicOrdering::Acquire)).unwrap()
                    && matches!(*source, OpcError::Execution(litchi_core::ExecutionError::Cancelled))
        ));
    }

    #[derive(Debug)]
    struct CancelOnCentralDirectorySink {
        accepted: Arc<AtomicUsize>,
        source: CancellationSource,
        armed: Arc<AtomicBool>,
        cancelled: bool,
        tail: Vec<u8>,
    }

    impl CancelOnCentralDirectorySink {
        fn saw_central_directory(&mut self, bytes: &[u8]) -> bool {
            let mut probe = self.tail.clone();
            probe.extend_from_slice(bytes);
            let found = probe.windows(4).any(|window| window == b"PK\x01\x02");
            self.tail.clear();
            self.tail
                .extend_from_slice(&probe[probe.len().saturating_sub(3)..]);
            found
        }
    }

    impl Write for CancelOnCentralDirectorySink {
        fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
            self.accepted
                .fetch_add(bytes.len(), AtomicOrdering::Relaxed);
            if self.armed.load(AtomicOrdering::Acquire)
                && !self.cancelled
                && self.saw_central_directory(bytes)
            {
                self.cancelled = true;
                self.source.cancel();
            }
            Ok(bytes.len())
        }

        fn flush(&mut self) -> io::Result<()> {
            Ok(())
        }
    }

    #[test]
    fn cancellation_during_central_directory_preserves_typed_progress() {
        let (source, context) = context_pair(16 * 1024 * 1024);
        let accepted = Arc::new(AtomicUsize::new(0));
        let armed = Arc::new(AtomicBool::new(false));
        let mut writer = StreamingWorkbookWriter::new(
            CancelOnCentralDirectorySink {
                accepted: Arc::clone(&accepted),
                source: source.clone(),
                armed: Arc::clone(&armed),
                cancelled: false,
                tail: Vec::new(),
            },
            context,
            StreamingWorkbookLimits::default(),
        )
        .unwrap();
        writer
            .write_row(1, [StreamingCell::new(1, StreamingCellValue::Blank)])
            .unwrap();
        armed.store(true, AtomicOrdering::Release);
        let error = writer.finish().unwrap_err();
        assert!(matches!(
            error,
            Error::Package(OpcError::IncompleteOutput { written, source })
                if written == u64::try_from(accepted.load(AtomicOrdering::Acquire)).unwrap()
                    && matches!(*source, OpcError::Execution(litchi_core::ExecutionError::Cancelled))
        ));
    }

    #[derive(Debug)]
    struct CancelOnFlushSink {
        accepted: Arc<AtomicUsize>,
        source: CancellationSource,
        armed: Arc<AtomicBool>,
        cancelled: bool,
    }

    impl Write for CancelOnFlushSink {
        fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
            self.accepted
                .fetch_add(bytes.len(), AtomicOrdering::Relaxed);
            Ok(bytes.len())
        }

        fn flush(&mut self) -> io::Result<()> {
            if self.armed.load(AtomicOrdering::Acquire) && !self.cancelled {
                self.cancelled = true;
                self.source.cancel();
            }
            Ok(())
        }
    }

    #[test]
    fn cancellation_during_final_flush_preserves_typed_progress() {
        let (source, context) = context_pair(16 * 1024 * 1024);
        let accepted = Arc::new(AtomicUsize::new(0));
        let armed = Arc::new(AtomicBool::new(false));
        let mut writer = StreamingWorkbookWriter::new(
            CancelOnFlushSink {
                accepted: Arc::clone(&accepted),
                source: source.clone(),
                armed: Arc::clone(&armed),
                cancelled: false,
            },
            context,
            StreamingWorkbookLimits::default(),
        )
        .unwrap();
        writer
            .write_row(1, [StreamingCell::new(1, StreamingCellValue::Blank)])
            .unwrap();
        armed.store(true, AtomicOrdering::Release);
        let error = writer.finish().unwrap_err();
        assert!(matches!(
            error,
            Error::Package(OpcError::IncompleteOutput { written, source })
                if written == u64::try_from(accepted.load(AtomicOrdering::Acquire)).unwrap()
                    && matches!(*source, OpcError::Execution(litchi_core::ExecutionError::Cancelled))
        ));
    }

    struct ShortSink(Vec<u8>);

    impl Write for ShortSink {
        fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
            let amount = bytes.len().min(1);
            self.0.extend_from_slice(&bytes[..amount]);
            Ok(amount)
        }

        fn flush(&mut self) -> io::Result<()> {
            Ok(())
        }
    }

    struct InterruptedOnce {
        output: Vec<u8>,
        interrupted: bool,
    }

    impl Write for InterruptedOnce {
        fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
            if !self.interrupted {
                self.interrupted = true;
                return Err(io::Error::new(io::ErrorKind::Interrupted, "retry"));
            }
            self.output.extend_from_slice(bytes);
            Ok(bytes.len())
        }

        fn flush(&mut self) -> io::Result<()> {
            Ok(())
        }
    }

    #[test]
    fn short_and_interrupted_sinks_are_retried() {
        let mut short = StreamingWorkbookWriter::new(
            ShortSink(Vec::new()),
            context(16 * 1024 * 1024),
            StreamingWorkbookLimits::default(),
        )
        .unwrap();
        short
            .write_row(
                1,
                [StreamingCell::new(1, StreamingCellValue::Text("short"))],
            )
            .unwrap();
        let bytes = short.finish().unwrap().0;
        assert!(crate::Workbook::from_bytes(bytes).is_ok());

        let mut interrupted = StreamingWorkbookWriter::new(
            InterruptedOnce {
                output: Vec::new(),
                interrupted: false,
            },
            context(16 * 1024 * 1024),
            StreamingWorkbookLimits::default(),
        )
        .unwrap();
        interrupted
            .write_row(1, [StreamingCell::new(1, StreamingCellValue::Blank)])
            .unwrap();
        assert!(interrupted.finish().is_ok());
    }

    struct WriteZeroSink;

    impl Write for WriteZeroSink {
        fn write(&mut self, _bytes: &[u8]) -> io::Result<usize> {
            Ok(0)
        }

        fn flush(&mut self) -> io::Result<()> {
            Ok(())
        }
    }

    #[test]
    fn write_zero_sink_fails_without_hanging() {
        let result = StreamingWorkbookWriter::new(
            WriteZeroSink,
            context(16 * 1024 * 1024),
            StreamingWorkbookLimits::default(),
        );
        assert!(result.is_err());
    }

    struct FlushFailSink {
        output: Vec<u8>,
        flushes: Arc<AtomicUsize>,
        fail_after: usize,
    }

    impl Write for FlushFailSink {
        fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
            self.output.extend_from_slice(bytes);
            Ok(bytes.len())
        }

        fn flush(&mut self) -> io::Result<()> {
            let flushes = self.flushes.fetch_add(1, AtomicOrdering::Relaxed);
            if flushes >= self.fail_after {
                return Err(io::Error::other("flush failure"));
            }
            Ok(())
        }
    }

    #[test]
    fn final_flush_failure_reports_partial_output() {
        let probe_flushes = Arc::new(AtomicUsize::new(0));
        let probe = StreamingWorkbookWriter::new(
            FlushFailSink {
                output: Vec::new(),
                flushes: Arc::clone(&probe_flushes),
                fail_after: usize::MAX,
            },
            context(16 * 1024 * 1024),
            StreamingWorkbookLimits::default(),
        )
        .unwrap();
        let fail_after = probe_flushes.load(AtomicOrdering::Relaxed);
        drop(probe);
        let writer = StreamingWorkbookWriter::new(
            FlushFailSink {
                output: Vec::new(),
                flushes: Arc::new(AtomicUsize::new(0)),
                fail_after,
            },
            context(16 * 1024 * 1024),
            StreamingWorkbookLimits::default(),
        )
        .unwrap();
        let error = match writer.finish() {
            Ok(_) => panic!("a sink with a failing flush must reject finish"),
            Err(error) => error,
        };
        assert!(matches!(
            error,
            Error::Package(OpcError::IncompleteOutput { written, .. }) if written > 0
        ));
    }

    #[test]
    fn rows_reuse_one_bounded_scratch_buffer() {
        let mut writer = writer();
        let capacity = writer.row_scratch.capacity();
        for row in 1..=1024 {
            writer
                .write_row(row, [StreamingCell::new(1, StreamingCellValue::Blank)])
                .unwrap();
            assert_eq!(writer.row_scratch.capacity(), capacity);
        }
        assert!(writer.finish().is_ok());
    }

    #[derive(Debug)]
    struct DigestSink {
        bytes: u64,
        hash: u64,
    }

    impl DigestSink {
        fn new() -> Self {
            Self {
                bytes: 0,
                hash: 0xcbf2_9ce4_8422_2325,
            }
        }
    }

    impl Write for DigestSink {
        fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
            for byte in bytes {
                self.hash ^= u64::from(*byte);
                self.hash = self.hash.wrapping_mul(0x1000_0000_01b3);
            }
            self.bytes = self
                .bytes
                .saturating_add(u64::try_from(bytes.len()).unwrap_or(u64::MAX));
            Ok(bytes.len())
        }

        fn flush(&mut self) -> io::Result<()> {
            Ok(())
        }
    }

    fn scaling_state(row_count: u32) -> (usize, usize, u32, u64, u64, DigestSink) {
        let mut writer = StreamingWorkbookWriter::new(
            DigestSink::new(),
            context_with_objects_and_work(250_000, 300_000),
            StreamingWorkbookLimits::default(),
        )
        .unwrap();
        for row in 1..=row_count {
            writer
                .write_row(row, [StreamingCell::new(1, StreamingCellValue::Blank)])
                .unwrap();
        }
        writer
            .write_row(MAX_ROWS, [StreamingCell::new(1, StreamingCellValue::Blank)])
            .unwrap();
        let state = (
            writer.row_scratch.capacity(),
            writer.row_scratch.len(),
            writer.last_row.unwrap_or(0),
            writer.cell_count(),
            writer.worksheet_xml_bytes(),
        );
        let sink = writer.finish().unwrap();
        (state.0, state.1, state.2, state.3, state.4, sink)
    }

    #[test]
    fn large_row_scaling_keeps_fixed_memory_without_retaining_output() {
        let small = scaling_state(1_024);
        let large = scaling_state(100_000);
        assert_eq!(small.0, large.0);
        assert_eq!(small.1, large.1);
        assert_eq!(small.2, MAX_ROWS);
        assert_eq!(large.2, MAX_ROWS);
        assert_eq!(small.3, 1_025);
        assert_eq!(large.3, 100_001);
        assert!(large.4 > small.4);
        assert!(large.5.bytes > small.5.bytes);
        assert_ne!(large.5.hash, small.5.hash);
    }
}
