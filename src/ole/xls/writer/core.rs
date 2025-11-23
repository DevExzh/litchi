//! XLS file writer implementation
//!
//! This module provides functionality to create and modify Microsoft Excel files
//! in the legacy binary format (.xls files) using the BIFF (Binary Interchange File Format).
//!
//! # Architecture
//!
//! The writer generates BIFF8 records and uses the OLE writer to create the
//! compound document structure. It supports:
//! - Creating workbooks with multiple worksheets
//! - Writing cell values (numbers, strings, formulas, booleans)
//! - Shared string table management
//! - Basic cell formatting
//!
//! # Example
//!
//! ```rust,no_run
//! use litchi::ole::xls::XlsWriter;
//!
//! let mut writer = XlsWriter::new();
//! let sheet = writer.add_worksheet("Sheet1")?;
//!
//! // Write some data
//! writer.write_string(sheet, 0, 0, "Hello")?;
//! writer.write_number(sheet, 0, 1, 42.0)?;
//!
//! // Save the file
//! writer.save("output.xls")?;
//! # Ok::<(), Box<dyn std::error::Error>>(())
//! ```

use super::super::error::{XlsError, XlsResult};
use super::biff;
use super::formatting::{CellStyle, ExtendedFormat, FormattingManager};
use crate::ole::writer::OleWriter;
use std::collections::HashMap;

mod conditional_format;
mod data_validation;
mod named_range;
mod worksheet;

pub use self::conditional_format::{
    XlsConditionalFormat, XlsConditionalFormatType, XlsConditionalPattern,
};
pub use self::data_validation::{
    XlsDataValidation, XlsDataValidationOperator, XlsDataValidationType,
};
pub use self::named_range::XlsDefinedName;
use self::named_range::XlsDefinedName as InternalDefinedName;
use self::worksheet::{
    AutoFilterRange, MergedRange, WritableCell, WritableWorksheet, XlsHyperlink,
};

fn column_to_letters(col: u16) -> String {
    let mut col_index = col as u32;
    let mut buf = Vec::new();

    loop {
        let rem = (col_index % 26) as u8;
        buf.push((b'A' + rem) as char);
        col_index /= 26;
        if col_index == 0 {
            break;
        }
        col_index -= 1;
    }

    buf.iter().rev().collect()
}

fn a1_cell(row: u32, col: u16) -> String {
    let col_str = column_to_letters(col);
    let row_idx = row + 1;
    format!("{col_str}{row_idx}")
}

/// Cell value type for writing
#[derive(Debug, Clone)]
pub enum XlsCellValue {
    /// String value
    String(String),
    /// Number value (f64)
    Number(f64),
    /// Boolean value
    Boolean(bool),
    /// Formula (stored as string)
    Formula(String),
    /// Blank/empty cell
    Blank,
}
/// XLS file writer
///
/// Provides methods to create and modify XLS (BIFF8) files.
pub struct XlsWriter {
    /// Worksheets to write
    worksheets: Vec<WritableWorksheet>,
    /// Shared string table
    shared_strings: Vec<String>,
    /// String to index mapping for deduplication
    string_map: HashMap<String, u32>,
    /// Workbook-level defined names (named ranges).
    defined_names: Vec<InternalDefinedName>,
    fmt: FormattingManager,
    /// Total number of string occurrences (including duplicates) for SST.cstTotal
    sst_total: u32,
    /// Use 1904 date system (Mac) instead of 1900 (Windows)
    use_1904_dates: bool,
}

impl XlsWriter {
    /// Create a new XLS writer
    pub fn new() -> Self {
        Self {
            worksheets: Vec::new(),
            shared_strings: Vec::new(),
            string_map: HashMap::new(),
            defined_names: Vec::new(),
            sst_total: 0,
            fmt: FormattingManager::new(),
            use_1904_dates: false,
        }
    }

    /// Add a new worksheet
    ///
    /// # Arguments
    ///
    /// * `name` - Worksheet name (max 31 characters)
    ///
    /// # Returns
    ///
    /// * `Result<usize, XlsError>` - Worksheet index or error
    pub fn add_worksheet(&mut self, name: &str) -> XlsResult<usize> {
        // Validate worksheet name
        if name.is_empty() || name.len() > 31 {
            return Err(XlsError::InvalidData(
                "Worksheet name must be 1-31 characters".to_string(),
            ));
        }

        // Check for duplicate names
        if self.worksheets.iter().any(|ws| ws.name == name) {
            return Err(XlsError::InvalidData(format!(
                "Worksheet '{}' already exists",
                name
            )));
        }

        let index = self.worksheets.len();
        self.worksheets
            .push(WritableWorksheet::new(name.to_string()));
        Ok(index)
    }

    /// Write a string value to a cell
    ///
    /// # Arguments
    ///
    /// * `sheet` - Worksheet index
    /// * `row` - Row index (0-based)
    /// * `col` - Column index (0-based)
    /// * `value` - String value
    pub fn write_string(&mut self, sheet: usize, row: u32, col: u16, value: &str) -> XlsResult<()> {
        self.write_string_with_format(sheet, row, col, value, 0)
    }

    pub fn write_string_with_format(
        &mut self,
        sheet: usize,
        row: u32,
        col: u16,
        value: &str,
        format_id: u16,
    ) -> XlsResult<()> {
        self.write_cell(
            sheet,
            row,
            col,
            XlsCellValue::String(value.to_string()),
            format_id,
        )
    }

    /// Write a number value to a cell
    ///
    /// # Arguments
    ///
    /// * `sheet` - Worksheet index
    /// * `row` - Row index (0-based)
    /// * `col` - Column index (0-based)
    /// * `value` - Numeric value
    pub fn write_number(&mut self, sheet: usize, row: u32, col: u16, value: f64) -> XlsResult<()> {
        self.write_number_with_format(sheet, row, col, value, 0)
    }

    pub fn write_number_with_format(
        &mut self,
        sheet: usize,
        row: u32,
        col: u16,
        value: f64,
        format_id: u16,
    ) -> XlsResult<()> {
        self.write_cell(sheet, row, col, XlsCellValue::Number(value), format_id)
    }

    /// Write a boolean value to a cell
    ///
    /// # Arguments
    ///
    /// * `sheet` - Worksheet index
    /// * `row` - Row index (0-based)
    /// * `col` - Column index (0-based)
    /// * `value` - Boolean value
    pub fn write_boolean(
        &mut self,
        sheet: usize,
        row: u32,
        col: u16,
        value: bool,
    ) -> XlsResult<()> {
        self.write_boolean_with_format(sheet, row, col, value, 0)
    }

    pub fn write_boolean_with_format(
        &mut self,
        sheet: usize,
        row: u32,
        col: u16,
        value: bool,
        format_id: u16,
    ) -> XlsResult<()> {
        self.write_cell(sheet, row, col, XlsCellValue::Boolean(value), format_id)
    }

    /// Write a formula to a cell
    ///
    /// # Arguments
    ///
    /// * `sheet` - Worksheet index
    /// * `row` - Row index (0-based)
    /// * `col` - Column index (0-based)
    /// * `formula` - Formula string (without leading '=')
    ///
    /// # Implementation Notes
    ///
    /// Formula tokenization is deferred as a future enhancement.
    /// Formulas are currently written as blank cells.
    pub fn write_formula(
        &mut self,
        sheet: usize,
        row: u32,
        col: u16,
        formula: &str,
    ) -> XlsResult<()> {
        self.write_formula_with_format(sheet, row, col, formula, 0)
    }

    pub fn write_formula_with_format(
        &mut self,
        sheet: usize,
        row: u32,
        col: u16,
        formula: &str,
        format_id: u16,
    ) -> XlsResult<()> {
        self.write_cell(
            sheet,
            row,
            col,
            XlsCellValue::Formula(formula.to_string()),
            format_id,
        )
    }

    /// Register a number format pattern and return its BIFF format index.
    ///
    /// This is a thin wrapper around the internal `FormattingManager`
    /// and mirrors Apache POI's `HSSFDataFormat.getFormat` API. The
    /// returned index can be stored in `ExtendedFormat.format_index`
    /// to apply number formats to cells.
    pub fn register_number_format(&mut self, pattern: &str) -> u16 {
        self.fmt.register_number_format(pattern)
    }

    /// Register a reusable cell style defined by `CellStyle`.
    ///
    /// The returned identifier can be passed to the `write_*_with_format`
    /// methods to apply this style to individual cells.
    pub fn add_cell_style(&mut self, style: CellStyle) -> u16 {
        self.fmt.register_cell_style(style)
    }

    pub fn add_cell_format(&mut self, format: ExtendedFormat) -> u16 {
        self.fmt.add_format(format)
    }

    /// Validate a defined name according to basic Excel constraints.
    ///
    /// This helper enforces only well-defined structural rules from the
    /// specification:
    /// - Name MUST NOT be empty.
    /// - Name length MUST be at most 255 characters (Lbl.cch is a byte).
    fn validate_defined_name(name: &str) -> XlsResult<()> {
        if name.is_empty() {
            return Err(XlsError::InvalidData(
                "Defined name must not be empty".to_string(),
            ));
        }

        let char_count = name.chars().count();
        if char_count > u8::MAX as usize {
            return Err(XlsError::InvalidData(
                "Defined name must be at most 255 characters".to_string(),
            ));
        }

        Ok(())
    }

    /// Set a hyperlink for a single cell.
    ///
    /// Row and column indices are 0-based, matching the rest of the XLS
    /// writer APIs. The hyperlink target can be a standard URL (http, https,
    /// ftp, mailto) or an internal reference such as `Sheet1!A1` or
    /// `internal:Sheet1!A1`.
    pub fn set_hyperlink(&mut self, sheet: usize, row: u32, col: u16, url: &str) -> XlsResult<()> {
        if row > u16::MAX as u32 {
            return Err(XlsError::InvalidData(
                "set_hyperlink: row index must be <= 65535 for BIFF8".to_string(),
            ));
        }

        if col >= 256 {
            return Err(XlsError::InvalidData(
                "set_hyperlink: column index must be < 256 for BIFF8".to_string(),
            ));
        }

        let worksheet = self
            .worksheets
            .get_mut(sheet)
            .ok_or_else(|| XlsError::WorksheetNotFound(format!("Sheet {}", sheet)))?;

        // Replace any existing hyperlink on this exact cell to match
        // XLSX writer semantics.
        worksheet.hyperlinks.retain(|h| {
            !(h.first_row == row && h.last_row == row && h.first_col == col && h.last_col == col)
        });

        worksheet.add_hyperlink(XlsHyperlink {
            first_row: row,
            last_row: row,
            first_col: col,
            last_col: col,
            url: url.to_string(),
        });

        Ok(())
    }

    pub fn set_auto_filter(
        &mut self,
        sheet: usize,
        first_row: u32,
        last_row: u32,
        first_col: u16,
        last_col: u16,
    ) -> XlsResult<()> {
        if first_row > last_row || first_col > last_col {
            return Err(XlsError::InvalidData(
                "set_auto_filter: first row/col must be <= last row/col".to_string(),
            ));
        }

        if last_row > u16::MAX as u32 {
            return Err(XlsError::InvalidData(
                "set_auto_filter: row index must be <= 65535 for BIFF8".to_string(),
            ));
        }

        if last_col >= 256 {
            return Err(XlsError::InvalidData(
                "set_auto_filter: column index must be < 256 for BIFF8".to_string(),
            ));
        }

        let worksheet = self
            .worksheets
            .get_mut(sheet)
            .ok_or_else(|| XlsError::WorksheetNotFound(format!("Sheet {}", sheet)))?;

        worksheet.auto_filter = Some(AutoFilterRange {
            first_row,
            last_row,
            first_col,
            last_col,
        });

        let itab = u16::try_from(sheet + 1).map_err(|_| {
            XlsError::InvalidData(
                "set_auto_filter: sheet index exceeds BIFF8 itab limit".to_string(),
            )
        })?;

        self.defined_names.retain(|n| {
            !(n.is_built_in && n.built_in_code == Some(0x0D) && n.local_sheet == Some(itab))
        });

        let start_ref = a1_cell(first_row, first_col);
        let end_ref = a1_cell(last_row, last_col);
        let reference = format!("{start_ref}:{end_ref}");

        self.defined_names.push(InternalDefinedName {
            name: "_FilterDatabase".to_string(),
            reference,
            comment: None,
            local_sheet: Some(itab),
            target_sheet: Some(sheet as u16),
            hidden: true,
            is_function: false,
            is_built_in: true,
            built_in_code: Some(0x0D),
        });

        Ok(())
    }

    /// Define a workbook-scoped named range.
    ///
    /// The reference must currently be a simple A1 or A1:B10 style range
    /// without sheet qualifiers. More complex formulas will be rejected
    /// at serialization time to avoid emitting invalid BIFF payloads.
    pub fn define_name(&mut self, name: &str, reference: &str) -> XlsResult<()> {
        Self::validate_defined_name(name)?;

        if self.worksheets.is_empty() {
            return Err(XlsError::InvalidData(
                "define_name: workbook must have at least one worksheet".to_string(),
            ));
        }

        // For now, workbook-scoped names that refer to cell ranges are
        // anchored to the first worksheet. Users who need explicit
        // sheet scoping can use `define_name_local`.
        let target_sheet = 0u16;

        self.defined_names.push(InternalDefinedName {
            name: name.to_string(),
            reference: reference.to_string(),
            comment: None,
            local_sheet: None,
            target_sheet: Some(target_sheet),
            hidden: false,
            is_function: false,
            is_built_in: false,
            built_in_code: None,
        });

        Ok(())
    }

    /// Define a sheet-scoped named range.
    ///
    /// `sheet` is a 0-based worksheet index.
    pub fn define_name_local(
        &mut self,
        name: &str,
        reference: &str,
        sheet: usize,
    ) -> XlsResult<()> {
        Self::validate_defined_name(name)?;

        let _ = self
            .worksheets
            .get(sheet)
            .ok_or_else(|| XlsError::WorksheetNotFound(format!("Sheet {}", sheet)))?;

        let itab = u16::try_from(sheet + 1).map_err(|_| {
            XlsError::InvalidData(
                "define_name_local: sheet index exceeds BIFF8 itab limit".to_string(),
            )
        })?;

        self.defined_names.push(InternalDefinedName {
            name: name.to_string(),
            reference: reference.to_string(),
            comment: None,
            local_sheet: Some(itab),
            target_sheet: Some(sheet as u16),
            hidden: false,
            is_function: false,
            is_built_in: false,
            built_in_code: None,
        });

        Ok(())
    }

    /// Define a workbook-scoped named range with a user-visible comment.
    pub fn define_name_with_comment(
        &mut self,
        name: &str,
        reference: &str,
        comment: &str,
    ) -> XlsResult<()> {
        Self::validate_defined_name(name)?;

        if self.worksheets.is_empty() {
            return Err(XlsError::InvalidData(
                "define_name_with_comment: workbook must have at least one worksheet".to_string(),
            ));
        }

        let target_sheet = 0u16;

        self.defined_names.push(InternalDefinedName {
            name: name.to_string(),
            reference: reference.to_string(),
            comment: Some(comment.to_string()),
            local_sheet: None,
            target_sheet: Some(target_sheet),
            hidden: false,
            is_function: false,
            is_built_in: false,
            built_in_code: None,
        });

        Ok(())
    }

    /// Remove all defined names with the given name.
    ///
    /// Returns `true` if at least one name was removed.
    pub fn remove_name(&mut self, name: &str) -> bool {
        let initial_len = self.defined_names.len();
        self.defined_names.retain(|n| n.name != name);
        self.defined_names.len() < initial_len
    }

    /// Get all defined names in this workbook.
    pub fn named_ranges(&self) -> &[XlsDefinedName] {
        &self.defined_names
    }

    /// Set the width of a column in character units.
    ///
    /// The column index is 0-based (0 = column A), matching the rest of the
    /// XLS writer API. The width is specified in the same units as Excel's
    /// UI, i.e. the number of characters of the "0" glyph in the default
    /// font. Internally this is converted to BIFF8 units of 1/256 characters
    /// for the COLINFO record.
    pub fn set_column_width(&mut self, sheet: usize, col: u16, width_chars: f64) -> XlsResult<()> {
        if col >= 256 {
            return Err(XlsError::InvalidData(
                "set_column_width: column index must be < 256 for BIFF8".to_string(),
            ));
        }

        if !(width_chars.is_finite()) || width_chars <= 0.0 {
            return Err(XlsError::InvalidData(
                "set_column_width: width must be a positive finite value".to_string(),
            ));
        }

        let max_units = 255u32 * 256u32; // Excel maximum column width
        let width_units_f = (width_chars * 256.0).round();
        if width_units_f <= 0.0 || width_units_f > max_units as f64 {
            return Err(XlsError::InvalidData(
                "set_column_width: width exceeds Excel's maximum (255 characters)".to_string(),
            ));
        }

        let width_units = width_units_f as u16;

        let worksheet = self
            .worksheets
            .get_mut(sheet)
            .ok_or_else(|| XlsError::WorksheetNotFound(format!("Sheet {}", sheet)))?;
        worksheet.set_column_width(col, width_units);
        Ok(())
    }

    /// Hide a column.
    pub fn hide_column(&mut self, sheet: usize, col: u16) -> XlsResult<()> {
        if col >= 256 {
            return Err(XlsError::InvalidData(
                "hide_column: column index must be < 256 for BIFF8".to_string(),
            ));
        }

        let worksheet = self
            .worksheets
            .get_mut(sheet)
            .ok_or_else(|| XlsError::WorksheetNotFound(format!("Sheet {}", sheet)))?;
        worksheet.hide_column(col);
        Ok(())
    }

    /// Show a previously hidden column.
    pub fn show_column(&mut self, sheet: usize, col: u16) -> XlsResult<()> {
        if col >= 256 {
            return Err(XlsError::InvalidData(
                "show_column: column index must be < 256 for BIFF8".to_string(),
            ));
        }

        let worksheet = self
            .worksheets
            .get_mut(sheet)
            .ok_or_else(|| XlsError::WorksheetNotFound(format!("Sheet {}", sheet)))?;
        worksheet.show_column(col);
        Ok(())
    }

    pub fn merge_cells(
        &mut self,
        sheet: usize,
        first_row: u32,
        last_row: u32,
        first_col: u16,
        last_col: u16,
    ) -> XlsResult<()> {
        if first_row > last_row || first_col > last_col {
            return Err(XlsError::InvalidData(
                "merge_cells: first row/col must be <= last row/col".to_string(),
            ));
        }

        let worksheet = self
            .worksheets
            .get_mut(sheet)
            .ok_or_else(|| XlsError::WorksheetNotFound(format!("Sheet {}", sheet)))?;

        worksheet.add_merged_range(MergedRange {
            first_row,
            last_row,
            first_col,
            last_col,
        });

        Ok(())
    }

    /// Configure freeze panes for the specified worksheet.
    ///
    /// Row and column indices are 0-based and represent the number of
    /// rows/columns at the top/left that remain frozen.
    pub fn freeze_panes(
        &mut self,
        sheet: usize,
        freeze_rows: u32,
        freeze_cols: u16,
    ) -> XlsResult<()> {
        let worksheet = self
            .worksheets
            .get_mut(sheet)
            .ok_or_else(|| XlsError::WorksheetNotFound(format!("Sheet {}", sheet)))?;

        if freeze_rows == 0 && freeze_cols == 0 {
            worksheet.clear_freeze_panes();
            return Ok(());
        }

        if freeze_rows > u16::MAX as u32 {
            return Err(XlsError::InvalidData(
                "freeze_panes: freeze_rows must be <= 65535".to_string(),
            ));
        }

        worksheet.set_freeze_panes(freeze_rows, freeze_cols);
        Ok(())
    }

    /// Remove any freeze panes from the specified worksheet.
    pub fn unfreeze_panes(&mut self, sheet: usize) -> XlsResult<()> {
        let worksheet = self
            .worksheets
            .get_mut(sheet)
            .ok_or_else(|| XlsError::WorksheetNotFound(format!("Sheet {}", sheet)))?;
        worksheet.clear_freeze_panes();
        Ok(())
    }

    /// Set the height of a row in points.
    ///
    /// The row index is 0-based (0 = first row), and the height is specified
    /// in typographic points. Internally this is converted to twips
    /// (1/20th of a point) for the BIFF8 ROW record.
    pub fn set_row_height(&mut self, sheet: usize, row: u32, height_points: f64) -> XlsResult<()> {
        if !(height_points.is_finite()) || height_points <= 0.0 {
            return Err(XlsError::InvalidData(
                "set_row_height: height must be a positive finite value".to_string(),
            ));
        }

        if row > u16::MAX as u32 {
            return Err(XlsError::InvalidData(
                "set_row_height: row index must be <= 65535 for BIFF8".to_string(),
            ));
        }

        let height_units_f = (height_points * 20.0).round();
        if height_units_f <= 0.0 || height_units_f > u16::MAX as f64 {
            return Err(XlsError::InvalidData(
                "set_row_height: height exceeds BIFF8 row height limit".to_string(),
            ));
        }

        let height_units = height_units_f as u16;

        let worksheet = self
            .worksheets
            .get_mut(sheet)
            .ok_or_else(|| XlsError::WorksheetNotFound(format!("Sheet {}", sheet)))?;
        worksheet.set_row_height(row, height_units);
        Ok(())
    }

    /// Hide a row.
    pub fn hide_row(&mut self, sheet: usize, row: u32) -> XlsResult<()> {
        if row > u16::MAX as u32 {
            return Err(XlsError::InvalidData(
                "hide_row: row index must be <= 65535 for BIFF8".to_string(),
            ));
        }

        let worksheet = self
            .worksheets
            .get_mut(sheet)
            .ok_or_else(|| XlsError::WorksheetNotFound(format!("Sheet {}", sheet)))?;
        worksheet.hide_row(row);
        Ok(())
    }

    /// Show a previously hidden row.
    pub fn show_row(&mut self, sheet: usize, row: u32) -> XlsResult<()> {
        if row > u16::MAX as u32 {
            return Err(XlsError::InvalidData(
                "show_row: row index must be <= 65535 for BIFF8".to_string(),
            ));
        }

        let worksheet = self
            .worksheets
            .get_mut(sheet)
            .ok_or_else(|| XlsError::WorksheetNotFound(format!("Sheet {}", sheet)))?;
        worksheet.show_row(row);
        Ok(())
    }

    /// Add a data validation rule to the specified worksheet.
    pub fn add_data_validation(
        &mut self,
        sheet: usize,
        validation: XlsDataValidation,
    ) -> XlsResult<()> {
        if validation.first_row > validation.last_row || validation.first_col > validation.last_col
        {
            return Err(XlsError::InvalidData(
                "add_data_validation: first row/col must be <= last row/col".to_string(),
            ));
        }

        if let Some(title) = validation.input_title.as_ref()
            && title.len() > 32
        {
            return Err(XlsError::InvalidData(
                "Input message title must be at most 32 characters".to_string(),
            ));
        }
        if let Some(text) = validation.input_message.as_ref()
            && text.len() > 255
        {
            return Err(XlsError::InvalidData(
                "Input message text must be at most 255 characters".to_string(),
            ));
        }
        if let Some(title) = validation.error_title.as_ref()
            && title.len() > 32
        {
            return Err(XlsError::InvalidData(
                "Error message title must be at most 32 characters".to_string(),
            ));
        }
        if let Some(text) = validation.error_message.as_ref()
            && text.len() > 255
        {
            return Err(XlsError::InvalidData(
                "Error message text must be at most 255 characters".to_string(),
            ));
        }

        let worksheet = self
            .worksheets
            .get_mut(sheet)
            .ok_or_else(|| XlsError::WorksheetNotFound(format!("Sheet {}", sheet)))?;

        worksheet.add_data_validation(validation);

        Ok(())
    }

    pub fn add_conditional_format(
        &mut self,
        sheet: usize,
        cf: XlsConditionalFormat,
    ) -> XlsResult<()> {
        if cf.first_row > cf.last_row || cf.first_col > cf.last_col {
            return Err(XlsError::InvalidData(
                "add_conditional_format: first row/col must be <= last row/col".to_string(),
            ));
        }

        let worksheet = self
            .worksheets
            .get_mut(sheet)
            .ok_or_else(|| XlsError::WorksheetNotFound(format!("Sheet {}", sheet)))?;

        worksheet.add_conditional_format(cf);

        Ok(())
    }

    fn write_cell(
        &mut self,
        sheet: usize,
        row: u32,
        col: u16,
        value: XlsCellValue,
        format_id: u16,
    ) -> XlsResult<()> {
        if self.fmt.get_format(format_id).is_none() {
            return Err(XlsError::InvalidFormat(format_id));
        }

        let worksheet = self
            .worksheets
            .get_mut(sheet)
            .ok_or_else(|| XlsError::WorksheetNotFound(format!("Sheet {}", sheet)))?;

        worksheet.add_cell(WritableCell {
            row,
            col,
            value,
            format_idx: format_id,
        });

        Ok(())
    }

    /// Set the date system (1900 vs 1904)
    ///
    /// # Arguments
    ///
    /// * `use_1904` - True to use 1904 date system (Mac), false for 1900 (Windows, default)
    pub fn set_1904_dates(&mut self, use_1904: bool) {
        self.use_1904_dates = use_1904;
    }

    /// Save the XLS file
    ///
    /// # Arguments
    ///
    /// * `path` - Output file path
    ///
    /// # Returns
    ///
    /// * `Result<(), XlsError>` - Success or error
    ///
    /// # Implementation Status
    ///
    /// ✅ Basic structure generation (BOF, EOF, workbook globals)
    /// ✅ Cell record generation (Number, LabelSST, BoolErr)
    /// ✅ Shared string table (SST)
    /// ❌ Formula tokenization (formulas stored as values currently)
    /// ❌ Cell formatting (XF records)
    /// ❌ Column widths / row heights
    /// ❌ Merged cells
    /// ❌ Named ranges
    pub fn save<P: AsRef<std::path::Path>>(&mut self, path: P) -> XlsResult<()> {
        // Build shared string table
        self.build_shared_strings();

        // Generate the Workbook stream
        let workbook_stream = self.generate_workbook_stream()?;

        // Create OLE compound document
        let mut ole_writer = OleWriter::new();
        ole_writer.create_stream(&["Workbook"], &workbook_stream)?;

        // Note: SummaryInformation and DocumentSummaryInformation streams are optional
        // They provide metadata like title, author, creation date, etc.
        // For now, we skip these as they're not required for a functional XLS file
        // They can be added in a future enhancement for complete metadata support

        // Save to file
        ole_writer.save(path)?;

        Ok(())
    }

    /// Write to a writer (useful for testing and in-memory generation)
    ///
    /// # Arguments
    ///
    /// * `writer` - Output writer
    ///
    /// # Returns
    ///
    /// * `Result<(), XlsError>` - Success or error
    pub fn write_to<W: std::io::Write + std::io::Seek>(&mut self, writer: &mut W) -> XlsResult<()> {
        // Build shared string table
        self.build_shared_strings();

        // Generate the Workbook stream
        let workbook_stream = self.generate_workbook_stream()?;

        // Create OLE compound document
        let mut ole_writer = OleWriter::new();
        ole_writer.create_stream(&["Workbook"], &workbook_stream)?;

        // Write to the provided writer
        ole_writer.write_to(writer)?;

        Ok(())
    }

    /// Build the shared string table from all string cells
    fn build_shared_strings(&mut self) {
        self.shared_strings.clear();
        self.string_map.clear();
        self.sst_total = 0;

        // Collect all unique strings from all worksheets
        for worksheet in &self.worksheets {
            for cell in worksheet.cells.values() {
                if let XlsCellValue::String(ref s) = cell.value {
                    // Count total occurrences
                    self.sst_total = self.sst_total.saturating_add(1);
                    // Insert unique strings
                    if !self.string_map.contains_key(s) {
                        let index = self.shared_strings.len() as u32;
                        self.string_map.insert(s.clone(), index);
                        self.shared_strings.push(s.clone());
                    }
                }
            }
        }
    }

    /// Generate the complete Workbook stream with all BIFF records
    fn generate_workbook_stream(&self) -> XlsResult<Vec<u8>> {
        let mut stream = Vec::new();

        // === Workbook Globals ===

        // BOF record (workbook)
        biff::write_bof(&mut stream, 0x0005)?;

        // CodePage record - BIFF8 requires Unicode codepage 1200 (0x04B0)
        biff::write_codepage(&mut stream, 0x04B0)?;

        // Date1904 record
        biff::write_date1904(&mut stream, self.use_1904_dates)?;

        // Window1 record (workbook window properties)
        biff::write_window1(&mut stream)?;

        // Write minimal formatting tables so XF index 0 is valid.
        // Order mirrors Apache POI's workbook creation:
        //  - FONT records
        //  - FORMAT records (built-in 0..7 + custom)
        //  - XF records (style and cell formats)
        self.fmt.write_fonts(&mut stream)?;
        self.fmt.write_number_formats(&mut stream)?;
        self.fmt.write_formats(&mut stream)?;

        // Built-in STYLE records and UseSelFS flag to align with Excel / POI
        // defaults. This makes standard cell styles (Normal, Currency, Percent,
        // etc.) visible to Excel even though we currently only use the default
        // cell XF (index 15) for all cells.
        biff::write_builtin_styles(&mut stream)?;
        biff::write_usesel_fs(&mut stream)?;

        // BoundSheet8 records (one per worksheet)
        // We need to calculate positions, so we'll write them after we know the sizes
        let mut boundsheet_positions = Vec::new();
        for worksheet in &self.worksheets {
            // Placeholder - we'll update positions later
            boundsheet_positions.push(stream.len());
            biff::write_boundsheet(&mut stream, 0, &worksheet.name)?;
        }

        // Internal SUPBOOK / EXTERNSHEET records are required for 3D
        // references used by defined names (NameParsedFormula). We keep
        // the model minimal by generating a single internal SUPBOOK and
        // one XTI entry per worksheet.
        if !self.defined_names.is_empty() && !self.worksheets.is_empty() {
            let sheet_count = u16::try_from(self.worksheets.len()).unwrap_or(u16::MAX);
            biff::write_supbook_internal(&mut stream, sheet_count)?;
            biff::write_externsheet_internal(&mut stream, sheet_count)?;
        }

        // NAME (Lbl) records for workbook- and sheet-scoped defined names.
        // These are stored in the globals substream and reference cell
        // areas using BIFF8 formula tokens.
        for defined_name in &self.defined_names {
            let rgce = defined_name.to_biff_formula()?;
            biff::write_name(&mut stream, defined_name, &rgce)?;
        }

        // SST record (shared string table)
        if !self.shared_strings.is_empty() {
            biff::write_sst(&mut stream, &self.shared_strings, self.sst_total)?;
        }

        // EOF record (end of workbook globals)
        biff::write_eof(&mut stream)?;

        // === Worksheets ===

        // Track actual worksheet positions
        let mut actual_positions = Vec::new();

        for worksheet in &self.worksheets {
            // Record the position of this worksheet's BOF
            let worksheet_pos = stream.len() as u32;
            actual_positions.push(worksheet_pos);

            // BOF record (worksheet)
            biff::write_bof(&mut stream, 0x0010)?;

            // DIMENSIONS record
            biff::write_dimensions(
                &mut stream,
                worksheet.first_row,
                worksheet.last_row,
                worksheet.first_col,
                worksheet.last_col,
            )?;

            // Required sheet records for worksheet substream per MS-XLS.
            //
            // Apache POI writes WINDOW2 first and then (optionally) PANE
            // immediately afterwards when freeze panes are configured. We
            // mirror that ordering here to avoid Excel interpreting the
            // pane as a generic split window.
            biff::write_wsbool(&mut stream)?;
            let has_freeze_panes = worksheet.freeze_panes.is_some();
            biff::write_window2(&mut stream, has_freeze_panes)?;

            if let Some(panes) = worksheet.freeze_panes {
                biff::write_pane(&mut stream, panes.freeze_rows, panes.freeze_cols)?;
            }

            if let Some(af) = worksheet.auto_filter {
                let width = u32::from(af.last_col).saturating_sub(u32::from(af.first_col)) + 1;
                let c_entries = u16::try_from(width).map_err(|_| {
                    XlsError::InvalidData(
                        "set_auto_filter: auto-filter column span exceeds BIFF8 limit".to_string(),
                    )
                })?;
                biff::write_autofilterinfo(&mut stream, c_entries)?;
            }

            // Column width / hidden state via COLINFO records.
            if !worksheet.column_widths.is_empty() || !worksheet.hidden_columns.is_empty() {
                use std::collections::BTreeSet;

                let mut cols = BTreeSet::<u16>::new();
                cols.extend(worksheet.column_widths.keys().copied());
                cols.extend(worksheet.hidden_columns.iter().copied());

                for col in cols {
                    let width_units = worksheet
                        .column_widths
                        .get(&col)
                        .copied()
                        // Default matches POI's ColumnInfoRecord constructor.
                        .unwrap_or(2275u16);
                    let hidden = worksheet.hidden_columns.contains(&col);
                    biff::write_colinfo(&mut stream, col, col, width_units, hidden)?;
                }
            }

            // Pre-compute row spans (first/last used column per row) for ROW records.
            use std::collections::HashMap as StdHashMap;
            let mut row_spans: StdHashMap<u32, (u16, u16)> = StdHashMap::new();

            for &(row, col) in worksheet.cells.keys() {
                let entry = row_spans.entry(row).or_insert((col, col.saturating_add(1)));
                if col < entry.0 {
                    entry.0 = col;
                }
                if col.saturating_add(1) > entry.1 {
                    entry.1 = col.saturating_add(1);
                }
            }

            // ROW records for rows with custom height or hidden state.
            if !worksheet.row_heights.is_empty() || !worksheet.hidden_rows.is_empty() {
                use std::collections::BTreeSet;

                let mut rows = BTreeSet::<u32>::new();
                rows.extend(worksheet.row_heights.keys().copied());
                rows.extend(worksheet.hidden_rows.iter().copied());

                for row in rows {
                    let (first_col, last_col_plus1) =
                        row_spans.get(&row).copied().unwrap_or((0, 0));
                    let height = worksheet
                        .row_heights
                        .get(&row)
                        // Default height matches POI's RowRecord constructor (0x00FF).
                        .copied()
                        .unwrap_or(0x00FFu16);
                    let hidden = worksheet.hidden_rows.contains(&row);
                    biff::write_row(&mut stream, row, first_col, last_col_plus1, height, hidden)?;
                }
            }

            // Cell records (sorted by row, then column)
            let mut sorted_cells: Vec<_> = worksheet.cells.iter().collect();
            sorted_cells.sort_by_key(|(k, _)| *k);

            for ((row, col), cell) in sorted_cells {
                let xf_index = self.fmt.cell_xf_index_for(cell.format_idx);
                match &cell.value {
                    XlsCellValue::Number(value) => {
                        biff::write_number(&mut stream, *row, *col, xf_index, *value)?;
                    },
                    XlsCellValue::String(s) => {
                        let sst_index = *self.string_map.get(s).unwrap();
                        biff::write_labelsst(&mut stream, *row, *col, xf_index, sst_index)?;
                    },
                    XlsCellValue::Boolean(value) => {
                        biff::write_boolerr(&mut stream, *row, *col, xf_index, *value)?;
                    },
                    XlsCellValue::Formula(_formula) => {
                        // Formula tokenization not yet implemented
                        // Write as blank cell for now
                        // Future enhancement: Parse formula to RPN tokens and write FORMULA record
                    },
                    XlsCellValue::Blank => {
                        // Skip blank cells
                    },
                }
            }

            // Hyperlink records for cells or ranges.
            for hyperlink in &worksheet.hyperlinks {
                biff::write_hyperlink(
                    &mut stream,
                    hyperlink.first_row,
                    hyperlink.last_row,
                    hyperlink.first_col,
                    hyperlink.last_col,
                    &hyperlink.url,
                )?;
            }

            if !worksheet.merged_ranges.is_empty() {
                biff::write_mergedcells(
                    &mut stream,
                    worksheet
                        .merged_ranges
                        .iter()
                        .map(|r| (r.first_row, r.last_row, r.first_col, r.last_col)),
                )?;
            }

            if !worksheet.data_validations.is_empty() {
                let dv_count = worksheet.data_validations.len() as u32;
                biff::write_dval(&mut stream, dv_count)?;

                for dv in &worksheet.data_validations {
                    let payload = dv.validation_type.to_biff_payload()?;

                    let ranges = [(dv.first_row, dv.last_row, dv.first_col, dv.last_col)];

                    let formula1 = payload.formula1.as_deref();
                    let formula2 = payload.formula2.as_deref();

                    let dv_config = biff::DvConfig {
                        data_type: payload.data_type,
                        operator: payload.operator,
                        error_style: 0, // errorStyle: STOP
                        empty_cell_allowed: true,
                        suppress_dropdown_arrow: false,
                        is_explicit_list_formula: payload.is_explicit_list,
                        show_prompt_on_cell_selected: dv.show_input_message,
                        prompt_title: dv.input_title.as_deref(),
                        prompt_text: dv.input_message.as_deref(),
                        show_error_on_invalid_value: dv.show_error_alert,
                        error_title: dv.error_title.as_deref(),
                        error_text: dv.error_message.as_deref(),
                        formula1,
                        formula2,
                    };

                    biff::write_dv(&mut stream, &dv_config, &ranges)?;
                }
            }

            if !worksheet.conditional_formats.is_empty() {
                for cf in &worksheet.conditional_formats {
                    let ranges = [(cf.first_row, cf.last_row, cf.first_col, cf.last_col)];

                    // One CFHEADER per rule with a single region keeps the
                    // implementation simple and matches Excel's expectations.
                    biff::write_cfheader(&mut stream, &ranges, 1)?;

                    let (condition_type, comparison_op, formula1, formula2) =
                        cf.format_type.to_biff_payload()?;

                    biff::write_cfrule(
                        &mut stream,
                        condition_type,
                        comparison_op,
                        &formula1,
                        &formula2,
                        cf.to_biff_pattern(),
                    )?;
                }
            }

            // EOF record (end of worksheet)
            biff::write_eof(&mut stream)?;
        }

        // Go back and update BoundSheet positions
        for (i, &pos) in actual_positions.iter().enumerate() {
            let boundsheet_pos = boundsheet_positions[i];
            // Position field starts at offset 4 in the record (after header)
            let pos_offset = boundsheet_pos + 4;
            stream[pos_offset..pos_offset + 4].copy_from_slice(&pos.to_le_bytes());
        }

        Ok(stream)
    }

    /// Get the number of worksheets in this workbook
    pub fn worksheet_count(&self) -> usize {
        self.worksheets.len()
    }

    /// Get worksheet name by index
    pub fn get_worksheet_name(&self, index: usize) -> Option<&str> {
        self.worksheets.get(index).map(|w| w.name.as_str())
    }

    // Implementation status notes:
    // ✅ Building shared string table (SST) with deduplication - IMPLEMENTED
    // ✅ Generating BIFF8 records for all cell types - IMPLEMENTED (Number, LabelSST, BoolErr)
    // ❌ Worksheet management (rename, delete, reorder) - Future enhancement
    // ❌ Cell formatting (fonts, colors, borders, number formats) - Future enhancement
    // ❌ Column widths and row heights - Future enhancement
    // ❌ Merged cells - Future enhancement
    // ✅ Named ranges (simple A1-style, workbook and sheet scoped) - IMPLEMENTED
    // ❌ Formulas (parsing and tokenization) - Future enhancement
}

impl Default for XlsWriter {
    fn default() -> Self {
        Self::new()
    }
}

/// Implementation notes for BIFF record generation:
///
/// All core BIFF8 records have been implemented in the `biff` module:
/// - ✅ write_bof() - Beginning of File (0x0809)
/// - ✅ write_eof() - End of File (0x000A)
/// - ✅ write_codepage() - Code page (0x0042)
/// - ✅ write_date1904() - Date system (0x0022)
/// - ✅ write_window1() - Workbook window properties (0x003D)
/// - ✅ write_boundsheet() - Sheet metadata (0x0085)
/// - ✅ write_dimensions() - Worksheet dimensions (0x0200)
/// - ✅ write_sst() - Shared string table with CONTINUE support (0x00FC)
/// - ✅ write_number() - Floating point cell (0x0203)
/// - ✅ write_labelsst() - String cell (0x00FD)
/// - ✅ write_boolerr() - Boolean/error cell (0x0205)
/// - ✅ write_continue() - Continuation record (0x003C)
///
/// Future enhancements:
/// - FORMULA record (0x0006) - For formula cells with RPN tokens
/// - XF records (0x00E0) - For cell formatting
/// - FONT records (0x0031) - For font definitions
/// - FORMAT records (0x041E) - For number formats
/// - COLINFO records (0x007D) - For column widths
/// - ROW records (0x0208) - For row heights
/// - MERGEDCELLS records (0x00E5) - For merged cell ranges
/// - NAME records (0x0018) - For named ranges
#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_create_writer() {
        let writer = XlsWriter::new();
        assert_eq!(writer.worksheets.len(), 0);
        assert_eq!(writer.shared_strings.len(), 0);
    }

    #[test]
    fn test_add_worksheet() {
        let mut writer = XlsWriter::new();
        let idx = writer.add_worksheet("Sheet1").unwrap();
        assert_eq!(idx, 0);
        assert_eq!(writer.worksheets.len(), 1);
        assert_eq!(writer.worksheets[0].name, "Sheet1");
    }

    #[test]
    fn test_write_string() {
        let mut writer = XlsWriter::new();
        let sheet = writer.add_worksheet("Sheet1").unwrap();
        writer.write_string(sheet, 0, 0, "Hello").unwrap();
        assert_eq!(writer.worksheets[0].cells.len(), 1);
    }

    #[test]
    fn test_write_number() {
        let mut writer = XlsWriter::new();
        let sheet = writer.add_worksheet("Sheet1").unwrap();
        writer.write_number(sheet, 0, 0, 42.5).unwrap();
        assert_eq!(writer.worksheets[0].cells.len(), 1);
    }
}
