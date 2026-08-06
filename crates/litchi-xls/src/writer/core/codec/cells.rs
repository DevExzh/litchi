use super::super::super::formatting::{CellStyle, ExtendedFormat};
use super::super::*;
use crate::error::{Error, Result};

impl Writer {
    /// Write a string value to a cell
    ///
    /// # Arguments
    ///
    /// * `sheet` - Worksheet index
    /// * `row` - Row index (0-based)
    /// * `col` - Column index (0-based)
    /// * `value` - String value
    pub fn write_string(&mut self, sheet: usize, row: u32, col: u16, value: &str) -> Result<()> {
        self.write_string_with_format(sheet, row, col, value, 0)
    }

    pub fn write_string_with_format(
        &mut self,
        sheet: usize,
        row: u32,
        col: u16,
        value: &str,
        format_id: u16,
    ) -> Result<()> {
        let pos = CellPos::try_new(row, col)?;
        self.write_cell(sheet, pos, CellValue::String(value.to_string()), format_id)
    }

    /// Write a number value to a cell
    ///
    /// # Arguments
    ///
    /// * `sheet` - Worksheet index
    /// * `row` - Row index (0-based)
    /// * `col` - Column index (0-based)
    /// * `value` - Numeric value
    pub fn write_number(&mut self, sheet: usize, row: u32, col: u16, value: f64) -> Result<()> {
        self.write_number_with_format(sheet, row, col, value, 0)
    }

    pub fn write_number_with_format(
        &mut self,
        sheet: usize,
        row: u32,
        col: u16,
        value: f64,
        format_id: u16,
    ) -> Result<()> {
        if !value.is_finite() {
            return Err(Error::InvalidData(
                "cell number must be finite for BIFF8 serialization".to_string(),
            ));
        }
        let pos = CellPos::try_new(row, col)?;
        self.write_cell(sheet, pos, CellValue::Number(value), format_id)
    }

    /// Write a boolean value to a cell
    ///
    /// # Arguments
    ///
    /// * `sheet` - Worksheet index
    /// * `row` - Row index (0-based)
    /// * `col` - Column index (0-based)
    /// * `value` - Boolean value
    pub fn write_boolean(&mut self, sheet: usize, row: u32, col: u16, value: bool) -> Result<()> {
        self.write_boolean_with_format(sheet, row, col, value, 0)
    }

    pub fn write_boolean_with_format(
        &mut self,
        sheet: usize,
        row: u32,
        col: u16,
        value: bool,
        format_id: u16,
    ) -> Result<()> {
        let pos = CellPos::try_new(row, col)?;
        self.write_cell(sheet, pos, CellValue::Boolean(value), format_id)
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
    /// The supported BIFF8 formula subset includes constants, cell/range
    /// references, arithmetic/comparison operators, and built-in functions
    /// recognized by [`FormulaTokenizer`](crate::writer::FormulaTokenizer).
    pub fn write_formula(&mut self, sheet: usize, row: u32, col: u16, formula: &str) -> Result<()> {
        self.write_formula_with_format(sheet, row, col, formula, 0)
    }

    pub fn write_formula_with_format(
        &mut self,
        sheet: usize,
        row: u32,
        col: u16,
        formula: &str,
        format_id: u16,
    ) -> Result<()> {
        let pos = CellPos::try_new(row, col)?;
        self.write_cell(
            sheet,
            pos,
            CellValue::Formula(formula.to_string()),
            format_id,
        )
    }

    /// Write a formula with explicit BIFF8 `Formula` metadata.
    ///
    /// The shared-formula flag is intentionally rejected until this writer
    /// owns the corresponding `ShrFmla` sequence. All other flags and the
    /// opaque application cache are emitted verbatim.
    pub fn write_formula_with_metadata(
        &mut self,
        sheet: usize,
        row: u32,
        col: u16,
        formula: &str,
        metadata: crate::FormulaMetadata,
    ) -> Result<()> {
        self.write_formula_with_format_and_metadata(sheet, row, col, formula, 0, metadata)
    }

    /// Write a formatted formula with explicit BIFF8 `Formula` metadata.
    pub fn write_formula_with_format_and_metadata(
        &mut self,
        sheet: usize,
        row: u32,
        col: u16,
        formula: &str,
        format_id: u16,
        metadata: crate::FormulaMetadata,
    ) -> Result<()> {
        crate::formula_metadata::validate_for_write(&metadata)?;
        let pos = CellPos::try_new(row, col)?;
        self.write_cell_with_formula_metadata(
            sheet,
            pos,
            CellValue::Formula(formula.to_string()),
            format_id,
            Some(metadata),
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

    /// Set a hyperlink for a single cell.
    ///
    /// Row and column indices are 0-based, matching the rest of the XLS
    /// writer APIs. The hyperlink target can be a standard URL (http, https,
    /// ftp, mailto) or an internal reference such as `Sheet1!A1` or
    /// `internal:Sheet1!A1`.
    pub fn set_hyperlink(&mut self, sheet: usize, row: u32, col: u16, url: &str) -> Result<()> {
        if row > u16::MAX as u32 {
            return Err(Error::InvalidData(
                "set_hyperlink: row index must be <= 65535 for BIFF8".to_string(),
            ));
        }

        if col >= 256 {
            return Err(Error::InvalidData(
                "set_hyperlink: column index must be < 256 for BIFF8".to_string(),
            ));
        }

        let worksheet = self
            .worksheets
            .get_mut(sheet)
            .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet {}", sheet)))?;

        // Replace any existing hyperlink on this exact cell to match
        // XLSX writer semantics.
        worksheet.hyperlinks.retain(|h| {
            !(h.first_row == row && h.last_row == row && h.first_col == col && h.last_col == col)
        });

        worksheet.add_hyperlink(Hyperlink {
            first_row: row,
            last_row: row,
            first_col: col,
            last_col: col,
            url: url.to_string(),
        });

        Ok(())
    }

    fn write_cell(
        &mut self,
        sheet: usize,
        pos: CellPos,
        value: CellValue,
        format_id: u16,
    ) -> Result<()> {
        self.write_cell_with_formula_metadata(sheet, pos, value, format_id, None)
    }

    fn write_cell_with_formula_metadata(
        &mut self,
        sheet: usize,
        pos: CellPos,
        value: CellValue,
        format_id: u16,
        formula_metadata: Option<crate::FormulaMetadata>,
    ) -> Result<()> {
        if self.fmt.get_format(format_id).is_none() {
            return Err(Error::InvalidFormat(format_id));
        }

        let worksheet = self
            .worksheets
            .get_mut(sheet)
            .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet {}", sheet)))?;

        worksheet.add_cell(
            WritableCell::new(pos, value, format_id, None).with_formula_metadata(formula_metadata),
        );

        Ok(())
    }
}
