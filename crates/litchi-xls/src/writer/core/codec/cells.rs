use super::super::super::formatting::{CellStyle, ExtendedFormat};
use super::super::{CellPos, CellValue, Hyperlink, WritableCell, Writer};
use crate::error::{Error, Result};
use crate::writer::formula::{FormulaTokenizer, encode_ptg_tokens};
use crate::writer::string_limits::{SHARED_STRING_UNITS, ensure_utf16_len_within};

/// Refuses a formula cell the workbook write would refuse, by running the
/// write's own steps for it — tokenizing, encoding and building its
/// `Formula` record, into a sink — when the cell is set. Temporary tokens
/// are released after validation rather than retained alongside every cell.
/// The steps depend only on the cell, its formula and its metadata.
fn check_formula_cell(pos: CellPos, formula: &str, metadata: crate::FormulaMetadata) -> Result<()> {
    let expression = formula.strip_prefix('=').unwrap_or(formula);
    let tokens = FormulaTokenizer::new().tokenize(expression)?;
    let encoded = encode_ptg_tokens(&tokens)?;
    crate::writer::biff::write_formula_with_metadata(
        &mut std::io::sink(),
        u32::from(pos.row()),
        u16::from(pos.col()),
        0,
        &encoded,
        metadata,
    )?;
    Ok(())
}

impl Writer {
    /// Write a string value to a cell
    ///
    /// # Arguments
    ///
    /// * `sheet` - Worksheet index
    /// * `row` - Row index (0-based)
    /// * `col` - Column index (0-based)
    /// * `value` - String value
    /// # Errors
    ///
    /// Returns [`Error::StringTooLong`] when `value` is longer than the 65,535
    /// UTF-16 code units a BIFF8 shared string can hold, leaving the cell
    /// unchanged, and an error if any other validation, decoding, encoding,
    /// or the requested operation fails.
    pub fn write_string(&mut self, sheet: usize, row: u32, col: u16, value: &str) -> Result<()> {
        self.write_string_with_format(sheet, row, col, value, 0)
    }

    /// # Errors
    ///
    /// Returns [`Error::StringTooLong`] when `value` is longer than the 65,535
    /// UTF-16 code units a BIFF8 shared string can hold, leaving the cell
    /// unchanged, and an error if any other validation, decoding, encoding,
    /// or the requested operation fails.
    pub fn write_string_with_format(
        &mut self,
        sheet: usize,
        row: u32,
        col: u16,
        value: &str,
        format_id: u16,
    ) -> Result<()> {
        let pos = CellPos::try_new(row, col)?;
        ensure_utf16_len_within(value, SHARED_STRING_UNITS, "shared string")?;
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
    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn write_number(&mut self, sheet: usize, row: u32, col: u16, value: f64) -> Result<()> {
        self.write_number_with_format(sheet, row, col, value, 0)
    }

    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
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
    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn write_boolean(&mut self, sheet: usize, row: u32, col: u16, value: bool) -> Result<()> {
        self.write_boolean_with_format(sheet, row, col, value, 0)
    }

    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
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
    /// # Errors
    ///
    /// Refuses, when it is called, a formula the workbook could not be
    /// written with: one the tokenizer rejects (such as `SUM(`), one that
    /// encodes to no tokens, one with a string constant longer than 255
    /// UTF-16 code units, or one whose `Formula` record would be longer than
    /// one BIFF8 record. A refused formula leaves the cell unchanged. Also
    /// returns an error for a cell outside the BIFF8 grid, an unknown
    /// worksheet or format, or a cell that belongs to a formula group.
    pub fn write_formula(&mut self, sheet: usize, row: u32, col: u16, formula: &str) -> Result<()> {
        self.write_formula_with_format(sheet, row, col, formula, 0)
    }

    /// # Errors
    ///
    /// Refuses a formula as [`Self::write_formula`] does, before staging it.
    pub fn write_formula_with_format(
        &mut self,
        sheet: usize,
        row: u32,
        col: u16,
        formula: &str,
        format_id: u16,
    ) -> Result<()> {
        let pos = CellPos::try_new(row, col)?;
        check_formula_cell(
            pos,
            formula,
            crate::FormulaMetadata::new().with_always_calculate(true),
        )?;
        self.stage_cell(
            sheet,
            pos,
            CellValue::Formula(formula.to_string()),
            format_id,
            None,
        )
    }

    /// Write a formula with explicit BIFF8 `Formula` metadata.
    ///
    /// The shared-formula flag is intentionally rejected until this writer
    /// owns the corresponding `ShrFmla` sequence. All other flags and the
    /// opaque application cache are emitted verbatim.
    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
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
    /// # Errors
    ///
    /// Refuses a formula as [`Self::write_formula`] does, and metadata the
    /// `Formula` record cannot carry for this cell, before staging it.
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
        check_formula_cell(pos, formula, metadata.clone())?;
        self.stage_cell(
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
    ///
    /// # Errors
    ///
    /// Returns [`Error::InvalidData`] for an empty pattern and
    /// [`Error::StringTooLong`] for one longer than the 255 UTF-16 code units
    /// a BIFF8 `Format` record holds. A refused pattern leaves the writer
    /// unchanged, so it still writes.
    pub fn register_number_format(&mut self, pattern: &str) -> Result<u16> {
        self.fmt.register_number_format(pattern)
    }

    /// Register a reusable cell style defined by `CellStyle`.
    ///
    /// The returned identifier can be passed to the `write_*_with_format`
    /// methods to apply this style to individual cells.
    ///
    /// # Errors
    ///
    /// Refuses the style's number format as [`Self::register_number_format`]
    /// does; a font a BIFF8 `Font` record cannot store (a name that is empty,
    /// holds a NUL or is longer than 31 UTF-16 code units, which is
    /// [`Error::StringTooLong`], or a height, weight, underline or color
    /// outside BIFF8's values); and, with [`Error::TooMany`], a font past the
    /// 1022 BIFF8 can address or an XF past the XF index space (4050 XF
    /// records once the workbook has `XFExt` records or a pivot table). Every
    /// check runs before anything is registered, so a refused style leaves
    /// the writer unchanged.
    pub fn add_cell_style(&mut self, style: CellStyle) -> Result<u16> {
        self.check_xf_capacity(1)?;
        self.fmt.register_cell_style(style)
    }

    /// Register a cell format (XF) and return the identifier the
    /// `write_*_with_format` methods take.
    ///
    /// # Errors
    ///
    /// Refuses a format whose font or number format index names nothing
    /// registered, and a format past the XF index space, as
    /// [`Self::add_cell_style`] does. A refused format leaves the writer
    /// unchanged.
    pub fn add_cell_format(&mut self, format: ExtendedFormat) -> Result<u16> {
        self.check_xf_capacity(1)?;
        self.fmt.add_format(format)
    }

    /// Set a hyperlink for a single cell.
    ///
    /// Row and column indices are 0-based, matching the rest of the XLS
    /// writer APIs. The hyperlink target can be a standard URL (http, https,
    /// ftp, mailto) or an internal reference such as `Sheet1!A1` or
    /// `internal:Sheet1!A1`. Surrounding whitespace is not part of the
    /// target, and an empty target removes the cell's hyperlink.
    /// # Errors
    ///
    /// Returns [`Error::StringTooLong`] for a target longer than one `HLink`
    /// record holds (4,093 UTF-16 code units for an internal location, 4,085
    /// for a URL) and [`Error::InvalidData`] for one containing a NUL
    /// character, leaving the worksheet unchanged; and an error for a cell
    /// outside the BIFF8 grid or a missing worksheet.
    pub fn set_hyperlink(&mut self, sheet: usize, row: u32, col: u16, url: &str) -> Result<()> {
        if row > u32::from(u16::MAX) {
            return Err(Error::InvalidData(
                "set_hyperlink: row index must be <= 65535 for BIFF8".to_string(),
            ));
        }

        if col >= 256 {
            return Err(Error::InvalidData(
                "set_hyperlink: column index must be < 256 for BIFF8".to_string(),
            ));
        }
        crate::writer::biff::validate_hyperlink_target(url)?;

        let worksheet = self
            .worksheets
            .get_mut(sheet)
            .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet {sheet}")))?;

        // Replace any existing hyperlink on this exact cell to match
        // XLSX writer semantics.
        worksheet.hyperlinks.retain(|h| {
            !(h.first_row == row && h.last_row == row && h.first_col == col && h.last_col == col)
        });

        if crate::writer::biff::classify_hyperlink(url).is_some() {
            worksheet.add_hyperlink(Hyperlink {
                first_row: row,
                last_row: row,
                first_col: col,
                last_col: col,
                url: url.to_string(),
            });
        }

        Ok(())
    }

    fn write_cell(
        &mut self,
        sheet: usize,
        pos: CellPos,
        value: CellValue,
        format_id: u16,
    ) -> Result<()> {
        self.stage_cell(sheet, pos, value, format_id, None)
    }

    /// Stages one validated cell without retaining derived formula tokens.
    fn stage_cell(
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
            .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet {sheet}")))?;

        if worksheet
            .cells
            .get(&(u32::from(pos.row()), u16::from(pos.col())))
            .and_then(|cell| cell.formula_metadata.as_ref())
            .is_some_and(|metadata| {
                metadata.shared_owner().is_some() || metadata.array_owner().is_some()
            })
        {
            return Err(Error::InvalidData(format!(
                "cell ({}, {}) belongs to a formula group and cannot be overwritten independently",
                pos.row(),
                pos.col()
            )));
        }

        let cell =
            WritableCell::new(pos, value, format_id, None).with_formula_metadata(formula_metadata);
        worksheet.add_cell(cell);

        Ok(())
    }
}
