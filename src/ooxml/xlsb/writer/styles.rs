//! Styles writer for XLSB format
//!
//! This module provides functionality to write style information (fonts, fills, borders, etc.)
//! to XLSB binary format files.

use crate::ooxml::xlsb::error::XlsbResult;
use crate::ooxml::xlsb::records::record_types;
use crate::ooxml::xlsb::styles::Border;
use crate::ooxml::xlsb::styles_table::{Fill, Font};
use crate::ooxml::xlsb::writer::RecordWriter;
use std::io::Write;

/// Styles writer for XLSB with support for custom fonts, fills, and borders
pub struct StylesWriter {
    /// Custom fonts (index 0 is the default font)
    fonts: Vec<Font>,
    /// Custom fills (indices 0-1 are default fills)
    fills: Vec<Fill>,
    /// Custom borders (index 0 is the default border)
    borders: Vec<Border>,
}

impl StylesWriter {
    /// Create a new styles writer with default styles
    pub fn new() -> Self {
        StylesWriter {
            fonts: vec![Font::default()],
            fills: vec![Fill::default(), Fill::default()],
            borders: vec![Border::default()],
        }
    }

    /// Add a custom font and return its index
    ///
    /// # Examples
    ///
    /// ```rust
    /// use litchi::ooxml::xlsb::writer::StylesWriter;
    /// use litchi::ooxml::xlsb::styles_table::Font;
    ///
    /// let mut styles = StylesWriter::new();
    /// let font = Font {
    ///     name: "Arial".to_string(),
    ///     size: 12.0,
    ///     color: Some(0xFF0000),
    ///     bold: true,
    ///     italic: false,
    ///     underline: false,
    ///     strike: false,
    ///};
    /// let font_id = styles.add_font(font);
    /// ```
    pub fn add_font(&mut self, font: Font) -> usize {
        self.fonts.push(font);
        self.fonts.len() - 1
    }

    /// Add a custom fill and return its index
    pub fn add_fill(&mut self, fill: Fill) -> usize {
        self.fills.push(fill);
        self.fills.len() - 1
    }

    /// Add a custom border and return its index
    pub fn add_border(&mut self, border: Border) -> usize {
        self.borders.push(border);
        self.borders.len() - 1
    }

    /// Write styles to binary format
    ///
    /// The layout is intentionally aligned with the minimal `styles.bin`
    /// produced by SheetJS and Excel:
    ///
    /// ```text
    /// BrtBeginStyleSheet
    ///   BrtBeginFmts / BrtFmt* / BrtEndFmts            (optional, currently empty)
    ///   BrtBeginFonts / BrtFont* / BrtEndFonts         (at least one Calibri font)
    ///   BrtBeginFills / BrtFill* / BrtEndFills         (two fills: none, gray125)
    ///   BrtBeginBorders / BrtBorder* / BrtEndBorders   (one default border)
    ///   BrtBeginCellStyleXFs / BrtXF / BrtEndCellStyleXFs
    ///   BrtBeginCellXFs / BrtXF / BrtEndCellXFs
    ///   BrtBeginStyles / BrtStyle / BrtEndStyles       ("Normal" style)
    ///   BrtBeginDXFs / BrtEndDXFs                      (empty)
    ///   BrtBeginTableStyles / BrtEndTableStyles        (no table styles, only defaults)
    /// BrtEndStyleSheet
    /// ```
    pub(crate) fn write<W: Write>(&self, writer: &mut RecordWriter<W>) -> XlsbResult<()> {
        // BrtBeginStyleSheet
        writer.write_record(record_types::BEGIN_STYLE_SHEET, &[])?;

        // Custom number formats are not currently emitted; we still write an
        // empty FMTS block to match SheetJS' minimal writer layout.
        self.write_default_formats(writer)?;

        // Fonts, fills, borders
        self.write_default_fonts(writer)?;
        self.write_default_fills(writer)?;
        self.write_default_borders(writer)?;

        // Cell style XFs and cell XFs
        self.write_default_cell_style_xfs(writer)?;
        self.write_default_cell_xfs(writer)?;

        // Named styles ("Normal"), DXFs and table styles
        self.write_default_styles(writer)?;
        self.write_default_dxfs(writer)?;
        self.write_default_table_styles(writer)?;

        // BrtEndStyleSheet
        writer.write_record(record_types::END_STYLE_SHEET, &[])?;

        Ok(())
    }

    /// Write default number formats
    fn write_default_formats<W: Write>(&self, writer: &mut RecordWriter<W>) -> XlsbResult<()> {
        let mut count_data = Vec::new();
        let mut temp_writer = RecordWriter::new(&mut count_data);
        temp_writer.write_u32(0)?; // count = 0 (no custom formats)
        writer.write_record(record_types::BEGIN_FMTS, &count_data)?;
        // Built-in formats don't need to be written
        writer.write_record(record_types::END_FMTS, &[])?;
        Ok(())
    }

    /// Write fonts (default and custom)
    fn write_default_fonts<W: Write>(&self, writer: &mut RecordWriter<W>) -> XlsbResult<()> {
        // For now we always serialize at least one default font. Additional
        // fonts added through `add_font` are serialized using a simplified
        // BrtFont layout compatible with SheetJS and Excel.
        let mut count_data = Vec::new();
        let mut temp_writer = RecordWriter::new(&mut count_data);
        let font_count = if self.fonts.is_empty() {
            1u32
        } else {
            self.fonts.len() as u32
        };
        temp_writer.write_u32(font_count)?;
        writer.write_record(record_types::BEGIN_FONTS, &count_data)?;

        if self.fonts.is_empty() {
            Self::write_font_record(writer, &Font::default())?;
        } else {
            for font in &self.fonts {
                Self::write_font_record(writer, font)?;
            }
        }

        writer.write_record(record_types::END_FONTS, &[])?;
        Ok(())
    }

    /// Write fills. To keep the writer simple and fully spec-compliant we
    /// currently emit the two standard fills used by Excel and SheetJS:
    /// `none` (pattern 0) and `gray125` (pattern 17). Custom fills tracked in
    /// the `fills` vector are not yet serialized.
    fn write_default_fills<W: Write>(&self, writer: &mut RecordWriter<W>) -> XlsbResult<()> {
        let mut count_data = Vec::new();
        let mut temp_writer = RecordWriter::new(&mut count_data);
        temp_writer.write_u32(2)?; // exactly two fills: none and gray125
        writer.write_record(record_types::BEGIN_FILLS, &count_data)?;

        // Pattern codes follow the XLSB Fls enumeration as used by SheetJS:
        // 0 = none, 17 = gray125.
        for pattern in [0u32, 17u32] {
            let mut fill_data = Vec::new();
            let mut temp = RecordWriter::new(&mut fill_data);

            // Fls (4 bytes)
            temp.write_u32(pattern)?;

            // BrtColor (FG) - automatic
            temp.write_u32(0)?;
            temp.write_u32(0)?;

            // BrtColor (BG) - automatic
            temp.write_u32(0)?;
            temp.write_u32(0)?;

            // 12 reserved u32 values (gradient / extra fields)
            for _ in 0..12 {
                temp.write_u32(0)?;
            }

            writer.write_record(record_types::FILL, &fill_data)?;
        }

        writer.write_record(record_types::END_FILLS, &[])?;
        Ok(())
    }

    /// Write borders. For now we emit a single default border using the same
    /// payload shape as SheetJS' `write_BrtBorder`, which is sufficient for
    /// default workbooks.
    fn write_default_borders<W: Write>(&self, writer: &mut RecordWriter<W>) -> XlsbResult<()> {
        let mut count_data = Vec::new();
        let mut temp_writer = RecordWriter::new(&mut count_data);
        temp_writer.write_u32(1)?; // one default border
        writer.write_record(record_types::BEGIN_BORDERS, &count_data)?;

        let mut border_data = Vec::new();
        let mut temp = RecordWriter::new(&mut border_data);

        // Diagonal flags (1 byte)
        temp.write_u8(0)?;

        // Five Blxf structures (top, bottom, left, right, diagonal), each:
        // 1 byte dg, 1 byte reserved, 4 bytes color, 4 bytes color.
        for _ in 0..5 {
            temp.write_u8(0)?; // dg
            temp.write_u8(0)?; // reserved
            temp.write_u32(0)?; // color 1
            temp.write_u32(0)?; // color 2
        }

        writer.write_record(record_types::BORDER, &border_data)?;
        writer.write_record(record_types::END_BORDERS, &[])?;
        Ok(())
    }

    /// Write default cell style XFs (BrtBeginCellStyleXFs / BrtXF / BrtEndCellStyleXFs).
    fn write_default_cell_style_xfs<W: Write>(
        &self,
        writer: &mut RecordWriter<W>,
    ) -> XlsbResult<()> {
        let mut count_data = Vec::new();
        let mut temp_writer = RecordWriter::new(&mut count_data);
        temp_writer.write_u32(1)?; // one style XF
        writer.write_record(record_types::BEGIN_CELL_STYLE_XFS, &count_data)?;

        // Parent index 0xFFFF to indicate no parent XF, matching SheetJS.
        Self::write_xf_record(writer, 0xFFFF)?;

        writer.write_record(record_types::END_CELL_STYLE_XFS, &[])?;
        Ok(())
    }

    /// Write default cell XFs (cell formats)
    fn write_default_cell_xfs<W: Write>(&self, writer: &mut RecordWriter<W>) -> XlsbResult<()> {
        let mut count_data = Vec::new();
        let mut temp_writer = RecordWriter::new(&mut count_data);
        temp_writer.write_u32(1)?; // one default XF
        writer.write_record(record_types::BEGIN_CELL_XFS, &count_data)?;

        // Parent index 0 for the normal cell XF.
        Self::write_xf_record(writer, 0)?;

        writer.write_record(record_types::END_CELL_XFS, &[])?;
        Ok(())
    }

    /// Write default styles table (BrtBeginStyles / BrtStyle / BrtEndStyles).
    fn write_default_styles<W: Write>(&self, writer: &mut RecordWriter<W>) -> XlsbResult<()> {
        let mut count_data = Vec::new();
        let mut temp_writer = RecordWriter::new(&mut count_data);
        temp_writer.write_u32(1)?; // one style: Normal
        writer.write_record(record_types::BEGIN_STYLES, &count_data)?;

        let mut style_data = Vec::new();
        let mut sw = RecordWriter::new(&mut style_data);

        // xfId (u32)
        sw.write_u32(0)?;
        // iUsageCount (u16) - minimal value observed in SheetJS
        sw.write_u16(1)?;
        // builtinId (u8) and iLevel (u8)
        sw.write_u8(0)?; // builtinId: Normal
        sw.write_u8(0)?; // iLevel
        // Name as XLWideString; for non-empty strings this matches
        // XLNullableWideString payload.
        sw.write_wide_string("Normal")?;

        writer.write_record(record_types::STYLE, &style_data)?;
        writer.write_record(record_types::END_STYLES, &[])?;
        Ok(())
    }

    /// Write an empty DXF table (BrtBeginDXFs / BrtEndDXFs).
    fn write_default_dxfs<W: Write>(&self, writer: &mut RecordWriter<W>) -> XlsbResult<()> {
        let mut data = Vec::new();
        let mut temp_writer = RecordWriter::new(&mut data);
        temp_writer.write_u32(0)?; // no DXFs
        writer.write_record(record_types::BEGIN_DXFS, &data)?;
        writer.write_record(record_types::END_DXFS, &[])?;
        Ok(())
    }

    /// Write minimal table styles (BrtBeginTableStyles / BrtEndTableStyles) with
    /// zero styles but default names matching SheetJS and Excel.
    fn write_default_table_styles<W: Write>(&self, writer: &mut RecordWriter<W>) -> XlsbResult<()> {
        let mut data = Vec::new();
        let mut temp_writer = RecordWriter::new(&mut data);
        temp_writer.write_u32(0)?; // cnt = 0
        temp_writer.write_wide_string("TableStyleMedium9")?;
        temp_writer.write_wide_string("PivotStyleMedium4")?;
        writer.write_record(record_types::BEGIN_TABLE_STYLES, &data)?;
        writer.write_record(record_types::END_TABLE_STYLES, &[])?;
        Ok(())
    }

    /// Helper to write a minimal BrtXF payload identical to SheetJS' writer.
    ///
    /// Layout (16 bytes):
    ///
    /// ```text
    /// u16 ixfeParent
    /// u16 numFmtId
    /// u16 iFont
    /// u16 iFill
    /// u16 ixBorder
    /// u8  trot
    /// u8  indent
    /// u8  flags
    /// u8  flags
    /// u8  xfGrbitAtr
    /// u8  reserved
    /// ```
    fn write_xf_record<W: Write>(writer: &mut RecordWriter<W>, ixfe_parent: u16) -> XlsbResult<()> {
        let mut xf_data = Vec::new();
        let mut temp_writer = RecordWriter::new(&mut xf_data);

        temp_writer.write_u16(ixfe_parent)?;
        temp_writer.write_u16(0)?; // numFmtId
        temp_writer.write_u16(0)?; // iFont
        temp_writer.write_u16(0)?; // iFill
        temp_writer.write_u16(0)?; // ixBorder
        temp_writer.write_u8(0)?; // trot
        temp_writer.write_u8(0)?; // indent
        temp_writer.write_u8(0)?; // flags
        temp_writer.write_u8(0)?; // flags
        temp_writer.write_u8(0)?; // xfGrbitAtr
        temp_writer.write_u8(0)?; // reserved

        writer.write_record(record_types::XF, &xf_data)?;
        Ok(())
    }

    /// Helper to write a BrtFont payload compatible with the simplified Font
    /// structure used by this crate. This closely follows SheetJS'
    /// `write_BrtFont` implementation.
    fn write_font_record<W: Write>(writer: &mut RecordWriter<W>, font: &Font) -> XlsbResult<()> {
        let mut font_data = Vec::new();
        let mut temp_writer = RecordWriter::new(&mut font_data);

        // Height in twips (size * 20)
        let height = (font.size * 20.0).round() as u16;
        temp_writer.write_u16(height)?;

        // FontFlags (2 bytes) – we currently map italic and strikeout only.
        let mut grbit: u8 = 0;
        if font.italic {
            grbit |= 0x02;
        }
        if font.strike {
            grbit |= 0x08;
        }
        temp_writer.write_u8(grbit)?;
        temp_writer.write_u8(0)?; // reserved

        // Weight: 0x02BC for bold, 0x0190 for normal.
        let weight = if font.bold { 0x02BC } else { 0x0190 };
        temp_writer.write_u16(weight)?;

        // Vertical alignment (baseline).
        temp_writer.write_u16(0)?;

        // Underline style: 1 = single underline, 0 = none.
        temp_writer.write_u8(if font.underline { 1 } else { 0 })?;

        // Family (2 = Swiss, matches Calibri in the default theme).
        temp_writer.write_u8(2)?;

        // Charset and reserved byte.
        temp_writer.write_u8(0)?; // charset
        temp_writer.write_u8(0)?; // reserved

        // BrtColor: we currently always emit automatic color (8 zero bytes).
        // TODO: honor `font.color` by mapping ARGB to BrtColor when needed.
        temp_writer.write_u32(0)?;
        temp_writer.write_u32(0)?;

        // Scheme: 2 = minor (Calibri in the default theme).
        temp_writer.write_u8(2)?;

        // Font name as XLWideString.
        temp_writer.write_wide_string(&font.name)?;

        writer.write_record(record_types::FONT, &font_data)?;
        Ok(())
    }
}

impl Default for StylesWriter {
    fn default() -> Self {
        Self::new()
    }
}
