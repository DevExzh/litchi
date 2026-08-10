#![cfg_attr(
    test,
    allow(
        clippy::float_cmp,
        reason = "style codec tests compare exact finite values just decoded from their own wire literals"
    )
)]
#![allow(
    clippy::cast_possible_truncation,
    clippy::cast_sign_loss,
    reason = "legacy module confines validated BIFF12 field narrowing or exact signed-bit reinterpretation to this codec boundary"
)]

//! Styles writer for XLSB format
//!
//! This module provides functionality to write style information (fonts, fills, borders, etc.)
//! to XLSB binary format files.

use crate::package::error::{Error, Result};
use crate::package::styles::{Border, BorderSide};
use crate::package::styles_table::{Fill, Font};
use crate::raw::Writer;
use crate::raw::kind;
use std::io::Write;

/// Differential formatting style for conditional formatting rules.
///
/// Each DXF is serialized as a `BrtDXF` record in styles.bin, referenced
/// by `dxf_id` in `BrtBeginCFRule`.
#[derive(Debug, Clone)]
pub struct DxfStyle {
    /// Fill foreground color (ARGB, e.g. `0xFFFF0000` for red).
    /// When set, the DXF applies a solid fill with this color.
    pub fill_fg_color: Option<u32>,
}

/// Styles writer for XLSB with support for custom fonts, fills, borders, and DXFs
pub struct StylesWriter {
    /// Custom fonts (index 0 is the default font)
    fonts: Vec<Font>,
    /// Custom fills (indices 0-1 are default fills)
    fills: Vec<Fill>,
    /// Custom borders (index 0 is the default border)
    borders: Vec<Border>,
    /// Differential formatting styles referenced by conditional formatting rules
    dxfs: Vec<DxfStyle>,
}

impl StylesWriter {
    /// Create a new styles writer with default styles
    pub fn new() -> Self {
        StylesWriter {
            fonts: vec![Font::default()],
            fills: vec![
                Fill::default(),
                Fill {
                    pattern_type: 17,
                    ..Fill::default()
                },
            ],
            borders: vec![Border::default()],
            dxfs: Vec::new(),
        }
    }

    /// Add a custom font and return its index
    ///
    /// # Examples
    ///
    /// ```ignore
    /// use litchi_xlsb::writer::StylesWriter;
    /// use litchi_xlsb::package::styles_table::Font;
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

    /// Add a differential formatting style (DXF) with a solid fill color and
    /// return its 0-based index.
    ///
    /// The returned index is used as `dxf_id` in
    /// [`Rule`](crate::conditional_formatting::Rule).
    ///
    /// # Arguments
    ///
    /// * `fill_color` — ARGB fill color (e.g. `0xFFFF0000` for opaque red)
    ///
    /// # Examples
    ///
    /// ```rust,ignore
    /// let dxf_idx = workbook.styles_mut().add_dxf_fill(0xFFFF0000);
    /// rule.dxf_id = Some(dxf_idx as u32);
    /// ```
    pub fn add_dxf_fill(&mut self, fill_color: u32) -> usize {
        self.dxfs.push(DxfStyle {
            fill_fg_color: Some(fill_color),
        });
        self.dxfs.len() - 1
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
    pub(crate) fn write<W: Write>(&self, writer: &mut Writer<W>) -> Result<()> {
        // BrtBeginStyleSheet
        writer.write_record(kind::BEGIN_STYLE_SHEET, &[])?;

        // Custom number formats are not currently emitted; we still write an
        // empty FMTS block to match SheetJS' minimal writer layout.
        self.write_default_formats(writer)?;

        // Fonts, fills, borders
        self.write_default_fonts(writer)?;
        self.write_fills(writer)?;
        self.write_borders(writer)?;

        // Cell style XFs and cell XFs
        self.write_default_cell_style_xfs(writer)?;
        self.write_default_cell_xfs(writer)?;

        // Named styles ("Normal"), DXFs and table styles
        self.write_default_styles(writer)?;
        self.write_default_dxfs(writer)?;
        self.write_default_table_styles(writer)?;

        // BrtEndStyleSheet
        writer.write_record(kind::END_STYLE_SHEET, &[])?;

        Ok(())
    }

    /// Write default number formats
    fn write_default_formats<W: Write>(&self, writer: &mut Writer<W>) -> Result<()> {
        let mut count_data = Vec::new();
        let mut temp_writer = Writer::new(&mut count_data);
        temp_writer.write_u32(0)?; // count = 0 (no custom formats)
        writer.write_record(kind::BEGIN_FMTS, &count_data)?;
        // Built-in formats don't need to be written
        writer.write_record(kind::END_FMTS, &[])?;
        Ok(())
    }

    /// Write fonts (default and custom)
    fn write_default_fonts<W: Write>(&self, writer: &mut Writer<W>) -> Result<()> {
        // For now we always serialize at least one default font. Additional
        // fonts added through `add_font` are serialized using a simplified
        // BrtFont layout used by SheetJS and Excel.
        let mut count_data = Vec::new();
        let mut temp_writer = Writer::new(&mut count_data);
        let font_count = if self.fonts.is_empty() {
            1u32
        } else {
            self.fonts.len() as u32
        };
        temp_writer.write_u32(font_count)?;
        writer.write_record(kind::BEGIN_FONTS, &count_data)?;

        if self.fonts.is_empty() {
            Self::write_font_record(writer, &Font::default())?;
        } else {
            for font in &self.fonts {
                Self::write_font_record(writer, font)?;
            }
        }

        writer.write_record(kind::END_FONTS, &[])?;
        Ok(())
    }

    /// Write the standard and user-added pattern fills.
    fn write_fills<W: Write>(&self, writer: &mut Writer<W>) -> Result<()> {
        let mut count_data = Vec::new();
        let mut temp_writer = Writer::new(&mut count_data);
        temp_writer.write_u32(self.fills.len() as u32)?;
        writer.write_record(kind::BEGIN_FILLS, &count_data)?;

        for fill in &self.fills {
            let fill_data = Self::encode_fill_payload(fill)?;
            writer.write_record(kind::FILL, &fill_data)?;
        }

        writer.write_record(kind::END_FILLS, &[])?;
        Ok(())
    }

    /// Write cell borders as `BrtBorder` records.
    fn write_borders<W: Write>(&self, writer: &mut Writer<W>) -> Result<()> {
        let mut count_data = Vec::new();
        let mut temp_writer = Writer::new(&mut count_data);
        temp_writer.write_u32(self.borders.len() as u32)?;
        writer.write_record(kind::BEGIN_BORDERS, &count_data)?;

        for border in &self.borders {
            let border_data = Self::encode_border_payload(border)?;
            writer.write_record(kind::BORDER, &border_data)?;
        }

        writer.write_record(kind::END_BORDERS, &[])?;
        Ok(())
    }

    fn write_blxf<W: Write>(writer: &mut Writer<W>, side: Option<&BorderSide>) -> Result<()> {
        if let Some(side) = side {
            writer.write_u8(side.style as u8)?;
            writer.write_u8(0)?;
            Self::write_brt_color(writer, side.color)?;
        } else {
            writer.write_u16(0)?;
            Self::write_brt_color(writer, None)?;
        }
        Ok(())
    }

    /// Write default cell style XFs (BrtBeginCellStyleXFs / BrtXF / BrtEndCellStyleXFs).
    fn write_default_cell_style_xfs<W: Write>(&self, writer: &mut Writer<W>) -> Result<()> {
        let mut count_data = Vec::new();
        let mut temp_writer = Writer::new(&mut count_data);
        temp_writer.write_u32(1)?; // one style XF
        writer.write_record(kind::BEGIN_CELL_STYLE_XFS, &count_data)?;

        // Parent index 0xFFFF to indicate no parent XF, matching SheetJS.
        Self::write_xf_record(writer, 0xFFFF)?;

        writer.write_record(kind::END_CELL_STYLE_XFS, &[])?;
        Ok(())
    }

    /// Write default cell XFs (cell formats)
    fn write_default_cell_xfs<W: Write>(&self, writer: &mut Writer<W>) -> Result<()> {
        let mut count_data = Vec::new();
        let mut temp_writer = Writer::new(&mut count_data);
        temp_writer.write_u32(1)?; // one default XF
        writer.write_record(kind::BEGIN_CELL_XFS, &count_data)?;

        // Parent index 0 for the normal cell XF.
        Self::write_xf_record(writer, 0)?;

        writer.write_record(kind::END_CELL_XFS, &[])?;
        Ok(())
    }

    /// Write default styles table (BrtBeginStyles / BrtStyle / BrtEndStyles).
    fn write_default_styles<W: Write>(&self, writer: &mut Writer<W>) -> Result<()> {
        let mut count_data = Vec::new();
        let mut temp_writer = Writer::new(&mut count_data);
        temp_writer.write_u32(1)?; // one style: Normal
        writer.write_record(kind::BEGIN_STYLES, &count_data)?;

        let mut style_data = Vec::new();
        let mut sw = Writer::new(&mut style_data);

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

        writer.write_record(kind::STYLE, &style_data)?;
        writer.write_record(kind::END_STYLES, &[])?;
        Ok(())
    }

    /// Write DXF table (`BrtBeginDXFs` / `BrtDXF`* / `BrtEndDXFs`).
    ///
    /// Each [`DxfStyle`] is serialized as a `BrtDXF` record per [MS-XLSB] 2.4.356.
    fn write_default_dxfs<W: Write>(&self, writer: &mut Writer<W>) -> Result<()> {
        // BrtBeginDXFs: count(u32)
        let mut data = Vec::new();
        let mut temp_writer = Writer::new(&mut data);
        temp_writer.write_u32(self.dxfs.len() as u32)?;
        writer.write_record(kind::BEGIN_DXFS, &data)?;

        for dxf in &self.dxfs {
            let payload = Self::serialize_dxf(dxf);
            writer.write_record(kind::DXF, &payload)?;
        }

        writer.write_record(kind::END_DXFS, &[])?;
        Ok(())
    }

    /// Serialize a single [`DxfStyle`] into a `BrtDXF` payload.
    ///
    /// Layout per [MS-XLSB] 2.4.356:
    ///   `flags(u16)` + `XFProps(reserved(u16) + cprops(u16) + xfPropArray)`
    fn serialize_dxf(dxf: &DxfStyle) -> Vec<u8> {
        let mut buf = Vec::with_capacity(32);

        // BrtDXF flags: unused(15 bits) + fNewBorder(1 bit) = 0
        buf.extend_from_slice(&0u16.to_le_bytes());

        // XFProps structure
        let mut props: Vec<Vec<u8>> = Vec::new();

        if let Some(color) = dxf.fill_fg_color {
            // XFProp 0x0000: FillPattern = FLSSOLID (1)
            props.push(Self::make_xf_prop(0x0000, &[0x01]));
            // XFProp 0x0002: background XFPropColor (RGBA)
            // CF DXFs use bgColor as the cell fill color per Excel/LO convention:
            // LO finalizeImport() swaps bgColor → patternColor when pattern is solid.
            props.push(Self::make_xf_prop(
                0x0002,
                &Self::make_xf_prop_color_rgba(color),
            ));
        }

        // XFProps: reserved(u16=0) + cprops(u16) + xfPropArray
        buf.extend_from_slice(&0u16.to_le_bytes()); // reserved
        buf.extend_from_slice(&(props.len() as u16).to_le_bytes()); // cprops
        for prop in &props {
            buf.extend_from_slice(prop);
        }

        buf
    }

    /// Build a single XFProp: `xfPropType(u16) + cb(u16) + xfPropDataBlob`.
    ///
    /// `cb` is the total size of this XFProp structure (4-byte header + data).
    fn make_xf_prop(prop_type: u16, data: &[u8]) -> Vec<u8> {
        let cb = (4 + data.len()) as u16;
        let mut buf = Vec::with_capacity(cb as usize);
        buf.extend_from_slice(&prop_type.to_le_bytes());
        buf.extend_from_slice(&cb.to_le_bytes());
        buf.extend_from_slice(data);
        buf
    }

    /// Build an `XFPropColor` blob (8 bytes) for a direct RGBA color.
    ///
    /// Layout per [MS-XLSB] 2.5.161:
    ///   `fValidRGBA(1 bit) + xclrType(7 bits)` + `icv(u8)` + `nTintShade(i16)` + `dwRgba(u32)`
    fn make_xf_prop_color_rgba(argb: u32) -> [u8; 8] {
        let mut buf = [0u8; 8];
        // byte 0: fValidRGBA=0 | xclrType=0x02 (RGBA) → (0x02 << 1) | 0 = 0x04
        buf[0] = 0x04;
        // byte 1: icv = 0 (ignored for RGBA type)
        buf[1] = 0;
        // bytes 2-3: nTintShade = 0 (no tint)
        buf[2] = 0;
        buf[3] = 0;
        // bytes 4-7: dwRgba in RGBA byte order.
        // ARGB input: 0xAARRGGBB → dwRgba expects RGBA: 0xRRGGBBAA
        let a = ((argb >> 24) & 0xFF) as u8;
        let r = ((argb >> 16) & 0xFF) as u8;
        let g = ((argb >> 8) & 0xFF) as u8;
        let b = (argb & 0xFF) as u8;
        buf[4] = r;
        buf[5] = g;
        buf[6] = b;
        buf[7] = a;
        buf
    }

    /// Write minimal table styles (BrtBeginTableStyles / BrtEndTableStyles) with
    /// zero styles but default names matching SheetJS and Excel.
    fn write_default_table_styles<W: Write>(&self, writer: &mut Writer<W>) -> Result<()> {
        let mut data = Vec::new();
        let mut temp_writer = Writer::new(&mut data);
        temp_writer.write_u32(0)?; // cnt = 0
        temp_writer.write_wide_string("TableStyleMedium9")?;
        temp_writer.write_wide_string("PivotStyleMedium4")?;
        writer.write_record(kind::BEGIN_TABLE_STYLES, &data)?;
        writer.write_record(kind::END_TABLE_STYLES, &[])?;
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
    fn write_xf_record<W: Write>(writer: &mut Writer<W>, ixfe_parent: u16) -> Result<()> {
        let mut xf_data = Vec::new();
        let mut temp_writer = Writer::new(&mut xf_data);

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

        writer.write_record(kind::XF, &xf_data)?;
        Ok(())
    }

    /// Helper to write a BrtFont payload for the simplified Font structure
    /// used by this crate. This closely follows SheetJS'
    /// `write_BrtFont` implementation.
    pub(crate) fn encode_font_payload(font: &Font) -> Result<Vec<u8>> {
        let height = (font.size * 20.0).round();
        if !font.size.is_finite() || !(20.0..=8191.0).contains(&height) {
            return Err(Error::Unrecognized {
                typ: "BrtFont height".to_string(),
                val: font.size.to_string(),
            });
        }
        let name_len = font.name.encode_utf16().count();
        if !(1..=31).contains(&name_len) {
            return Err(Error::Unrecognized {
                typ: "BrtFont name length".to_string(),
                val: name_len.to_string(),
            });
        }

        let mut font_data = Vec::new();
        let mut temp_writer = Writer::new(&mut font_data);

        // Height in twips (size * 20)
        temp_writer.write_u16(height as u16)?;

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

        Self::write_brt_color(&mut temp_writer, font.color)?;

        // Scheme: 2 = minor (Calibri in the default theme).
        temp_writer.write_u8(2)?;

        // Font name as XLWideString.
        temp_writer.write_wide_string(&font.name)?;

        Ok(font_data)
    }

    fn write_font_record<W: Write>(writer: &mut Writer<W>, font: &Font) -> Result<()> {
        writer.write_record(kind::FONT, &Self::encode_font_payload(font)?)?;
        Ok(())
    }

    pub(crate) fn encode_fill_payload(fill: &Fill) -> Result<Vec<u8>> {
        if !matches!(fill.pattern_type, 0..=18) {
            return Err(Error::UnsupportedFeature(format!(
                "XLSB fill pattern {} is not a supported pattern fill",
                fill.pattern_type
            )));
        }
        let mut data = Vec::new();
        let mut writer = Writer::new(&mut data);
        writer.write_u32(fill.pattern_type)?;
        Self::write_brt_color(&mut writer, fill.fg_color)?;
        Self::write_brt_color(&mut writer, fill.bg_color)?;
        for _ in 0..12 {
            writer.write_u32(0)?;
        }
        Ok(data)
    }

    pub(crate) fn encode_border_payload(border: &Border) -> Result<Vec<u8>> {
        if border.vertical.is_some() || border.horizontal.is_some() {
            return Err(Error::UnsupportedFeature(
                "vertical and horizontal table borders are not BrtBorder fields".to_string(),
            ));
        }
        if (border.diagonal_down || border.diagonal_up) && border.diagonal.is_none() {
            return Err(Error::Unrecognized {
                typ: "BrtBorder diagonal".to_string(),
                val: "direction is set without a diagonal border".to_string(),
            });
        }
        let mut data = Vec::new();
        let mut writer = Writer::new(&mut data);
        writer.write_u8(u8::from(border.diagonal_down) | (u8::from(border.diagonal_up) << 1))?;
        Self::write_blxf(&mut writer, border.top.as_ref())?;
        Self::write_blxf(&mut writer, border.bottom.as_ref())?;
        Self::write_blxf(&mut writer, border.left.as_ref())?;
        Self::write_blxf(&mut writer, border.right.as_ref())?;
        Self::write_blxf(&mut writer, border.diagonal.as_ref())?;
        Ok(data)
    }

    /// Write an automatic or direct ARGB `BrtColor`.
    fn write_brt_color<W: Write>(writer: &mut Writer<W>, color: Option<u32>) -> Result<()> {
        if let Some(argb) = color {
            // fValidRGB=1, xColorType=2 (direct ARGB).
            writer.write_u8((2 << 1) | 1)?;
            writer.write_u8(0)?; // index is ignored for direct colors
            writer.write_u16(0)?; // no tint or shade
            writer.write_u8(((argb >> 16) & 0xFF) as u8)?;
            writer.write_u8(((argb >> 8) & 0xFF) as u8)?;
            writer.write_u8((argb & 0xFF) as u8)?;
            writer.write_u8(((argb >> 24) & 0xFF) as u8)?;
        } else {
            writer.write_u32(0)?;
            writer.write_u32(0)?;
        }
        Ok(())
    }
}

impl Default for StylesWriter {
    fn default() -> Self {
        Self::new()
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::package::styles::BorderStyle;
    use crate::package::styles_table::StylesTable;

    #[test]
    fn test_styles_writer_new() {
        let writer = StylesWriter::new();
        assert_eq!(writer.fonts.len(), 1); // Default font
        assert_eq!(writer.fills.len(), 2); // Two default fills
        assert_eq!(writer.borders.len(), 1); // Default border
        assert!(writer.dxfs.is_empty());
    }

    #[test]
    fn test_styles_writer_default() {
        let writer: StylesWriter = Default::default();
        assert_eq!(writer.fonts.len(), 1);
        assert_eq!(writer.fills.len(), 2);
    }

    #[test]
    fn test_add_font() {
        let mut writer = StylesWriter::new();
        let font = Font {
            name: "Arial".to_string(),
            size: 12.0,
            color: None,
            bold: true,
            italic: false,
            underline: false,
            strike: false,
        };
        let idx = writer.add_font(font);
        assert_eq!(idx, 1); // Index after default font
        assert_eq!(writer.fonts.len(), 2);
    }

    #[test]
    fn test_add_fill() {
        let mut writer = StylesWriter::new();
        let fill = Fill::default();
        let idx = writer.add_fill(fill);
        assert_eq!(idx, 2); // Index after default fills
        assert_eq!(writer.fills.len(), 3);
    }

    #[test]
    fn writes_and_reads_direct_font_and_pattern_fill_colors() {
        let mut styles = StylesWriter::new();
        styles.add_font(Font {
            name: "Arial".to_string(),
            size: 12.0,
            color: Some(0x80402010),
            bold: true,
            italic: true,
            underline: true,
            strike: true,
        });
        styles.add_fill(Fill {
            pattern_type: 1,
            fg_color: Some(0xFFFF0000),
            bg_color: Some(0x4000FF00),
        });
        styles.add_border(Border {
            top: Some(BorderSide {
                style: BorderStyle::Thin,
                color: Some(0xFF112233),
            }),
            diagonal: Some(BorderSide {
                style: BorderStyle::Double,
                color: Some(0x80402010),
            }),
            diagonal_down: true,
            ..Border::default()
        });

        let mut bytes = Vec::new();
        styles.write(&mut Writer::new(&mut bytes)).unwrap();
        let parsed = StylesTable::from_bytes(&bytes).unwrap();

        let font = parsed.get_font(1).unwrap();
        assert_eq!(font.name, "Arial");
        assert_eq!(font.size, 12.0);
        assert_eq!(font.color, Some(0x80402010));
        assert!(font.bold);
        assert!(font.italic);
        assert!(font.underline);
        assert!(font.strike);

        assert_eq!(parsed.get_fill(1).unwrap().pattern_type, 17);
        let fill = parsed.get_fill(2).unwrap();
        assert_eq!(fill.pattern_type, 1);
        assert_eq!(fill.fg_color, Some(0xFFFF0000));
        assert_eq!(fill.bg_color, Some(0x4000FF00));

        let border = parsed.get_border(1).unwrap();
        assert_eq!(border.top.as_ref().unwrap().style, BorderStyle::Thin);
        assert_eq!(border.top.as_ref().unwrap().color, Some(0xFF112233));
        assert_eq!(border.diagonal.as_ref().unwrap().style, BorderStyle::Double);
        assert_eq!(border.diagonal.as_ref().unwrap().color, Some(0x80402010));
        assert!(border.diagonal_down);
        assert!(!border.diagonal_up);
    }

    #[test]
    fn direct_brt_color_uses_rgba_byte_order() {
        let mut bytes = Vec::new();
        StylesWriter::write_brt_color(&mut Writer::new(&mut bytes), Some(0x80402010)).unwrap();
        assert_eq!(bytes, [5, 0, 0, 0, 0x40, 0x20, 0x10, 0x80]);
    }

    #[test]
    fn test_add_border() {
        let mut writer = StylesWriter::new();
        let border = Border::default();
        let idx = writer.add_border(border);
        assert_eq!(idx, 1); // Index after default border
        assert_eq!(writer.borders.len(), 2);
    }

    #[test]
    fn test_add_dxf_fill() {
        let mut writer = StylesWriter::new();
        let idx1 = writer.add_dxf_fill(0xFFFF0000); // Red
        let idx2 = writer.add_dxf_fill(0xFF00FF00); // Green

        assert_eq!(idx1, 0);
        assert_eq!(idx2, 1);
        assert_eq!(writer.dxfs.len(), 2);
        assert_eq!(writer.dxfs[0].fill_fg_color, Some(0xFFFF0000));
        assert_eq!(writer.dxfs[1].fill_fg_color, Some(0xFF00FF00));
    }

    #[test]
    fn test_serialize_dxf_empty() {
        let dxf = DxfStyle {
            fill_fg_color: None,
        };
        let payload = StylesWriter::serialize_dxf(&dxf);
        assert!(!payload.is_empty());
        // Should have flags (2 bytes) + XFProps header (4 bytes)
        assert_eq!(payload.len(), 6);
    }

    #[test]
    fn test_serialize_dxf_with_color() {
        let dxf = DxfStyle {
            fill_fg_color: Some(0xFFFF0000), // Red
        };
        let payload = StylesWriter::serialize_dxf(&dxf);
        assert!(!payload.is_empty());
        // Should include fill pattern prop and color prop
    }

    #[test]
    fn test_make_xf_prop() {
        let prop = StylesWriter::make_xf_prop(0x0000, &[0x01]);
        assert_eq!(prop.len(), 5); // 2 bytes type + 2 bytes cb + 1 byte data
    }

    #[test]
    fn test_make_xf_prop_color_rgba() {
        let color = StylesWriter::make_xf_prop_color_rgba(0xFFFF0000); // ARGB red
        assert_eq!(color.len(), 8);
        // Check RGBA byte order (RRGGBBAA)
        assert_eq!(color[4], 0xFF); // R
        assert_eq!(color[5], 0x00); // G
        assert_eq!(color[6], 0x00); // B
        assert_eq!(color[7], 0xFF); // A
    }

    #[test]
    fn test_dxf_style_clone() {
        let dxf = DxfStyle {
            fill_fg_color: Some(0xFF00FF00),
        };
        let cloned = dxf.clone();
        assert_eq!(cloned.fill_fg_color, dxf.fill_fg_color);
    }

    #[test]
    fn test_dxf_style_debug() {
        let dxf = DxfStyle {
            fill_fg_color: Some(0xFF0000FF),
        };
        let debug_str = format!("{:?}", dxf);
        assert!(debug_str.contains("DxfStyle"));
    }
}
