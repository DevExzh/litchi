//! Strict BIFF12 styles-part codec.
//!
//! The record layouts are the checked-in [MS-XLSB] Brt* definitions. Host
//! package and compatibility concerns stay outside this module.

use super::model::{
    Alignment, Border, BorderSide, BorderStyle, CellFormat, Fill, Font, HorizontalAlignment, Table,
    VerticalAlignment,
};
use crate::leniency::{LayoutDefect, ReservedFieldDefect, Tolerance, ToleranceReport};
use crate::raw::{Cursor, Records, kind};
use std::collections::TryReserveError;
use thiserror::Error;

/// Result returned by the styles codec.
pub type Result<T> = std::result::Result<T, Error>;

/// Error returned by the bounded styles codec.
#[derive(Debug, Error)]
#[non_exhaustive]
pub enum Error {
    /// A BIFF12 header, payload, scalar, or string failed wire validation.
    #[error(transparent)]
    Wire(#[from] crate::raw::Error),
    /// A fixed-width record field was truncated.
    #[error("invalid length: expected {expected}, found {found}")]
    InvalidLength { expected: usize, found: usize },
    /// A well-framed field violated its checked-in domain.
    #[error("unrecognized {typ}: {val}")]
    Unrecognized { typ: String, val: String },
    /// A bounded collection could not reserve the required memory.
    #[error("allocation failed for {resource}: {source}")]
    Allocation {
        resource: &'static str,
        source: TryReserveError,
    },
}

/// Read one complete XLSB styles stream, refusing every defect.
///
/// This standalone codec function cannot report repairs, so it rejects the
/// deviations an opened workbook repairs and records; see
/// [`crate::leniency`].
pub fn read(bytes: &[u8]) -> Result<Table> {
    read_with(bytes, Tolerance::refusing())
}

/// Read one complete XLSB styles stream, recording every Excel-compatible
/// repair into `report`.
pub(crate) fn read_with_report(bytes: &[u8], report: &mut ToleranceReport) -> Result<Table> {
    read_with(bytes, Tolerance::recording(report))
}

fn read_with(bytes: &[u8], mut tolerance: Tolerance<'_>) -> Result<Table> {
    let mut table = Table::default();
    let mut in_fonts = false;
    let mut in_fills = false;
    let mut in_borders = false;
    let mut in_fmts = false;
    let mut in_cell_xfs = false;
    let mut in_cell_style_xfs = false;
    let mut fmt_ordinal = 0u32;

    for record in Records::new(bytes) {
        let record = record?;
        let data = record.payload();
        match record.kind() {
            kind::BEGIN_FONTS => {
                in_fonts = true;
                table.fonts.clear();
            },
            kind::END_FONTS => in_fonts = false,
            kind::FONT if in_fonts => {
                reserve(&mut table.fonts, 1, "styles fonts")?;
                let ordinal = ordinal(table.fonts.len());
                table
                    .fonts
                    .push(parse_font_with(data, ordinal, tolerance.reborrow())?);
            },
            kind::BEGIN_FILLS => {
                in_fills = true;
                table.fills.clear();
            },
            kind::END_FILLS => in_fills = false,
            kind::FILL if in_fills => {
                reserve(&mut table.fills, 1, "styles fills")?;
                let ordinal = ordinal(table.fills.len());
                table
                    .fills
                    .push(parse_fill_with(data, ordinal, tolerance.reborrow())?);
            },
            kind::BEGIN_BORDERS => {
                in_borders = true;
                table.borders.clear();
            },
            kind::END_BORDERS => in_borders = false,
            kind::BORDER if in_borders => {
                reserve(&mut table.borders, 1, "styles borders")?;
                let ordinal = ordinal(table.borders.len());
                table
                    .borders
                    .push(parse_border_with(data, ordinal, tolerance.reborrow())?);
            },
            kind::BEGIN_FMTS => in_fmts = true,
            kind::END_FMTS => in_fmts = false,
            kind::FMT if in_fmts => {
                let (id, format_code) =
                    parse_num_fmt_with(data, fmt_ordinal, tolerance.reborrow())?;
                fmt_ordinal = fmt_ordinal.saturating_add(1);
                table
                    .num_fmts
                    .try_reserve(1)
                    .map_err(|source| allocation("styles number formats", source))?;
                table.num_fmts.insert(id, format_code);
            },
            kind::BEGIN_CELL_XFS => {
                in_cell_xfs = true;
                table.cell_xfs.clear();
            },
            kind::END_CELL_XFS => in_cell_xfs = false,
            kind::XF if in_cell_xfs => {
                reserve(&mut table.cell_xfs, 1, "styles cell formats")?;
                table.cell_xfs.push(parse_cell_format(data)?);
            },
            kind::BEGIN_CELL_STYLE_XFS => {
                in_cell_style_xfs = true;
                table.cell_style_xfs.clear();
            },
            kind::END_CELL_STYLE_XFS => in_cell_style_xfs = false,
            kind::XF if in_cell_style_xfs => {
                reserve(&mut table.cell_style_xfs, 1, "styles cell-style formats")?;
                table.cell_style_xfs.push(parse_cell_format(data)?);
            },
            _ => {},
        }
    }

    Ok(table)
}

/// A collection ordinal for the tolerance report, saturated to `u32`.
fn ordinal(index: usize) -> u32 {
    u32::try_from(index).unwrap_or(u32::MAX)
}

/// Parse a BrtFont payload.
pub fn parse_font(data: &[u8]) -> Result<Font> {
    parse_font_with(data, 0, Tolerance::refusing())
}

fn parse_font_with(data: &[u8], ordinal: u32, mut tolerance: Tolerance<'_>) -> Result<Font> {
    const FONT_NAME_OFFSET: usize = 21;
    if data.len() < FONT_NAME_OFFSET {
        return Err(Error::InvalidLength {
            expected: FONT_NAME_OFFSET,
            found: data.len(),
        });
    }

    let height = read_u16(data, 0)?;
    let size = f64::from(height) / 20.0;
    let flags = read_u16(data, 2)?;
    let italic = flags & 0x0002 != 0;
    let strike = flags & 0x0008 != 0;
    let bold = read_u16(data, 4)? >= 0x02BC;
    let underline = data[8] != 0;
    let color = direct_color(data, 12, kind::FONT, ordinal, &mut tolerance)?;
    let (name, _) = decode_string(&data[FONT_NAME_OFFSET..])?;

    Ok(Font {
        size,
        name,
        bold,
        italic,
        underline,
        strike,
        color,
    })
}

/// Parse a BrtFill payload.
pub fn parse_fill(data: &[u8]) -> Result<Fill> {
    parse_fill_with(data, 0, Tolerance::refusing())
}

fn parse_fill_with(data: &[u8], ordinal: u32, mut tolerance: Tolerance<'_>) -> Result<Fill> {
    const FIXED_FILL_SIZE: usize = 68;
    if data.len() < FIXED_FILL_SIZE {
        return Err(Error::InvalidLength {
            expected: FIXED_FILL_SIZE,
            found: data.len(),
        });
    }

    let pattern_type = read_u32(data, 0)?;
    let (fg_color, bg_color) = if pattern_type == 0x28 {
        (None, None)
    } else {
        (
            direct_color(data, 4, kind::FILL, ordinal, &mut tolerance)?,
            direct_color(data, 12, kind::FILL, ordinal, &mut tolerance)?,
        )
    };

    Ok(Fill {
        pattern_type,
        fg_color,
        bg_color,
    })
}

/// Parse one direct BrtColor into compact ARGB form.
///
/// A direct color whose `fValidRGB` bit is clear is refused here; an opened
/// workbook reads it as no color and records the repair.
pub fn parse_direct_color(data: &[u8], offset: usize) -> Result<Option<u32>> {
    direct_color(data, offset, kind::FONT, 0, &mut Tolerance::refusing())
}

fn direct_color(
    data: &[u8],
    offset: usize,
    record_type: crate::raw::Kind,
    ordinal: u32,
    tolerance: &mut Tolerance<'_>,
) -> Result<Option<u32>> {
    let end = offset.checked_add(8).ok_or(Error::InvalidLength {
        expected: usize::MAX,
        found: data.len(),
    })?;
    if end > data.len() {
        return Err(Error::InvalidLength {
            expected: end,
            found: data.len(),
        });
    }

    let flags = data[offset];
    let valid_rgb = flags & 1 != 0;
    let color_type = flags >> 1;
    if color_type != 2 {
        return Ok(None);
    }
    if !valid_rgb {
        // MS-XLSB 2.5.17: with fValidRGB clear the RGB bytes carry no color.
        tolerance.tolerate(
            LayoutDefect::ColorWithoutValidRgb,
            record_type,
            ordinal,
            u32::from(flags),
            || Error::Unrecognized {
                typ: "BrtColor".to_string(),
                val: "direct RGB color is not marked valid".to_string(),
            },
        )?;
        return Ok(None);
    }

    let red = u32::from(data[offset + 4]);
    let green = u32::from(data[offset + 5]);
    let blue = u32::from(data[offset + 6]);
    let alpha = u32::from(data[offset + 7]);
    Ok(Some((alpha << 24) | (red << 16) | (green << 8) | blue))
}

/// Parse a BrtBorder payload.
pub fn parse_border(data: &[u8]) -> Result<Border> {
    parse_border_with(data, 0, Tolerance::refusing())
}

fn parse_border_with(data: &[u8], ordinal: u32, mut tolerance: Tolerance<'_>) -> Result<Border> {
    const BORDER_SIZE: usize = 51;
    if data.len() < BORDER_SIZE {
        return Err(Error::InvalidLength {
            expected: BORDER_SIZE,
            found: data.len(),
        });
    }

    let flags = data[0];
    // MS-XLSB 2.4.43: only fBdrDiagDown and fBdrDiagUp are defined; the
    // reserved bits MUST be ignored.
    if flags & !0x03 != 0 {
        tolerance.tolerate(
            ReservedFieldDefect::BorderFlags,
            kind::BORDER,
            ordinal,
            u32::from(flags & !0x03),
            || Error::Unrecognized {
                typ: "BrtBorder flags".to_string(),
                val: format!("0x{flags:02X}"),
            },
        )?;
    }

    let mut side =
        |range: std::ops::Range<usize>| parse_side(&data[range], ordinal, &mut tolerance);
    Ok(Border {
        top: side(1..11)?,
        bottom: side(11..21)?,
        left: side(21..31)?,
        right: side(31..41)?,
        diagonal: side(41..51)?,
        vertical: None,
        horizontal: None,
        diagonal_down: flags & 1 != 0,
        diagonal_up: flags & 2 != 0,
    })
}

/// Parse a BrtFmt payload.
pub fn parse_num_fmt(data: &[u8]) -> Result<(u32, String)> {
    parse_num_fmt_with(data, 0, Tolerance::refusing())
}

fn parse_num_fmt_with(
    data: &[u8],
    ordinal: u32,
    mut tolerance: Tolerance<'_>,
) -> Result<(u32, String)> {
    // ifmt plus the XLWideString character count; the code-length domain is
    // checked separately below.
    if data.len() < 6 {
        return Err(Error::InvalidLength {
            expected: 6,
            found: data.len(),
        });
    }

    let id = u32::from(read_u16(data, 0)?);
    if !matches!(id, 5..=8 | 23..=26 | 41..=44 | 63..=66 | 164..=382) {
        tolerance.tolerate(LayoutDefect::NumberFormatId, kind::FMT, ordinal, id, || {
            Error::Unrecognized {
                typ: "BrtFmt ifmt".to_string(),
                val: id.to_string(),
            }
        })?;
    }
    let (format_code, consumed) = decode_string(&data[2..])?;
    let length = format_code.encode_utf16().count();
    if !(1..=255).contains(&length) {
        tolerance.tolerate(
            LayoutDefect::NumberFormatCodeLength,
            kind::FMT,
            ordinal,
            u32::try_from(length).unwrap_or(u32::MAX),
            || Error::Unrecognized {
                typ: "BrtFmt stFmtCode length".to_string(),
                val: length.to_string(),
            },
        )?;
    }
    let trailing = data.len().saturating_sub(consumed.saturating_add(2));
    tolerance.tolerate_trailing(kind::FMT, ordinal, trailing, || Error::Unrecognized {
        typ: "BrtFmt".to_string(),
        val: format!("{trailing} trailing bytes"),
    })?;

    Ok((id, format_code))
}

/// Parse the compact fields retained from a BrtXF payload.
pub fn parse_cell_format(data: &[u8]) -> Result<CellFormat> {
    if data.len() < 16 {
        return Err(Error::InvalidLength {
            expected: 16,
            found: data.len(),
        });
    }

    Ok(CellFormat {
        font_id: u32::from(read_u16(data, 4)?),
        fill_id: u32::from(read_u16(data, 6)?),
        border_id: u32::from(read_u16(data, 8)?),
        num_fmt_id: u32::from(read_u16(data, 2)?),
        alignment: parse_alignment(data),
    })
}

fn parse_alignment(data: &[u8]) -> Option<Alignment> {
    let rotation = data[10];
    let indent = data[11];
    let alignment_flags = data[12];
    let property_flags = data[13];
    if rotation == 0 && indent == 0 && alignment_flags == 0 && property_flags & 0x0F == 0 {
        return None;
    }

    Some(Alignment {
        horizontal: HorizontalAlignment::from_u8(alignment_flags & 0x07),
        vertical: VerticalAlignment::from_u8((alignment_flags >> 3) & 0x07),
        rotation,
        indent,
        text_direction: (property_flags >> 2) & 0x03,
        wrap_text: alignment_flags & 0x40 != 0,
        shrink_to_fit: property_flags & 0x01 != 0,
    })
}

fn parse_side(
    data: &[u8],
    ordinal: u32,
    tolerance: &mut Tolerance<'_>,
) -> Result<Option<BorderSide>> {
    let style_byte = data[0];
    let refusal = || Error::Unrecognized {
        typ: "Blxf".to_string(),
        val: format!("style 0x{style_byte:02X}, reserved 0x{:02X}", data[1]),
    };
    // MS-XLSB 2.5.5 Blxf: the reserved byte MUST be 0 and MUST be ignored.
    if data[1] != 0 {
        tolerance.tolerate(
            ReservedFieldDefect::BorderSideReserved,
            kind::BORDER,
            ordinal,
            u32::from(data[1]),
            refusal,
        )?;
    }
    if style_byte > 13 {
        tolerance.tolerate(
            LayoutDefect::BorderStyle,
            kind::BORDER,
            ordinal,
            u32::from(style_byte),
            refusal,
        )?;
        return Ok(None);
    }

    let style = BorderStyle::from_u8(style_byte);
    if style == BorderStyle::None {
        return Ok(None);
    }

    Ok(Some(BorderSide {
        style,
        color: direct_color(data, 2, kind::BORDER, ordinal, tolerance)?,
    }))
}

fn decode_string(data: &[u8]) -> Result<(String, usize)> {
    let mut cursor = Cursor::new(data, "XLWideString");
    let value = cursor.read_wide_string()?;
    Ok((value, cursor.position()))
}

fn read_u16(data: &[u8], offset: usize) -> Result<u16> {
    let end = offset.checked_add(2).ok_or(Error::InvalidLength {
        expected: usize::MAX,
        found: data.len(),
    })?;
    let bytes = data.get(offset..end).ok_or(Error::InvalidLength {
        expected: end,
        found: data.len(),
    })?;
    Ok(u16::from_le_bytes([bytes[0], bytes[1]]))
}

fn read_u32(data: &[u8], offset: usize) -> Result<u32> {
    let end = offset.checked_add(4).ok_or(Error::InvalidLength {
        expected: usize::MAX,
        found: data.len(),
    })?;
    let bytes = data.get(offset..end).ok_or(Error::InvalidLength {
        expected: end,
        found: data.len(),
    })?;
    Ok(u32::from_le_bytes([bytes[0], bytes[1], bytes[2], bytes[3]]))
}

fn reserve<T>(values: &mut Vec<T>, additional: usize, resource: &'static str) -> Result<()> {
    values
        .try_reserve(additional)
        .map_err(|source| allocation(resource, source))
}

fn allocation(resource: &'static str, source: TryReserveError) -> Error {
    Error::Allocation { resource, source }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn direct_color(argb: u32) -> [u8; 8] {
        [
            5,
            0,
            0,
            0,
            ((argb >> 16) & 0xFF) as u8,
            ((argb >> 8) & 0xFF) as u8,
            (argb & 0xFF) as u8,
            ((argb >> 24) & 0xFF) as u8,
        ]
    }

    #[test]
    fn parses_direct_color_and_default_table() {
        assert!(matches!(
            parse_direct_color(&direct_color(0x80402010), 0),
            Ok(Some(0x80402010))
        ));
        assert_eq!(Table::default().get_num_fmt(14), Some("mm-dd-yy"));
    }

    #[test]
    fn rejects_truncated_and_invalid_color_payloads() {
        assert!(matches!(
            parse_font(&[0; 20]),
            Err(Error::InvalidLength { .. })
        ));
        assert!(matches!(
            parse_fill(&[0; 67]),
            Err(Error::InvalidLength { .. })
        ));
        assert!(matches!(
            parse_direct_color(&[4, 0, 0, 0, 1, 2, 3, 4], 0),
            Err(Error::Unrecognized { .. })
        ));
    }
}
