//! Strict BIFF12 styles-part codec.
//!
//! The record layouts are the checked-in [MS-XLSB] Brt* definitions. Host
//! package and compatibility concerns stay outside this module.

use super::model::{
    Alignment, Border, BorderSide, BorderStyle, CellFormat, Fill, Font, HorizontalAlignment, Table,
    VerticalAlignment,
};
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

/// Read one complete XLSB styles stream.
pub fn read(bytes: &[u8]) -> Result<Table> {
    let mut table = Table::default();
    let mut in_fonts = false;
    let mut in_fills = false;
    let mut in_borders = false;
    let mut in_fmts = false;
    let mut in_cell_xfs = false;
    let mut in_cell_style_xfs = false;

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
                table.fonts.push(parse_font(data)?);
            },
            kind::BEGIN_FILLS => {
                in_fills = true;
                table.fills.clear();
            },
            kind::END_FILLS => in_fills = false,
            kind::FILL if in_fills => {
                reserve(&mut table.fills, 1, "styles fills")?;
                table.fills.push(parse_fill(data)?);
            },
            kind::BEGIN_BORDERS => {
                in_borders = true;
                table.borders.clear();
            },
            kind::END_BORDERS => in_borders = false,
            kind::BORDER if in_borders => {
                reserve(&mut table.borders, 1, "styles borders")?;
                table.borders.push(parse_border(data)?);
            },
            kind::BEGIN_FMTS => in_fmts = true,
            kind::END_FMTS => in_fmts = false,
            kind::FMT if in_fmts => {
                let (id, format_code) = parse_num_fmt(data)?;
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

/// Parse a BrtFont payload.
pub fn parse_font(data: &[u8]) -> Result<Font> {
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
    let color = parse_direct_color(data, 12)?;
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
        (parse_direct_color(data, 4)?, parse_direct_color(data, 12)?)
    };

    Ok(Fill {
        pattern_type,
        fg_color,
        bg_color,
    })
}

/// Parse one direct BrtColor into compact ARGB form.
pub fn parse_direct_color(data: &[u8], offset: usize) -> Result<Option<u32>> {
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
        return Err(Error::Unrecognized {
            typ: "BrtColor".to_string(),
            val: "direct RGB color is not marked valid".to_string(),
        });
    }

    let red = u32::from(data[offset + 4]);
    let green = u32::from(data[offset + 5]);
    let blue = u32::from(data[offset + 6]);
    let alpha = u32::from(data[offset + 7]);
    Ok(Some((alpha << 24) | (red << 16) | (green << 8) | blue))
}

/// Parse a BrtBorder payload.
pub fn parse_border(data: &[u8]) -> Result<Border> {
    const BORDER_SIZE: usize = 51;
    if data.len() < BORDER_SIZE {
        return Err(Error::InvalidLength {
            expected: BORDER_SIZE,
            found: data.len(),
        });
    }

    let flags = data[0];
    if flags & !0x03 != 0 {
        return Err(Error::Unrecognized {
            typ: "BrtBorder flags".to_string(),
            val: format!("0x{flags:02X}"),
        });
    }

    Ok(Border {
        top: parse_side(&data[1..11])?,
        bottom: parse_side(&data[11..21])?,
        left: parse_side(&data[21..31])?,
        right: parse_side(&data[31..41])?,
        diagonal: parse_side(&data[41..51])?,
        vertical: None,
        horizontal: None,
        diagonal_down: flags & 1 != 0,
        diagonal_up: flags & 2 != 0,
    })
}

/// Parse a BrtFmt payload.
pub fn parse_num_fmt(data: &[u8]) -> Result<(u32, String)> {
    if data.len() < 8 {
        return Err(Error::InvalidLength {
            expected: 8,
            found: data.len(),
        });
    }

    let id = u32::from(read_u16(data, 0)?);
    if !matches!(id, 5..=8 | 23..=26 | 41..=44 | 63..=66 | 164..=382) {
        return Err(Error::Unrecognized {
            typ: "BrtFmt ifmt".to_string(),
            val: id.to_string(),
        });
    }
    let (format_code, consumed) = decode_string(&data[2..])?;
    let length = format_code.encode_utf16().count();
    if !(1..=255).contains(&length) {
        return Err(Error::Unrecognized {
            typ: "BrtFmt stFmtCode length".to_string(),
            val: length.to_string(),
        });
    }
    if consumed.checked_add(2) != Some(data.len()) {
        return Err(Error::Unrecognized {
            typ: "BrtFmt".to_string(),
            val: format!("{} trailing bytes", data.len() - consumed - 2),
        });
    }

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

fn parse_side(data: &[u8]) -> Result<Option<BorderSide>> {
    let style_byte = data[0];
    if style_byte > 13 || data[1] != 0 {
        return Err(Error::Unrecognized {
            typ: "Blxf".to_string(),
            val: format!("style 0x{style_byte:02X}, reserved 0x{:02X}", data[1]),
        });
    }

    let style = BorderStyle::from_u8(style_byte);
    if style == BorderStyle::None {
        return Ok(None);
    }

    Ok(Some(BorderSide {
        style,
        color: parse_direct_color(data, 2)?,
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
