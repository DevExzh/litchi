//! Word 97 `Brc80` and `SPgbProp` codecs for section page borders.

use super::model::{ApplyTo, Art, Border, Borders, Color, Depth, Error, Offset, Style};
use crate::doc::package::{DocError, Result};
use crate::sprm::Sprm;

/// Decode one fixed-size `Brc80` section-border operand.
pub(crate) fn decode_brc80(sprm: &Sprm, name: &str) -> Result<Option<Border>> {
    let operand = sprm.operand_bytes();
    if operand.len() != 4 {
        return corrupted(&format!("{name} operand must contain exactly 4 bytes"));
    }

    let style = match operand[1] {
        0x00 => None,
        0x01 => Some(Style::Single),
        0x03 => Some(Style::Double),
        0x05 => Some(Style::Thick),
        0x06 => Some(Style::Dotted),
        0x07 => Some(Style::Dashed),
        0x08 => Some(Style::DotDash),
        0x09 => Some(Style::DotDotDash),
        0x0A => Some(Style::Triple),
        0x0B => Some(Style::ThinThickSmallGap),
        0x0C => Some(Style::ThickThinSmallGap),
        0x0D => Some(Style::ThinThickThinSmallGap),
        0x0E => Some(Style::ThinThickMediumGap),
        0x0F => Some(Style::ThickThinMediumGap),
        0x10 => Some(Style::ThinThickThinMediumGap),
        0x11 => Some(Style::ThinThickLargeGap),
        0x12 => Some(Style::ThickThinLargeGap),
        0x13 => Some(Style::ThinThickThinLargeGap),
        0x14 => Some(Style::Wave),
        0x15 => Some(Style::DoubleWave),
        0x16 => Some(Style::DashSmallGap),
        0x17 => Some(Style::DashDotStroked),
        0x18 => Some(Style::ThreeDEmboss),
        0x19 => Some(Style::ThreeDEngrave),
        code => Some(Style::Art(Art::try_from(code).map_err(|error| {
            DocError::Corrupted(format!(
                "{name} contains invalid Brc80 border type {code:#04x}: {error}"
            ))
        })?)),
    };

    let color = match operand[2] {
        0x00 => Color::Automatic,
        0x01 => Color::Black,
        0x02 => Color::Blue,
        0x03 => Color::Cyan,
        0x04 => Color::Green,
        0x05 => Color::Magenta,
        0x06 => Color::Red,
        0x07 => Color::Yellow,
        0x08 => Color::White,
        0x09 => Color::DarkBlue,
        0x0A => Color::DarkCyan,
        0x0B => Color::DarkGreen,
        0x0C => Color::DarkMagenta,
        0x0D => Color::DarkRed,
        0x0E => Color::DarkYellow,
        0x0F => Color::DarkGray,
        0x10 => Color::LightGray,
        invalid => {
            return corrupted(&format!(
                "{name} contains invalid Ico color index {invalid:#04x}"
            ));
        },
    };

    let effects = operand[3];
    if effects & 0x80 != 0 {
        return corrupted(&format!(
            "{name} contains reserved Brc80 effect bits {:#04x}",
            effects & 0x80
        ));
    }
    let Some(style) = style else {
        return Ok(None);
    };
    Ok(Some(Border {
        style,
        width_eighth_points: operand[0],
        color,
        spacing_points: effects & 0x1F,
        shadow: effects & 0x20 != 0,
        frame: effects & 0x40 != 0,
    }))
}

/// Decode the two-byte `SPgbProp` section-border placement operand.
pub(crate) fn decode_pgb_prop(borders: &mut Borders, sprm: &Sprm) -> Result<()> {
    let operand = sprm.operand_bytes();
    if operand.len() != 2 {
        return corrupted("sprmSPgbProp operand must contain exactly 2 bytes");
    }
    if operand[1] != 0 {
        return corrupted("sprmSPgbProp reserved byte must be zero");
    }
    borders.apply_to = match operand[0] & 0x07 {
        0 => ApplyTo::AllPages,
        1 => ApplyTo::FirstPage,
        2 => ApplyTo::AllButFirstPage,
        _ => return corrupted("sprmSPgbProp contains an invalid PgbApplyTo value"),
    };
    borders.depth = match (operand[0] >> 3) & 0x03 {
        0 => Depth::InFront,
        1 => Depth::Behind,
        _ => return corrupted("sprmSPgbProp contains an invalid PgbPageDepth value"),
    };
    borders.offset_from = match (operand[0] >> 5) & 0x07 {
        0 => Offset::Text,
        1 => Offset::PageEdge,
        _ => return corrupted("sprmSPgbProp contains an invalid PgbOffsetFrom value"),
    };
    Ok(())
}

/// Encode all present page borders and non-default placement into a SEPX.
pub(crate) fn encode_sepx(
    output: &mut Vec<u8>,
    borders: &Borders,
) -> std::result::Result<(), Error> {
    borders.validate()?;
    for (opcode, border) in [
        (crate::sprm_operations::SPRM_S_BRC_TOP80, borders.top),
        (crate::sprm_operations::SPRM_S_BRC_LEFT80, borders.left),
        (crate::sprm_operations::SPRM_S_BRC_BOTTOM80, borders.bottom),
        (crate::sprm_operations::SPRM_S_BRC_RIGHT80, borders.right),
    ] {
        if let Some(border) = border {
            encode_brc80(output, opcode, border)?;
        }
    }
    if borders.apply_to != ApplyTo::AllPages
        || borders.depth != Depth::InFront
        || borders.offset_from != Offset::Text
    {
        encode_pgb_prop(output, borders);
    }
    Ok(())
}

fn encode_brc80(
    output: &mut Vec<u8>,
    opcode: u16,
    border: Border,
) -> std::result::Result<(), Error> {
    border.validate()?;
    let style = match border.style {
        Style::Single => 0x01,
        Style::Double => 0x03,
        Style::Thick => 0x05,
        Style::Dotted => 0x06,
        Style::Dashed => 0x07,
        Style::DotDash => 0x08,
        Style::DotDotDash => 0x09,
        Style::Triple => 0x0A,
        Style::ThinThickSmallGap => 0x0B,
        Style::ThickThinSmallGap => 0x0C,
        Style::ThinThickThinSmallGap => 0x0D,
        Style::ThinThickMediumGap => 0x0E,
        Style::ThickThinMediumGap => 0x0F,
        Style::ThinThickThinMediumGap => 0x10,
        Style::ThinThickLargeGap => 0x11,
        Style::ThickThinLargeGap => 0x12,
        Style::ThinThickThinLargeGap => 0x13,
        Style::Wave => 0x14,
        Style::DoubleWave => 0x15,
        Style::DashSmallGap => 0x16,
        Style::DashDotStroked => 0x17,
        Style::ThreeDEmboss => 0x18,
        Style::ThreeDEngrave => 0x19,
        Style::Art(art) => art.code(),
    };
    let color = match border.color {
        Color::Automatic => 0x00,
        Color::Black => 0x01,
        Color::Blue => 0x02,
        Color::Cyan => 0x03,
        Color::Green => 0x04,
        Color::Magenta => 0x05,
        Color::Red => 0x06,
        Color::Yellow => 0x07,
        Color::White => 0x08,
        Color::DarkBlue => 0x09,
        Color::DarkCyan => 0x0A,
        Color::DarkGreen => 0x0B,
        Color::DarkMagenta => 0x0C,
        Color::DarkRed => 0x0D,
        Color::DarkYellow => 0x0E,
        Color::DarkGray => 0x0F,
        Color::LightGray => 0x10,
    };
    output.extend_from_slice(&opcode.to_le_bytes());
    output.extend_from_slice(&[
        border.width_eighth_points,
        style,
        color,
        border.spacing_points | u8::from(border.shadow) << 5 | u8::from(border.frame) << 6,
    ]);
    Ok(())
}

fn encode_pgb_prop(output: &mut Vec<u8>, borders: &Borders) {
    let apply_to = match borders.apply_to {
        ApplyTo::AllPages => 0,
        ApplyTo::FirstPage => 1,
        ApplyTo::AllButFirstPage => 2,
    };
    let depth = match borders.depth {
        Depth::InFront => 0,
        Depth::Behind => 1,
    };
    let offset_from = match borders.offset_from {
        Offset::Text => 0,
        Offset::PageEdge => 1,
    };
    output.extend_from_slice(&crate::sprm_operations::SPRM_S_PGB_PROP.to_le_bytes());
    output.push(apply_to | depth << 3 | offset_from << 5);
    output.push(0);
}

fn corrupted<T>(message: &str) -> Result<T> {
    Err(DocError::Corrupted(message.to_string()))
}
