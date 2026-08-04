//! XLSB border record parsing
//!
//! This module implements parsing for BrtBorder records according to the MS-XLSB specification.
//! Reference: [MS-XLSB] Section 2.4.55 - BrtBorder

use crate::xlsb::error::{Error, Result};

/// Border side information
#[derive(Debug, Clone)]
pub struct BorderSide {
    pub style: BorderStyle,
    pub color: Option<u32>,
}

/// Border styles matching Excel's border styles
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
#[repr(u8)]
pub enum BorderStyle {
    #[default]
    None = 0,
    Thin = 1,
    Medium = 2,
    Dashed = 3,
    Dotted = 4,
    Thick = 5,
    Double = 6,
    Hair = 7,
    MediumDashed = 8,
    DashDot = 9,
    MediumDashDot = 10,
    DashDotDot = 11,
    MediumDashDotDot = 12,
    SlantDashDot = 13,
}

impl BorderStyle {
    /// Convert from u8 value
    pub fn from_u8(value: u8) -> Self {
        match value {
            1 => BorderStyle::Thin,
            2 => BorderStyle::Medium,
            3 => BorderStyle::Dashed,
            4 => BorderStyle::Dotted,
            5 => BorderStyle::Thick,
            6 => BorderStyle::Double,
            7 => BorderStyle::Hair,
            8 => BorderStyle::MediumDashed,
            9 => BorderStyle::DashDot,
            10 => BorderStyle::MediumDashDot,
            11 => BorderStyle::DashDotDot,
            12 => BorderStyle::MediumDashDotDot,
            13 => BorderStyle::SlantDashDot,
            _ => BorderStyle::None,
        }
    }
}

/// Border container
#[derive(Debug, Clone, Default)]
pub struct Border {
    pub top: Option<BorderSide>,
    pub bottom: Option<BorderSide>,
    pub left: Option<BorderSide>,
    pub right: Option<BorderSide>,
    pub diagonal: Option<BorderSide>,
    pub vertical: Option<BorderSide>,
    pub horizontal: Option<BorderSide>,
    pub diagonal_down: bool,
    pub diagonal_up: bool,
}

impl Border {
    /// Parse border from BrtBorder record data
    ///
    /// # BrtBorder Structure (MS-XLSB Section 2.4.55)
    ///
    /// The BrtBorder record specifies border formatting properties.
    /// Each border side is a 10-byte `Blxf`: style, reserved byte, and an
    /// 8-byte `BrtColor`.
    pub fn parse(data: &[u8]) -> Result<Self> {
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
            top: Self::parse_blxf(&data[1..11])?,
            bottom: Self::parse_blxf(&data[11..21])?,
            left: Self::parse_blxf(&data[21..31])?,
            right: Self::parse_blxf(&data[31..41])?,
            diagonal: Self::parse_blxf(&data[41..51])?,
            vertical: None,
            horizontal: None,
            diagonal_down: flags & 1 != 0,
            diagonal_up: flags & 2 != 0,
        })
    }

    fn parse_blxf(data: &[u8]) -> Result<Option<BorderSide>> {
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

        let color = Self::parse_direct_color(&data[2..10])?;
        Ok(Some(BorderSide { style, color }))
    }

    fn parse_direct_color(data: &[u8]) -> Result<Option<u32>> {
        let valid_rgb = data[0] & 1 != 0;
        let color_type = data[0] >> 1;
        if color_type != 2 {
            return Ok(None);
        }
        if !valid_rgb {
            return Err(Error::Unrecognized {
                typ: "BrtColor".to_string(),
                val: "direct RGB color is not marked valid".to_string(),
            });
        }
        Ok(Some(
            (u32::from(data[7]) << 24)
                | (u32::from(data[4]) << 16)
                | (u32::from(data[5]) << 8)
                | u32::from(data[6]),
        ))
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_border_style_conversion() {
        assert_eq!(BorderStyle::from_u8(0), BorderStyle::None);
        assert_eq!(BorderStyle::from_u8(1), BorderStyle::Thin);
        assert_eq!(BorderStyle::from_u8(5), BorderStyle::Thick);
        assert_eq!(BorderStyle::from_u8(255), BorderStyle::None);
    }

    #[test]
    fn test_empty_border() {
        assert!(matches!(
            Border::parse(&[]),
            Err(Error::InvalidLength { .. })
        ));
    }

    #[test]
    fn test_border_with_top() {
        let mut data = vec![0; 51];
        data[1] = 1;
        data[3..11].copy_from_slice(&[5, 0, 0, 0, 0x40, 0x20, 0x10, 0x80]);
        let border = Border::parse(&data).unwrap();
        assert!(border.top.is_some());
        let top = border.top.unwrap();
        assert_eq!(top.style, BorderStyle::Thin);
        assert_eq!(top.color, Some(0x80402010));
        assert!(border.bottom.is_none());
    }

    #[test]
    fn parses_diagonal_flags_and_side() {
        let mut data = vec![0; 51];
        data[0] = 3;
        data[41] = 6;
        let border = Border::parse(&data).unwrap();
        assert!(border.diagonal_down);
        assert!(border.diagonal_up);
        assert_eq!(border.diagonal.unwrap().style, BorderStyle::Double);
    }
}
