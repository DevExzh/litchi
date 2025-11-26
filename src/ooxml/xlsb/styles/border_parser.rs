//! XLSB border record parsing
//!
//! This module implements parsing for BrtBorder records according to the MS-XLSB specification.
//! Reference: [MS-XLSB] Section 2.4.55 - BrtBorder

use crate::common::binary;
use crate::ooxml::xlsb::error::XlsbResult;

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
}

impl Border {
    /// Parse border from BrtBorder record data
    ///
    /// # BrtBorder Structure (MS-XLSB Section 2.4.55)
    ///
    /// The BrtBorder record specifies border formatting properties.
    /// Each border side is encoded with:
    /// - 1 byte: border style
    /// - 4 bytes: ARGB color (optional)
    pub fn parse(data: &[u8]) -> XlsbResult<Self> {
        if data.is_empty() {
            return Ok(Border::default());
        }

        let mut offset = 0;
        let mut border = Border::default();

        // Read flags to determine which borders are present
        if data.len() < 2 {
            return Ok(border);
        }

        let flags = binary::read_u16_le_at(data, offset)?;
        offset += 2;

        // Bit flags indicate which borders are defined
        // Bit 0: top
        // Bit 1: bottom
        // Bit 2: left
        // Bit 3: right
        // Bit 4: diagonal
        // Bit 5: vertical (for table borders)
        // Bit 6: horizontal (for table borders)

        // Parse top border
        if (flags & 0x01) != 0
            && let Some(side) = Self::parse_border_side(data, &mut offset)?
        {
            border.top = Some(side);
        }

        // Parse bottom border
        if (flags & 0x02) != 0
            && let Some(side) = Self::parse_border_side(data, &mut offset)?
        {
            border.bottom = Some(side);
        }

        // Parse left border
        if (flags & 0x04) != 0
            && let Some(side) = Self::parse_border_side(data, &mut offset)?
        {
            border.left = Some(side);
        }

        // Parse right border
        if (flags & 0x08) != 0
            && let Some(side) = Self::parse_border_side(data, &mut offset)?
        {
            border.right = Some(side);
        }

        // Parse diagonal border
        if (flags & 0x10) != 0
            && let Some(side) = Self::parse_border_side(data, &mut offset)?
        {
            border.diagonal = Some(side);
        }

        // Parse vertical border (for table borders)
        if (flags & 0x20) != 0
            && let Some(side) = Self::parse_border_side(data, &mut offset)?
        {
            border.vertical = Some(side);
        }

        // Parse horizontal border (for table borders)
        if (flags & 0x40) != 0
            && let Some(side) = Self::parse_border_side(data, &mut offset)?
        {
            border.horizontal = Some(side);
        }

        Ok(border)
    }

    /// Parse a single border side
    ///
    /// Format:
    /// - 1 byte: border style
    /// - 4 bytes: color (ARGB) - optional based on style
    fn parse_border_side(data: &[u8], offset: &mut usize) -> XlsbResult<Option<BorderSide>> {
        if *offset >= data.len() {
            return Ok(None);
        }

        // Read border style
        let style_byte = data[*offset];
        *offset += 1;

        let style = BorderStyle::from_u8(style_byte);

        // If style is None, no color follows
        if style == BorderStyle::None {
            return Ok(None);
        }

        // Read color (4 bytes ARGB)
        let color = if *offset + 4 <= data.len() {
            let c = binary::read_u32_le_at(data, *offset)?;
            *offset += 4;
            Some(c)
        } else {
            None
        };

        Ok(Some(BorderSide { style, color }))
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
        let border = Border::parse(&[]).unwrap();
        assert!(border.top.is_none());
        assert!(border.bottom.is_none());
    }

    #[test]
    fn test_border_with_top() {
        // Flags: 0x01 (top border present)
        // Style: 0x01 (Thin)
        // Color: 0x00000000 (black)
        let data = vec![0x01, 0x00, 0x01, 0x00, 0x00, 0x00, 0x00];
        let border = Border::parse(&data).unwrap();
        assert!(border.top.is_some());
        let top = border.top.unwrap();
        assert_eq!(top.style, BorderStyle::Thin);
    }
}
