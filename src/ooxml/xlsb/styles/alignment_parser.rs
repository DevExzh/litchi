//! XLSB alignment record parsing
//!
//! This module implements parsing for cell alignment within BrtXF records
//! according to the MS-XLSB specification.
//! Reference: [MS-XLSB] Section 2.5.148 - XFProps

use crate::common::binary;
use crate::ooxml::xlsb::error::XlsbResult;

/// Horizontal alignment values
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[repr(u8)]
pub enum HorizontalAlignment {
    General = 0,
    Left = 1,
    Center = 2,
    Right = 3,
    Fill = 4,
    Justify = 5,
    CenterContinuous = 6,
    Distributed = 7,
}

impl HorizontalAlignment {
    pub fn from_u8(value: u8) -> Self {
        match value {
            1 => HorizontalAlignment::Left,
            2 => HorizontalAlignment::Center,
            3 => HorizontalAlignment::Right,
            4 => HorizontalAlignment::Fill,
            5 => HorizontalAlignment::Justify,
            6 => HorizontalAlignment::CenterContinuous,
            7 => HorizontalAlignment::Distributed,
            _ => HorizontalAlignment::General,
        }
    }
}

/// Vertical alignment values
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[repr(u8)]
pub enum VerticalAlignment {
    Top = 0,
    Center = 1,
    Bottom = 2,
    Justify = 3,
    Distributed = 4,
}

impl VerticalAlignment {
    pub fn from_u8(value: u8) -> Self {
        match value {
            1 => VerticalAlignment::Center,
            2 => VerticalAlignment::Bottom,
            3 => VerticalAlignment::Justify,
            4 => VerticalAlignment::Distributed,
            _ => VerticalAlignment::Top,
        }
    }
}

/// Cell alignment information
///
/// Fields ordered for compact memory layout with minimal padding.
/// Enums are typically 1-byte each with repr(u8), followed by u8 and bool fields.
#[derive(Debug, Clone)]
pub struct Alignment {
    pub horizontal: HorizontalAlignment,
    pub vertical: VerticalAlignment,
    pub rotation: u8,
    pub indent: u8,
    pub text_direction: u8,
    pub wrap_text: bool,
    pub shrink_to_fit: bool,
}

impl Default for Alignment {
    fn default() -> Self {
        Alignment {
            horizontal: HorizontalAlignment::General,
            vertical: VerticalAlignment::Bottom,
            rotation: 0,
            indent: 0,
            text_direction: 0,
            wrap_text: false,
            shrink_to_fit: false,
        }
    }
}

impl Alignment {
    /// Parse alignment from XF record data
    ///
    /// # XFProps Structure (MS-XLSB Section 2.5.148)
    ///
    /// The alignment is encoded in a bitfield structure:
    /// - Bits 0-2: horizontal alignment (3 bits)
    /// - Bits 3-4: vertical alignment (2 bits)
    /// - Bit 5: wrap text (1 bit)
    /// - Bits 6-9: rotation (4 bits for angle/90)
    /// - Bit 10: shrink to fit (1 bit)
    /// - Bits 11-14: indent level (4 bits)
    /// - Bits 15-16: text direction (2 bits)
    pub fn parse(data: &[u8], offset: usize) -> XlsbResult<Option<Self>> {
        // XF record layout (simplified):
        // Offset 0-1: font ID (u16)
        // Offset 2-3: num fmt ID (u16)
        // Offset 4-5: fill ID (u16)
        // Offset 6-7: border ID (u16)
        // Offset 8-9: XF flags (u16) - indicates what follows
        // Offset 10+: alignment data (if present)

        if offset + 10 > data.len() {
            return Ok(None);
        }

        // Read XF flags to determine if alignment is present
        let xf_flags = binary::read_u16_le_at(data, offset + 8)?;

        // Bit 4 of xf_flags indicates if alignment is present
        let has_alignment = (xf_flags & 0x0010) != 0;

        if !has_alignment || offset + 12 > data.len() {
            return Ok(None);
        }

        // Read alignment flags (2 bytes starting at offset+10)
        let align_flags = binary::read_u16_le_at(data, offset + 10)?;

        // Extract alignment properties from bitfield
        let horizontal = HorizontalAlignment::from_u8((align_flags & 0x07) as u8);
        let vertical = VerticalAlignment::from_u8(((align_flags >> 3) & 0x03) as u8);
        let wrap_text = (align_flags & 0x0020) != 0;
        let rotation = ((align_flags >> 6) & 0x000F) as u8;
        let shrink_to_fit = (align_flags & 0x0400) != 0;
        let indent = ((align_flags >> 11) & 0x000F) as u8;
        let text_direction = ((align_flags >> 15) & 0x0003) as u8;

        // Convert rotation from 90-degree units to degrees
        let rotation_degrees = rotation.saturating_mul(15); // Excel uses 15-degree increments

        Ok(Some(Alignment {
            horizontal,
            vertical,
            rotation: rotation_degrees,
            indent,
            text_direction,
            wrap_text,
            shrink_to_fit,
        }))
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_horizontal_alignment_conversion() {
        assert_eq!(
            HorizontalAlignment::from_u8(0),
            HorizontalAlignment::General
        );
        assert_eq!(HorizontalAlignment::from_u8(1), HorizontalAlignment::Left);
        assert_eq!(HorizontalAlignment::from_u8(2), HorizontalAlignment::Center);
        assert_eq!(HorizontalAlignment::from_u8(3), HorizontalAlignment::Right);
    }

    #[test]
    fn test_vertical_alignment_conversion() {
        assert_eq!(VerticalAlignment::from_u8(0), VerticalAlignment::Top);
        assert_eq!(VerticalAlignment::from_u8(1), VerticalAlignment::Center);
        assert_eq!(VerticalAlignment::from_u8(2), VerticalAlignment::Bottom);
    }

    #[test]
    fn test_default_alignment() {
        let align = Alignment::default();
        assert_eq!(align.horizontal, HorizontalAlignment::General);
        assert_eq!(align.vertical, VerticalAlignment::Bottom);
        assert!(!align.wrap_text);
        assert!(!align.shrink_to_fit);
    }
}
