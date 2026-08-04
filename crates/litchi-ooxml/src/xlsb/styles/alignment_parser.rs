//! XLSB alignment record parsing
//!
//! This module implements parsing for cell alignment within BrtXF records
//! according to the MS-XLSB specification.
//! Reference: [MS-XLSB] Section 2.4.865 - BrtXF

use crate::xlsb::error::Result;

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
    /// # BrtXF Structure (MS-XLSB Section 2.4.865)
    ///
    /// Rotation and indentation are stored as bytes 10 and 11. Alignment and
    /// protection flags occupy bytes 12 and 13.
    pub fn parse(data: &[u8], offset: usize) -> Result<Option<Self>> {
        if offset + 16 > data.len() {
            return Ok(None);
        }

        let rotation = data[offset + 10];
        let indent = data[offset + 11];
        let alignment_flags = data[offset + 12];
        let property_flags = data[offset + 13];
        if rotation == 0 && indent == 0 && alignment_flags == 0 && property_flags & 0x0F == 0 {
            return Ok(None);
        }

        let horizontal = HorizontalAlignment::from_u8(alignment_flags & 0x07);
        let vertical = VerticalAlignment::from_u8((alignment_flags >> 3) & 0x07);
        let wrap_text = alignment_flags & 0x40 != 0;
        let shrink_to_fit = property_flags & 0x01 != 0;
        let text_direction = (property_flags >> 2) & 0x03;

        Ok(Some(Alignment {
            horizontal,
            vertical,
            rotation,
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

    #[test]
    fn parses_brt_xf_alignment_at_spec_offsets() {
        let mut data = [0u8; 16];
        data[10] = 45;
        data[11] = 7;
        data[12] = 3 | (2 << 3) | 0x40;
        data[13] = 1 | (2 << 2);

        let alignment = Alignment::parse(&data, 0).unwrap().unwrap();
        assert_eq!(alignment.horizontal, HorizontalAlignment::Right);
        assert_eq!(alignment.vertical, VerticalAlignment::Bottom);
        assert_eq!(alignment.rotation, 45);
        assert_eq!(alignment.indent, 7);
        assert_eq!(alignment.text_direction, 2);
        assert!(alignment.wrap_text);
        assert!(alignment.shrink_to_fit);
    }

    #[test]
    fn omits_default_and_truncated_alignment() {
        assert!(Alignment::parse(&[0; 16], 0).unwrap().is_none());
        assert!(Alignment::parse(&[0; 15], 0).unwrap().is_none());
    }
}
