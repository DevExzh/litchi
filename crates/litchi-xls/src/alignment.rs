//! BIFF8 XF cell and style alignment metadata.

use super::error::{Error, Result};
use super::leniency::{FormattingDefect, ToleranceLog};

/// Horizontal alignment of cell text.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum HorizontalAlignment {
    General,
    Left,
    Center,
    Right,
    Fill,
    Justify,
    CenterAcrossSelection,
    Distributed,
}

/// Vertical alignment of cell text.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum VerticalAlignment {
    Top,
    Center,
    Bottom,
    Justify,
    Distributed,
}

/// Rotation applied to cell text.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum TextRotation {
    None,
    CounterClockwise(u8),
    Clockwise(u8),
    Vertical,
}

/// Logical reading order of cell text.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum ReadingOrder {
    Context,
    LeftToRight,
    RightToLeft,
}

/// Alignment properties stored in a BIFF8 cell or style XF record.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct CellAlignment {
    horizontal: HorizontalAlignment,
    vertical: VerticalAlignment,
    wrap_text: bool,
    justify_last_line: bool,
    rotation: TextRotation,
    indent: u8,
    shrink_to_fit: bool,
    reading_order: ReadingOrder,
}

impl Default for CellAlignment {
    fn default() -> Self {
        Self {
            horizontal: HorizontalAlignment::General,
            vertical: VerticalAlignment::Bottom,
            wrap_text: false,
            justify_last_line: false,
            rotation: TextRotation::None,
            indent: 0,
            shrink_to_fit: false,
            reading_order: ReadingOrder::Context,
        }
    }
}

impl CellAlignment {
    pub fn horizontal(&self) -> HorizontalAlignment {
        self.horizontal
    }

    pub fn vertical(&self) -> VerticalAlignment {
        self.vertical
    }

    pub fn wraps_text(&self) -> bool {
        self.wrap_text
    }

    pub fn justifies_last_line(&self) -> bool {
        self.justify_last_line
    }

    pub fn rotation(&self) -> TextRotation {
        self.rotation
    }

    pub fn indent(&self) -> u8 {
        self.indent
    }

    pub fn shrinks_to_fit(&self) -> bool {
        self.shrink_to_fit
    }

    pub fn reading_order(&self) -> ReadingOrder {
        self.reading_order
    }

    /// Parse the three XF alignment bytes under an explicit leniency policy.
    ///
    /// `xf_index` locates the owning XF record in the recorded report. The only
    /// tolerated defect is `fJustLast` without distributed horizontal
    /// alignment, which is cleared; invalid vertical alignment, rotation,
    /// reading order, and reserved bits stay hard errors because each of those
    /// would otherwise silently invent a layout the file never specified.
    pub(crate) fn parse(
        alignment_options: u8,
        rotation: u8,
        indentation_options: u8,
        xf_index: u16,
        tolerance: &mut ToleranceLog,
    ) -> Result<Self> {
        let horizontal = match alignment_options & 0x07 {
            0 => HorizontalAlignment::General,
            1 => HorizontalAlignment::Left,
            2 => HorizontalAlignment::Center,
            3 => HorizontalAlignment::Right,
            4 => HorizontalAlignment::Fill,
            5 => HorizontalAlignment::Justify,
            6 => HorizontalAlignment::CenterAcrossSelection,
            7 => HorizontalAlignment::Distributed,
            _ => unreachable!(),
        };
        let vertical_value = (alignment_options >> 4) & 0x07;
        let vertical = match vertical_value {
            0 => VerticalAlignment::Top,
            1 => VerticalAlignment::Center,
            2 => VerticalAlignment::Bottom,
            3 => VerticalAlignment::Justify,
            4 => VerticalAlignment::Distributed,
            value => return Err(invalid(format!("XF vertical alignment {value} is invalid"))),
        };
        let mut justify_last_line = alignment_options & 0x80 != 0;
        if justify_last_line && horizontal != HorizontalAlignment::Distributed {
            tolerance.tolerate(
                FormattingDefect::AlignmentJustifyLastLine,
                u32::from(xf_index),
                u32::from(alignment_options & 0x07),
                || invalid("XF justify-last-line requires distributed horizontal alignment"),
            )?;
            justify_last_line = false;
        }

        let rotation = match rotation {
            0 => TextRotation::None,
            1..=90 => TextRotation::CounterClockwise(rotation),
            91..=180 => TextRotation::Clockwise(rotation - 90),
            255 => TextRotation::Vertical,
            value => return Err(invalid(format!("XF text rotation {value} is invalid"))),
        };

        if indentation_options & 0x20 != 0 {
            return Err(invalid("XF alignment reserved bit is set"));
        }
        let reading_order = match indentation_options >> 6 {
            0 => ReadingOrder::Context,
            1 => ReadingOrder::LeftToRight,
            2 => ReadingOrder::RightToLeft,
            value => return Err(invalid(format!("XF reading order {value} is invalid"))),
        };

        Ok(Self {
            horizontal,
            vertical,
            wrap_text: alignment_options & 0x08 != 0,
            justify_last_line,
            rotation,
            indent: indentation_options & 0x0f,
            shrink_to_fit: indentation_options & 0x10 != 0,
            reading_order,
        })
    }
}

fn invalid(message: impl Into<String>) -> Error {
    Error::InvalidData(message.into())
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::leniency::Leniency;

    /// Strict-mode shim for the three-byte alignment payload.
    fn parse(
        alignment_options: u8,
        rotation: u8,
        indentation_options: u8,
    ) -> Result<CellAlignment> {
        CellAlignment::parse(
            alignment_options,
            rotation,
            indentation_options,
            0,
            &mut ToleranceLog::new(Leniency::Strict),
        )
    }

    #[test]
    fn parses_every_horizontal_and_vertical_alignment() {
        let horizontal = [
            HorizontalAlignment::General,
            HorizontalAlignment::Left,
            HorizontalAlignment::Center,
            HorizontalAlignment::Right,
            HorizontalAlignment::Fill,
            HorizontalAlignment::Justify,
            HorizontalAlignment::CenterAcrossSelection,
            HorizontalAlignment::Distributed,
        ];
        for (value, expected) in horizontal.into_iter().enumerate() {
            let parsed = parse(value as u8 | 0x20, 0, 0).unwrap();
            assert_eq!(parsed.horizontal(), expected);
            assert_eq!(parsed.vertical(), VerticalAlignment::Bottom);
        }

        let vertical = [
            VerticalAlignment::Top,
            VerticalAlignment::Center,
            VerticalAlignment::Bottom,
            VerticalAlignment::Justify,
            VerticalAlignment::Distributed,
        ];
        for (value, expected) in vertical.into_iter().enumerate() {
            let parsed = parse((value as u8) << 4, 0, 0).unwrap();
            assert_eq!(parsed.vertical(), expected);
        }
    }

    #[test]
    fn parses_flags_indent_and_reading_orders() {
        let parsed = parse(0x80 | 0x08 | 0x07, 0, 0x10 | 0x0f).unwrap();
        assert!(parsed.wraps_text());
        assert!(parsed.justifies_last_line());
        assert_eq!(parsed.indent(), 15);
        assert!(parsed.shrinks_to_fit());

        assert_eq!(
            parse(0x20, 0, 0).unwrap().reading_order(),
            ReadingOrder::Context
        );
        assert_eq!(
            parse(0x20, 0, 0x40).unwrap().reading_order(),
            ReadingOrder::LeftToRight
        );
        assert_eq!(
            parse(0x20, 0, 0x80).unwrap().reading_order(),
            ReadingOrder::RightToLeft
        );
    }

    #[test]
    fn parses_rotation_boundaries() {
        for (value, expected) in [
            (0, TextRotation::None),
            (1, TextRotation::CounterClockwise(1)),
            (90, TextRotation::CounterClockwise(90)),
            (91, TextRotation::Clockwise(1)),
            (180, TextRotation::Clockwise(90)),
            (255, TextRotation::Vertical),
        ] {
            assert_eq!(parse(0x20, value, 0).unwrap().rotation(), expected);
        }
    }

    #[test]
    fn rejects_invalid_enum_and_reserved_values() {
        for value in 5..=7 {
            assert!(parse(value << 4, 0, 0).is_err());
        }
        assert!(parse(0x80, 0, 0).is_err());
        assert!(parse(0x20, 181, 0).is_err());
        assert!(parse(0x20, 254, 0).is_err());
        assert!(parse(0x20, 0, 0x20).is_err());
        assert!(parse(0x20, 0, 0xc0).is_err());
    }

    #[test]
    fn a_lenient_policy_clears_justify_last_line_without_distributed_alignment() {
        // fJustLast set with alcH = Left (1), which MS-XLS forbids.
        const NON_DISTRIBUTED_JUST_LAST: u8 = 0x80 | 0x20 | 0x01;
        const XF_INDEX: u16 = 42;

        let mut tolerance = ToleranceLog::new(Leniency::TolerateFormattingDefects);
        let parsed =
            CellAlignment::parse(NON_DISTRIBUTED_JUST_LAST, 0, 0, XF_INDEX, &mut tolerance)
                .expect("a lenient policy repairs the flag");
        assert!(!parsed.justifies_last_line());
        assert_eq!(parsed.horizontal(), HorizontalAlignment::Left);

        let report = tolerance.into_report();
        assert_eq!(report.count(FormattingDefect::AlignmentJustifyLastLine), 1);
        let entry = report.defects()[0];
        assert_eq!(entry.ordinal(), u32::from(XF_INDEX));
        assert_eq!(entry.observed(), u32::from(HorizontalAlignment::Left as u8));
    }

    #[test]
    fn a_lenient_policy_still_rejects_every_other_alignment_defect() {
        // Only the justify-last-line contradiction is cosmetic; an unknown
        // vertical alignment, rotation, or reading order would otherwise make
        // the reader invent a layout the file never specified.
        let mut tolerance = ToleranceLog::new(Leniency::TolerateFormattingDefects);
        for (alignment_options, rotation, indentation_options) in [
            (0x50u8, 0u8, 0u8),
            (0x20, 181, 0),
            (0x20, 0, 0x20),
            (0x20, 0, 0xc0),
        ] {
            assert!(
                CellAlignment::parse(
                    alignment_options,
                    rotation,
                    indentation_options,
                    0,
                    &mut tolerance,
                )
                .is_err()
            );
        }
        assert!(tolerance.into_report().is_clean());
    }

    #[test]
    fn justify_last_line_survives_when_the_file_is_conforming() {
        const DISTRIBUTED_JUST_LAST: u8 = 0x80 | 0x20 | 0x07;
        for leniency in [Leniency::Strict, Leniency::TolerateFormattingDefects] {
            let mut tolerance = ToleranceLog::new(leniency);
            let parsed = CellAlignment::parse(DISTRIBUTED_JUST_LAST, 0, 0, 0, &mut tolerance)
                .expect("distributed alignment permits fJustLast");
            assert!(parsed.justifies_last_line());
            assert!(tolerance.into_report().is_clean());
        }
    }
}
