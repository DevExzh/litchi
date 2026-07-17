//! BIFF8 XF cell and style alignment metadata.

use super::error::{XlsError, XlsResult};

/// Horizontal alignment of cell text.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum XlsHorizontalAlignment {
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
pub enum XlsVerticalAlignment {
    Top,
    Center,
    Bottom,
    Justify,
    Distributed,
}

/// Rotation applied to cell text.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum XlsTextRotation {
    None,
    CounterClockwise(u8),
    Clockwise(u8),
    Vertical,
}

/// Logical reading order of cell text.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum XlsReadingOrder {
    Context,
    LeftToRight,
    RightToLeft,
}

/// Alignment properties stored in a BIFF8 cell or style XF record.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct XlsCellAlignment {
    horizontal: XlsHorizontalAlignment,
    vertical: XlsVerticalAlignment,
    wrap_text: bool,
    justify_last_line: bool,
    rotation: XlsTextRotation,
    indent: u8,
    shrink_to_fit: bool,
    reading_order: XlsReadingOrder,
}

impl Default for XlsCellAlignment {
    fn default() -> Self {
        Self {
            horizontal: XlsHorizontalAlignment::General,
            vertical: XlsVerticalAlignment::Bottom,
            wrap_text: false,
            justify_last_line: false,
            rotation: XlsTextRotation::None,
            indent: 0,
            shrink_to_fit: false,
            reading_order: XlsReadingOrder::Context,
        }
    }
}

impl XlsCellAlignment {
    pub fn horizontal(&self) -> XlsHorizontalAlignment {
        self.horizontal
    }

    pub fn vertical(&self) -> XlsVerticalAlignment {
        self.vertical
    }

    pub fn wraps_text(&self) -> bool {
        self.wrap_text
    }

    pub fn justifies_last_line(&self) -> bool {
        self.justify_last_line
    }

    pub fn rotation(&self) -> XlsTextRotation {
        self.rotation
    }

    pub fn indent(&self) -> u8 {
        self.indent
    }

    pub fn shrinks_to_fit(&self) -> bool {
        self.shrink_to_fit
    }

    pub fn reading_order(&self) -> XlsReadingOrder {
        self.reading_order
    }

    pub(crate) fn parse(
        alignment_options: u8,
        rotation: u8,
        indentation_options: u8,
    ) -> XlsResult<Self> {
        let horizontal = match alignment_options & 0x07 {
            0 => XlsHorizontalAlignment::General,
            1 => XlsHorizontalAlignment::Left,
            2 => XlsHorizontalAlignment::Center,
            3 => XlsHorizontalAlignment::Right,
            4 => XlsHorizontalAlignment::Fill,
            5 => XlsHorizontalAlignment::Justify,
            6 => XlsHorizontalAlignment::CenterAcrossSelection,
            7 => XlsHorizontalAlignment::Distributed,
            _ => unreachable!(),
        };
        let vertical_value = (alignment_options >> 4) & 0x07;
        let vertical = match vertical_value {
            0 => XlsVerticalAlignment::Top,
            1 => XlsVerticalAlignment::Center,
            2 => XlsVerticalAlignment::Bottom,
            3 => XlsVerticalAlignment::Justify,
            4 => XlsVerticalAlignment::Distributed,
            value => return Err(invalid(format!("XF vertical alignment {value} is invalid"))),
        };
        let justify_last_line = alignment_options & 0x80 != 0;
        if justify_last_line && horizontal != XlsHorizontalAlignment::Distributed {
            return Err(invalid(
                "XF justify-last-line requires distributed horizontal alignment",
            ));
        }

        let rotation = match rotation {
            0 => XlsTextRotation::None,
            1..=90 => XlsTextRotation::CounterClockwise(rotation),
            91..=180 => XlsTextRotation::Clockwise(rotation - 90),
            255 => XlsTextRotation::Vertical,
            value => return Err(invalid(format!("XF text rotation {value} is invalid"))),
        };

        if indentation_options & 0x20 != 0 {
            return Err(invalid("XF alignment reserved bit is set"));
        }
        let reading_order = match indentation_options >> 6 {
            0 => XlsReadingOrder::Context,
            1 => XlsReadingOrder::LeftToRight,
            2 => XlsReadingOrder::RightToLeft,
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

fn invalid(message: impl Into<String>) -> XlsError {
    XlsError::InvalidData(message.into())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn parses_every_horizontal_and_vertical_alignment() {
        let horizontal = [
            XlsHorizontalAlignment::General,
            XlsHorizontalAlignment::Left,
            XlsHorizontalAlignment::Center,
            XlsHorizontalAlignment::Right,
            XlsHorizontalAlignment::Fill,
            XlsHorizontalAlignment::Justify,
            XlsHorizontalAlignment::CenterAcrossSelection,
            XlsHorizontalAlignment::Distributed,
        ];
        for (value, expected) in horizontal.into_iter().enumerate() {
            let parsed = XlsCellAlignment::parse(value as u8 | 0x20, 0, 0).unwrap();
            assert_eq!(parsed.horizontal(), expected);
            assert_eq!(parsed.vertical(), XlsVerticalAlignment::Bottom);
        }

        let vertical = [
            XlsVerticalAlignment::Top,
            XlsVerticalAlignment::Center,
            XlsVerticalAlignment::Bottom,
            XlsVerticalAlignment::Justify,
            XlsVerticalAlignment::Distributed,
        ];
        for (value, expected) in vertical.into_iter().enumerate() {
            let parsed = XlsCellAlignment::parse((value as u8) << 4, 0, 0).unwrap();
            assert_eq!(parsed.vertical(), expected);
        }
    }

    #[test]
    fn parses_flags_indent_and_reading_orders() {
        let parsed = XlsCellAlignment::parse(0x80 | 0x08 | 0x07, 0, 0x10 | 0x0f).unwrap();
        assert!(parsed.wraps_text());
        assert!(parsed.justifies_last_line());
        assert_eq!(parsed.indent(), 15);
        assert!(parsed.shrinks_to_fit());

        assert_eq!(
            XlsCellAlignment::parse(0x20, 0, 0).unwrap().reading_order(),
            XlsReadingOrder::Context
        );
        assert_eq!(
            XlsCellAlignment::parse(0x20, 0, 0x40)
                .unwrap()
                .reading_order(),
            XlsReadingOrder::LeftToRight
        );
        assert_eq!(
            XlsCellAlignment::parse(0x20, 0, 0x80)
                .unwrap()
                .reading_order(),
            XlsReadingOrder::RightToLeft
        );
    }

    #[test]
    fn parses_rotation_boundaries() {
        for (value, expected) in [
            (0, XlsTextRotation::None),
            (1, XlsTextRotation::CounterClockwise(1)),
            (90, XlsTextRotation::CounterClockwise(90)),
            (91, XlsTextRotation::Clockwise(1)),
            (180, XlsTextRotation::Clockwise(90)),
            (255, XlsTextRotation::Vertical),
        ] {
            assert_eq!(
                XlsCellAlignment::parse(0x20, value, 0).unwrap().rotation(),
                expected
            );
        }
    }

    #[test]
    fn rejects_invalid_enum_and_reserved_values() {
        for value in 5..=7 {
            assert!(XlsCellAlignment::parse(value << 4, 0, 0).is_err());
        }
        assert!(XlsCellAlignment::parse(0x80, 0, 0).is_err());
        assert!(XlsCellAlignment::parse(0x20, 181, 0).is_err());
        assert!(XlsCellAlignment::parse(0x20, 254, 0).is_err());
        assert!(XlsCellAlignment::parse(0x20, 0, 0x20).is_err());
        assert!(XlsCellAlignment::parse(0x20, 0, 0xc0).is_err());
    }
}
