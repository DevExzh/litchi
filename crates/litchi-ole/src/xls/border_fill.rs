//! BIFF8 XF border and fill metadata.

use crate::xls::{XlsColor, XlsError, XlsPalette, XlsResult};

/// A BIFF8 cell border line style.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
#[non_exhaustive]
pub enum XlsBorderStyle {
    #[default]
    None,
    Thin,
    Medium,
    Dashed,
    Dotted,
    Thick,
    Double,
    Hair,
    MediumDashed,
    DashDot,
    MediumDashDot,
    DashDotDot,
    MediumDashDotDot,
    SlantedDashDot,
}

impl XlsBorderStyle {
    fn from_bits(value: u32) -> XlsResult<Self> {
        match value {
            0 => Ok(Self::None),
            1 => Ok(Self::Thin),
            2 => Ok(Self::Medium),
            3 => Ok(Self::Dashed),
            4 => Ok(Self::Dotted),
            5 => Ok(Self::Thick),
            6 => Ok(Self::Double),
            7 => Ok(Self::Hair),
            8 => Ok(Self::MediumDashed),
            9 => Ok(Self::DashDot),
            10 => Ok(Self::MediumDashDot),
            11 => Ok(Self::DashDotDot),
            12 => Ok(Self::MediumDashDotDot),
            13 => Ok(Self::SlantedDashDot),
            _ => Err(XlsError::InvalidData(format!(
                "reserved BIFF8 border style {value}"
            ))),
        }
    }
}

/// The style and indexed color of one cell border edge.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct XlsBorderSide {
    style: XlsBorderStyle,
    color_index: u16,
}

impl XlsBorderSide {
    pub fn style(&self) -> XlsBorderStyle {
        self.style
    }

    pub fn color_index(&self) -> u16 {
        self.color_index
    }

    pub fn color(&self, palette: &XlsPalette) -> Option<XlsColor> {
        palette.color(self.color_index)
    }
}

/// All border metadata stored by a BIFF8 XF record.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct XlsCellBorders {
    left: XlsBorderSide,
    right: XlsBorderSide,
    top: XlsBorderSide,
    bottom: XlsBorderSide,
    diagonal: XlsBorderSide,
    diagonal_down: bool,
    diagonal_up: bool,
}

impl XlsCellBorders {
    pub fn left(&self) -> &XlsBorderSide {
        &self.left
    }

    pub fn right(&self) -> &XlsBorderSide {
        &self.right
    }

    pub fn top(&self) -> &XlsBorderSide {
        &self.top
    }

    pub fn bottom(&self) -> &XlsBorderSide {
        &self.bottom
    }

    pub fn diagonal(&self) -> &XlsBorderSide {
        &self.diagonal
    }

    /// Returns whether the diagonal runs from top-left to bottom-right.
    pub fn diagonal_down(&self) -> bool {
        self.diagonal_down
    }

    /// Returns whether the diagonal runs from bottom-left to top-right.
    pub fn diagonal_up(&self) -> bool {
        self.diagonal_up
    }
}

/// A BIFF8 cell fill pattern.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
#[non_exhaustive]
pub enum XlsFillPattern {
    #[default]
    None,
    Solid,
    MediumGray,
    DarkGray,
    LightGray,
    DarkHorizontal,
    DarkVertical,
    DarkDown,
    DarkUp,
    DarkGrid,
    DarkTrellis,
    LightHorizontal,
    LightVertical,
    LightDown,
    LightUp,
    LightGrid,
    LightTrellis,
    Gray125,
    Gray0625,
}

impl XlsFillPattern {
    fn from_bits(value: u32) -> XlsResult<Self> {
        match value {
            0 => Ok(Self::None),
            1 => Ok(Self::Solid),
            2 => Ok(Self::MediumGray),
            3 => Ok(Self::DarkGray),
            4 => Ok(Self::LightGray),
            5 => Ok(Self::DarkHorizontal),
            6 => Ok(Self::DarkVertical),
            7 => Ok(Self::DarkDown),
            8 => Ok(Self::DarkUp),
            9 => Ok(Self::DarkGrid),
            10 => Ok(Self::DarkTrellis),
            11 => Ok(Self::LightHorizontal),
            12 => Ok(Self::LightVertical),
            13 => Ok(Self::LightDown),
            14 => Ok(Self::LightUp),
            15 => Ok(Self::LightGrid),
            16 => Ok(Self::LightTrellis),
            17 => Ok(Self::Gray125),
            18 => Ok(Self::Gray0625),
            _ => Err(XlsError::InvalidData(format!(
                "reserved BIFF8 fill pattern {value}"
            ))),
        }
    }
}

/// Fill pattern and foreground/background indexed colors from a BIFF8 XF.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct XlsCellFill {
    pattern: XlsFillPattern,
    foreground_color_index: u16,
    background_color_index: u16,
}

impl XlsCellFill {
    pub fn pattern(&self) -> XlsFillPattern {
        self.pattern
    }

    pub fn foreground_color_index(&self) -> u16 {
        self.foreground_color_index
    }

    pub fn background_color_index(&self) -> u16 {
        self.background_color_index
    }

    pub fn foreground_color(&self, palette: &XlsPalette) -> Option<XlsColor> {
        palette.color(self.foreground_color_index)
    }

    pub fn background_color(&self, palette: &XlsPalette) -> Option<XlsColor> {
        palette.color(self.background_color_index)
    }
}

pub(crate) fn parse_xf_border_fill(
    data: &[u8],
    style_xf: bool,
) -> XlsResult<(XlsCellBorders, XlsCellFill)> {
    debug_assert!(data.len() >= 20);
    let border1 = u32::from_le_bytes([data[10], data[11], data[12], data[13]]);
    let border2 = u32::from_le_bytes([data[14], data[15], data[16], data[17]]);
    let area = u16::from_le_bytes([data[18], data[19]]);

    let _ = style_xf;

    let diagonal_down = border1 & (1 << 30) != 0;
    let diagonal_up = border1 & (1 << 31) != 0;
    let diagonal_style = XlsBorderStyle::from_bits((border2 >> 21) & 0x0f)?;
    if (diagonal_down || diagonal_up) != (diagonal_style != XlsBorderStyle::None) {
        return Err(XlsError::InvalidData(
            "BIFF8 diagonal border direction and style contradict each other".to_string(),
        ));
    }

    let left_color = ((border1 >> 16) & 0x7f) as u16;
    let right_color = ((border1 >> 23) & 0x7f) as u16;
    let top_color = (border2 & 0x7f) as u16;
    let bottom_color = ((border2 >> 7) & 0x7f) as u16;
    let diagonal_color = ((border2 >> 14) & 0x7f) as u16;
    let foreground_color = (area & 0x7f) as u16;
    let background_color = ((area >> 7) & 0x7f) as u16;
    for (name, color) in [
        ("left border", left_color),
        ("right border", right_color),
        ("top border", top_color),
        ("bottom border", bottom_color),
        ("diagonal border", diagonal_color),
        ("fill foreground", foreground_color),
        ("fill background", background_color),
    ] {
        if color > 0x41 && color != 0x7f {
            return Err(XlsError::InvalidData(format!(
                "invalid BIFF8 {name} color index {color:#04x}"
            )));
        }
    }

    let borders = XlsCellBorders {
        left: XlsBorderSide {
            style: XlsBorderStyle::from_bits(border1 & 0x0f)?,
            color_index: left_color,
        },
        right: XlsBorderSide {
            style: XlsBorderStyle::from_bits((border1 >> 4) & 0x0f)?,
            color_index: right_color,
        },
        top: XlsBorderSide {
            style: XlsBorderStyle::from_bits((border1 >> 8) & 0x0f)?,
            color_index: top_color,
        },
        bottom: XlsBorderSide {
            style: XlsBorderStyle::from_bits((border1 >> 12) & 0x0f)?,
            color_index: bottom_color,
        },
        diagonal: XlsBorderSide {
            style: diagonal_style,
            color_index: diagonal_color,
        },
        diagonal_down,
        diagonal_up,
    };
    let fill = XlsCellFill {
        pattern: XlsFillPattern::from_bits(border2 >> 26)?,
        foreground_color_index: foreground_color,
        background_color_index: background_color,
    };
    Ok((borders, fill))
}

#[cfg(test)]
mod tests {
    use super::*;

    const BORDER_STYLES: [XlsBorderStyle; 14] = [
        XlsBorderStyle::None,
        XlsBorderStyle::Thin,
        XlsBorderStyle::Medium,
        XlsBorderStyle::Dashed,
        XlsBorderStyle::Dotted,
        XlsBorderStyle::Thick,
        XlsBorderStyle::Double,
        XlsBorderStyle::Hair,
        XlsBorderStyle::MediumDashed,
        XlsBorderStyle::DashDot,
        XlsBorderStyle::MediumDashDot,
        XlsBorderStyle::DashDotDot,
        XlsBorderStyle::MediumDashDotDot,
        XlsBorderStyle::SlantedDashDot,
    ];

    const FILL_PATTERNS: [XlsFillPattern; 19] = [
        XlsFillPattern::None,
        XlsFillPattern::Solid,
        XlsFillPattern::MediumGray,
        XlsFillPattern::DarkGray,
        XlsFillPattern::LightGray,
        XlsFillPattern::DarkHorizontal,
        XlsFillPattern::DarkVertical,
        XlsFillPattern::DarkDown,
        XlsFillPattern::DarkUp,
        XlsFillPattern::DarkGrid,
        XlsFillPattern::DarkTrellis,
        XlsFillPattern::LightHorizontal,
        XlsFillPattern::LightVertical,
        XlsFillPattern::LightDown,
        XlsFillPattern::LightUp,
        XlsFillPattern::LightGrid,
        XlsFillPattern::LightTrellis,
        XlsFillPattern::Gray125,
        XlsFillPattern::Gray0625,
    ];

    fn xf(border1: u32, border2: u32, area: u16) -> [u8; 20] {
        let mut data = [0; 20];
        data[10..14].copy_from_slice(&border1.to_le_bytes());
        data[14..18].copy_from_slice(&border2.to_le_bytes());
        data[18..20].copy_from_slice(&area.to_le_bytes());
        data
    }

    fn parse_cell(data: &[u8]) -> XlsResult<(XlsCellBorders, XlsCellFill)> {
        parse_xf_border_fill(data, false)
    }

    #[test]
    fn decodes_every_border_style() {
        for (bits, expected) in BORDER_STYLES.into_iter().enumerate() {
            assert_eq!(
                parse_cell(&xf(bits as u32, 0, 0))
                    .unwrap()
                    .0
                    .left()
                    .style(),
                expected
            );
        }
    }

    #[test]
    fn decodes_diagonal_directions_and_colors() {
        let border1 = (7 << 16) | (8 << 23) | (1 << 30) | (1 << 31);
        let border2 = 9 | (10 << 7) | (11 << 14) | (6 << 21);
        let area = 12 | (13 << 7);
        let (borders, fill) = parse_cell(&xf(border1, border2, area)).unwrap();
        assert!(borders.diagonal_down());
        assert!(borders.diagonal_up());
        assert_eq!(borders.diagonal().style(), XlsBorderStyle::Double);
        assert_eq!(borders.diagonal().color_index(), 11);
        assert_eq!(borders.left().color_index(), 7);
        assert_eq!(borders.right().color_index(), 8);
        assert_eq!(borders.top().color_index(), 9);
        assert_eq!(borders.bottom().color_index(), 10);
        assert_eq!(fill.foreground_color_index(), 12);
        assert_eq!(fill.background_color_index(), 13);
    }

    #[test]
    fn decodes_every_fill_pattern() {
        for (bits, expected) in FILL_PATTERNS.into_iter().enumerate() {
            assert_eq!(
                parse_cell(&xf(0, (bits as u32) << 26, 0))
                    .unwrap()
                    .1
                    .pattern(),
                expected
            );
        }
    }

    #[test]
    fn rejects_invalid_and_contradictory_encodings() {
        assert!(parse_cell(&xf(14, 0, 0)).is_err());
        assert!(parse_cell(&xf(0, 19 << 26, 0)).is_err());
        assert!(parse_cell(&xf(0, 1 << 21, 0)).is_err());
        assert!(parse_cell(&xf(1 << 30, 0, 0)).is_err());
        assert!(parse_cell(&xf(0x42 << 16, 0, 0)).is_err());
    }

    #[test]
    fn style_xf_ignores_must_zero_reserved_bits() {
        assert!(parse_xf_border_fill(&xf(0, 1 << 25, 0), true).is_ok());
        assert!(parse_xf_border_fill(&xf(0, 0, 1 << 14), true).is_ok());
        assert!(parse_xf_border_fill(&xf(0, 0, 1 << 15), true).is_ok());
    }

    #[test]
    fn cell_xf_preserves_extension_and_button_bits() {
        assert!(parse_xf_border_fill(&xf(0, 1 << 25, 0), false).is_ok());
        assert!(parse_xf_border_fill(&xf(0, 0, 1 << 14), false).is_ok());
    }

    #[test]
    fn resolves_palette_colors() {
        let (borders, fill) =
            parse_cell(&xf(8 << 16, 0, 9 | (10 << 7))).unwrap();
        let palette = XlsPalette::default();
        assert!(borders.left().color(&palette).is_some());
        assert!(fill.foreground_color(&palette).is_some());
        assert!(fill.background_color(&palette).is_some());
    }
}
