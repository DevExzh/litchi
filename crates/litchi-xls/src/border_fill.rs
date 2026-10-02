//! BIFF8 XF border and fill metadata.

use crate::leniency::{CellFormatDefect, ToleranceReport};
use crate::{Color, Error, Palette, Result};

/// A BIFF8 cell border line style.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
#[non_exhaustive]
pub enum BorderStyle {
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

impl BorderStyle {
    /// The style `value` names, or [`BorderStyle::None`] — the default
    /// LibreOffice substitutes — recorded as
    /// [`CellFormatDefect::BorderStyle`] against XF `xf_index`.
    fn repaired(value: u32, xf_index: u16, tolerance: &mut ToleranceReport) -> Self {
        Self::from_bits(value).unwrap_or_else(|_| {
            tolerance.record(CellFormatDefect::BorderStyle, u32::from(xf_index), value);
            Self::None
        })
    }

    fn from_bits(value: u32) -> Result<Self> {
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
            _ => Err(Error::InvalidData(format!(
                "reserved BIFF8 border style {value}"
            ))),
        }
    }
}

/// The style and indexed color of one cell border edge.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct BorderSide {
    style: BorderStyle,
    color_index: u16,
}

impl BorderSide {
    #[must_use]
    pub fn style(&self) -> BorderStyle {
        self.style
    }

    #[must_use]
    pub fn color_index(&self) -> u16 {
        self.color_index
    }

    #[must_use]
    pub fn color(&self, palette: &Palette) -> Option<Color> {
        palette.color(self.color_index)
    }
}

/// All border metadata stored by a BIFF8 XF record.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct CellBorders {
    left: BorderSide,
    right: BorderSide,
    top: BorderSide,
    bottom: BorderSide,
    diagonal: BorderSide,
    diagonal_down: bool,
    diagonal_up: bool,
}

impl CellBorders {
    #[must_use]
    pub fn left(&self) -> &BorderSide {
        &self.left
    }

    #[must_use]
    pub fn right(&self) -> &BorderSide {
        &self.right
    }

    #[must_use]
    pub fn top(&self) -> &BorderSide {
        &self.top
    }

    #[must_use]
    pub fn bottom(&self) -> &BorderSide {
        &self.bottom
    }

    #[must_use]
    pub fn diagonal(&self) -> &BorderSide {
        &self.diagonal
    }

    /// Returns whether the diagonal runs from top-left to bottom-right.
    #[must_use]
    pub fn diagonal_down(&self) -> bool {
        self.diagonal_down
    }

    /// Returns whether the diagonal runs from bottom-left to top-right.
    #[must_use]
    pub fn diagonal_up(&self) -> bool {
        self.diagonal_up
    }
}

/// A BIFF8 cell fill pattern.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
#[non_exhaustive]
pub enum FillPattern {
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

impl FillPattern {
    /// The pattern `value` names, or [`FillPattern::None`] — the default
    /// LibreOffice substitutes — recorded as
    /// [`CellFormatDefect::FillPattern`] against XF `xf_index`.
    fn repaired(value: u32, xf_index: u16, tolerance: &mut ToleranceReport) -> Self {
        Self::from_bits(value).unwrap_or_else(|_| {
            tolerance.record(CellFormatDefect::FillPattern, u32::from(xf_index), value);
            Self::None
        })
    }

    fn from_bits(value: u32) -> Result<Self> {
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
            _ => Err(Error::InvalidData(format!(
                "reserved BIFF8 fill pattern {value}"
            ))),
        }
    }
}

/// Fill pattern and foreground/background indexed colors from a BIFF8 XF.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct CellFill {
    pattern: FillPattern,
    foreground_color_index: u16,
    background_color_index: u16,
}

impl CellFill {
    #[must_use]
    pub fn pattern(&self) -> FillPattern {
        self.pattern
    }

    #[must_use]
    pub fn foreground_color_index(&self) -> u16 {
        self.foreground_color_index
    }

    #[must_use]
    pub fn background_color_index(&self) -> u16 {
        self.background_color_index
    }

    #[must_use]
    pub fn foreground_color(&self, palette: &Palette) -> Option<Color> {
        palette.color(self.foreground_color_index)
    }

    #[must_use]
    pub fn background_color(&self, palette: &Palette) -> Option<Color> {
        palette.color(self.background_color_index)
    }
}

/// The automatic 7-bit border/fill color index, the substitute for an index
/// outside the palette.
const AUTOMATIC_COLOR: u16 = 0x7f;

/// Parse the border and fill fields of the XF `xf_index`, applying the
/// Excel-compatible repairs under every policy and recording each as a
/// [`CellFormatDefect`]: an undefined border style or fill pattern reads as
/// none, a color index outside the palette as automatic (0x7F), and a
/// diagonal whose direction flags and style contradict each other is kept
/// exactly as stored (it draws no line, as in LibreOffice).
pub(crate) fn parse_xf_border_fill(
    data: &[u8],
    xf_index: u16,
    tolerance: &mut ToleranceReport,
) -> (CellBorders, CellFill) {
    debug_assert!(data.len() >= 20);
    let border1 = u32::from_le_bytes([data[10], data[11], data[12], data[13]]);
    let border2 = u32::from_le_bytes([data[14], data[15], data[16], data[17]]);
    let area = u16::from_le_bytes([data[18], data[19]]);

    let diagonal_down = border1 & (1 << 30) != 0;
    let diagonal_up = border1 & (1 << 31) != 0;
    let diagonal_style = BorderStyle::repaired((border2 >> 21) & 0x0f, xf_index, tolerance);
    if (diagonal_down || diagonal_up) != (diagonal_style != BorderStyle::None) {
        tolerance.record(
            CellFormatDefect::DiagonalBorder,
            u32::from(xf_index),
            (border1 >> 30) | ((border2 >> 21) & 0x0f) << 2,
        );
    }

    let mut color = |value: u16| -> u16 {
        if value > 0x41 && value != AUTOMATIC_COLOR {
            tolerance.record(
                CellFormatDefect::ColorIndex,
                u32::from(xf_index),
                u32::from(value),
            );
            AUTOMATIC_COLOR
        } else {
            value
        }
    };
    let left_color = color(((border1 >> 16) & 0x7f) as u16);
    let right_color = color(((border1 >> 23) & 0x7f) as u16);
    let top_color = color((border2 & 0x7f) as u16);
    let bottom_color = color(((border2 >> 7) & 0x7f) as u16);
    let diagonal_color = color(((border2 >> 14) & 0x7f) as u16);
    let foreground_color = color(area & 0x7f);
    let background_color = color((area >> 7) & 0x7f);

    let borders = CellBorders {
        left: BorderSide {
            style: BorderStyle::repaired(border1 & 0x0f, xf_index, tolerance),
            color_index: left_color,
        },
        right: BorderSide {
            style: BorderStyle::repaired((border1 >> 4) & 0x0f, xf_index, tolerance),
            color_index: right_color,
        },
        top: BorderSide {
            style: BorderStyle::repaired((border1 >> 8) & 0x0f, xf_index, tolerance),
            color_index: top_color,
        },
        bottom: BorderSide {
            style: BorderStyle::repaired((border1 >> 12) & 0x0f, xf_index, tolerance),
            color_index: bottom_color,
        },
        diagonal: BorderSide {
            style: diagonal_style,
            color_index: diagonal_color,
        },
        diagonal_down,
        diagonal_up,
    };
    let fill = CellFill {
        pattern: FillPattern::repaired(border2 >> 26, xf_index, tolerance),
        foreground_color_index: foreground_color,
        background_color_index: background_color,
    };
    (borders, fill)
}

#[cfg(test)]
mod tests {
    use super::*;

    const BORDER_STYLES: [BorderStyle; 14] = [
        BorderStyle::None,
        BorderStyle::Thin,
        BorderStyle::Medium,
        BorderStyle::Dashed,
        BorderStyle::Dotted,
        BorderStyle::Thick,
        BorderStyle::Double,
        BorderStyle::Hair,
        BorderStyle::MediumDashed,
        BorderStyle::DashDot,
        BorderStyle::MediumDashDot,
        BorderStyle::DashDotDot,
        BorderStyle::MediumDashDotDot,
        BorderStyle::SlantedDashDot,
    ];

    const FILL_PATTERNS: [FillPattern; 19] = [
        FillPattern::None,
        FillPattern::Solid,
        FillPattern::MediumGray,
        FillPattern::DarkGray,
        FillPattern::LightGray,
        FillPattern::DarkHorizontal,
        FillPattern::DarkVertical,
        FillPattern::DarkDown,
        FillPattern::DarkUp,
        FillPattern::DarkGrid,
        FillPattern::DarkTrellis,
        FillPattern::LightHorizontal,
        FillPattern::LightVertical,
        FillPattern::LightDown,
        FillPattern::LightUp,
        FillPattern::LightGrid,
        FillPattern::LightTrellis,
        FillPattern::Gray125,
        FillPattern::Gray0625,
    ];

    fn xf(border1: u32, border2: u32, area: u16) -> [u8; 20] {
        let mut data = [0; 20];
        data[10..14].copy_from_slice(&border1.to_le_bytes());
        data[14..18].copy_from_slice(&border2.to_le_bytes());
        data[18..20].copy_from_slice(&area.to_le_bytes());
        data
    }

    /// Conforming-only shim: refuses anything the parser had to repair.
    fn parse_cell(data: &[u8]) -> Result<(CellBorders, CellFill)> {
        let mut report = ToleranceReport::default();
        let parsed = parse_xf_border_fill(data, 0, &mut report);
        if report.is_clean() {
            Ok(parsed)
        } else {
            Err(Error::InvalidData(format!("repaired: {report:?}")))
        }
    }

    fn parse_repaired(data: &[u8]) -> ((CellBorders, CellFill), ToleranceReport) {
        let mut report = ToleranceReport::default();
        let parsed = parse_xf_border_fill(data, 33, &mut report);
        (parsed, report)
    }

    #[test]
    fn decodes_every_border_style() {
        for (bits, expected) in BORDER_STYLES.into_iter().enumerate() {
            assert_eq!(
                parse_cell(&xf(bits as u32, 0, 0)).unwrap().0.left().style(),
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
        assert_eq!(borders.diagonal().style(), BorderStyle::Double);
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

    fn single_repair(report: &ToleranceReport) -> (CellFormatDefect, u32) {
        assert_eq!(report.total(), 1, "{report:?}");
        let entry = report.defects()[0];
        assert_eq!(entry.ordinal(), 33, "the XF index");
        let crate::leniency::Defect::CellFormat(defect) = entry.defect() else {
            panic!("unexpected class {entry:?}");
        };
        (defect, entry.observed())
    }

    #[test]
    fn an_undefined_border_style_reads_as_none() {
        let ((borders, _), report) = parse_repaired(&xf(14 << 4, 0, 0));
        assert_eq!(borders.right().style(), BorderStyle::None);
        assert_eq!(single_repair(&report), (CellFormatDefect::BorderStyle, 14));
    }

    #[test]
    fn an_undefined_fill_pattern_reads_as_none() {
        let ((_, fill), report) = parse_repaired(&xf(0, 19 << 26, 0));
        assert_eq!(fill.pattern(), FillPattern::None);
        assert_eq!(single_repair(&report), (CellFormatDefect::FillPattern, 19));
    }

    #[test]
    fn an_out_of_palette_color_reads_as_automatic() {
        let ((borders, _), report) = parse_repaired(&xf(0x42 << 16, 0, 0));
        assert_eq!(borders.left().color_index(), 0x7f);
        assert_eq!(single_repair(&report), (CellFormatDefect::ColorIndex, 0x42));
        let ((_, fill), report) = parse_repaired(&xf(0, 0, 0x50 << 7));
        assert_eq!(fill.background_color_index(), 0x7f);
        assert_eq!(single_repair(&report), (CellFormatDefect::ColorIndex, 0x50));
    }

    #[test]
    fn a_contradictory_diagonal_is_kept_as_stored() {
        // A style without a direction.
        let ((borders, _), report) = parse_repaired(&xf(0, 1 << 21, 0));
        assert_eq!(borders.diagonal().style(), BorderStyle::Thin);
        assert!(!borders.diagonal_down() && !borders.diagonal_up());
        assert_eq!(
            single_repair(&report),
            (CellFormatDefect::DiagonalBorder, 1 << 2)
        );
        // A direction without a style.
        let ((borders, _), report) = parse_repaired(&xf(1 << 30, 0, 0));
        assert_eq!(borders.diagonal().style(), BorderStyle::None);
        assert!(borders.diagonal_down());
        assert_eq!(
            single_repair(&report),
            (CellFormatDefect::DiagonalBorder, 1)
        );
    }

    #[test]
    fn style_xf_ignores_must_zero_reserved_bits() {
        assert!(parse_cell(&xf(0, 1 << 25, 0)).is_ok());
        assert!(parse_cell(&xf(0, 0, 1 << 14)).is_ok());
        assert!(parse_cell(&xf(0, 0, 1 << 15)).is_ok());
    }

    #[test]
    fn resolves_palette_colors() {
        let (borders, fill) = parse_cell(&xf(8 << 16, 0, 9 | (10 << 7))).unwrap();
        let palette = Palette::default();
        assert!(borders.left().color(&palette).is_some());
        assert!(fill.foreground_color(&palette).is_some());
        assert!(fill.background_color(&palette).is_some());
    }
}
