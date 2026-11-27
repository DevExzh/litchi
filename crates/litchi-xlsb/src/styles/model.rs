//! Package-neutral XLSB style values.

use std::collections::HashMap;

/// Cell horizontal alignment from the BrtXF flags.
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
    /// Convert the wire value, using the general fallback.
    #[must_use]
    pub const fn from_u8(value: u8) -> Self {
        match value {
            1 => Self::Left,
            2 => Self::Center,
            3 => Self::Right,
            4 => Self::Fill,
            5 => Self::Justify,
            6 => Self::CenterContinuous,
            7 => Self::Distributed,
            _ => Self::General,
        }
    }
}

/// Cell vertical alignment from the BrtXF flags.
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
    /// Convert the wire value, using the top fallback.
    #[must_use]
    pub const fn from_u8(value: u8) -> Self {
        match value {
            1 => Self::Center,
            2 => Self::Bottom,
            3 => Self::Justify,
            4 => Self::Distributed,
            _ => Self::Top,
        }
    }
}

/// Compact, host-neutral alignment information carried by a cell format.
#[derive(Debug, Clone, PartialEq, Eq)]
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
        Self {
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

/// Border side style from a Blxf value.
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
    /// Convert the wire style, using None for unknown values.
    #[must_use]
    pub const fn from_u8(value: u8) -> Self {
        match value {
            1 => Self::Thin,
            2 => Self::Medium,
            3 => Self::Dashed,
            4 => Self::Dotted,
            5 => Self::Thick,
            6 => Self::Double,
            7 => Self::Hair,
            8 => Self::MediumDashed,
            9 => Self::DashDot,
            10 => Self::MediumDashDot,
            11 => Self::DashDotDot,
            12 => Self::MediumDashDotDot,
            13 => Self::SlantDashDot,
            _ => Self::None,
        }
    }
}

/// One border side and its optional direct ARGB color.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct BorderSide {
    pub style: BorderStyle,
    pub color: Option<u32>,
}

/// Compact border representation for the five BrtBorder sides.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
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

/// Font information used by the XLSB styles part.
#[derive(Debug, Clone, PartialEq)]
pub struct Font {
    pub name: String,
    pub size: f64,
    pub color: Option<u32>,
    pub bold: bool,
    pub italic: bool,
    pub underline: bool,
    pub strike: bool,
}

impl Default for Font {
    fn default() -> Self {
        Self {
            name: "Calibri".to_string(),
            size: 11.0,
            color: None,
            bold: false,
            italic: false,
            underline: false,
            strike: false,
        }
    }
}

/// Pattern fill information from a BrtFill record.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct Fill {
    pub pattern_type: u32,
    pub fg_color: Option<u32>,
    pub bg_color: Option<u32>,
}

/// A number format declared by a BrtFmt record.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct NumberFormat {
    pub id: u32,
    pub format_code: String,
}

/// The compact formatting fields retained from a BrtXF record.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct CellFormat {
    pub font_id: u32,
    pub fill_id: u32,
    pub border_id: u32,
    pub num_fmt_id: u32,
    pub alignment: Option<Alignment>,
}

/// A complete styles part model.
#[derive(Debug, Clone)]
pub struct Table {
    pub fonts: Vec<Font>,
    pub fills: Vec<Fill>,
    pub borders: Vec<Border>,
    pub num_fmts: HashMap<u32, String>,
    pub cell_xfs: Vec<CellFormat>,
    pub cell_style_xfs: Vec<CellFormat>,
}

impl Default for Table {
    fn default() -> Self {
        Self {
            fonts: vec![Font::default()],
            fills: vec![Fill::default()],
            borders: vec![Border::default()],
            num_fmts: Self::builtin_formats(),
            cell_xfs: vec![CellFormat::default()],
            cell_style_xfs: vec![CellFormat::default()],
        }
    }
}

impl Table {
    /// Get a cell XF by zero-based index.
    #[must_use]
    pub fn get_cell_format(&self, index: usize) -> Option<&CellFormat> {
        self.cell_xfs.get(index)
    }

    /// Get a font by zero-based index.
    #[must_use]
    pub fn get_font(&self, index: usize) -> Option<&Font> {
        self.fonts.get(index)
    }

    /// Get a fill by zero-based index.
    #[must_use]
    pub fn get_fill(&self, index: usize) -> Option<&Fill> {
        self.fills.get(index)
    }

    /// Get a border by zero-based index.
    #[must_use]
    pub fn get_border(&self, index: usize) -> Option<&Border> {
        self.borders.get(index)
    }

    /// Get a number-format code by identifier.
    #[must_use]
    pub fn get_num_fmt(&self, id: u32) -> Option<&str> {
        self.num_fmts.get(&id).map(String::as_str)
    }

    /// Check whether a built-in or custom format is date-like.
    #[must_use]
    pub fn is_date_format(&self, num_fmt_id: u32) -> bool {
        if matches!(num_fmt_id, 14..=22 | 27..=36 | 45..=47 | 50..=58) {
            return true;
        }

        self.get_num_fmt(num_fmt_id)
            .map(|format_code| {
                let format_lower = format_code.to_lowercase();
                format_lower.contains('y')
                    || format_lower.contains('m')
                    || format_lower.contains('d')
                    || format_lower.contains('h')
                    || format_lower.contains('s')
            })
            .unwrap_or(false)
    }

    fn builtin_formats() -> HashMap<u32, String> {
        [
            (0, "General"),
            (1, "0"),
            (2, "0.00"),
            (3, "#,##0"),
            (4, "#,##0.00"),
            (9, "0%"),
            (10, "0.00%"),
            (11, "0.00E+00"),
            (12, "# ?/?"),
            (13, "# ??/??"),
            (14, "mm-dd-yy"),
            (15, "d-mmm-yy"),
            (16, "d-mmm"),
            (17, "mmm-yy"),
            (18, "h:mm AM/PM"),
            (19, "h:mm:ss AM/PM"),
            (20, "h:mm"),
            (21, "h:mm:ss"),
            (22, "m/d/yy h:mm"),
            (37, "#,##0 ;(#,##0)"),
            (38, "#,##0 ;[Red](#,##0)"),
            (39, "#,##0.00;(#,##0.00)"),
            (40, "#,##0.00;[Red](#,##0.00)"),
            (45, "mm:ss"),
            (46, "[h]:mm:ss"),
            (47, "mmss.0"),
            (48, "##0.0E+0"),
            (49, "@"),
        ]
        .into_iter()
        .map(|(id, code)| (id, code.to_string()))
        .collect()
    }
}
