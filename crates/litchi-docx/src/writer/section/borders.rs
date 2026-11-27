//! `WordprocessingML` section page-border vocabulary.
//!
//! The section module owns the `sectPr` container; this focused module owns
//! the page-border model used by that container.

pub use super::{Art, Color, Display, OffsetFrom, Style, ZOrder};

/// One page-border edge (`CT_Border`, ECMA-376 §17.6.16).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Border {
    pub style: Style,
    /// Border size in eighths of a point for line styles, points for art.
    pub size: Option<u32>,
    /// Border offset space in points (`0..=31`).
    pub space: Option<u32>,
    pub color: Option<Color>,
    pub shadow: bool,
    pub frame: bool,
}

/// Section page-border settings (`CT_PageBorders`, ECMA-376 §17.6.16).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Borders {
    pub offset_from: OffsetFrom,
    pub z_order: ZOrder,
    pub display: Display,
    pub top: Option<Border>,
    pub left: Option<Border>,
    pub bottom: Option<Border>,
    pub right: Option<Border>,
}

impl Default for Borders {
    fn default() -> Self {
        Self {
            offset_from: OffsetFrom::Page,
            z_order: ZOrder::Back,
            display: Display::AllPages,
            top: None,
            left: None,
            bottom: None,
            right: None,
        }
    }
}
