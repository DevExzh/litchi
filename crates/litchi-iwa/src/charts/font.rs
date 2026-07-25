//! Chart-wide font identity, face-trait, and point-size CRUD.

mod native;

use crate::Result;
use crate::text::{TextFont, TextPointSize};

pub(crate) use native::{
    chart_font, chart_font_size, reset_chart_font, reset_chart_font_size, set_chart_font,
    set_chart_font_size,
};

/// Effective chart-wide font identity and face traits.
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct ChartFont {
    font: TextFont,
    bold: bool,
    italic: bool,
}

impl ChartFont {
    /// Construct a chart font with regular face traits.
    pub const fn new(font: TextFont) -> Self {
        Self {
            font,
            bold: false,
            italic: false,
        }
    }

    /// Construct a named chart font from a validated PostScript identifier.
    pub fn named(name: impl Into<String>) -> Result<Self> {
        TextFont::named(name).map(Self::new)
    }

    /// Borrow the effective font identity.
    pub const fn font(&self) -> &TextFont {
        &self.font
    }

    /// Whether the chart uses bold face traits.
    pub const fn bold(&self) -> bool {
        self.bold
    }

    /// Whether the chart uses italic face traits.
    pub const fn italic(&self) -> bool {
        self.italic
    }

    /// Enable or disable bold face traits.
    pub const fn with_bold(mut self, bold: bool) -> Self {
        self.bold = bold;
        self
    }

    /// Enable or disable italic face traits.
    pub const fn with_italic(mut self, italic: bool) -> Self {
        self.italic = italic;
        self
    }
}

/// Positive chart-wide character size in typographic points.
#[derive(Debug, Clone, Copy, PartialEq, PartialOrd, Default)]
pub struct ChartFontSize(TextPointSize);

impl ChartFontSize {
    /// Native default used by source-created charts.
    pub const TWELVE: Self = Self(TextPointSize::TWELVE);

    /// Construct a finite chart character size greater than zero.
    pub fn from_points(points: f32) -> Result<Self> {
        TextPointSize::from_points(points).map(Self)
    }

    /// Return the character size in typographic points.
    pub const fn points(self) -> f32 {
        self.0.points()
    }

    pub(crate) const fn text_point_size(self) -> TextPointSize {
        self.0
    }
}

impl From<TextPointSize> for ChartFontSize {
    fn from(size: TextPointSize) -> Self {
        Self(size)
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn chart_font_values_are_strict_and_typed() {
        let font = ChartFont::named("AvenirNext-DemiBold")
            .unwrap()
            .with_bold(true);
        assert_eq!(font.font().name(), Some("AvenirNext-DemiBold"));
        assert!(font.bold());
        assert!(!font.italic());
        assert!(ChartFont::named(" AvenirNext-Regular").is_err());

        assert_eq!(ChartFontSize::TWELVE.points(), 12.0);
        assert_eq!(ChartFontSize::from_points(13.3).unwrap().points(), 13.3);
        assert!(ChartFontSize::from_points(0.0).is_err());
        assert!(ChartFontSize::from_points(f32::NAN).is_err());
    }
}
