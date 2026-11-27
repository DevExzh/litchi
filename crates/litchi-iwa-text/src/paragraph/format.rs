//! Archive-free paragraph formatting values shared by iWork formats.
//!
//! This module contains only semantic inspector values. Native discriminants,
//! protobuf presence, inheritance, and package mutation remain in the concrete
//! IWA adapter.

use litchi_iwa_common::{
    color::Rgba,
    shape::stroke::{Pattern, Width},
};

use super::{
    border::{Offset, Sides},
    direction::WritingDirection,
    flow::Flow,
    tabs::{DecimalCharacter, DefaultInterval, Stops},
};
use crate::appearance::ParagraphBackground;

/// Validation failures produced by paragraph formatting constructors.
#[derive(Debug, Clone, Copy, PartialEq, Eq, thiserror::Error)]
#[non_exhaustive]
pub enum Error {
    /// A visible paragraph border has no selected edge.
    #[error("a visible paragraph border must contain at least one side")]
    EmptyBorder,
    /// Rounded corners are meaningful only when all four edges are present.
    #[error("rounded paragraph-border corners require all four sides")]
    RoundedCornersRequireAllSides,
    /// An indentation distance is NaN or infinite.
    #[error("paragraph indentation must be finite")]
    IndentNonFinite,
    /// An indentation distance is negative.
    #[error("paragraph indentation must be nonnegative")]
    IndentNegative,
    /// Paragraph spacing is NaN or infinite.
    #[error("paragraph spacing must be finite")]
    SpacingNonFinite,
    /// Paragraph spacing is negative.
    #[error("paragraph spacing must be nonnegative")]
    SpacingNegative,
    /// Relative line spacing is NaN or infinite.
    #[error("paragraph line-spacing multiplier must be finite")]
    LineSpacingMultipleNonFinite,
    /// Relative line spacing is zero or negative.
    #[error("paragraph line-spacing multiplier must be greater than zero")]
    LineSpacingMultipleNonPositive,
    /// Absolute line spacing is NaN or infinite.
    #[error("paragraph line-spacing distance must be finite")]
    LineSpacingPointsNonFinite,
    /// Absolute line spacing is zero or negative.
    #[error("paragraph line-spacing distance must be greater than zero")]
    LineSpacingPointsNonPositive,
}

/// Result type for checked paragraph formatting values.
pub type Result<T> = std::result::Result<T, Error>;

/// One uniform paragraph-border appearance.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct Border {
    color: Rgba,
    width: Width,
    pattern: Pattern,
    sides: Sides,
    offset: Offset,
    rounded_corners: bool,
}

impl Border {
    /// Construct a paragraph border accepted by the native iWork inspector.
    ///
    /// The scalar color, width, and offset values are already validated by
    /// their respective semantic leaves before they reach this constructor.
    ///
    /// # Errors
    ///
    /// Returns [`Error::EmptyBorder`] when no edge is selected, or
    /// [`Error::RoundedCornersRequireAllSides`] when rounded corners are
    /// requested without all four edges.
    pub fn new(
        color: Rgba,
        width: Width,
        pattern: Pattern,
        sides: Sides,
        offset: Offset,
        rounded_corners: bool,
    ) -> Result<Self> {
        if sides.is_empty() {
            return Err(Error::EmptyBorder);
        }
        if rounded_corners && !sides.is_all() {
            return Err(Error::RoundedCornersRequireAllSides);
        }
        Ok(Self {
            color,
            width,
            pattern,
            sides,
            offset,
            rounded_corners,
        })
    }

    /// Return the border color.
    #[must_use]
    pub const fn color(self) -> Rgba {
        self.color
    }

    /// Return the border width.
    #[must_use]
    pub const fn width(self) -> Width {
        self.width
    }

    /// Return the border pattern.
    #[must_use]
    pub const fn pattern(self) -> Pattern {
        self.pattern
    }

    /// Return the selected border edges.
    #[must_use]
    pub const fn sides(self) -> Sides {
        self.sides
    }

    /// Return the gap between paragraph text and the border.
    #[must_use]
    pub const fn offset(self) -> Offset {
        self.offset
    }

    /// Return whether the border uses rounded corners.
    #[must_use]
    pub const fn has_rounded_corners(self) -> bool {
        self.rounded_corners
    }
}

/// Effective paragraph-border setting.
#[derive(Debug, Clone, Copy, PartialEq, Default)]
pub enum Borders {
    /// Do not paint a paragraph border.
    #[default]
    None,
    /// Paint one uniform border around the selected edges.
    Bordered(Border),
}

/// Nonnegative paragraph indentation measured in typographic points.
#[repr(transparent)]
#[derive(Debug, Clone, Copy, PartialEq, PartialOrd, Default)]
pub struct IndentPoints(f32);

impl IndentPoints {
    /// Zero indentation.
    pub const ZERO: Self = Self(0.0);

    /// Construct a finite, nonnegative indentation distance.
    ///
    /// # Errors
    ///
    /// Returns [`Error::IndentNonFinite`] for NaN or infinity and
    /// [`Error::IndentNegative`] for a negative distance.
    pub fn from_points(points: f32) -> Result<Self> {
        if !points.is_finite() {
            return Err(Error::IndentNonFinite);
        }
        if points < 0.0 {
            return Err(Error::IndentNegative);
        }
        Ok(if points == 0.0 {
            Self::ZERO
        } else {
            Self(points)
        })
    }

    /// Return the indentation distance in points.
    #[must_use]
    pub const fn points(self) -> f32 {
        self.0
    }
}

/// First-line, left, and right paragraph indentation.
///
/// `first_line` and `left` are absolute distances from the left text
/// boundary. A hanging indent therefore has `first_line < left`.
#[repr(C)]
#[derive(Debug, Clone, Copy, PartialEq, Default)]
pub struct Indents {
    /// Distance of the first line from the left text boundary.
    pub first_line: IndentPoints,
    /// Distance of subsequent lines from the left text boundary.
    pub left: IndentPoints,
    /// Distance of all lines from the right text boundary.
    pub right: IndentPoints,
}

impl Indents {
    /// No indentation on any edge.
    pub const NONE: Self = Self {
        first_line: IndentPoints::ZERO,
        left: IndentPoints::ZERO,
        right: IndentPoints::ZERO,
    };

    /// Construct first-line, left, and right indentation values.
    #[must_use]
    pub const fn new(first_line: IndentPoints, left: IndentPoints, right: IndentPoints) -> Self {
        Self {
            first_line,
            left,
            right,
        }
    }
}

/// Nonnegative paragraph spacing measured in typographic points.
#[repr(transparent)]
#[derive(Debug, Clone, Copy, PartialEq, PartialOrd, Default)]
pub struct SpacingPoints(f32);

impl SpacingPoints {
    /// Zero paragraph spacing.
    pub const ZERO: Self = Self(0.0);

    /// Construct a finite, nonnegative paragraph-spacing distance.
    ///
    /// # Errors
    ///
    /// Returns [`Error::SpacingNonFinite`] for NaN or infinity and
    /// [`Error::SpacingNegative`] for a negative distance.
    pub fn from_points(points: f32) -> Result<Self> {
        if !points.is_finite() {
            return Err(Error::SpacingNonFinite);
        }
        if points < 0.0 {
            return Err(Error::SpacingNegative);
        }
        Ok(if points == 0.0 {
            Self::ZERO
        } else {
            Self(points)
        })
    }

    /// Return the spacing distance in points.
    #[must_use]
    pub const fn points(self) -> f32 {
        self.0
    }
}

/// Space inserted before and after a paragraph.
#[repr(C)]
#[derive(Debug, Clone, Copy, PartialEq, Default)]
pub struct Spacing {
    /// Space before the paragraph.
    pub before: SpacingPoints,
    /// Space after the paragraph.
    pub after: SpacingPoints,
}

impl Spacing {
    /// No space before or after the paragraph.
    pub const NONE: Self = Self {
        before: SpacingPoints::ZERO,
        after: SpacingPoints::ZERO,
    };

    /// Construct before and after spacing values.
    #[must_use]
    pub const fn new(before: SpacingPoints, after: SpacingPoints) -> Self {
        Self { before, after }
    }
}

/// Positive multiplier used by relative paragraph line spacing.
#[repr(transparent)]
#[derive(Debug, Clone, Copy, PartialEq, PartialOrd)]
pub struct LineSpacingMultiple(f32);

impl LineSpacingMultiple {
    /// Single line spacing.
    pub const SINGLE: Self = Self(1.0);
    /// One-and-a-fifth line spacing.
    pub const ONE_POINT_TWO: Self = Self(1.2);
    /// One-and-a-half line spacing.
    pub const ONE_POINT_FIVE: Self = Self(1.5);
    /// Double line spacing.
    pub const DOUBLE: Self = Self(2.0);

    /// Construct a finite multiplier greater than zero.
    ///
    /// # Errors
    ///
    /// Returns [`Error::LineSpacingMultipleNonFinite`] for NaN or infinity
    /// and [`Error::LineSpacingMultipleNonPositive`] for zero or a negative
    /// multiplier.
    pub fn new(value: f32) -> Result<Self> {
        if !value.is_finite() {
            return Err(Error::LineSpacingMultipleNonFinite);
        }
        if value <= 0.0 {
            return Err(Error::LineSpacingMultipleNonPositive);
        }
        Ok(Self(value))
    }

    /// Return the relative line-spacing multiplier.
    #[must_use]
    pub const fn get(self) -> f32 {
        self.0
    }
}

impl Default for LineSpacingMultiple {
    fn default() -> Self {
        Self::SINGLE
    }
}

/// Positive absolute line-spacing distance in typographic points.
#[repr(transparent)]
#[derive(Debug, Clone, Copy, PartialEq, PartialOrd)]
pub struct LineSpacingPoints(f32);

impl LineSpacingPoints {
    /// Construct a finite line-spacing distance greater than zero.
    ///
    /// # Errors
    ///
    /// Returns [`Error::LineSpacingPointsNonFinite`] for NaN or infinity and
    /// [`Error::LineSpacingPointsNonPositive`] for zero or a negative
    /// distance.
    pub fn from_points(points: f32) -> Result<Self> {
        if !points.is_finite() {
            return Err(Error::LineSpacingPointsNonFinite);
        }
        if points <= 0.0 {
            return Err(Error::LineSpacingPointsNonPositive);
        }
        Ok(Self(points))
    }

    /// Return the absolute line-spacing distance in points.
    #[must_use]
    pub const fn points(self) -> f32 {
        self.0
    }
}

/// Native paragraph line-spacing modes.
#[derive(Debug, Clone, Copy, PartialEq)]
pub enum LineSpacing {
    /// Scale the font's natural line height by a positive multiplier.
    Relative(LineSpacingMultiple),
    /// Use at least the specified point distance.
    AtLeast(LineSpacingPoints),
    /// Use exactly the specified point distance.
    Exactly(LineSpacingPoints),
    /// Cap the line distance at the specified point distance.
    Maximum(LineSpacingPoints),
    /// Add the specified point distance between adjacent lines.
    Between(LineSpacingPoints),
}

impl Default for LineSpacing {
    fn default() -> Self {
        Self::Relative(LineSpacingMultiple::SINGLE)
    }
}

/// Horizontal paragraph alignment.
#[repr(u8)]
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
pub enum Alignment {
    /// Follow the paragraph writing direction.
    #[default]
    Natural,
    /// Align to the left edge.
    Left,
    /// Center the paragraph.
    Center,
    /// Align to the right edge.
    Right,
    /// Stretch text to both edges.
    Justified,
}

/// Complete archive-free paragraph formatting.
#[derive(Debug, Clone, PartialEq, Default)]
pub struct Format {
    /// Solid fill painted across the paragraph's layout box.
    pub background: ParagraphBackground,
    /// Lines painted around the paragraph's layout box.
    pub borders: Borders,
    /// Pagination, break, widow/orphan, and hyphenation behavior.
    pub flow: Flow,
    /// Base direction used to lay out bidirectional paragraph text.
    pub writing_direction: WritingDirection,
    /// Horizontal paragraph alignment.
    pub alignment: Alignment,
    /// Native line-spacing mode and amount.
    pub line_spacing: LineSpacing,
    /// Space inserted before and after the paragraph.
    pub spacing: Spacing,
    /// First-line, left, and right paragraph indentation.
    pub indents: Indents,
    /// Character used to align decimal tab stops.
    pub decimal_tab_character: DecimalCharacter,
    /// Distance between implicit paragraph tab stops.
    pub default_tab_interval: DefaultInterval,
    /// Explicit ruler tab stops inherited by the paragraph.
    pub tab_stops: Stops,
}

impl Format {
    /// Construct the native semantic defaults.
    #[must_use]
    pub fn new() -> Self {
        Self::default()
    }
}

#[cfg(test)]
mod tests {
    use std::mem::size_of;

    use super::*;

    #[test]
    fn formatting_scalars_are_strict_and_compact() {
        assert_eq!(size_of::<IndentPoints>(), size_of::<f32>());
        assert_eq!(size_of::<SpacingPoints>(), size_of::<f32>());
        assert_eq!(size_of::<LineSpacingMultiple>(), size_of::<f32>());
        assert_eq!(size_of::<LineSpacingPoints>(), size_of::<f32>());
        assert_eq!(IndentPoints::from_points(-0.0).unwrap().points(), 0.0);
        assert_eq!(SpacingPoints::from_points(12.5).unwrap().points(), 12.5);
        assert!(matches!(
            IndentPoints::from_points(f32::NAN),
            Err(Error::IndentNonFinite)
        ));
        assert!(matches!(
            SpacingPoints::from_points(-1.0),
            Err(Error::SpacingNegative)
        ));
        assert!(matches!(
            LineSpacingMultiple::new(0.0),
            Err(Error::LineSpacingMultipleNonPositive)
        ));
        assert!(matches!(
            LineSpacingPoints::from_points(f32::INFINITY),
            Err(Error::LineSpacingPointsNonFinite)
        ));
    }

    #[test]
    fn border_validation_stays_in_the_semantic_leaf() {
        let width = Width::new(1.0).unwrap();
        assert!(matches!(
            Border::new(
                Rgba::black(),
                width,
                Pattern::Solid,
                Sides::NONE,
                Offset::DEFAULT,
                false,
            ),
            Err(Error::EmptyBorder)
        ));
        assert!(matches!(
            Border::new(
                Rgba::black(),
                width,
                Pattern::Solid,
                Sides::TOP,
                Offset::DEFAULT,
                true,
            ),
            Err(Error::RoundedCornersRequireAllSides)
        ));
        let border = Border::new(
            Rgba::black(),
            width,
            Pattern::Solid,
            Sides::ALL,
            Offset::DEFAULT,
            true,
        )
        .unwrap();
        assert!(border.has_rounded_corners());
        assert_eq!(border.sides(), Sides::ALL);
    }

    #[test]
    fn aggregate_defaults_are_archive_free() {
        let format = Format::new();
        assert_eq!(format.borders, Borders::None);
        assert_eq!(format.alignment, Alignment::Natural);
        assert_eq!(format.line_spacing, LineSpacing::default());
        assert_eq!(format.spacing, Spacing::NONE);
        assert_eq!(format.indents, Indents::NONE);
        assert!(format.tab_stops.is_empty());
    }
}
