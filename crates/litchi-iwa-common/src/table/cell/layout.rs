//! Dependency-free text-layout vocabulary for table cells.

/// Validation failures for a table-cell inset.
#[derive(Clone, Copy, Debug, PartialEq, Eq, thiserror::Error)]
pub enum Error {
    /// The supplied distance was NaN or infinite.
    #[error("table-cell inset must be finite")]
    NonFinite,
    /// The supplied distance was less than zero.
    #[error("table-cell inset must be non-negative")]
    Negative,
}

/// Result type for table-cell layout construction.
pub type Result<T> = std::result::Result<T, Error>;

/// Whether text remains on one line or wraps inside its table cell.
#[repr(u8)]
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
pub enum TextWrap {
    /// Keep text on one line.
    #[default]
    Unwrapped,
    /// Wrap text within the cell's content box.
    Wrapped,
}

/// Vertical placement of text inside a table cell.
#[repr(u8)]
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
pub enum VerticalAlignment {
    /// Place text at the top edge of the content box.
    #[default]
    Top,
    /// Center text vertically in the content box.
    Middle,
    /// Place text at the bottom edge of the content box.
    Bottom,
}

/// Finite, non-negative distance from one cell edge, measured in points.
#[repr(transparent)]
#[derive(Debug, Clone, Copy, PartialEq, PartialOrd)]
pub struct Inset(f32);

impl Inset {
    /// A zero-point inset.
    pub const ZERO: Self = Self(0.0);

    /// Construct an inset after validating its finite, non-negative domain.
    ///
    /// # Errors
    ///
    /// Returns [`Error::NonFinite`] for NaN or infinite input and
    /// [`Error::Negative`] for a finite value below zero.
    pub fn from_points(points: f32) -> Result<Self> {
        if !points.is_finite() {
            return Err(Error::NonFinite);
        }
        if points < 0.0 {
            return Err(Error::Negative);
        }
        Ok(Self(points))
    }

    /// Return the distance in points.
    #[must_use]
    pub const fn points(self) -> f32 {
        self.0
    }
}

/// Independently typed text insets for all four cell edges.
#[repr(C)]
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct Insets {
    left: Inset,
    top: Inset,
    right: Inset,
    bottom: Inset,
}

impl Insets {
    /// Zero inset on every edge.
    pub const ZERO: Self = Self::uniform(Inset::ZERO);

    /// Construct insets in left, top, right, bottom order.
    #[must_use]
    pub const fn new(left: Inset, top: Inset, right: Inset, bottom: Inset) -> Self {
        Self {
            left,
            top,
            right,
            bottom,
        }
    }

    /// Construct equal insets on all four edges.
    #[must_use]
    pub const fn uniform(inset: Inset) -> Self {
        Self::new(inset, inset, inset, inset)
    }

    /// Return the left inset.
    #[must_use]
    pub const fn left(self) -> Inset {
        self.left
    }

    /// Return the top inset.
    #[must_use]
    pub const fn top(self) -> Inset {
        self.top
    }

    /// Return the right inset.
    #[must_use]
    pub const fn right(self) -> Inset {
        self.right
    }

    /// Return the bottom inset.
    #[must_use]
    pub const fn bottom(self) -> Inset {
        self.bottom
    }
}

impl Default for Insets {
    fn default() -> Self {
        Self::ZERO
    }
}

/// Effective composable text layout for one native table cell.
#[repr(C)]
#[derive(Debug, Clone, Copy, PartialEq, Default)]
pub struct Layout {
    text_wrap: TextWrap,
    vertical_alignment: VerticalAlignment,
    insets: Insets,
}

impl Layout {
    /// Construct a complete cell layout.
    #[must_use]
    pub const fn new(
        text_wrap: TextWrap,
        vertical_alignment: VerticalAlignment,
        insets: Insets,
    ) -> Self {
        Self {
            text_wrap,
            vertical_alignment,
            insets,
        }
    }

    /// Return the text-wrapping mode.
    #[must_use]
    pub const fn text_wrap(self) -> TextWrap {
        self.text_wrap
    }

    /// Return the vertical-alignment mode.
    #[must_use]
    pub const fn vertical_alignment(self) -> VerticalAlignment {
        self.vertical_alignment
    }

    /// Return the four edge insets.
    #[must_use]
    pub const fn insets(self) -> Insets {
        self.insets
    }

    /// Return a layout with a different text-wrapping mode.
    #[must_use]
    pub const fn with_text_wrap(mut self, text_wrap: TextWrap) -> Self {
        self.text_wrap = text_wrap;
        self
    }

    /// Return a layout with a different vertical-alignment mode.
    #[must_use]
    pub const fn with_vertical_alignment(mut self, vertical_alignment: VerticalAlignment) -> Self {
        self.vertical_alignment = vertical_alignment;
        self
    }

    /// Return a layout with different edge insets.
    #[must_use]
    pub const fn with_insets(mut self, insets: Insets) -> Self {
        self.insets = insets;
        self
    }
}

#[cfg(test)]
mod tests {
    use std::mem::size_of;

    use super::{Error, Inset, Insets, Layout, TextWrap, VerticalAlignment};

    #[test]
    fn layout_values_are_compact_and_composable() {
        assert_eq!(size_of::<Inset>(), 4);
        assert_eq!(size_of::<Insets>(), 16);
        assert_eq!(size_of::<Layout>(), 20);

        let inset = Inset::from_points(6.0).unwrap_or_else(|_| panic!("valid inset"));
        let layout = Layout::default()
            .with_text_wrap(TextWrap::Wrapped)
            .with_vertical_alignment(VerticalAlignment::Middle)
            .with_insets(Insets::uniform(inset));
        assert_eq!(layout.insets().left().points().to_bits(), 6.0_f32.to_bits());
        assert_eq!(layout.vertical_alignment(), VerticalAlignment::Middle);
    }

    #[test]
    fn inset_values_reject_non_finite_and_negative_input() {
        assert_eq!(Inset::from_points(-0.01), Err(Error::Negative));
        assert_eq!(Inset::from_points(f32::NAN), Err(Error::NonFinite));
        assert_eq!(Inset::from_points(f32::INFINITY), Err(Error::NonFinite));
        assert_eq!(Inset::from_points(f32::NEG_INFINITY), Err(Error::NonFinite));
    }
}
