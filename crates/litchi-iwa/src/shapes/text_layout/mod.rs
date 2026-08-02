//! Typed frame-level layout for text inside shapes and ordinary text boxes.
//!
//! iWork stores vertical alignment, edge insets, and autosizing on the shape
//! style. Text-box columns are another independently composable shape-style
//! property; document-body columns use a separate column-style archive.

mod native;
mod style;

use crate::{Error, Result};

pub(crate) use style::{reset_shape_text_layout, set_shape_text_layout, shape_text_layout};

/// Vertical placement of text inside a shape or ordinary text box.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
pub enum ShapeTextVerticalAlignment {
    #[default]
    Top,
    Middle,
    Bottom,
    Justified,
}

/// Whether iWork should shrink text until it fits inside the shape.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
pub enum ShapeTextAutoSize {
    #[default]
    Fixed,
    ShrinkToFit,
}

/// Finite, non-negative distance from one shape edge, measured in points.
#[derive(Debug, Clone, Copy, PartialEq, PartialOrd)]
pub struct ShapeTextInset(f32);

impl ShapeTextInset {
    pub const ZERO: Self = Self(0.0);

    pub fn from_points(points: f32) -> Result<Self> {
        if !points.is_finite() || points < 0.0 {
            return Err(Error::ParseError(
                "iWork shape text inset must be finite and non-negative".to_owned(),
            ));
        }
        Ok(Self(points))
    }

    pub const fn points(self) -> f32 {
        self.0
    }
}

/// Independently typed text insets for all four shape edges.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct ShapeTextInsets {
    left: ShapeTextInset,
    top: ShapeTextInset,
    right: ShapeTextInset,
    bottom: ShapeTextInset,
}

impl ShapeTextInsets {
    pub const ZERO: Self = Self::uniform(ShapeTextInset::ZERO);

    pub const fn new(
        left: ShapeTextInset,
        top: ShapeTextInset,
        right: ShapeTextInset,
        bottom: ShapeTextInset,
    ) -> Self {
        Self {
            left,
            top,
            right,
            bottom,
        }
    }

    pub const fn uniform(inset: ShapeTextInset) -> Self {
        Self::new(inset, inset, inset, inset)
    }

    pub const fn left(self) -> ShapeTextInset {
        self.left
    }

    pub const fn top(self) -> ShapeTextInset {
        self.top
    }

    pub const fn right(self) -> ShapeTextInset {
        self.right
    }

    pub const fn bottom(self) -> ShapeTextInset {
        self.bottom
    }
}

impl Default for ShapeTextInsets {
    fn default() -> Self {
        Self::ZERO
    }
}

/// Composable frame-level text layout stored in a native shape style.
#[derive(Debug, Clone, Copy, PartialEq, Default)]
pub struct ShapeTextLayout {
    vertical_alignment: ShapeTextVerticalAlignment,
    insets: ShapeTextInsets,
    auto_size: ShapeTextAutoSize,
}

impl ShapeTextLayout {
    pub const fn new(
        vertical_alignment: ShapeTextVerticalAlignment,
        insets: ShapeTextInsets,
        auto_size: ShapeTextAutoSize,
    ) -> Self {
        Self {
            vertical_alignment,
            insets,
            auto_size,
        }
    }

    pub const fn vertical_alignment(self) -> ShapeTextVerticalAlignment {
        self.vertical_alignment
    }

    pub const fn insets(self) -> ShapeTextInsets {
        self.insets
    }

    pub const fn auto_size(self) -> ShapeTextAutoSize {
        self.auto_size
    }

    pub const fn with_vertical_alignment(
        mut self,
        vertical_alignment: ShapeTextVerticalAlignment,
    ) -> Self {
        self.vertical_alignment = vertical_alignment;
        self
    }

    pub const fn with_insets(mut self, insets: ShapeTextInsets) -> Self {
        self.insets = insets;
        self
    }

    pub const fn with_auto_size(mut self, auto_size: ShapeTextAutoSize) -> Self {
        self.auto_size = auto_size;
        self
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn text_insets_reject_invalid_distances() {
        assert!(ShapeTextInset::from_points(-0.01).is_err());
        assert!(ShapeTextInset::from_points(f32::NAN).is_err());
        assert!(ShapeTextInset::from_points(f32::INFINITY).is_err());
    }

    #[test]
    fn layout_builders_preserve_strong_types() {
        let inset = ShapeTextInset::from_points(12.0).unwrap();
        let layout = ShapeTextLayout::default()
            .with_vertical_alignment(ShapeTextVerticalAlignment::Middle)
            .with_insets(ShapeTextInsets::uniform(inset))
            .with_auto_size(ShapeTextAutoSize::ShrinkToFit);
        assert_eq!(layout.insets(), ShapeTextInsets::uniform(inset));
        assert_eq!(layout.auto_size(), ShapeTextAutoSize::ShrinkToFit);
    }
}
