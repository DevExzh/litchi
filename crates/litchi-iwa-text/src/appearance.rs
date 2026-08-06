//! Archive-free text appearance values shared by iWork format owners.
//!
//! The native IWA adapter remains responsible for protobuf conversion,
//! inheritance, and package mutation. This module owns only the compact
//! semantic compositions exposed by the text inspector.

use litchi_iwa_common::{
    color::Rgba,
    shape::{
        shadow::{
            Angle, Appearance as ShadowAppearance, BlurRadius, Drop as DropShadow, Offset, Opacity,
            Shadow as ShapeShadow,
        },
        stroke::{Pattern, Stroke, Width},
    },
};

/// Validation failures produced by text appearance conversions.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Error {
    /// A shape shadow family other than a drop shadow was supplied for text.
    UnsupportedShadowFamily,
}

impl std::fmt::Display for Error {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::UnsupportedShadowFamily => {
                formatter.write_str("text appearance supports only drop shadows")
            },
        }
    }
}

impl std::error::Error for Error {}

/// Result returned by checked text-appearance conversions.
pub type Result<T> = std::result::Result<T, Error>;

/// Effective outline applied to uniformly styled text.
#[derive(Debug, Clone, Copy, PartialEq, Default)]
pub enum Outline {
    /// Render text without an outline.
    #[default]
    None,
    /// Render text using a typed native stroke.
    Stroke(Stroke),
}

impl Outline {
    /// Construct the exact one-point outline written by the native inspector.
    #[must_use]
    pub fn standard() -> Self {
        Self::Stroke(Stroke::new(
            Rgba::transparent_black(),
            Width::ONE,
            Pattern::Solid,
        ))
    }
}

/// Effective drop shadow applied to uniformly styled text.
#[derive(Debug, Clone, Copy, PartialEq, Default)]
pub enum Shadow {
    /// Render text without a shadow.
    #[default]
    None,
    /// Render text with a typed native drop shadow.
    Drop(DropShadow),
}

impl Shadow {
    /// Construct the exact shadow written by the native inspector checkbox.
    #[must_use]
    pub const fn standard() -> Self {
        Self::Drop(DropShadow::new(
            ShadowAppearance::new(
                Rgba::black(),
                BlurRadius::ONE_POINT,
                Offset::FIVE_POINTS,
                Opacity::OPAQUE,
            ),
            Angle::FORTY_FIVE_DEGREES,
        ))
    }

    /// Convert this text shadow into the neutral shape-shadow vocabulary.
    #[must_use]
    pub const fn into_shape_shadow(self) -> ShapeShadow {
        match self {
            Self::None => ShapeShadow::Disabled,
            Self::Drop(shadow) => ShapeShadow::Drop(shadow),
        }
    }

    /// Restrict a neutral shape shadow to the drop-shadow family supported by
    /// the text inspector.
    ///
    /// # Errors
    ///
    /// Returns [`Error::UnsupportedShadowFamily`] for contact and curved
    /// shape shadows.
    pub fn from_shape_shadow(shape_shadow: ShapeShadow) -> Result<Self> {
        match shape_shadow {
            ShapeShadow::Disabled => Ok(Self::None),
            ShapeShadow::Drop(drop_shadow) => Ok(Self::Drop(drop_shadow)),
            ShapeShadow::Contact(_) | ShapeShadow::Curved(_) => Err(Error::UnsupportedShadowFamily),
        }
    }
}

/// Effective solid background painted behind uniformly styled text glyphs.
#[derive(Debug, Clone, Copy, PartialEq, Default)]
pub enum Background {
    /// Do not paint a background behind the text glyph run.
    #[default]
    None,
    /// Paint a solid color behind the text glyph run.
    Color(Rgba),
}

/// Effective solid fill painted across a paragraph's layout box.
#[derive(Debug, Clone, Copy, PartialEq, Default)]
pub enum ParagraphBackground {
    /// Leave the paragraph layout box unfilled.
    #[default]
    None,
    /// Paint the paragraph layout box with a solid color.
    Color(Rgba),
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn standard_appearances_are_typed_and_compact() {
        assert!(matches!(Outline::standard(), Outline::Stroke(_)));
        assert!(matches!(Shadow::standard(), Shadow::Drop(_)));
        assert_eq!(Background::default(), Background::None);
        assert_eq!(ParagraphBackground::default(), ParagraphBackground::None);
    }

    #[test]
    fn only_drop_shadows_are_valid_for_text() {
        let contact = ShapeShadow::Contact(litchi_iwa_common::shape::shadow::Contact::new(
            ShadowAppearance::new(
                Rgba::black(),
                BlurRadius::ZERO,
                Offset::ZERO,
                Opacity::OPAQUE,
            ),
            litchi_iwa_common::shape::shadow::Perspective::LEVEL,
        ));
        assert_eq!(
            Shadow::from_shape_shadow(contact),
            Err(Error::UnsupportedShadowFamily)
        );
    }
}
