//! Typed native fills for ordinary iWork drawing shapes.

mod gradient;
mod image;
mod native;
mod style;

use super::RgbaColor;

pub use gradient::{
    ShapeGradient, ShapeGradientAngle, ShapeGradientKind, ShapeGradientOpacity, ShapeGradientStop,
    ShapeGradientStopMidpoint, ShapeGradientStopPosition,
};
pub use image::{ShapeImageDataIdentifier, ShapeImageFill, ShapeImageFillTechnique};
pub(crate) use native::{fill_from_native, fill_to_native, image_data_identifier};
pub(crate) use style::{remove_orphaned_image_asset, validate_image_asset};
pub(crate) use style::{reset_shape_fill, set_shape_fill, set_shape_image_fill_data, shape_fill};

/// Standard shape fills shared by Pages, Numbers, and Keynote.
#[derive(Debug, Clone, PartialEq, Default)]
pub enum ShapeFill {
    /// An empty native `FillArchive` (the applications display “No Fill”).
    #[default]
    None,
    /// One normalized RGB color.
    Solid(RgbaColor),
    /// A validated simple or advanced linear/radial gradient.
    Gradient(ShapeGradient),
    /// A validated simple or tinted image fill.
    Image(ShapeImageFill),
}
