//! Typed native fills for ordinary iWork drawing shapes.

mod gradient;
mod native;
mod style;

use super::RgbaColor;

pub use gradient::{
    ShapeGradient, ShapeGradientAngle, ShapeGradientKind, ShapeGradientOpacity, ShapeGradientStop,
    ShapeGradientStopMidpoint, ShapeGradientStopPosition,
};
pub(crate) use style::{reset_shape_fill, set_shape_fill, shape_fill};

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
}
