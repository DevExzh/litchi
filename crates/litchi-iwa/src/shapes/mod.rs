//! Shape and Drawing Element Support
//!
//! This module provides support for extracting text and metadata from shapes,
//! text boxes, and other drawing elements in iWork documents.
//!
//! Shapes in iWork documents can contain text (text boxes), images, or be
//! purely visual elements. This module helps extract meaningful content
//! from these elements.

mod caption;
mod color;
mod effects;
mod fill;
mod geometry;
pub mod line;
mod line_end;
mod path;
mod properties;
mod shadow;
mod stroke;
mod text_columns;
pub mod text_extractor;
mod text_layout;

pub use caption::DrawableTitleCaption;
pub use color::{RgbColorSpace, Rgba, RgbaColor};
pub(crate) use color::{color_from_native, color_to_native};
pub(crate) use effects::{reset_shape_effects, set_shape_effects, shape_effects};
pub use fill::{ShapeFill, ShapeImageDataIdentifier, ShapeImageFill, ShapeImageFillTechnique};
pub(crate) use fill::{
    fill_from_native, fill_to_native, image_data_identifier, remove_orphaned_image_asset,
    validate_image_asset,
};
pub(crate) use fill::{reset_shape_fill, set_shape_fill, set_shape_image_fill_data, shape_fill};
pub use geometry::{DrawableFlipAxis, DrawableGeometry, DrawablePoint, DrawableSize};
pub(crate) use geometry::{
    flip_drawable_geometry, geometry_from_drawable, offset_drawable_geometry,
    patch_drawable_geometry, restore_drawable_original_size, set_shape_geometry, shape_geometry,
};
pub use line::LineSegment;
pub(crate) use line::{
    line_geometry, line_path_source, line_segments_match, set_shape_line_segment,
    shape_line_segment,
};
pub use line_end::{Endpoint, Endpoints};
pub(crate) use line_end::{
    insert_style_variation, remove_style_variation, set_shape_line_endpoints, shape_line_endpoints,
};
pub use litchi_iwa_common::shape::effects::{Effects, Opacity, Reflection, ReflectionOpacity};
pub use litchi_iwa_common::shape::fill::Opacity as GradientOpacity;
pub use litchi_iwa_common::shape::fill::{Angle, Gradient, Kind, Stop, StopMidpoint, StopPosition};
pub use litchi_iwa_common::shape::path::{
    CornerRadius, InnerRadiusRatio, PolygonSides, Preset, StarPoints,
};
pub use path::ShapePathKind;
pub(crate) use path::{set_shape_preset, shape_path_kind, shape_path_source, shape_preset};

impl From<litchi_iwa_common::shape::path::Error> for crate::Error {
    fn from(error: litchi_iwa_common::shape::path::Error) -> Self {
        Self::ParseError(error.to_string())
    }
}

pub use properties::DrawableProperties;
pub(crate) use properties::{
    drawable_properties, patch_wrapped_drawable_properties, set_shape_properties, shape_properties,
};
pub use shadow::{
    Appearance, BlurRadius, Contact, Curve, Curved, Drop, Offset, Perspective, Shadow,
};
// The protected paragraph-alignment fixture is migrated by a concurrent agent.
// Keep its test-only imports resolving without exposing a production alias.
#[cfg(test)]
#[doc(hidden)]
pub use shadow::{
    Angle as ShapeShadowAngle, Appearance as ShapeShadowAppearance,
    BlurRadius as ShapeShadowBlurRadius, Contact as ShapeContactShadow, Drop as ShapeDropShadow,
    Offset as ShapeShadowOffset, Opacity as ShapeShadowOpacity,
    Perspective as ShapeShadowPerspective, Shadow as ShapeShadow,
};
pub(crate) use shadow::{
    reset_shape_shadow, set_shape_shadow, shadow_from_native, shadow_to_native, shape_shadow,
};
pub use stroke::{Cap, Join, LineStyle, MiterLimit, Pattern, Stroke, Width};
#[cfg(test)]
#[doc(hidden)]
pub use stroke::{Pattern as StrokePattern, Width as StrokeWidth};
pub(crate) use stroke::{
    empty_stroke_archive, reset_shape_stroke, set_shape_stroke, shape_stroke, stroke_from_native,
    stroke_to_native,
};
pub(crate) use text_columns::{
    reset_shape_text_columns, set_shape_text_columns, shape_text_columns,
};
pub use text_extractor::ShapeTextExtractor;
pub(crate) use text_layout::{reset_shape_text_layout, set_shape_text_layout, shape_text_layout};

impl From<litchi_iwa_common::shape::fill::Error> for crate::Error {
    fn from(error: litchi_iwa_common::shape::fill::Error) -> Self {
        Self::ParseError(error.to_string())
    }
}

impl From<litchi_iwa_common::shape::stroke::Error> for crate::Error {
    fn from(error: litchi_iwa_common::shape::stroke::Error) -> Self {
        Self::ParseError(error.to_string())
    }
}

impl From<litchi_iwa_common::shape::shadow::Error> for crate::Error {
    fn from(error: litchi_iwa_common::shape::shadow::Error) -> Self {
        Self::ParseError(error.to_string())
    }
}
