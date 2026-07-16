//! Shape and Drawing Element Support
//!
//! This module provides support for extracting text and metadata from shapes,
//! text boxes, and other drawing elements in iWork documents.
//!
//! Shapes in iWork documents can contain text (text boxes), images, or be
//! purely visual elements. This module helps extract meaningful content
//! from these elements.

mod color;
mod effects;
mod fill;
mod geometry;
mod line;
mod line_end;
mod path;
mod properties;
mod shadow;
mod stroke;
pub mod text_extractor;

pub use color::{RgbColorSpace, RgbaColor};
pub use effects::{ShapeEffects, ShapeOpacity, ShapeReflection, ShapeReflectionOpacity};
pub(crate) use effects::{reset_shape_effects, set_shape_effects, shape_effects};
pub use fill::{
    ShapeFill, ShapeGradient, ShapeGradientAngle, ShapeGradientKind, ShapeGradientOpacity,
    ShapeGradientStop, ShapeGradientStopMidpoint, ShapeGradientStopPosition,
    ShapeImageDataIdentifier, ShapeImageFill, ShapeImageFillTechnique,
};
pub(crate) use fill::{reset_shape_fill, set_shape_fill, set_shape_image_fill_data, shape_fill};
pub use geometry::{DrawableGeometry, DrawablePoint, DrawableSize};
pub(crate) use geometry::{
    geometry_from_drawable, patch_drawable_geometry, set_shape_geometry, shape_geometry,
};
pub use line::LineSegment;
pub(crate) use line::{
    line_geometry, line_path_source, line_segments_match, set_shape_line_segment,
    shape_line_segment,
};
pub use line_end::{LineEndpoint, LineEndpoints};
pub(crate) use line_end::{set_shape_line_endpoints, shape_line_endpoints};
pub use path::{
    ShapeCornerRadius, ShapePathKind, ShapePolygonSides, ShapePreset, ShapeStarInnerRatio,
    ShapeStarPoints,
};
pub(crate) use path::{set_shape_preset, shape_path_kind, shape_path_source, shape_preset};
pub use properties::DrawableProperties;
pub(crate) use properties::{set_shape_properties, shape_properties};
pub use shadow::{
    ShapeContactShadow, ShapeCurvedShadow, ShapeDropShadow, ShapeShadow, ShapeShadowAngle,
    ShapeShadowAppearance, ShapeShadowBlurRadius, ShapeShadowCurve, ShapeShadowOffset,
    ShapeShadowOpacity, ShapeShadowPerspective,
};
pub(crate) use shadow::{reset_shape_shadow, set_shape_shadow, shape_shadow};
pub use stroke::{
    LineStyle, ShapeStroke, StrokeCap, StrokeJoin, StrokeMiterLimit, StrokePattern, StrokeWidth,
};
pub(crate) use stroke::{reset_shape_stroke, set_shape_stroke, shape_stroke};
pub use text_extractor::ShapeTextExtractor;
