//! Named drawing resources stored in ODF common styles.
//!
//! Each resource family owns its typed model, bounded XML codec, and focused
//! regression coverage. Package and flat-document accessors live in `style`.

pub mod fill_image;
pub mod gradient;
pub mod hatch;
pub mod marker;
pub mod opacity;
pub mod stroke_dash;
pub mod style;

pub use fill_image::{
    FillImage, FillImageActuate, FillImageLength, FillImageLengthUnit, FillImageLink,
    FillImageLinkKind, FillImageShow, FillImageSource, FillImages, parse_drawing_fill_images,
};
pub use gradient::{
    Gradient, GradientAngle, GradientCoordinate, GradientCoordinateUnit, GradientIntensity,
    GradientPercent, GradientSpreadMethod, GradientStopOffset, Gradients, LegacyGradient,
    LegacyGradientStyle, LibreOfficeGradientColorType, LibreOfficeGradientStop, RgbColor,
    SvgGradientCommon, SvgGradientStop, SvgLinearGradient, SvgRadialGradient,
    parse_drawing_gradients,
};
pub use hatch::{
    Hatch, HatchLength, HatchLengthUnit, HatchRotation, HatchStyle, Hatches, parse_drawing_hatches,
};
pub use marker::{Marker, MarkerPathData, MarkerViewBox, Markers, parse_drawing_markers};
pub use opacity::{
    Opacities, Opacity, OpacityAngle, OpacityGeometryPercent, OpacityPercent, OpacityStop,
    OpacityStopValue, OpacityStyle, parse_drawing_opacities,
};
pub use stroke_dash::{
    StrokeDash, StrokeDashMeasure, StrokeDashMeasureUnit, StrokeDashStyle, StrokeDashes,
    parse_drawing_stroke_dashes,
};
