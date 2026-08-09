//! Typed `DrawingML` color choices and transforms.
//!
//! The `color` owner models the ECMA-376 `EG_ColorChoice` and
//! `EG_ColorTransform` groups shared by the `WordprocessingML`, `PresentationML`,
//! `SpreadsheetML`, and `SpreadsheetML` binary drawing projections. Unsupported
//! choices and extensions remain bounded [`Unknown`] values instead of being
//! silently discarded.
//!
//! The semantic values live in [`model`], the fragment codec and structural
//! limits live in [`codec`], and the focused conformance checks live in
//! [`tests`]. Format crates own the surrounding fill, line, shape, and package
//! relationships.

mod codec;
mod model;
mod validation;

#[cfg(test)]
mod tests;

pub use codec::{MAX_DEPTH, MAX_NODES, MAX_TRANSFORMS, MAX_XML_BYTES, read, write};
pub use model::{
    Angle, Base, FixedPercentage, Hsl, Percentage, PositiveAngle, PositiveFixedPercentage,
    PositivePercentage, Preset, Rgb, ScRgb, Scheme, System, Transform, Transformed, Unknown, Value,
};
