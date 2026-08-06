//! Typed DrawingML color choices.
//!
//! The `color` owner is intentionally limited to the two color choices shared
//! by the WordprocessingML, PresentationML, SpreadsheetML, and SpreadsheetML
//! binary drawing projections: `srgbClr` and `schemeClr`. Other valid
//! DrawingML color choices, extensions, and color transforms are retained as
//! bounded [`Unknown`] values instead of being silently discarded.
//!
//! The semantic values live in [`model`], the fragment codec and structural
//! limits live in [`codec`], and the focused conformance checks live in
//! [`tests`]. Format crates own the surrounding fill, line, shape, and package
//! relationships.

mod codec;
mod model;

#[cfg(test)]
mod tests;

pub use codec::{MAX_DEPTH, MAX_NODES, MAX_XML_BYTES, read, write};
pub use model::{Rgb, Scheme, Unknown, Value};
