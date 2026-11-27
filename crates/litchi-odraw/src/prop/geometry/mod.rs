//! Typed custom geometry from `[MS-ODRAW]` shape properties.
//!
//! The public facade is intentionally small: [`Geometry`] borrows the
//! `pVertices_complex` and `pSegmentInfo_complex` arrays from the containing
//! property table, while [`Points`] and [`PathInfos`] expose checked typed
//! views without copying their wire data.  The implementation is split into
//! semantic models, wire decoding, and cross-property validation.

mod codec;
mod model;
mod validation;

pub use codec::parse;
pub use model::{
    Coordinate, EscapeKind, Geometry, Instruction, PathInfo, PathInfos, PathKind, Point, Points,
};

#[cfg(test)]
mod tests;
