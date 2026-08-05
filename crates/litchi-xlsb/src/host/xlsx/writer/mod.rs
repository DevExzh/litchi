//! DrawingML authoring helpers used by XLSB worksheet drawings.

pub mod shape;

pub use super::shapes::Geometry;
pub use shape::{ConnectionEndSpec, ConnectionShapeSpec, DrawingObjectSpec, GroupSpec, ShapeSpec};
