//! Bounded `DrawingML` diagram data-model facade.
//!
//! The semantic model, XML codec, validation, and tests are kept in separate
//! layers while this module preserves the established public facade.
mod codec;
mod model;
#[cfg(test)]
mod tests;
mod validation;

pub use model::{
    Conformance, Connection, ConnectionType, DiagramDataModel, Id, IdError, Point, PointType,
};

pub(crate) use crate::diagram::{DGM_NAMESPACE, DGM_NAMESPACE_STRICT};

pub(super) const MAX_DATA_MODEL_XML: usize = 16 * 1024 * 1024;
pub(super) const MAX_NODES: usize = 200_000;
pub(super) const MAX_DEPTH: usize = 128;
pub(super) const MAX_POINTS: usize = 100_000;
pub(super) const MAX_CONNECTIONS: usize = 100_000;
pub(super) const MAX_TEXT_BYTES: usize = 1024 * 1024;
pub(super) const DML_NAMESPACE: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
pub(super) const DML_NAMESPACE_STRICT: &str = "http://purl.oclc.org/ooxml/drawingml/main";
/// Recursion guard for [`DiagramDataModel::node_tree`] on cyclic graphs.
pub(super) const MAX_TREE_DEPTH: u32 = 64;
