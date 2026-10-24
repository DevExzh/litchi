//! Wire-level model3d XML codec.

mod xml;

use crate::{Error, LimitKind, Result};

use super::{
    MAX_CHILDREN, MAX_DEPTH, MAX_FRAGMENT_BYTES, MAX_NODES, MAX_RASTER_CHILDREN, MAX_XML_BYTES,
};

/// Caller-selected bounds for one model3d XML operation.
///
/// The fixed schema ceilings remain hard caps; this policy lets a host select
/// a narrower input, retained-fragment, output, or tree budget without exposing
/// a limit-less codec twin.  `write_to` uses the output budget after a counting
/// preflight and then streams directly to the caller's sink.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
#[must_use]
pub struct Limits {
    /// Maximum complete model3d XML accepted by [`read`].
    pub max_xml_bytes: usize,
    /// Maximum one opaque child fragment accepted by [`opaque`].
    pub max_fragment_bytes: usize,
    /// Maximum serialized model3d XML produced by a write operation.
    pub max_output_bytes: usize,
    /// Maximum model3d children retained by a read operation.
    pub max_children: usize,
    /// Maximum raster children retained by a read operation.
    pub max_raster_children: usize,
    /// Maximum nested XML depth inspected by a read operation.
    pub max_depth: usize,
    /// Maximum nested XML nodes inspected by a read operation.
    pub max_nodes: usize,
}

impl Limits {
    /// The historical model3d bounds.
    pub const DEFAULT: Self = Self {
        max_xml_bytes: MAX_XML_BYTES,
        max_fragment_bytes: MAX_FRAGMENT_BYTES,
        max_output_bytes: MAX_XML_BYTES,
        max_children: MAX_CHILDREN,
        max_raster_children: MAX_RASTER_CHILDREN,
        max_depth: MAX_DEPTH,
        max_nodes: MAX_NODES,
    };

    /// Construct a policy with the historical tree ceilings and selected byte
    /// budgets.
    #[must_use]
    pub const fn new(
        max_xml_bytes: usize,
        max_fragment_bytes: usize,
        max_output_bytes: usize,
    ) -> Self {
        Self {
            max_xml_bytes,
            max_fragment_bytes,
            max_output_bytes,
            ..Self::DEFAULT
        }
    }

    pub(crate) fn validate(self) -> Result<Self> {
        validate_bound(LimitKind::XmlBytes, self.max_xml_bytes, MAX_XML_BYTES)?;
        validate_bound(
            LimitKind::InertFragmentBytes,
            self.max_fragment_bytes,
            MAX_FRAGMENT_BYTES,
        )?;
        validate_bound(LimitKind::OutputBytes, self.max_output_bytes, MAX_XML_BYTES)?;
        validate_bound(LimitKind::ModelChildren, self.max_children, MAX_CHILDREN)?;
        validate_bound(
            LimitKind::RasterChildren,
            self.max_raster_children,
            MAX_RASTER_CHILDREN,
        )?;
        validate_bound(LimitKind::XmlDepth, self.max_depth, MAX_DEPTH)?;
        validate_bound(LimitKind::XmlNodes, self.max_nodes, MAX_NODES)?;
        Ok(self)
    }
}

impl Default for Limits {
    fn default() -> Self {
        Self::DEFAULT
    }
}

fn validate_bound(kind: LimitKind, value: usize, maximum: usize) -> Result<()> {
    if value == 0 || value > maximum {
        return Err(Error::invalid_limit(kind, value, maximum));
    }
    Ok(())
}

pub use xml::{opaque, read, write, write_to};
