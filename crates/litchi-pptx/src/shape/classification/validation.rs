//! Bounds and value validation for classification edits.

use super::model::{Outcome, Snapshot};
use crate::{Error, Result};

pub(super) const MAX_SHAPE_BYTES: usize = 64 * 1024 * 1024;
pub(super) const MAX_SHAPE_NODES: usize = 1_000_000;
pub(super) const MAX_SHAPE_DEPTH: usize = 512;
pub(super) const MAX_EXTENSION_BYTES: usize = 1 << 20;
pub(super) const MAX_EXTENSION_COUNT: usize = 1_024;

pub(super) fn validate_snapshot(snapshot: &Snapshot) -> Result<()> {
    snapshot.validate()
}

pub(super) fn validate_outcome(outcome: Outcome) -> Result<()> {
    // Keeping this as a separate seam makes future schema additions explicit
    // instead of allowing arbitrary lexical values through the writer.
    let wire = outcome.wire();
    if wire.is_empty() {
        return Err(Error::Invalid("classification outcome is empty".into()));
    }
    Ok(())
}

pub(super) fn count_node(nodes: &mut usize) -> Result<()> {
    *nodes = nodes.checked_add(1).ok_or(Error::Limit {
        resource: "classification shape XML nodes",
        limit: MAX_SHAPE_NODES,
    })?;
    if *nodes > MAX_SHAPE_NODES {
        return Err(Error::Limit {
            resource: "classification shape XML nodes",
            limit: MAX_SHAPE_NODES,
        });
    }
    Ok(())
}

pub(super) fn push_unknown(values: &mut Vec<super::model::Opaque>, xml: Vec<u8>) -> Result<()> {
    if values.len() >= MAX_EXTENSION_COUNT {
        return Err(Error::Limit {
            resource: "classification opaque extensions",
            limit: MAX_EXTENSION_COUNT,
        });
    }
    values.push(super::model::Opaque::from_wire(xml)?);
    Ok(())
}

pub(super) fn xml_error(error: impl std::fmt::Display) -> Error {
    Error::Xml(error.to_string())
}

pub(super) fn invalid(message: impl Into<String>) -> Error {
    Error::Invalid(message.into())
}
