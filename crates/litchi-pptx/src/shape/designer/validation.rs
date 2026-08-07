//! Bounds and value validation for Designer design-element edits.

use super::model::{Opaque, Snapshot};
use crate::{Error, Result};

pub(super) const MAX_SHAPE_BYTES: usize = 64 * 1024 * 1024;
pub(super) const MAX_SHAPE_NODES: usize = 1_000_000;
pub(super) const MAX_SHAPE_DEPTH: usize = 512;
pub(super) const MAX_EXTENSION_BYTES: usize = 1 << 20;
pub(super) const MAX_EXTENSION_COUNT: usize = 1_024;

pub(super) fn validate_snapshot(snapshot: &Snapshot) -> Result<()> {
    snapshot.validate()
}

pub(super) fn validate_value(value: bool) -> Result<()> {
    let _ = value;
    Ok(())
}

pub(super) fn validate_tag_text(
    value: &str,
    limits: super::model::Limits,
    resource: &'static str,
) -> Result<()> {
    if value.len() > limits.string_bytes() {
        return Err(Error::Limit {
            resource,
            limit: limits.string_bytes(),
        });
    }
    if value.chars().all(is_xml_10_char) {
        Ok(())
    } else {
        Err(invalid(format!(
            "{resource} contains a character forbidden by XML 1.0"
        )))
    }
}

fn is_xml_10_char(value: char) -> bool {
    matches!(
        value,
        '\u{9}' | '\u{A}' | '\u{D}' | '\u{20}'..='\u{D7FF}' | '\u{E000}'..='\u{FFFD}' | '\u{10000}'..='\u{10FFFF}'
    )
}

pub(super) fn count_node(nodes: &mut usize) -> Result<()> {
    *nodes = nodes.checked_add(1).ok_or(Error::Limit {
        resource: "designer shape XML nodes",
        limit: MAX_SHAPE_NODES,
    })?;
    if *nodes > MAX_SHAPE_NODES {
        return Err(Error::Limit {
            resource: "designer shape XML nodes",
            limit: MAX_SHAPE_NODES,
        });
    }
    Ok(())
}

pub(super) fn push_unknown(values: &mut Vec<Opaque>, xml: Vec<u8>) -> Result<()> {
    if values.len() >= MAX_EXTENSION_COUNT {
        return Err(Error::Limit {
            resource: "designer opaque extensions",
            limit: MAX_EXTENSION_COUNT,
        });
    }
    values.push(Opaque::from_wire(xml)?);
    Ok(())
}

pub(super) fn xml_error(error: impl std::fmt::Display) -> Error {
    Error::Xml(error.to_string())
}

pub(super) fn invalid(message: impl Into<String>) -> Error {
    Error::Invalid(message.into())
}
