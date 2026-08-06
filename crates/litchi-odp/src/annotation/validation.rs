//! Validation shared by the ODP annotation editor.

use super::{Anchor, Annotation, Position};
use litchi_core::{Error, Result};

pub(crate) const MAX_XML_BYTES: usize = 64 * 1024 * 1024;
pub(crate) const MAX_ANNOTATIONS: usize = 65_536;
pub(crate) const MAX_EVENTS: usize = 1_000_000;
pub(crate) const MAX_FRAGMENT_BYTES: usize = 16 * 1024 * 1024;

pub(crate) fn validate_annotation(annotation: &Annotation) -> Result<()> {
    annotation.validate()?;
    if let Some(name) = annotation.name() {
        if name.is_empty() {
            return invalid("annotation office:name cannot be empty");
        }
        if name.len() > MAX_FRAGMENT_BYTES {
            return invalid("annotation office:name exceeds the size limit");
        }
        if name.chars().any(char::is_control) {
            return invalid("annotation office:name contains an XML control character");
        }
    }
    Ok(())
}

pub(crate) fn validate_anchor(anchor: &Anchor) -> Result<()> {
    match anchor.position() {
        Position::Page { .. } => Ok(()),
        Position::Shape { name, .. } => super::model::validate_shape_name(name),
    }
}

pub(crate) fn bounds(index: usize, len: usize) -> Error {
    Error::InvalidFormat(format!(
        "annotation index {index} is out of bounds for {len} entries"
    ))
}

pub(crate) fn invalid<T>(message: impl Into<String>) -> Result<T> {
    Err(Error::InvalidFormat(message.into()))
}
