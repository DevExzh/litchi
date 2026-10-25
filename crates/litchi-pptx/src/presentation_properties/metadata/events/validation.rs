//! Semantic and safety validation for inert slide-show event metadata.

use super::codec::Limits;
use super::model::{Draft, Kind};
use crate::{Error, Result};

pub(crate) const MAX_EVENTS: usize = 65_536;
pub(crate) const MAX_SLIDE_XML_BYTES: usize = 32 * 1024 * 1024;
pub(crate) const MAX_TOTAL_SLIDE_XML_BYTES: usize = 256 * 1024 * 1024;
pub(crate) const MAX_XML_NODES: usize = 250_000;
pub(crate) const MAX_XML_DEPTH: usize = 128;
pub(crate) const MAX_EVENT_STRING_BYTES: usize = 64;

/// Validate one typed event without resolving its target or executing it.
pub(crate) fn validate_draft(value: &Draft, limits: Limits) -> Result<()> {
    if value.time().as_str().len() > limits.max_event_string_bytes() {
        return Err(limit(
            crate::error::CodecLimitKind::StringBytes,
            value.time().as_str().len(),
            limits.max_event_string_bytes(),
        ));
    }
    if let Kind::Seek { at } = value.kind()
        && at.as_str().len() > limits.max_event_string_bytes()
    {
        return Err(limit(
            crate::error::CodecLimitKind::StringBytes,
            at.as_str().len(),
            limits.max_event_string_bytes(),
        ));
    }
    Ok(())
}

/// Validate the bounded, ordered event collection.
pub(crate) fn validate_events(values: &[Draft], limits: Limits) -> Result<()> {
    if values.is_empty() {
        return Err(invalid(
            "slide-show event storage requires at least one event",
        ));
    }
    if values.len() > limits.max_events() {
        return Err(limit(
            crate::error::CodecLimitKind::Events,
            values.len(),
            limits.max_events(),
        ));
    }
    for value in values {
        validate_draft(value, limits)?;
    }
    Ok(())
}

pub(crate) fn validate_extension_uri(value: &str, expected: &str) -> Result<()> {
    if value == expected {
        Ok(())
    } else {
        Err(invalid(
            "slide-show event extension URI is not the MS-PPTX URI",
        ))
    }
}

fn invalid(message: impl Into<String>) -> Error {
    Error::Invalid(message.into())
}

fn limit(kind: crate::error::CodecLimitKind, observed: usize, maximum: usize) -> Error {
    Error::Limit {
        kind: crate::error::LimitKind::Codec(kind),
        observed,
        maximum,
    }
}
