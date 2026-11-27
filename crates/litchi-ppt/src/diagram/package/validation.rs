//! Structural and resource checks for slide-boundary publication.

use crate::package::{Error, Result};

use super::model::SlideLimits;

pub(super) fn validate_limits(limits: SlideLimits) -> Result<()> {
    if limits.max_slide_bytes < 8 {
        return Err(Error::InvalidFormat(
            "SlideContainer limit must include a record header".into(),
        ));
    }
    if limits.diagram.max_build_list_bytes < 8 {
        return Err(Error::InvalidFormat(
            "diagram BuildList limit must include a record header".into(),
        ));
    }
    Ok(())
}

pub(super) fn invalid(message: impl Into<String>) -> Error {
    Error::InvalidFormat(message.into())
}

pub(super) fn corrupted(message: impl Into<String>) -> Error {
    Error::Corrupted(message.into())
}
