//! Structural and semantic validation for `PowerPoint` anchors.

use crate::package::{Error, Result};

use super::model::{Limits, RECT_LEN, SMALL_RECT_LEN};

pub(super) fn bounds(left: i32, top: i32, right: i32, bottom: i32) -> Result<()> {
    if left > right {
        return invalid("shape anchor left exceeds right");
    }
    if top > bottom {
        return invalid("shape anchor top exceeds bottom");
    }
    Ok(())
}

pub(super) fn payload_len(length: usize, limits: Limits) -> Result<()> {
    if length != SMALL_RECT_LEN && length != RECT_LEN {
        return corrupted(format!(
            "OfficeArtClientAnchor recLen must be 8 or 16, got {length}"
        ));
    }
    if length > limits.max_payload_bytes {
        return corrupted("OfficeArtClientAnchor payload exceeds the configured limit");
    }
    Ok(())
}

pub(super) fn header(version: u16, instance: u16, record_type: u16) -> Result<()> {
    if version != 0 {
        return corrupted(format!(
            "OfficeArtClientAnchor recVer must be 0, got {version}"
        ));
    }
    if instance != 0 {
        return corrupted(format!(
            "OfficeArtClientAnchor recInstance must be 0, got {instance}"
        ));
    }
    if record_type != super::codec::RECORD_TYPE {
        return corrupted(format!(
            "OfficeArtClientAnchor recType must be 0xF010, got 0x{record_type:04X}"
        ));
    }
    Ok(())
}

pub(super) fn corrupted(message: impl Into<String>) -> Result<()> {
    Err(Error::Corrupted(message.into()))
}

pub(super) fn invalid(message: impl Into<String>) -> Result<()> {
    Err(Error::InvalidFormat(message.into()))
}
