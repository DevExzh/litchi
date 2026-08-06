//! Bounded semantic and relationship validation for media track metadata.

use litchi_ooxml_common::xml::is_ncname;

use super::model::{Caption, CaptionTarget, MediaMetadata, TracksInfo};
use crate::{Error, Result};

pub(crate) const MAX_TRACKS: usize = 4_096;
pub(crate) const MAX_STRING_BYTES: usize = 1024 * 1024;

pub(crate) fn validate_metadata(value: &MediaMetadata) -> Result<()> {
    if value.key.slide_part_name.is_empty() || value.key.slide_part_name.len() > MAX_STRING_BYTES {
        return Err(invalid("media track slide part name is empty or too long"));
    }
    if value.key.shape_id == 0 {
        return Err(invalid("media track shape ID must be non-zero"));
    }
    if let Some(relationship_id) = &value.media_relationship_id {
        validate_relationship_id(relationship_id, "media relationship ID")?;
    }
    if let Some(tracks) = &value.tracks_info {
        validate_tracks(tracks)?;
    }
    Ok(())
}

pub(crate) fn validate_tracks(value: &TracksInfo) -> Result<()> {
    if value.captions.len() > MAX_TRACKS {
        return Err(Error::Limit {
            resource: "media caption track count",
            limit: MAX_TRACKS,
        });
    }
    let mut ids = std::collections::HashSet::with_capacity(value.captions.len());
    for caption in &value.captions {
        validate_caption(caption)?;
        if !ids.insert(caption.id.as_str()) {
            return Err(invalid(format!(
                "duplicate media caption track ID '{}'",
                caption.id
            )));
        }
    }
    Ok(())
}

pub(crate) fn validate_caption(value: &Caption) -> Result<()> {
    if !is_guid(&value.id) {
        return Err(invalid(format!(
            "invalid media caption track GUID '{}'",
            value.id
        )));
    }
    bounded(&value.id, "media caption track ID")?;
    if value.label.is_empty() {
        return Err(invalid("media caption track label cannot be empty"));
    }
    bounded(&value.label, "media caption track label")?;
    if let Some(language) = &value.language {
        bounded(language, "media caption language")?;
        if language.trim().is_empty() {
            return Err(invalid("media caption language cannot be blank"));
        }
    }
    match &value.target {
        CaptionTarget::Internal {
            part_name,
            content_type,
        } => {
            bounded(part_name, "media caption target part name")?;
            bounded(content_type, "media caption target content type")?;
            if part_name.is_empty() || content_type.is_empty() {
                return Err(invalid("internal media caption target is incomplete"));
            }
        },
        CaptionTarget::External { target } => {
            bounded(target, "external media caption target")?;
            if target.is_empty() {
                return Err(invalid("external media caption target cannot be empty"));
            }
        },
    }
    Ok(())
}

pub(crate) fn validate_relationship_id(value: &str, label: &str) -> Result<()> {
    bounded(value, label)?;
    if value.is_empty() || !is_ncname(value) {
        return Err(invalid(format!("{label} is not a valid XML NCName")));
    }
    Ok(())
}

pub(crate) fn parse_boolean(value: &str, label: &str) -> Result<bool> {
    match value {
        "true" | "1" => Ok(true),
        "false" | "0" => Ok(false),
        _ => Err(invalid(format!("{label} is not an XML boolean"))),
    }
}

fn bounded(value: &str, _label: &str) -> Result<()> {
    if value.len() > MAX_STRING_BYTES {
        return Err(Error::Limit {
            resource: "media track string",
            limit: MAX_STRING_BYTES,
        });
    }
    Ok(())
}

fn is_guid(value: &str) -> bool {
    let bytes = value.as_bytes();
    if bytes.len() != 38 || bytes[0] != b'{' || bytes[37] != b'}' {
        return false;
    }
    for (index, byte) in bytes[1..37].iter().copied().enumerate() {
        if matches!(index, 8 | 13 | 18 | 23) {
            if byte != b'-' {
                return false;
            }
        } else if !byte.is_ascii_hexdigit() {
            return false;
        }
    }
    true
}

fn invalid(message: impl Into<String>) -> Error {
    Error::Invalid(message.into())
}
