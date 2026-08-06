use super::super::model::{Frame, Kind, Mode};
use crate::presentation::embedded::{MAX_ATTRIBUTE_BYTES, invalid, limit};
use crate::{Error, Result};

pub(crate) const MAX_SLIDE_BYTES: usize = 32 * 1024 * 1024;
pub(crate) const MAX_PART_BYTES: usize = 64 * 1024 * 1024;
const MAX_PAYLOAD_BYTES: usize = 32 * 1024 * 1024;
const MAX_NAME_CHARS: usize = 256;
pub(crate) const MAX_OBJECTS: usize = 4_096;
pub(crate) const MAX_RELATIONSHIPS: usize = 16_384;

pub(crate) fn validate_source(bytes: &[u8]) -> Result<()> {
    if bytes.len() > MAX_SLIDE_BYTES {
        return Err(limit("OLE slide XML bytes", MAX_SLIDE_BYTES));
    }
    Ok(())
}

pub(crate) fn validate_payload(bytes: &[u8]) -> Result<()> {
    if bytes.is_empty() {
        return Err(invalid("embedded OLE payload cannot be empty"));
    }
    if bytes.len() > MAX_PART_BYTES || bytes.len() > MAX_PAYLOAD_BYTES {
        return Err(limit("embedded OLE payload bytes", MAX_PART_BYTES));
    }
    Ok(())
}

pub(crate) fn validate_text(value: &str, label: &'static str, allow_empty: bool) -> Result<()> {
    if value.len() > MAX_ATTRIBUTE_BYTES {
        return Err(limit(label, MAX_ATTRIBUTE_BYTES));
    }
    if !allow_empty && value.is_empty() {
        return Err(invalid(format!("{label} cannot be empty")));
    }
    if value.chars().any(|character| {
        matches!(
            character,
            '\u{0000}'..='\u{0008}' | '\u{000B}'..='\u{000C}' | '\u{000E}'..='\u{001F}'
        )
    }) {
        return Err(invalid(format!(
            "{label} contains an invalid XML character"
        )));
    }
    Ok(())
}

pub(crate) fn validate_program(value: &str) -> Result<()> {
    let mut chars = value.chars();
    if value.chars().count() > 39
        || !chars
            .next()
            .is_some_and(|value| value.is_ascii_alphabetic())
        || !chars.all(|value| value.is_ascii_alphanumeric() || value == '.')
    {
        return Err(invalid("invalid OLE program ID"));
    }
    Ok(())
}

pub(crate) fn validate_name(value: &str) -> Result<()> {
    if value.is_empty() || value.chars().count() > MAX_NAME_CHARS {
        return Err(invalid("OLE object name is empty or too long"));
    }
    validate_text(value, "OLE object name", false)
}

pub(crate) fn validate_anchor(anchor: Frame) -> Result<()> {
    if anchor.cx <= 0 || anchor.cy <= 0 {
        return Err(invalid("OLE anchor extents must be positive"));
    }
    Ok(())
}

pub(crate) fn validate_target(
    mode: Mode,
    target: Option<&str>,
    payload: Option<&[u8]>,
) -> Result<()> {
    match mode {
        Mode::Embedded => {
            if target.is_some() {
                return Err(invalid(
                    "embedded OLE objects cannot use an external target",
                ));
            }
            let payload = payload.ok_or_else(|| invalid("embedded OLE object has no payload"))?;
            validate_payload(payload)?;
        },
        Mode::Linked => {
            let target = target.ok_or_else(|| invalid("linked OLE object has no target"))?;
            validate_text(target, "OLE link target", false)?;
            if payload.is_some() {
                return Err(invalid(
                    "linked OLE objects cannot carry an embedded payload",
                ));
            }
        },
    }
    Ok(())
}

pub(crate) fn validate_kind(kind: Kind) -> Result<()> {
    let _ = kind;
    Ok(())
}

pub(crate) fn invalid_revision() -> Error {
    invalid("OLE slide patch source is stale")
}
