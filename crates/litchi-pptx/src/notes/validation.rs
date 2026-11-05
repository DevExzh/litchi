//! Semantic bounds for detached `PresentationML` notes transactions.

use std::collections::HashSet;

use litchi_opc::{PackURI, TargetMode};

use super::model::Link;
use super::transaction::PartState;
use super::{Limits, invalid, limit};
use crate::{Error, Result};

pub(crate) const MAX_RELATIONSHIP_FIELD_BYTES: usize = 16 * 1024;
pub(crate) const MAX_RELATIONSHIPS: usize = 4_096;

pub(crate) fn validate_link(link: &Link) -> Result<()> {
    validate_id(link.id(), "notes relationship ID")?;
    validate_field(link.relationship_type(), "notes relationship type")?;
    validate_field(link.target(), "notes relationship target")?;
    if link.target_mode() == TargetMode::Internal && link.target().is_empty() {
        return Err(invalid("notes internal relationship target is empty"));
    }
    Ok(())
}

pub(crate) fn validate_links(links: &[Link]) -> Result<()> {
    if links.len() > MAX_RELATIONSHIPS {
        return Err(limit(
            crate::error::CodecLimitKind::Relationships,
            links.len(),
            MAX_RELATIONSHIPS,
        ));
    }
    let mut ids = HashSet::new();
    ids.try_reserve(links.len())
        .map_err(|source| Error::Allocation {
            resource: "notes relationship IDs",
            source,
        })?;
    for link in links {
        validate_link(link)?;
        if !ids.insert(link.id()) {
            return Err(invalid("notes relationship IDs must be unique"));
        }
    }
    Ok(())
}

pub(crate) fn validate_parts_with_limits(parts: &[PartState], limits: Limits) -> Result<()> {
    let mut total = 0usize;
    let mut names = HashSet::new();
    names
        .try_reserve(parts.len())
        .map_err(|source| Error::Allocation {
            resource: "notes part names",
            source,
        })?;
    for part in parts {
        let name = part.name.as_str();
        if name.len() > MAX_RELATIONSHIP_FIELD_BYTES {
            return Err(limit(
                crate::error::CodecLimitKind::NameBytes,
                name.len(),
                MAX_RELATIONSHIP_FIELD_BYTES,
            ));
        }
        PackURI::new(name).map_err(Error::Invalid)?;
        if part.content_type.is_empty() {
            return Err(invalid("notes part content type is empty"));
        }
        validate_links(&part.relationships)?;
        total = total
            .checked_add(part.data.len())
            .ok_or_else(|| invalid("notes transaction bytes overflow"))?;
        if total > limits.max_total_bytes {
            return Err(limit(
                crate::error::CodecLimitKind::PayloadBytes,
                total,
                limits.max_total_bytes,
            ));
        }
        if !names.insert(name.to_owned()) {
            return Err(invalid("notes transaction contains duplicate part names"));
        }
    }
    Ok(())
}

pub(crate) fn validate_id(value: &str, label: &'static str) -> Result<()> {
    if value.is_empty() {
        return Err(invalid(format!("{label} is empty")));
    }
    if value.len() > MAX_RELATIONSHIP_FIELD_BYTES {
        return Err(limit(
            crate::error::CodecLimitKind::RelationshipBytes,
            value.len(),
            MAX_RELATIONSHIP_FIELD_BYTES,
        ));
    }
    let mut bytes = value.bytes();
    let Some(first) = bytes.next() else {
        return Err(invalid(format!("{label} is empty")));
    };
    if !(first.is_ascii_alphabetic() || first == b'_')
        || !bytes.all(|byte| byte.is_ascii_alphanumeric() || matches!(byte, b'_' | b'-' | b'.'))
    {
        return Err(invalid(format!("invalid {label}")));
    }
    Ok(())
}

pub(crate) fn validate_field(value: &str, label: &'static str) -> Result<()> {
    if value.len() > MAX_RELATIONSHIP_FIELD_BYTES {
        return Err(limit(
            crate::error::CodecLimitKind::RelationshipBytes,
            value.len(),
            MAX_RELATIONSHIP_FIELD_BYTES,
        ));
    }
    if value.chars().any(|character| {
        matches!(
            character,
            '\u{0000}'..='\u{0008}'
                | '\u{000B}'..='\u{000C}'
                | '\u{000E}'..='\u{001F}'
        )
    }) {
        return Err(invalid(format!(
            "{label} contains an invalid XML character"
        )));
    }
    Ok(())
}
