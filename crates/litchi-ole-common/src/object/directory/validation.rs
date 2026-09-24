//! Semantic checks for the typed CFB directory projection.

use super::model::{EntryKind, Metadata};
use super::{catalog::Projection, codec};
use litchi_cfb::DirectoryEntry;
use litchi_cfb::OleError;
use std::collections::HashSet;

fn validate_shape(metadata: Metadata) -> Result<(), OleError> {
    if metadata.sid().raw() > super::model::MAX_REGULAR_SID {
        return Err(OleError::InvalidFormat(
            "CFB directory metadata contains an invalid SID".into(),
        ));
    }

    for link in [
        metadata.links().left(),
        metadata.links().right(),
        metadata.links().child(),
    ]
    .into_iter()
    .flatten()
    {
        if link == metadata.sid() {
            return Err(OleError::InvalidFormat(
                "CFB directory metadata contains a self-referential link".into(),
            ));
        }
    }

    match metadata.kind() {
        EntryKind::Storage => {
            if metadata.stream_size() != 0 || metadata.uses_mini_stream() {
                return Err(OleError::InvalidFormat(
                    "CFB storage metadata contains stream-only fields".into(),
                ));
            }
        },
        EntryKind::Stream => {
            if metadata.class_id().is_some() || metadata.links().child().is_some() {
                return Err(OleError::InvalidFormat(
                    "CFB stream metadata contains storage-only fields".into(),
                ));
            }
        },
        EntryKind::Root => {
            if metadata.sid().raw() != 0
                || metadata.links().left().is_some()
                || metadata.links().right().is_some()
                || metadata.uses_mini_stream()
            {
                return Err(OleError::InvalidFormat(
                    "CFB root metadata violates root-directory invariants".into(),
                ));
            }
        },
    }
    Ok(())
}

/// Validates fields whose shape is needed to expose a source projection.
///
/// CFB producers in the wild leave stale values in fields that are
/// meaningless for a particular directory kind.  A source-preserving
/// catalog must be able to retain those values for a no-op edit, while typed
/// edits use [`validate`] below before publishing a new value.
pub(super) fn validate_source(metadata: Metadata) -> Result<(), OleError> {
    validate_shape(metadata)
}

/// Validates a typed metadata value before a mutation is committed.
pub(super) fn validate(metadata: Metadata) -> Result<(), OleError> {
    validate_shape(metadata)?;
    match metadata.kind() {
        EntryKind::Storage => {
            // [MS-CFB] 2.6.1 requires the storage starting-sector field to
            // be zero.  ENDOFCHAIN is tolerated only while replaying a raw
            // source value through `validate_source`.
            if metadata.start_sector() != 0 {
                return Err(OleError::InvalidFormat(
                    "CFB storage metadata must use a zero starting sector".into(),
                ));
            }
        },
        EntryKind::Stream => {
            // Stream timestamps are required to be zero on newly published
            // typed values.  Existing nonconforming source words remain
            // available to source-preserving no-op catalogs.
            if metadata.creation_time() != 0 || metadata.modified_time() != 0 {
                return Err(OleError::InvalidFormat(
                    "CFB stream metadata must use zero timestamps".into(),
                ));
            }
        },
        EntryKind::Root => {
            if metadata.creation_time() != 0 {
                return Err(OleError::InvalidFormat(
                    "CFB root metadata must use a zero creation time".into(),
                ));
            }
        },
    }
    Ok(())
}

pub(crate) fn validate_catalog(
    entries: &[DirectoryEntry],
    limits: super::model::Limits,
) -> Result<Vec<Projection>, OleError> {
    limits.validate()?;
    if entries.len() > limits.max_entries {
        return Err(OleError::InvalidFormat(format!(
            "CFB directory catalog contains {} entries, exceeding limit {}",
            entries.len(),
            limits.max_entries
        )));
    }
    let root = entries.first().ok_or_else(|| {
        OleError::InvalidFormat("CFB directory catalog must begin with SID-zero Root Entry".into())
    })?;
    if root.sid != 0 || root.entry_type != EntryKind::Root.raw() || root.name != "Root Entry" {
        return Err(OleError::InvalidFormat(
            "CFB directory catalog must begin with SID-zero Root Entry".into(),
        ));
    }

    let mut seen = HashSet::new();
    seen.try_reserve(entries.len())
        .map_err(|source| OleError::Allocation {
            resource: "CFB directory catalog SID index",
            source,
        })?;
    let mut pending = Vec::new();
    pending
        .try_reserve(entries.len())
        .map_err(|source| OleError::Allocation {
            resource: "CFB directory catalog validation stack",
            source,
        })?;
    for entry in entries {
        pending.push((entry, 0usize));
    }

    let mut total_bytes = 0usize;
    let mut raw_children = 0usize;
    while let Some((entry, depth)) = pending.pop() {
        if depth > limits.max_raw_depth {
            return Err(OleError::InvalidFormat(
                "CFB directory raw child depth exceeds the catalog limit".into(),
            ));
        }
        seen.try_reserve(1).map_err(|source| OleError::Allocation {
            resource: "CFB directory catalog SID index",
            source,
        })?;
        if !seen.insert(entry.sid) {
            return Err(OleError::InvalidFormat(format!(
                "CFB directory catalog contains duplicate SID {}",
                entry.sid
            )));
        }
        validate_name(&entry.name, limits.max_name_bytes)?;
        if entry.clsid.len() > limits.max_name_bytes {
            return Err(OleError::InvalidFormat(
                "CFB directory CLSID exceeds the catalog name limit".into(),
            ));
        }
        total_bytes = total_bytes
            .checked_add(entry.name.len())
            .and_then(|value| value.checked_add(entry.clsid.len()))
            .ok_or_else(|| {
                OleError::InvalidFormat("CFB directory catalog size overflows".into())
            })?;
        if total_bytes > limits.max_total_bytes {
            return Err(OleError::InvalidFormat(
                "CFB directory catalog metadata exceeds the total byte limit".into(),
            ));
        }
        let _ = codec::decode_links(entry)?;
        for child in &entry.children {
            raw_children = raw_children
                .checked_add(1)
                .ok_or_else(|| OleError::InvalidFormat("CFB raw child count overflows".into()))?;
            if raw_children > limits.max_raw_children {
                return Err(OleError::InvalidFormat(
                    "CFB directory raw child count exceeds the catalog limit".into(),
                ));
            }
            pending
                .try_reserve(1)
                .map_err(|source| OleError::Allocation {
                    resource: "CFB directory catalog validation stack",
                    source,
                })?;
            pending.push((child, depth + 1));
        }
    }

    entries.iter().map(codec::project).collect()
}

fn validate_name(name: &str, max_bytes: usize) -> Result<(), OleError> {
    if name.is_empty() {
        return Err(OleError::InvalidFormat(
            "CFB directory entry name is empty".into(),
        ));
    }
    if name.len() > max_bytes {
        return Err(OleError::InvalidFormat(
            "CFB directory entry name exceeds the catalog limit".into(),
        ));
    }
    if name.contains('\0') {
        return Err(OleError::InvalidFormat(
            "CFB directory entry name contains NUL".into(),
        ));
    }
    if name
        .chars()
        .any(|character| matches!(character, '/' | '\\' | ':' | '!'))
    {
        return Err(OleError::InvalidFormat(
            "CFB directory entry name contains a forbidden character".into(),
        ));
    }
    if name.encode_utf16().count() > 31 {
        return Err(OleError::InvalidFormat(
            "CFB directory entry name exceeds 31 UTF-16 code units".into(),
        ));
    }
    Ok(())
}

pub(crate) fn validate_name_for_edit(
    name: &str,
    limits: super::model::Limits,
) -> Result<(), OleError> {
    validate_name(name, limits.max_name_bytes)
}
