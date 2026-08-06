//! Semantic checks for the typed CFB directory projection.

use super::model::{EntryKind, Metadata};
use litchi_cfb::OleError;
use litchi_cfb::consts::ENDOFCHAIN;

pub(super) fn validate(metadata: Metadata) -> Result<(), OleError> {
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
            if metadata.stream_size() != 0
                || metadata.uses_mini_stream()
                || !matches!(metadata.start_sector(), 0 | ENDOFCHAIN)
            {
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
