//! Semantic and wire-boundary validation for bookmark summaries.

use super::model::{Bookmark, Summary};
use crate::consts::RecordType;
use crate::package::{Error, Result};
use crate::records::Record;
use std::collections::HashSet;

pub(super) const NAME_BYTES: usize = 64;
pub(super) const ENTITY_BYTES: usize = 68;
pub(super) const MAX_VALUE_BYTES: usize = 510;
pub(super) const MAX_BOOKMARKS: usize = 4_096;

impl Summary {
    /// Validate summary IDs against document `TextBookMarkAtom` identifiers.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn validate_text_bookmark_ids(
        &self,
        text_bookmark_ids: impl IntoIterator<Item = u32>,
    ) -> Result<()> {
        let mut text_ids = HashSet::new();
        for id in text_bookmark_ids {
            if text_ids.len() >= MAX_BOOKMARKS {
                return corrupted(format!(
                    "text bookmark collection exceeds {MAX_BOOKMARKS} entries"
                ));
            }
            if !text_ids.insert(id) {
                return corrupted(format!("duplicate TextBookMarkAtom ID {id}"));
            }
        }
        for bookmark in &self.bookmarks {
            if !text_ids.contains(&bookmark.id) {
                return corrupted(format!(
                    "summary bookmark ID {} has no TextBookMarkAtom",
                    bookmark.id
                ));
            }
        }
        if text_ids.len() != self.bookmarks.len() {
            return corrupted("a TextBookMarkAtom has no summary bookmark entity");
        }
        self.validate_seed(text_ids)
    }

    pub(super) fn validate_seed(&self, other_ids: impl IntoIterator<Item = u32>) -> Result<()> {
        validate_seed(self, other_ids)
    }
}

/// Validate the complete authored summary before it reaches the record codec.
pub(super) fn validate_summary(summary: &Summary) -> Result<()> {
    if summary.bookmarks.len() > MAX_BOOKMARKS {
        return corrupted(format!(
            "bookmark collection exceeds {MAX_BOOKMARKS} entries"
        ));
    }

    let mut ids = HashSet::with_capacity(summary.bookmarks.len());
    for bookmark in &summary.bookmarks {
        validate_bookmark(bookmark)?;
        if !ids.insert(bookmark.id) {
            return corrupted(format!("duplicate PowerPoint bookmark ID {}", bookmark.id));
        }
    }
    validate_seed(summary, std::iter::empty())
}

/// Validate one semantic bookmark independently of its containing summary.
pub(super) fn validate_bookmark(bookmark: &Bookmark) -> Result<()> {
    if bookmark.container_instance > 0x0fff {
        return corrupted("bookmark container instance exceeds 12 bits");
    }

    let name_units = bookmark.name.encode_utf16().collect::<Vec<_>>();
    if name_units.is_empty() || name_units.len() > NAME_BYTES / 2 || name_units.contains(&0) {
        return corrupted("bookmarkName must contain 1 through 32 non-null UTF-16 code units");
    }

    let value_units = bookmark.value.encode_utf16().collect::<Vec<_>>();
    if value_units.contains(&0)
        || value_units
            .iter()
            .any(|unit| matches!(*unit, 0x0001..=0x001f | 0x007f..=0x009f))
    {
        return corrupted("BookmarkValueAtom contains a non-printable character");
    }
    let encoded_units = if value_units.is_empty() {
        1
    } else {
        value_units.len()
    };
    if encoded_units > MAX_VALUE_BYTES / 2 {
        return corrupted("BookmarkValueAtom exceeds 255 UTF-16 code units");
    }
    Ok(())
}

fn validate_seed(summary: &Summary, other_ids: impl IntoIterator<Item = u32>) -> Result<()> {
    let max_id = summary
        .bookmarks
        .iter()
        .map(|bookmark| bookmark.id)
        .chain(other_ids)
        .max();
    if max_id.is_some_and(|id| summary.id_seed <= id) {
        return corrupted("bookmark ID seed must exceed every existing bookmark ID");
    }
    Ok(())
}

pub(super) fn require_header(
    record: &Record,
    version: u16,
    instance: u16,
    record_type: RecordType,
    context: &str,
) -> Result<()> {
    if record.version != version
        || record.instance != instance
        || record.record_type_raw != record_type.as_u16()
    {
        return corrupted(format!("invalid {context} record header"));
    }
    Ok(())
}

pub(super) fn corrupted<T>(message: impl Into<String>) -> Result<T> {
    Err(Error::Corrupted(message.into()))
}
