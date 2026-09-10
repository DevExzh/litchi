//! Inert RTF tracked-move bookmark ranges.
//!
//! RTF 1.9.1 pairs each `\mvfmf`/`\mvfml` (Move From) or
//! `\mvtof`/`\mvtol` (Move To) start/end range by an alphanumeric tag.  The
//! same tag links the two locations when both are present.  This module
//! retains each complete main-body start/end range as metadata, but does not expose a
//! cross-location pair object or apply the specification's deleted/inserted
//! fallback when one location is absent; it never executes a move operation.

use crate::{RtfError, RtfResult};
use std::borrow::Cow;

pub(crate) const MAX_MOVE_BOOKMARKS: usize = 65_536;
pub(crate) const MAX_MOVE_BOOKMARK_TAG_BYTES: usize = 20;
pub(crate) const MAX_MOVE_BOOKMARK_TOTAL_BYTES: usize = 16 * 1_048_576;

/// Whether a move bookmark identifies the source or destination location.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum MoveBookmarkKind {
    /// `\mvfmf`/`\mvfml`, the source of the move.
    From,
    /// `\mvtof`/`\mvtol`, the destination of the move.
    To,
}

impl MoveBookmarkKind {
    pub(crate) const fn start_control(self) -> &'static str {
        match self {
            Self::From => "mvfmf",
            Self::To => "mvtof",
        }
    }

    pub(crate) const fn end_control(self) -> &'static str {
        match self {
            Self::From => "mvfml",
            Self::To => "mvtol",
        }
    }
}

/// One inert, complete start/end tracked-move bookmark range.
///
/// `kind` identifies the Move From or Move To location.  A second range with
/// the same tag and the opposite kind is the corresponding cross-location
/// range in the source; this bounded model retains that shared tag without
/// exposing a pair index or applying move semantics.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct MoveBookmark<'a> {
    /// Move source or destination.
    pub kind: MoveBookmarkKind,
    /// Unique alphanumeric tag matching this range's controls and, when
    /// present, the opposite-kind move location.
    pub tag: Cow<'a, str>,
    /// Revision-table author index encoded in the opening six-byte payload.
    pub author: u16,
    /// Packed RTF DTTM value encoded in the opening six-byte payload.
    pub date: u32,
    /// UTF-8 byte offset where the move bookmark opens in body text.
    pub position: usize,
    /// Body text covered by this move bookmark.
    pub content: Cow<'a, str>,
}

impl<'a> MoveBookmark<'a> {
    /// Construct a validated inert move bookmark.
    pub fn new(
        kind: MoveBookmarkKind,
        tag: Cow<'a, str>,
        author: u16,
        date: u32,
        position: usize,
        content: Cow<'a, str>,
    ) -> RtfResult<Self> {
        let bookmark = Self {
            kind,
            tag,
            author,
            date,
            position,
            content,
        };
        bookmark.validate()?;
        Ok(bookmark)
    }

    pub(crate) fn validate(&self) -> RtfResult<()> {
        if self.tag.is_empty() || self.tag.len() > MAX_MOVE_BOOKMARK_TAG_BYTES {
            return Err(RtfError::MalformedDocument(
                "RTF move-bookmark tag must contain 1..=20 bytes".to_string(),
            ));
        }
        if !self.tag.bytes().all(|byte| byte.is_ascii_alphanumeric()) {
            return Err(RtfError::MalformedDocument(
                "RTF move-bookmark tag must be alphanumeric".to_string(),
            ));
        }
        if self.content.len() > MAX_MOVE_BOOKMARK_TOTAL_BYTES {
            return Err(RtfError::MalformedDocument(
                "RTF move-bookmark content exceeds the safety limit".to_string(),
            ));
        }
        Ok(())
    }

    pub(crate) fn into_owned(self) -> MoveBookmark<'static> {
        MoveBookmark {
            kind: self.kind,
            tag: Cow::Owned(self.tag.into_owned()),
            author: self.author,
            date: self.date,
            position: self.position,
            content: Cow::Owned(self.content.into_owned()),
        }
    }
}
