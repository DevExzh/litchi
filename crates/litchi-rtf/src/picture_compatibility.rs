//! Positional Word picture compatibility wrappers.

use crate::{RtfError, RtfResult};

pub const MAX_PICTURE_COMPATIBILITY_RECORDS: usize = 65_536;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum PictureCompatibilityKind {
    /// Preferred Word 97-2002 shape picture (`shppict`).
    ShapePicture,
    /// Compatibility fallback picture (`nonshppict`).
    NonShapePicture,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct PictureCompatibilityRecord {
    /// UTF-8 byte position in visible document text.
    pub position: usize,
    /// Wrapper kind.
    pub kind: PictureCompatibilityKind,
    /// Index into `RtfDocument::pictures()`.
    pub picture_index: usize,
}

impl PictureCompatibilityRecord {
    pub fn validate(&self, body: &str, picture_count: usize) -> RtfResult<()> {
        if body.get(self.position..self.position).is_none() {
            return Err(RtfError::MalformedDocument(
                "RTF picture-compatibility position is not a UTF-8 body boundary".to_string(),
            ));
        }
        if self.picture_index >= picture_count {
            return Err(RtfError::MalformedDocument(
                "RTF picture-compatibility index is outside the picture store".to_string(),
            ));
        }
        Ok(())
    }
}
