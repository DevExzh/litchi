use super::shape::Anchor;
use crate::{Error, Result};

/// One ordered rich-text run in a comment.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct CommentTextRunWrite {
    pub character_index: u16,
    pub font_index: u16,
}

/// Options for a canonical BIFF8 cell comment.
#[derive(Debug, Clone, Default)]
pub struct CommentWriteOptions {
    pub visible: bool,
    pub shared: bool,
    pub anchor: Option<Anchor>,
    pub text_runs: Vec<CommentTextRunWrite>,
    pub font_when_empty: u16,
    /// Stable GUID override. When absent, the writer derives a deterministic,
    /// worksheet- and cell-specific GUID.
    pub guid: Option<[u8; 16]>,
}

#[derive(Debug, Clone)]
pub(super) struct WritableComment {
    pub row: u16,
    pub column: u8,
    pub author: String,
    pub text: String,
    pub options: CommentWriteOptions,
}

impl WritableComment {
    pub(super) fn try_new(
        row: u16,
        column: u8,
        author: &str,
        text: &str,
        options: CommentWriteOptions,
    ) -> Result<Self> {
        let mut owned_author = String::new();
        owned_author
            .try_reserve_exact(author.len())
            .map_err(|_| Error::Allocation("reserving comment-author storage"))?;
        owned_author.push_str(author);

        let mut owned_text = String::new();
        owned_text
            .try_reserve_exact(text.len())
            .map_err(|_| Error::Allocation("reserving comment-text storage"))?;
        owned_text.push_str(text);

        Ok(Self {
            row,
            column,
            author: owned_author,
            text: owned_text,
            options,
        })
    }

    pub(super) fn anchor(&self) -> Anchor {
        self.options
            .anchor
            .unwrap_or_else(|| Anchor::default_for_cell(self.row, self.column))
    }
}

pub(crate) fn validate_comment(
    row: u32,
    column: u16,
    author: &str,
    text: &str,
    options: &CommentWriteOptions,
) -> Result<(u16, u8)> {
    let row = u16::try_from(row).map_err(|_| {
        Error::InvalidData("comment row exceeds the BIFF8 limit of 65535".to_string())
    })?;
    let column = u8::try_from(column).map_err(|_| {
        Error::InvalidData("comment column exceeds the BIFF8 limit of 255".to_string())
    })?;
    let author_len = author.encode_utf16().count();
    if !(1..=54).contains(&author_len) {
        return Err(Error::InvalidData(
            "comment author length must be 1..=54 UTF-16 code units".to_string(),
        ));
    }
    let text_len = text.encode_utf16().count();
    if text_len > usize::from(u16::MAX) {
        return Err(Error::InvalidData(
            "comment text exceeds 65535 UTF-16 code units".to_string(),
        ));
    }
    if text_len == 0 && !options.text_runs.is_empty() {
        return Err(Error::InvalidData(
            "empty comment text cannot contain formatting runs".to_string(),
        ));
    }
    if !options.text_runs.is_empty() {
        if options.text_runs[0].character_index != 0 {
            return Err(Error::InvalidData(
                "the first comment formatting run must start at character zero".to_string(),
            ));
        }
        let mut previous = None;
        for run in &options.text_runs {
            if usize::from(run.character_index) >= text_len
                || previous.is_some_and(|value| value >= run.character_index)
            {
                return Err(Error::InvalidData(
                    "comment formatting runs must be strictly ordered within the text".to_string(),
                ));
            }
            previous = Some(run.character_index);
        }
    }
    let run_count = if text_len == 0 {
        0
    } else {
        options.text_runs.len().max(1)
    };
    let run_bytes = run_count
        .checked_add(1)
        .and_then(|count| count.checked_mul(8))
        .ok_or_else(|| Error::InvalidData("comment formatting run size overflows".to_string()))?;
    if run_bytes > 65_528 {
        return Err(Error::InvalidData(
            "comment formatting runs exceed the BIFF8 cbRuns limit".to_string(),
        ));
    }
    Ok((row, column))
}

pub(crate) fn deterministic_comment_guid(
    sheet_index: usize,
    row: u16,
    column: u8,
    object_id: u16,
) -> [u8; 16] {
    let mut guid = [0u8; 16];
    guid[0..4].copy_from_slice(b"LTCM");
    guid[4..8].copy_from_slice(&(sheet_index as u32).to_le_bytes());
    guid[8..10].copy_from_slice(&row.to_le_bytes());
    guid[10] = column;
    guid[11] = 0x40;
    guid[12..14].copy_from_slice(&object_id.to_le_bytes());
    guid[14] = 0x80;
    guid[15] = 1;
    guid
}
