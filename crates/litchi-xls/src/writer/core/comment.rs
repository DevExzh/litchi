use crate::{XlsError, XlsResult};

/// Cell-relative rectangle used by the OfficeArt comment shape.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct XlsCommentAnchor {
    pub move_with_cells: bool,
    pub size_with_cells: bool,
    pub first_column: u16,
    pub first_column_offset: u16,
    pub first_row: u32,
    pub first_row_offset: u16,
    pub last_column: u16,
    pub last_column_offset: u16,
    pub last_row: u32,
    pub last_row_offset: u16,
}

impl XlsCommentAnchor {
    pub(crate) fn validate(&self) -> XlsResult<()> {
        if self.first_column > 255
            || self.last_column > 255
            || self.first_row > 65_535
            || self.last_row > 65_535
            || self.first_column_offset > 1023
            || self.last_column_offset > 1023
            || self.first_row_offset > 255
            || self.last_row_offset > 255
            || self.first_column > self.last_column
            || self.first_row > self.last_row
            || (self.move_with_cells && !self.size_with_cells)
        {
            return Err(XlsError::InvalidData(
                "comment anchor is outside BIFF8 bounds or has invalid movement flags".to_string(),
            ));
        }
        Ok(())
    }

    pub(crate) fn default_for_cell(row: u16, column: u8) -> Self {
        let first_column = u16::from(column).saturating_add(1).min(252);
        let first_row = u32::from(row).min(65_531);
        Self {
            move_with_cells: true,
            size_with_cells: true,
            first_column,
            first_column_offset: 0,
            first_row,
            first_row_offset: 0,
            last_column: first_column + 3,
            last_column_offset: 0,
            last_row: first_row + 4,
            last_row_offset: 0,
        }
    }
}

/// One ordered rich-text run in a comment.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct XlsCommentTextRunWrite {
    pub character_index: u16,
    pub font_index: u16,
}

/// Options for a canonical BIFF8 cell comment.
#[derive(Debug, Clone, Default)]
pub struct XlsCommentWriteOptions {
    pub visible: bool,
    pub shared: bool,
    pub anchor: Option<XlsCommentAnchor>,
    pub text_runs: Vec<XlsCommentTextRunWrite>,
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
    pub options: XlsCommentWriteOptions,
}

impl WritableComment {
    pub(super) fn anchor(&self) -> XlsCommentAnchor {
        self.options
            .anchor
            .unwrap_or_else(|| XlsCommentAnchor::default_for_cell(self.row, self.column))
    }
}

pub(crate) fn validate_comment(
    row: u32,
    column: u16,
    author: &str,
    text: &str,
    options: &XlsCommentWriteOptions,
) -> XlsResult<(u16, u8)> {
    let row = u16::try_from(row).map_err(|_| {
        XlsError::InvalidData("comment row exceeds the BIFF8 limit of 65535".to_string())
    })?;
    let column = u8::try_from(column).map_err(|_| {
        XlsError::InvalidData("comment column exceeds the BIFF8 limit of 255".to_string())
    })?;
    let author_len = author.encode_utf16().count();
    if !(1..=54).contains(&author_len) {
        return Err(XlsError::InvalidData(
            "comment author length must be 1..=54 UTF-16 code units".to_string(),
        ));
    }
    let text_len = text.encode_utf16().count();
    if text_len > usize::from(u16::MAX) {
        return Err(XlsError::InvalidData(
            "comment text exceeds 65535 UTF-16 code units".to_string(),
        ));
    }
    if let Some(anchor) = options.anchor {
        anchor.validate()?;
    }
    if text_len == 0 && !options.text_runs.is_empty() {
        return Err(XlsError::InvalidData(
            "empty comment text cannot contain formatting runs".to_string(),
        ));
    }
    if !options.text_runs.is_empty() {
        if options.text_runs[0].character_index != 0 {
            return Err(XlsError::InvalidData(
                "the first comment formatting run must start at character zero".to_string(),
            ));
        }
        let mut previous = None;
        for run in &options.text_runs {
            if usize::from(run.character_index) >= text_len
                || previous.is_some_and(|value| value >= run.character_index)
            {
                return Err(XlsError::InvalidData(
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
        .ok_or_else(|| {
            XlsError::InvalidData("comment formatting run size overflows".to_string())
        })?;
    if run_bytes > 65_528 {
        return Err(XlsError::InvalidData(
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
