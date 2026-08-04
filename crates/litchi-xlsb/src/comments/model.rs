//! Semantic XLSB comment values.

/// A font-formatting run in a comment's rich string.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct CommentRun {
    /// UTF-16 character index where this run begins.
    pub character_index: u16,
    /// Workbook font index for the run.
    pub font_id: u16,
}

/// A cell comment in a comments part.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Comment {
    /// Row (zero-based).
    pub row: u32,
    /// Column (zero-based).
    pub col: u32,
    /// Comment author.
    pub author: String,
    /// Plain comment text.
    pub text: String,
    /// Rich-string font runs.
    pub runs: Vec<CommentRun>,
    /// Comment GUID from `BrtBeginComment`.
    pub guid: [u8; 16],
    /// Optional GUID carried by an alternate-content `BrtUid` record.
    pub alternate_guid: Option<[u8; 16]>,
    /// Host-side visibility state. The comments stream itself does not carry
    /// this value, so the codec reads it as `false` and does not serialize it.
    pub visible: bool,
}

impl Comment {
    /// Create a comment with no formatting runs or identifiers.
    #[must_use]
    pub fn new(row: u32, col: u32, author: String, text: String) -> Self {
        Self {
            row,
            col,
            author,
            text,
            runs: Vec::new(),
            guid: [0; 16],
            alternate_guid: None,
            visible: false,
        }
    }

    /// Set the host-side visibility state.
    pub fn set_visible(&mut self, visible: bool) {
        self.visible = visible;
    }
}
