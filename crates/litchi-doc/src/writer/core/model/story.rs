use super::{CharacterFormatting, ParagraphFormatting, WriteError};
use std::collections::HashMap;

pub(in crate::writer::core) type NoteStoryData = (Vec<u8>, Vec<u8>, u32);
pub(in crate::writer::core) struct HeaderStoryData {
    pub(in crate::writer::core) plcfhdd: Vec<u8>,
    pub(in crate::writer::core) fields: Vec<u8>,
    pub(in crate::writer::core) char_count: u32,
    /// Story-relative anchor CPs of header floating items with their kind
    /// (in story order, which is CP-ascending by construction).
    pub(in crate::writer::core) shape_anchor_cps: Vec<(u32, FloatingAnchorKind)>,
}

/// `PlcfHdd` slot of the odd page header, which Word uses as the default
/// header when the document does not use facing pages.
pub(in crate::writer::core) const HEADER_SLOT_ODD: usize = 7;
/// `PlcfHdd` slot of the even page header.
pub(in crate::writer::core) const HEADER_SLOT_EVEN: usize = 6;
/// `PlcfHdd` slot of the first page header.
pub(in crate::writer::core) const HEADER_SLOT_FIRST: usize = 10;

/// Which header a floating text box or picture is anchored in.
///
/// The writer emits the section properties each kind needs automatically:
/// even headers enable DOP `fFacingPages`, first-page headers enable SEP
/// `fTitlePage`, because appending the anchor creates the corresponding
/// header story.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum HeaderKind {
    /// Odd page header; Word's default header.
    Odd,
    /// Even page header (requires facing pages, enabled automatically).
    Even,
    /// First page header (requires a different first page, enabled
    /// automatically).
    FirstPage,
}

impl HeaderKind {
    /// The `PlcfHdd` slot holding this header kind's story.
    pub(in crate::writer::core) fn slot(self) -> usize {
        match self {
            Self::Odd => HEADER_SLOT_ODD,
            Self::Even => HEADER_SLOT_EVEN,
            Self::FirstPage => HEADER_SLOT_FIRST,
        }
    }
}

/// A floating-item anchor paragraph appended to a header's paragraphs.
pub(in crate::writer::core) struct HeaderAnchor {
    /// `PlcfHdd` slot of the header holding the anchor.
    pub(in crate::writer::core) slot: usize,
    /// Paragraph index within that slot's paragraph list.
    pub(in crate::writer::core) paragraph_index: usize,
    /// Which floating item the anchor belongs to.
    pub(in crate::writer::core) kind: FloatingAnchorKind,
}

pub(in crate::writer::core) struct CommentStoryData {
    pub(in crate::writer::core) owners: Vec<u8>,
    pub(in crate::writer::core) references: Vec<u8>,
    pub(in crate::writer::core) text_positions: Vec<u8>,
    pub(in crate::writer::core) bookmark_names: Vec<u8>,
    pub(in crate::writer::core) bookmark_starts: Vec<u8>,
    pub(in crate::writer::core) bookmark_ends: Vec<u8>,
    pub(in crate::writer::core) extended_metadata: Vec<u8>,
    pub(in crate::writer::core) char_count: u32,
}

pub(in crate::writer::core) struct BookmarkTableData {
    pub(in crate::writer::core) names: Vec<u8>,
    pub(in crate::writer::core) starts: Vec<u8>,
    pub(in crate::writer::core) ends: Vec<u8>,
}

pub(in crate::writer::core) struct RevisionWriterData {
    pub(in crate::writer::core) indexes: HashMap<String, u16>,
    pub(in crate::writer::core) table: Vec<u8>,
}

#[derive(Clone, Copy)]
pub(in crate::writer::core) enum MainReferenceKind {
    Footnote,
    Endnote,
    Comment,
}

/// Represents a text run with formatting
#[derive(Debug, Clone)]
pub(in crate::writer::core) struct TextRun {
    /// Text content
    pub(in crate::writer::core) text: String,
    /// Character formatting
    pub(in crate::writer::core) formatting: CharacterFormatting,
    /// Index into `Writer::pictures` when this run is a picture
    /// (a single 0x0001 inline or 0x0008 floating picture character).
    pub(in crate::writer::core) picture_index: Option<u32>,
    /// Index into `Writer::shapes` when this run is a floating
    /// drawing-shape anchor (a single 0x0008 character).
    pub(in crate::writer::core) shape_index: Option<u32>,
}

/// Represents a paragraph
#[derive(Debug, Clone)]
#[allow(
    dead_code,
    reason = "reserved DOC structure retained for format completeness or future round-trip support"
)] // Reserved for future implementation
pub(in crate::writer::core) struct WritableParagraph {
    /// Text runs in this paragraph
    pub(in crate::writer::core) runs: Vec<TextRun>,
    /// Paragraph formatting
    pub(in crate::writer::core) formatting: ParagraphFormatting,
}

pub(in crate::writer::core) fn writable_paragraph_from_runs(
    runs: Vec<(String, CharacterFormatting)>,
    formatting: ParagraphFormatting,
) -> WritableParagraph {
    let runs = if runs.is_empty() {
        vec![TextRun {
            text: String::new(),
            formatting: CharacterFormatting::default(),
            picture_index: None,
            shape_index: None,
        }]
    } else {
        runs.into_iter()
            .map(|(text, formatting)| TextRun {
                text,
                formatting,
                picture_index: None,
                shape_index: None,
            })
            .collect()
    };
    WritableParagraph { runs, formatting }
}

/// One formatted paragraph in a header or footer story.
///
/// The paragraph owns its runs so callers can build content without tying the
/// writer to temporary buffers. Paragraph marks are emitted by the writer and
/// therefore MUST NOT be embedded in run text passed to [`Self::from_runs`].
#[derive(Debug, Clone)]
pub struct HeaderFooterParagraph {
    pub(in crate::writer::core) runs: Vec<(String, CharacterFormatting)>,
    pub(in crate::writer::core) formatting: ParagraphFormatting,
}

impl HeaderFooterParagraph {
    /// Construct a plain paragraph containing one run.
    pub fn plain(text: impl Into<String>) -> Self {
        Self {
            runs: vec![(text.into(), CharacterFormatting::default())],
            formatting: ParagraphFormatting::default(),
        }
    }

    /// Construct a paragraph from owned character runs and paragraph formatting.
    #[must_use]
    pub fn from_runs(
        runs: Vec<(String, CharacterFormatting)>,
        formatting: ParagraphFormatting,
    ) -> Self {
        Self { runs, formatting }
    }

    /// Construct a paragraph containing one inert Word field.
    ///
    /// The instruction is transported verbatim and is never evaluated.
    pub fn field(
        instruction: impl Into<String>,
        result: impl Into<String>,
        result_formatting: CharacterFormatting,
    ) -> Result<Self, WriteError> {
        let instruction = instruction.into();
        let result = result.into();
        if instruction.is_empty() {
            return Err(WriteError::InvalidData(
                "DOC header/footer field instruction is empty".to_string(),
            ));
        }
        if instruction
            .chars()
            .chain(result.chars())
            .any(|character| character == '\r' || matches!(character as u32, 0x0013..=0x0015))
        {
            return Err(WriteError::InvalidData(
                "DOC header/footer field text contains a structural control character".to_string(),
            ));
        }
        let special = CharacterFormatting {
            special: Some(true),
            ..CharacterFormatting::default()
        };
        let hidden = CharacterFormatting {
            field_vanish: Some(true),
            ..CharacterFormatting::default()
        };
        Ok(Self::from_runs(
            vec![
                ("\u{0013}".to_string(), special.clone()),
                (instruction, hidden),
                ("\u{0014}".to_string(), special.clone()),
                (result, result_formatting),
                ("\u{0015}".to_string(), special),
            ],
            ParagraphFormatting::default(),
        ))
    }

    /// Character runs in document order.
    #[must_use]
    pub fn runs(&self) -> &[(String, CharacterFormatting)] {
        &self.runs
    }

    /// Paragraph-level formatting.
    #[must_use]
    pub fn formatting(&self) -> &ParagraphFormatting {
        &self.formatting
    }
}

/// Represents a table cell
#[derive(Debug, Clone)]
pub(in crate::writer::core) struct TableCell {
    /// Paragraphs in the cell
    pub(in crate::writer::core) paragraphs: Vec<WritableParagraph>,
}

/// Represents a table row
#[derive(Debug, Clone)]
pub(in crate::writer::core) struct TableRow {
    /// Cells in the row
    pub(in crate::writer::core) cells: Vec<TableCell>,
    /// Row and cell layout encoded in the row mark's TAP properties.
    pub(in crate::writer::core) formatting: crate::writer::tap::TableRow,
}

/// Represents a table
#[derive(Debug, Clone)]
pub(in crate::writer::core) struct WritableTable {
    /// Rows in the table
    pub(in crate::writer::core) rows: Vec<TableRow>,
}

/// A picture queued for embedding, with its placement mode.
#[derive(Debug, Clone)]
pub(in crate::writer::core) struct WriterPicture {
    /// The picture data and display dimensions.
    pub(in crate::writer::core) picture: crate::writer::images::Picture,
    /// Shape id allocated at insert time (shared sequence with shapes).
    pub(in crate::writer::core) shape_id: u32,
    /// Position and wrapping when the picture floats; `None` for inline.
    pub(in crate::writer::core) floating: Option<crate::writer::images::FloatingPosition>,
}

/// A primitive drawing shape queued for the drawing layer.
#[derive(Debug, Clone)]
pub(in crate::writer::core) struct WriterShape {
    /// The shape geometry, size, and colors.
    pub(in crate::writer::core) shape: crate::writer::shapes::Shape,
    /// Shape id allocated at insert time (shared sequence with pictures).
    pub(in crate::writer::core) shape_id: u32,
    /// Position and wrapping.
    pub(in crate::writer::core) position: crate::writer::images::FloatingPosition,
    /// Textbox story text when the shape is a text box.
    pub(in crate::writer::core) text: Option<String>,
}

/// What kind of floating content a 0x0008 anchor character refers to.
#[derive(Debug, Clone, Copy)]
pub(in crate::writer::core) enum FloatingAnchorKind {
    /// Index into `Writer::pictures`.
    Picture(u32),
    /// Index into `Writer::shapes`.
    Shape(u32),
}
