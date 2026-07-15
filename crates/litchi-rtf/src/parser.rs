//! RTF parser that builds document structure from tokens.

use super::error::{RtfError, RtfResult};
use super::lexer::{ControlWord, Token};
use super::types::*;
use bumpalo::Bump;
use encoding_rs::Encoding;
use litchi_core::encoding::codepage_to_encoding;
use smallvec::SmallVec;
use std::borrow::Cow;
use std::cell::RefCell;
use std::collections::HashMap;
use std::num::NonZeroU16;

/// RTF destination type - determines if we're in document body or header
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum Destination {
    /// Main document body - text should be extracted
    DocumentBody,
    /// Font table - should be skipped
    FontTable,
    /// Color table - should be skipped
    ColorTable,
    /// Stylesheet - should be skipped
    StyleSheet,
    /// Document info - should be skipped
    Info,
    /// Picture data - extract and process embedded images
    Picture,
    /// Embedded object - extract if possible
    Object,
    /// Result of embedded object rendering - should be skipped
    Result,
    /// Field instruction
    FieldInstruction,
    /// Field result
    FieldResult,
    /// Header content
    Header,
    /// Footer content
    Footer,
    /// Footnote content
    Footnote,
    /// Endnote content
    Endnote,
    /// Revision/track changes
    #[allow(dead_code)]
    Revision,
    /// Other destinations - should be skipped
    Other,
}

#[derive(Clone, Copy)]
enum InfoTextField {
    Title,
    Subject,
    Author,
    Manager,
    Company,
    Operator,
    Category,
    Keywords,
    Comment,
}

#[derive(Clone, Copy)]
enum InfoTimeField {
    Creation,
    Revision,
    Print,
    Backup,
}

#[derive(Default)]
struct InfoTimestamp {
    year: Option<i32>,
    month: Option<i32>,
    day: Option<i32>,
    hour: Option<i32>,
    minute: Option<i32>,
    second: Option<i32>,
}

const MAX_INFO_TEXT_BYTES: usize = 1_048_576;
const MAX_BOOKMARKS: usize = 65_536;
const MAX_BOOKMARK_NAME_BYTES: usize = 65_536;
const MAX_ANNOTATIONS: usize = 65_536;
const MAX_ANNOTATION_TEXT_BYTES: usize = 4 * 1_048_576;

struct OpenBookmark {
    name: String,
    position: usize,
    first_column: Option<i32>,
    last_column: Option<i32>,
    is_public: bool,
    order: usize,
}

struct BookmarkSpan {
    bookmark: OpenBookmark,
    end: usize,
}

/// Parser state for tracking formatting context.
#[derive(Debug, Clone)]
struct State {
    /// Current character formatting
    formatting: Formatting,
    /// Current paragraph properties
    paragraph: Paragraph,
    /// Unicode skip count (characters to skip after \u)
    unicode_skip: i32,
    /// Whether we're inside a table
    in_table: bool,
    /// Cell boundaries for current row (in twips)
    cell_boundaries: SmallVec<[i32; 8]>,
    /// Current destination (for skipping non-document content)
    destination: Destination,
    /// Current text encoding
    encoding: &'static Encoding,
}

impl Default for State {
    fn default() -> Self {
        Self {
            formatting: Formatting::default(),
            paragraph: Paragraph::default(),
            unicode_skip: 1,
            in_table: false,
            cell_boundaries: SmallVec::new(),
            destination: Destination::DocumentBody,
            encoding: encoding_rs::WINDOWS_1252, // Default ANSI encoding
        }
    }
}

/// RTF Parser.
pub struct Parser<'a> {
    /// Token stream
    tokens: &'a [Token<'a>],
    /// Current position in token stream
    pos: usize,
    /// State stack (for handling groups)
    states: Vec<State>,
    /// Font table
    font_table: RefCell<FontTable<'a>>,
    /// Color table
    color_table: RefCell<ColorTable>,
    /// Parsed style blocks
    blocks: Vec<StyleBlock<'a>>,
    /// Arena for temporary allocations
    arena: &'a Bump,
    /// Extracted tables
    tables: Vec<super::table::Table<'a>>,
    /// Current table being built
    current_table: Option<super::table::Table<'a>>,
    /// Current row being built
    current_row: Option<super::table::Row<'a>>,
    /// Current cell text buffer
    current_cell_text: SmallVec<[u8; 128]>,
    /// Extracted pictures
    pictures: Vec<super::picture::Picture<'a>>,
    /// Extracted fields
    fields: Vec<super::field::Field<'a>>,
    /// List table
    list_table: super::list::ListTable<'a>,
    /// List override table
    list_override_table: super::list::ListOverrideTable,
    /// Sections
    sections: Vec<super::section::Section<'a>>,
    /// Bookmarks
    bookmarks: super::bookmark::BookmarkTable<'a>,
    /// Open bookmark ranges, indexed by name.
    open_bookmarks: HashMap<String, Vec<OpenBookmark>>,
    /// Completed bookmark ranges awaiting content reconstruction.
    bookmark_spans: Vec<BookmarkSpan>,
    /// UTF-8 byte length of body text emitted into style blocks.
    body_text_len: usize,
    /// Stable source order for bookmark ranges.
    next_bookmark_order: usize,
    /// Shapes
    shapes: Vec<super::shape::Shape<'a>>,
    /// Shape groups
    shape_groups: Vec<super::shape::ShapeGroup<'a>>,
    /// Stylesheet
    stylesheet: super::stylesheet::StyleSheet<'a>,
    /// Document information
    info: super::info::DocumentInfo<'a>,
    /// Annotations
    annotations: Vec<super::annotation::Annotation<'a>>,
    /// Parsed annotation reference ranges by numeric identifier.
    annotation_ranges: HashMap<i32, (usize, Option<usize>)>,
    /// Author metadata immediately preceding an annotation destination.
    pending_annotation_author: String,
    /// Author initials immediately preceding an annotation destination.
    pending_annotation_initials: String,
    /// Footnotes and endnotes
    notes: Vec<super::section::Note<'a>>,
    /// Track changes/revisions
    revisions: Vec<super::annotation::Revision<'a>>,
    /// Current header/footer being parsed
    #[allow(dead_code)]
    current_header_footer: Option<super::section::HeaderFooter<'a>>,
    /// Current note being parsed (content buffer)
    current_note_buffer: SmallVec<[u8; 256]>,
    /// Current header/footer type being parsed
    current_hf_type: Option<super::section::HeaderFooterType>,
}

impl<'a> Parser<'a> {
    /// Create a new parser.
    pub fn new(tokens: &'a [Token<'a>], arena: &'a Bump) -> Self {
        Self {
            tokens,
            pos: 0,
            states: vec![State::default()],
            font_table: RefCell::new(FontTable::new()),
            color_table: RefCell::new(ColorTable::new()),
            blocks: Vec::new(),
            arena,
            tables: Vec::new(),
            current_table: None,
            current_row: None,
            current_cell_text: SmallVec::new(),
            pictures: Vec::new(),
            fields: Vec::new(),
            list_table: super::list::ListTable::new(),
            list_override_table: super::list::ListOverrideTable::new(),
            sections: Vec::new(),
            bookmarks: super::bookmark::BookmarkTable::new(),
            open_bookmarks: HashMap::new(),
            bookmark_spans: Vec::new(),
            body_text_len: 0,
            next_bookmark_order: 0,
            shapes: Vec::new(),
            shape_groups: Vec::new(),
            stylesheet: super::stylesheet::StyleSheet::new(),
            info: super::info::DocumentInfo::new(),
            annotations: Vec::new(),
            annotation_ranges: HashMap::new(),
            pending_annotation_author: String::new(),
            pending_annotation_initials: String::new(),
            notes: Vec::new(),
            revisions: Vec::new(),
            current_header_footer: None,
            current_note_buffer: SmallVec::new(),
            current_hf_type: None,
        }
    }

    /// Parse the token stream into a document.
    pub fn parse(mut self) -> RtfResult<ParsedDocument<'a>> {
        // Validate document structure
        if self.tokens.is_empty() {
            return Err(RtfError::MalformedDocument(
                "Empty token stream".to_string(),
            ));
        }

        // Expect opening brace
        if !matches!(self.tokens.first(), Some(Token::OpenBrace)) {
            return Err(RtfError::MalformedDocument(
                "Document must start with {".to_string(),
            ));
        }

        // Parse document content
        self.parse_group()?;

        // Finalize any remaining table
        self.finalize_table();
        self.finalize_bookmarks()?;

        Ok(ParsedDocument {
            font_table: self.font_table.into_inner(),
            color_table: self.color_table.into_inner(),
            blocks: self.blocks,
            tables: self.tables,
            pictures: self.pictures,
            fields: self.fields,
            list_table: self.list_table,
            list_override_table: self.list_override_table,
            sections: self.sections,
            bookmarks: self.bookmarks,
            shapes: self.shapes,
            shape_groups: self.shape_groups,
            stylesheet: self.stylesheet,
            info: self.info,
            annotations: self.annotations,
            notes: self.notes,
            revisions: self.revisions,
        })
    }

    /// Parse a group (content between braces).
    fn parse_group(&mut self) -> RtfResult<()> {
        self.expect_token(Token::OpenBrace)?;

        // Push new state (inherit from parent)
        if let Some(current) = self.states.last() {
            self.states.push(current.clone());
        } else {
            self.states.push(State::default());
        }

        // Check if this is a special group (header, destination, etc.)
        if self.pos < self.tokens.len() {
            match &self.tokens[self.pos] {
                Token::Control(ControlWord::FontTable) => {
                    // Mark this as font table destination
                    if let Some(state) = self.states.last_mut() {
                        state.destination = Destination::FontTable;
                    }
                    self.parse_font_table()?;
                    self.skip_until_close_brace()?;
                    self.states.pop();
                    return Ok(());
                },
                Token::Control(ControlWord::ColorTable) => {
                    // Mark this as color table destination
                    if let Some(state) = self.states.last_mut() {
                        state.destination = Destination::ColorTable;
                    }
                    self.parse_color_table()?;
                    self.skip_until_close_brace()?;
                    self.states.pop();
                    return Ok(());
                },
                Token::Control(ControlWord::IgnorableDestination) => {
                    if matches!(
                        self.tokens.get(self.pos + 1),
                        Some(Token::Control(
                            ControlWord::BookmarkStart | ControlWord::BookmarkEnd
                        ))
                    ) {
                        self.parse_bookmark_destination()?;
                        self.states.pop();
                        return Ok(());
                    }
                    match self.tokens.get(self.pos + 1) {
                        Some(Token::Control(ControlWord::AnnotationAuthor)) => {
                            self.pending_annotation_author =
                                self.parse_ignorable_text_destination()?;
                            self.states.pop();
                            return Ok(());
                        },
                        Some(Token::Control(ControlWord::AnnotationInitials)) => {
                            self.pending_annotation_initials =
                                self.parse_ignorable_text_destination()?;
                            self.states.pop();
                            return Ok(());
                        },
                        Some(Token::Control(ControlWord::AnnotationRangeStart)) => {
                            self.parse_annotation_range_marker(true)?;
                            self.states.pop();
                            return Ok(());
                        },
                        Some(Token::Control(ControlWord::AnnotationRangeEnd)) => {
                            self.parse_annotation_range_marker(false)?;
                            self.states.pop();
                            return Ok(());
                        },
                        Some(Token::Control(ControlWord::Annotation)) => {
                            self.parse_annotation_destination()?;
                            self.states.pop();
                            return Ok(());
                        },
                        _ => {},
                    }
                    // Mark as other destination and skip
                    if let Some(state) = self.states.last_mut() {
                        state.destination = Destination::Other;
                    }
                    self.skip_until_close_brace()?;
                    self.states.pop();
                    return Ok(());
                },
                Token::Control(ControlWord::StyleSheet) => {
                    // Mark as stylesheet destination and skip
                    if let Some(state) = self.states.last_mut() {
                        state.destination = Destination::StyleSheet;
                    }
                    self.skip_until_close_brace()?;
                    self.states.pop();
                    return Ok(());
                },
                Token::Control(ControlWord::Info) => {
                    // Parse document metadata without adding it to body text.
                    if let Some(state) = self.states.last_mut() {
                        state.destination = Destination::Info;
                    }
                    self.parse_info()?;
                    self.states.pop();
                    return Ok(());
                },
                Token::Control(ControlWord::Picture) => {
                    // Mark as picture destination and extract
                    if let Some(state) = self.states.last_mut() {
                        state.destination = Destination::Picture;
                    }
                    self.parse_picture()?;
                    self.skip_until_close_brace()?;
                    self.states.pop();
                    return Ok(());
                },
                Token::Control(ControlWord::Object) => {
                    // Mark as object destination
                    // Embedded objects in RTF files include:
                    // - MathType/Equation Editor equations
                    // - Excel charts and spreadsheets
                    // - Visio diagrams
                    // - Other OLE-embedded content
                    //
                    // For basic support, we extract object metadata and skip the binary data.
                    // Full OLE parsing would require:
                    // 1. Parse the OLE object structure from hex-encoded binary data
                    // 2. Identify the object type (CLSID/ProgID)
                    // 3. Extract and decode the object's native format
                    // 4. Convert to suitable representation (e.g., LaTeX for equations)
                    if let Some(state) = self.states.last_mut() {
                        state.destination = Destination::Object;
                    }
                    self.skip_until_close_brace()?;
                    self.states.pop();
                    return Ok(());
                },
                Token::Control(ControlWord::Result) => {
                    // Mark as result destination and skip
                    // This contains the rendered result of an embedded object
                    if let Some(state) = self.states.last_mut() {
                        state.destination = Destination::Result;
                    }
                    self.skip_until_close_brace()?;
                    self.states.pop();
                    return Ok(());
                },
                Token::Control(ControlWord::Field) => {
                    // Parse field group
                    self.parse_field()?;
                    self.skip_until_close_brace()?;
                    self.states.pop();
                    return Ok(());
                },
                Token::Control(ControlWord::Header) => {
                    self.pos += 1;
                    if let Some(state) = self.states.last_mut() {
                        state.destination = Destination::Header;
                    }
                    self.current_hf_type = Some(super::section::HeaderFooterType::Header);
                    self.parse_header_footer_content()?;
                    self.states.pop();
                    return Ok(());
                },
                Token::Control(ControlWord::HeaderFirst) => {
                    self.pos += 1;
                    if let Some(state) = self.states.last_mut() {
                        state.destination = Destination::Header;
                    }
                    self.current_hf_type = Some(super::section::HeaderFooterType::HeaderFirst);
                    self.parse_header_footer_content()?;
                    self.states.pop();
                    return Ok(());
                },
                Token::Control(ControlWord::HeaderLeft) => {
                    self.pos += 1;
                    if let Some(state) = self.states.last_mut() {
                        state.destination = Destination::Header;
                    }
                    self.current_hf_type = Some(super::section::HeaderFooterType::HeaderLeft);
                    self.parse_header_footer_content()?;
                    self.states.pop();
                    return Ok(());
                },
                Token::Control(ControlWord::HeaderRight) => {
                    self.pos += 1;
                    if let Some(state) = self.states.last_mut() {
                        state.destination = Destination::Header;
                    }
                    self.current_hf_type = Some(super::section::HeaderFooterType::HeaderRight);
                    self.parse_header_footer_content()?;
                    self.states.pop();
                    return Ok(());
                },
                Token::Control(ControlWord::Footer) => {
                    self.pos += 1;
                    if let Some(state) = self.states.last_mut() {
                        state.destination = Destination::Footer;
                    }
                    self.current_hf_type = Some(super::section::HeaderFooterType::Footer);
                    self.parse_header_footer_content()?;
                    self.states.pop();
                    return Ok(());
                },
                Token::Control(ControlWord::FooterFirst) => {
                    self.pos += 1;
                    if let Some(state) = self.states.last_mut() {
                        state.destination = Destination::Footer;
                    }
                    self.current_hf_type = Some(super::section::HeaderFooterType::FooterFirst);
                    self.parse_header_footer_content()?;
                    self.states.pop();
                    return Ok(());
                },
                Token::Control(ControlWord::FooterLeft) => {
                    self.pos += 1;
                    if let Some(state) = self.states.last_mut() {
                        state.destination = Destination::Footer;
                    }
                    self.current_hf_type = Some(super::section::HeaderFooterType::FooterLeft);
                    self.parse_header_footer_content()?;
                    self.states.pop();
                    return Ok(());
                },
                Token::Control(ControlWord::FooterRight) => {
                    self.pos += 1;
                    if let Some(state) = self.states.last_mut() {
                        state.destination = Destination::Footer;
                    }
                    self.current_hf_type = Some(super::section::HeaderFooterType::FooterRight);
                    self.parse_header_footer_content()?;
                    self.states.pop();
                    return Ok(());
                },
                Token::Control(ControlWord::Footnote) => {
                    self.pos += 1;
                    if let Some(state) = self.states.last_mut() {
                        state.destination = Destination::Footnote;
                    }
                    self.parse_note(true)?;
                    self.states.pop();
                    return Ok(());
                },
                Token::Control(ControlWord::Endnote) => {
                    self.pos += 1;
                    if let Some(state) = self.states.last_mut() {
                        state.destination = Destination::Endnote;
                    }
                    self.parse_note(false)?;
                    self.states.pop();
                    return Ok(());
                },
                _ => {},
            }
        }

        // Parse group content
        self.parse_content()?;

        // Pop state
        self.states.pop();

        Ok(())
    }

    /// Parse group content (text and control words).
    fn parse_content(&mut self) -> RtfResult<()> {
        let mut text_buffer = SmallVec::<[u8; 256]>::new();

        while self.pos < self.tokens.len() {
            match &self.tokens[self.pos] {
                Token::CloseBrace => {
                    // Flush any buffered text
                    if !text_buffer.is_empty() {
                        self.flush_text_buffer(&mut text_buffer)?;
                    }
                    self.pos += 1;
                    return Ok(());
                },
                Token::OpenBrace => {
                    // Flush text before entering nested group
                    if !text_buffer.is_empty() {
                        self.flush_text_buffer(&mut text_buffer)?;
                    }
                    self.parse_group()?;
                },
                Token::Control(control) => {
                    match control {
                        ControlWord::Par | ControlWord::Line => {
                            self.pos += 1;
                            // Paragraph break - flush current text
                            if !text_buffer.is_empty() {
                                self.flush_text_buffer(&mut text_buffer)?;
                            }
                            text_buffer.push(b'\n');
                        },
                        ControlWord::Tab => {
                            self.pos += 1;
                            text_buffer.push(b'\t');
                        },
                        ControlWord::Unicode(code) => {
                            // Handle Unicode character with potential fallback
                            if !text_buffer.is_empty() {
                                self.flush_text_buffer(&mut text_buffer)?;
                            }
                            self.parse_unicode_sequence(*code)?;
                        },
                        _ => {
                            self.pos += 1;
                            // Apply formatting changes
                            self.apply_control_word(control)?;
                        },
                    }
                },
                Token::Text(text) => {
                    self.pos += 1;
                    // Skip empty text tokens
                    if text.is_empty() {
                        continue;
                    }
                    // Check if we're in a table
                    if self.current_state().map(|s| s.in_table).unwrap_or(false) {
                        // Accumulate in cell text buffer
                        self.current_cell_text.extend_from_slice(text.as_bytes());
                    } else {
                        // Regular text accumulation
                        text_buffer.extend_from_slice(text.as_bytes());
                    }
                },
                Token::Binary(_) => {
                    // Skip binary data for now
                    self.pos += 1;
                },
            }
        }

        // Flush remaining text
        if !text_buffer.is_empty() {
            self.flush_text_buffer(&mut text_buffer)?;
        }

        Ok(())
    }

    /// Flush text buffer to a style block.
    fn flush_text_buffer(&mut self, buffer: &mut SmallVec<[u8; 256]>) -> RtfResult<()> {
        if buffer.is_empty() {
            return Ok(());
        }

        let state = self.current_state()?;

        // Only create blocks for text in the document body
        // Skip text from font tables, color tables, stylesheets, etc.
        if state.destination == Destination::DocumentBody {
            // The bytes in the buffer came from a string that was decoded with Windows-1252.
            // Each character in that string represents a byte value (0x00-0xFF).
            // We need to recover the original bytes, then decode with the correct encoding.
            //
            // Since Windows-1252 characters U+0000-U+00FF map 1:1 to byte values 0x00-0xFF
            // (with some exceptions in the 0x80-0x9F range), we can reconstruct the
            // original bytes by taking the lower 8 bits of each character's code point.
            //
            // Note: buffer contains UTF-8 bytes of the string. We need to decode to chars first.
            let original_bytes: SmallVec<[u8; 256]> = std::str::from_utf8(buffer)
                .unwrap_or("")
                .chars()
                .map(|c| c as u8) // Take lower 8 bits
                .collect();

            // Now decode using the correct encoding
            let (decoded_str, _, _) = state.encoding.decode(&original_bytes);

            // Allocate in arena and create block
            let text = self.arena.alloc_str(&decoded_str);
            let block = StyleBlock::new(Cow::Borrowed(text), state.formatting, state.paragraph);
            self.body_text_len = self.body_text_len.checked_add(text.len()).ok_or_else(|| {
                RtfError::MalformedDocument("RTF body text length overflow".to_string())
            })?;
            self.blocks.push(block);
        }

        buffer.clear();
        Ok(())
    }

    /// Apply a control word to the current state.
    fn apply_control_word(&mut self, control: &ControlWord) -> RtfResult<()> {
        let state = self.current_state_mut()?;

        match control {
            // Font formatting
            ControlWord::FontNumber(n) => {
                state.formatting.font_ref = *n as FontRef;
            },
            ControlWord::FontSize(size) => {
                if let Some(nz) = NonZeroU16::new((*size).max(0) as u16) {
                    state.formatting.font_size = nz;
                }
            },
            ControlWord::ColorForeground(c) => {
                state.formatting.color_ref = *c as ColorRef;
            },

            // Character formatting
            ControlWord::Bold(b) => state.formatting.bold = *b,
            ControlWord::Italic(b) => state.formatting.italic = *b,
            ControlWord::Underline(b) => {
                state.formatting.underline = if *b {
                    super::types::UnderlineStyle::Single
                } else {
                    super::types::UnderlineStyle::None
                }
            },
            ControlWord::UnderlineNone => {
                state.formatting.underline = super::types::UnderlineStyle::None
            },
            ControlWord::UnderlineDouble => {
                state.formatting.underline = super::types::UnderlineStyle::Double
            },
            ControlWord::UnderlineDotted => {
                state.formatting.underline = super::types::UnderlineStyle::Dotted
            },
            ControlWord::UnderlineDashed => {
                state.formatting.underline = super::types::UnderlineStyle::Dashed
            },
            ControlWord::UnderlineDashDot => {
                state.formatting.underline = super::types::UnderlineStyle::DashDot
            },
            ControlWord::UnderlineDashDotDot => {
                state.formatting.underline = super::types::UnderlineStyle::DashDotDot
            },
            ControlWord::UnderlineWords => {
                state.formatting.underline = super::types::UnderlineStyle::Words
            },
            ControlWord::UnderlineThick => {
                state.formatting.underline = super::types::UnderlineStyle::Thick
            },
            ControlWord::UnderlineWave => {
                state.formatting.underline = super::types::UnderlineStyle::Wave
            },
            ControlWord::Strike(b) => state.formatting.strike = *b,
            ControlWord::DoubleStrike(b) => state.formatting.double_strike = *b,
            ControlWord::Superscript(b) => state.formatting.superscript = *b,
            ControlWord::Subscript(b) => state.formatting.subscript = *b,
            ControlWord::SmallCaps(b) => state.formatting.smallcaps = *b,
            ControlWord::AllCaps(b) => state.formatting.all_caps = *b,
            ControlWord::Hidden(b) => state.formatting.hidden = *b,
            ControlWord::Outline(b) => state.formatting.outline = *b,
            ControlWord::Shadow(b) => state.formatting.shadow = *b,
            ControlWord::Emboss(b) => state.formatting.emboss = *b,
            ControlWord::Imprint(b) => state.formatting.imprint = *b,
            ControlWord::CharSpacing(n) => state.formatting.char_spacing = *n,
            ControlWord::CharScale(n) => state.formatting.char_scale = *n,
            ControlWord::Kerning(n) => state.formatting.kerning = *n,
            ControlWord::Highlight(c) => state.formatting.highlight_color = Some(*c as ColorRef),
            ControlWord::Plain => {
                // Reset to default formatting
                state.formatting = Formatting::default();
            },

            // Paragraph alignment
            ControlWord::LeftAlign => state.paragraph.alignment = Alignment::Left,
            ControlWord::RightAlign => state.paragraph.alignment = Alignment::Right,
            ControlWord::Center => state.paragraph.alignment = Alignment::Center,
            ControlWord::Justify => state.paragraph.alignment = Alignment::Justify,
            ControlWord::Pard => {
                // Reset to default paragraph properties
                state.paragraph = Paragraph::default();
            },

            // Paragraph spacing
            ControlWord::SpaceBefore(n) => state.paragraph.spacing.before = *n,
            ControlWord::SpaceAfter(n) => state.paragraph.spacing.after = *n,
            ControlWord::SpaceBetween(n) => state.paragraph.spacing.line = *n,
            ControlWord::LineMultiple(b) => state.paragraph.spacing.line_multiple = *b,

            // Paragraph indentation
            ControlWord::LeftIndent(n) => state.paragraph.indentation.left = *n,
            ControlWord::RightIndent(n) => state.paragraph.indentation.right = *n,
            ControlWord::FirstLineIndent(n) => state.paragraph.indentation.first_line = *n,

            // Paragraph additional properties
            ControlWord::KeepTogether => state.paragraph.keep_together = true,
            ControlWord::KeepNext => state.paragraph.keep_next = true,
            ControlWord::PageBreakBefore => state.paragraph.page_break_before = true,
            ControlWord::WidowControl => state.paragraph.widow_control = true,

            // Unicode
            ControlWord::UnicodeSkip(n) => state.unicode_skip = *n,
            ControlWord::Unicode(code) => {
                // Unicode characters are handled separately during text parsing
                // since they may span multiple tokens with fallback characters
                // The control word itself doesn't add text here
                let _ = code; // Suppress unused warning
            },

            // Character encoding
            ControlWord::AnsiCodePage(cp) => {
                // Set encoding based on Windows code page
                if let Some(encoding) = codepage_to_encoding(*cp as u32) {
                    state.encoding = encoding;
                }
            },

            // Table control words
            ControlWord::InTable => {
                state.in_table = true;
            },
            ControlWord::TableRowDefaults => {
                // Start a new row definition
                state.cell_boundaries.clear();
                self.start_table_if_needed();
            },
            ControlWord::CellX(boundary) => {
                // Cell boundary definition
                state.cell_boundaries.push(*boundary);
            },
            ControlWord::TableCell => {
                // Cell break - finalize current cell
                self.finalize_cell();
            },
            ControlWord::TableRow => {
                // Row break - finalize current row
                self.finalize_row();
            },

            _ => {
                // Ignore unknown or unhandled control words
            },
        }

        Ok(())
    }

    /// Parse font table.
    fn parse_font_table(&mut self) -> RtfResult<()> {
        self.pos += 1; // Skip \fonttbl

        while self.pos < self.tokens.len() {
            match &self.tokens[self.pos] {
                Token::CloseBrace => {
                    return Ok(());
                },
                Token::OpenBrace => {
                    self.parse_font_entry()?;
                },
                _ => {
                    self.pos += 1;
                },
            }
        }

        Ok(())
    }

    /// Parse a single font table entry.
    fn parse_font_entry(&mut self) -> RtfResult<()> {
        self.pos += 1; // Skip {

        let mut font_num = 0;
        let mut font_family = FontFamily::Nil;
        let mut charset = 0;
        let mut name_parts = SmallVec::<[&str; 4]>::new();

        while self.pos < self.tokens.len() {
            match &self.tokens[self.pos] {
                Token::CloseBrace => {
                    self.pos += 1;
                    break;
                },
                Token::OpenBrace => {
                    // Skip nested groups (e.g., {\*\panose ...})
                    self.skip_group()?;
                },
                Token::Control(ControlWord::FontNumber(n)) => {
                    font_num = *n as FontRef;
                    self.pos += 1;
                },
                Token::Control(ControlWord::FontFamily(family)) => {
                    font_family = match *family {
                        "roman" => FontFamily::Roman,
                        "swiss" => FontFamily::Swiss,
                        "modern" => FontFamily::Modern,
                        "script" => FontFamily::Script,
                        "decor" => FontFamily::Decor,
                        "tech" => FontFamily::Tech,
                        _ => FontFamily::Nil,
                    };
                    self.pos += 1;
                },
                Token::Control(ControlWord::FontCharset(cs)) => {
                    charset = *cs as u8;
                    self.pos += 1;
                },
                Token::Text(text) => {
                    // Font name (may contain semicolon at the end)
                    let trimmed = text.trim_end_matches(';').trim();
                    if !trimmed.is_empty() {
                        name_parts.push(trimmed);
                    }
                    self.pos += 1;
                },
                _ => {
                    self.pos += 1;
                },
            }
        }

        // Combine name parts
        let name = if name_parts.is_empty() {
            Cow::Borrowed("")
        } else {
            let combined = name_parts.join(" ");
            let allocated = self.arena.alloc_str(&combined);
            Cow::Borrowed(allocated)
        };

        let font = Font::new(name, font_family, charset);
        self.font_table.borrow_mut().insert(font_num, font);

        Ok(())
    }

    /// Parse color table.
    fn parse_color_table(&mut self) -> RtfResult<()> {
        self.pos += 1; // Skip \colortbl

        let mut current_red = 0;
        let mut current_green = 0;
        let mut current_blue = 0;

        while self.pos < self.tokens.len() {
            match &self.tokens[self.pos] {
                Token::CloseBrace => {
                    // Add final color if any
                    let color = Color::new(current_red, current_green, current_blue);
                    self.color_table.borrow_mut().add(color);
                    return Ok(());
                },
                Token::Control(ControlWord::Red(r)) => {
                    current_red = (*r).clamp(0, 255) as u8;
                    self.pos += 1;
                },
                Token::Control(ControlWord::Green(g)) => {
                    current_green = (*g).clamp(0, 255) as u8;
                    self.pos += 1;
                },
                Token::Control(ControlWord::Blue(b)) => {
                    current_blue = (*b).clamp(0, 255) as u8;
                    self.pos += 1;
                },
                Token::Text(text) if text.trim() == ";" => {
                    // Color separator - add current color
                    let color = Color::new(current_red, current_green, current_blue);
                    self.color_table.borrow_mut().add(color);
                    current_red = 0;
                    current_green = 0;
                    current_blue = 0;
                    self.pos += 1;
                },
                _ => {
                    self.pos += 1;
                },
            }
        }

        Ok(())
    }

    /// Parse the standard RTF `info` destination.
    fn parse_info(&mut self) -> RtfResult<()> {
        self.pos += 1; // `info`
        while self.pos < self.tokens.len() {
            match self.tokens.get(self.pos) {
                Some(Token::OpenBrace) => {
                    self.pos += 1;
                    let control = self.tokens.get(self.pos).cloned();
                    match control {
                        Some(Token::Control(ControlWord::Title)) => {
                            self.parse_info_text(InfoTextField::Title)?;
                        },
                        Some(Token::Control(ControlWord::Subject)) => {
                            self.parse_info_text(InfoTextField::Subject)?;
                        },
                        Some(Token::Control(ControlWord::Author)) => {
                            self.parse_info_text(InfoTextField::Author)?;
                        },
                        Some(Token::Control(ControlWord::Manager)) => {
                            self.parse_info_text(InfoTextField::Manager)?;
                        },
                        Some(Token::Control(ControlWord::Company)) => {
                            self.parse_info_text(InfoTextField::Company)?;
                        },
                        Some(Token::Control(ControlWord::Operator)) => {
                            self.parse_info_text(InfoTextField::Operator)?;
                        },
                        Some(Token::Control(ControlWord::Category)) => {
                            self.parse_info_text(InfoTextField::Category)?;
                        },
                        Some(Token::Control(ControlWord::Keywords)) => {
                            self.parse_info_text(InfoTextField::Keywords)?;
                        },
                        Some(Token::Control(ControlWord::Comment | ControlWord::DocComment)) => {
                            self.parse_info_text(InfoTextField::Comment)?;
                        },
                        Some(Token::Control(ControlWord::CreationTime)) => {
                            self.parse_info_time(InfoTimeField::Creation)?;
                        },
                        Some(Token::Control(ControlWord::RevisionTime)) => {
                            self.parse_info_time(InfoTimeField::Revision)?;
                        },
                        Some(Token::Control(ControlWord::PrintTime)) => {
                            self.parse_info_time(InfoTimeField::Print)?;
                        },
                        Some(Token::Control(ControlWord::BackupTime)) => {
                            self.parse_info_time(InfoTimeField::Backup)?;
                        },
                        _ => self.skip_open_info_group()?,
                    }
                },
                Some(Token::CloseBrace) => {
                    self.pos += 1;
                    return Ok(());
                },
                Some(Token::Control(control)) => {
                    match control {
                        ControlWord::InfoVersion(value) => self.info.version = Some(*value),
                        ControlWord::InfoRevision(value) => self.info.revision = Some(*value),
                        ControlWord::EditingTime(value) => self.info.editing_time = Some(*value),
                        ControlWord::NumberOfPages(value) => self.info.pages = Some(*value),
                        ControlWord::NumberOfWords(value) => self.info.words = Some(*value),
                        ControlWord::NumberOfCharacters(value) => {
                            self.info.characters = Some(*value);
                        },
                        ControlWord::NumberOfCharactersWithSpaces(value) => {
                            self.info.characters_with_spaces = Some(*value);
                        },
                        ControlWord::DocumentId(value) => self.info.id = Some(*value),
                        _ => {},
                    }
                    self.pos += 1;
                },
                Some(_) => self.pos += 1,
                None => break,
            }
        }
        Err(RtfError::UnexpectedEof)
    }

    fn parse_bookmark_destination(&mut self) -> RtfResult<()> {
        self.pos += 1; // ignorable-destination marker
        let is_start = match self.tokens.get(self.pos) {
            Some(Token::Control(ControlWord::BookmarkStart)) => true,
            Some(Token::Control(ControlWord::BookmarkEnd)) => false,
            _ => {
                return Err(RtfError::MalformedDocument(
                    "invalid bookmark destination".into(),
                ));
            },
        };
        self.pos += 1;

        let mut name = String::new();
        let mut first_column = None;
        let mut last_column = None;
        let mut is_public = false;
        let mut depth = 1usize;
        let mut unicode_skip = self.current_state()?.unicode_skip.max(0) as usize;
        let mut fallback_skip = 0usize;
        while self.pos < self.tokens.len() && depth > 0 {
            match self.tokens.get(self.pos) {
                Some(Token::OpenBrace) => depth += 1,
                Some(Token::CloseBrace) => depth -= 1,
                Some(Token::Text(text)) => {
                    let skipped = fallback_skip.min(text.chars().count());
                    fallback_skip -= skipped;
                    name.extend(text.chars().skip(skipped));
                },
                Some(Token::Control(ControlWord::BookmarkFirstColumn(value))) => {
                    first_column = Some(*value);
                },
                Some(Token::Control(ControlWord::BookmarkLastColumn(value))) => {
                    last_column = Some(*value);
                },
                Some(Token::Control(ControlWord::BookmarkPublic)) => is_public = true,
                Some(Token::Control(ControlWord::Unicode(_))) => {
                    let mut utf16 = SmallVec::<[u16; 4]>::new();
                    while let Some(Token::Control(ControlWord::Unicode(code))) =
                        self.tokens.get(self.pos)
                    {
                        utf16.push(*code as u16);
                        self.pos += 1;
                    }
                    name.push_str(&String::from_utf16(&utf16).map_err(|error| {
                        RtfError::InvalidUnicode(format!("invalid Unicode bookmark name: {error}"))
                    })?);
                    fallback_skip = unicode_skip.saturating_mul(utf16.len());
                    continue;
                },
                Some(Token::Control(ControlWord::UnicodeSkip(value))) => {
                    unicode_skip = (*value).max(0) as usize;
                },
                _ => {},
            }
            self.pos += 1;
            if name.len() > MAX_BOOKMARK_NAME_BYTES {
                return Err(RtfError::MalformedDocument(
                    "RTF bookmark name exceeds the safety limit".to_string(),
                ));
            }
        }
        if depth != 0 {
            return Err(RtfError::UnexpectedEof);
        }
        let name = name.trim_end_matches(['\r', '\n']).to_string();
        if name.is_empty() {
            return Ok(());
        }

        if is_start {
            if self.next_bookmark_order >= MAX_BOOKMARKS {
                return Err(RtfError::MalformedDocument(
                    "RTF bookmark count exceeds the safety limit".to_string(),
                ));
            }
            let bookmark = OpenBookmark {
                name: name.clone(),
                position: self.body_text_len,
                first_column,
                last_column,
                is_public,
                order: self.next_bookmark_order,
            };
            self.next_bookmark_order += 1;
            self.open_bookmarks.entry(name).or_default().push(bookmark);
        } else if let Some(open) = self.open_bookmarks.get_mut(&name).and_then(Vec::pop) {
            self.bookmark_spans.push(BookmarkSpan {
                bookmark: open,
                end: self.body_text_len,
            });
        }
        Ok(())
    }

    fn finalize_bookmarks(&mut self) -> RtfResult<()> {
        for bookmarks in self.open_bookmarks.values_mut() {
            for bookmark in bookmarks.drain(..) {
                self.bookmark_spans.push(BookmarkSpan {
                    bookmark,
                    end: self.body_text_len,
                });
            }
        }
        self.bookmark_spans
            .sort_unstable_by_key(|span| span.bookmark.order);
        if self.bookmark_spans.is_empty() {
            return Ok(());
        }

        let mut body = String::with_capacity(self.body_text_len);
        for block in &self.blocks {
            body.push_str(block.text.as_ref());
        }
        for span in self.bookmark_spans.drain(..) {
            let content = body.get(span.bookmark.position..span.end).ok_or_else(|| {
                RtfError::MalformedDocument("bookmark does not align to body text".to_string())
            })?;
            self.bookmarks.add(super::bookmark::Bookmark {
                name: Cow::Owned(span.bookmark.name),
                position: span.bookmark.position,
                content: Cow::Owned(content.to_string()),
                first_column: span.bookmark.first_column,
                last_column: span.bookmark.last_column,
                is_public: span.bookmark.is_public,
            });
        }
        Ok(())
    }

    fn parse_ignorable_text_destination(&mut self) -> RtfResult<String> {
        self.pos += 2; // ignorable marker and destination control word
        let mut value = String::new();
        let mut depth = 1usize;
        let mut unicode_skip = self.current_state()?.unicode_skip.max(0) as usize;
        let mut fallback_skip = 0usize;
        while self.pos < self.tokens.len() && depth > 0 {
            match self.tokens.get(self.pos) {
                Some(Token::OpenBrace) => depth += 1,
                Some(Token::CloseBrace) => depth -= 1,
                Some(Token::Text(text)) => {
                    let skipped = fallback_skip.min(text.chars().count());
                    fallback_skip -= skipped;
                    value.extend(text.chars().skip(skipped));
                },
                Some(Token::Control(ControlWord::Unicode(_))) => {
                    let mut utf16 = SmallVec::<[u16; 4]>::new();
                    while let Some(Token::Control(ControlWord::Unicode(code))) =
                        self.tokens.get(self.pos)
                    {
                        utf16.push(*code as u16);
                        self.pos += 1;
                    }
                    value.push_str(&String::from_utf16(&utf16).map_err(|error| {
                        RtfError::InvalidUnicode(format!(
                            "invalid Unicode annotation metadata: {error}"
                        ))
                    })?);
                    fallback_skip = unicode_skip.saturating_mul(utf16.len());
                    continue;
                },
                Some(Token::Control(ControlWord::UnicodeSkip(count))) => {
                    unicode_skip = (*count).max(0) as usize;
                },
                _ => {},
            }
            self.pos += 1;
            if value.len() > MAX_BOOKMARK_NAME_BYTES {
                return Err(RtfError::MalformedDocument(
                    "RTF annotation destination exceeds the safety limit".to_string(),
                ));
            }
        }
        if depth != 0 {
            return Err(RtfError::UnexpectedEof);
        }
        Ok(value.trim_end_matches(['\r', '\n']).to_string())
    }

    fn parse_annotation_range_marker(&mut self, is_start: bool) -> RtfResult<()> {
        let value = self.parse_ignorable_text_destination()?;
        let Ok(reference) = value.trim().parse::<i32>() else {
            return Ok(());
        };
        if !self.annotation_ranges.contains_key(&reference)
            && self.annotation_ranges.len() >= MAX_ANNOTATIONS
        {
            return Err(RtfError::MalformedDocument(
                "RTF annotation range count exceeds the safety limit".to_string(),
            ));
        }
        if is_start {
            self.annotation_ranges
                .insert(reference, (self.body_text_len, None));
        } else {
            self.annotation_ranges
                .entry(reference)
                .and_modify(|range| range.1 = Some(self.body_text_len))
                .or_insert((self.body_text_len, Some(self.body_text_len)));
        }
        Ok(())
    }

    fn parse_annotation_destination(&mut self) -> RtfResult<()> {
        if self.annotations.len() >= MAX_ANNOTATIONS {
            return Err(RtfError::MalformedDocument(
                "RTF annotation count exceeds the safety limit".to_string(),
            ));
        }
        self.pos += 2; // ignorable marker and annotation destination
        let mut reference = None;
        let mut date = None;
        let mut parent_id = None;
        let mut icon = None;
        let mut time = None;
        let mut text = String::new();
        let mut depth = 1usize;
        let mut unicode_skip = self.current_state()?.unicode_skip.max(0) as usize;
        let mut fallback_skip = 0usize;
        while self.pos < self.tokens.len() && depth > 0 {
            match self.tokens.get(self.pos) {
                Some(Token::OpenBrace) => {
                    let nested =
                        match (self.tokens.get(self.pos + 1), self.tokens.get(self.pos + 2)) {
                            (
                                Some(Token::Control(ControlWord::IgnorableDestination)),
                                Some(Token::Control(control)),
                            ) => Some(*control),
                            _ => None,
                        };
                    match nested {
                        Some(ControlWord::AnnotationReference) => {
                            let value = self.parse_nested_annotation_value()?;
                            reference = value.trim().parse::<i32>().ok();
                        },
                        Some(ControlWord::AnnotationDate) => {
                            date = Some(self.parse_nested_annotation_value()?);
                        },
                        Some(ControlWord::AnnotationParent) => {
                            parent_id = Some(self.parse_nested_annotation_value()?);
                        },
                        Some(ControlWord::AnnotationIcon) => {
                            icon = Some(self.parse_nested_annotation_value()?);
                        },
                        Some(ControlWord::AnnotationTime) => {
                            time = Some(self.parse_nested_annotation_value()?);
                        },
                        _ => {
                            depth += 1;
                            self.pos += 1;
                        },
                    }
                    continue;
                },
                Some(Token::CloseBrace) => depth -= 1,
                Some(Token::Text(value)) => {
                    let skipped = fallback_skip.min(value.chars().count());
                    fallback_skip -= skipped;
                    text.extend(value.chars().skip(skipped));
                },
                Some(Token::Control(ControlWord::Unicode(_))) => {
                    let mut utf16 = SmallVec::<[u16; 4]>::new();
                    while let Some(Token::Control(ControlWord::Unicode(code))) =
                        self.tokens.get(self.pos)
                    {
                        utf16.push(*code as u16);
                        self.pos += 1;
                    }
                    text.push_str(&String::from_utf16(&utf16).map_err(|error| {
                        RtfError::InvalidUnicode(format!(
                            "invalid Unicode annotation text: {error}"
                        ))
                    })?);
                    fallback_skip = unicode_skip.saturating_mul(utf16.len());
                    continue;
                },
                Some(Token::Control(ControlWord::UnicodeSkip(count))) => {
                    unicode_skip = (*count).max(0) as usize;
                },
                Some(Token::Control(ControlWord::Par | ControlWord::Line)) => text.push('\n'),
                Some(Token::Control(ControlWord::Tab)) => text.push('\t'),
                _ => {},
            }
            self.pos += 1;
            if text.len() > MAX_ANNOTATION_TEXT_BYTES {
                return Err(RtfError::MalformedDocument(
                    "RTF annotation text exceeds the safety limit".to_string(),
                ));
            }
        }
        if depth != 0 {
            return Err(RtfError::UnexpectedEof);
        }

        let id = reference.unwrap_or(0);
        let (position, range_end) = self
            .annotation_ranges
            .get(&id)
            .map(|(start, end)| (*start, end.unwrap_or(*start)))
            .unwrap_or((self.body_text_len, self.body_text_len));
        self.annotations.push(super::annotation::Annotation {
            annotation_type: super::annotation::AnnotationType::Comment,
            id,
            author: Cow::Owned(std::mem::take(&mut self.pending_annotation_author)),
            initials: Cow::Owned(std::mem::take(&mut self.pending_annotation_initials)),
            date: date.map(Cow::Owned),
            text: Cow::Owned(text.trim_end_matches(['\r', '\n']).to_string()),
            position,
            range_end,
            parent_id: parent_id.map(Cow::Owned),
            icon: icon.map(Cow::Owned),
            time: time.map(Cow::Owned),
        });
        Ok(())
    }

    fn parse_nested_annotation_value(&mut self) -> RtfResult<String> {
        self.pos += 3; // opening brace, ignorable marker, destination
        let mut value = String::new();
        let mut depth = 1usize;
        while self.pos < self.tokens.len() && depth > 0 {
            match self.tokens.get(self.pos) {
                Some(Token::OpenBrace) => depth += 1,
                Some(Token::CloseBrace) => depth -= 1,
                Some(Token::Text(text)) => value.push_str(text),
                _ => {},
            }
            self.pos += 1;
            if value.len() > MAX_BOOKMARK_NAME_BYTES {
                return Err(RtfError::MalformedDocument(
                    "RTF annotation metadata exceeds the safety limit".to_string(),
                ));
            }
        }
        if depth != 0 {
            return Err(RtfError::UnexpectedEof);
        }
        Ok(value.trim_end_matches(['\r', '\n']).to_string())
    }

    fn parse_info_text(&mut self, field: InfoTextField) -> RtfResult<()> {
        self.pos += 1; // destination control word
        let mut value = String::new();
        let mut depth = 1usize;
        let mut fallback_skip = 0usize;
        while self.pos < self.tokens.len() && depth > 0 {
            match self.tokens.get(self.pos) {
                Some(Token::OpenBrace) => {
                    depth += 1;
                    self.pos += 1;
                },
                Some(Token::CloseBrace) => {
                    depth -= 1;
                    self.pos += 1;
                },
                Some(Token::Text(text)) => {
                    let skipped = fallback_skip.min(text.chars().count());
                    fallback_skip -= skipped;
                    let remainder: String = text.chars().skip(skipped).collect();
                    value.push_str(&remainder);
                    self.pos += 1;
                },
                Some(Token::Control(ControlWord::Unicode(_))) => {
                    let mut utf16 = SmallVec::<[u16; 4]>::new();
                    while let Some(Token::Control(ControlWord::Unicode(code))) =
                        self.tokens.get(self.pos)
                    {
                        utf16.push(*code as u16);
                        self.pos += 1;
                    }
                    value.push_str(&String::from_utf16(&utf16).map_err(|error| {
                        RtfError::InvalidUnicode(format!("Invalid info Unicode: {error}"))
                    })?);
                    fallback_skip =
                        self.current_state()?.unicode_skip.max(0) as usize * utf16.len();
                },
                Some(Token::Control(ControlWord::UnicodeSkip(count))) => {
                    self.current_state_mut()?.unicode_skip = *count;
                    self.pos += 1;
                },
                Some(_) => self.pos += 1,
                None => break,
            }
            if value.len() > MAX_INFO_TEXT_BYTES {
                return Err(RtfError::MalformedDocument(
                    "RTF info text exceeds the metadata safety limit".to_string(),
                ));
            }
        }
        if depth != 0 {
            return Err(RtfError::UnexpectedEof);
        }
        let allocated = self.arena.alloc_str(value.trim_end_matches(['\r', '\n']));
        let value = Some(Cow::Borrowed(&*allocated));
        match field {
            InfoTextField::Title => self.info.title = value,
            InfoTextField::Subject => self.info.subject = value,
            InfoTextField::Author => self.info.author = value,
            InfoTextField::Manager => self.info.manager = value,
            InfoTextField::Company => self.info.company = value,
            InfoTextField::Operator => self.info.operator = value,
            InfoTextField::Category => self.info.category = value,
            InfoTextField::Keywords => self.info.keywords = value,
            InfoTextField::Comment => self.info.comment = value,
        }
        Ok(())
    }

    fn parse_info_time(&mut self, field: InfoTimeField) -> RtfResult<()> {
        self.pos += 1; // time destination
        let mut timestamp = InfoTimestamp::default();
        let mut depth = 1usize;
        while self.pos < self.tokens.len() && depth > 0 {
            match self.tokens.get(self.pos) {
                Some(Token::OpenBrace) => depth += 1,
                Some(Token::CloseBrace) => depth -= 1,
                Some(Token::Control(control)) => match control {
                    ControlWord::Year(value) => timestamp.year = Some(*value),
                    ControlWord::Month(value) => timestamp.month = Some(*value),
                    ControlWord::Day(value) => timestamp.day = Some(*value),
                    ControlWord::Hour(value) => timestamp.hour = Some(*value),
                    ControlWord::Minute(value) => timestamp.minute = Some(*value),
                    ControlWord::Second(value) => timestamp.second = Some(*value),
                    _ => {},
                },
                _ => {},
            }
            self.pos += 1;
        }
        if depth != 0 {
            return Err(RtfError::UnexpectedEof);
        }
        let serialized = format!(
            "{:04}-{:02}-{:02}T{:02}:{:02}:{:02}",
            timestamp.year.unwrap_or(0),
            timestamp.month.unwrap_or(0),
            timestamp.day.unwrap_or(0),
            timestamp.hour.unwrap_or(0),
            timestamp.minute.unwrap_or(0),
            timestamp.second.unwrap_or(0),
        );
        let allocated = self.arena.alloc_str(&serialized);
        let value = Some(Cow::Borrowed(&*allocated));
        match field {
            InfoTimeField::Creation => self.info.creation_time = value,
            InfoTimeField::Revision => self.info.revision_time = value,
            InfoTimeField::Print => self.info.print_time = value,
            InfoTimeField::Backup => self.info.backup_time = value,
        }
        Ok(())
    }

    fn skip_open_info_group(&mut self) -> RtfResult<()> {
        let mut depth = 1usize;
        while self.pos < self.tokens.len() && depth > 0 {
            match self.tokens.get(self.pos) {
                Some(Token::OpenBrace) => depth += 1,
                Some(Token::CloseBrace) => depth -= 1,
                _ => {},
            }
            self.pos += 1;
        }
        (depth == 0).then_some(()).ok_or(RtfError::UnexpectedEof)
    }

    /// Skip tokens until closing brace.
    fn skip_until_close_brace(&mut self) -> RtfResult<()> {
        let mut depth = 1;

        while self.pos < self.tokens.len() && depth > 0 {
            match &self.tokens[self.pos] {
                Token::OpenBrace => depth += 1,
                Token::CloseBrace => depth -= 1,
                _ => {},
            }
            self.pos += 1;
        }

        Ok(())
    }

    /// Skip an entire group starting from the OpenBrace token.
    fn skip_group(&mut self) -> RtfResult<()> {
        // Must be positioned at OpenBrace
        if !matches!(self.tokens.get(self.pos), Some(Token::OpenBrace)) {
            return Ok(());
        }

        self.pos += 1; // Skip the OpenBrace
        let mut depth = 1;

        while self.pos < self.tokens.len() && depth > 0 {
            match &self.tokens[self.pos] {
                Token::OpenBrace => depth += 1,
                Token::CloseBrace => depth -= 1,
                _ => {},
            }
            self.pos += 1;
        }

        Ok(())
    }

    /// Expect a specific token.
    fn expect_token(&mut self, expected: Token) -> RtfResult<()> {
        if self.pos >= self.tokens.len() {
            return Err(RtfError::UnexpectedEof);
        }

        if self.tokens[self.pos] != expected {
            return Err(RtfError::ParserError(format!(
                "Expected {:?}, found {:?}",
                expected, self.tokens[self.pos]
            )));
        }

        self.pos += 1;
        Ok(())
    }

    /// Get current state (mutable).
    fn current_state_mut(&mut self) -> RtfResult<&mut State> {
        self.states
            .last_mut()
            .ok_or_else(|| RtfError::ParserError("No parser state available".to_string()))
    }

    /// Get current state (immutable).
    fn current_state(&self) -> RtfResult<&State> {
        self.states
            .last()
            .ok_or_else(|| RtfError::ParserError("No parser state available".to_string()))
    }

    /// Parse Unicode character sequence with fallback handling.
    ///
    /// RTF Unicode format: `\uN` where N is a signed 16-bit decimal value
    /// Followed by `\ucN` fallback characters (usually ANSI representation)
    ///
    /// Handles compound Unicode characters (surrogate pairs for emoji, etc.)
    fn parse_unicode_sequence(&mut self, first_code: i32) -> RtfResult<()> {
        let skip_count = self.current_state()?.unicode_skip as usize;

        // Collect all consecutive unicode values (for surrogate pairs)
        let mut unicode_values = SmallVec::<[u16; 4]>::new();

        // Convert signed 16-bit value to unsigned
        unicode_values.push(first_code as u16);
        self.pos += 1;

        // Look ahead for additional Unicode characters (compound characters)
        while self.pos < self.tokens.len() {
            if let Token::Control(ControlWord::Unicode(code)) = &self.tokens[self.pos] {
                unicode_values.push(*code as u16);
                self.pos += 1;
            } else {
                break;
            }
        }

        // Skip fallback characters based on unicode_skip count
        // Fallback chars are for non-Unicode readers (usually hex escapes or plain ASCII)
        let mut fallback_skip = skip_count * unicode_values.len();
        let mut fallback_remainder = None;

        // Handle fallback: skip the next N characters/tokens
        while fallback_skip > 0 && self.pos < self.tokens.len() {
            match &self.tokens[self.pos] {
                Token::Text(text) => {
                    let character_count = text.chars().count();
                    if character_count <= fallback_skip {
                        fallback_skip -= character_count;
                        self.pos += 1;
                    } else {
                        fallback_remainder =
                            Some(text.chars().skip(fallback_skip).collect::<String>());
                        fallback_skip = 0;
                        self.pos += 1;
                    }
                },
                Token::Control(ControlWord::Unicode(_)) => {
                    // Next unicode, don't skip
                    break;
                },
                _ => {
                    // Treat other tokens as single character
                    fallback_skip = fallback_skip.saturating_sub(1);
                    self.pos += 1;
                },
            }
        }

        // Convert Unicode values to UTF-8 string
        let unicode_str = String::from_utf16(&unicode_values)
            .map_err(|e| RtfError::InvalidUnicode(format!("Invalid Unicode sequence: {}", e)))?;

        let state = self.current_state()?.clone();
        if state.in_table {
            self.current_cell_text
                .extend_from_slice(unicode_str.as_bytes());
            if let Some(remainder) = fallback_remainder {
                self.current_cell_text
                    .extend_from_slice(remainder.as_bytes());
            }
        } else {
            // Add the Unicode sequence to the document as its own formatted block.
            let allocated = self.arena.alloc_str(&unicode_str);
            let block =
                StyleBlock::new(Cow::Borrowed(allocated), state.formatting, state.paragraph);
            self.body_text_len =
                self.body_text_len
                    .checked_add(allocated.len())
                    .ok_or_else(|| {
                        RtfError::MalformedDocument("RTF body text length overflow".into())
                    })?;
            self.blocks.push(block);

            // A fallback and subsequent text often share one lexer token. Preserve
            // the portion after the configured fallback character count.
            if let Some(remainder) = fallback_remainder {
                let mut buffer = SmallVec::<[u8; 256]>::from_slice(remainder.as_bytes());
                self.flush_text_buffer(&mut buffer)?;
            }
        }

        Ok(())
    }

    /// Start a table if not already started.
    fn start_table_if_needed(&mut self) {
        if self.current_table.is_none() {
            self.current_table = Some(super::table::Table::new());
        }
        if self.current_row.is_none() {
            self.current_row = Some(super::table::Row::new());
        }
    }

    /// Finalize the current cell and add it to the current row.
    fn finalize_cell(&mut self) {
        if !self.current_cell_text.is_empty() {
            // Convert cell text to string
            if let Ok(text_str) = std::str::from_utf8(&self.current_cell_text) {
                let allocated = self.arena.alloc_str(text_str);
                let cell = super::table::Cell::new(Cow::Borrowed(allocated));

                // Add cell to current row
                if let Some(row) = &mut self.current_row {
                    row.add_cell(cell);
                }
            }

            // Clear cell buffer
            self.current_cell_text.clear();
        }
    }

    /// Finalize the current row and add it to the current table.
    fn finalize_row(&mut self) {
        // Finalize any pending cell
        self.finalize_cell();

        // Add row to table
        if let (Some(table), Some(row)) = (&mut self.current_table, self.current_row.take())
            && row.cell_count() > 0
        {
            table.add_row(row);
        }

        // Start a new row for next cells
        self.current_row = Some(super::table::Row::new());
    }

    /// Finalize the current table and add it to the tables list.
    fn finalize_table(&mut self) {
        // Finalize any pending row
        if self.current_row.is_some() {
            self.finalize_row();
        }

        // Add table to tables list
        if let Some(table) = self.current_table.take()
            && table.row_count() > 0
        {
            self.tables.push(table);
        }
    }

    /// Parse picture/image content.
    ///
    /// Pictures in RTF have the format:
    /// {\pict\emfblip\picw<width>\pich<height>...<hex data>}
    fn parse_picture(&mut self) -> RtfResult<()> {
        self.pos += 1; // Skip \pict

        let mut image_type = super::picture::ImageType::Unknown;
        let mut width = None;
        let mut height = None;
        let mut goal_width = None;
        let mut goal_height = None;
        let mut scale_x = None;
        let mut scale_y = None;
        let mut hex_data = SmallVec::<[u8; 512]>::new();

        // Parse picture properties and data
        while self.pos < self.tokens.len() {
            match &self.tokens[self.pos] {
                Token::CloseBrace => {
                    break;
                },
                Token::Control(control) => {
                    self.pos += 1;
                    match control {
                        ControlWord::Emfblip => image_type = super::picture::ImageType::Emf,
                        ControlWord::Pngblip => image_type = super::picture::ImageType::Png,
                        ControlWord::Jpegblip => image_type = super::picture::ImageType::Jpeg,
                        ControlWord::Macpict => image_type = super::picture::ImageType::Pict,
                        ControlWord::Wmetafile(_) | ControlWord::Pmmetafile(_) => {
                            image_type = super::picture::ImageType::Wmf
                        },
                        ControlWord::Dibitmap(_) | ControlWord::Wbitmap(_) => {
                            image_type = super::picture::ImageType::Dib
                        },
                        ControlWord::PictureWidth(w) => width = Some(*w),
                        ControlWord::PictureHeight(h) => height = Some(*h),
                        ControlWord::PictureGoalWidth(w) => goal_width = Some(*w),
                        ControlWord::PictureGoalHeight(h) => goal_height = Some(*h),
                        ControlWord::PictureScaleX(s) => scale_x = Some(*s),
                        ControlWord::PictureScaleY(s) => scale_y = Some(*s),
                        _ => {},
                    }
                },
                Token::Text(text) => {
                    // Accumulate hex-encoded image data
                    hex_data.extend_from_slice(text.as_bytes());
                    self.pos += 1;
                },
                Token::Binary(_) => {
                    // Skip binary data for now
                    self.pos += 1;
                },
                Token::OpenBrace => {
                    // Skip nested groups
                    self.skip_group()?;
                },
            }
        }

        // Decode hex data to binary
        if !hex_data.is_empty()
            && let Ok(hex_str) = std::str::from_utf8(&hex_data)
            && let Ok(decoded) = litchi_core::encoding::decode_hex_data(hex_str)
        {
            // If type not specified, try to detect from data
            if image_type == super::picture::ImageType::Unknown {
                image_type = super::picture::detect_image_type(&decoded);
            }

            // Allocate in arena and create picture
            let data_alloc = self.arena.alloc_slice_copy(&decoded);
            let mut picture = super::picture::Picture::new(image_type, Cow::Borrowed(data_alloc));
            picture.width = width;
            picture.height = height;
            picture.goal_width = goal_width;
            picture.goal_height = goal_height;
            picture.scale_x = scale_x;
            picture.scale_y = scale_y;

            self.pictures.push(picture);
        }

        Ok(())
    }

    /// Parse field content.
    ///
    /// Fields in RTF have the format:
    /// {\field{\*\fldinst INSTRUCTION}{\fldrslt RESULT}}
    fn parse_field(&mut self) -> RtfResult<()> {
        self.pos += 1; // Skip \field

        let mut instruction = SmallVec::<[u8; 128]>::new();
        let mut result = SmallVec::<[u8; 128]>::new();
        let mut in_instruction;
        let mut in_result;

        // Parse field groups
        while self.pos < self.tokens.len() {
            match &self.tokens[self.pos] {
                Token::CloseBrace => {
                    // End of outer field group
                    break;
                },
                Token::OpenBrace => {
                    self.pos += 1;
                    // Check for fldinst or fldrslt
                    if self.pos < self.tokens.len() {
                        // Look for \*\fldinst or \fldrslt
                        let is_ignorable = matches!(
                            self.tokens.get(self.pos),
                            Some(Token::Control(ControlWord::IgnorableDestination))
                        );
                        if is_ignorable {
                            self.pos += 1;
                        }

                        if let Some(Token::Control(ControlWord::FieldInstruction)) =
                            self.tokens.get(self.pos)
                        {
                            self.pos += 1;
                            in_instruction = true;
                            in_result = false;
                            if let Some(state) = self.states.last_mut() {
                                state.destination = Destination::FieldInstruction;
                            }
                        } else if let Some(Token::Control(ControlWord::FieldResult)) =
                            self.tokens.get(self.pos)
                        {
                            self.pos += 1;
                            in_instruction = false;
                            in_result = true;
                            if let Some(state) = self.states.last_mut() {
                                state.destination = Destination::FieldResult;
                            }
                        } else {
                            // Skip unknown nested groups
                            self.skip_until_close_brace()?;
                            continue;
                        }

                        // Collect text until closing brace
                        while self.pos < self.tokens.len() {
                            match &self.tokens[self.pos] {
                                Token::CloseBrace => {
                                    self.pos += 1;
                                    break;
                                },
                                Token::Text(text) => {
                                    if in_instruction {
                                        instruction.extend_from_slice(text.as_bytes());
                                    } else if in_result {
                                        result.extend_from_slice(text.as_bytes());
                                    }
                                    self.pos += 1;
                                },
                                Token::OpenBrace => {
                                    // Skip nested groups
                                    self.skip_group()?;
                                },
                                _ => {
                                    self.pos += 1;
                                },
                            }
                        }
                    }
                },
                _ => {
                    self.pos += 1;
                },
            }
        }

        // Create field if we have instruction
        if !instruction.is_empty()
            && let Ok(inst_str) = std::str::from_utf8(&instruction)
        {
            // Allocate instruction in arena first
            let inst_alloc = self.arena.alloc_str(inst_str);

            // Parse field type from allocated instruction
            let mut field = super::field::Field::parse_instruction(inst_alloc);
            field.instruction = Cow::Borrowed(inst_alloc);

            // Add result if available
            if !result.is_empty()
                && let Ok(res_str) = std::str::from_utf8(&result)
            {
                let res_alloc = self.arena.alloc_str(res_str);
                field.result = Cow::Borrowed(res_alloc);
            }

            self.fields.push(field);
        }

        Ok(())
    }

    /// Parse header or footer content.
    fn parse_header_footer_content(&mut self) -> RtfResult<()> {
        let hf_type = self
            .current_hf_type
            .ok_or_else(|| RtfError::MalformedDocument("Header/footer type not set".to_string()))?;

        let mut hf = super::section::HeaderFooter::new(hf_type);
        let mut text_buffer = SmallVec::<[u8; 256]>::new();
        let default_state = State::default();

        while self.pos < self.tokens.len() {
            match &self.tokens[self.pos] {
                Token::CloseBrace => {
                    if !text_buffer.is_empty() {
                        if let Ok(text) = std::str::from_utf8(&text_buffer) {
                            let state = self.current_state().ok().unwrap_or(&default_state);
                            let text_alloc = self.arena.alloc_str(text);
                            let para = super::section::HeaderFooterParagraph::new(
                                Cow::Borrowed(text_alloc),
                                state.formatting,
                                state.paragraph,
                            );
                            hf.add_paragraph(para);
                        }
                        text_buffer.clear();
                    }
                    self.pos += 1;
                    break;
                },
                Token::OpenBrace => {
                    if !text_buffer.is_empty() {
                        if let Ok(text) = std::str::from_utf8(&text_buffer) {
                            let state = self.current_state().ok().unwrap_or(&default_state);
                            let text_alloc = self.arena.alloc_str(text);
                            let para = super::section::HeaderFooterParagraph::new(
                                Cow::Borrowed(text_alloc),
                                state.formatting,
                                state.paragraph,
                            );
                            hf.add_paragraph(para);
                        }
                        text_buffer.clear();
                    }
                    self.parse_group()?;
                },
                Token::Control(ControlWord::Par | ControlWord::Line) => {
                    self.pos += 1;
                    if !text_buffer.is_empty() {
                        if let Ok(text) = std::str::from_utf8(&text_buffer) {
                            let state = self.current_state().ok().unwrap_or(&default_state);
                            let text_alloc = self.arena.alloc_str(text);
                            let para = super::section::HeaderFooterParagraph::new(
                                Cow::Borrowed(text_alloc),
                                state.formatting,
                                state.paragraph,
                            );
                            hf.add_paragraph(para);
                        }
                        text_buffer.clear();
                    }
                },
                Token::Control(ControlWord::Tab) => {
                    self.pos += 1;
                    text_buffer.push(b'\t');
                },
                Token::Control(ControlWord::Unicode(code)) => {
                    let decoded = self.parse_destination_unicode_sequence(*code)?;
                    text_buffer.extend_from_slice(decoded.as_bytes());
                },
                Token::Control(control) => {
                    self.pos += 1;
                    self.apply_control_word(control)?;
                },
                Token::Text(text) => {
                    self.pos += 1;
                    text_buffer.extend_from_slice(text.as_bytes());
                },
                _ => {
                    self.pos += 1;
                },
            }
        }

        // Add header/footer to the current section or create a new section
        if let Some(section) = self.sections.last_mut() {
            section.add_header_footer(hf);
        } else {
            let mut section = super::section::Section::new();
            section.add_header_footer(hf);
            self.sections.push(section);
        }

        self.current_hf_type = None;
        Ok(())
    }

    fn parse_destination_unicode_sequence(&mut self, first_code: i32) -> RtfResult<String> {
        let skip_count = self.current_state()?.unicode_skip.max(0) as usize;
        let mut utf16 = SmallVec::<[u16; 4]>::new();
        utf16.push(first_code as u16);
        self.pos += 1;
        while let Some(Token::Control(ControlWord::Unicode(code))) = self.tokens.get(self.pos) {
            utf16.push(*code as u16);
            self.pos += 1;
        }

        let mut fallback_skip = skip_count.saturating_mul(utf16.len());
        let mut remainder = String::new();
        while fallback_skip > 0 && self.pos < self.tokens.len() {
            match self.tokens.get(self.pos) {
                Some(Token::Text(text)) => {
                    let count = text.chars().count();
                    if count <= fallback_skip {
                        fallback_skip -= count;
                    } else {
                        remainder.extend(text.chars().skip(fallback_skip));
                        fallback_skip = 0;
                    }
                    self.pos += 1;
                },
                Some(Token::Control(ControlWord::Unicode(_))) => break,
                Some(_) => {
                    fallback_skip = fallback_skip.saturating_sub(1);
                    self.pos += 1;
                },
                None => break,
            }
        }
        let mut decoded = String::from_utf16(&utf16).map_err(|error| {
            RtfError::InvalidUnicode(format!("invalid destination Unicode: {error}"))
        })?;
        decoded.push_str(&remainder);
        Ok(decoded)
    }

    /// Parse footnote or endnote content.
    fn parse_note(&mut self, is_footnote: bool) -> RtfResult<()> {
        self.current_note_buffer.clear();
        let mut reference = String::from(if is_footnote { "1" } else { "i" });

        while self.pos < self.tokens.len() {
            match &self.tokens[self.pos] {
                Token::CloseBrace => {
                    self.pos += 1;
                    break;
                },
                Token::OpenBrace => {
                    self.parse_group()?;
                },
                Token::Control(ControlWord::FootnoteNumber(n)) => {
                    self.pos += 1;
                    reference = n.to_string();
                },
                Token::Control(ControlWord::Tab) => {
                    self.pos += 1;
                    self.current_note_buffer.push(b'\t');
                },
                Token::Control(control) => {
                    self.pos += 1;
                    self.apply_control_word(control)?;
                },
                Token::Text(text) => {
                    self.pos += 1;
                    self.current_note_buffer.extend_from_slice(text.as_bytes());
                },
                _ => {
                    self.pos += 1;
                },
            }
        }

        if !self.current_note_buffer.is_empty()
            && let Ok(content) = std::str::from_utf8(&self.current_note_buffer)
        {
            let content_alloc = self.arena.alloc_str(content);
            let mut note = if is_footnote {
                super::section::Note::footnote(Cow::Owned(reference), Cow::Borrowed(content_alloc))
            } else {
                super::section::Note::endnote(Cow::Owned(reference), Cow::Borrowed(content_alloc))
            };

            if let Ok(state) = self.current_state() {
                note.formatting = state.formatting;
            }

            self.notes.push(note);
        }

        Ok(())
    }
}

/// Parsed RTF document.
///
/// This is an intermediate representation produced by the parser
/// before being converted into the final `RtfDocument` structure.
/// All fields are public to allow direct access during document construction.
pub struct ParsedDocument<'a> {
    /// Font table
    pub font_table: FontTable<'a>,
    /// Color table
    pub color_table: ColorTable,
    /// Style blocks
    pub blocks: Vec<StyleBlock<'a>>,
    /// Extracted tables
    pub tables: Vec<super::table::Table<'a>>,
    /// Extracted pictures
    pub pictures: Vec<super::picture::Picture<'a>>,
    /// Extracted fields
    pub fields: Vec<super::field::Field<'a>>,
    /// List table
    pub list_table: super::list::ListTable<'a>,
    /// List override table
    pub list_override_table: super::list::ListOverrideTable,
    /// Sections
    pub sections: Vec<super::section::Section<'a>>,
    /// Bookmarks
    pub bookmarks: super::bookmark::BookmarkTable<'a>,
    /// Shapes
    pub shapes: Vec<super::shape::Shape<'a>>,
    /// Shape groups
    pub shape_groups: Vec<super::shape::ShapeGroup<'a>>,
    /// Stylesheet
    pub stylesheet: super::stylesheet::StyleSheet<'a>,
    /// Document information
    pub info: super::info::DocumentInfo<'a>,
    /// Annotations
    pub annotations: Vec<super::annotation::Annotation<'a>>,
    /// Footnotes and endnotes
    pub notes: Vec<super::section::Note<'a>>,
    /// Track changes/revisions
    pub revisions: Vec<super::annotation::Revision<'a>>,
}
