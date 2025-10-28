//! RTF parser that builds document structure from tokens.

use super::error::{RtfError, RtfResult};
use super::lexer::{ControlWord, Token};
use super::types::*;
use crate::common::encoding::codepage_to_encoding;
use bumpalo::Bump;
use encoding_rs::Encoding;
use smallvec::SmallVec;
use std::borrow::Cow;
use std::cell::RefCell;
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
    /// Picture data - should be skipped
    /// TODO: Future enhancement - extract and process embedded images
    Picture,
    /// Embedded object - should be skipped  
    /// TODO: Future enhancement - extract OLE objects (e.g., MathType equations)
    Object,
    /// Result of embedded object rendering - should be skipped
    Result,
    /// Other destinations - should be skipped
    Other,
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

        Ok(ParsedDocument {
            font_table: self.font_table.into_inner(),
            color_table: self.color_table.into_inner(),
            blocks: self.blocks,
            tables: self.tables,
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
                    // Mark as info destination and skip
                    if let Some(state) = self.states.last_mut() {
                        state.destination = Destination::Info;
                    }
                    self.skip_until_close_brace()?;
                    self.states.pop();
                    return Ok(());
                },
                Token::Control(ControlWord::Picture) => {
                    // Mark as picture destination and skip
                    // TODO: Future enhancement - extract and process embedded images
                    if let Some(state) = self.states.last_mut() {
                        state.destination = Destination::Picture;
                    }
                    self.skip_until_close_brace()?;
                    self.states.pop();
                    return Ok(());
                },
                Token::Control(ControlWord::Object) => {
                    // Mark as object destination and skip
                    // TODO: Future enhancement - extract OLE objects
                    // Embedded objects in RTF files include:
                    // - MathType/Equation Editor equations
                    // - Excel charts and spreadsheets
                    // - Visio diagrams
                    // - Other OLE-embedded content
                    //
                    // To properly handle these, we would need to:
                    // 1. Parse the OLE object structure from the hex-encoded binary data
                    // 2. Identify the object type (CLSID/ProgID)
                    // 3. Extract and decode the object's native format
                    // 4. Convert to a suitable representation (e.g., LaTeX for equations, PNG for charts)
                    //
                    // For now, we skip these objects to avoid polluting the text output with
                    // hex-encoded binary data.
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
            ControlWord::Underline(b) => state.formatting.underline = *b,
            ControlWord::UnderlineNone => state.formatting.underline = false,
            ControlWord::Strike(b) => state.formatting.strike = *b,
            ControlWord::Superscript(b) => state.formatting.superscript = *b,
            ControlWord::Subscript(b) => state.formatting.subscript = *b,
            ControlWord::SmallCaps(b) => state.formatting.smallcaps = *b,
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

        // Handle fallback: skip the next N characters/tokens
        while fallback_skip > 0 && self.pos < self.tokens.len() {
            match &self.tokens[self.pos] {
                Token::Text(text) => {
                    let text_len = text.len();
                    if text_len <= fallback_skip {
                        fallback_skip -= text_len;
                        self.pos += 1;
                    } else {
                        // Partial text consumption - not ideal but handle it
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

        // Add to document
        let allocated = self.arena.alloc_str(&unicode_str);
        let state = self.current_state()?;
        let block = StyleBlock::new(Cow::Borrowed(allocated), state.formatting, state.paragraph);
        self.blocks.push(block);

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
}

/// Parsed RTF document.
pub struct ParsedDocument<'a> {
    /// Font table
    pub font_table: FontTable<'a>,
    /// Color table
    pub color_table: ColorTable,
    /// Style blocks
    pub blocks: Vec<StyleBlock<'a>>,
    /// Extracted tables
    pub tables: Vec<super::table::Table<'a>>,
}
