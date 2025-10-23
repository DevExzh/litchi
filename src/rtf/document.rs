//! RTF document representation.

use super::error::{RtfError, RtfResult};
use super::lexer::Lexer;
use super::parser::Parser;
use super::types::{ColorTable, FontTable, Paragraph as RtfParagraph, Run, StyleBlock};
use bumpalo::Bump;
use std::borrow::Cow;
use std::path::Path;

/// RTF Document.
///
/// This is the main entry point for parsing RTF documents.
/// It provides access to the document's text content, paragraphs, runs, and tables.
pub struct RtfDocument<'a> {
    /// Font table
    font_table: FontTable<'a>,
    /// Color table
    color_table: ColorTable,
    /// Style blocks
    blocks: Vec<StyleBlock<'a>>,
    /// Extracted tables
    tables: Vec<super::table::Table<'a>>,
    /// Arena allocator (kept to maintain lifetime)
    _arena: Bump,
}

impl<'a> RtfDocument<'a> {
    /// Parse an RTF document from a string.
    ///
    /// This method automatically detects and decompresses compressed RTF data.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi::rtf::RtfDocument;
    ///
    /// let rtf = r#"{\rtf1\ansi Hello World!\par}"#;
    /// let doc = RtfDocument::parse(rtf)?;
    /// let text = doc.text();
    /// # Ok::<(), litchi::rtf::RtfError>(())
    /// ```
    pub fn parse(input: &str) -> RtfResult<RtfDocument<'static>> {
        Self::parse_internal(input.as_bytes())
    }

    /// Parse RTF from bytes (handles both compressed and uncompressed)
    fn parse_internal(bytes: &[u8]) -> RtfResult<RtfDocument<'static>> {
        // Check if it's compressed RTF
        let input = if super::compressed::is_compressed_rtf(bytes) {
            // Decompress first
            let decompressed = super::compressed::decompress(bytes)?;
            String::from_utf8(decompressed).map_err(|e| {
                RtfError::InvalidUnicode(format!("Invalid UTF-8 after decompression: {}", e))
            })?
        } else {
            // Try to parse as UTF-8 string
            std::str::from_utf8(bytes)
                .map_err(|e| RtfError::InvalidUnicode(format!("Invalid UTF-8: {}", e)))?
                .to_string()
        };

        Self::parse_string(&input)
    }

    /// Parse an RTF document from a UTF-8 string (internal)
    fn parse_string(input: &str) -> RtfResult<RtfDocument<'static>> {
        // Create arena for temporary allocations during parsing
        let arena = Bump::new();

        // Lexer phase
        let mut lexer = Lexer::new(input, &arena);
        let tokens = lexer.tokenize()?;

        // Parser phase
        let parser = Parser::new(&tokens, &arena);
        let parsed = parser.parse()?;

        // Convert parsed document to owned document
        // We need to convert Cow::Borrowed to Cow::Owned to detach from input lifetime
        let owned_blocks: Vec<StyleBlock<'static>> = parsed
            .blocks
            .into_iter()
            .map(|block| StyleBlock {
                text: Cow::Owned(block.text.into_owned()),
                formatting: block.formatting,
                paragraph: block.paragraph,
            })
            .collect();

        // Convert font table to owned
        let owned_font_table = FontTable {
            fonts: parsed
                .font_table
                .fonts
                .into_iter()
                .map(|font| super::types::Font {
                    name: Cow::Owned(font.name.into_owned()),
                    family: font.family,
                    charset: font.charset,
                })
                .collect(),
        };

        // Convert tables to owned
        let owned_tables: Vec<super::table::Table<'static>> = parsed
            .tables
            .into_iter()
            .map(|table| {
                let mut owned_table = super::table::Table::new();
                for row in table.rows() {
                    let mut owned_row = super::table::Row::new();
                    for cell in row.cells() {
                        let owned_cell =
                            super::table::Cell::new(Cow::Owned(cell.text().to_string()));
                        owned_row.add_cell(owned_cell);
                    }
                    owned_table.add_row(owned_row);
                }
                owned_table
            })
            .collect();

        Ok(RtfDocument {
            font_table: owned_font_table,
            color_table: parsed.color_table,
            blocks: owned_blocks,
            tables: owned_tables,
            _arena: arena,
        })
    }

    /// Parse an RTF document from a file.
    ///
    /// This method automatically detects and handles compressed RTF files.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi::rtf::RtfDocument;
    ///
    /// let doc = RtfDocument::open("document.rtf")?;
    /// let text = doc.text();
    /// # Ok::<(), litchi::rtf::RtfError>(())
    /// ```
    pub fn open<P: AsRef<Path>>(path: P) -> RtfResult<RtfDocument<'static>> {
        let bytes = std::fs::read(path)
            .map_err(|e| RtfError::ParserError(format!("Failed to read file: {}", e)))?;
        Self::parse_internal(&bytes)
    }

    /// Parse an RTF document from bytes.
    ///
    /// This method automatically detects and decompresses compressed RTF data.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi::rtf::RtfDocument;
    ///
    /// let bytes = std::fs::read("document.rtf")?;
    /// let doc = RtfDocument::from_bytes(&bytes)?;
    /// let text = doc.text();
    /// # Ok::<(), litchi::rtf::RtfError>(())
    /// ```
    pub fn from_bytes(bytes: &[u8]) -> RtfResult<RtfDocument<'static>> {
        Self::parse_internal(bytes)
    }

    /// Get all text content from the document.
    ///
    /// This concatenates all text blocks with their natural separators.
    pub fn text(&self) -> String {
        self.blocks
            .iter()
            .map(|block| block.text.as_ref())
            .collect::<Vec<&str>>()
            .join("")
    }

    /// Get the number of paragraphs in the document.
    ///
    /// Paragraphs are determined by paragraph breaks in the RTF source.
    pub fn paragraph_count(&self) -> usize {
        self.paragraphs().len()
    }

    /// Get all paragraphs in the document.
    ///
    /// This groups style blocks into paragraphs based on newline characters.
    pub fn paragraphs(&self) -> Vec<RtfParagraph> {
        let mut paragraphs = Vec::new();
        let mut current_para = RtfParagraph::default();
        let mut has_content = false;

        for block in &self.blocks {
            let text = block.text.as_ref();

            // Split on newlines to detect paragraph boundaries
            let parts: Vec<&str> = text.split('\n').collect();

            for (i, part) in parts.iter().enumerate() {
                if !part.is_empty() {
                    // Inherit paragraph properties from the style block
                    current_para = block.paragraph;
                    has_content = true;
                }

                // If this is not the last part, we have a paragraph break
                if i < parts.len() - 1 && has_content {
                    paragraphs.push(current_para);
                    current_para = RtfParagraph::default();
                    has_content = false;
                }
            }
        }

        // Add final paragraph if it has content
        if has_content {
            paragraphs.push(current_para);
        }

        paragraphs
    }

    /// Get all runs in the document.
    ///
    /// A run is a contiguous block of text with the same formatting.
    pub fn runs(&self) -> Vec<Run<'_>> {
        self.blocks
            .iter()
            .map(|block| Run::new(block.text.clone(), block.formatting))
            .collect()
    }

    /// Get all tables in the document.
    ///
    /// Returns all tables extracted from the RTF document.
    pub fn tables(&self) -> &[super::table::Table<'_>] {
        &self.tables
    }

    /// Get the font table.
    pub fn font_table(&self) -> &FontTable<'_> {
        &self.font_table
    }

    /// Get the color table.
    pub fn color_table(&self) -> &ColorTable {
        &self.color_table
    }

    /// Get all style blocks.
    pub fn blocks(&self) -> &[StyleBlock<'_>] {
        &self.blocks
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_simple_document() {
        let rtf = r#"{\rtf1\ansi Hello World!\par}"#;
        let doc = RtfDocument::parse(rtf).unwrap();
        let text = doc.text();
        assert!(text.contains("Hello World"));
    }

    #[test]
    fn test_formatted_text() {
        let rtf = r#"{\rtf1\ansi{\b Bold}{\i Italic}\par}"#;
        let doc = RtfDocument::parse(rtf).unwrap();
        let runs = doc.runs();
        assert!(!runs.is_empty());
    }
}
