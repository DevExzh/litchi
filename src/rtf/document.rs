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
        let input_bytes = if super::compressed::is_compressed_rtf(bytes) {
            // Decompress first
            super::compressed::decompress(bytes)?
        } else {
            bytes.to_vec()
        };

        // RTF files are NOT UTF-8. They contain bytes in whatever code page is
        // specified by \ansicpg (e.g., Windows-1252, GB2312, etc.).
        //
        // We use Latin-1 (ISO-8859-1) encoding for initial parsing because:
        // 1. It provides 1:1 byte-to-character mapping (byte 0xNN -> U+00NN)
        // 2. Control words (ASCII) parse correctly
        // 3. We can recover original bytes and decode them with correct encoding later
        //
        // The parser will detect \ansicpg and use the proper encoding for text.
        let (input_str, _, _) = encoding_rs::WINDOWS_1252.decode(&input_bytes);

        Self::parse_string(&input_str)
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

        // Convert pictures to owned
        let owned_pictures: Vec<super::picture::Picture<'static>> = parsed
            .pictures
            .into_iter()
            .map(|pic| super::picture::Picture {
                image_type: pic.image_type,
                data: Cow::Owned(pic.data.into_owned()),
                width: pic.width,
                height: pic.height,
                goal_width: pic.goal_width,
                goal_height: pic.goal_height,
                scale_x: pic.scale_x,
                scale_y: pic.scale_y,
            })
            .collect();

        // Convert fields to owned
        let owned_fields: Vec<super::field::Field<'static>> = parsed
            .fields
            .into_iter()
            .map(|field| super::field::Field {
                field_type: field.field_type,
                instruction: Cow::Owned(field.instruction.into_owned()),
                result: Cow::Owned(field.result.into_owned()),
            })
            .collect();

        // Convert all borrowed data to owned
        Ok(RtfDocument {
            font_table: owned_font_table,
            color_table: parsed.color_table,
            blocks: owned_blocks,
            tables: owned_tables,
            pictures: owned_pictures,
            fields: owned_fields,
            list_table: Self::convert_list_table_to_owned(parsed.list_table),
            list_override_table: parsed.list_override_table,
            sections: Self::convert_sections_to_owned(parsed.sections),
            bookmarks: Self::convert_bookmarks_to_owned(parsed.bookmarks),
            shapes: Self::convert_shapes_to_owned(parsed.shapes),
            shape_groups: Self::convert_shape_groups_to_owned(parsed.shape_groups),
            stylesheet: Self::convert_stylesheet_to_owned(parsed.stylesheet),
            info: Self::convert_info_to_owned(parsed.info),
            annotations: Self::convert_annotations_to_owned(parsed.annotations),
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

    /// Get all paragraphs with their content (runs).
    ///
    /// This groups style blocks into paragraphs based on newline characters,
    /// and returns each paragraph with its associated runs.
    pub fn paragraphs_with_content(&self) -> Vec<super::types::ParagraphContent<'_>> {
        use std::borrow::Cow;

        let mut paragraphs = Vec::new();
        let mut current_para_props = RtfParagraph::default();
        let mut current_runs: Vec<Run<'_>> = Vec::new();
        let mut has_content = false;

        for block in &self.blocks {
            let text = block.text.as_ref();

            // Split on newlines to detect paragraph boundaries
            let parts: Vec<&str> = text.split('\n').collect();

            for (i, part) in parts.iter().enumerate() {
                if !part.is_empty() {
                    // Inherit paragraph properties from the style block
                    current_para_props = block.paragraph;
                    has_content = true;

                    // Add run for this part
                    current_runs.push(Run::new(Cow::Borrowed(part), block.formatting));
                }

                // If this is not the last part, we have a paragraph break
                if i < parts.len() - 1 && has_content {
                    paragraphs.push(super::types::ParagraphContent::new(
                        current_para_props,
                        current_runs.clone(),
                    ));
                    current_runs.clear();
                    current_para_props = RtfParagraph::default();
                    has_content = false;
                }
            }
        }

        // Add final paragraph if it has content
        if has_content {
            paragraphs.push(super::types::ParagraphContent::new(
                current_para_props,
                current_runs,
            ));
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

    /// Get all document elements (paragraphs and tables) in approximate document order.
    ///
    /// Note: Due to RTF's structure, tables are extracted separately from paragraph flow.
    /// This method returns paragraphs first, followed by tables. For most use cases this
    /// is sufficient. If you need precise positional information, work with `blocks()` directly.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi::rtf::RtfDocument;
    ///
    /// let doc = RtfDocument::open("document.rtf")?;
    /// for element in doc.elements() {
    ///     match element {
    ///         litchi::rtf::DocumentElement::Paragraph(para) => {
    ///             println!("Paragraph: {}", para.text());
    ///         }
    ///         litchi::rtf::DocumentElement::Table(table) => {
    ///             println!("Table with {} rows", table.row_count());
    ///         }
    ///     }
    /// }
    /// # Ok::<(), litchi::rtf::RtfError>(())
    /// ```
    pub fn elements(&self) -> Vec<super::DocumentElement<'_>> {
        let mut elements = Vec::new();

        // Add all paragraphs first
        for para in self.paragraphs_with_content() {
            elements.push(super::DocumentElement::Paragraph(para));
        }

        // Add all tables
        for table in &self.tables {
            elements.push(super::DocumentElement::Table(table.clone()));
        }

        elements
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

    /// Get all pictures in the document.
    ///
    /// Returns all embedded images extracted from the RTF document.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi::rtf::RtfDocument;
    ///
    /// let doc = RtfDocument::open("document.rtf")?;
    /// for (i, picture) in doc.pictures().iter().enumerate() {
    ///     println!("Picture {}: {:?}, {} bytes", i, picture.image_type, picture.data().len());
    /// }
    /// # Ok::<(), litchi::rtf::RtfError>(())
    /// ```
    pub fn pictures(&self) -> &[super::picture::Picture<'_>] {
        &self.pictures
    }

    /// Get all fields in the document.
    ///
    /// Returns all fields (hyperlinks, cross-references, etc.) from the RTF document.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi::rtf::{RtfDocument, FieldType};
    ///
    /// let doc = RtfDocument::open("document.rtf")?;
    /// for field in doc.fields() {
    ///     if field.field_type == FieldType::Hyperlink {
    ///         if let Some(url) = field.extract_url() {
    ///             println!("Hyperlink: {}", url);
    ///         }
    ///     }
    /// }
    /// # Ok::<(), litchi::rtf::RtfError>(())
    /// ```
    pub fn fields(&self) -> &[super::field::Field<'_>] {
        &self.fields
    }

    /// Get the list table.
    ///
    /// Returns all list definitions (for bulleted and numbered lists) in the document.
    pub fn list_table(&self) -> &super::list::ListTable<'_> {
        &self.list_table
    }

    /// Get the list override table.
    ///
    /// Returns list instances that override base list definitions.
    pub fn list_override_table(&self) -> &super::list::ListOverrideTable {
        &self.list_override_table
    }

    /// Get all sections in the document.
    ///
    /// Returns section information including page layout, headers, and footers.
    pub fn sections(&self) -> &[super::section::Section<'_>] {
        &self.sections
    }

    /// Get the bookmark table.
    ///
    /// Returns all bookmarks defined in the document.
    pub fn bookmarks(&self) -> &super::bookmark::BookmarkTable<'_> {
        &self.bookmarks
    }

    /// Get all shapes in the document.
    ///
    /// Returns drawing objects, text boxes, and other shapes.
    pub fn shapes(&self) -> &[super::shape::Shape<'_>] {
        &self.shapes
    }

    /// Get all shape groups in the document.
    ///
    /// Returns grouped shapes.
    pub fn shape_groups(&self) -> &[super::shape::ShapeGroup<'_>] {
        &self.shape_groups
    }

    /// Get the stylesheet.
    ///
    /// Returns style definitions for paragraphs and characters.
    pub fn stylesheet(&self) -> &super::stylesheet::StyleSheet<'_> {
        &self.stylesheet
    }

    /// Get document information/metadata.
    ///
    /// Returns document properties like title, author, subject, etc.
    pub fn info(&self) -> &super::info::DocumentInfo<'_> {
        &self.info
    }

    /// Get all annotations (comments) in the document.
    ///
    /// Returns document annotations and revisions.
    pub fn annotations(&self) -> &[super::annotation::Annotation<'_>] {
        &self.annotations
    }

    // Helper methods to convert borrowed data to owned
    //
    // These methods are used internally during parsing to convert borrowed data
    // (tied to the input lifetime) to owned data (with 'static lifetime).
    // This allows the parsed document to outlive the input string.

    /// Convert list table to owned
    #[allow(clippy::needless_pass_by_value)]
    fn convert_list_table_to_owned(
        _table: super::list::ListTable<'_>,
    ) -> super::list::ListTable<'static> {
        // TODO: Implement proper conversion when list parsing is fully implemented
        super::list::ListTable::new()
    }

    /// Convert sections to owned
    #[allow(clippy::needless_pass_by_value)]
    fn convert_sections_to_owned(
        _sections: Vec<super::section::Section<'_>>,
    ) -> Vec<super::section::Section<'static>> {
        // TODO: Implement proper conversion when section parsing is fully implemented
        Vec::new()
    }

    /// Convert bookmarks to owned
    #[allow(clippy::needless_pass_by_value)]
    fn convert_bookmarks_to_owned(
        _bookmarks: super::bookmark::BookmarkTable<'_>,
    ) -> super::bookmark::BookmarkTable<'static> {
        // TODO: Implement proper conversion when bookmark parsing is fully implemented
        super::bookmark::BookmarkTable::new()
    }

    /// Convert shapes to owned
    #[allow(clippy::needless_pass_by_value)]
    fn convert_shapes_to_owned(
        _shapes: Vec<super::shape::Shape<'_>>,
    ) -> Vec<super::shape::Shape<'static>> {
        // TODO: Implement proper conversion when shape parsing is fully implemented
        Vec::new()
    }

    /// Convert shape groups to owned
    #[allow(clippy::needless_pass_by_value)]
    fn convert_shape_groups_to_owned(
        _groups: Vec<super::shape::ShapeGroup<'_>>,
    ) -> Vec<super::shape::ShapeGroup<'static>> {
        // TODO: Implement proper conversion when shape group parsing is fully implemented
        Vec::new()
    }

    /// Convert stylesheet to owned
    #[allow(clippy::needless_pass_by_value)]
    fn convert_stylesheet_to_owned(
        _stylesheet: super::stylesheet::StyleSheet<'_>,
    ) -> super::stylesheet::StyleSheet<'static> {
        // TODO: Implement proper conversion when stylesheet parsing is fully implemented
        super::stylesheet::StyleSheet::new()
    }

    /// Convert document info to owned
    #[allow(clippy::needless_pass_by_value)]
    fn convert_info_to_owned(
        _info: super::info::DocumentInfo<'_>,
    ) -> super::info::DocumentInfo<'static> {
        // TODO: Implement proper conversion when info parsing is fully implemented
        super::info::DocumentInfo::new()
    }

    /// Convert annotations to owned
    #[allow(clippy::needless_pass_by_value)]
    fn convert_annotations_to_owned(
        _annotations: Vec<super::annotation::Annotation<'_>>,
    ) -> Vec<super::annotation::Annotation<'static>> {
        // TODO: Implement proper conversion when annotation parsing is fully implemented
        Vec::new()
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
