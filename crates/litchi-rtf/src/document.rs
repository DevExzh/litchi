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
    /// Footnotes and endnotes
    notes: Vec<super::section::Note<'a>>,
    /// Track changes/revisions
    revisions: Vec<super::annotation::Revision<'a>>,
}

impl<'a> RtfDocument<'a> {
    /// Parse an RTF document from a string.
    ///
    /// This method automatically detects and decompresses compressed RTF data.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_rtf::RtfDocument;
    ///
    /// let rtf = r#"{\rtf1\ansi Hello World!\par}"#;
    /// let doc = RtfDocument::parse(rtf)?;
    /// let text = doc.text();
    /// # Ok::<(), litchi_rtf::RtfError>(())
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
            notes: Self::convert_notes_to_owned(parsed.notes),
            revisions: Self::convert_revisions_to_owned(parsed.revisions),
        })
    }

    /// Parse an RTF document from a file.
    ///
    /// This method automatically detects and handles compressed RTF files.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_rtf::RtfDocument;
    ///
    /// let doc = RtfDocument::open("document.rtf")?;
    /// let text = doc.text();
    /// # Ok::<(), litchi_rtf::RtfError>(())
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
    /// use litchi_rtf::RtfDocument;
    ///
    /// let bytes = std::fs::read("document.rtf").map_err(|e| format!("IO error: {}", e))?;
    /// let doc = RtfDocument::from_bytes(&bytes)?;
    /// let text = doc.text();
    /// # Ok::<(), Box<dyn std::error::Error>>(())
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
    /// use litchi_rtf::RtfDocument;
    ///
    /// let doc = RtfDocument::open("document.rtf")?;
    /// for element in doc.elements() {
    ///     match element {
    ///         litchi_rtf::DocumentElement::Paragraph(para) => {
    ///             println!("Paragraph: {}", para.text());
    ///         }
    ///         litchi_rtf::DocumentElement::Table(table) => {
    ///             println!("Table with {} rows", table.row_count());
    ///         }
    ///     }
    /// }
    /// # Ok::<(), litchi_rtf::RtfError>(())
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
    /// use litchi_rtf::RtfDocument;
    ///
    /// let doc = RtfDocument::open("document.rtf")?;
    /// for (i, picture) in doc.pictures().iter().enumerate() {
    ///     println!("Picture {}: {:?}, {} bytes", i, picture.image_type, picture.data().len());
    /// }
    /// # Ok::<(), litchi_rtf::RtfError>(())
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
    /// use litchi_rtf::{RtfDocument, FieldType};
    ///
    /// let doc = RtfDocument::open("document.rtf")?;
    /// for field in doc.fields() {
    ///     if field.field_type == FieldType::Hyperlink {
    ///         if let Some(url) = field.extract_url() {
    ///             println!("Hyperlink: {}", url);
    ///         }
    ///     }
    /// }
    /// # Ok::<(), litchi_rtf::RtfError>(())
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
    fn convert_list_table_to_owned(
        table: super::list::ListTable<'_>,
    ) -> super::list::ListTable<'static> {
        let mut owned = super::list::ListTable::new();
        for list in table.lists() {
            owned.add(super::list::List {
                id: list.id,
                template_id: list.template_id,
                simple: list.simple,
                hybrid: list.hybrid,
                name: Cow::Owned(list.name.clone().into_owned()),
                levels: list
                    .levels
                    .iter()
                    .map(|level| super::list::ListLevel {
                        level: level.level,
                        level_type: level.level_type,
                        number_text: Cow::Owned(level.number_text.clone().into_owned()),
                        start_at: level.start_at,
                        justification: level.justification,
                        follow_previous: level.follow_previous,
                        follow: level.follow,
                        font_ref: level.font_ref,
                        indent: level.indent,
                        space: level.space,
                    })
                    .collect(),
            });
        }
        owned
    }

    /// Convert sections to owned
    fn convert_sections_to_owned(
        sections: Vec<super::section::Section<'_>>,
    ) -> Vec<super::section::Section<'static>> {
        sections
            .into_iter()
            .map(|section| super::section::Section {
                properties: section.properties,
                headers_footers: section
                    .headers_footers
                    .into_iter()
                    .map(|header_footer| super::section::HeaderFooter {
                        header_type: header_footer.header_type,
                        paragraphs: header_footer
                            .paragraphs
                            .into_iter()
                            .map(|paragraph| super::section::HeaderFooterParagraph {
                                text: Cow::Owned(paragraph.text.into_owned()),
                                formatting: paragraph.formatting,
                                paragraph: paragraph.paragraph,
                            })
                            .collect(),
                    })
                    .collect(),
            })
            .collect()
    }

    /// Convert bookmarks to owned
    fn convert_bookmarks_to_owned(
        bookmarks: super::bookmark::BookmarkTable<'_>,
    ) -> super::bookmark::BookmarkTable<'static> {
        let mut owned = super::bookmark::BookmarkTable::new();
        for bookmark in bookmarks.bookmarks() {
            owned.add(super::bookmark::Bookmark {
                name: Cow::Owned(bookmark.name.clone().into_owned()),
                position: bookmark.position,
                content: Cow::Owned(bookmark.content.clone().into_owned()),
                first_column: bookmark.first_column,
                last_column: bookmark.last_column,
                is_public: bookmark.is_public,
            });
        }
        owned
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
    fn convert_stylesheet_to_owned(
        stylesheet: super::stylesheet::StyleSheet<'_>,
    ) -> super::stylesheet::StyleSheet<'static> {
        let mut owned = super::stylesheet::StyleSheet::new();
        for style in stylesheet.styles() {
            owned.add(super::stylesheet::Style {
                id: style.id,
                name: Cow::Owned(style.name.clone().into_owned()),
                style_type: style.style_type,
                based_on: style.based_on,
                next_style: style.next_style,
                linked_style: style.linked_style,
                formatting: style.formatting,
                paragraph: style.paragraph,
                builtin: style.builtin,
                hidden: style.hidden,
                additive: style.additive,
                auto_update: style.auto_update,
                locked: style.locked,
                semi_hidden: style.semi_hidden,
                unhide_when_used: style.unhide_when_used,
                quick_format: style.quick_format,
                priority: style.priority,
                revision_id: style.revision_id,
                personal: style.personal,
                compose: style.compose,
                reply: style.reply,
            });
        }
        owned
    }

    /// Convert document info to owned
    fn convert_info_to_owned(
        info: super::info::DocumentInfo<'_>,
    ) -> super::info::DocumentInfo<'static> {
        super::info::DocumentInfo {
            title: info.title.map(|value| Cow::Owned(value.into_owned())),
            subject: info.subject.map(|value| Cow::Owned(value.into_owned())),
            author: info.author.map(|value| Cow::Owned(value.into_owned())),
            manager: info.manager.map(|value| Cow::Owned(value.into_owned())),
            company: info.company.map(|value| Cow::Owned(value.into_owned())),
            operator: info.operator.map(|value| Cow::Owned(value.into_owned())),
            category: info.category.map(|value| Cow::Owned(value.into_owned())),
            keywords: info.keywords.map(|value| Cow::Owned(value.into_owned())),
            comment: info.comment.map(|value| Cow::Owned(value.into_owned())),
            version: info.version,
            revision: info.revision,
            creation_time: info
                .creation_time
                .map(|value| Cow::Owned(value.into_owned())),
            revision_time: info
                .revision_time
                .map(|value| Cow::Owned(value.into_owned())),
            print_time: info.print_time.map(|value| Cow::Owned(value.into_owned())),
            backup_time: info.backup_time.map(|value| Cow::Owned(value.into_owned())),
            editing_time: info.editing_time,
            pages: info.pages,
            words: info.words,
            characters: info.characters,
            characters_with_spaces: info.characters_with_spaces,
            id: info.id,
        }
    }

    /// Convert annotations to owned
    fn convert_annotations_to_owned(
        annotations: Vec<super::annotation::Annotation<'_>>,
    ) -> Vec<super::annotation::Annotation<'static>> {
        annotations
            .into_iter()
            .map(|annotation| super::annotation::Annotation {
                annotation_type: annotation.annotation_type,
                id: annotation.id,
                author: Cow::Owned(annotation.author.into_owned()),
                initials: Cow::Owned(annotation.initials.into_owned()),
                date: annotation.date.map(|value| Cow::Owned(value.into_owned())),
                text: Cow::Owned(annotation.text.into_owned()),
                position: annotation.position,
                range_end: annotation.range_end,
                parent_id: annotation
                    .parent_id
                    .map(|value| Cow::Owned(value.into_owned())),
                icon: annotation.icon.map(|value| Cow::Owned(value.into_owned())),
                time: annotation.time.map(|value| Cow::Owned(value.into_owned())),
            })
            .collect()
    }

    /// Convert notes to owned
    fn convert_notes_to_owned(
        notes: Vec<super::section::Note<'_>>,
    ) -> Vec<super::section::Note<'static>> {
        notes
            .into_iter()
            .map(|note| super::section::Note {
                is_footnote: note.is_footnote,
                reference: Cow::Owned(note.reference.into_owned()),
                content: Cow::Owned(note.content.into_owned()),
                formatting: note.formatting,
            })
            .collect()
    }

    /// Convert revisions to owned
    fn convert_revisions_to_owned(
        revisions: Vec<super::annotation::Revision<'_>>,
    ) -> Vec<super::annotation::Revision<'static>> {
        revisions
            .into_iter()
            .map(|rev| super::annotation::Revision {
                revision_type: rev.revision_type,
                author: Cow::Owned(rev.author.into_owned()),
                date: rev.date.map(|d| Cow::Owned(d.into_owned())),
                id: rev.id,
                content: Cow::Owned(rev.content.into_owned()),
            })
            .collect()
    }

    /// Get all footnotes and endnotes in the document.
    pub fn notes(&self) -> &[super::section::Note<'_>] {
        &self.notes
    }

    /// Get all footnotes in the document.
    pub fn footnotes(&self) -> Vec<&super::section::Note<'_>> {
        self.notes.iter().filter(|n| n.is_footnote).collect()
    }

    /// Get all endnotes in the document.
    pub fn endnotes(&self) -> Vec<&super::section::Note<'_>> {
        self.notes.iter().filter(|n| !n.is_footnote).collect()
    }

    /// Get all track changes/revisions in the document.
    pub fn revisions(&self) -> &[super::annotation::Revision<'_>] {
        &self.revisions
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::{Alignment, ListFollow, ListJustification, ListLevelType, StyleType};

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

    #[test]
    fn parses_complete_document_info_without_leaking_into_body() {
        let rtf = r#"{\rtf1\ansi\ansicpg1252
            {\info
                {\title Annual \u20320? Report}
                {\subject Results}{\author Ada}{\manager Grace}
                {\company Caf\'e9 Corp \u8364?}{\operator Linus}{\category Finance}
                {\keywords alpha; beta}{\comment Reviewed}
                {\creatim\yr2025\mo7\dy14\hr9\min8\sec7}
                {\revtim\yr2026\mo1\dy2\hr3\min4\sec5}
                {\printim\yr2026\mo2\dy3}{\buptim\yr2024\mo12\dy31}
                \version4\vern9\edmins120\nofpages8\nofwords900
                \nofchars4200\nofcharsws5000\id77
            }
            Body text\par}"#;
        let doc = RtfDocument::parse(rtf).unwrap();
        let info = doc.info();
        assert_eq!(info.title.as_deref(), Some("Annual 你 Report"));
        assert_eq!(info.subject.as_deref(), Some("Results"));
        assert_eq!(info.author.as_deref(), Some("Ada"));
        assert_eq!(info.manager.as_deref(), Some("Grace"));
        assert_eq!(info.company.as_deref(), Some("Café Corp €"));
        assert_eq!(info.operator.as_deref(), Some("Linus"));
        assert_eq!(info.category.as_deref(), Some("Finance"));
        assert_eq!(info.keywords.as_deref(), Some("alpha; beta"));
        assert_eq!(info.comment.as_deref(), Some("Reviewed"));
        assert_eq!(info.creation_time.as_deref(), Some("2025-07-14T09:08:07"));
        assert_eq!(info.revision_time.as_deref(), Some("2026-01-02T03:04:05"));
        assert_eq!(info.print_time.as_deref(), Some("2026-02-03T00:00:00"));
        assert_eq!(info.backup_time.as_deref(), Some("2024-12-31T00:00:00"));
        assert_eq!(info.version, Some(4));
        assert_eq!(info.revision, Some(9));
        assert_eq!(info.editing_time, Some(120));
        assert_eq!(info.pages, Some(8));
        assert_eq!(info.words, Some(900));
        assert_eq!(info.characters, Some(4200));
        assert_eq!(info.characters_with_spaces, Some(5000));
        assert_eq!(info.id, Some(77));
        assert_eq!(doc.text().trim(), "Body text");
    }

    #[test]
    fn ignores_unknown_nested_info_destinations() {
        let rtf = r#"{\rtf1{\info{\*\unknown nested {data}}{\title Kept}}Text}"#;
        let doc = RtfDocument::parse(rtf).unwrap();
        assert_eq!(doc.info().title.as_deref(), Some("Kept"));
        assert_eq!(doc.text(), "Text");
    }

    #[test]
    fn parses_complete_stylesheet_without_leaking_names_into_body() {
        let rtf = r#"{\rtf1\ansi
            {\stylesheet
                {\s0\fs22\ql\snext0\sqformat\spriority0 Normal;}
                {\s1\b\qc\sb120\li240\keepn\sbasedon0\snext0\slink2
                    \sautoupd\shidden\slocked\ssemihidden\sunhideused\sqformat
                    \spriority9\styrsid42\spersonal\scompose\sreply Heading \u20320?;}
                {\*\cs2\i\additive\sbasedon0\slink1 Emphasis;}
                {\*\ds3 Section Style;}
                {\*\ts4{\*\unknown ignored} Table Style;}
            }
            Body}"#;
        let doc = RtfDocument::parse(rtf).unwrap();
        assert_eq!(doc.text().trim(), "Body");
        assert_eq!(doc.stylesheet().styles().len(), 5);

        let heading = doc.stylesheet().get_typed(StyleType::Paragraph, 1).unwrap();
        assert_eq!(heading.name, "Heading 你");
        assert_eq!(heading.based_on, Some(0));
        assert_eq!(heading.next_style, Some(0));
        assert_eq!(heading.linked_style, Some(2));
        assert!(heading.formatting.bold);
        let paragraph = heading.paragraph.unwrap();
        assert_eq!(paragraph.alignment, Alignment::Center);
        assert_eq!(paragraph.spacing.before, 120);
        assert_eq!(paragraph.indentation.left, 240);
        assert!(paragraph.keep_next);
        assert!(heading.auto_update);
        assert!(heading.hidden);
        assert!(heading.locked);
        assert!(heading.semi_hidden);
        assert!(heading.unhide_when_used);
        assert!(heading.quick_format);
        assert_eq!(heading.priority, Some(9));
        assert_eq!(heading.revision_id, Some(42));
        assert!(heading.personal);
        assert!(heading.compose);
        assert!(heading.reply);

        let emphasis = doc.stylesheet().get_typed(StyleType::Character, 2).unwrap();
        assert_eq!(emphasis.name, "Emphasis");
        assert!(emphasis.formatting.italic);
        assert!(emphasis.additive);
        assert!(emphasis.paragraph.is_none());
        assert!(doc.stylesheet().get_typed(StyleType::Section, 3).is_some());
        assert!(doc.stylesheet().get_typed(StyleType::Table, 4).is_some());
    }

    #[test]
    fn parses_list_and_override_tables_without_leaking_labels() {
        let rtf = r#"{\rtf1\ansi
            {\*\listtable
                {\list\listtemplateid42\listhybrid
                    {\listlevel\levelnfc0\leveljc2\levelfollow1\levelstartat3
                        \levelspace120\levelindent360
                        {\leveltext\'02\'00.;}{\levelnumbers\'01;}\f2}
                    {\listlevel\levelnfc23\leveljc0\levelfollow2\levelstartat1
                        {\leveltext\'01\u8226?;}{\levelnumbers;}}
                    {\listname Outline;}\listid77}
            }
            {\listoverridetable
                {\listoverride\listid77\listoverridecount1
                    {\lfolevel\listoverridestartat\levelstartat9}\ls4}}
            Body}"#;
        let doc = RtfDocument::parse(rtf).unwrap();
        assert_eq!(doc.text().trim(), "Body");
        assert_eq!(doc.list_table().lists().len(), 1);

        let list = doc.list_table().get(77).unwrap();
        assert_eq!(list.template_id, 42);
        assert!(!list.simple);
        assert!(list.hybrid);
        assert_eq!(list.name, "Outline");
        assert_eq!(list.levels.len(), 2);
        let decimal = &list.levels[0];
        assert_eq!(decimal.level_type, ListLevelType::Decimal);
        assert_eq!(decimal.number_text, "\0.");
        assert_eq!(decimal.start_at, 3);
        assert_eq!(decimal.justification, ListJustification::Right);
        assert_eq!(decimal.follow, ListFollow::Space);
        assert_eq!(decimal.font_ref, 2);
        assert_eq!(decimal.indent, 360);
        assert_eq!(decimal.space, 120);
        let bullet = &list.levels[1];
        assert_eq!(bullet.level_type, ListLevelType::Bullet);
        assert_eq!(bullet.number_text, "•");
        assert_eq!(bullet.follow, ListFollow::Nothing);

        let list_override = doc.list_override_table().get(4).unwrap();
        assert_eq!(list_override.list_id, 77);
        assert_eq!(list_override.level_count_override, Some(1));
        assert_eq!(list_override.start_at_override, Some(9));
    }

    #[test]
    fn preserves_paragraph_list_instance_and_level() {
        let doc = RtfDocument::parse(r#"{\rtf1\pard\ls4\ilvl2 Listed text}"#).unwrap();
        assert_eq!(doc.text(), "Listed text");
        let paragraph = doc.blocks().last().unwrap().paragraph;
        assert_eq!(paragraph.list_override, Some(4));
        assert_eq!(paragraph.list_level, Some(2));
    }

    #[test]
    fn parses_nested_bookmarks_with_range_metadata() {
        let rtf = r#"{\rtf1\ansi Before {\*\bkmkstart\bkmkcolf1\bkmkcoll3\bkmkpub Outer}alpha {\*\bkmkstart Inner}\u20320?{\*\bkmkend Inner} omega{\*\bkmkend Outer} After}"#;
        let doc = RtfDocument::parse(rtf).unwrap();
        assert_eq!(doc.text(), "Before alpha 你 omega After");

        let outer = doc.bookmarks().get("Outer").unwrap();
        assert_eq!(outer.position, "Before ".len());
        assert_eq!(outer.content, "alpha 你 omega");
        assert_eq!(outer.first_column, Some(1));
        assert_eq!(outer.last_column, Some(3));
        assert!(outer.is_public);

        let inner = doc.bookmarks().get("Inner").unwrap();
        assert_eq!(inner.position, "Before alpha ".len());
        assert_eq!(inner.content, "你");
    }

    #[test]
    fn parses_annotation_range_author_and_body_without_text_leakage() {
        let rtf = r#"{\rtf1\ansi aaa {\*\atrfstart 7}bbb{\*\atrfend 7}{\*\atnid MM}{\*\atnauthor Max Mustermann}\chatn{\*\annotation{\*\atnref 7}{\*\atndate 667322855}{\*\atnparent root}{\*\atnicn 2}{\*\atntime 42}Comment \u20320?} ccc}"#;
        let doc = RtfDocument::parse(rtf).unwrap();
        assert_eq!(doc.text(), "aaa bbb ccc");
        assert_eq!(doc.annotations().len(), 1);

        let annotation = &doc.annotations()[0];
        assert_eq!(annotation.id, 7);
        assert_eq!(annotation.author, "Max Mustermann");
        assert_eq!(annotation.initials, "MM");
        assert_eq!(annotation.date.as_deref(), Some("667322855"));
        assert_eq!(annotation.text, "Comment 你");
        assert_eq!(annotation.position, "aaa ".len());
        assert_eq!(annotation.range_end, "aaa bbb".len());
        assert_eq!(annotation.parent_id.as_deref(), Some("root"));
        assert_eq!(annotation.icon.as_deref(), Some("2"));
        assert_eq!(annotation.time.as_deref(), Some("42"));
    }

    #[test]
    fn preserves_parsed_headers_and_footers_in_owned_document() {
        let rtf = r#"{\rtf1\ansi\sectd\sbkeven\pgwsxn10000\pghsxn14000\marglsxn900\margrsxn800\margtsxn700\margbsxn600\guttersxn120\headery300\footery400\lndscpsxn\cols2\colsx360\pgnstarts5\pgnucrm\vertalc\linemod1\lineppage{\header Main \u20320? header\par Second line}{\footer Page footer}Body}"#;
        let doc = RtfDocument::parse(rtf).unwrap();
        assert_eq!(doc.text(), "Body");
        assert_eq!(doc.sections().len(), 1);
        let section = &doc.sections()[0];
        assert_eq!(
            section.properties.break_type,
            crate::SectionBreakType::EvenPage
        );
        assert_eq!(section.properties.page_width, 10000);
        assert_eq!(section.properties.page_height, 14000);
        assert_eq!(section.properties.margin_left, 900);
        assert_eq!(section.properties.margin_right, 800);
        assert_eq!(section.properties.margin_top, 700);
        assert_eq!(section.properties.margin_bottom, 600);
        assert_eq!(section.properties.margin_gutter, 120);
        assert_eq!(section.properties.header_distance, 300);
        assert_eq!(section.properties.footer_distance, 400);
        assert_eq!(
            section.properties.orientation,
            crate::PageOrientation::Landscape
        );
        assert_eq!(section.properties.columns, 2);
        assert_eq!(section.properties.column_space, 360);
        assert_eq!(section.properties.page_number_start, 5);
        assert_eq!(
            section.properties.page_number_format,
            crate::PageNumberFormat::UpperRoman
        );
        assert_eq!(
            section.properties.vertical_alignment,
            crate::VerticalAlignment::Center
        );
        assert!(section.properties.line_numbering);
        assert!(section.properties.line_number_restart);
        assert_eq!(
            section
                .get_header(super::super::section::HeaderFooterType::Header)
                .unwrap()
                .text(),
            "Main 你 header\nSecond line"
        );
        assert_eq!(
            section
                .get_header(super::super::section::HeaderFooterType::Footer)
                .unwrap()
                .text(),
            "Page footer"
        );
    }
}
