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
    /// Ordered positional legacy form fields.
    form_fields: Vec<super::form_field::FormField<'a>>,
    /// Inert producer provenance from the generator destination.
    generator: Option<crate::DocumentGenerator<'a>>,
    /// Ordered revision-save/session provenance.
    revision_save: Option<crate::RevisionSaveMetadata>,
    /// Ordered inert XML namespace table; `Some([])` preserves an empty table.
    xml_namespaces: Option<Vec<crate::XmlNamespace<'a>>>,
    /// Inert Office theme package and optional color-scheme mapping bytes.
    theme: Option<crate::DocumentTheme<'a>>,
    /// Inert latent-style defaults and ordered exceptions.
    latent_styles: Option<crate::LatentStyles<'a>>,
    /// Inert custom XML data-store bytes.
    data_store: Option<crate::DocumentDataStore<'a>>,
    /// Document-level defaults for mathematical layout.
    math_properties: Option<crate::DocumentMathProperties>,
    /// Default primary, East Asian, and complex-script languages.
    language_defaults: crate::DocumentLanguageDefaults,
    /// Explicit default bidirectional precedence for document text.
    document_direction: Option<crate::TextDirection>,
    /// Whether the document gutter is positioned on the right.
    gutter_on_right: bool,
    /// Embedded and linked objects
    objects: Vec<super::object::EmbeddedObject<'a>>,
    /// Ordered inert document-variable metadata
    document_variables: Vec<super::document_variable::DocumentVariable<'a>>,
    /// Ordered inert user-defined document properties
    user_properties: Vec<super::user_property::UserProperty<'a>>,
    /// Ordered inert index and table-of-contents source marks
    navigation_entries: Vec<super::navigation_entry::NavigationEntry<'a>>,
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
    /// Ordered revision-author table referenced by revision author indices.
    revision_authors: Vec<super::annotation::RevisionAuthor<'a>>,
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

    /// Parse RTF from its original byte representation.
    ///
    /// Use this entry point when the document can contain `bin` destinations or
    /// legacy-code-page bytes that are not valid UTF-8.
    pub fn parse_bytes(input: &[u8]) -> RtfResult<RtfDocument<'static>> {
        Self::parse_internal(input)
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
        let input_str: String = input_bytes.iter().map(|byte| char::from(*byte)).collect();

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
                owned_table.set_direction(table.direction());
                for row in table.rows() {
                    let mut owned_row = super::table::Row::new();
                    owned_row.set_direction(row.direction());
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

        let owned_objects = parsed
            .objects
            .into_iter()
            .map(|object| super::object::EmbeddedObject {
                kind: object.kind,
                class_name: Cow::Owned(object.class_name.into_owned()),
                name: Cow::Owned(object.name.into_owned()),
                width: object.width,
                height: object.height,
                locked: object.locked,
                update_requested: object.update_requested,
                set_size: object.set_size,
                result_text: Cow::Owned(object.result_text.into_owned()),
                result_picture_indices: object.result_picture_indices,
                data: object.data,
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
            form_fields: parsed
                .form_fields
                .into_iter()
                .map(super::form_field::FormField::into_owned)
                .collect(),
            generator: parsed
                .generator
                .map(crate::DocumentGenerator::into_owned),
            revision_save: parsed.revision_save,
            xml_namespaces: parsed.saw_xml_namespace_table.then(|| {
                parsed
                    .xml_namespaces
                    .into_iter()
                    .map(crate::XmlNamespace::into_owned)
                    .collect()
            }),
            theme: parsed.theme.map(crate::DocumentTheme::into_owned),
            latent_styles: parsed.latent_styles.map(crate::LatentStyles::into_owned),
            data_store: parsed.data_store.map(crate::DocumentDataStore::into_owned),
            math_properties: parsed.math_properties,
            language_defaults: parsed.language_defaults,
            document_direction: parsed.document_direction,
            gutter_on_right: parsed.gutter_on_right,
            objects: owned_objects,
            document_variables: parsed
                .document_variables
                .into_iter()
                .map(super::document_variable::DocumentVariable::into_owned)
                .collect(),
            user_properties: parsed
                .user_properties
                .into_iter()
                .map(super::user_property::UserProperty::into_owned)
                .collect(),
            navigation_entries: parsed
                .navigation_entries
                .into_iter()
                .map(super::navigation_entry::NavigationEntry::into_owned)
                .collect(),
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
            revision_authors: parsed
                .revision_authors
                .into_iter()
                .map(super::annotation::RevisionAuthor::into_owned)
                .collect(),
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

    /// Return ordered positional legacy form fields.
    pub fn form_fields(&self) -> &[super::form_field::FormField<'_>] {
        &self.form_fields
    }

    /// Return inert producer provenance from the RTF generator destination.
    pub fn generator(&self) -> Option<&crate::DocumentGenerator<'_>> {
        self.generator.as_ref()
    }

    /// Set validated inert producer provenance.
    pub fn set_generator(&mut self, generator: crate::DocumentGenerator<'a>) -> RtfResult<()> {
        generator.validate()?;
        self.generator = Some(generator);
        Ok(())
    }

    /// Remove producer provenance metadata.
    pub fn clear_generator(&mut self) {
        self.generator = None;
    }

    /// Return ordered revision-save/session provenance.
    pub fn revision_save_metadata(&self) -> Option<&crate::RevisionSaveMetadata> {
        self.revision_save.as_ref()
    }

    /// Replace revision-save/session provenance after full validation.
    pub fn set_revision_save_metadata(
        &mut self,
        metadata: crate::RevisionSaveMetadata,
    ) -> RtfResult<()> {
        metadata.validate()?;
        self.revision_save = Some(metadata);
        Ok(())
    }

    /// Remove revision-save/session provenance.
    pub fn clear_revision_save_metadata(&mut self) {
        self.revision_save = None;
    }

    /// Return the ordered inert XML namespace table, preserving empty-table presence.
    pub fn xml_namespaces(&self) -> Option<&[crate::XmlNamespace<'_>]> {
        self.xml_namespaces.as_deref()
    }

    /// Replace the XML namespace table after full validation.
    pub fn set_xml_namespaces(
        &mut self,
        namespaces: Vec<crate::XmlNamespace<'a>>,
    ) -> RtfResult<()> {
        Self::validate_xml_namespaces(&namespaces)?;
        self.xml_namespaces = Some(namespaces);
        Ok(())
    }

    /// Append one inert XML namespace entry, creating the table if absent.
    pub fn push_xml_namespace(&mut self, namespace: crate::XmlNamespace<'a>) -> RtfResult<()> {
        namespace.validate()?;
        let was_present = self.xml_namespaces.is_some();
        let mut namespaces = self.xml_namespaces.take().unwrap_or_default();
        namespaces.push(namespace);
        if let Err(error) = Self::validate_xml_namespaces(&namespaces) {
            namespaces.pop();
            self.xml_namespaces = was_present.then_some(namespaces);
            return Err(error);
        }
        self.xml_namespaces = Some(namespaces);
        Ok(())
    }

    /// Remove the XML namespace table entirely.
    pub fn clear_xml_namespaces(&mut self) {
        self.xml_namespaces = None;
    }

    /// Return inert Office theme bytes without interpreting their contents.
    pub fn theme(&self) -> Option<&crate::DocumentTheme<'_>> {
        self.theme.as_ref()
    }

    /// Replace inert Office theme bytes after bounds validation.
    pub fn set_theme(&mut self, theme: crate::DocumentTheme<'a>) -> RtfResult<()> {
        theme.validate()?;
        self.theme = Some(theme);
        Ok(())
    }

    /// Remove theme and color-scheme mapping payloads.
    pub fn clear_theme(&mut self) {
        self.theme = None;
    }

    /// Return inert latent-style defaults and ordered exceptions.
    pub fn latent_styles(&self) -> Option<&crate::LatentStyles<'_>> {
        self.latent_styles.as_ref()
    }

    /// Replace latent-style metadata after full validation.
    pub fn set_latent_styles(&mut self, styles: crate::LatentStyles<'a>) -> RtfResult<()> {
        styles.validate()?;
        self.latent_styles = Some(styles);
        Ok(())
    }

    /// Remove latent-style metadata.
    pub fn clear_latent_styles(&mut self) {
        self.latent_styles = None;
    }

    /// Return inert custom XML data-store bytes without interpreting them.
    pub fn data_store(&self) -> Option<&crate::DocumentDataStore<'_>> {
        self.data_store.as_ref()
    }

    /// Replace inert data-store bytes after bounds validation.
    pub fn set_data_store(&mut self, data_store: crate::DocumentDataStore<'a>) -> RtfResult<()> {
        data_store.validate()?;
        self.data_store = Some(data_store);
        Ok(())
    }

    /// Remove the custom XML data-store payload.
    pub fn clear_data_store(&mut self) {
        self.data_store = None;
    }

    /// Return document-level mathematical layout defaults.
    pub fn math_properties(&self) -> Option<&crate::DocumentMathProperties> {
        self.math_properties.as_ref()
    }

    /// Replace document-level mathematical layout defaults after validation.
    pub fn set_math_properties(
        &mut self,
        properties: crate::DocumentMathProperties,
    ) -> RtfResult<()> {
        properties.validate()?;
        self.math_properties = Some(properties);
        Ok(())
    }

    /// Remove document-level mathematical layout defaults.
    pub fn clear_math_properties(&mut self) {
        self.math_properties = None;
    }

    /// Return language defaults declared by the RTF header.
    pub fn language_defaults(&self) -> &crate::DocumentLanguageDefaults {
        &self.language_defaults
    }

    /// Replace document language defaults.
    pub fn set_language_defaults(
        &mut self,
        defaults: crate::DocumentLanguageDefaults,
    ) -> RtfResult<()> {
        defaults.validate()?;
        self.language_defaults = defaults;
        Ok(())
    }

    /// Remove all document language defaults.
    pub fn clear_language_defaults(&mut self) {
        self.language_defaults = crate::DocumentLanguageDefaults::default();
    }

    /// Return the explicit document-wide bidirectional precedence.
    pub fn document_direction(&self) -> Option<crate::TextDirection> {
        self.document_direction
    }

    /// Set the explicit document-wide bidirectional precedence.
    pub fn set_document_direction(&mut self, direction: crate::TextDirection) {
        self.document_direction = Some(direction);
    }

    /// Remove the explicit document-wide bidirectional precedence.
    pub fn clear_document_direction(&mut self) {
        self.document_direction = None;
    }

    /// Return whether the document gutter is positioned on the right.
    pub fn gutter_on_right(&self) -> bool {
        self.gutter_on_right
    }

    /// Position the document gutter on the right when `enabled` is true.
    pub fn set_gutter_on_right(&mut self, enabled: bool) {
        self.gutter_on_right = enabled;
    }

    fn validate_xml_namespaces(namespaces: &[crate::XmlNamespace<'_>]) -> RtfResult<()> {
        if namespaces.len() > crate::xml_namespace::MAX_XML_NAMESPACES {
            return Err(RtfError::MalformedDocument(
                "RTF XML namespace count exceeds the safety limit".to_string(),
            ));
        }
        let mut total = 0usize;
        for (index, namespace) in namespaces.iter().enumerate() {
            namespace.validate()?;
            if namespaces[..index]
                .iter()
                .any(|existing| existing.id == namespace.id)
            {
                return Err(RtfError::MalformedDocument(
                    "RTF XML namespace IDs must be unique".to_string(),
                ));
            }
            total = total.checked_add(namespace.namespace.len()).ok_or_else(|| {
                RtfError::MalformedDocument(
                    "RTF XML namespace aggregate size overflow".to_string(),
                )
            })?;
            if total > crate::xml_namespace::MAX_XML_NAMESPACE_TOTAL_BYTES {
                return Err(RtfError::MalformedDocument(
                    "RTF XML namespace aggregate text exceeds the safety limit".to_string(),
                ));
            }
        }
        Ok(())
    }

    /// Append inert form-field metadata at a valid visible body range.
    pub fn push_form_field(
        &mut self,
        field: super::form_field::FormField<'a>,
    ) -> RtfResult<()> {
        field.validate()?;
        if self.form_fields.len() >= super::form_field::MAX_FORM_FIELDS {
            return Err(RtfError::MalformedDocument(
                "RTF form-field count exceeds the safety limit".to_string(),
            ));
        }
        let body = self.text();
        let result = body.get(field.position..field.range_end).ok_or_else(|| {
            RtfError::MalformedDocument(
                "RTF form-field range is outside body text or splits a character".to_string(),
            )
        })?;
        if result != field.result_text {
            return Err(RtfError::MalformedDocument(
                "RTF form-field result does not match its visible body range".to_string(),
            ));
        }
        if field.position != field.range_end
            && self.form_fields.iter().any(|existing| {
                existing.position != existing.range_end
                    && field.position < existing.range_end
                    && existing.position < field.range_end
            })
        {
            return Err(RtfError::MalformedDocument(
                "RTF form-field result ranges cannot overlap".to_string(),
            ));
        }
        let total = self
            .form_fields
            .iter()
            .try_fold(field.text_bytes().unwrap_or(usize::MAX), |total, existing| {
                total.checked_add(existing.text_bytes()?)
            })
            .ok_or_else(|| {
                RtfError::MalformedDocument("RTF form-field aggregate size overflow".to_string())
            })?;
        if total > super::form_field::MAX_FORM_FIELD_TOTAL_BYTES {
            return Err(RtfError::MalformedDocument(
                "RTF form-field aggregate text exceeds the safety limit".to_string(),
            ));
        }
        self.form_fields.push(field);
        Ok(())
    }

    /// Remove all legacy form-field metadata without changing visible body text.
    pub fn clear_form_fields(&mut self) {
        self.form_fields.clear();
    }

    /// Return embedded and linked object records without activating their content.
    pub fn objects(&self) -> &[super::object::EmbeddedObject<'_>] {
        &self.objects
    }

    /// Return ordered inert RTF document-variable name/value pairs.
    pub fn document_variables(&self) -> &[super::document_variable::DocumentVariable<'_>] {
        &self.document_variables
    }

    /// Append a document variable without evaluating or resolving it.
    pub fn push_document_variable(
        &mut self,
        variable: super::document_variable::DocumentVariable<'a>,
    ) -> RtfResult<()> {
        variable.validate()?;
        if self.document_variables.len()
            >= super::document_variable::MAX_DOCUMENT_VARIABLES
        {
            return Err(RtfError::MalformedDocument(
                "RTF document-variable count limit exceeded".to_string(),
            ));
        }
        let aggregate = self.document_variables.iter().try_fold(
            variable.name.len() + variable.value.len(),
            |size, existing| {
                size.checked_add(existing.name.len())
                    .and_then(|size| size.checked_add(existing.value.len()))
            },
        );
        if aggregate.is_none_or(|size| {
            size > super::document_variable::MAX_DOCUMENT_VARIABLE_TEXT_BYTES
        }) {
            return Err(RtfError::MalformedDocument(
                "RTF document-variable aggregate text limit exceeded".to_string(),
            ));
        }
        self.document_variables.push(variable);
        Ok(())
    }

    /// Remove all document variables.
    pub fn clear_document_variables(&mut self) {
        self.document_variables.clear();
    }

    /// Return ordered, inert RTF user-defined document properties.
    pub fn user_properties(&self) -> &[super::user_property::UserProperty<'_>] {
        &self.user_properties
    }

    /// Append a unique user-defined property without evaluating its value or link.
    pub fn push_user_property(
        &mut self,
        property: super::user_property::UserProperty<'a>,
    ) -> RtfResult<()> {
        property.validate()?;
        if self.user_properties.len() >= super::user_property::MAX_USER_PROPERTIES {
            return Err(RtfError::MalformedDocument(
                "RTF user-property count limit exceeded".to_string(),
            ));
        }
        if self
            .user_properties
            .iter()
            .any(|existing| existing.name == property.name)
        {
            return Err(RtfError::MalformedDocument(format!(
                "duplicate RTF user-property name: {}",
                property.name
            )));
        }
        let aggregate = property.text_bytes().and_then(|initial| {
            self.user_properties.iter().try_fold(initial, |size, existing| {
                size.checked_add(existing.text_bytes()?)
            })
        });
        if aggregate.is_none_or(|size| {
            size > super::user_property::MAX_USER_PROPERTY_TEXT_BYTES
        }) {
            return Err(RtfError::MalformedDocument(
                "RTF user-property aggregate text limit exceeded".to_string(),
            ));
        }
        self.user_properties.push(property);
        Ok(())
    }

    /// Remove all user-defined properties.
    pub fn clear_user_properties(&mut self) {
        self.user_properties.clear();
    }

    /// Return ordered, inert index and table-of-contents source marks.
    pub fn navigation_entries(&self) -> &[super::navigation_entry::NavigationEntry<'_>] {
        &self.navigation_entries
    }

    /// Append an inert source mark at a valid UTF-8 body position.
    pub fn push_navigation_entry(
        &mut self,
        entry: super::navigation_entry::NavigationEntry<'a>,
    ) -> RtfResult<()> {
        entry.validate()?;
        let body = self.text();
        if body.get(entry.position()..entry.position()).is_none() {
            return Err(RtfError::MalformedDocument(
                "RTF navigation-entry position is outside body text or splits a character"
                    .to_string(),
            ));
        }
        if self.navigation_entries.len()
            >= super::navigation_entry::MAX_NAVIGATION_ENTRIES
        {
            return Err(RtfError::MalformedDocument(
                "RTF navigation-entry count limit exceeded".to_string(),
            ));
        }
        let aggregate = entry.text_bytes().and_then(|initial| {
            self.navigation_entries
                .iter()
                .try_fold(initial, |size, existing| {
                    size.checked_add(existing.text_bytes()?)
                })
        });
        if aggregate.is_none_or(|size| {
            size > super::navigation_entry::MAX_NAVIGATION_ENTRY_TEXT_TOTAL_BYTES
        }) {
            return Err(RtfError::MalformedDocument(
                "RTF navigation-entry aggregate text limit exceeded".to_string(),
            ));
        }
        self.navigation_entries.push(entry);
        Ok(())
    }

    /// Remove all index and table-of-contents source marks.
    pub fn clear_navigation_entries(&mut self) {
        self.navigation_entries.clear();
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

    /// Return inert document and revision-protection metadata.
    pub fn protection(&self) -> &crate::DocumentProtection<'_> {
        &self.info.protection
    }

    /// Replace inert document-protection metadata.
    pub fn set_protection(&mut self, protection: crate::DocumentProtection<'a>) -> RtfResult<()> {
        protection.validate()?;
        self.info.protection = protection;
        Ok(())
    }

    /// Remove all document-protection metadata.
    pub fn clear_protection(&mut self) {
        self.info.protection = crate::DocumentProtection::default();
    }

    /// Get all annotations (comments) in the document.
    ///
    /// Returns document annotations and revisions.
    pub fn annotations(&self) -> &[super::annotation::Annotation<'_>] {
        &self.annotations
    }

    /// Append an inert comment annotation after validating its body range.
    pub fn push_annotation(
        &mut self,
        annotation: super::annotation::Annotation<'a>,
    ) -> RtfResult<()> {
        annotation.validate()?;
        let body = self.text();
        if body.get(annotation.position..annotation.range_end).is_none() {
            return Err(RtfError::MalformedDocument(
                "RTF annotation range is outside body text or splits a character".to_string(),
            ));
        }
        if self.annotations.len() >= super::annotation::MAX_ANNOTATIONS {
            return Err(RtfError::MalformedDocument(
                "RTF annotation count limit exceeded".to_string(),
            ));
        }
        if annotation.has_reference
            && self
                .annotations
                .iter()
                .any(|existing| existing.has_reference && existing.id == annotation.id)
        {
            return Err(RtfError::MalformedDocument(
                "duplicate RTF annotation reference".to_string(),
            ));
        }
        let aggregate = annotation.text_bytes().and_then(|initial| {
            self.annotations.iter().try_fold(initial, |size, existing| {
                size.checked_add(existing.text_bytes()?)
            })
        });
        if aggregate.is_none_or(|size| {
            size > super::annotation::MAX_ANNOTATION_TEXT_TOTAL_BYTES
        }) {
            return Err(RtfError::MalformedDocument(
                "RTF annotation aggregate text limit exceeded".to_string(),
            ));
        }
        self.annotations.push(annotation);
        Ok(())
    }

    /// Remove all comment annotations.
    pub fn clear_annotations(&mut self) {
        self.annotations.clear();
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
    fn convert_shapes_to_owned(
        shapes: Vec<super::shape::Shape<'_>>,
    ) -> Vec<super::shape::Shape<'static>> {
        shapes
            .into_iter()
            .map(Self::convert_shape_to_owned)
            .collect()
    }

    /// Convert shape groups to owned
    fn convert_shape_groups_to_owned(
        groups: Vec<super::shape::ShapeGroup<'_>>,
    ) -> Vec<super::shape::ShapeGroup<'static>> {
        groups
            .into_iter()
            .map(Self::convert_shape_group_to_owned)
            .collect()
    }

    fn convert_shape_group_to_owned(
        group: super::shape::ShapeGroup<'_>,
    ) -> super::shape::ShapeGroup<'static> {
        super::shape::ShapeGroup {
            name: Cow::Owned(group.name.into_owned()),
            shapes: group
                .shapes
                .into_iter()
                .map(Self::convert_shape_to_owned)
                .collect(),
            groups: group
                .groups
                .into_iter()
                .map(Self::convert_shape_group_to_owned)
                .collect(),
            geometry: group.geometry,
            properties: group
                .properties
                .into_iter()
                .map(Self::convert_shape_property_to_owned)
                .collect(),
        }
    }

    fn convert_shape_to_owned(shape: super::shape::Shape<'_>) -> super::shape::Shape<'static> {
        super::shape::Shape {
            shape_type: shape.shape_type,
            geometry: shape.geometry,
            fill: shape.fill,
            border: shape.border,
            line: shape.line,
            text: Cow::Owned(shape.text.into_owned()),
            text_formatting: shape.text_formatting,
            wrap_mode: shape.wrap_mode,
            behind_doc: shape.behind_doc,
            is_background: shape.is_background,
            locked: shape.locked,
            name: Cow::Owned(shape.name.into_owned()),
            properties: shape
                .properties
                .into_iter()
                .map(Self::convert_shape_property_to_owned)
                .collect(),
        }
    }

    fn convert_shape_property_to_owned(
        property: super::shape::ShapeProperty<'_>,
    ) -> super::shape::ShapeProperty<'static> {
        super::shape::ShapeProperty {
            name: Cow::Owned(property.name.into_owned()),
            value: Cow::Owned(property.value.into_owned()),
        }
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
            document_comment: info.document_comment.map(|value| Cow::Owned(value.into_owned())),
            hyperlink_base: info.hyperlink_base.map(|value| Cow::Owned(value.into_owned())),
            version: info.version,
            revision: info.revision,
            creation_time: info
                .creation_time
                .map(|value| Cow::Owned(value.into_owned())),
            creation_timestamp: info.creation_timestamp,
            revision_time: info
                .revision_time
                .map(|value| Cow::Owned(value.into_owned())),
            revision_timestamp: info.revision_timestamp,
            print_time: info.print_time.map(|value| Cow::Owned(value.into_owned())),
            print_timestamp: info.print_timestamp,
            backup_time: info.backup_time.map(|value| Cow::Owned(value.into_owned())),
            backup_timestamp: info.backup_timestamp,
            editing_time: info.editing_time,
            pages: info.pages,
            words: info.words,
            characters: info.characters,
            characters_with_spaces: info.characters_with_spaces,
            id: info.id,
            protection: info.protection.into_owned(),
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
                has_reference: annotation.has_reference,
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
                position: rev.position,
                range_end: rev.range_end,
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

    /// Get the ordered revision-author table.
    pub fn revision_authors(&self) -> &[super::annotation::RevisionAuthor<'_>] {
        &self.revision_authors
    }

    /// Append an entry to the ordered revision-author table.
    pub fn push_revision_author(
        &mut self,
        author: super::annotation::RevisionAuthor<'a>,
    ) -> RtfResult<()> {
        author.validate()?;
        if self.revision_authors.len() >= super::annotation::MAX_REVISION_AUTHORS {
            return Err(RtfError::MalformedDocument(
                "RTF revision author count exceeds the safety limit".to_string(),
            ));
        }
        let total = self
            .revision_authors
            .iter()
            .try_fold(author.name.len(), |total, existing| {
                total.checked_add(existing.name.len())
            })
            .ok_or_else(|| {
                RtfError::MalformedDocument(
                    "RTF aggregate revision-author size overflow".to_string(),
                )
            })?;
        if total > super::annotation::MAX_REVISION_AUTHOR_TEXT_TOTAL_BYTES {
            return Err(RtfError::MalformedDocument(
                "RTF aggregate revision-author text exceeds the safety limit".to_string(),
            ));
        }
        self.revision_authors.push(author);
        Ok(())
    }

    /// Remove the revision-author table when no revision still references it.
    pub fn clear_revision_authors(&mut self) -> RtfResult<()> {
        if !self.revisions.is_empty() {
            return Err(RtfError::MalformedDocument(
                "cannot clear an RTF revision-author table while revisions reference it"
                    .to_string(),
            ));
        }
        self.revision_authors.clear();
        Ok(())
    }

    /// Append a validated tracked change.
    pub fn push_revision(
        &mut self,
        revision: super::annotation::Revision<'a>,
    ) -> RtfResult<()> {
        revision.validate()?;
        if self.revisions.len() >= super::annotation::MAX_REVISIONS {
            return Err(RtfError::MalformedDocument(
                "RTF revision count exceeds the safety limit".to_string(),
            ));
        }
        let author_index = usize::try_from(revision.id).map_err(|_| {
            RtfError::MalformedDocument(
                "RTF revision author index cannot be negative".to_string(),
            )
        })?;
        let author = self.revision_authors.get(author_index).ok_or_else(|| {
            RtfError::MalformedDocument(
                "RTF revision author index is outside revtbl".to_string(),
            )
        })?;
        if author.name != revision.author {
            return Err(RtfError::MalformedDocument(
                "RTF revision author does not match its revtbl entry".to_string(),
            ));
        }

        let body = self.text();
        match revision.revision_type {
            super::annotation::RevisionType::Insertion => {
                let content = body
                    .get(revision.position..revision.range_end)
                    .ok_or_else(|| {
                        RtfError::MalformedDocument(
                            "RTF insertion range is outside body text or splits a character"
                                .to_string(),
                        )
                    })?;
                if content != revision.content {
                    return Err(RtfError::MalformedDocument(
                        "RTF insertion content does not match its visible body range".to_string(),
                    ));
                }
                if self.revisions.iter().any(|existing| {
                    existing.revision_type == super::annotation::RevisionType::Insertion
                        && revision.position < existing.range_end
                        && existing.position < revision.range_end
                }) {
                    return Err(RtfError::MalformedDocument(
                        "RTF insertion ranges cannot overlap".to_string(),
                    ));
                }
            },
            super::annotation::RevisionType::Deletion => {
                if body.get(..revision.position).is_none() {
                    return Err(RtfError::MalformedDocument(
                        "RTF deletion position is outside body text or splits a character"
                            .to_string(),
                    ));
                }
            },
            _ => {
                return Err(RtfError::MalformedDocument(
                    "this RTF revision kind has no lossless scoped-run representation"
                        .to_string(),
                ));
            },
        }
        let total = self
            .revisions
            .iter()
            .try_fold(revision.content.len(), |total, existing| {
                total.checked_add(existing.content.len())
            })
            .ok_or_else(|| {
                RtfError::MalformedDocument("RTF aggregate revision size overflow".to_string())
            })?;
        if total > super::annotation::MAX_REVISION_TEXT_TOTAL_BYTES {
            return Err(RtfError::MalformedDocument(
                "RTF aggregate revision text exceeds the safety limit".to_string(),
            ));
        }
        self.revisions.push(revision);
        Ok(())
    }

    /// Remove all tracked changes while retaining the ordered author table.
    pub fn clear_revisions(&mut self) {
        self.revisions.clear();
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::{Alignment, ListFollow, ListJustification, ListLevelType, RevisionType, StyleType};

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
    fn parses_and_owns_shapes_and_shape_groups() {
        let rtf = r#"{\rtf1\ansi
            {\shp\shpleft10\shptop20\shpright310\shpbottom140
                \shprotation15\shpz4\shpwr4\shplockanchor{\*\shpinst
                    {\sp{\sn shapeType}{\sv 202}}
                    {\sp{\sn wzName}{\sv Owned Text Box}}
                    {\sp{\sn fBehindDocument}{\sv 1}}
                    {\sp{\sn fLockPosition}{\sv 1}}
                    {\sp{\sn fillType}{\sv 7}}
                    {\sp{\sn fillColor}{\sv 66051}}
                    {\sp{\sn fillBackColor}{\sv 263430}}
                    {\sp{\sn fillOpacity}{\sv 32768}}
                    {\sp{\sn fLine}{\sv 0}}
                    {\sp{\sn lineColor}{\sv 460809}}
                    {\sp{\sn lineWidth}{\sv 12700}}
                    {\sp{\sn futureOfficeArtProperty}{\sv retained}}}
                {\shptxt Hello \u20320?}}
            {\shpgrp\shpleft1\shptop2\shpright801\shpbottom602
                {\sp{\sn wzName}{\sv Owned Group}}
                {\shp\shpinst1\shpleft5\shptop6\shpwidth70\shpheight80\shpfblwtxt1}
                {\shp\shpinst3\shpleft15\shptop16\shpwidth90\shpheight100}
                {\shpgrp\shpleft100\shptop110\shpright400\shpbottom510
                    {\sp{\sn wzName}{\sv Owned Nested Group}}
                    {\shp\shpinst20\shpleft1\shptop2\shpwidth3\shpheight4}}}
        }"#;
        let doc = RtfDocument::parse(rtf).unwrap();

        let shape = &doc.shapes()[0];
        assert_eq!(shape.shape_type, crate::ShapeType::TextBox);
        assert_eq!(shape.geometry.x, 10);
        assert_eq!(shape.geometry.y, 20);
        assert_eq!(shape.geometry.width, 300);
        assert_eq!(shape.geometry.height, 120);
        assert_eq!(shape.geometry.rotation, 15);
        assert_eq!(shape.geometry.z_order, 4);
        assert_eq!(shape.text, "Hello 你");
        assert_eq!(shape.name, "Owned Text Box");
        assert!(shape.behind_doc);
        assert!(shape.locked);
        assert_eq!(shape.wrap_mode, crate::WrapMode::Tight);
        assert_eq!(shape.fill.fill_type, crate::FillType::Gradient);
        assert_eq!(shape.fill.color.raw(), 66_051);
        assert_eq!(shape.fill.color.red(), 1);
        assert_eq!(shape.fill.color.green(), 2);
        assert_eq!(shape.fill.color.blue(), 3);
        assert_eq!(shape.fill.color2.unwrap().raw(), 263_430);
        assert_eq!(shape.fill.opacity.raw(), 32_768);
        assert_eq!(shape.fill.opacity.as_fraction(), 0.5);
        assert!(!shape.line.visible);
        assert_eq!(shape.line.color.raw(), 460_809);
        assert_eq!(shape.line.width_emu, 12_700);
        assert!(shape.properties.iter().any(|property| {
            property.name == "futureOfficeArtProperty" && property.value == "retained"
        }));
        assert!(
            shape
                .properties
                .iter()
                .all(|property| matches!(property.name, Cow::Owned(_)))
        );
        assert!(matches!(shape.text, Cow::Owned(_)));

        let group = &doc.shape_groups()[0];
        assert_eq!(group.geometry, crate::ShapeGeometry::new(1, 2, 800, 600));
        assert_eq!(group.shapes.len(), 2);
        assert_eq!(group.shapes[0].shape_type, crate::ShapeType::Rectangle);
        assert!(group.shapes[0].behind_doc);
        assert_eq!(group.shapes[0].wrap_mode, crate::WrapMode::Behind);
        assert_eq!(group.shapes[1].shape_type, crate::ShapeType::Ellipse);
        assert_eq!(group.name, "Owned Group");
        assert!(matches!(group.name, Cow::Owned(_)));
        assert!(
            group
                .properties
                .iter()
                .all(|property| matches!(property.value, Cow::Owned(_)))
        );
        assert_eq!(group.groups().len(), 1);
        let nested = &group.groups()[0];
        assert_eq!(nested.name, "Owned Nested Group");
        assert_eq!(
            nested.geometry,
            crate::ShapeGeometry::new(100, 110, 300, 400)
        );
        assert_eq!(nested.shapes().len(), 1);
        assert_eq!(nested.shapes()[0].shape_type, crate::ShapeType::Line);
        assert!(matches!(nested.name, Cow::Owned(_)));
    }

    #[test]
    fn rejects_excessively_nested_shape_groups() {
        let mut rtf = String::from("{\\rtf1");
        for _ in 0..=64 {
            rtf.push_str("{\\shpgrp");
        }
        for _ in 0..=64 {
            rtf.push('}');
        }
        rtf.push('}');

        let error = match RtfDocument::parse(&rtf) {
            Ok(_) => panic!("excessive shape-group nesting should fail"),
            Err(error) => error,
        };
        assert!(
            matches!(error, RtfError::MalformedDocument(message) if message.contains("shape group nesting"))
        );
    }

    #[test]
    fn parses_real_background_shape_fixture_with_trailing_newline() {
        let rtf = include_str!("../../../test-data/rtf/background.rtf");
        let doc = RtfDocument::parse(rtf).unwrap();

        assert_eq!(doc.shapes().len(), 2);
        assert_eq!(doc.shapes()[0].geometry.x, 2633);
        assert_eq!(doc.shapes()[0].geometry.width, 2220);
        assert_eq!(doc.shapes()[0].fill.color.raw(), 5_880_731);
        assert_eq!(doc.shapes()[0].fill.fill_type, crate::FillType::Solid);
        assert_eq!(doc.shapes()[0].properties.len(), 4);
        assert_eq!(doc.shapes()[1].geometry.x, 488);
        assert_eq!(doc.shapes()[1].geometry.width, 1515);
        assert_eq!(doc.shapes()[1].fill.color.raw(), 5_066_944);
        assert_eq!(
            doc.text(),
            "First should be foreground, the second should be background.\n"
        );
    }

    #[test]
    fn preserves_real_watermark_office_art_properties() {
        let rtf = include_str!("../../../test-data/rtf/watermark.rtf");
        let doc = RtfDocument::parse(rtf).unwrap();

        assert_eq!(doc.shapes().len(), 3);
        let shape = &doc.shapes()[0];
        assert_eq!(shape.shape_type, crate::ShapeType::Custom(136));
        assert_eq!(shape.geometry.rotation, 315);
        assert_eq!(shape.fill.color.raw(), 6_108_695);
        assert_eq!(shape.fill.opacity.raw(), 32_768);
        assert!(!shape.line.visible);
        assert_eq!(shape.name, "PowerPlusWaterMarkObject142907");
        assert!(shape.behind_doc);
        assert_eq!(shape.property("gtextUNICODE"), Some("ASAP"));
    }

    #[test]
    fn parses_shape_from_ignorable_page_background_destination() {
        let rtf = include_str!("../../../test-data/rtf/page-background.rtf");
        let doc = RtfDocument::parse(rtf).unwrap();

        assert_eq!(doc.shapes().len(), 1);
        let shape = &doc.shapes()[0];
        assert_eq!(shape.shape_type, crate::ShapeType::Rectangle);
        assert_eq!(shape.fill.color.raw(), 5_296_274);
        assert!(shape.is_background);
        assert!(shape.behind_doc);
        assert_eq!(shape.property("bWMode"), Some("9"));
        assert!(!doc.text().contains("shapeType"));
        assert!(!doc.text().contains("fillColor"));
    }

    #[test]
    fn extracts_embedded_object_metadata_and_native_bytes_without_activation() {
        let rtf = r#"{\rtf1\ansi Before {\object\objemb\objw1440\objh720\objlock\objupdate\objsetsize
            {\*\objclass Package}{\*\objname Owned \u20320? Object}
            {\*\objdata
                01050000 02000000 08000000 5061636b61676500
                00000000 00000000 08000000 d0cf11e0a1b11ae1}
            {\result fallback {\pict\pngblip\picw10\pich20 89504e470d0a1a0a}}} After}"#;
        let doc = RtfDocument::parse(rtf).unwrap();

        assert_eq!(doc.text(), "Before  After");
        assert_eq!(doc.objects().len(), 1);
        let object = &doc.objects()[0];
        assert_eq!(object.kind, crate::ObjectKind::Embedded);
        assert_eq!(object.class_name, "Package");
        assert_eq!(object.name, "Owned 你 Object");
        assert_eq!(object.width, 1440);
        assert_eq!(object.height, 720);
        assert!(object.locked);
        assert!(object.update_requested);
        assert!(object.set_size);
        assert!(matches!(object.class_name, Cow::Owned(_)));
        assert_eq!(object.result_text, "fallback");
        assert_eq!(object.result_picture_indices, [0]);
        assert!(matches!(object.result_text, Cow::Owned(_)));
        assert_eq!(doc.pictures().len(), 1);
        assert_eq!(doc.pictures()[0].image_type, crate::ImageType::Png);
        assert_eq!(doc.pictures()[0].width, Some(10));
        assert_eq!(doc.pictures()[0].height, Some(20));

        let header = object.ole_header().unwrap();
        assert_eq!(header.ole_version, 0x501);
        assert_eq!(header.format_id, 2);
        assert_eq!(header.class_name, b"Package");
        assert!(header.is_compound_file());
    }

    #[test]
    fn rejects_invalid_embedded_object_hex() {
        let error = match RtfDocument::parse(r#"{\rtf1{\object{\*\objdata 0xz}}}"#) {
            Ok(_) => panic!("invalid object hex should fail"),
            Err(error) => error,
        };
        assert!(
            matches!(error, RtfError::MalformedDocument(message) if message.contains("non-hexadecimal"))
        );
    }

    #[test]
    fn preserves_exact_binary_picture_payload() {
        let payload = [
            0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, b'{', b'\\', b'}',
        ];
        let mut rtf = br"{\rtf1{\pict\pngblip\bin11 ".to_vec();
        rtf.extend_from_slice(&payload);
        rtf.extend_from_slice(b"}}");

        let doc = RtfDocument::parse_bytes(&rtf).unwrap();
        assert_eq!(doc.pictures().len(), 1);
        assert_eq!(doc.pictures()[0].image_type, crate::ImageType::Png);
        assert_eq!(doc.pictures()[0].data(), payload);
    }

    #[test]
    fn rejects_unclosed_document_and_destination_groups() {
        for rtf in [
            r#"{\rtf1 body"#,
            r#"{\rtf1{\*\unknown destination"#,
            r#"{\rtf1{\shp\shpleft1}"#,
            r#"{\rtf1{\shp{\sp{\sn shapeType}{\sv 1}"#,
        ] {
            assert!(matches!(
                RtfDocument::parse(rtf),
                Err(RtfError::UnexpectedEof)
            ));
        }
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
    fn parses_tracked_insertions_and_deletions_with_author_ranges() {
        let rtf = r#"{\rtf1\ansi
            {\*\revtbl {Unknown;}{Max \u20320?;}}
            Before {\deleted\revauthdel1\revdttmdel1199059860 old \u20320? text}
            and {\revised\revauth1\revdttm-1501115711 new text} after}"#;
        let doc = RtfDocument::parse(rtf).unwrap();
        let body = doc.text();
        assert!(!body.contains("old 你 text"));
        assert!(body.contains("and new text after"));
        assert_eq!(doc.revisions().len(), 2);

        let deletion = &doc.revisions()[0];
        assert_eq!(deletion.revision_type, RevisionType::Deletion);
        assert_eq!(deletion.id, 1);
        assert_eq!(deletion.author, "Max 你");
        assert_eq!(deletion.date.as_deref(), Some("1199059860"));
        assert_eq!(deletion.content, "old 你 text");
        assert_eq!(deletion.position, deletion.range_end);
        assert!(!body.contains(deletion.content.as_ref()));

        let insertion = &doc.revisions()[1];
        assert_eq!(insertion.revision_type, RevisionType::Insertion);
        assert_eq!(insertion.author, "Max 你");
        assert_eq!(insertion.date.as_deref(), Some("-1501115711"));
        assert_eq!(insertion.content, "new text");
        assert_eq!(
            body.get(insertion.position..insertion.range_end),
            Some(insertion.content.as_ref())
        );
    }

    #[test]
    fn revision_toggle_boundaries_flush_preceding_text() {
        let doc = RtfDocument::parse(
            r#"{\rtf1{\*\revtbl Unknown;}plain \revised\revauth0\revdttm1 changed\revised0 plain}"#,
        )
        .unwrap();
        assert_eq!(doc.text(), "plain changedplain");
        assert_eq!(doc.revisions().len(), 1);
        assert_eq!(doc.revisions()[0].content, "changed");
        assert_eq!(doc.revisions()[0].author, "Unknown");
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

    #[test]
    fn decodes_hex_escapes_after_declared_codepage() {
        let cyrillic =
            RtfDocument::parse(r#"{\rtf1\ansi\ansicpg1251 \'cf\'f0\'e8\'e2\'e5\'f2}"#).unwrap();
        assert_eq!(cyrillic.text(), "Привет");

        let japanese = RtfDocument::parse(r#"{\rtf1\ansi\ansicpg932 \'82\'a0\'82\'a2}"#).unwrap();
        assert_eq!(japanese.text(), "あい");
    }

    #[test]
    fn decodes_macintosh_and_exact_dos_character_sets() {
        let scoped = RtfDocument::parse(r#"{\rtf1\ansi \'80{\mac \'80}\'80}"#).unwrap();
        assert_eq!(scoped.text(), "€Ä€");

        let cp437 = RtfDocument::parse(r#"{\rtf1\pc \'9b}"#).unwrap();
        assert_eq!(cp437.text(), "¢");
        let cp850 = RtfDocument::parse(r#"{\rtf1\pca \'9b}"#).unwrap();
        assert_eq!(cp850.text(), "ø");
        let explicit_cp437 = RtfDocument::parse(r#"{\rtf1\ansi\ansicpg437 \'9b}"#).unwrap();
        assert_eq!(explicit_cp437.text(), "¢");
    }

    #[test]
    fn decodes_unescaped_legacy_bytes_and_semantic_control_symbols() {
        let mut bytes = br#"{\rtf1\ansi\ansicpg1252 "#.to_vec();
        bytes.push(0xE9);
        bytes.push(b'}');
        assert_eq!(RtfDocument::parse_bytes(&bytes).unwrap().text(), "é");

        let symbols = RtfDocument::parse(r#"{\rtf1\pc A\~B\-C\_D}"#).unwrap();
        assert_eq!(symbols.text(), "A\u{00A0}B\u{00AD}C\u{2011}D");
    }

    #[test]
    fn decodes_declared_codepage_in_semantic_destinations() {
        let doc = RtfDocument::parse(
            r#"{\rtf1\ansi\ansicpg1251{\info{\author \'cf\'f0\'e8\'e2\'e5\'f2\~X}}{\*\revtbl {\'cf\'f0\'e8\'e2\'e5\'f2;}}{\header \'cf\'f0\'e8\'e2\'e5\'f2\_X}Body{\footnote \'cf\'f0\'e8\'e2\'e5\'f2\~X}{\revised\revauth0 X}}"#,
        )
        .unwrap();

        assert_eq!(doc.info().author.as_deref(), Some("Привет\u{00A0}X"));
        assert_eq!(
            doc.sections()[0].headers_footers[0].text(),
            "Привет\u{2011}X"
        );
        assert_eq!(doc.notes()[0].content, "Привет\u{00A0}X");
        assert_eq!(doc.revisions()[0].author, "Привет");
    }
}
