use super::*;

/// Append one textbox story (main or header) to the text stream.
///
/// Per text box: its paragraphs (each `\r`-terminated, with `\n`/`\r`/`"\r\n"`
/// as input separators) plus a trailing CR; one story-final CR is included
/// in the returned story character count. Returns the story-relative start
/// CP of each text box and the total story length (a ccp value).

impl Writer {
    /// Create a new DOC writer
    pub fn new() -> Self {
        Self {
            paragraphs: Vec::new(),
            tables: Vec::new(),
            properties: HashMap::new(),
            header_even: None,
            header_odd: None,
            header_first: None,
            footer_even: None,
            footer_odd: None,
            footer_first: None,
            footnotes: Vec::new(),
            endnotes: Vec::new(),
            comments: Vec::new(),
            bookmarks: Vec::new(),
            smart_tags: Vec::new(),
            smart_tag_recognizer_ranges: Vec::new(),
            proofing_tables: ProofingTables::default(),
            associated_strings: DocumentAssociatedStrings::default(),
            saved_by_table: None,
            glossary_metadata: None,
            attached_glossary: None,
            section_formatting_revision: None,
            section_columns: None,
            section_right_to_left: false,
            section_text_flow: crate::TextFlow::default(),
            section_page_borders: None,
            numbering: NumberingWriter::new(),
            styles: Vec::new(),
            pictures: Vec::new(),
            shapes: Vec::new(),
            header_shapes: Vec::new(),
            header_pictures: Vec::new(),
            header_anchors: Vec::new(),
            next_shape_id: crate::writer::images::FIRST_PICTURE_SHAPE_ID,
            encryption: None,
            vba_project: None,
        }
    }

    /// Protect the generated document with a password-to-open profile.
    ///
    /// Validation is atomic: an invalid password or profile leaves any previous
    /// password setting unchanged.
    pub fn set_password(
        &mut self,
        password: impl Into<String>,
        profile: EncryptionProfile,
    ) -> Result<(), WriteError> {
        let password = Zeroizing::new(password.into());
        validate_writer_password(profile, password.as_str()).map_err(WriteError::InvalidData)?;
        self.encryption = Some(WriterEncryption { profile, password });
        Ok(())
    }

    /// Remove password-to-open protection and wipe the stored password.
    pub fn clear_password(&mut self) {
        self.encryption = None;
    }

    /// Return the configured password-to-open profile without exposing the password.
    pub fn encryption_profile(&self) -> Option<EncryptionProfile> {
        self.encryption.as_ref().map(|value| value.profile)
    }

    /// Configure a complete inert VBA project with safe default limits.
    pub fn set_vba(&mut self, project: litchi_vba::build::Project) -> Result<(), WriteError> {
        self.set_vba_with(project, &litchi_vba::Limits::default())
    }

    /// Configure a complete inert VBA project using explicit resource limits.
    ///
    /// Validation and serialization complete before writer state is changed.
    pub fn set_vba_with(
        &mut self,
        project: litchi_vba::build::Project,
        limits: &litchi_vba::Limits,
    ) -> Result<(), WriteError> {
        let payload = project.finish(limits)?;
        self.put_vba(payload);
        Ok(())
    }

    /// Configure an already validated and serialized inert VBA project.
    ///
    /// Import standalone CFB bytes through [`litchi_vba::Payload::read`] first.
    pub fn put_vba(&mut self, payload: litchi_vba::Payload) {
        self.vba_project = Some(payload);
    }

    /// Remove the configured VBA project storage.
    pub fn clear_vba(&mut self) {
        self.vba_project = None;
    }

    /// Whether a complete VBA project is configured for output.
    pub fn has_vba(&self) -> bool {
        self.vba_project.is_some()
    }

    /// Insert or replace a spelling or grammar proofing-state table.
    ///
    /// Character positions use the concatenated DOC document-part coordinate
    /// space. The final CP ceiling is validated when output is generated.
    pub fn set_proofing_table(&mut self, table: ProofingStateTable) -> Option<ProofingStateTable> {
        self.proofing_tables.set(table)
    }

    /// Replace both optional proofing tables.
    pub fn set_proofing_tables(&mut self, tables: ProofingTables) {
        self.proofing_tables = tables;
    }

    /// Access one configured proofing table.
    pub fn proofing_table(&self, feature: ProofingFeature) -> Option<&ProofingStateTable> {
        self.proofing_tables.get(feature)
    }

    /// Remove and return one configured proofing table.
    pub fn clear_proofing_table(&mut self, feature: ProofingFeature) -> Option<ProofingStateTable> {
        self.proofing_tables.remove(feature)
    }

    /// Replace all 18 associated-document string slots.
    pub fn set_associated_strings(&mut self, strings: DocumentAssociatedStrings) {
        self.associated_strings = strings;
    }

    /// Access the associated-document string table that will be written.
    pub fn associated_strings(&self) -> &DocumentAssociatedStrings {
        &self.associated_strings
    }

    /// Replace one associated-document string slot atomically.
    pub fn set_associated_string(
        &mut self,
        slot: AssociatedStringSlot,
        value: impl Into<String>,
    ) -> Result<String, WriteError> {
        self.associated_strings
            .set(slot, value)
            .map_err(|error| WriteError::InvalidData(error.to_string()))
    }

    /// Reset all associated-document string slots to empty strings.
    ///
    /// The mandatory `SttbfAssoc` structure is still emitted.
    pub fn reset_associated_strings(&mut self) {
        self.associated_strings = DocumentAssociatedStrings::default();
    }

    /// Configure the optional Word 97/2000 save-history table.
    pub fn set_saved_by_table(&mut self, table: SavedByTable) -> Option<SavedByTable> {
        self.saved_by_table.replace(table)
    }

    /// Access the configured save-history table.
    pub fn saved_by_table(&self) -> Option<&SavedByTable> {
        self.saved_by_table.as_ref()
    }

    /// Remove and return the configured save-history table.
    pub fn clear_saved_by_table(&mut self) -> Option<SavedByTable> {
        self.saved_by_table.take()
    }

    /// Configure this output as a glossary-only DOC.
    ///
    /// Item ranges use main-story UTF-16 character positions and may cover
    /// formatted text, tables, drawings, or pictures. The metadata's `ccpText`
    /// is checked against the generated main story before output is modified.
    pub fn set_glossary_metadata(
        &mut self,
        metadata: GlossaryMetadata,
    ) -> Option<GlossaryMetadata> {
        self.glossary_metadata.replace(metadata)
    }

    /// Access the configured glossary-only metadata.
    pub fn glossary_metadata(&self) -> Option<&GlossaryMetadata> {
        self.glossary_metadata.as_ref()
    }

    /// Return this writer to ordinary-document output.
    pub fn clear_glossary_metadata(&mut self) -> Option<GlossaryMetadata> {
        self.glossary_metadata.take()
    }

    /// Attach a distinct glossary-only document to this template.
    ///
    /// The attached writer must have [`Writer::set_glossary_metadata`]
    /// configured. Its main story becomes the template's AutoText story.
    ///
    /// # Errors
    ///
    /// Returns an error for nested or independently encrypted glossary
    /// documents and independent VBA projects. Those configurations cannot be
    /// represented by the shared DOC stream topology.
    pub fn set_attached_glossary(
        &mut self,
        glossary: Writer,
    ) -> Result<Option<Writer>, WriteError> {
        glossary.validate_as_attached_glossary()?;
        Ok(self
            .attached_glossary
            .replace(Box::new(glossary))
            .map(|previous| *previous))
    }

    /// Access the attached glossary writer.
    pub fn attached_glossary(&self) -> Option<&Writer> {
        self.attached_glossary.as_deref()
    }

    /// Mutably access the attached glossary writer.
    pub fn attached_glossary_mut(&mut self) -> Option<&mut Writer> {
        self.attached_glossary.as_deref_mut()
    }

    /// Remove and return the attached glossary writer.
    pub fn clear_attached_glossary(&mut self) -> Option<Writer> {
        self.attached_glossary.take().map(|glossary| *glossary)
    }
}

impl Writer {
    /// Add a custom paragraph, character, table, or numbering style.
    ///
    /// Custom styles occupy consecutive indices beginning at 15. The returned
    /// index can be used by the corresponding formatting properties.
    pub fn add_style(
        &mut self,
        style: crate::writer::stylesheet::StyleDefinition,
    ) -> Result<u16, WriteError> {
        let index = 15usize
            .checked_add(self.styles.len())
            .and_then(|index| u16::try_from(index).ok())
            .filter(|index| *index <= 0x0FFC)
            .ok_or_else(|| {
                WriteError::InvalidData("DOC stylesheet exceeds 4093 style slots".to_string())
            })?;
        self.styles.push(style);
        Ok(index)
    }

    /// Add a paragraph with plain text
    ///
    /// # Arguments
    ///
    /// * `text` - Paragraph text
    ///
    /// # Returns
    ///
    /// * `Result<(), WriteError>` - Success or error
    pub fn add_paragraph(&mut self, text: &str) -> Result<(), WriteError> {
        self.paragraphs.push(WritableParagraph {
            runs: vec![TextRun {
                text: text.to_string(),
                formatting: CharacterFormatting::default(),
                picture_index: None,
                shape_index: None,
            }],
            formatting: ParagraphFormatting::default(),
        });
        Ok(())
    }

    /// Add a paragraph with paragraph formatting (default character formatting)
    pub fn add_formatted_paragraph(
        &mut self,
        text: &str,
        para_fmt: ParagraphFormatting,
    ) -> Result<(), WriteError> {
        self.add_paragraph_with_format(text, CharacterFormatting::default(), para_fmt)
    }

    /// Add a paragraph with formatting
    ///
    /// # Arguments
    ///
    /// * `text` - Paragraph text
    /// * `char_fmt` - Character formatting
    /// * `para_fmt` - Paragraph formatting
    pub fn add_paragraph_with_format(
        &mut self,
        text: &str,
        char_fmt: CharacterFormatting,
        para_fmt: ParagraphFormatting,
    ) -> Result<(), WriteError> {
        self.paragraphs.push(WritableParagraph {
            runs: vec![TextRun {
                text: text.to_string(),
                formatting: char_fmt,
                picture_index: None,
                shape_index: None,
            }],
            formatting: para_fmt,
        });
        Ok(())
    }

    /// Add a paragraph composed of multiple runs (rich text)
    ///
    /// Each tuple is (text, character formatting) and the whole paragraph shares the
    /// given paragraph formatting.
    pub fn add_paragraph_runs(
        &mut self,
        runs: Vec<(String, CharacterFormatting)>,
        para_fmt: ParagraphFormatting,
    ) -> Result<(), WriteError> {
        if runs.is_empty() {
            return self.add_paragraph_with_format("", CharacterFormatting::default(), para_fmt);
        }
        let mut wruns = Vec::with_capacity(runs.len());
        for (text, formatting) in runs {
            wruns.push(TextRun {
                text,
                formatting,
                picture_index: None,
                shape_index: None,
            });
        }
        self.paragraphs.push(WritableParagraph {
            runs: wruns,
            formatting: para_fmt,
        });
        Ok(())
    }

    /// Insert an inline picture as its own paragraph.
    ///
    /// The picture is written as a single 0x0001 picture character with
    /// sprmCFSpec and sprmCPicLocation applied ([MS-DOC] 1.3); the character
    /// points to an OfficeArtWordDrawing block (PICF + OfficeArtSpContainer +
    /// OfficeArtFBSE with an embedded BLIP) in the Data stream. The image
    /// bytes are stored verbatim — no re-encoding is performed.
    pub fn insert_picture(
        &mut self,
        picture: crate::writer::images::Picture,
    ) -> Result<(), WriteError> {
        self.insert_picture_run(picture, None, "\u{0001}")
    }

    /// Insert a floating picture anchored to its own paragraph.
    ///
    /// The anchor is a single 0x0008 character with sprmCFSpec and
    /// sprmCPicLocation applied ([MS-DOC] 1.3). The picture data is stored
    /// like an inline picture's, and the anchor character position is
    /// recorded in the Main Document's PlcfSpa together with an
    /// OfficeArtContent drawing group (fcDggInfo) holding the picture-frame
    /// shape, so readers can resolve the anchor to position and image.
    pub fn insert_floating_picture(
        &mut self,
        picture: crate::writer::images::Picture,
        position: crate::writer::images::FloatingPosition,
    ) -> Result<(), WriteError> {
        self.insert_picture_run(picture, Some(position), "\u{0008}")
    }

    /// Shared tail of `insert_picture`/`insert_floating_picture`: queue the
    /// picture and append a single-character anchor paragraph.
    fn insert_picture_run(
        &mut self,
        picture: crate::writer::images::Picture,
        floating: Option<crate::writer::images::FloatingPosition>,
        anchor: &str,
    ) -> Result<(), WriteError> {
        let picture_index = u32::try_from(self.pictures.len()).map_err(|_| {
            WriteError::InvalidData("DOC picture count exceeds the 32-bit range".to_string())
        })?;
        let shape_id = self.allocate_shape_id()?;
        self.pictures.push(WriterPicture {
            picture,
            shape_id,
            floating,
        });
        self.paragraphs.push(WritableParagraph {
            runs: vec![TextRun {
                text: anchor.to_string(),
                formatting: CharacterFormatting {
                    special: Some(true),
                    ..CharacterFormatting::default()
                },
                picture_index: Some(picture_index),
                shape_index: None,
            }],
            formatting: ParagraphFormatting::default(),
        });
        Ok(())
    }

    /// Shared tail of `insert_floating_shape`/`insert_floating_text_box`.
    fn insert_shape_run(
        &mut self,
        shape: crate::writer::shapes::Shape,
        position: crate::writer::images::FloatingPosition,
        text: Option<String>,
    ) -> Result<(), WriteError> {
        let shape_index = u32::try_from(self.shapes.len()).map_err(|_| {
            WriteError::InvalidData("DOC shape count exceeds the 32-bit range".to_string())
        })?;
        let shape_id = self.allocate_shape_id()?;
        self.shapes.push(WriterShape {
            shape,
            shape_id,
            position,
            text,
        });
        self.paragraphs.push(WritableParagraph {
            runs: vec![TextRun {
                text: "\u{0008}".to_string(),
                formatting: CharacterFormatting {
                    special: Some(true),
                    ..CharacterFormatting::default()
                },
                picture_index: None,
                shape_index: Some(shape_index),
            }],
            formatting: ParagraphFormatting::default(),
        });
        Ok(())
    }

    /// Insert a floating primitive drawing shape anchored to its own paragraph.
    ///
    /// The anchor is a single 0x0008 character with sprmCFSpec applied, and
    /// the shape is emitted into the document's drawing group (fcDggInfo
    /// OfficeArtContent) with its position recorded in the Main Document's
    /// PlcfSpa — the same mechanism as floating pictures ([MS-DOC] 1.3).
    ///
    /// Shape text (text boxes) is not supported; see
    /// [`crate::writer::shapes::Shape`].
    pub fn insert_floating_shape(
        &mut self,
        shape: crate::writer::shapes::Shape,
        position: crate::writer::images::FloatingPosition,
    ) -> Result<(), WriteError> {
        self.insert_shape_run(shape, position, None)
    }

    /// Insert a floating text box anchored to its own paragraph.
    ///
    /// Anchoring and positioning work like [`Self::insert_floating_shape`],
    /// but the shape is emitted as an msosptTextBox with an
    /// OfficeArtClientTextbox record whose TXID links it to an entry in the
    /// textbox story ([MS-DOC] PlcftxbxTxt). The story text is appended to
    /// the WordDocument stream after the endnote story and counted in
    /// ccpTxbx. The text is plain: `\n` (or `\r` / `"\r\n"`) separates
    /// paragraphs; no character or paragraph formatting is applied.
    pub fn insert_floating_text_box(
        &mut self,
        shape: crate::writer::shapes::Shape,
        position: crate::writer::images::FloatingPosition,
        text: impl Into<String>,
    ) -> Result<(), WriteError> {
        self.insert_shape_run(shape, position, Some(text.into()))
    }

    /// Insert a text box anchored in the given header.
    ///
    /// The anchor is a single 0x0008 paragraph appended to that header's
    /// paragraphs (created when absent); position and wrapping work like
    /// [`Self::insert_floating_text_box`], but the shape position is recorded
    /// in the Header Document's PlcfSpaHdr, the text goes to the header
    /// textbox story (counted in ccpHdrTxbx, linked through PlcfHdrtxbxTxt),
    /// and the shape joins the Header Document drawing of the fcDggInfo
    /// OfficeArtContent. Header floating items use their own shape-id cluster
    /// starting at 2049, so they never collide with main-story shapes.
    ///
    /// Set or replace header paragraphs BEFORE calling this method: the
    /// anchor lives in paragraphs this method appends, and replacing the
    /// header's paragraph list afterwards drops the anchor.
    pub fn insert_header_text_box(
        &mut self,
        kind: HeaderKind,
        shape: crate::writer::shapes::Shape,
        position: crate::writer::images::FloatingPosition,
        text: impl Into<String>,
    ) -> Result<(), WriteError> {
        let shape_id = self.allocate_header_shape_id()?;
        let item_index = u32::try_from(self.header_shapes.len()).map_err(|_| {
            WriteError::InvalidData(
                "DOC header text box count exceeds the 32-bit range".to_string(),
            )
        })?;
        self.header_shapes.push(WriterShape {
            shape,
            shape_id,
            position,
            text: Some(text.into()),
        });
        self.append_header_anchor(kind, FloatingAnchorKind::Shape(item_index))
    }

    /// Insert a floating picture anchored in the given header (the classic
    /// letterhead logo / watermark pattern).
    ///
    /// Anchoring works like [`Self::insert_header_text_box`]: the picture is
    /// written as a PICF block with an embedded BLIP in the Data stream
    /// (bytes stored verbatim), referenced by sprmCPicLocation on the 0x0008
    /// anchor character, positioned through the PlcfSpaHdr, and rendered as a
    /// picture-frame shape in the Header Document drawing.
    pub fn insert_header_picture(
        &mut self,
        kind: HeaderKind,
        picture: crate::writer::images::Picture,
        position: crate::writer::images::FloatingPosition,
    ) -> Result<(), WriteError> {
        let shape_id = self.allocate_header_shape_id()?;
        let item_index = u32::try_from(self.header_pictures.len()).map_err(|_| {
            WriteError::InvalidData("DOC header picture count exceeds the 32-bit range".to_string())
        })?;
        self.header_pictures.push(WriterPicture {
            picture,
            shape_id,
            floating: Some(position),
        });
        self.append_header_anchor(kind, FloatingAnchorKind::Picture(item_index))
    }

    /// Allocate the next header-drawing shape id from the header cluster.
    fn allocate_header_shape_id(&mut self) -> Result<u32, WriteError> {
        let count = self.header_shapes.len() + self.header_pictures.len();
        let index = u32::try_from(count).map_err(|_| {
            WriteError::InvalidData(
                "DOC header floating item count exceeds the 32-bit range".to_string(),
            )
        })?;
        Ok(crate::writer::images::HEADER_FIRST_SHAPE_ID + index)
    }

    /// Append a 0x0008 anchor paragraph to the given header and record it.
    fn append_header_anchor(
        &mut self,
        kind: HeaderKind,
        anchor_kind: FloatingAnchorKind,
    ) -> Result<(), WriteError> {
        let paragraphs = match kind {
            HeaderKind::Odd => &mut self.header_odd,
            HeaderKind::Even => &mut self.header_even,
            HeaderKind::FirstPage => &mut self.header_first,
        };
        let paragraphs = paragraphs.get_or_insert_with(Vec::new);
        let paragraph_index = paragraphs.len();
        paragraphs.push(HeaderFooterParagraph::from_runs(
            vec![(
                "\u{0008}".to_string(),
                CharacterFormatting {
                    special: Some(true),
                    ..CharacterFormatting::default()
                },
            )],
            ParagraphFormatting::default(),
        ));
        self.header_anchors.push(HeaderAnchor {
            slot: kind.slot(),
            paragraph_index,
            kind: anchor_kind,
        });
        Ok(())
    }

    /// Allocate the next shape id from the sequence shared by pictures and
    /// drawing shapes (group shape ids start one below the first picture id).
    fn allocate_shape_id(&mut self) -> Result<u32, WriteError> {
        let shape_id = self.next_shape_id;
        self.next_shape_id = self
            .next_shape_id
            .checked_add(1)
            .ok_or_else(|| WriteError::InvalidData("DOC shape ids exhausted".to_string()))?;
        Ok(shape_id)
    }

    /// Add a hyperlink paragraph using Word field codes (HYPERLINK)
    ///
    /// This creates a field sequence:
    /// - 0x0013 (field begin, fSpec=1)
    /// - Instruction text: `HYPERLINK "url"` (field-vanished)
    /// - 0x0014 (field separator, fSpec=1)
    /// - Display text
    /// - 0x0015 (field end, fSpec=1)
    ///
    /// # Arguments
    /// - `display_text` - Visible link text shown in the document
    /// - `url` - Target URL for the hyperlink (quotes will be escaped)
    /// - `para_fmt` - Paragraph formatting to apply to this paragraph
    pub fn add_hyperlink(
        &mut self,
        display_text: &str,
        url: &str,
        mut para_fmt: ParagraphFormatting,
    ) -> Result<(), WriteError> {
        // Stage: Implementing hyperlinks using field codes

        // Escape quotes inside URL by doubling them per Word field syntax
        let escaped = url.replace('"', "\"\"");
        let instr = format!("HYPERLINK \"{}\"", escaped);

        // Default hyperlink visual style (blue + single underline)
        let link_fmt = CharacterFormatting {
            underline: Some(true),
            color: Some((0x00, 0x00, 0xFF)),
            ..CharacterFormatting::default()
        };

        // Field begin/separator/end special chars
        let spec_fmt = CharacterFormatting {
            special: Some(true),
            ..CharacterFormatting::default()
        };

        // Field instruction should be hidden (vanished) but not special
        let instr_fmt = CharacterFormatting {
            field_vanish: Some(true),
            ..CharacterFormatting::default()
        };

        let runs = vec![
            ("\u{0013}".to_string(), spec_fmt.clone()), // fldBegin
            (instr, instr_fmt),                         // instruction text (hidden)
            ("\u{0014}".to_string(), spec_fmt.clone()), // fldSep
            (display_text.to_string(), link_fmt),       // display text
            ("\u{0015}".to_string(), spec_fmt),         // fldEnd
        ];

        // Keep consistent paragraph spacing defaults for hyperlink paragraph (no auto spacing)
        if para_fmt.space_before_auto.is_none() {
            para_fmt.space_before_auto = Some(false);
        }
        if para_fmt.space_after_auto.is_none() {
            para_fmt.space_after_auto = Some(false);
        }

        self.add_paragraph_runs(runs, para_fmt)
    }

    /// Set the odd-page header text (HeaderStories index 7)
    pub fn set_odd_header(&mut self, text: &str) {
        self.header_odd = Some(vec![HeaderFooterParagraph::plain(text)]);
    }
    /// Set the even-page header text (HeaderStories index 6)
    pub fn set_even_header(&mut self, text: &str) {
        self.header_even = Some(vec![HeaderFooterParagraph::plain(text)]);
    }
    /// Set the first-page header text (HeaderStories index 10)
    pub fn set_first_header(&mut self, text: &str) {
        self.header_first = Some(vec![HeaderFooterParagraph::plain(text)]);
    }
    /// Set the odd-page footer text (HeaderStories index 9)
    pub fn set_odd_footer(&mut self, text: &str) {
        self.footer_odd = Some(vec![HeaderFooterParagraph::plain(text)]);
    }
    /// Set the even-page footer text (HeaderStories index 8)
    pub fn set_even_footer(&mut self, text: &str) {
        self.footer_even = Some(vec![HeaderFooterParagraph::plain(text)]);
    }
    /// Set the first-page footer text (HeaderStories index 11)
    pub fn set_first_footer(&mut self, text: &str) {
        self.footer_first = Some(vec![HeaderFooterParagraph::plain(text)]);
    }

    /// Set formatted odd-page header paragraphs (HeaderStories index 7).
    pub fn set_odd_header_paragraphs(
        &mut self,
        paragraphs: Vec<HeaderFooterParagraph>,
    ) -> Result<(), WriteError> {
        validate_header_footer_paragraphs(&paragraphs)?;
        self.header_odd = Some(paragraphs);
        Ok(())
    }

    /// Set formatted even-page header paragraphs (HeaderStories index 6).
    pub fn set_even_header_paragraphs(
        &mut self,
        paragraphs: Vec<HeaderFooterParagraph>,
    ) -> Result<(), WriteError> {
        validate_header_footer_paragraphs(&paragraphs)?;
        self.header_even = Some(paragraphs);
        Ok(())
    }

    /// Set formatted first-page header paragraphs (HeaderStories index 10).
    pub fn set_first_header_paragraphs(
        &mut self,
        paragraphs: Vec<HeaderFooterParagraph>,
    ) -> Result<(), WriteError> {
        validate_header_footer_paragraphs(&paragraphs)?;
        self.header_first = Some(paragraphs);
        Ok(())
    }

    /// Set formatted odd-page footer paragraphs (HeaderStories index 9).
    pub fn set_odd_footer_paragraphs(
        &mut self,
        paragraphs: Vec<HeaderFooterParagraph>,
    ) -> Result<(), WriteError> {
        validate_header_footer_paragraphs(&paragraphs)?;
        self.footer_odd = Some(paragraphs);
        Ok(())
    }

    /// Set formatted even-page footer paragraphs (HeaderStories index 8).
    pub fn set_even_footer_paragraphs(
        &mut self,
        paragraphs: Vec<HeaderFooterParagraph>,
    ) -> Result<(), WriteError> {
        validate_header_footer_paragraphs(&paragraphs)?;
        self.footer_even = Some(paragraphs);
        Ok(())
    }

    /// Set formatted first-page footer paragraphs (HeaderStories index 11).
    pub fn set_first_footer_paragraphs(
        &mut self,
        paragraphs: Vec<HeaderFooterParagraph>,
    ) -> Result<(), WriteError> {
        validate_header_footer_paragraphs(&paragraphs)?;
        self.footer_first = Some(paragraphs);
        Ok(())
    }

    /// Add a footnote to the document.
    ///
    /// The `ref_position` in `FootnoteEntry` is the character position
    /// in the main document where the footnote reference marker appears.
    pub fn add_footnote(&mut self, entry: FootnoteEntry) {
        self.footnotes.push(entry);
    }

    /// Add an endnote to the document.
    pub fn add_endnote(&mut self, entry: FootnoteEntry) {
        self.endnotes.push(entry);
    }

    /// Add a point or ranged comment to the document.
    pub fn add_comment(&mut self, entry: CommentEntry) {
        self.comments.push(entry);
    }

    /// Add a standard bookmark to the document.
    pub fn add_bookmark(&mut self, entry: BookmarkEntry) {
        self.bookmarks.push(entry);
    }

    /// Add an inert smart-tag bookmark and property bag.
    pub fn add_smart_tag(&mut self, entry: SmartTagEntry) {
        self.smart_tags.push(entry);
    }

    /// Add one contiguous smart-tag recognizer-state range.
    ///
    /// Ranges are serialized in insertion order and must form a contiguous CP
    /// sequence when the document is saved.
    pub fn add_smart_tag_recognizer_range(&mut self, range: SmartTagRecognizerRange) {
        self.smart_tag_recognizer_ranges.push(range);
    }

    /// Add a list structure definition.
    pub fn add_list(&mut self, list: ListStructure) {
        self.numbering.add_list(list);
    }

    /// Add a list format override.
    pub fn add_list_override(&mut self, lfo: ListFormatOverride) {
        self.numbering.add_override(lfo);
    }

    /// Set names parallel to the document's list definitions.
    pub fn set_list_names(&mut self, table: ListNamesTable) {
        self.numbering.set_list_names(table);
    }

    /// Set template codes parallel to the document's list definitions.
    pub fn set_list_templates(&mut self, table: ListTemplateTable) {
        self.numbering.set_list_templates(table);
    }
}

impl Writer {
    pub fn add_table(&mut self, rows: usize, cols: usize) -> Result<usize, WriteError> {
        if rows == 0 || cols == 0 {
            return Err(WriteError::InvalidData(
                "Table must have at least 1 row and 1 column".to_string(),
            ));
        }
        if cols > 63 {
            return Err(WriteError::InvalidData(
                "DOC table rows cannot exceed 63 cells".to_string(),
            ));
        }

        let mut table = WritableTable { rows: Vec::new() };

        for _ in 0..rows {
            let mut row = TableRow {
                cells: Vec::new(),
                formatting: crate::writer::tap::TableRow {
                    cells: Vec::with_capacity(cols),
                    ..crate::writer::tap::TableRow::default()
                },
            };
            for _ in 0..cols {
                row.cells.push(TableCell {
                    paragraphs: vec![WritableParagraph {
                        runs: vec![TextRun {
                            text: String::new(),
                            formatting: CharacterFormatting::default(),
                            picture_index: None,
                            shape_index: None,
                        }],
                        formatting: ParagraphFormatting::default(),
                    }],
                });
            }
            for index in 0..cols {
                const DEFAULT_TABLE_WIDTH: u32 = 8640;
                let left = DEFAULT_TABLE_WIDTH * index as u32 / cols as u32;
                let right = DEFAULT_TABLE_WIDTH * (index + 1) as u32 / cols as u32;
                row.formatting.cells.push(crate::writer::tap::TableCell {
                    width: (right - left) as u16,
                    merged: false,
                    ..crate::writer::tap::TableCell::default()
                });
            }
            table.rows.push(row);
        }

        let index = self.tables.len();
        self.tables.push(table);
        Ok(index)
    }

    /// Mark the document section's properties as a tracked formatting change.
    ///
    /// The legacy writer currently emits one section spanning the document, so
    /// this revision applies to that complete section.
    pub fn set_section_formatting_revision(&mut self, revision: FormattingRevision) {
        self.section_formatting_revision = Some(revision);
    }

    /// Set validated column geometry for the writer's single section.
    pub fn set_section_columns(
        &mut self,
        columns: crate::section::columns::Layout,
    ) -> Result<(), WriteError> {
        columns
            .validate()
            .map_err(|error| WriteError::InvalidData(error.to_string()))?;
        self.section_columns = Some(columns);
        Ok(())
    }

    /// Return explicit section column geometry, or `None` for the file-format default.
    pub fn section_columns(&self) -> Option<&crate::section::columns::Layout> {
        self.section_columns.as_ref()
    }

    /// Remove the explicit column override and restore the single-column default.
    pub fn clear_section_columns(&mut self) {
        self.section_columns = None;
    }

    /// Select left-to-right or right-to-left section column population order.
    pub fn set_section_right_to_left(&mut self, value: bool) {
        self.section_right_to_left = value;
    }

    pub fn section_right_to_left(&self) -> bool {
        self.section_right_to_left
    }

    /// Set the section-wide text-flow mode.
    pub fn set_section_text_flow(&mut self, value: crate::TextFlow) {
        self.section_text_flow = value;
    }

    pub fn section_text_flow(&self) -> crate::TextFlow {
        self.section_text_flow
    }

    /// Set validated page borders for the writer's single section.
    pub fn set_section_page_borders(
        &mut self,
        borders: crate::section::borders::Borders,
    ) -> Result<(), WriteError> {
        borders
            .validate()
            .map_err(|error| WriteError::InvalidData(error.to_string()))?;
        self.section_page_borders = Some(borders);
        Ok(())
    }

    /// Return explicit page borders, or `None` for the file-format default.
    pub fn section_page_borders(&self) -> Option<&crate::section::borders::Borders> {
        self.section_page_borders.as_ref()
    }

    /// Remove all explicit page-border edges and placement controls.
    pub fn clear_section_page_borders(&mut self) {
        self.section_page_borders = None;
    }

    /// Set text in a table cell
    ///
    /// # Arguments
    ///
    /// * `table_idx` - Table index
    /// * `row` - Row index (0-based)
    /// * `col` - Column index (0-based)
    /// * `text` - Cell text
    pub fn set_table_cell_text(
        &mut self,
        table_idx: usize,
        row: usize,
        col: usize,
        text: &str,
    ) -> Result<(), WriteError> {
        self.set_table_cell_paragraph_runs(
            table_idx,
            row,
            col,
            vec![(text.to_string(), CharacterFormatting::default())],
            ParagraphFormatting::default(),
        )
    }

    /// Replace a table cell with one paragraph composed of formatted runs.
    pub fn set_table_cell_paragraph_runs(
        &mut self,
        table_idx: usize,
        row: usize,
        col: usize,
        runs: Vec<(String, CharacterFormatting)>,
        formatting: ParagraphFormatting,
    ) -> Result<(), WriteError> {
        let paragraph = writable_paragraph_from_runs(runs, formatting);
        self.table_cell_mut(table_idx, row, col)?.paragraphs = vec![paragraph];
        Ok(())
    }

    /// Append a paragraph composed of formatted runs to a table cell.
    pub fn append_table_cell_paragraph_runs(
        &mut self,
        table_idx: usize,
        row: usize,
        col: usize,
        runs: Vec<(String, CharacterFormatting)>,
        formatting: ParagraphFormatting,
    ) -> Result<(), WriteError> {
        let paragraph = writable_paragraph_from_runs(runs, formatting);
        self.table_cell_mut(table_idx, row, col)?
            .paragraphs
            .push(paragraph);
        Ok(())
    }

    fn table_cell_mut(
        &mut self,
        table_idx: usize,
        row: usize,
        col: usize,
    ) -> Result<&mut TableCell, WriteError> {
        let table = self
            .tables
            .get_mut(table_idx)
            .ok_or_else(|| WriteError::InvalidData(format!("Table {} not found", table_idx)))?;

        let row_data = table
            .rows
            .get_mut(row)
            .ok_or_else(|| WriteError::InvalidData(format!("Row {} not found", row)))?;

        let cell = row_data
            .cells
            .get_mut(col)
            .ok_or_else(|| WriteError::InvalidData(format!("Column {} not found", col)))?;
        Ok(cell)
    }

    /// Set the widths, horizontal merges, height, and header state for a table row.
    pub fn set_table_row_formatting(
        &mut self,
        table_idx: usize,
        row: usize,
        formatting: crate::writer::tap::TableRow,
    ) -> Result<(), WriteError> {
        let table = self
            .tables
            .get_mut(table_idx)
            .ok_or_else(|| WriteError::InvalidData(format!("Table {table_idx} not found")))?;
        let row_data = table
            .rows
            .get_mut(row)
            .ok_or_else(|| WriteError::InvalidData(format!("Row {row} not found")))?;
        if formatting.cells.len() != row_data.cells.len() {
            return Err(WriteError::InvalidData(format!(
                "Row {row} formatting has {} cells but the row contains {}",
                formatting.cells.len(),
                row_data.cells.len()
            )));
        }
        crate::writer::tap::generate_row_sprms(&formatting)
            .map_err(|error| WriteError::InvalidData(error.to_string()))?;
        row_data.formatting = formatting;
        Ok(())
    }

    /// Set a document property
    ///
    /// # Arguments
    ///
    /// * `name` - Property name (e.g., "Title", "Author", "Subject")
    /// * `value` - Property value
    pub fn set_property(&mut self, name: &str, value: &str) {
        self.properties.insert(name.to_string(), value.to_string());
    }
}

impl Default for Writer {
    fn default() -> Self {
        Self::new()
    }
}

// Implementation deferred - DOC binary format functions:
// These would be needed for full DOC file generation:
// - FIB (File Information Block) generation
// - Piece table builder for text storage
// - SPRM generation for CHP (Character Properties)
// - SPRM generation for PAP (Paragraph Properties)
// - FKP (Formatted Disk Page) builder
// - TAP (Table Properties) builder
//
// Recommendation: Use the DOCX writer (fully implemented) for production use.
