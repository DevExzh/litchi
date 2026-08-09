use super::prelude::*;

impl Document {
    /// Get image binary data for an embedded image.
    ///
    /// This method extracts the image data from the `WordDocument` stream.
    /// The data is returned as a `Cow` to minimize copying when possible.
    ///
    /// # Arguments
    ///
    /// * `image` - Reference to an Image obtained from `Run::image()`
    ///
    /// # Example
    ///
    /// ```rust,ignore
    /// for para in doc.paragraphs()? {
    ///     for run in para.runs()? {
    ///         if let Some(img) = run.image() {
    ///             let data = doc.image_data(img)?;
    ///             let pic_type = img.picture_type(&doc.word_document())?;
    ///             // Process image data...
    ///         }
    ///     }
    /// }
    /// ```
    pub fn image_data(
        &self,
        image: &crate::image::Image,
    ) -> std::result::Result<litchi_odraw::image::File<'_>, crate::image::ImageError> {
        // Use the appropriate stream based on pic_offset
        let data_stream = self.get_data_stream(image.pic_offset()).ok_or(
            crate::image::ImageError::InvalidPicOffset(image.pic_offset()),
        )?;
        let word_document = self.word_document();
        image.data(data_stream, word_document)
    }

    /// Get a reference to the `WordDocument` stream.
    ///
    /// This is useful for low-level image operations.
    #[inline]
    #[must_use]
    pub fn word_document(&self) -> &[u8] {
        &self.word_document
    }

    /// Get the appropriate stream for picture data based on `pic_offset`.
    ///
    /// According to Apache POI's `PicturesTable.getData()`:
    /// - If Data stream exists and `pic_offset` < `data_stream.len()`, use Data stream
    /// - Otherwise use `WordDocument` stream
    ///
    /// This is because pictures are typically stored in the Data stream,
    /// not the `WordDocument` stream.
    pub(super) fn get_data_stream(&self, offset: u32) -> Option<&[u8]> {
        if let Some(data_stream) = &self.data_stream
            && (offset as usize) < data_stream.len()
        {
            return Some(data_stream.as_slice());
        }
        None
    }

    /// Get a reference to the Data stream (if available).
    ///
    /// The Data stream contains embedded pictures and OLE objects.
    #[inline]
    #[must_use]
    pub fn data_stream(&self) -> Option<&[u8]> {
        self.data_stream.as_deref()
    }

    /// Get the floating-shape anchors of the Main Document.
    ///
    /// Each entry maps the character position of a 0x0008 floating-shape
    /// anchor character to its positioning attributes ([MS-DOC] Spa): the
    /// shape id (which matches the `spid` of the shape's `OfficeArtFSP`), the
    /// position rectangle in twips, the position origins, and the
    /// text-wrapping style. Returns an empty slice when the document has no
    /// floating shapes in the main story.
    #[inline]
    #[must_use]
    pub fn shape_positions(&self) -> &[crate::parts::spa::ShapeAnchor] {
        &self.shape_anchors
    }

    /// Get the floating-shape anchors of the Header Document.
    ///
    /// Like [`Self::shape_positions`], but for shapes anchored in the
    /// header/footer story (positions from the `PlcfSpaHdr`). Returns an empty
    /// slice when the document has no floating shapes in the header story.
    #[inline]
    #[must_use]
    pub fn header_shape_positions(&self) -> &[crate::parts::spa::ShapeAnchor] {
        &self.header_shape_anchors
    }

    /// Map a header-story-relative character position to the header it
    /// belongs to.
    fn header_kind_at_cp(
        &self,
        story_relative_cp: u32,
    ) -> Option<crate::parts::headers::HeaderFooterType> {
        let (story_base, _) = self.fib.get_header_range()?;
        let absolute_cp = story_base.checked_add(story_relative_cp)?;
        self.headers_table.as_ref().and_then(|table| {
            table
                .stories()
                .iter()
                .find(|story| {
                    story.story_type.is_header()
                        && absolute_cp >= story.start_cp
                        && absolute_cp < story.end_cp
                })
                .map(|story| story.story_type)
        })
    }

    /// Get the header type containing a header-story character position.
    ///
    /// Floating-shape anchors in the Header Document carry CPs relative to
    /// the start of the header story (see [`Self::header_shape_positions`]);
    /// this maps such a CP to the header (odd, even, or first-page) whose
    /// story range contains it. Returns `None` when the document has no
    /// matching header story.
    #[must_use]
    pub fn header_story_kind_at_cp(
        &self,
        cp: u32,
    ) -> Option<crate::parts::headers::HeaderFooterType> {
        self.header_kind_at_cp(cp)
    }

    /// Resolve text box entries against a textbox story range.
    ///
    /// For header-story text boxes, the header kind is resolved through the
    /// box's shape: its Spa anchor CP lives in the header story (the textbox
    /// story has its own CP space), and the header owning that CP answers
    /// the kind.
    fn resolve_text_boxes(
        &self,
        entries: &[crate::parts::textbox::TextBoxEntry],
        story_range: Option<(u32, u32)>,
        in_header_story: bool,
    ) -> Vec<crate::parts::textbox::TextBox> {
        let Some((story_start, _)) = story_range else {
            return Vec::new();
        };
        entries
            .iter()
            .map(|entry| {
                let raw = self
                    .text_extractor
                    .text_at_range(story_start + entry.start_cp, story_start + entry.end_cp);
                // The range of each text box ends with a trailing CR.
                let text = raw.strip_suffix('\r').unwrap_or(raw);
                let header_kind = if in_header_story {
                    self.header_shape_anchors
                        .iter()
                        .find(|anchor| anchor.spa.shape_id == entry.shape_id)
                        .and_then(|anchor| self.header_kind_at_cp(anchor.cp))
                } else {
                    None
                };
                crate::parts::textbox::TextBox {
                    shape_id: entry.shape_id,
                    text: text.to_string(),
                    header_kind,
                }
            })
            .collect()
    }

    /// Get the text boxes of the document with their plain-text content.
    ///
    /// The text comes from the textbox story (the subdocument counted by
    /// ccpTxbx); each entry's `shape_id` matches the `spid` of the shape's
    /// `OfficeArtFSP` record in the drawing layer and the `lid` of its Spa.
    /// Paragraphs within a text box are separated by '\r'. Returns an empty
    /// vector when the document has no textbox story.
    #[must_use]
    pub fn text_boxes(&self) -> Vec<crate::parts::textbox::TextBox> {
        self.resolve_text_boxes(&self.textbox_entries, self.fib.get_textbox_range(), false)
    }

    /// Get the text boxes anchored in the header/footer story.
    ///
    /// Like [`Self::text_boxes`], but for the header textbox story (counted
    /// by ccpHdrTxbx, linked through `PlcfHdrtxbxTxt`). Each entry's
    /// `header_kind` reports the header (odd, even, or first-page) the box is
    /// anchored in. Returns an empty vector when the document has no header
    /// textbox story.
    #[must_use]
    pub fn header_text_boxes(&self) -> Vec<crate::parts::textbox::TextBox> {
        self.resolve_text_boxes(
            &self.header_textbox_entries,
            self.fib.get_header_textbox_range(),
            true,
        )
    }
}
