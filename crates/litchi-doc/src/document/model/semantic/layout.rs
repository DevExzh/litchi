use super::prelude::*;

impl Document {
    /// Get one paragraph by its zero-based position without materializing the full collection.
    ///
    /// The position spans the subdocuments in the same order as [`Self::paragraphs`]. Missing
    /// positions return `Ok(None)`.
    pub fn paragraph_at(&self, position: Position) -> Result<Option<Paragraph>> {
        let mut remaining = position.get();
        let text = self.0.text_extractor.text_shared();
        let text_ranges = self.0.text_extractor.cp_to_byte_shared();

        for (_, start_cp, end_cp) in self.0.fib.get_all_subdoc_ranges() {
            if start_cp >= end_cp {
                continue;
            }

            let para_extractor =
                ParagraphExtractor::new_with_range_and_stylesheet_and_shared_ranges(
                    Arc::clone(&text),
                    Arc::clone(&text_ranges),
                    self.0.pap_bin_table.as_ref(),
                    self.0.chp_bin_table.as_ref(),
                    (start_cp, end_cp),
                    self.0.stylesheet.as_ref(),
                )?;
            let (extracted, available) = para_extractor.extract_paragraph_at(remaining)?;
            if let Some(extracted) = extracted {
                let mut paragraphs = Vec::with_capacity(1);
                self.convert_to_paragraphs(vec![extracted], &mut paragraphs)?;
                return Ok(paragraphs.pop());
            }
            remaining = remaining.saturating_sub(available);
        }

        Ok(None)
    }

    /// Get all paragraphs in the document.
    ///
    /// Returns a vector of `Paragraph` objects representing paragraphs
    /// from all subdocuments (main, headers, footers, footnotes, etc.).
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_doc::Package;
    ///
    /// let pkg = Package::open("document.doc")?;
    /// let doc = pkg.document();
    ///
    /// for para in doc.paragraphs()? {
    ///     println!("Paragraph: {}", para.text()?);
    /// }
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn paragraphs(&self) -> Result<Vec<Paragraph>> {
        let mut all_paragraphs = Vec::new();

        // Reuse the decoded text and CP boundary map across all subdocument extractors.
        let text = self.0.text_extractor.text_shared();
        let text_ranges = self.0.text_extractor.cp_to_byte_shared();

        // Get all subdocument ranges from FIB
        let subdoc_ranges = self.0.fib.get_all_subdoc_ranges();

        // Pre-allocate if we know the approximate size
        if let Some((_, _, last_end)) = subdoc_ranges.last() {
            // Rough estimate: one paragraph per 100 characters
            let estimate = (*last_end as usize) / 100;
            all_paragraphs.reserve(estimate.max(16));
        }

        // Parse each subdocument range
        for (_subdoc_name, start_cp, end_cp) in subdoc_ranges {
            if start_cp >= end_cp {
                continue;
            }

            // Create extractor for this CP range - text and CP boundaries are shared via Arc
            // clones (cheap pointer copies). Pass bin-table references to avoid re-parsing.
            let para_extractor =
                ParagraphExtractor::new_with_range_and_stylesheet_and_shared_ranges(
                    Arc::clone(&text),
                    Arc::clone(&text_ranges),
                    self.0.pap_bin_table.as_ref(),
                    self.0.chp_bin_table.as_ref(),
                    (start_cp, end_cp),
                    self.0.stylesheet.as_ref(),
                )?;

            let extracted_paras = para_extractor.extract_paragraphs()?;

            // Convert to Paragraph objects and add to result
            self.convert_to_paragraphs(extracted_paras, &mut all_paragraphs)?;
        }

        Ok(all_paragraphs)
    }

    // fn has_picture(&self, picture_offset: u32) -> bool {}

    /// Convert extracted paragraph data to Paragraph objects.
    ///
    /// This is a helper method used by `paragraphs()` to convert the raw extracted
    /// paragraph data into high-level Paragraph objects with formula and image support.
    pub(super) fn convert_to_paragraphs(
        &self,
        extracted_paras: Vec<ExtractedParagraph>,
        output: &mut Vec<Paragraph>,
    ) -> Result<()> {
        // Pre-allocate run vectors based on estimated size
        let mut object_name_buffer = String::with_capacity(32);

        for extracted in extracted_paras {
            let ExtractedParagraph {
                properties: para_props,
                runs,
                range,
                marked,
                ..
            } = extracted;
            // Pre-allocate run storage
            let mut run_objects = Vec::with_capacity(runs.len());

            // Create runs for the paragraph, checking for MTEF formulas, images, and OLE2 objects
            for run in runs {
                let mut converted = self.convert_run(
                    run.text,
                    run.properties,
                    run.range.start(),
                    &mut object_name_buffer,
                )?;
                converted.set_cp_range(run.range);
                run_objects.push(converted);
            }

            for run in &mut run_objects {
                run.resolve_revisions(&self.0.revision_authors)?;
            }

            // Create paragraph with runs and properties
            // Following Apache POI's design: text is stored in runs, not duplicated in paragraph
            // Pass empty string since runs contain all the text
            let mut para = Paragraph::new(String::new());
            para.set_runs(run_objects);
            para.set_properties(para_props);
            para.set_cp_range(range, marked);
            para.resolve_revision(&self.0.revision_authors)?;
            output.push(para);
        }
        Ok(())
    }

    /// Build one run, checking for MTEF formulas, images, and OLE2 objects.
    fn convert_run(
        &self,
        text: String,
        props: crate::parts::chp::CharacterProperties,
        cp: u32,
        object_name_buffer: &mut String,
    ) -> Result<Run> {
        use crate::image::extract_image;

        // Primary matching: Use pic_offset to find MTEF data (most reliable)
        if let Some(pic_offset) = props.pic_offset {
            // Skip zero offsets as they're likely invalid
            if pic_offset > 0 {
                // Reuse buffer to avoid repeated allocations
                object_name_buffer.clear();
                object_name_buffer.push('_');
                object_name_buffer.push_str(&pic_offset.to_string());

                if let Some(mtef_ast) = self.0.parsed_mtef.get(object_name_buffer.as_str()) {
                    // Found matching formula - create run with MTEF AST (Arc::clone is cheap)
                    return Ok(Run::with_mtef_formula(text, props, Arc::clone(mtef_ast)));
                }
            }
        }

        // Secondary matching: Check if this is an OLE2 object without pic_offset
        if props.is_ole2
            && Self::is_potential_mtef_formula(&text)
            && let Some(mtef_ast) = self.parse_mtef_for_text(&text)
        {
            return Ok(Run::with_mtef_formula(text, props, mtef_ast));
        }

        if props.is_spec && matches!(text.as_str(), "\u{0008}" | "\u{0008}\u{0015}") {
            if !props.is_obj
                && !props.is_ole2
                && !props.is_data
                && let Some(image) = self.floating_image_at(cp, props.pic_offset.unwrap_or(0))?
            {
                return Ok(Run::with_image(text, props, image));
            }
            return Ok(Run::new(text, props));
        }

        // Check for embedded images
        // According to Apache POI, pictures are stored in Data stream if available
        if let Some(pic_offset) = props.pic_offset
            && let Some(data_stream) = self.get_data_stream(pic_offset)
            && let Ok(Some(image)) = extract_image(data_stream, &text, &props)
        {
            return Ok(Run::with_image(text, props, image));
        }

        // Regular run without formula or image
        Ok(Run::new(text, props))
    }

    /// Get all tables in the document.
    ///
    /// Extracts tables by grouping paragraphs that have table markers.
    /// Based on Apache POI's `TableIterator` algorithm.
    ///
    /// Returns a vector of `Table` objects representing tables
    /// in the document.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_doc::Package;
    ///
    /// let pkg = Package::open("document.doc")?;
    /// let doc = pkg.document();
    ///
    /// for table in doc.tables()? {
    ///     println!("Table with {} rows", table.row_count()?);
    /// }
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn tables(&self) -> Result<Vec<Table>> {
        let mut tables = Vec::new();
        self.group_elements(|element| {
            if let Grouped::Table(table) = element {
                tables.push(table);
            }
        })?;
        Ok(tables)
    }

    /// Get all document elements (paragraphs and tables) in document order.
    ///
    /// This method extracts paragraphs once and identifies which paragraphs belong to tables,
    /// returning an ordered vector of `DocumentElement` objects that preserves the document structure.
    /// This is more efficient than calling `paragraphs()` and `tables()` separately, and it
    /// maintains the correct order of elements for sequential processing (e.g., Markdown conversion).
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_doc::Package;
    /// use litchi_doc::Element as DocumentElement;
    ///
    /// let pkg = Package::open("document.doc")?;
    /// let doc = pkg.document();
    ///
    /// for element in doc.elements()? {
    ///     match element {
    ///         DocumentElement::Paragraph(para) => {
    ///             println!("Paragraph: {}", para.text()?);
    ///         }
    ///         DocumentElement::Table(table) => {
    ///             println!("Table with {} rows", table.row_count()?);
    ///         }
    ///     }
    /// }
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    ///
    /// # Performance
    ///
    /// This method is optimized to extract paragraphs only once and identify tables
    /// by scanning paragraph properties, which is significantly faster than calling
    /// `paragraphs()` and `tables()` separately.
    pub fn elements(&self) -> Result<Vec<crate::Element>> {
        let mut elements = Vec::new();
        self.group_elements(|element| {
            elements.push(match element {
                Grouped::Paragraph(paragraph) => crate::Element::Paragraph(Box::new(paragraph)),
                Grouped::Table(table) => crate::Element::Table(Box::new(table)),
            });
        })?;
        Ok(elements)
    }

    /// Group [`Self::paragraphs`] into top-level elements, in order.
    ///
    /// A table is a paragraph of table depth 1 and every following paragraph
    /// of depth 1 or more ([`crate::parts::pap::ParagraphProperties::table_depth`]);
    /// paragraphs of nested tables stay inside it, and a nested-table
    /// paragraph no top-level table encloses forms no element.
    fn group_elements(&self, mut emit: impl FnMut(Grouped)) -> Result<()> {
        use super::super::validation::{BlockGrouper, Placement};

        let mut grouper = BlockGrouper::default();
        let mut table = Vec::new();
        for paragraph in self.paragraphs()? {
            match grouper.place(paragraph.properties().table_depth()) {
                Placement::TableContinue => table.push(paragraph),
                Placement::TableStart => {
                    flush_table(&mut table, &mut emit);
                    table.push(paragraph);
                },
                Placement::Paragraph => {
                    flush_table(&mut table, &mut emit);
                    emit(Grouped::Paragraph(paragraph));
                },
                Placement::Orphan => flush_table(&mut table, &mut emit),
            }
        }
        flush_table(&mut table, &mut emit);
        Ok(())
    }
}

/// One top-level element, before the public API boxes it.
#[allow(
    clippy::large_enum_variant,
    reason = "a transient value moved straight to its consumer; boxing the paragraph here would add one allocation per paragraph that `tables` discards"
)]
enum Grouped {
    Paragraph(Paragraph),
    Table(Table),
}

/// Emit the open top-level table, if any, and leave `table` empty.
fn flush_table(table: &mut Vec<Paragraph>, emit: &mut impl FnMut(Grouped)) {
    if table.is_empty() {
        return;
    }
    if let Some(built) = super::super::validation::table_from_paragraphs(std::mem::take(table), 1) {
        emit(Grouped::Table(built));
    }
}
