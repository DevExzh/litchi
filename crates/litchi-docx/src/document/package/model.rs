#![expect(
    clippy::match_same_arms,
    reason = "separate arms document distinct OOXML grammar cases"
)]
#![expect(
    clippy::shadow_reuse,
    reason = "parser bindings are intentionally refined after validation"
)]
//! Semantic document queries for the package-bound `WordprocessingML` facade.

use crate::alt::Chunk;
use crate::error::{Error, Result};
use crate::numbering::Suffix;
use crate::paragraph::Paragraph;
use crate::parts::DocumentPart;
use crate::smart_tag::SmartTag;
use crate::statistics::{
    Statistics, count_characters, count_characters_no_spaces, count_words, estimate_line_count,
    estimate_page_count,
};
use crate::table::Table;
use litchi_opc::OpcPackage;

use super::super::model::{Block, Document, Element};

impl<'a> Document<'a> {
    /// Get all text content from the document.
    ///
    /// This extracts all text from all paragraphs in the document,
    /// concatenated together.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_docx::Package;
    ///
    /// let pkg = Package::open("document.docx")?;
    /// let doc = pkg.document()?;
    /// let text = doc.text()?;
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn text(&self) -> Result<String> {
        self.part.extract_text()
    }

    /// Return numbered paragraphs with resolved, typed list markers.
    ///
    /// This is separate from [`Self::text`], whose behavior remains unchanged.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn list_items(&self) -> Result<Vec<crate::list::ListItem>> {
        Ok(self
            .resolved_list_items()?
            .into_iter()
            .map(crate::list::ListItem::from)
            .collect())
    }

    /// Return numbered paragraphs with resolved format and counter semantics.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn resolved_list_items(&self) -> Result<Vec<crate::list::ResolvedListItem>> {
        let numbering = self.numbering()?.unwrap_or_default();
        let paragraphs = self.paragraphs()?;
        let mut styles = self.styles()?;
        let mut counters = crate::list::ListCounterState::new();
        let mut items = Vec::new();

        for (paragraph_index, paragraph) in paragraphs.iter().enumerate() {
            if let Some(item) = resolve_list_item(
                paragraph_index,
                paragraph,
                &numbering,
                &mut styles,
                &mut counters,
            )? {
                items.push(item);
            }
        }
        Ok(items)
    }

    /// Return top-level document elements paired with resolved list metadata.
    ///
    /// The returned vector is aligned one-to-one with [`Self::elements`], so a
    /// table never shifts the metadata associated with following paragraphs.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn elements_with_resolved_list_items(
        &self,
    ) -> Result<Vec<(Element, Option<crate::list::ResolvedListItem>)>> {
        let numbering = self.numbering()?.unwrap_or_default();
        let elements = self.elements()?;
        let mut styles = self.styles()?;
        let mut counters = crate::list::ListCounterState::new();
        let mut paragraph_index = 0usize;
        let mut resolved = Vec::with_capacity(elements.len());

        for element in elements {
            let item = match &element {
                Element::Paragraph(paragraph) => {
                    let item = resolve_list_item(
                        paragraph_index,
                        paragraph,
                        &numbering,
                        &mut styles,
                        &mut counters,
                    )?;
                    paragraph_index = paragraph_index.checked_add(1).ok_or_else(|| {
                        Error::InvalidFormat("document paragraph index overflow".to_owned())
                    })?;
                    item
                },
                Element::Table(_) => None,
                Element::Unknown(_) => None,
            };
            resolved.push((element, item));
        }
        Ok(resolved)
    }

    /// Extract paragraph text with resolvable list labels prepended.
    ///
    /// Picture bullets and unsupported formats remain available as typed markers
    /// through [`Self::list_items`] and are not replaced with fabricated text here.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn text_with_list_labels(&self) -> Result<String> {
        let paragraphs = self.paragraphs()?;
        let items = self.list_items()?;
        let mut by_paragraph = items.into_iter().peekable();
        let mut output = String::new();
        for (index, paragraph) in paragraphs.iter().enumerate() {
            if index != 0 {
                output.push('\n');
            }
            if by_paragraph
                .peek()
                .is_some_and(|item| item.paragraph_index == index)
                && let Some(item) = by_paragraph.next()
                && let crate::list::ListMarker::Text(label) = item.marker
            {
                output.push_str(&label);
                match item.suffix {
                    Suffix::Tab => output.push('\t'),
                    Suffix::Space => output.push(' '),
                    Suffix::Nothing => {},
                }
            }
            output.push_str(&paragraph.text()?);
        }
        Ok(output)
    }

    /// Get the number of paragraphs in the document.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_docx::Package;
    ///
    /// let pkg = Package::open("document.docx")?;
    /// let doc = pkg.document()?;
    /// let count = doc.paragraph_count()?;
    /// println!("Paragraphs: {}", count);
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn paragraph_count(&self) -> Result<usize> {
        self.part.paragraph_count()
    }

    /// Get the number of tables in the document.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_docx::Package;
    ///
    /// let pkg = Package::open("document.docx")?;
    /// let doc = pkg.document()?;
    /// let count = doc.table_count()?;
    /// println!("Tables: {}", count);
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn table_count(&self) -> Result<usize> {
        self.part.table_count()
    }

    /// Get access to the underlying document part.
    ///
    /// This provides lower-level access to the document XML.
    #[inline]
    #[must_use]
    pub fn part(&self) -> &DocumentPart<'a> {
        &self.part
    }

    /// Get all paragraphs in the document.
    ///
    /// Returns a vector of `Paragraph` objects representing all `<w:p>`
    /// elements in the document body.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_docx::Package;
    ///
    /// let pkg = Package::open("document.docx")?;
    /// let doc = pkg.document()?;
    ///
    /// for para in doc.paragraphs()? {
    ///     println!("Paragraph: {}", para.text()?);
    ///
    ///     // Access runs within the paragraph
    ///     for run in para.runs()? {
    ///         println!("  Run: {} (bold: {:?})", run.text()?, run.bold()?);
    ///     }
    /// }
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn paragraphs(&self) -> Result<Vec<Paragraph>> {
        // Convert SmallVec to Vec for API compatibility
        Ok(self.part.paragraphs()?.into_iter().collect())
    }

    /// Get all tables in the document.
    ///
    /// Returns a vector of `Table` objects representing all `<w:tbl>`
    /// elements in the document body.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_docx::Package;
    ///
    /// let pkg = Package::open("document.docx")?;
    /// let doc = pkg.document()?;
    ///
    /// for table in doc.tables()? {
    ///     println!("Table with {} rows", table.row_count()?);
    ///
    ///     for (row_idx, row) in table.rows()?.iter().enumerate() {
    ///         for (col_idx, cell) in row.cells()?.iter().enumerate() {
    ///             println!("Cell [{},{}]: {}", row_idx, col_idx, cell.text()?);
    ///         }
    ///     }
    /// }
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn tables(&self) -> Result<Vec<Table>> {
        // Convert SmallVec to Vec for API compatibility
        Ok(self.part.tables()?.into_iter().collect())
    }

    /// Return all run-level smart tags in document order.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn smart_tags(&self) -> Result<Vec<SmartTag>> {
        let mut tags = Vec::new();
        for paragraph in self.part.paragraphs()? {
            tags.extend(paragraph.smart_tags()?);
        }
        Ok(tags)
    }

    /// Return the number of run-level smart tags in the document.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn smart_tag_count(&self) -> Result<usize> {
        Ok(self.smart_tags()?.len())
    }

    /// Get all document elements (paragraphs and tables) in document order.
    ///
    /// This method extracts both paragraphs and tables in a single pass,
    /// returning an ordered vector that preserves the document structure.
    /// This is more efficient than calling `paragraphs()` and `tables()` separately,
    /// and it maintains the correct order of elements for sequential processing.
    ///
    /// # Examples
    ///
    /// ```ignore
    /// use litchi_docx::Package;
    /// use litchi_docx::Element;
    ///
    /// let pkg = Package::open("document.docx")?;
    /// let doc = pkg.document()?;
    ///
    /// for element in doc.elements()? {
    ///     match element {
    ///         Element::Paragraph(para) => {
    ///             println!("Paragraph: {}", para.text()?);
    ///         }
    ///         Element::Table(table) => {
    ///             println!("Table with {} rows", table.row_count()?);
    ///         }
    ///     }
    /// }
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    ///
    /// # Performance
    ///
    /// Uses a single-pass XML parser that is significantly faster than
    /// calling `paragraphs()` and `tables()` separately.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn elements(&self) -> Result<Vec<Element>> {
        self.part.elements()
    }

    /// Return paragraphs, tables, and alternative-format anchors in document order.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn blocks(&self) -> Result<Vec<Block>> {
        self.part.blocks()
    }

    /// Return all alternative-format import anchors in XML order.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn alts(&self) -> Result<Vec<Chunk>> {
        self.part.alts()
    }

    /// Get a specific paragraph by index.
    ///
    /// # Arguments
    /// * `index` - Zero-based index of the paragraph
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_docx::Package;
    ///
    /// let pkg = Package::open("document.docx")?;
    /// let doc = pkg.document()?;
    ///
    /// // Get first paragraph
    /// if let Some(para) = doc.paragraph(0)? {
    ///     println!("First paragraph: {}", para.text()?);
    /// }
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn paragraph(&self, index: usize) -> Result<Option<Paragraph>> {
        self.part.paragraph(index)
    }

    /// Get a specific table by index.
    ///
    /// # Arguments
    /// * `index` - Zero-based index of the table
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_docx::Package;
    ///
    /// let pkg = Package::open("document.docx")?;
    /// let doc = pkg.document()?;
    ///
    /// // Get first table
    /// if let Some(table) = doc.table(0)? {
    ///     println!("Table has {} rows", table.row_count()?);
    /// }
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn table(&self, index: usize) -> Result<Option<Table>> {
        let tables = self.tables()?;
        Ok(tables.into_iter().nth(index))
    }

    /// Extract all text from a specific range of paragraphs.
    ///
    /// # Arguments
    /// * `start` - Starting paragraph index (inclusive)
    /// * `end` - Ending paragraph index (exclusive)
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_docx::Package;
    ///
    /// let pkg = Package::open("document.docx")?;
    /// let doc = pkg.document()?;
    ///
    /// // Get text from paragraphs 5-10
    /// let text = doc.text_range(5, 10)?;
    /// println!("{}", text);
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn text_range(&self, start: usize, end: usize) -> Result<String> {
        let paragraphs = self.paragraphs()?;
        let mut result = String::new();

        for (idx, para) in paragraphs.into_iter().enumerate() {
            if idx >= end {
                break;
            }
            if idx >= start {
                if !result.is_empty() {
                    result.push('\n');
                }
                result.push_str(&para.text()?);
            }
        }

        Ok(result)
    }

    /// Check if the document contains any tables.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_docx::Package;
    ///
    /// let pkg = Package::open("document.docx")?;
    /// let doc = pkg.document()?;
    ///
    /// if doc.has_tables()? {
    ///     println!("Document contains tables");
    /// }
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn has_tables(&self) -> Result<bool> {
        Ok(self.table_count()? > 0)
    }

    /// Get the underlying OPC package reference.
    ///
    /// This provides access to low-level package operations.
    #[inline]
    #[must_use]
    pub fn opc_package(&self) -> &OpcPackage {
        self.opc
    }

    /// Search for text in the document.
    ///
    /// Returns the indices of paragraphs that contain the search text.
    ///
    /// # Arguments
    /// * `query` - Text to search for (case-sensitive)
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_docx::Package;
    ///
    /// let pkg = Package::open("document.docx")?;
    /// let doc = pkg.document()?;
    ///
    /// // Find paragraphs containing "important"
    /// let matches = doc.search("important")?;
    /// println!("Found {} matching paragraphs", matches.len());
    ///
    /// for idx in matches {
    ///     if let Some(para) = doc.paragraph(idx)? {
    ///         println!("Match in paragraph {}: {}", idx, para.text()?);
    ///     }
    /// }
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn search(&self, query: &str) -> Result<Vec<usize>> {
        let paragraphs = self.paragraphs()?;
        let mut matches = Vec::new();

        for (idx, para) in paragraphs.iter().enumerate() {
            if para.text()?.contains(query) {
                matches.push(idx);
            }
        }

        Ok(matches)
    }

    /// Search for text in the document (case-insensitive).
    ///
    /// Returns the indices of paragraphs that contain the search text.
    ///
    /// # Arguments
    /// * `query` - Text to search for (case-insensitive)
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_docx::Package;
    ///
    /// let pkg = Package::open("document.docx")?;
    /// let doc = pkg.document()?;
    ///
    /// // Find paragraphs containing "important" (case-insensitive)
    /// let matches = doc.search_ignore_case("IMPORTANT")?;
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn search_ignore_case(&self, query: &str) -> Result<Vec<usize>> {
        let paragraphs = self.paragraphs()?;
        let query_lower = query.to_lowercase();
        let mut matches = Vec::new();

        for (idx, para) in paragraphs.iter().enumerate() {
            if para.text()?.to_lowercase().contains(&query_lower) {
                matches.push(idx);
            }
        }

        Ok(matches)
    }
    /// Get document statistics.
    ///
    /// Calculates comprehensive statistics about the document including
    /// word count, character count, paragraph count, and other metrics.
    ///
    /// # Performance
    ///
    /// Statistics are calculated on-demand by parsing the entire document.
    /// For large documents, consider caching the result if you need to
    /// access statistics multiple times.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_docx::Package;
    ///
    /// let pkg = Package::open("document.docx")?;
    /// let doc = pkg.document()?;
    ///
    /// let stats = doc.statistics()?;
    /// println!("Words: {}", stats.word_count());
    /// println!("Characters: {}", stats.character_count());
    /// println!("Paragraphs: {}", stats.paragraph_count());
    /// println!("Pages (estimate): {}", stats.page_count());
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn statistics(&self) -> Result<Statistics> {
        // Get all text content
        let text = self.text()?;

        // Calculate text statistics
        let word_count = count_words(&text);
        let character_count = count_characters(&text);
        let character_count_no_spaces = count_characters_no_spaces(&text);

        // Get paragraph and table counts
        let paragraph_count = self.paragraph_count()?;
        let table_count = self.table_count()?;

        // Estimate lines and pages (80 chars/line, 45 lines/page)
        let line_count = estimate_line_count(&text, 80);
        let page_count = estimate_page_count(line_count, 45);

        // Count images and drawings across all paragraphs
        let mut image_count = 0usize;
        let mut drawing_count = 0usize;

        for para in self.paragraphs()? {
            image_count = image_count
                .checked_add(para.images()?.len())
                .ok_or_else(|| Error::InvalidFormat("DOCX image count overflow".into()))?;
            drawing_count = drawing_count
                .checked_add(para.drawing_objects()?.len())
                .ok_or_else(|| Error::InvalidFormat("DOCX drawing count overflow".into()))?;
        }

        Ok(Statistics::from_counts(
            word_count,
            character_count,
            character_count_no_spaces,
            paragraph_count,
            line_count,
            page_count,
            table_count,
            image_count,
            drawing_count,
        ))
    }

    // ========================================
    // READING FEATURES - ALL IMPLEMENTED ✅
    // ========================================
    // ✅ Text extraction: text(), paragraph_count(), table_count()
    // ✅ Paragraphs: paragraphs(), paragraph(), text_range()
    // ✅ Tables: tables(), table(), has_tables()
    // ✅ Sections: sections() with page properties, margins, orientation
    // ✅ Styles: styles() with full style information
    // ✅ Headers/Footers: headers(), footers(), header(), footer()
    // ✅ Hyperlinks: hyperlinks(), hyperlink_count()
    // ✅ Footnotes/Endnotes: footnotes(), endnotes(), footnote_count(), endnote_count()
    // ✅ Comments: comments(), comment_count()
    // ✅ Bookmarks: bookmarks(), bookmark_count()
    // ✅ Fields: fields(), field_count()
    // ✅ Numbering: numbering() with abstract numbering and instances
    // ✅ Document Settings: settings(), is_protected()
    // ✅ Document Variables: document_variables()
    // ✅ Statistics: statistics() with word/character/page counts
    // ✅ Theme: theme() with color and font schemes
    // ✅ Content Controls: content_controls()
    // ✅ Custom XML: custom_xml()
    // ✅ Search: search(), search_ignore_case()
    //
    // ========================================
    // WRITE OPERATIONS
    // ========================================
    // Note: Write operations are primarily handled by the MutableDocument API
    // in the writer module. See src/ooxml/docx/writer/doc.rs for full API.
    //
    // ✅ COMPLETED: Modification operations (ECMA-376 Section 17.2.2)
    // - Add/insert/remove paragraphs: add_paragraph(), insert_paragraph(),
    //   remove_paragraph() on MutableDocument
    // - Add/insert/remove tables: add_table(), insert_table(), remove_table()
    // - Modify runs: MutableRun::bold(), italic(), font_name(), color()
    //   reached through the paragraph()/table() accessors
    // ✅ COMPLETED: Note and comment removal
    // - remove_footnote()/remove_endnote() strip typed w:footnoteReference /
    //   w:endnoteReference runs (ECMA-376 Sections 17.11.10, 17.11.2)
    // - remove_comment() drops authored w:comment entries (ECMA-376
    //   Section 17.13.4.2); note/comment IDs stay unique after removals
    //
    // ✅ COMPLETED: Track changes reading (November 2024)
    // - See revision.rs module and Paragraph::revisions() method
    // - Full support for insert, delete, move, and format revisions
    // - Includes author, date, and revision ID tracking
    //
    // ✅ Table of contents: table_of_contents(), table_of_contents_count()
    // - Typed inert field discovery and cached results; no pagination,
    //   automatic refresh, or structured-document-tag mutation
    //
    // ✅ Tables of authorities: tables_of_authorities(), table_of_authorities_entries()
    // - Typed inert TOA/TA discovery; no citation search, pagination, or refresh
    //
    // ✅ Generated indexes: indexes(), index_entries()
    // - Typed inert INDEX/XE discovery; no marker search, sorting, pagination, or refresh
    //
    // ✅ Mail merge field discovery: merge_fields(), merge_field_names(), typed_merge_fields(),
    //    mail_merge_data_fields(), mail_merge_counters()
    // ✅ COMPLETED: Mail merge settings mutation (MS-DOCX Section 17.16.5.35)
    // - Typed inert `w:mailMerge` model: Settings, DataSourceObject, FieldMap,
    //   Recipients (see crate::mail_merge)
    // - Package-level create/replace/remove: Package::set_mail_merge(),
    //   update_mail_merge(), update_mail_merge_recipients(), clear_mail_merge()
    // - Connection strings, queries, and data sources stay inert typed data;
    //   executing the merge against a data source is out of scope by design
    //
    // ✅ COMPLETED: Typed run breaks (MS-DOCX Section 17.3.3.1)
    // - Run::breaks() preserves text-wrapping, page, and column breaks plus clear behavior
    // - Rendered pagination hints remain distinguishable from authored breaks
    // ✅ COMPLETED: Section-break insertion and mutation
    // - MutableDocument::insert_section_break() authors paragraph-level w:pPr/w:sectPr;
    //   section_mut() edits the body-final section
    // - section_break(), update_section_break(), remove_section_break(),
    //   move_section_break() cover per-section page setup mutation
    //
    // ✅ Watermarks: typed VML header discovery plus mutable add/remove support
    //
    // ✅ COMPLETED: Images reading (November 2024)
    // - See image.rs module and Paragraph::images() method
    // - Full support for inline images with lazy loading
    //
    // ✅ COMPLETED: Drawing objects - shapes, text boxes (November 2024)
    // - See drawing.rs module and Paragraph::drawing_objects() method
    // - Full support for shapes, text boxes, inline/anchored positions
    // - 20+ standard shape types (rectangle, ellipse, arrows, etc.)
    //
    // ✅ Smart tags: namespace-aware reading and mutable nested-tag writing
}

fn resolve_list_item(
    paragraph_index: usize,
    paragraph: &Paragraph,
    numbering: &crate::numbering::Collection,
    styles: &mut crate::styles::Styles<'_>,
    counters: &mut crate::list::ListCounterState,
) -> Result<Option<crate::list::ResolvedListItem>> {
    let direct = paragraph.numbering()?;
    let style_id = paragraph.style_id()?;
    let inherited = if direct.is_none() {
        match style_id.as_deref() {
            Some(style_id) => styles.resolved_numbering(style_id)?,
            None => None,
        }
    } else {
        None
    };
    let associated = if direct.is_none() && inherited.is_none() {
        resolve_style_association(style_id.as_deref(), numbering)?
    } else {
        None
    };
    let Some(mut properties) = direct.or(inherited).or(associated) else {
        return Ok(None);
    };
    if properties.num_id == 0 {
        return Ok(None);
    }

    let mut linked_num_ids = std::collections::HashSet::new();
    loop {
        if !linked_num_ids.insert(properties.num_id) {
            return Err(Error::InvalidFormat(format!(
                "numbering style-link cycle at numId {}",
                properties.num_id
            )));
        }
        let num = numbering.get_num(properties.num_id).ok_or_else(|| {
            Error::InvalidFormat(format!(
                "paragraph references missing numId {}",
                properties.num_id
            ))
        })?;
        let abstract_num = numbering
            .get_abstract_num(num.abstract_num_id())
            .ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "numId {} references a missing abstract numbering definition",
                    properties.num_id
                ))
            })?;
        let Some(style_link) = abstract_num.num_style_link() else {
            break;
        };
        let linked = styles.resolved_numbering(style_link)?.ok_or_else(|| {
            Error::InvalidFormat(format!("numbering style link '{style_link}' has no numPr"))
        })?;
        if linked.num_id == 0 {
            return Ok(None);
        }
        properties.num_id = linked.num_id;
    }

    let resolved = counters.advance_resolved(numbering, properties)?;
    Ok(Some(crate::list::ResolvedListItem {
        paragraph_index,
        numbering: properties,
        format: resolved.format,
        kind: resolved.kind,
        value: resolved.value,
        marker: resolved.marker,
        suffix: resolved.suffix,
        text: paragraph.text()?,
    }))
}

fn resolve_style_association(
    style_id: Option<&str>,
    numbering: &crate::numbering::Collection,
) -> Result<Option<crate::numbering::Paragraph>> {
    let Some(style_id) = style_id else {
        return Ok(None);
    };
    let mut found = None;
    for num in numbering.nums() {
        let abstract_num = numbering
            .get_abstract_num(num.abstract_num_id())
            .ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "numId {} references a missing abstract numbering definition",
                    num.id()
                ))
            })?;
        for level in abstract_num.levels() {
            if level.paragraph_style.as_deref() != Some(style_id) {
                continue;
            }
            let candidate = crate::numbering::Paragraph {
                num_id: num.id(),
                level: level.level,
            };
            if found.is_some() {
                return Err(Error::InvalidFormat(format!(
                    "paragraph style '{style_id}' has ambiguous numbering associations"
                )));
            }
            found = Some(candidate);
        }
    }
    Ok(found)
}
