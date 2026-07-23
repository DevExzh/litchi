use super::super::OleFile;
use super::bookmark::Bookmark;
/// Document - the main API for working with Word document content.
use super::comment::Comment;
use super::encryption::decrypt_document_streams;
use super::footnote::Footnote;
use super::header_footer::HeaderFooter;
use super::hyperlink::Hyperlink;
use super::package::{DocError, DocOpenOptions, Result};
use super::paragraph::{Paragraph, Run};
use super::parts::associated_strings::DocumentAssociatedStrings;
use super::parts::bookmarks::BookmarksTable;
use super::parts::chp_bin_table::ChpBinTable;
use super::parts::comments::CommentsTable;
use super::parts::fib::FileInformationBlock;
use super::parts::fields::{
    AdvanceField, Field, FieldStory, FieldText, FieldsTable, GoToButtonField, IfField,
    MacroButtonField, MailMergeConditionalControlField, MailMergeCounterField, MailMergeNextField,
    MailMergeRecipientField, MergeField, PromptField, UserIdentityField,
};
use super::parts::footnotes::{EndnotesTable, FootnotesTable};
use super::parts::headers::HeadersTable;
use super::parts::hyperlinks::HyperlinksTable;
use super::parts::list_names::ListNamesTable;
use super::parts::list_templates::ListTemplateTable;
use super::parts::numbering::{ListTables, ParagraphListBinding};
use super::parts::pap_bin_table::PapBinTable;
use super::parts::paragraph_extractor::{ExtractedParagraph, ParagraphExtractor};
use super::parts::piece_table::PieceTable;
use super::parts::proofing::ProofingTables;
use super::parts::revisions::RevisionAuthorTable;
use super::parts::sections::SectionsTable;
use super::parts::styles::StyleSheet;
use super::parts::text::TextExtractor;
use super::table::Table;
#[cfg(feature = "formula")]
use crate::mtef_extractor::MtefExtractor;
use std::collections::HashMap;
use std::io::{Read, Seek};
use std::sync::Arc;

/// A Word document (.doc).
///
/// This is the main API for reading and manipulating legacy Word document content.
/// It provides access to paragraphs, tables, and other document elements.
///
/// # Examples
///
/// ```rust,no_run
/// use litchi_ole::doc::Package;
///
/// let mut pkg = Package::open("document.doc")?;
/// let doc = pkg.document()?;
///
/// // Extract all text
/// let text = doc.text()?;
/// println!("Document text: {}", text);
///
/// // Get paragraph count
/// let count = doc.paragraph_count()?;
/// println!("Number of paragraphs: {}", count);
/// # Ok::<(), Box<dyn std::error::Error>>(())
/// ```
pub struct Document {
    /// File Information Block from WordDocument stream
    fib: FileInformationBlock,
    /// The WordDocument stream - main document binary data
    /// Used during initialization for TextExtractor and ChpBinTable parsing
    #[allow(dead_code)] // False positive: used during initialization via parse_chp_bin_table
    word_document: Vec<u8>,
    /// The Data stream - contains embedded objects, pictures, etc.
    /// According to Apache POI, pictures are stored here, not in WordDocument stream.
    data_stream: Option<Vec<u8>>,
    /// Text extractor - holds the extracted document text
    text_extractor: TextExtractor,
    /// Character property bin table - parsed once and shared across all paragraph extractors
    chp_bin_table: Option<ChpBinTable>,
    /// Paragraph property bin table - parsed once and shared across all paragraph extractors
    pap_bin_table: Option<PapBinTable>,
    /// Fields table - contains field information (embedded equations, hyperlinks, etc.)
    /// Used during initialization for hyperlink extraction; exposed via `fields_table()` accessor.
    fields_table: Option<FieldsTable>,
    /// Headers and footers table
    headers_table: Option<HeadersTable>,
    /// Footnotes table
    footnotes_table: Option<FootnotesTable>,
    /// Endnotes table
    endnotes_table: Option<EndnotesTable>,
    /// Comments table
    comments_table: CommentsTable,
    /// Standard bookmark tables
    bookmarks_table: BookmarksTable,
    /// Revision-mark authors
    revision_authors: RevisionAuthorTable,
    /// Fixed associated-document strings
    associated_strings: Option<DocumentAssociatedStrings>,
    /// Names parallel to list definitions for LISTNUM fields
    list_names: Option<ListNamesTable>,
    /// List-level template codes parallel to list definitions
    list_templates: Option<ListTemplateTable>,
    /// Deferred strict spelling/grammar proofing metadata parse
    proofing_tables: Result<ProofingTables>,
    /// Section ranges, layout, and property revision marks
    sections: SectionsTable,
    /// Hyperlinks table
    hyperlinks_table: Option<HyperlinksTable>,
    /// List/numbering tables
    list_tables: Option<ListTables>,
    /// Word 97+ stylesheet, including raw style UPX property sets.
    stylesheet: Option<StyleSheet>,
    /// Extracted MTEF data from OLE streams (stream_name -> mtef_data)
    #[allow(dead_code)] // Stored for debugging and raw access
    mtef_data: std::collections::HashMap<String, Vec<u8>>,
    /// Parsed MTEF formulas rendered while their temporary parser arena is alive.
    /// Owned strings avoid a self-referential document and remain cheap to share.
    parsed_mtef: std::collections::HashMap<String, Arc<str>>,
}

#[cfg(all(test, feature = "formula"))]
mod owned_mtef_tests {
    use super::*;

    #[test]
    fn malformed_multiple_formulas_are_independently_owned_and_dropped() {
        let mut inputs = HashMap::new();
        inputs.insert("equation-a".to_string(), vec![0xAA; 7]);
        inputs.insert("equation-b".to_string(), vec![0xBB; 13]);

        let rendered = Document::parse_all_mtef_data(&inputs).expect("malformed formulas render");
        assert_eq!(rendered.len(), 2);
        assert!(rendered["equation-a"].contains("Invalid MTEF format"));
        assert!(rendered["equation-b"].contains("Invalid MTEF format"));
        assert!(!Arc::ptr_eq(
            &rendered["equation-a"],
            &rendered["equation-b"]
        ));

        let retained = Arc::clone(&rendered["equation-a"]);
        let weak = Arc::downgrade(&retained);
        drop(retained);
        drop(rendered);
        assert!(weak.upgrade().is_none());
    }
}

impl Document {
    /// Create a new Document from an OLE file.
    ///
    /// This is typically called internally by `Package::document()`.
    pub(crate) fn from_ole<R: Read + Seek>(ole: &mut OleFile<R>) -> Result<Self> {
        Self::from_ole_with_options(ole, DocOpenOptions::default())
    }

    /// Create a new Document from an OLE file with password-to-open options.
    pub(crate) fn from_ole_with_options<R: Read + Seek>(
        ole: &mut OleFile<R>,
        options: DocOpenOptions<'_>,
    ) -> Result<Self> {
        // Read the WordDocument stream (main document stream)
        let mut word_document = ole
            .open_stream(&["WordDocument"])
            .map_err(|_| DocError::StreamNotFound("WordDocument".to_string()))?;

        // Parse the File Information Block (FIB) from the start of WordDocument
        let mut fib = FileInformationBlock::parse(&word_document)?;

        // Determine which table stream to use (0Table or 1Table)
        let table_stream_name = if fib.which_table_stream() {
            "1Table"
        } else {
            "0Table"
        };

        // Read the table stream
        let mut table_stream = ole
            .open_stream(&[table_stream_name])
            .map_err(|_| DocError::StreamNotFound(table_stream_name.to_string()))?;

        // Read the Data stream (optional - contains embedded pictures and objects)
        // According to Apache POI, pictures are stored in Data stream, not WordDocument stream
        let mut data_stream = ole.open_stream(&["Data"]).ok();

        if fib.is_encrypted() {
            decrypt_document_streams(
                &fib,
                &mut word_document,
                &mut table_stream,
                data_stream.as_deref_mut(),
                options.password,
            )?;
            // FibBase is clear, but the rest of the FIB was encrypted and must be
            // reparsed before any offsets or character counts are consulted.
            fib = FileInformationBlock::parse(&word_document)?;
        }

        // Create text extractor
        let text_extractor = TextExtractor::new(&fib, &word_document, &table_stream)?;

        // Parse fields table to identify embedded equations and hyperlinks
        let fields_table = Some(FieldsTable::parse(&fib, &table_stream)?);

        // Parse headers/footers table
        let headers_table = HeadersTable::parse(&fib, &table_stream).ok();

        // Parse footnotes and endnotes tables
        let footnotes_table = FootnotesTable::parse(&fib, &table_stream).ok();
        let endnotes_table = EndnotesTable::parse(&fib, &table_stream).ok();
        let comments_table = CommentsTable::parse(&fib, &table_stream)?;
        let bookmarks_table = BookmarksTable::parse(&fib, &table_stream)?;
        let revision_authors = RevisionAuthorTable::parse(&fib, &table_stream)?;
        let associated_strings = DocumentAssociatedStrings::parse(&fib, &table_stream)?;
        let list_names = ListNamesTable::parse(&fib, &table_stream)?;
        let list_templates = ListTemplateTable::parse(&fib, &table_stream)?;
        let proofing_tables = ProofingTables::parse(&fib, &table_stream);
        let sections =
            SectionsTable::parse(&fib, &table_stream, &word_document, &revision_authors)?;

        // Parse hyperlinks from fields table
        let hyperlinks_table = fields_table.as_ref().and_then(|ft| {
            HyperlinksTable::from_fields(ft, |start, end| {
                Ok(text_extractor.text_at_range(start, end).to_string())
            })
            .ok()
        });

        // Parse list/numbering tables
        let list_tables = ListTables::parse(&fib, &table_stream).ok();

        // Word 97+ files are required to carry a stylesheet. Older Word files use
        // a different FIB and stylesheet representation that is not interpreted here.
        let mut stylesheet = (fib.version() >= 0x00C1)
            .then(|| StyleSheet::parse(&fib, &table_stream))
            .transpose()?;
        if let Some(stylesheet) = &mut stylesheet {
            stylesheet.resolve_revision_authors(&revision_authors)?;
        }

        // Extract MTEF data from OLE streams
        let mtef_data = Self::extract_mtef_data(ole)?;

        // Parse MTEF data into AST nodes
        #[cfg(feature = "formula")]
        let parsed_mtef = Self::parse_all_mtef_data(&mtef_data)?;
        #[cfg(not(feature = "formula"))]
        let parsed_mtef = Self::parse_all_mtef_data(&mtef_data)?;

        // Reconstruct both property bin tables from one shared piece-table parse.
        let piece_table = Self::parse_piece_table(&fib, &table_stream);
        let chp_bin_table = piece_table.as_ref().and_then(|piece_table| {
            Self::table_slice(&fib, &table_stream, 12)
                .and_then(|data| ChpBinTable::parse(data, &word_document, piece_table))
        });
        let pap_bin_table = if let (Some(piece_table), Some(data)) = (
            piece_table.as_ref(),
            Self::table_slice(&fib, &table_stream, 13),
        ) {
            PapBinTable::parse(
                data,
                &word_document,
                data_stream.as_deref(),
                piece_table,
                stylesheet.as_ref(),
            )?
        } else {
            None
        };

        Ok(Self {
            fib,
            word_document,
            data_stream,
            text_extractor,
            chp_bin_table,
            pap_bin_table,
            fields_table,
            headers_table,
            footnotes_table,
            endnotes_table,
            comments_table,
            bookmarks_table,
            revision_authors,
            associated_strings,
            list_names,
            list_templates,
            proofing_tables,
            sections,
            hyperlinks_table,
            list_tables,
            stylesheet,
            mtef_data,
            parsed_mtef,
        })
    }

    /// Extract MTEF data from OLE streams during document initialization
    ///
    /// This method extracts embedded equation objects from the ObjectPool directory.
    /// Each embedded equation is stored as a separate OLE object within ObjectPool.
    #[cfg(feature = "formula")]
    fn extract_mtef_data<R: Read + Seek>(ole: &mut OleFile<R>) -> Result<HashMap<String, Vec<u8>>> {
        // Extract all MTEF formulas from ObjectPool (the primary location for embedded equations)
        let mtef_data = MtefExtractor::extract_all_mtef_from_objectpool(ole)
            .map_err(|e| DocError::InvalidFormat(format!("Failed to extract MTEF data: {}", e)))?;

        // Also try direct stream names for compatibility with older formats
        let mut all_mtef = mtef_data;
        let direct_stream_names = ["Equation Native", "MSWordEquation", "Equation.3"];

        for stream_name in &direct_stream_names {
            if let Ok(Some(data)) = MtefExtractor::extract_mtef_from_stream(ole, &[stream_name]) {
                all_mtef.insert(stream_name.to_string(), data);
            }
        }

        Ok(all_mtef)
    }

    /// Extract MTEF data fallback (when formula feature is disabled)
    #[cfg(not(feature = "formula"))]
    fn extract_mtef_data<R: Read + Seek>(
        _ole: &mut OleFile<R>,
    ) -> Result<HashMap<String, Vec<u8>>> {
        Ok(HashMap::new())
    }

    fn table_slice<'a>(
        fib: &FileInformationBlock,
        table_stream: &'a [u8],
        pointer_index: usize,
    ) -> Option<&'a [u8]> {
        let (offset, length) = fib.get_table_pointer(pointer_index)?;
        let start = usize::try_from(offset).ok()?;
        let length = usize::try_from(length).ok()?;
        if length == 0 {
            return None;
        }
        table_stream.get(start..start.checked_add(length)?)
    }

    /// Parse the CLX once for both property bin tables.
    fn parse_piece_table(fib: &FileInformationBlock, table_stream: &[u8]) -> Option<PieceTable> {
        Self::table_slice(fib, table_stream, 33).and_then(PieceTable::parse)
    }

    /// Parse each MTEF stream in a scoped arena and retain an owned rendering.
    ///
    /// `MathNode` is intentionally arena-borrowing. Converting before the local
    /// arena is dropped keeps `Document` an ordinary owning type with no dependent
    /// fields, leaked allocations, or extended lifetimes.
    #[cfg(feature = "formula")]
    fn parse_all_mtef_data(
        mtef_data: &HashMap<String, Vec<u8>>,
    ) -> Result<HashMap<String, Arc<str>>> {
        let mut parsed_mtef = HashMap::new();

        for (stream_name, data) in mtef_data {
            let formula = litchi_formula::Formula::new();
            let mut parser = litchi_formula::MtefParser::new(formula.arena(), data);

            if parser.is_valid() {
                match parser.parse() {
                    Ok(nodes) if !nodes.is_empty() => {
                        let mut converter = litchi_formula::LatexConverter::new();
                        let rendered = converter.convert_nodes(&nodes).map_err(|error| {
                            DocError::InvalidFormat(format!(
                                "Failed to render MTEF formula {stream_name}: {error}"
                            ))
                        })?;
                        parsed_mtef.insert(stream_name.clone(), Arc::<str>::from(rendered));
                    },
                    Ok(_) => {},
                    Err(e) => {
                        parsed_mtef.insert(
                            stream_name.clone(),
                            Arc::<str>::from(format!("[Formula parsing error: {e}]")),
                        );
                    },
                }
            } else {
                parsed_mtef.insert(
                    stream_name.clone(),
                    Arc::<str>::from(format!("[Invalid MTEF format ({} bytes)]", data.len())),
                );
            }
        }

        Ok(parsed_mtef)
    }

    /// Parse all extracted MTEF data fallback (when formula feature is disabled)
    #[cfg(not(feature = "formula"))]
    fn parse_all_mtef_data(
        _mtef_data: &HashMap<String, Vec<u8>>,
    ) -> Result<HashMap<String, Arc<str>>> {
        Ok(HashMap::new())
    }

    /// Check if text indicates a potential MTEF formula
    fn is_potential_mtef_formula(text: &str) -> bool {
        let text = text.trim();

        // Common indicators of MathType equations in text
        text.contains("MathType")
            || text.contains("MTExtra")
            || text.contains("\\")
            || text.contains("{")
            || text.contains("}")
            || (text.len() > 10 && (text.contains("^") || text.contains("_")))
    }

    /// Parse MTEF data for a given text pattern
    #[cfg(feature = "formula")]
    fn parse_mtef_for_text(&self, _text: &str) -> Option<Arc<str>> {
        // For now, try to find any parsed MTEF data
        // In a more sophisticated implementation, we'd match specific text patterns
        // to specific MTEF streams

        for parsed_ast in self.parsed_mtef.values() {
            if !parsed_ast.is_empty() {
                return Some(Arc::clone(parsed_ast));
            }
        }

        None
    }

    /// Parse MTEF data for a given text pattern (fallback when formula feature is disabled)
    #[cfg(not(feature = "formula"))]
    fn parse_mtef_for_text(&self, _text: &str) -> Option<Arc<str>> {
        None
    }

    /// Get all text content from the document.
    ///
    /// This extracts all text from the document, concatenated together.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_ole::doc::Package;
    ///
    /// let mut pkg = Package::open("document.doc")?;
    /// let doc = pkg.document()?;
    /// let text = doc.text()?;
    /// println!("{}", text);
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn text(&self) -> Result<String> {
        self.text_extractor.extract_all_text()
    }

    /// Get the number of paragraphs in the document.
    ///
    /// This method counts the same logical paragraphs returned by [`Self::paragraphs`].
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_ole::doc::Package;
    ///
    /// let mut pkg = Package::open("document.doc")?;
    /// let doc = pkg.document()?;
    /// let count = doc.paragraph_count()?;
    /// println!("Paragraphs: {}", count);
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn paragraph_count(&self) -> Result<usize> {
        Ok(self.paragraphs()?.len())
    }

    /// Get the number of tables in the document.
    ///
    /// Counts top-level tables (table_level == 1) by scanning paragraph properties
    /// for table markers. Based on Apache POI's table detection algorithm.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_ole::doc::Package;
    ///
    /// let mut pkg = Package::open("document.doc")?;
    /// let doc = pkg.document()?;
    /// let count = doc.table_count()?;
    /// println!("Tables: {}", count);
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn table_count(&self) -> Result<usize> {
        // Count tables by iterating through paragraphs and tracking table boundaries
        // A new table starts when we encounter a paragraph with in_table=true and
        // table_level=1 after a paragraph that was not in a table or had a different level
        let paragraphs = self.paragraphs()?;
        let mut table_count = 0;
        let mut in_table_level_1 = false;

        for para in paragraphs {
            let props = para.properties();

            // Check if this paragraph is in a top-level table (level 1)
            if props.in_table && props.table_nesting_level == 1 {
                // If we weren't previously in a level-1 table, this is a new table
                if !in_table_level_1 {
                    table_count += 1;
                    in_table_level_1 = true;
                }
            } else {
                // We've exited the table
                in_table_level_1 = false;
            }
        }

        Ok(table_count)
    }

    /// Get access to the File Information Block.
    ///
    /// This provides lower-level access to document properties and structure.
    #[inline]
    pub fn fib(&self) -> &FileInformationBlock {
        &self.fib
    }

    /// Get the parsed Word 97+ stylesheet.
    ///
    /// Null fixed-index slots are retained, and each non-empty style exposes its
    /// exact UPX property payloads for subsequent inheritance and formatting.
    pub fn stylesheet(&self) -> Option<&StyleSheet> {
        self.stylesheet.as_ref()
    }

    /// Get the document's fixed associated-string metadata table.
    ///
    /// Template and mail-merge paths are inert strings and are never opened.
    pub fn associated_strings(&self) -> Option<&DocumentAssociatedStrings> {
        self.associated_strings.as_ref()
    }

    /// Get the ordered `LISTNUM` list-name metadata table.
    pub fn list_names(&self) -> Option<&ListNamesTable> {
        self.list_names.as_ref()
    }

    /// Get list-level template codes parallel to the document's list definitions.
    pub fn list_templates(&self) -> Option<&ListTemplateTable> {
        self.list_templates.as_ref()
    }

    /// Strictly access spelling and grammar proofing-state ranges.
    ///
    /// Parsing is deferred so nonconforming producer caches do not prevent the document's
    /// primary text from opening. Any malformed PLCF is reported when this metadata is requested.
    pub fn proofing_tables(&self) -> Result<&ProofingTables> {
        self.proofing_tables
            .as_ref()
            .map_err(|error| DocError::Corrupted(format!("invalid proofing metadata: {error}")))
    }

    /// Get access to the fields table (if parsed).
    ///
    /// Contains information about all fields in the main document,
    /// including embedded objects and hyperlinks.
    #[inline]
    pub fn fields_table(&self) -> Option<&FieldsTable> {
        self.fields_table.as_ref()
    }

    /// Get stored instruction and cached-result text for every field story.
    ///
    /// The returned text follows the field-range rules in MS-DOC section
    /// 2.8.25. It is read from the document's existing text only: fields are
    /// never evaluated or refreshed, DDE conversations are never started,
    /// external paths are never opened, OLE objects are never activated, and
    /// macro instructions are never resolved, loaded, or executed.
    pub fn fields(&self) -> Result<Vec<FieldText>> {
        let Some(fields) = &self.fields_table else {
            return Ok(Vec::new());
        };

        fields.field_texts(|story, start, end| self.field_story_text(story, start, end))
    }

    /// Get typed, inert `MACROBUTTON` fields in story and source order.
    ///
    /// Returned values expose only stored macro or command names, button text,
    /// cached results, and field state. This method never resolves, loads,
    /// invokes, or otherwise executes a macro or command.
    pub fn macro_button_fields(&self) -> Result<Vec<MacroButtonField>> {
        let fields = self.fields()?;
        Ok(fields.iter().filter_map(FieldText::macro_button).collect())
    }

    /// Get the number of typed, inert `MACROBUTTON` fields.
    pub fn macro_button_field_count(&self) -> Result<usize> {
        Ok(self.macro_button_fields()?.len())
    }

    /// Get typed, inert `GOTOBUTTON` fields in story and source order.
    ///
    /// Returned values expose only stored destinations, button text, cached
    /// results, and field state. This method never resolves a destination,
    /// changes the insertion point, activates a jump, or refreshes a field.
    pub fn go_to_button_fields(&self) -> Result<Vec<GoToButtonField>> {
        let fields = self.fields()?;
        Ok(fields.iter().filter_map(FieldText::go_to_button).collect())
    }

    /// Get the number of typed, inert `GOTOBUTTON` fields.
    pub fn go_to_button_field_count(&self) -> Result<usize> {
        Ok(self.go_to_button_fields()?.len())
    }

    /// Get typed, inert `MERGEFIELD` fields in story and source order.
    ///
    /// Returned values expose only stored data-column names, switches, cached
    /// results, and field state. This method never opens a data source, resolves
    /// records, performs a merge, or refreshes a field result.
    pub fn merge_fields(&self) -> Result<Vec<MergeField>> {
        let fields = self.fields()?;
        Ok(fields.iter().filter_map(FieldText::merge_field).collect())
    }

    /// Get the number of typed, inert `MERGEFIELD` fields.
    pub fn merge_field_count(&self) -> Result<usize> {
        Ok(self.merge_fields()?.len())
    }

    /// Get typed, inert `MERGEREC` and `MERGESEQ` fields in story and source order.
    ///
    /// Returned values expose only stored kinds, cached results, and field
    /// state. This method never selects or counts records, opens a data source,
    /// performs a merge, or refreshes a field result.
    pub fn mail_merge_counters(&self) -> Result<Vec<MailMergeCounterField>> {
        let fields = self.fields()?;
        Ok(fields
            .iter()
            .filter_map(FieldText::mail_merge_counter)
            .collect())
    }

    /// Get the number of typed, inert mail-merge counter fields.
    pub fn mail_merge_counter_count(&self) -> Result<usize> {
        Ok(self.mail_merge_counters()?.len())
    }

    /// Get typed, inert `NEXT` mail-merge control fields in story and source order.
    ///
    /// Returned values expose only stored cached results and field state. This
    /// method never advances a record, opens a data source, performs a merge,
    /// or refreshes a field result.
    pub fn mail_merge_next_fields(&self) -> Result<Vec<MailMergeNextField>> {
        let fields = self.fields()?;
        Ok(fields
            .iter()
            .filter_map(FieldText::mail_merge_next)
            .collect())
    }

    /// Get the number of typed, inert `NEXT` mail-merge control fields.
    pub fn mail_merge_next_field_count(&self) -> Result<usize> {
        Ok(self.mail_merge_next_fields()?.len())
    }

    /// Get typed, inert `NEXTIF` and `SKIPIF` fields in story and source order.
    ///
    /// Returned values expose only stored comparison text, cached results, and
    /// field state. This method never evaluates a comparison, changes record
    /// selection, opens a data source, performs a merge, or refreshes a field
    /// result.
    pub fn mail_merge_conditional_controls(&self) -> Result<Vec<MailMergeConditionalControlField>> {
        let fields = self.fields()?;
        Ok(fields
            .iter()
            .filter_map(FieldText::mail_merge_conditional_control)
            .collect())
    }

    /// Get the number of typed, inert conditional mail-merge control fields.
    pub fn mail_merge_conditional_control_count(&self) -> Result<usize> {
        Ok(self.mail_merge_conditional_controls()?.len())
    }

    /// Get typed, inert `IF` fields in story and source order.
    ///
    /// Returned values expose only stored expression text, cached results, and
    /// field state. This method never parses or evaluates an expression,
    /// resolves field values, or refreshes a field result.
    pub fn if_fields(&self) -> Result<Vec<IfField>> {
        let fields = self.fields()?;
        Ok(fields.iter().filter_map(FieldText::if_field).collect())
    }

    /// Get the number of typed, inert `IF` fields.
    pub fn if_field_count(&self) -> Result<usize> {
        Ok(self.if_fields()?.len())
    }

    /// Get typed, inert `ASK` and `FILLIN` fields in story and source order.
    ///
    /// Returned values expose only stored prompt, bookmark, default-response,
    /// cached-result, and field-state metadata. This method never displays a
    /// prompt, captures a response, creates or updates a bookmark, performs a
    /// merge, or refreshes a field result.
    pub fn prompt_fields(&self) -> Result<Vec<PromptField>> {
        let fields = self.fields()?;
        Ok(fields.iter().filter_map(FieldText::prompt_field).collect())
    }

    /// Get the number of typed, inert `ASK` and `FILLIN` fields.
    pub fn prompt_field_count(&self) -> Result<usize> {
        Ok(self.prompt_fields()?.len())
    }

    /// Get typed, inert user-identity fields in story and source order.
    ///
    /// Returned values expose only stored kind, override, formatting, cached
    /// result, and field state. This method never reads or modifies a host
    /// user's identity, applies formatting, or refreshes a field.
    pub fn user_identity_fields(&self) -> Result<Vec<UserIdentityField>> {
        let fields = self.fields()?;
        Ok(fields
            .iter()
            .filter_map(FieldText::user_identity_field)
            .collect())
    }

    /// Get the number of typed, inert user-identity fields.
    pub fn user_identity_field_count(&self) -> Result<usize> {
        Ok(self.user_identity_fields()?.len())
    }

    /// Get typed, inert `ADVANCE` fields in story and source order.
    ///
    /// Returned values expose only stored point adjustments, cached results,
    /// and field state. This method never moves text, changes layout, reflows
    /// content, or refreshes a field.
    pub fn advance_fields(&self) -> Result<Vec<AdvanceField>> {
        let fields = self.fields()?;
        Ok(fields.iter().filter_map(FieldText::advance_field).collect())
    }

    /// Get the number of typed, inert `ADVANCE` fields.
    pub fn advance_field_count(&self) -> Result<usize> {
        Ok(self.advance_fields()?.len())
    }

    /// Get typed, inert `ADDRESSBLOCK` and `GREETINGLINE` fields in story and source
    /// order.
    ///
    /// Returned values expose stored recipient layout, locale, country, fallback,
    /// cached-result, and field-state metadata only. This method never opens a
    /// data source, selects a record, performs a merge, expands placeholders,
    /// generates text, or refreshes a field result.
    pub fn mail_merge_recipient_fields(&self) -> Result<Vec<MailMergeRecipientField>> {
        let fields = self.fields()?;
        Ok(fields
            .iter()
            .filter_map(FieldText::mail_merge_recipient_field)
            .collect())
    }

    /// Get the number of typed, inert `ADDRESSBLOCK` and `GREETINGLINE` fields.
    pub fn mail_merge_recipient_field_count(&self) -> Result<usize> {
        Ok(self.mail_merge_recipient_fields()?.len())
    }

    /// Get stored instruction and cached-result text for one parsed field.
    ///
    /// Field positions are relative to their FieldStory. This method reads only
    /// that stored text range and performs no field evaluation or external
    /// action.
    pub fn field_text(&self, field: &Field) -> Result<FieldText> {
        FieldText::from_field(field, |start, end| {
            self.field_story_text(field.story, start, end)
        })
    }

    fn field_story_text(&self, story: FieldStory, start: u32, end: u32) -> Result<String> {
        if start > end {
            return Err(DocError::Corrupted(
                "field text range has its start after its end".to_string(),
            ));
        }

        let (story_start, story_end) = self.field_story_range(story)?;
        let start = story_start
            .checked_add(start)
            .ok_or_else(|| DocError::Corrupted("field text range start overflows".to_string()))?;
        let end = story_start
            .checked_add(end)
            .ok_or_else(|| DocError::Corrupted("field text range end overflows".to_string()))?;
        if end > story_end {
            return Err(DocError::Corrupted(
                "field text range exceeds its document story".to_string(),
            ));
        }

        Ok(self.text_extractor.text_at_range(start, end).to_string())
    }

    fn field_story_range(&self, story: FieldStory) -> Result<(u32, u32)> {
        let range = match story {
            FieldStory::Main => Some(self.fib.get_main_doc_range()),
            FieldStory::Header => self.fib.get_header_range(),
            FieldStory::Footnote => self.fib.get_footnote_range(),
            FieldStory::Comment => self.fib.get_comment_range(),
            FieldStory::Endnote => self.fib.get_endnote_range(),
            FieldStory::Textbox => self.fib.get_textbox_range(),
            FieldStory::HeaderTextbox => self.fib.get_header_textbox_range(),
        };
        range.ok_or_else(|| {
            DocError::Corrupted(format!(
                "field table refers to absent {} story",
                match story {
                    FieldStory::Main => "main document",
                    FieldStory::Header => "header/footer",
                    FieldStory::Footnote => "footnote",
                    FieldStory::Comment => "comment",
                    FieldStory::Endnote => "endnote",
                    FieldStory::Textbox => "textbox",
                    FieldStory::HeaderTextbox => "header textbox",
                }
            ))
        })
    }

    // ──────────────────────────────────────────────────────────────────
    // Headers / Footers
    // ──────────────────────────────────────────────────────────────────

    /// Get all headers and footers in the document.
    ///
    /// Each section can have up to six stories: first-page header/footer,
    /// even-page header/footer, and odd-page (default) header/footer.
    /// Empty stories (where start_cp == end_cp) are omitted.
    ///
    /// # Example
    ///
    /// ```rust,ignore
    /// for hf in doc.headers_footers()? {
    ///     println!("{:?}: {}", hf.header_footer_type, hf.text());
    /// }
    /// ```
    pub fn headers_footers(&self) -> Result<Vec<HeaderFooter>> {
        let table = match &self.headers_table {
            Some(t) => t,
            None => return Ok(Vec::new()),
        };

        let mut result = Vec::new();
        for story in table.stories() {
            if story.is_empty() {
                continue;
            }

            let text = self
                .text_extractor
                .text_at_range(story.start_cp, story.end_cp)
                .to_string();

            // Extract paragraphs for this header/footer range
            let paragraphs = self.extract_paragraphs_for_range(story.start_cp, story.end_cp)?;

            let mut hf = HeaderFooter::new(story.story_type, text);
            hf.paragraphs = paragraphs;
            result.push(hf);
        }

        Ok(result)
    }

    /// Get only headers (filtering out footers).
    pub fn headers(&self) -> Result<Vec<HeaderFooter>> {
        Ok(self
            .headers_footers()?
            .into_iter()
            .filter(|hf| hf.is_header())
            .collect())
    }

    /// Get only footers (filtering out headers).
    pub fn footers(&self) -> Result<Vec<HeaderFooter>> {
        Ok(self
            .headers_footers()?
            .into_iter()
            .filter(|hf| hf.is_footer())
            .collect())
    }

    // ──────────────────────────────────────────────────────────────────
    // Footnotes / Endnotes
    // ──────────────────────────────────────────────────────────────────

    /// Get all footnotes in the document.
    ///
    /// Each footnote contains its reference position in the main document,
    /// the footnote number, and the footnote text with paragraphs.
    ///
    /// # Example
    ///
    /// ```rust,ignore
    /// for note in doc.footnotes()? {
    ///     println!("Footnote {}: {}", note.number, note.text());
    /// }
    /// ```
    pub fn footnotes(&self) -> Result<Vec<Footnote>> {
        let table = match &self.footnotes_table {
            Some(t) => t,
            None => return Ok(Vec::new()),
        };

        let mut result = Vec::with_capacity(table.count());
        for reference in table.references() {
            let text = self
                .text_extractor
                .text_at_range(reference.text_start_cp, reference.text_end_cp)
                .to_string();

            let paragraphs =
                self.extract_paragraphs_for_range(reference.text_start_cp, reference.text_end_cp)?;

            let mut note = Footnote::new(reference.ref_cp, reference.descriptor.number, text);
            note.paragraphs = paragraphs;
            result.push(note);
        }

        Ok(result)
    }

    /// Get all endnotes in the document.
    ///
    /// Endnotes share the same structure as footnotes but are placed
    /// at the end of the document or section.
    ///
    /// # Example
    ///
    /// ```rust,ignore
    /// for note in doc.endnotes()? {
    ///     println!("Endnote {}: {}", note.number, note.text());
    /// }
    /// ```
    pub fn endnotes(&self) -> Result<Vec<Footnote>> {
        let table = match &self.endnotes_table {
            Some(t) => t,
            None => return Ok(Vec::new()),
        };

        let mut result = Vec::with_capacity(table.count());
        for reference in table.references() {
            let text = self
                .text_extractor
                .text_at_range(reference.text_start_cp, reference.text_end_cp)
                .to_string();

            let paragraphs =
                self.extract_paragraphs_for_range(reference.text_start_cp, reference.text_end_cp)?;

            let mut note = Footnote::new(reference.ref_cp, reference.descriptor.number, text);
            note.paragraphs = paragraphs;
            result.push(note);
        }

        Ok(result)
    }

    // ──────────────────────────────────────────────────────────────────
    // Bookmarks
    // ──────────────────────────────────────────────────────────────────

    /// Get all standard bookmarks in start-CP order.
    pub fn bookmarks(&self) -> Result<Vec<Bookmark>> {
        Ok(self.bookmarks_table.bookmarks().to_vec())
    }

    /// Get author names used by tracked revisions and related annotations.
    pub fn revision_authors(&self) -> &[String] {
        self.revision_authors.authors()
    }

    /// Get section property revision marks in document order.
    pub fn section_revisions(&self) -> &[super::revision::SectionRevisionMark] {
        self.sections.revisions()
    }

    /// Get sections in main-document character-position order.
    pub fn sections(&self) -> &[super::section::DocSection] {
        self.sections.sections()
    }

    /// Find the section containing `cp` using half-open section ranges.
    pub fn section_at_cp(&self, cp: u32) -> Option<&super::section::DocSection> {
        self.sections.section_at_cp(cp)
    }

    // ──────────────────────────────────────────────────────────────────
    // Comments
    // ──────────────────────────────────────────────────────────────────

    /// Get all comments in main-document reference order.
    pub fn comments(&self) -> Result<Vec<Comment>> {
        let mut result = Vec::with_capacity(self.comments_table.count());
        for reference in self.comments_table.references() {
            let reference_end = reference
                .reference_cp
                .checked_add(1)
                .ok_or_else(|| DocError::Corrupted("comment reference CP overflows".to_string()))?;
            let marker_end = reference
                .marker_cp
                .checked_add(1)
                .ok_or_else(|| DocError::Corrupted("comment marker CP overflows".to_string()))?;
            if self
                .text_extractor
                .text_at_range(reference.reference_cp, reference_end)
                != "\u{5}"
                || self
                    .text_extractor
                    .text_at_range(reference.marker_cp, marker_end)
                    != "\u{5}"
            {
                return Err(DocError::Corrupted(
                    "comment reference or story does not begin with U+0005".to_string(),
                ));
            }
            if let Some(chp_table) = &self.chp_bin_table
                && (!chp_table
                    .runs_in_range(reference.reference_cp, reference_end)
                    .any(|run| run.properties.is_spec)
                    || !chp_table
                        .runs_in_range(reference.marker_cp, marker_end)
                        .any(|run| run.properties.is_spec))
            {
                return Err(DocError::Corrupted(
                    "comment reference or story marker is missing sprmCFSpec".to_string(),
                ));
            }

            let body_start = reference.marker_cp.checked_add(1).ok_or_else(|| {
                DocError::Corrupted("comment body start CP overflows".to_string())
            })?;
            let paragraph_mark_cp = reference
                .text_end_cp
                .checked_sub(1)
                .ok_or_else(|| DocError::Corrupted("comment story range is empty".to_string()))?;
            if self
                .text_extractor
                .text_at_range(paragraph_mark_cp, reference.text_end_cp)
                != "\r"
            {
                return Err(DocError::Corrupted(
                    "comment story does not end with a paragraph mark".to_string(),
                ));
            }
            let text = self
                .text_extractor
                .text_at_range(body_start, reference.text_end_cp)
                .to_string();
            let paragraphs =
                self.extract_paragraphs_for_range(body_start, reference.text_end_cp)?;
            let mut comment = Comment::new(
                reference.reference_cp,
                reference.author.clone(),
                reference.descriptor.initials.clone(),
                reference.descriptor.bookmark_tag,
                text,
            );
            comment.range_start = reference.range_start_cp;
            comment.range_end = reference.range_end_cp;
            comment.extended_metadata = reference.extended_metadata;
            comment.paragraphs = paragraphs;
            result.push(comment);
        }
        Ok(result)
    }

    // ──────────────────────────────────────────────────────────────────
    // Hyperlinks
    // ──────────────────────────────────────────────────────────────────

    /// Get all hyperlinks in the document.
    ///
    /// Hyperlinks are extracted from HYPERLINK fields in the main document.
    /// Each hyperlink includes the destination URL/path, display text, and type.
    ///
    /// # Example
    ///
    /// ```rust,ignore
    /// for link in doc.hyperlinks()? {
    ///     println!("{} -> {}", link.display_text(), link.destination());
    /// }
    /// ```
    pub fn hyperlinks(&self) -> Result<Vec<Hyperlink>> {
        let table = match &self.hyperlinks_table {
            Some(t) => t,
            None => return Ok(Vec::new()),
        };

        Ok(table
            .hyperlinks()
            .iter()
            .map(Hyperlink::from_internal)
            .collect())
    }

    /// Find hyperlinks at a specific character position in the document.
    pub fn hyperlinks_at_position(&self, cp: u32) -> Vec<Hyperlink> {
        match &self.hyperlinks_table {
            Some(t) => t
                .find_at_position(cp)
                .into_iter()
                .map(Hyperlink::from_internal)
                .collect(),
            None => Vec::new(),
        }
    }

    // ──────────────────────────────────────────────────────────────────
    // Numbering / Lists
    // ──────────────────────────────────────────────────────────────────

    /// Get the list tables (list definitions and overrides).
    ///
    /// Use this to look up list formatting for individual paragraphs
    /// via their `list_format_override` and `list_level` properties.
    ///
    /// # Example
    ///
    /// ```rust,ignore
    /// if let Some(tables) = doc.list_tables() {
    ///     for para in doc.paragraphs()? {
    ///         if let Some(info) = doc.paragraph_list_info(&para) {
    ///             println!("Level {}: {:?}", info.level, info.number_format);
    ///         }
    ///     }
    /// }
    /// ```
    pub fn list_tables(&self) -> Option<&ListTables> {
        self.list_tables.as_ref()
    }

    /// Resolve a non-empty `LISTNUM` name by zero-based `PlfLst` definition index.
    ///
    /// Entries beyond the list-definition array are ignored as required by `[MS-DOC]`.
    pub fn list_name_for_definition_index(&self, index: usize) -> Option<&str> {
        let definition_count = self.list_tables.as_ref()?.structures().len();
        if index >= definition_count {
            return None;
        }
        self.list_names.as_ref()?.name(index)
    }

    /// Get list/numbering information for a specific paragraph.
    ///
    /// Returns `Some(ListLevel)` if the paragraph is part of a list,
    /// `None` otherwise. Any `LFOLVL` start-at or formatting overrides
    /// attached to the paragraph's LFO are applied to the result.
    pub fn paragraph_list_info(
        &self,
        paragraph: &Paragraph,
    ) -> Option<super::parts::numbering::ListLevel> {
        let binding = self.paragraph_list_binding(paragraph)?;
        let mut level = binding.effective_level().clone();
        level.start_at = binding.effective_start_at();
        Some(level)
    }

    /// Resolve typed list metadata for a paragraph without cloning list data.
    ///
    /// The returned binding exposes the selected `LSTF`, `LFO`, base `LVL`,
    /// optional `LFOLVL`, effective formatting, start value, and the
    /// preserve-indents bit encoded by a negative `sprmPIlfo`.
    pub fn paragraph_list_binding(
        &self,
        paragraph: &Paragraph,
    ) -> Option<ParagraphListBinding<'_>> {
        let properties = paragraph.properties();
        let signed_lfo = properties.list_format_override?;
        let level = properties.list_level.unwrap_or(0);
        self.list_tables.as_ref()?.bind_paragraph(signed_lfo, level)
    }

    // ──────────────────────────────────────────────────────────────────
    // Internal helpers for subdocument extraction
    // ──────────────────────────────────────────────────────────────────

    /// Extract paragraphs for a specific character position range.
    ///
    /// Used internally to get paragraphs for subdocuments like
    /// headers, footers, footnotes, and endnotes.
    fn extract_paragraphs_for_range(&self, start_cp: u32, end_cp: u32) -> Result<Vec<Paragraph>> {
        if start_cp >= end_cp {
            return Ok(Vec::new());
        }

        let text = Arc::new(self.text()?);

        let para_extractor = ParagraphExtractor::new_with_range_and_stylesheet(
            Arc::clone(&text),
            self.pap_bin_table.as_ref(),
            self.chp_bin_table.as_ref(),
            (start_cp, end_cp),
            self.stylesheet.as_ref(),
        )?;

        let extracted = para_extractor.extract_paragraphs()?;
        let mut paragraphs = Vec::with_capacity(extracted.len());
        self.convert_to_paragraphs(extracted, &mut paragraphs)?;
        Ok(paragraphs)
    }

    /// Get image binary data for an embedded image.
    ///
    /// This method extracts the image data from the WordDocument stream.
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
    #[cfg(feature = "imgconv")]
    pub fn image_data(
        &self,
        image: &super::image::Image,
    ) -> std::result::Result<crate::extractor::ExtractedImage<'_>, super::image::ImageError> {
        // Use the appropriate stream based on pic_offset
        let data_stream = self.get_data_stream(image.pic_offset()).ok_or(
            super::image::ImageError::InvalidPicOffset(image.pic_offset()),
        )?;
        let word_document = self.word_document();
        image.data(data_stream, word_document)
    }

    /// Get a reference to the WordDocument stream.
    ///
    /// This is useful for low-level image operations.
    #[inline]
    pub fn word_document(&self) -> &[u8] {
        &self.word_document
    }

    /// Get the appropriate stream for picture data based on pic_offset.
    ///
    /// According to Apache POI's PicturesTable.getData():
    /// - If Data stream exists and pic_offset < data_stream.len(), use Data stream
    /// - Otherwise use WordDocument stream
    ///
    /// This is because pictures are typically stored in the Data stream,
    /// not the WordDocument stream.
    fn get_data_stream(&self, offset: u32) -> Option<&[u8]> {
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
    pub fn data_stream(&self) -> Option<&[u8]> {
        self.data_stream.as_deref()
    }

    /// Get all paragraphs in the document.
    ///
    /// Returns a vector of `Paragraph` objects representing paragraphs
    /// from all subdocuments (main, headers, footers, footnotes, etc.).
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_ole::doc::Package;
    ///
    /// let mut pkg = Package::open("document.doc")?;
    /// let doc = pkg.document()?;
    ///
    /// for para in doc.paragraphs()? {
    ///     println!("Paragraph: {}", para.text()?);
    /// }
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn paragraphs(&self) -> Result<Vec<Paragraph>> {
        let mut all_paragraphs = Vec::new();

        // Wrap text in Arc to share across all extractors without cloning (thread-safe)
        let text = Arc::new(self.text()?);

        // Get all subdocument ranges from FIB
        let subdoc_ranges = self.fib.get_all_subdoc_ranges();

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

            // Create extractor for this CP range - text is shared via Arc::clone (cheap pointer copy)
            // Pass ChpBinTable reference to avoid re-parsing
            let para_extractor = ParagraphExtractor::new_with_range_and_stylesheet(
                Arc::clone(&text),
                self.pap_bin_table.as_ref(),
                self.chp_bin_table.as_ref(),
                (start_cp, end_cp),
                self.stylesheet.as_ref(),
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
    /// This is a helper method used by paragraphs() to convert the raw extracted
    /// paragraph data into high-level Paragraph objects with formula and image support.
    fn convert_to_paragraphs(
        &self,
        extracted_paras: Vec<ExtractedParagraph>,
        output: &mut Vec<Paragraph>,
    ) -> Result<()> {
        use super::image::extract_image;

        // Pre-allocate run vectors based on estimated size
        let mut object_name_buffer = String::with_capacity(32);

        for (_para_text, para_props, runs) in extracted_paras {
            // Pre-allocate run storage
            let mut run_objects = Vec::with_capacity(runs.len());

            // Create runs for the paragraph, checking for MTEF formulas, images, and OLE2 objects
            for (text, props) in runs {
                // Primary matching: Use pic_offset to find MTEF data (most reliable)
                if let Some(pic_offset) = props.pic_offset {
                    // Skip zero offsets as they're likely invalid
                    if pic_offset > 0 {
                        // Reuse buffer to avoid repeated allocations
                        object_name_buffer.clear();
                        use std::fmt::Write;
                        let _ = write!(object_name_buffer, "_{}", pic_offset);

                        if let Some(mtef_ast) = self.parsed_mtef.get(object_name_buffer.as_str()) {
                            // Found matching formula - create run with MTEF AST (Arc::clone is cheap)
                            run_objects.push(Run::with_mtef_formula(
                                text,
                                props,
                                Arc::clone(mtef_ast),
                            ));
                            continue;
                        }
                    }
                }

                // Secondary matching: Check if this is an OLE2 object without pic_offset
                if props.is_ole2
                    && Self::is_potential_mtef_formula(&text)
                    && let Some(mtef_ast) = self.parse_mtef_for_text(&text)
                {
                    run_objects.push(Run::with_mtef_formula(text, props, mtef_ast));
                    continue;
                }

                // Check for embedded images
                // According to Apache POI, pictures are stored in Data stream if available
                if let Some(pic_offset) = props.pic_offset
                    && let Some(data_stream) = self.get_data_stream(pic_offset)
                    && let Ok(Some(image)) = extract_image(data_stream, &text, &props)
                {
                    run_objects.push(Run::with_image(text, props, image));
                    continue;
                }

                // Regular run without formula or image
                run_objects.push(Run::new(text, props));
            }

            for run in &mut run_objects {
                run.resolve_revisions(&self.revision_authors)?;
            }

            // Create paragraph with runs and properties
            // Following Apache POI's design: text is stored in runs, not duplicated in paragraph
            // Pass empty string since runs contain all the text
            let mut para = Paragraph::new(String::new());
            para.set_runs(run_objects);
            para.set_properties(para_props);
            para.resolve_revision(&self.revision_authors)?;
            output.push(para);
        }
        Ok(())
    }

    /// Get all tables in the document.
    ///
    /// Extracts tables by grouping paragraphs that have table markers.
    /// Based on Apache POI's TableIterator algorithm.
    ///
    /// Returns a vector of `Table` objects representing tables
    /// in the document.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_ole::doc::Package;
    ///
    /// let mut pkg = Package::open("document.doc")?;
    /// let doc = pkg.document()?;
    ///
    /// for table in doc.tables()? {
    ///     println!("Table with {} rows", table.row_count()?);
    /// }
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn tables(&self) -> Result<Vec<Table>> {
        self.extract_tables_from_paragraphs(&self.paragraphs()?, 1)
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
    /// use litchi_ole::doc::Package;
    /// use litchi_ole::doc::DocElement as DocumentElement;
    ///
    /// let mut pkg = Package::open("document.doc")?;
    /// let doc = pkg.document()?;
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
    pub fn elements(&self) -> Result<Vec<super::DocElement>> {
        use super::DocElement;

        // Extract all paragraphs once
        let paragraphs = self.paragraphs()?;
        let mut elements = Vec::new();
        let mut i = 0;

        while i < paragraphs.len() {
            let para = &paragraphs[i];
            let props = para.properties();

            // Check if this paragraph starts a top-level table (level 1)
            if props.in_table && props.table_nesting_level == 1 {
                // Found the start of a table - collect all paragraphs in this table
                let mut table_paras = Vec::new();

                // Collect paragraphs until we exit the table
                while i < paragraphs.len() {
                    let current_para = &paragraphs[i];
                    let current_props = current_para.properties();

                    if !current_props.in_table || current_props.table_nesting_level < 1 {
                        // Exited the table
                        break;
                    }

                    table_paras.push(current_para.clone());
                    i += 1;
                }

                // Extract rows from the collected table paragraphs
                let rows = self.extract_rows_from_table_paragraphs(&table_paras, 1)?;

                if !rows.is_empty() {
                    let properties = rows.first().and_then(|row| row.properties()).cloned();
                    let table = if let Some(properties) = properties {
                        Table::with_properties(rows, properties)
                    } else {
                        Table::new(rows)
                    };
                    elements.push(DocElement::Table(Box::new(table)));
                }
            } else if !props.in_table {
                // This is a regular paragraph (not in a table)
                elements.push(DocElement::Paragraph(Box::new(para.clone())));
                i += 1;
            } else {
                // This paragraph is in a nested table (level > 1), skip it
                // as it will be processed as part of its parent table
                i += 1;
            }
        }

        Ok(elements)
    }

    /// Extract tables from a list of paragraphs at a specific nesting level.
    ///
    /// This is based on Apache POI's table extraction algorithm that scans
    /// paragraphs for table markers and groups them into Table structures.
    ///
    /// # Arguments
    ///
    /// * `paragraphs` - List of paragraphs to scan
    /// * `level` - Table nesting level to extract (1 for top-level tables)
    ///
    /// # Returns
    ///
    /// Vector of Table objects found at the specified nesting level
    fn extract_tables_from_paragraphs(
        &self,
        paragraphs: &[Paragraph],
        level: i32,
    ) -> Result<Vec<Table>> {
        let mut tables = Vec::new();
        let mut i = 0;

        while i < paragraphs.len() {
            let para = &paragraphs[i];
            let props = para.properties();

            // Check if this paragraph starts a table at the requested level
            if props.in_table && props.table_nesting_level == level {
                // Found the start of a table - collect all paragraphs in this table
                let mut table_paras = Vec::new();

                // Collect paragraphs until we exit the table
                while i < paragraphs.len() {
                    let current_para = &paragraphs[i];
                    let current_props = current_para.properties();

                    if !current_props.in_table || current_props.table_nesting_level < level {
                        // Exited the table
                        break;
                    }

                    table_paras.push(current_para.clone());
                    i += 1;
                }

                // Now extract rows from the collected table paragraphs
                let rows = self.extract_rows_from_table_paragraphs(&table_paras, level)?;

                if !rows.is_empty() {
                    let properties = rows.first().and_then(|row| row.properties()).cloned();
                    if let Some(properties) = properties {
                        tables.push(Table::with_properties(rows, properties));
                    } else {
                        tables.push(Table::new(rows));
                    }
                }
            } else {
                i += 1;
            }
        }

        Ok(tables)
    }

    /// Extract rows from table paragraphs.
    ///
    /// Groups consecutive paragraphs into rows based on the is_table_row_end marker.
    /// Based on Apache POI's Table.initRows() logic.
    ///
    /// # Arguments
    ///
    /// * `table_paras` - Paragraphs belonging to a table
    /// * `level` - Table nesting level
    ///
    /// # Returns
    ///
    /// Vector of Row objects
    fn extract_rows_from_table_paragraphs(
        &self,
        table_paras: &[Paragraph],
        level: i32,
    ) -> Result<Vec<super::table::Row>> {
        use super::table::Row;

        let mut rows = Vec::new();
        let mut current_row_paras = Vec::new();

        for para in table_paras {
            let props = para.properties();

            // Skip paragraphs from nested tables (higher level)
            if props.table_nesting_level > level {
                continue;
            }

            // Add paragraph to current row
            current_row_paras.push(para.clone());

            // Check if this paragraph marks the end of a row
            if props.is_table_row_end && props.table_nesting_level == level {
                // End of row - create cells from the collected paragraphs
                let cells = self.extract_cells_from_row_paragraphs(
                    &current_row_paras,
                    props.table_properties.as_ref(),
                )?;

                if !cells.is_empty() {
                    rows.push(Row::with_metadata(
                        cells,
                        props.table_properties.clone(),
                        para.table_formatting_revision().cloned(),
                        props.table_properties_preserved_for_revision,
                    ));
                }

                current_row_paras.clear();
            }
        }

        // Handle any remaining paragraphs (incomplete row)
        if !current_row_paras.is_empty() {
            let last = current_row_paras
                .last()
                .expect("non-empty row paragraph collection");
            let cells = self.extract_cells_from_row_paragraphs(
                &current_row_paras,
                last.properties().table_properties.as_ref(),
            )?;
            if !cells.is_empty() {
                rows.push(Row::with_metadata(
                    cells,
                    last.properties().table_properties.clone(),
                    last.table_formatting_revision().cloned(),
                    last.properties().table_properties_preserved_for_revision,
                ));
            }
        }

        super::table::apply_table_cell_styles(&mut rows);
        Ok(rows)
    }

    /// Extract cells from row paragraphs.
    ///
    /// Each cell typically consists of one or more paragraphs.
    /// Cell marks delimit groups of one or more paragraphs, while TAP properties
    /// provide the corresponding per-cell formatting.
    ///
    /// # Arguments
    ///
    /// * `row_paras` - Paragraphs belonging to a row
    ///
    /// # Returns
    ///
    /// Vector of Cell objects
    fn extract_cells_from_row_paragraphs(
        &self,
        row_paras: &[Paragraph],
        table_properties: Option<&super::parts::tap::TableProperties>,
    ) -> Result<Vec<super::table::Cell>> {
        use super::table::Cell;

        let mut cells = Vec::new();
        let mut cell_paragraphs = Vec::new();
        for para in row_paras {
            let props = para.properties();

            // Skip the row-end marker paragraph as it doesn't contain cell content
            if props.is_table_row_end {
                continue;
            }

            cell_paragraphs.push(para.clone());
            if props.is_table_cell_end {
                let properties = table_properties
                    .and_then(|tap| tap.cell_properties.get(cells.len()))
                    .cloned();
                cells.push(Cell::with_properties(
                    std::mem::take(&mut cell_paragraphs),
                    properties,
                ));
            }
        }

        if !cell_paragraphs.is_empty() {
            let properties = table_properties
                .and_then(|tap| tap.cell_properties.get(cells.len()))
                .cloned();
            cells.push(Cell::with_properties(cell_paragraphs, properties));
        }

        if cells.is_empty() && !row_paras.is_empty() {
            cells.push(Cell::new(String::new()));
        }

        Ok(cells)
    }
}

#[cfg(test)]
mod tests {
    #[cfg(feature = "imgconv")]
    use super::super::{Image, ImageError, Package};
    #[cfg(feature = "imgconv")]
    use std::path::Path;

    #[cfg(feature = "imgconv")]
    #[test]
    fn test_extract_png_image_from_doc() {
        let base = Path::new(env!("CARGO_MANIFEST_DIR"))
            .parent()
            .unwrap()
            .parent()
            .unwrap();
        let doc_path = base
            .join("test-data")
            .join("ole")
            .join("doc")
            .join("PngPicture.doc");

        let mut pkg = Package::open(&doc_path).expect("open doc");
        let doc = pkg.document().expect("load document");
        const PNG_SIGNATURE: [u8; 8] = [0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A];

        let mut found_signature = doc
            .word_document
            .windows(PNG_SIGNATURE.len())
            .any(|window| window == PNG_SIGNATURE);

        if let Some(data_stream) = doc.data_stream.as_ref() {
            found_signature |= data_stream
                .windows(PNG_SIGNATURE.len())
                .any(|window| window == PNG_SIGNATURE);
        }

        assert!(
            found_signature,
            "expected PNG signature in document streams"
        );
    }

    #[cfg(feature = "imgconv")]
    #[test]
    fn test_image_data_with_invalid_offset() {
        let base = Path::new(env!("CARGO_MANIFEST_DIR"))
            .parent()
            .unwrap()
            .parent()
            .unwrap();
        let doc_path = base
            .join("test-data")
            .join("ole")
            .join("doc")
            .join("PngPicture.doc");

        let mut pkg = Package::open(&doc_path).expect("open doc");
        let doc = pkg.document().expect("load document");

        let img = Image::new(u32::MAX);
        let err = doc.image_data(&img).expect_err("expected invalid offset");
        assert!(matches!(err, ImageError::InvalidPicOffset(_)));
    }
}
