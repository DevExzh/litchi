use super::super::OleFile;
/// Document - the main API for working with Word document content.
use super::package::{DocError, Result};
use super::paragraph::{Paragraph, Run};
use super::parts::chp_bin_table::ChpBinTable;
use super::parts::fib::FileInformationBlock;
use super::parts::fields::FieldsTable;
use super::parts::paragraph_extractor::{ExtractedParagraph, ParagraphExtractor};
use super::parts::text::TextExtractor;
use super::table::Table;
#[cfg(feature = "formula")]
use crate::ole::mtef_extractor::MtefExtractor;
use std::collections::HashMap;
use std::io::{Read, Seek};
use std::sync::Arc;

/// Type alias for parsed MTEF formula data with arena allocations
#[cfg(feature = "formula")]
type ParsedMtefData = (
    Vec<crate::formula::Formula<'static>>,
    Vec<Box<[u8]>>,
    HashMap<String, Arc<Vec<crate::formula::MathNode<'static>>>>,
);

/// A Word document (.doc).
///
/// This is the main API for reading and manipulating legacy Word document content.
/// It provides access to paragraphs, tables, and other document elements.
///
/// # Examples
///
/// ```rust,no_run
/// use litchi::doc::Package;
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
    /// The table stream (0Table or 1Table) - contains formatting and structure
    table_stream: Vec<u8>,
    /// Text extractor - holds the extracted document text
    text_extractor: TextExtractor,
    /// Character property bin table - parsed once and shared across all paragraph extractors
    chp_bin_table: Option<ChpBinTable>,
    /// Fields table - contains field information (embedded equations, hyperlinks, etc.)
    #[allow(dead_code)] // Stored for future field extraction features
    fields_table: Option<FieldsTable>,
    /// Extracted MTEF data from OLE streams (stream_name -> mtef_data)
    #[allow(dead_code)] // Stored for debugging and raw access
    mtef_data: std::collections::HashMap<String, Vec<u8>>,
    /// Formula arenas that own the memory for parsed formulas
    /// These must be stored to keep the arena allocations alive for the lifetime of Document
    #[cfg(feature = "formula")]
    #[allow(dead_code)] // Stored for arena lifetime management, not directly accessed
    formula_arenas: Vec<crate::formula::Formula<'static>>,
    /// Data buffers that store the MTEF binary data with 'static lifetime
    /// These must be stored to keep the buffer allocations alive for the lifetime of Document
    #[cfg(feature = "formula")]
    #[allow(dead_code)] // Stored for buffer lifetime management, not directly accessed
    data_buffers: Vec<Box<[u8]>>,
    /// Parsed MTEF formulas (stream_name -> parsed_ast)
    /// Using Arc to share AST nodes across multiple runs without cloning (thread-safe)
    #[cfg(feature = "formula")]
    parsed_mtef: std::collections::HashMap<String, Arc<Vec<crate::formula::MathNode<'static>>>>,
    /// Parsed MTEF formulas placeholder (when formula feature is disabled)
    #[cfg(not(feature = "formula"))]
    parsed_mtef: std::collections::HashMap<String, Arc<Vec<()>>>,
}

impl Document {
    /// Create a new Document from an OLE file.
    ///
    /// This is typically called internally by `Package::document()`.
    pub(crate) fn from_ole<R: Read + Seek>(ole: &mut OleFile<R>) -> Result<Self> {
        // Read the WordDocument stream (main document stream)
        let word_document = ole
            .open_stream(&["WordDocument"])
            .map_err(|_| DocError::StreamNotFound("WordDocument".to_string()))?;

        // Parse the File Information Block (FIB) from the start of WordDocument
        let fib = FileInformationBlock::parse(&word_document)?;

        // Determine which table stream to use (0Table or 1Table)
        let table_stream_name = if fib.which_table_stream() {
            "1Table"
        } else {
            "0Table"
        };

        // Read the table stream
        let table_stream = ole
            .open_stream(&[table_stream_name])
            .map_err(|_| DocError::StreamNotFound(table_stream_name.to_string()))?;

        // Create text extractor
        let text_extractor = TextExtractor::new(&fib, &word_document, &table_stream)?;

        // Parse fields table to identify embedded equations
        let fields_table = FieldsTable::parse(&fib, &table_stream).ok();

        // Extract MTEF data from OLE streams
        let mtef_data = Self::extract_mtef_data(ole)?;

        // Parse MTEF data into AST nodes
        #[cfg(feature = "formula")]
        let (formula_arenas, data_buffers, parsed_mtef) = Self::parse_all_mtef_data(&mtef_data)?;
        #[cfg(not(feature = "formula"))]
        let parsed_mtef = Self::parse_all_mtef_data(&mtef_data)?;

        // Parse ChpBinTable once here to avoid re-parsing for each subdocument
        // This is a major performance optimization since ChpBinTable::parse is expensive
        let chp_bin_table = Self::parse_chp_bin_table(&fib, &table_stream, &word_document)?;

        Ok(Self {
            fib,
            word_document,
            table_stream,
            text_extractor,
            chp_bin_table,
            fields_table,
            mtef_data,
            #[cfg(feature = "formula")]
            formula_arenas,
            #[cfg(feature = "formula")]
            data_buffers,
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

    /// Parse ChpBinTable once during document initialization.
    ///
    /// This is a performance optimization - parsing ChpBinTable is expensive,
    /// so we do it once and share the result across all paragraph extractors.
    fn parse_chp_bin_table(
        fib: &FileInformationBlock,
        table_stream: &[u8],
        word_document: &[u8],
    ) -> Result<Option<ChpBinTable>> {
        use super::parts::piece_table::PieceTable;

        // Parse piece table (required for FC-to-CP conversion in ChpBinTable)
        // According to [MS-DOC], fcClx is at FIB offset 0x01A2
        // In FibRgFcLcb97 (starting at FIB offset 154), this is index 33
        let piece_table = if let Some((offset, length)) = fib.get_table_pointer(33) {
            if length > 0 && (offset as usize) < table_stream.len() {
                let clx_data = &table_stream[offset as usize..];
                let clx_len = length.min((table_stream.len() - offset as usize) as u32) as usize;
                PieceTable::parse(&clx_data[..clx_len])
            } else {
                None
            }
        } else {
            None
        };

        // Parse ChpBinTable if we have a piece table
        // Index 12 in FibRgFcLcb97 is fcPlcfBteChpx/lcbPlcfBteChpx (PLCFBTECHPX)
        let chp_bin_table = if let (Some((offset, length)), Some(pt)) =
            (fib.get_table_pointer(12), &piece_table)
        {
            if length > 0 && (offset as usize) < table_stream.len() {
                let chp_data = &table_stream[offset as usize..];
                let chp_len = length.min((table_stream.len() - offset as usize) as u32) as usize;
                if chp_len >= 8 {
                    // Parse CHPBinTable (PlcfBteChpx with FKP pages)
                    // FKP pages are in WordDocument stream, not table stream!
                    ChpBinTable::parse(&chp_data[..chp_len], word_document, pt)
                } else {
                    None
                }
            } else {
                None
            }
        } else {
            None
        };

        Ok(chp_bin_table)
    }

    /// Parse all extracted MTEF data into AST nodes using proper arena allocation.
    ///
    /// This function creates Formula arenas for each MTEF stream and stores them
    /// to ensure the arena allocations remain valid for the lifetime of the Document.
    /// Both arenas and data buffers are kept in Vecs so they can be properly dropped
    /// when the Document is dropped, completely avoiding memory leaks.
    ///
    /// # Safety
    ///
    /// This function uses `unsafe` to extend lifetimes to 'static. This is safe because:
    /// - The formula arenas are stored in the returned Vec and owned by Document
    /// - The data buffers are stored in the returned Vec and owned by Document  
    /// - Both will live as long as the Document struct
    /// - The MathNode references remain valid because they point into these owned arenas
    #[cfg(feature = "formula")]
    fn parse_all_mtef_data(mtef_data: &HashMap<String, Vec<u8>>) -> Result<ParsedMtefData> {
        let mut formula_arenas = Vec::new();
        let mut data_buffers = Vec::new();
        let mut parsed_mtef = HashMap::new();

        for (stream_name, data) in mtef_data {
            // Create a formula arena for parsing
            let formula = crate::formula::Formula::new();

            // Clone data into a boxed slice - we'll store this to avoid leaking
            let data_box = data.clone().into_boxed_slice();

            // Get 'static references for parsing
            // Safety: We store both the arena and data buffer in the Document,
            // so they will live as long as the Document. The 'static lifetime is sound.
            let arena_ref: &'static bumpalo::Bump = unsafe {
                std::mem::transmute::<&bumpalo::Bump, &'static bumpalo::Bump>(formula.arena())
            };
            let data_ptr: &'static [u8] =
                unsafe { std::mem::transmute::<&[u8], &'static [u8]>(data_box.as_ref()) };

            // Parse the MTEF data
            let mut parser = crate::formula::MtefParser::new(arena_ref, data_ptr);

            if parser.is_valid() {
                match parser.parse() {
                    Ok(nodes) if !nodes.is_empty() => {
                        // Successfully parsed - store the AST nodes in Arc for sharing, arena, and buffer
                        parsed_mtef.insert(stream_name.clone(), Arc::new(nodes));
                        formula_arenas.push(formula);
                        data_buffers.push(data_box);
                    },
                    Ok(_) => {
                        // Empty result - skip, arena and buffer will be dropped
                    },
                    Err(e) => {
                        // Parse error - store placeholder text using the arena
                        let error_text =
                            arena_ref.alloc_str(&format!("[Formula parsing error: {}]", e));
                        parsed_mtef.insert(
                            stream_name.clone(),
                            Arc::new(vec![crate::formula::MathNode::Text(
                                std::borrow::Cow::Borrowed(error_text),
                            )]),
                        );
                        formula_arenas.push(formula);
                        data_buffers.push(data_box);
                    },
                }
            } else {
                // Invalid MTEF format - store placeholder using the arena
                let error_text =
                    arena_ref.alloc_str(&format!("[Invalid MTEF format ({} bytes)]", data.len()));
                parsed_mtef.insert(
                    stream_name.clone(),
                    Arc::new(vec![crate::formula::MathNode::Text(
                        std::borrow::Cow::Borrowed(error_text),
                    )]),
                );
                formula_arenas.push(formula);
                data_buffers.push(data_box);
            }
        }

        Ok((formula_arenas, data_buffers, parsed_mtef))
    }

    /// Parse all extracted MTEF data fallback (when formula feature is disabled)
    #[cfg(not(feature = "formula"))]
    fn parse_all_mtef_data(
        _mtef_data: &HashMap<String, Vec<u8>>,
    ) -> Result<HashMap<String, Arc<Vec<()>>>> {
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
    fn parse_mtef_for_text(
        &self,
        _text: &str,
    ) -> Option<Arc<Vec<crate::formula::MathNode<'static>>>> {
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
    fn parse_mtef_for_text(&self, _text: &str) -> Option<Arc<Vec<()>>> {
        None
    }

    /// Get all text content from the document.
    ///
    /// This extracts all text from the document, concatenated together.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi::doc::Package;
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
    /// This method counts paragraphs using the PAPBinTable (Paragraph Properties Binary Table)
    /// which provides accurate paragraph boundaries from the document's binary structures.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi::doc::Package;
    ///
    /// let mut pkg = Package::open("document.doc")?;
    /// let doc = pkg.document()?;
    /// let count = doc.paragraph_count()?;
    /// println!("Paragraphs: {}", count);
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn paragraph_count(&self) -> Result<usize> {
        // Parse PAP PLCF to get accurate paragraph count
        // Based on Apache POI's PAPBinTable approach
        use crate::ole::plcf::PlcfParser;

        // Get PAP bin table location from FIB
        // Index 13 in FibRgFcLcb97 is fcPlcfBtePapx/lcbPlcfBtePapx
        if let Some((offset, length)) = self.fib.get_table_pointer(13)
            && length > 0
            && (offset as usize) < self.table_stream.len()
        {
            let pap_data = &self.table_stream[offset as usize..];
            let pap_len = length.min((self.table_stream.len() - offset as usize) as u32) as usize;

            // Each entry in PAP PLCF represents a paragraph boundary
            if let Some(plcf) = PlcfParser::parse(&pap_data[..pap_len], 4) {
                // PLCF count represents the number of paragraph boundaries
                // The actual paragraph count is the number of intervals
                return Ok(plcf.count().saturating_sub(1).max(0));
            }
        }

        // Fallback: count from extracted paragraphs
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
    /// use litchi::doc::Package;
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

    /// Get all paragraphs in the document.
    ///
    /// Returns a vector of `Paragraph` objects representing paragraphs
    /// from all subdocuments (main, headers, footers, footnotes, etc.).
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi::doc::Package;
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
            let para_extractor = ParagraphExtractor::new_with_range(
                &self.fib,
                &self.table_stream,
                Arc::clone(&text),
                self.chp_bin_table.as_ref(),
                (start_cp, end_cp),
            )?;

            let extracted_paras = para_extractor.extract_paragraphs()?;

            // Convert to Paragraph objects and add to result
            self.convert_to_paragraphs(extracted_paras, &mut all_paragraphs);
        }

        Ok(all_paragraphs)
    }

    /// Convert extracted paragraph data to Paragraph objects.
    ///
    /// This is a helper method used by paragraphs() to convert the raw extracted
    /// paragraph data into high-level Paragraph objects with formula matching.
    fn convert_to_paragraphs(
        &self,
        extracted_paras: Vec<ExtractedParagraph>,
        output: &mut Vec<Paragraph>,
    ) {
        // Pre-allocate run vectors based on estimated size
        let mut object_name_buffer = String::with_capacity(32);

        for (_para_text, para_props, runs) in extracted_paras {
            // Pre-allocate run storage
            let mut run_objects = Vec::with_capacity(runs.len());

            // Create runs for the paragraph, checking for MTEF formulas and OLE2 objects
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

                // Regular run without formula
                run_objects.push(Run::new(text, props));
            }

            // Create paragraph with runs and properties
            // Following Apache POI's design: text is stored in runs, not duplicated in paragraph
            // Pass empty string since runs contain all the text
            let mut para = Paragraph::new(String::new());
            para.set_runs(run_objects);
            para.set_properties(para_props);
            output.push(para);
        }
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
    /// use litchi::doc::Package;
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
    /// use litchi::doc::Package;
    /// use litchi::DocumentElement;
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
    pub fn elements(&self) -> Result<Vec<crate::document::DocumentElement>> {
        use crate::document::DocumentElement;

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
                    elements.push(DocumentElement::Table(crate::document::Table::Doc(
                        Table::new(rows),
                    )));
                }
            } else if !props.in_table {
                // This is a regular paragraph (not in a table)
                elements.push(DocumentElement::Paragraph(crate::document::Paragraph::Doc(
                    para.clone(),
                )));
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
                    tables.push(Table::new(rows));
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
                let cells = self.extract_cells_from_row_paragraphs(&current_row_paras)?;

                if !cells.is_empty() {
                    rows.push(Row::new(cells));
                }

                current_row_paras.clear();
            }
        }

        // Handle any remaining paragraphs (incomplete row)
        if !current_row_paras.is_empty() {
            let cells = self.extract_cells_from_row_paragraphs(&current_row_paras)?;
            if !cells.is_empty() {
                rows.push(Row::new(cells));
            }
        }

        Ok(rows)
    }

    /// Extract cells from row paragraphs.
    ///
    /// Each cell typically consists of one or more paragraphs.
    /// The exact cell boundaries are determined by table properties (TAP).
    /// For now, we create a simple cell structure from the paragraphs.
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
    ) -> Result<Vec<super::table::Cell>> {
        use super::table::Cell;

        // For a proper implementation, we'd need to parse TAP (Table Properties)
        // to get exact cell boundaries. For now, we create one cell per paragraph
        // which is a simplified approach but works for basic tables.

        let mut cells = Vec::new();

        // Group paragraphs into cells
        // In Word's binary format, cell boundaries are marked in table properties
        // For now, we use a simple heuristic: each paragraph is a cell
        // unless it's the row-end marker
        for para in row_paras {
            let props = para.properties();

            // Skip the row-end marker paragraph as it doesn't contain cell content
            if props.is_table_row_end {
                continue;
            }

            // Create a cell with this paragraph
            let cell = Cell::with_properties(vec![para.clone()], None);
            cells.push(cell);
        }

        // If we have no cells but have a row-end marker, create at least one empty cell
        if cells.is_empty() && !row_paras.is_empty() {
            let text = row_paras
                .iter()
                .filter_map(|p| p.text().ok())
                .collect::<Vec<_>>()
                .join(" ");
            cells.push(Cell::new(text));
        }

        Ok(cells)
    }
}

#[cfg(test)]
mod tests {
    // Tests will be added as implementation progresses
}
