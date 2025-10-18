/// Document - the main API for working with Word document content.
use super::package::{DocError, Result};
use super::paragraph::{Paragraph, Run};
use super::parts::fib::FileInformationBlock;
use super::parts::text::TextExtractor;
use super::parts::paragraph_extractor::ParagraphExtractor;
use super::parts::fields::FieldsTable;
use super::parts::pap::ParagraphProperties;
use super::parts::chp::CharacterProperties;
use super::table::Table;
use super::super::OleFile;
use crate::ole::mtef_extractor::MtefExtractor;
use std::collections::HashMap;
use std::io::{Read, Seek};

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
    word_document: Vec<u8>,
    /// The table stream (0Table or 1Table) - contains formatting and structure
    table_stream: Vec<u8>,
    /// Text extractor - holds the extracted document text
    text_extractor: TextExtractor,
    /// Fields table - contains field information (embedded equations, hyperlinks, etc.)
    #[allow(dead_code)] // Stored for future field extraction features
    fields_table: Option<FieldsTable>,
    /// Extracted MTEF data from OLE streams (stream_name -> mtef_data)
    #[allow(dead_code)] // Stored for debugging and raw access
    mtef_data: std::collections::HashMap<String, Vec<u8>>,
    /// Parsed MTEF formulas (stream_name -> parsed_ast)
    parsed_mtef: std::collections::HashMap<String, Vec<crate::formula::MathNode<'static>>>,
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
        let table_stream_name = if fib.which_table_stream() { "1Table" } else { "0Table" };

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
        let parsed_mtef = Self::parse_all_mtef_data(&mtef_data)?;

        Ok(Self {
            fib,
            word_document,
            table_stream,
            text_extractor,
            fields_table,
            mtef_data,
            parsed_mtef,
        })
    }

    /// Extract MTEF data from OLE streams during document initialization
    ///
    /// This method extracts embedded equation objects from the ObjectPool directory.
    /// Each embedded equation is stored as a separate OLE object within ObjectPool.
    fn extract_mtef_data<R: Read + Seek>(ole: &mut OleFile<R>) -> Result<HashMap<String, Vec<u8>>> {
        // Extract all MTEF formulas from ObjectPool (the primary location for embedded equations)
        let mtef_data = MtefExtractor::extract_all_mtef_from_objectpool(ole)
            .map_err(|e| DocError::InvalidFormat(format!("Failed to extract MTEF data: {}", e)))?;

        // Also try direct stream names for compatibility with older formats
        let mut all_mtef = mtef_data;
        let direct_stream_names = [
            "Equation Native",
            "MSWordEquation",
            "Equation.3",
        ];

        for stream_name in &direct_stream_names {
            if let Ok(Some(data)) = MtefExtractor::extract_mtef_from_stream(ole, &[stream_name]) {
                all_mtef.insert(stream_name.to_string(), data);
            }
        }

        Ok(all_mtef)
    }

    /// Parse all extracted MTEF data into AST nodes
    fn parse_all_mtef_data(mtef_data: &HashMap<String, Vec<u8>>) -> Result<HashMap<String, Vec<crate::formula::MathNode<'static>>>> {
        let mut parsed_mtef = HashMap::new();

        for (stream_name, data) in mtef_data {
            // Create a formula arena for parsing
            let formula = crate::formula::Formula::new();
            
            // Clone data to extend its lifetime for the parser
            // We'll need to leak the arena to make the parsed nodes 'static
            // This is necessary because we're storing them in the Document
            let arena_box = Box::new(formula);
            let arena_ptr = Box::leak(arena_box);
            
            // Create a buffer that will live as long as we need
            let data_box = data.clone().into_boxed_slice();
            let data_ptr: &'static [u8] = Box::leak(data_box);
            
            // Parse the MTEF data
            let mut parser = crate::formula::MtefParser::new(arena_ptr.arena(), data_ptr);
            
            eprintln!("DEBUG: Parsing MTEF stream '{}', {} bytes, is_valid={}", stream_name, data.len(), parser.is_valid());

            if parser.is_valid() {
                match parser.parse() {
                    Ok(nodes) if !nodes.is_empty() => {
                        // Successfully parsed - store the AST nodes
                        parsed_mtef.insert(stream_name.clone(), nodes);
                    }
                    Ok(_) => {
                        // Empty result - skip
                    }
                    Err(e) => {
                        // Parse error - store placeholder text
                        // We need to create a new arena for the placeholder
                        let placeholder_formula = crate::formula::Formula::new();
                        let placeholder_arena = Box::leak(Box::new(placeholder_formula));
                        let error_text = placeholder_arena.arena().alloc_str(&format!("[Formula parsing error: {}]", e));
                        parsed_mtef.insert(stream_name.clone(), vec![crate::formula::MathNode::Text(
                            std::borrow::Cow::Borrowed(error_text)
                        )]);
                    }
                }
            } else {
                // Invalid MTEF format - store placeholder
                let placeholder_formula = crate::formula::Formula::new();
                let placeholder_arena = Box::leak(Box::new(placeholder_formula));
                let error_text = placeholder_arena.arena().alloc_str(&format!("[Invalid MTEF format ({} bytes)]", data.len()));
                parsed_mtef.insert(stream_name.clone(), vec![crate::formula::MathNode::Text(
                    std::borrow::Cow::Borrowed(error_text)
                )]);
            }
        }

        Ok(parsed_mtef)
    }

    /// Check if text indicates a potential MTEF formula
    fn is_potential_mtef_formula(text: &str) -> bool {
        let text = text.trim();

        // Common indicators of MathType equations in text
        text.contains("MathType") ||
        text.contains("MTExtra") ||
        text.contains("\\") ||
        text.contains("{") ||
        text.contains("}") ||
        (text.len() > 10 && (text.contains("^") || text.contains("_")))
    }

    /// Parse MTEF data for a given text pattern
    fn parse_mtef_for_text(&self, _text: &str) -> Option<Vec<crate::formula::MathNode<'static>>> {
        // For now, try to find any parsed MTEF data
        // In a more sophisticated implementation, we'd match specific text patterns
        // to specific MTEF streams

        for parsed_ast in self.parsed_mtef.values() {
            if !parsed_ast.is_empty() {
                return Some(parsed_ast.clone());
            }
        }

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
        // TODO: Implement proper paragraph counting from binary structures
        // For now, approximate by counting newlines
        Ok(self.text()?.lines().count())
    }

    /// Get the number of tables in the document.
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
        // TODO: Implement proper table counting from binary structures
        Ok(0)
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
        let text = self.text()?;
        
        // Get all subdocument ranges from FIB
        let subdoc_ranges = self.fib.get_all_subdoc_ranges();
        
        eprintln!("DEBUG: Found {} subdocument ranges", subdoc_ranges.len());
        for (name, start, end) in &subdoc_ranges {
            eprintln!("DEBUG:   {}: CP range {}..{} ({} chars)", name, start, end, end - start);
        }
        
        // Parse each subdocument range
        for (subdoc_name, start_cp, end_cp) in subdoc_ranges {
            if start_cp >= end_cp {
                continue;
            }
            
            eprintln!("DEBUG: Parsing subdocument '{}' (CP {}..{})", subdoc_name, start_cp, end_cp);
            
            // Create extractor for this CP range
            let para_extractor = ParagraphExtractor::new_with_range(
                &self.fib,
                &self.table_stream,
                &self.word_document,
                text.clone(),
                (start_cp, end_cp),
            )?;
            
            let extracted_paras = para_extractor.extract_paragraphs()?;
            eprintln!("DEBUG:   Extracted {} paragraphs from '{}'", extracted_paras.len(), subdoc_name);
            
            // Convert to Paragraph objects and add to result
            self.convert_to_paragraphs(extracted_paras, &mut all_paragraphs);
        }
        
        eprintln!("DEBUG: Total paragraphs extracted: {}", all_paragraphs.len());
        Ok(all_paragraphs)
    }
    
    /// Convert extracted paragraph data to Paragraph objects.
    ///
    /// This is a helper method used by paragraphs() to convert the raw extracted
    /// paragraph data into high-level Paragraph objects with formula matching.
    fn convert_to_paragraphs(
        &self,
        extracted_paras: Vec<(String, ParagraphProperties, Vec<(String, CharacterProperties)>)>,
        output: &mut Vec<Paragraph>,
    ) {
        for (para_text, para_props, runs) in extracted_paras {
            // Create runs for the paragraph, checking for MTEF formulas and OLE2 objects
            let run_objects: Vec<Run> = runs
                .into_iter()
                .map(|(text, props)| {
                    // Primary matching: Use pic_offset to find MTEF data (most reliable)
                    if let Some(pic_offset) = props.pic_offset {
                        // Skip zero offsets as they're likely invalid
                        if pic_offset > 0 {
                            let object_name = format!("_{}", pic_offset);
                            if let Some(mtef_ast) = self.parsed_mtef.get(&object_name) {
                                // Found matching formula - create run with MTEF AST
                                return Run::with_mtef_formula(text, props, mtef_ast.clone());
                            }
                        }
                    }
                    
                    // Secondary matching: Check if this is an OLE2 object without pic_offset
                    if props.is_ole2 && Self::is_potential_mtef_formula(&text) 
                        && let Some(mtef_ast) = self.parse_mtef_for_text(&text) {
                        return Run::with_mtef_formula(text, props, mtef_ast);
                    }
                    
                    // Regular run without formula
                    Run::new(text, props)
                })
                .collect();

            // Create paragraph with runs and properties
            let mut para = Paragraph::new(para_text);
            para.set_runs(run_objects);
            para.set_properties(para_props);
            output.push(para);
        }
    }

    /// Get all tables in the document.
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
        // TODO: Implement proper table extraction from binary structures
        Ok(Vec::new())
    }
}

#[cfg(test)]
mod tests {
    // Tests will be added as implementation progresses
}

