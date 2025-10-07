/// Document - the main API for working with Word document content.
use super::package::{DocError, Result};
use super::paragraph::{Paragraph, Run};
use super::parts::fib::FileInformationBlock;
use super::parts::text::TextExtractor;
use super::parts::paragraph_extractor::ParagraphExtractor;
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
    /// The table stream (0Table or 1Table) - contains formatting and structure
    table_stream: Vec<u8>,
    /// Text extractor - holds the extracted document text
    text_extractor: TextExtractor,
    /// Extracted MTEF data from OLE streams (stream_name -> mtef_data)
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

        // Extract MTEF data from OLE streams
        let mtef_data = Self::extract_mtef_data(ole)?;

        // Parse MTEF data into AST nodes
        let parsed_mtef = Self::parse_all_mtef_data(&mtef_data)?;

        Ok(Self {
            fib,
            table_stream,
            text_extractor,
            mtef_data,
            parsed_mtef,
        })
    }

    /// Extract MTEF data from OLE streams during document initialization
    fn extract_mtef_data<R: Read + Seek>(ole: &mut OleFile<R>) -> Result<HashMap<String, Vec<u8>>> {
        let mut mtef_data = HashMap::new();

        // Common MTEF stream names in Word documents
        let mtef_stream_names = [
            "Equation Native",
            "MSWordEquation",
            "Equation.3",
        ];

        for stream_name in &mtef_stream_names {
            if let Ok(Some(data)) = MtefExtractor::extract_mtef_data_from_stream(ole, stream_name) {
                mtef_data.insert(stream_name.to_string(), data);
            }
        }

        Ok(mtef_data)
    }

    /// Parse all extracted MTEF data into AST nodes
    fn parse_all_mtef_data(mtef_data: &HashMap<String, Vec<u8>>) -> Result<HashMap<String, Vec<crate::formula::MathNode<'static>>>> {
        let mut parsed_mtef = HashMap::new();

        for (stream_name, data) in mtef_data {
            // Try to parse the MTEF data
            // let formula = crate::formula::Formula::new();
            // let mut parser = crate::formula::MtefParser::new(formula.arena(), data);

            // if parser.is_valid() && let Ok(nodes) = parser.parse() && !nodes.is_empty() {
                parsed_mtef.insert(stream_name.clone(), vec![crate::formula::MathNode::Text(
                    std::borrow::Cow::Owned(format!("MTEF Formula ({} bytes)", data.len()))
                )]);
            // }
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
    /// for para in doc.paragraphs()? {
    ///     println!("Paragraph: {}", para.text()?);
    /// }
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn paragraphs(&self) -> Result<Vec<Paragraph>> {
        // Use the proper paragraph extractor to parse from binary structures
        let text = self.text()?;
        let para_extractor = ParagraphExtractor::new(
            &self.fib,
            &self.table_stream,
            text,
        )?;

        let extracted_paras = para_extractor.extract_paragraphs()?;

        // Convert to Paragraph objects
        let mut paragraphs = Vec::with_capacity(extracted_paras.len());
        for (para_text, _para_props, runs) in extracted_paras {
            // Create runs for the paragraph, checking for MTEF formulas
            let run_objects: Vec<Run> = runs
                .into_iter()
                .map(|(text, props)| {
                    // Check if this run text indicates a potential MTEF formula
                    if Self::is_potential_mtef_formula(&text) {
                        // Try to find corresponding MTEF data and parse it
                        if let Some(mtef_ast) = self.parse_mtef_for_text(&text) {
                            Run::with_mtef_formula(text, props, mtef_ast)
                        } else {
                            Run::new(text, props)
                        }
                    } else {
                        Run::new(text, props)
                    }
                })
                .collect();

            // Create paragraph with runs
            let mut para = Paragraph::new(para_text);
            para.set_runs(run_objects);
            paragraphs.push(para);
        }

        Ok(paragraphs)
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

