/// Document - the main API for working with Word document content.
use super::package::{DocError, Result};
use super::paragraph::Paragraph;
use super::parts::fib::FileInformationBlock;
use super::parts::text::TextExtractor;
use super::table::Table;
use super::super::OleFile;
use std::fs::File;

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
    /// The main document text
    word_document: Vec<u8>,
    /// The table stream (0Table or 1Table)
    table_stream: Vec<u8>,
    /// Text extractor
    text_extractor: TextExtractor,
}

impl Document {
    /// Create a new Document from an OLE file.
    ///
    /// This is typically called internally by `Package::document()`.
    pub(crate) fn from_ole(ole: &mut OleFile<File>) -> Result<Self> {
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

        Ok(Self {
            fib,
            word_document,
            table_stream,
            text_extractor,
        })
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
        // TODO: Implement proper paragraph extraction from binary structures
        // For now, split text into lines as a simple approximation
        let text = self.text()?;
        Ok(text
            .lines()
            .map(|line| Paragraph::new(line.to_string()))
            .collect())
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

