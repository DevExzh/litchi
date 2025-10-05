/// Document - the main API for working with Word document content.
use crate::ooxml::docx::paragraph::Paragraph;
use crate::ooxml::docx::parts::DocumentPart;
use crate::ooxml::docx::table::Table;
use crate::ooxml::error::Result;

/// A Word document.
///
/// This is the main API for reading and manipulating Word document content.
/// It provides access to paragraphs, tables, sections, styles, and other
/// document elements.
///
/// # Examples
///
/// ```rust,no_run
/// use litchi::ooxml::docx::Package;
///
/// let pkg = Package::open("document.docx")?;
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
pub struct Document<'a> {
    /// The underlying document part
    part: DocumentPart<'a>,
}

impl<'a> Document<'a> {
    /// Create a new Document from a DocumentPart.
    ///
    /// This is typically called internally by `Package::document()`.
    #[inline]
    pub(crate) fn new(part: DocumentPart<'a>) -> Self {
        Self { part }
    }

    /// Get all text content from the document.
    ///
    /// This extracts all text from all paragraphs in the document,
    /// concatenated together.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi::ooxml::docx::Package;
    ///
    /// let pkg = Package::open("document.docx")?;
    /// let doc = pkg.document()?;
    /// let text = doc.text()?;
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn text(&self) -> Result<String> {
        self.part.extract_text()
    }

    /// Get the number of paragraphs in the document.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi::ooxml::docx::Package;
    ///
    /// let pkg = Package::open("document.docx")?;
    /// let doc = pkg.document()?;
    /// let count = doc.paragraph_count()?;
    /// println!("Paragraphs: {}", count);
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn paragraph_count(&self) -> Result<usize> {
        self.part.paragraph_count()
    }

    /// Get the number of tables in the document.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi::ooxml::docx::Package;
    ///
    /// let pkg = Package::open("document.docx")?;
    /// let doc = pkg.document()?;
    /// let count = doc.table_count()?;
    /// println!("Tables: {}", count);
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn table_count(&self) -> Result<usize> {
        self.part.table_count()
    }

    /// Get access to the underlying document part.
    ///
    /// This provides lower-level access to the document XML.
    #[inline]
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
    /// use litchi::ooxml::docx::Package;
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
    /// use litchi::ooxml::docx::Package;
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
    pub fn tables(&self) -> Result<Vec<Table>> {
        // Convert SmallVec to Vec for API compatibility
        Ok(self.part.tables()?.into_iter().collect())
    }

    // TODO: Add more methods:
    // - sections() -> Iterator<Section>
    // - styles() -> Styles
    // - add_paragraph() -> Paragraph (writing support)
    // - add_table() -> Table (writing support)
    // - save() (writing support)
}

// Note: Paragraph, Run, Table, Row, Cell are now in separate modules:
// - paragraph.rs: Paragraph and Run
// - table.rs: Table, Row, Cell

/// A section in a Word document.
///
/// Represents a `<w:sectPr>` element in the document XML.
///
/// # Future API
///
/// ```rust,ignore
/// impl Section {
///     pub fn page_size(&self) -> (Length, Length);
///     pub fn margins(&self) -> Margins;
///     pub fn header(&self) -> Option<Header>;
///     pub fn footer(&self) -> Option<Footer>;
/// }
/// ```
pub struct Section {
    // TODO: Implement section structure
}

#[cfg(test)]
mod tests {
    // Tests will be added as implementation progresses
}
