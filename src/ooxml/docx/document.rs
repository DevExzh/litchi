use crate::ooxml::docx::parts::DocumentPart;
/// Document - the main API for working with Word document content.
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

    // TODO: Add more methods:
    // - paragraphs() -> Iterator<Paragraph>
    // - tables() -> Iterator<Table>
    // - sections() -> Iterator<Section>
    // - styles() -> Styles
    // - add_paragraph() -> Paragraph
    // - add_table() -> Table
    // - save()
}

/// A paragraph in a Word document.
///
/// Represents a `<w:p>` element in the document XML.
///
/// # Future API
///
/// ```rust,ignore
/// impl Paragraph {
///     pub fn text(&self) -> String;
///     pub fn runs(&self) -> impl Iterator<Item = Run>;
///     pub fn style(&self) -> Option<&str>;
///     pub fn add_run(&mut self, text: &str) -> Run;
/// }
/// ```
pub struct Paragraph {
    // TODO: Implement paragraph structure
}

/// A table in a Word document.
///
/// Represents a `<w:tbl>` element in the document XML.
///
/// # Future API
///
/// ```rust,ignore
/// impl Table {
///     pub fn rows(&self) -> impl Iterator<Item = Row>;
///     pub fn row_count(&self) -> usize;
///     pub fn column_count(&self) -> usize;
///     pub fn cell(&self, row: usize, col: usize) -> Option<Cell>;
/// }
/// ```
pub struct Table {
    // TODO: Implement table structure
}

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
