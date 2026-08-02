//! Document element types for representing ordered content.

use super::Paragraph;
#[cfg(any(feature = "doc", feature = "ooxml", feature = "rtf", feature = "odf"))]
use super::Table;

/// An element in a document's natural content order.
///
/// This enum represents the natural order of elements as they appear in a document,
/// which is essential for proper Markdown conversion and other sequential operations.
/// Table elements are available when a table-capable document feature is enabled.
///
/// # Examples
///
/// ```rust,no_run
/// use litchi::Document;
///
/// let doc = Document::open("document.docx")?;
///
/// // Process elements in document order
/// for element in doc.elements()? {
///     if let Some(para) = element.as_paragraph() {
///         println!("Paragraph: {}", para.text()?);
///     }
/// }
/// # Ok::<(), litchi::common::Error>(())
/// ```
#[derive(Debug, Clone)]
pub enum DocumentElement {
    /// A paragraph element (boxed to reduce enum size)
    Paragraph(Box<Paragraph>),
    /// A table element (boxed to reduce enum size from 12KB to ~224 bytes)
    #[cfg(any(feature = "doc", feature = "ooxml", feature = "rtf", feature = "odf"))]
    Table(Box<Table>),
}

impl DocumentElement {
    /// Check if this element is a paragraph.
    #[inline]
    pub fn is_paragraph(&self) -> bool {
        match self {
            DocumentElement::Paragraph(_) => true,
            #[cfg(any(feature = "doc", feature = "ooxml", feature = "rtf", feature = "odf"))]
            DocumentElement::Table(_) => false,
        }
    }

    /// Check if this element is a table.
    #[cfg(any(feature = "doc", feature = "ooxml", feature = "rtf", feature = "odf"))]
    #[inline]
    pub fn is_table(&self) -> bool {
        match self {
            DocumentElement::Paragraph(_) => false,
            DocumentElement::Table(_) => true,
        }
    }

    /// Get a reference to the paragraph, if this is a paragraph element.
    ///
    /// Returns `None` if this is a table element.
    #[inline]
    pub fn as_paragraph(&self) -> Option<&Paragraph> {
        match self {
            DocumentElement::Paragraph(p) => Some(p),
            #[cfg(any(feature = "doc", feature = "ooxml", feature = "rtf", feature = "odf"))]
            DocumentElement::Table(_) => None,
        }
    }

    /// Get a reference to the table, if this is a table element.
    ///
    /// Returns `None` if this is a paragraph element.
    #[cfg(any(feature = "doc", feature = "ooxml", feature = "rtf", feature = "odf"))]
    #[inline]
    pub fn as_table(&self) -> Option<&Table> {
        match self {
            DocumentElement::Table(t) => Some(t.as_ref()),
            _ => None,
        }
    }

    /// Consume this element and return the paragraph, if this is a paragraph element.
    ///
    /// Returns `None` if this is a table element.
    #[inline]
    pub fn into_paragraph(self) -> Option<Paragraph> {
        match self {
            DocumentElement::Paragraph(p) => Some(*p),
            #[cfg(any(feature = "doc", feature = "ooxml", feature = "rtf", feature = "odf"))]
            DocumentElement::Table(_) => None,
        }
    }

    /// Consume this element and return the table, if this is a table element.
    ///
    /// Returns `None` if this is a paragraph element.
    #[cfg(any(feature = "doc", feature = "ooxml", feature = "rtf", feature = "odf"))]
    #[inline]
    pub fn into_table(self) -> Option<Table> {
        match self {
            DocumentElement::Table(t) => Some(*t),
            _ => None,
        }
    }
}

#[cfg(all(test, feature = "ooxml"))]
mod tests {
    use super::*;

    fn paragraph_element() -> DocumentElement {
        DocumentElement::Paragraph(Box::new(Paragraph::Docx(
            crate::ooxml::docx::Paragraph::new(Vec::new()),
        )))
    }

    fn table_element() -> DocumentElement {
        DocumentElement::Table(Box::new(Table::Docx(Box::new(
            crate::ooxml::docx::Table::new(Vec::new()),
        ))))
    }

    #[test]
    fn test_document_element_is_paragraph() {
        let element = paragraph_element();
        assert!(element.is_paragraph());
        assert!(!element.is_table());
    }

    #[test]
    fn test_document_element_is_table() {
        let element = table_element();
        assert!(element.is_table());
        assert!(!element.is_paragraph());
    }

    #[test]
    fn test_document_element_as_paragraph() {
        let element = paragraph_element();
        assert!(element.as_paragraph().is_some());
        assert!(element.as_table().is_none());
    }

    #[test]
    fn test_document_element_as_table() {
        let element = table_element();
        assert!(element.as_table().is_some());
        assert!(element.as_paragraph().is_none());
    }

    #[test]
    fn test_document_element_into_paragraph() {
        let element = paragraph_element();
        assert!(element.into_paragraph().is_some());
    }

    #[test]
    fn test_document_element_into_table() {
        let element = table_element();
        assert!(element.into_table().is_some());
    }
}
