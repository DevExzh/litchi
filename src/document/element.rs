//! Document element types for representing ordered content.

use super::{Paragraph, Table};

/// A document element that can be either a paragraph or a table.
///
/// This enum represents the natural order of elements as they appear in a document,
/// which is essential for proper Markdown conversion and other sequential operations.
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
///     match element {
///         litchi::DocumentElement::Paragraph(para) => {
///             println!("Paragraph: {}", para.text()?);
///         }
///         litchi::DocumentElement::Table(table) => {
///             println!("Table with {} rows", table.row_count()?);
///         }
///     }
/// }
/// # Ok::<(), litchi::common::Error>(())
/// ```
#[derive(Debug, Clone)]
pub enum DocumentElement {
    /// A paragraph element (boxed to reduce enum size)
    Paragraph(Box<Paragraph>),
    /// A table element (boxed to reduce enum size from 12KB to ~224 bytes)
    Table(Box<Table>),
}

impl DocumentElement {
    /// Check if this element is a paragraph.
    #[inline]
    pub fn is_paragraph(&self) -> bool {
        matches!(self, DocumentElement::Paragraph(_))
    }

    /// Check if this element is a table.
    #[inline]
    pub fn is_table(&self) -> bool {
        matches!(self, DocumentElement::Table(_))
    }

    /// Get a reference to the paragraph, if this is a paragraph element.
    ///
    /// Returns `None` if this is a table element.
    #[inline]
    pub fn as_paragraph(&self) -> Option<&Paragraph> {
        match self {
            DocumentElement::Paragraph(p) => Some(p),
            _ => None,
        }
    }

    /// Get a reference to the table, if this is a table element.
    ///
    /// Returns `None` if this is a paragraph element.
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
            _ => None,
        }
    }

    /// Consume this element and return the table, if this is a table element.
    ///
    /// Returns `None` if this is a paragraph element.
    #[inline]
    pub fn into_table(self) -> Option<Table> {
        match self {
            DocumentElement::Table(t) => Some(*t),
            _ => None,
        }
    }
}
