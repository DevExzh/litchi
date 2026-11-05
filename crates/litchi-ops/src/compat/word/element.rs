//! Ordered document elements, moved from the facade's `document::element`.

use super::{Paragraph, Table};

/// An element in a document's natural content order: a paragraph or a
/// table, as [`DocumentView::elements`](crate::DocumentView::elements)
/// returns them.
///
/// **Scheduled for deletion in M1**, with the rest of [`crate::compat`].
#[derive(Debug, Clone)]
#[non_exhaustive]
pub enum DocumentElement {
    /// A paragraph (boxed to keep the enum small).
    Paragraph(Box<Paragraph>),
    /// A table (boxed to keep the enum small).
    Table(Box<Table>),
}

impl DocumentElement {
    /// Whether this element is a paragraph.
    ///
    /// Scheduled for deletion in M1.
    #[must_use]
    pub fn is_paragraph(&self) -> bool {
        matches!(self, Self::Paragraph(_))
    }

    /// Whether this element is a table.
    ///
    /// Scheduled for deletion in M1.
    #[must_use]
    pub fn is_table(&self) -> bool {
        matches!(self, Self::Table(_))
    }

    /// The paragraph, when this element is one.
    ///
    /// Scheduled for deletion in M1.
    #[must_use]
    pub fn as_paragraph(&self) -> Option<&Paragraph> {
        match self {
            Self::Paragraph(paragraph) => Some(paragraph),
            Self::Table(_) => None,
        }
    }

    /// The table, when this element is one.
    ///
    /// Scheduled for deletion in M1.
    #[must_use]
    pub fn as_table(&self) -> Option<&Table> {
        match self {
            Self::Table(table) => Some(table),
            Self::Paragraph(_) => None,
        }
    }

    /// Consume the element and return its paragraph, when it is one.
    ///
    /// Scheduled for deletion in M1.
    #[must_use]
    pub fn into_paragraph(self) -> Option<Paragraph> {
        match self {
            Self::Paragraph(paragraph) => Some(*paragraph),
            Self::Table(_) => None,
        }
    }

    /// Consume the element and return its table, when it is one.
    ///
    /// Scheduled for deletion in M1.
    #[must_use]
    pub fn into_table(self) -> Option<Table> {
        match self {
            Self::Table(table) => Some(*table),
            Self::Paragraph(_) => None,
        }
    }
}

#[cfg(all(test, feature = "docx"))]
mod tests {
    use super::*;

    fn paragraph_element() -> DocumentElement {
        DocumentElement::Paragraph(Box::new(Paragraph::Docx(litchi_docx::Paragraph::new(
            Vec::new(),
        ))))
    }

    fn table_element() -> DocumentElement {
        DocumentElement::Table(Box::new(Table::Docx(Box::new(litchi_docx::Table::new(
            Vec::new(),
        )))))
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
