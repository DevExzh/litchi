//! Mutable document state and in-memory element operations.

use crate::core::OwnedPackage;
use crate::elements::table::Table;
use crate::elements::text::{Heading, Hyperlink, List, Paragraph};
use litchi_core::{Metadata, Result};

/// Document element type for tracking insertion order
#[derive(Debug, Clone)]
pub(super) enum DocumentElement {
    /// A paragraph element
    Paragraph(Paragraph),
    /// A heading element
    Heading(Heading),
    /// A table element
    Table(Table),
    /// A text list element
    List(List),
    /// A standalone drawing frame (image or text box) at body level
    Frame(crate::elements::element::Element),
}

/// A mutable ODT document that supports in-place modifications.
///
/// This struct wraps an ODT document and provides methods to modify its content,
/// including adding, updating, and removing paragraphs, tables, and other elements.
///
/// # Examples
///
/// ```no_run
/// use litchi_odt::{Document, MutableDocument};
///
/// # fn main() -> litchi_core::Result<()> {
/// // Open an existing document
/// let doc = Document::open("input.odt")?;
/// let mut mutable_doc = MutableDocument::from_document(doc)?;
///
/// // Modify the document
/// mutable_doc.add_paragraph("New paragraph")?;
/// mutable_doc.remove_paragraph(0)?;
///
/// // Save the modified document
/// mutable_doc.save("output.odt")?;
/// # Ok(())
/// # }
/// ```
pub struct MutableDocument {
    /// Document elements in insertion order (paragraphs and tables mixed)
    pub(super) elements: Vec<DocumentElement>,
    /// Document metadata (mutable)
    pub(super) metadata: Metadata,
    /// Original MIME type
    pub(super) mimetype: String,
    /// Original styles XML (preserved as-is for now)
    pub(super) styles_xml: Option<String>,
    /// Original package used to retain auxiliary package parts during rewriting.
    pub(super) source_package: Option<OwnedPackage>,
    /// Authoritative original content XML used by byte-preserving inline mutations.
    pub(super) content_xml: Option<String>,
    /// Authored picture payloads written into the package on save.
    pub(super) pending_images: Vec<crate::frame::Part>,
    /// Monotonic counter for authored frame names (1-based).
    pub(super) next_frame_number: usize,
}

impl MutableDocument {
    /// Create a new empty mutable document.
    ///
    /// # Examples
    ///
    /// ```
    /// use litchi_odt::MutableDocument;
    ///
    /// let doc = MutableDocument::new();
    /// ```
    pub fn new() -> Self {
        Self {
            elements: Vec::new(),
            metadata: Metadata::default(),
            mimetype: "application/vnd.oasis.opendocument.text".to_string(),
            styles_xml: None,
            source_package: None,
            content_xml: None,
            pending_images: Vec::new(),
            next_frame_number: 1,
        }
    }
}

impl MutableDocument {
    /// Get all paragraphs in the document.
    pub fn paragraphs(&self) -> Vec<&Paragraph> {
        self.elements
            .iter()
            .filter_map(|elem| {
                if let DocumentElement::Paragraph(p) = elem {
                    Some(p)
                } else {
                    None
                }
            })
            .collect()
    }

    /// Get all paragraphs (owned) in the document.
    pub fn paragraphs_owned(&self) -> Vec<Paragraph> {
        self.elements
            .iter()
            .filter_map(|elem| {
                if let DocumentElement::Paragraph(p) = elem {
                    Some(p.clone())
                } else {
                    None
                }
            })
            .collect()
    }

    /// Get all headings in the document.
    pub fn headings(&self) -> Vec<&Heading> {
        self.elements
            .iter()
            .filter_map(|element| match element {
                DocumentElement::Heading(heading) => Some(heading),
                _ => None,
            })
            .collect()
    }

    /// Get all lists in the document.
    pub fn lists(&self) -> Vec<&List> {
        self.elements
            .iter()
            .filter_map(|element| match element {
                DocumentElement::List(list) => Some(list),
                _ => None,
            })
            .collect()
    }

    /// Get all tables in the document.
    pub fn tables(&self) -> Vec<&Table> {
        self.elements
            .iter()
            .filter_map(|elem| {
                if let DocumentElement::Table(t) = elem {
                    Some(t)
                } else {
                    None
                }
            })
            .collect()
    }

    /// Get all tables (owned) in the document.
    pub fn tables_owned(&self) -> Vec<Table> {
        self.elements
            .iter()
            .filter_map(|elem| {
                if let DocumentElement::Table(t) = elem {
                    Some(t.clone())
                } else {
                    None
                }
            })
            .collect()
    }

    /// Get the document metadata.
    pub fn metadata(&self) -> &Metadata {
        &self.metadata
    }

    /// Get a mutable reference to the document metadata.
    pub fn metadata_mut(&mut self) -> &mut Metadata {
        &mut self.metadata
    }

    /// Add a new paragraph to the end of the document.
    ///
    /// # Arguments
    ///
    /// * `text` - Text content for the paragraph
    ///
    /// # Examples
    ///
    /// ```no_run
    /// use litchi_odt::MutableDocument;
    ///
    /// # fn main() -> litchi_core::Result<()> {
    /// let mut doc = MutableDocument::new();
    /// doc.add_paragraph("Hello, World!")?;
    /// doc.add_paragraph("Second paragraph")?;
    /// # Ok(())
    /// # }
    /// ```
    pub fn add_paragraph(&mut self, text: &str) -> Result<()> {
        let mut para = Paragraph::new();
        para.set_text(text);
        self.invalidate_content_xml();
        self.elements.push(DocumentElement::Paragraph(para));
        Ok(())
    }

    /// Append a paragraph containing one simple ODF hyperlink.
    ///
    /// The target is inert metadata and is never followed by the library.
    pub fn add_hyperlink(&mut self, href: impl AsRef<str>, text: impl AsRef<str>) -> Result<()> {
        let hyperlink = Hyperlink::with_href(href, text)?;
        self.add_hyperlink_element(hyperlink)
    }

    /// Append a paragraph containing a fully configured ODF hyperlink.
    pub fn add_hyperlink_element(&mut self, hyperlink: Hyperlink) -> Result<()> {
        let mut paragraph = Paragraph::new();
        paragraph.add_hyperlink(hyperlink)?;
        self.invalidate_content_xml();
        self.elements.push(DocumentElement::Paragraph(paragraph));
        Ok(())
    }

    /// Add a heading to the end of the document.
    pub fn add_heading(&mut self, text: &str, level: u8) -> Result<()> {
        if !(1..=6).contains(&level) {
            return Err(litchi_core::Error::InvalidFormat(
                "Heading level must be between 1 and 6".to_string(),
            ));
        }

        let mut heading = Heading::new(level);
        heading.set_text(text);
        self.invalidate_content_xml();
        self.elements.push(DocumentElement::Heading(heading));
        Ok(())
    }

    /// Add an existing list element to the end of the document.
    pub fn add_list(&mut self, list: List) -> Result<()> {
        self.invalidate_content_xml();
        self.elements.push(DocumentElement::List(list));
        Ok(())
    }

    /// Insert a paragraph at a specific index.
    ///
    /// # Arguments
    ///
    /// * `index` - Position to insert at (0-based)
    /// * `text` - Text content for the paragraph
    ///
    /// # Examples
    ///
    /// ```no_run
    /// use litchi_odt::MutableDocument;
    ///
    /// # fn main() -> litchi_core::Result<()> {
    /// let mut doc = MutableDocument::new();
    /// doc.add_paragraph("First")?;
    /// doc.add_paragraph("Third")?;
    /// doc.insert_paragraph(1, "Second")?;
    /// # Ok(())
    /// # }
    /// ```
    pub fn insert_paragraph(&mut self, index: usize, text: &str) -> Result<()> {
        let mut para = Paragraph::new();
        para.set_text(text);

        if index <= self.elements.len() {
            self.invalidate_content_xml();
            self.elements
                .insert(index, DocumentElement::Paragraph(para));
            Ok(())
        } else {
            Err(litchi_core::Error::InvalidFormat(format!(
                "Index {} out of bounds (length: {})",
                index,
                self.elements.len()
            )))
        }
    }
}

impl MutableDocument {
    pub fn remove_paragraph(&mut self, index: usize) -> Result<Paragraph> {
        // Find the index of the nth paragraph
        let mut para_count = 0;
        let mut element_index = None;

        for (i, elem) in self.elements.iter().enumerate() {
            if matches!(elem, DocumentElement::Paragraph(_)) {
                if para_count == index {
                    element_index = Some(i);
                    break;
                }
                para_count += 1;
            }
        }

        if let Some(idx) = element_index {
            self.invalidate_content_xml();
            if let DocumentElement::Paragraph(para) = self.elements.remove(idx) {
                Ok(para)
            } else {
                unreachable!()
            }
        } else {
            Err(litchi_core::Error::InvalidFormat(format!(
                "Paragraph index {} out of bounds (found {} paragraphs)",
                index, para_count
            )))
        }
    }

    /// Update a paragraph at a specific index.
    ///
    /// # Arguments
    ///
    /// * `index` - Index of the paragraph to update (0-based)
    /// * `text` - New text content
    ///
    /// # Examples
    ///
    /// ```no_run
    /// use litchi_odt::MutableDocument;
    ///
    /// # fn main() -> litchi_core::Result<()> {
    /// let mut doc = MutableDocument::new();
    /// doc.add_paragraph("Old text")?;
    /// doc.update_paragraph(0, "New text")?;
    /// # Ok(())
    /// # }
    /// ```
    pub fn update_paragraph(&mut self, index: usize, text: &str) -> Result<()> {
        // Find the index of the nth paragraph
        let mut para_count = 0;
        let mut element_index = None;

        for (i, elem) in self.elements.iter().enumerate() {
            if matches!(elem, DocumentElement::Paragraph(_)) {
                if para_count == index {
                    element_index = Some(i);
                    break;
                }
                para_count += 1;
            }
        }

        if let Some(idx) = element_index {
            self.invalidate_content_xml();
            if let DocumentElement::Paragraph(ref mut para) = self.elements[idx] {
                para.set_text(text);
                Ok(())
            } else {
                unreachable!()
            }
        } else {
            Err(litchi_core::Error::InvalidFormat(format!(
                "Paragraph index {} out of bounds (found {} paragraphs)",
                index, para_count
            )))
        }
    }

    /// Clear all paragraphs from the document.
    ///
    /// # Examples
    ///
    /// ```no_run
    /// use litchi_odt::MutableDocument;
    ///
    /// # fn main() -> litchi_core::Result<()> {
    /// let mut doc = MutableDocument::new();
    /// doc.add_paragraph("Paragraph 1")?;
    /// doc.add_paragraph("Paragraph 2")?;
    /// doc.clear_paragraphs();
    /// assert_eq!(doc.paragraphs().len(), 0);
    /// # Ok(())
    /// # }
    /// ```
    pub fn clear_paragraphs(&mut self) {
        self.invalidate_content_xml();
        self.elements
            .retain(|elem| !matches!(elem, DocumentElement::Paragraph(_)));
    }

    /// Add a table to the document.
    ///
    /// # Arguments
    ///
    /// * `table` - Table to add
    ///
    /// # Examples
    ///
    /// ```no_run
    /// use litchi_odt::MutableDocument;
    /// use litchi_odt::elements::table::Table;
    ///
    /// # fn main() -> litchi_core::Result<()> {
    /// let mut doc = MutableDocument::new();
    /// let mut table = Table::new();
    /// table.set_name("Table1");
    /// doc.add_table(table)?;
    /// # Ok(())
    /// # }
    /// ```
    pub fn add_table(&mut self, table: Table) -> Result<()> {
        self.invalidate_content_xml();
        self.elements.push(DocumentElement::Table(table));
        Ok(())
    }

    /// Remove a table at a specific index.
    ///
    /// # Arguments
    ///
    /// * `index` - Index of the table to remove (0-based)
    ///
    /// # Examples
    ///
    /// ```no_run
    /// use litchi_odt::MutableDocument;
    /// use litchi_odt::elements::table::Table;
    ///
    /// # fn main() -> litchi_core::Result<()> {
    /// let mut doc = MutableDocument::new();
    /// doc.add_table(Table::new())?;
    /// doc.remove_table(0)?;
    /// # Ok(())
    /// # }
    /// ```
    pub fn remove_table(&mut self, index: usize) -> Result<Table> {
        // Find the index of the nth table
        let mut table_count = 0;
        let mut element_index = None;

        for (i, elem) in self.elements.iter().enumerate() {
            if matches!(elem, DocumentElement::Table(_)) {
                if table_count == index {
                    element_index = Some(i);
                    break;
                }
                table_count += 1;
            }
        }

        if let Some(idx) = element_index {
            self.invalidate_content_xml();
            if let DocumentElement::Table(table) = self.elements.remove(idx) {
                Ok(table)
            } else {
                unreachable!()
            }
        } else {
            Err(litchi_core::Error::InvalidFormat(format!(
                "Table index {} out of bounds (found {} tables)",
                index, table_count
            )))
        }
    }

    /// Clear all tables from the document.
    pub fn clear_tables(&mut self) {
        self.invalidate_content_xml();
        self.elements
            .retain(|elem| !matches!(elem, DocumentElement::Table(_)));
    }

    /// Clear all content (paragraphs and tables) from the document.
    pub fn clear_content(&mut self) {
        self.invalidate_content_xml();
        self.elements.clear();
    }
}

impl Default for MutableDocument {
    fn default() -> Self {
        Self::new()
    }
}
