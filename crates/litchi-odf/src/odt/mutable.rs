//! Mutable document structure for in-place modifications.
//!
//! This module provides a mutable wrapper around ODT documents that allows
//! for in-place modification of content, styles, and metadata.

use crate::core::{OdfStructure, OwnedPackage, PackageWriter};
use crate::elements::parser::DocumentOrderElement;
use crate::elements::table::Table;
use crate::elements::text::{Heading, List, Paragraph};
use crate::odt::Document;
use crate::odt::header_footer::{
    HeaderFooterKind, MasterPage, add_master_page, parse_master_pages, set_region_text,
    set_region_xml,
};
use crate::odt::page_layout::{PageLayout, parse_page_layouts, set_page_layout_xml};
use litchi_core::{Metadata, Result, xml::escape_xml};
use std::path::Path;

/// Document element type for tracking insertion order
#[derive(Debug, Clone)]
enum DocumentElement {
    /// A paragraph element
    Paragraph(Paragraph),
    /// A heading element
    Heading(Heading),
    /// A table element
    Table(Table),
    /// A text list element
    List(List),
}

/// A mutable ODT document that supports in-place modifications.
///
/// This struct wraps an ODT document and provides methods to modify its content,
/// including adding, updating, and removing paragraphs, tables, and other elements.
///
/// # Examples
///
/// ```no_run
/// use litchi_odf::{Document, MutableDocument};
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
    elements: Vec<DocumentElement>,
    /// Document metadata (mutable)
    metadata: Metadata,
    /// Original MIME type
    mimetype: String,
    /// Original styles XML (preserved as-is for now)
    styles_xml: Option<String>,
    /// Original package used to retain auxiliary package parts during rewriting.
    source_package: Option<OwnedPackage>,
}

impl MutableDocument {
    /// Create a mutable document from an existing Document.
    ///
    /// This parses the document structure into mutable elements.
    ///
    /// # Arguments
    ///
    /// * `doc` - The document to make mutable
    ///
    /// # Examples
    ///
    /// ```no_run
    /// use litchi_odf::{Document, MutableDocument};
    ///
    /// # fn main() -> litchi_core::Result<()> {
    /// let doc = Document::open("document.odt")?;
    /// let mut mutable_doc = MutableDocument::from_document(doc)?;
    /// # Ok(())
    /// # }
    /// ```
    pub fn from_document(doc: Document) -> Result<Self> {
        let source_elements = doc.elements()?;
        let metadata = doc.metadata()?;

        // Get MIME type from package
        let mimetype = "application/vnd.oasis.opendocument.text".to_string();

        // Extract styles XML from the document's package
        let styles_xml = doc
            .get_file("styles.xml")
            .ok()
            .and_then(|bytes| String::from_utf8(bytes).ok());

        let elements = source_elements
            .into_iter()
            .map(|element| match element {
                DocumentOrderElement::Paragraph(paragraph) => DocumentElement::Paragraph(paragraph),
                DocumentOrderElement::Heading(heading) => DocumentElement::Heading(heading),
                DocumentOrderElement::Table(table) => DocumentElement::Table(table),
                DocumentOrderElement::List(list) => DocumentElement::List(list),
            })
            .collect();
        let source_package = Some(doc.into_package());

        Ok(Self {
            elements,
            metadata,
            mimetype,
            styles_xml,
            source_package,
        })
    }

    /// Create a new empty mutable document.
    ///
    /// # Examples
    ///
    /// ```
    /// use litchi_odf::MutableDocument;
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
        }
    }

    /// Parse the document's master pages and current header/footer regions.
    pub fn master_pages(&self) -> Result<Vec<MasterPage>> {
        self.styles_xml
            .as_deref()
            .map_or_else(|| Ok(Vec::new()), parse_master_pages)
    }

    /// Parse automatic page layouts, their properties, and header/footer styles.
    pub fn page_layouts(&self) -> Result<Vec<PageLayout>> {
        self.styles_xml
            .as_deref()
            .map_or_else(|| Ok(Vec::new()), parse_page_layouts)
    }

    /// Replace one page layout with a complete XML fragment.
    ///
    /// The fragment must be exactly one self-contained `style:page-layout`
    /// element whose `style:name` matches `page_layout_name`. This supports all
    /// page properties and nested header/footer styles while preserving every
    /// unrelated byte in `styles.xml`.
    pub fn set_page_layout_xml(
        &mut self,
        page_layout_name: &str,
        page_layout_xml: &str,
    ) -> Result<()> {
        let styles = self.styles_xml.as_deref().ok_or_else(|| {
            litchi_core::Error::InvalidFormat(
                "document has no styles.xml page layout to modify".to_string(),
            )
        })?;
        self.styles_xml = Some(set_page_layout_xml(
            styles,
            page_layout_name,
            page_layout_xml,
        )?);
        Ok(())
    }

    /// Create or replace typed header/footer properties in one page layout.
    pub fn set_page_layout_header_footer_properties(
        &mut self,
        page_layout_name: &str,
        region: crate::PageHeaderFooterRegion,
        properties: &crate::HeaderFooterStyleProperties,
    ) -> Result<()> {
        let styles = self.styles_xml.as_deref().ok_or_else(|| litchi_core::Error::InvalidFormat("document has no styles.xml page layout to modify".to_string()))?;
        let layouts = parse_page_layouts(styles)?;
        let layout = layouts.iter().find(|layout| layout.name == page_layout_name).ok_or_else(|| litchi_core::Error::InvalidFormat(format!("page layout '{page_layout_name}' does not exist")))?;
        let replacement = crate::header_footer_properties::replace_page_layout_region_properties(layout, region, properties)?;
        self.styles_xml = Some(set_page_layout_xml(styles, page_layout_name, &replacement)?);
        Ok(())
    }

    /// Create or replace typed columns in one existing page layout.
    pub fn set_page_layout_columns(
        &mut self,
        page_layout_name: &str,
        columns: &crate::StyleColumns,
    ) -> Result<()> {
        let styles = self.styles_xml.as_deref().ok_or_else(|| {
            litchi_core::Error::InvalidFormat(
                "document has no styles.xml page layout to modify".to_string(),
            )
        })?;
        let layouts = parse_page_layouts(styles)?;
        let layout = layouts
            .iter()
            .find(|layout| layout.name == page_layout_name)
            .ok_or_else(|| {
                litchi_core::Error::InvalidFormat(format!(
                    "page layout '{page_layout_name}' does not exist"
                ))
            })?;
        let replacement = crate::style_columns::replace_page_layout_columns(layout, columns)?;
        self.styles_xml = Some(set_page_layout_xml(
            styles,
            page_layout_name,
            &replacement,
        )?);
        Ok(())
    }

    /// Create or replace the typed footnote separator in one existing page layout.
    pub fn set_page_layout_footnote_separator(
        &mut self,
        page_layout_name: &str,
        separator: &crate::StyleFootnoteSeparator,
    ) -> Result<()> {
        let styles = self.styles_xml.as_deref().ok_or_else(|| {
            litchi_core::Error::InvalidFormat(
                "document has no styles.xml page layout to modify".to_string(),
            )
        })?;
        let layouts = parse_page_layouts(styles)?;
        let layout = layouts
            .iter()
            .find(|layout| layout.name == page_layout_name)
            .ok_or_else(|| {
                litchi_core::Error::InvalidFormat(format!(
                    "page layout '{page_layout_name}' does not exist"
                ))
            })?;
        let replacement =
            crate::footnote_separator::replace_page_layout_footnote_separator(layout, separator)?;
        self.styles_xml = Some(set_page_layout_xml(
            styles,
            page_layout_name,
            &replacement,
        )?);
        Ok(())
    }

    /// Add an empty master page and its referenced page layout.
    /// Replace one existing named list level's modern label alignment.
    pub fn set_list_level_label_alignment(&mut self,item:&crate::ListStyleLevelLabelAlignment)->Result<()>{let styles=self.styles_xml.as_deref().ok_or_else(||litchi_core::Error::InvalidFormat("document has no styles.xml list style to modify".to_string()))?;self.styles_xml=Some(crate::list_label_alignment::replace_list_level_label_alignment_xml(styles,item)?);Ok(())}

    /// Add an empty master page and its referenced page layout.
    /// Replace, insert, or remove one existing paragraph style's direct drop cap.
    pub fn set_paragraph_style_drop_cap(&mut self, style: &crate::ParagraphStyleDropCap) -> Result<()> {
        let styles = self.styles_xml.as_deref().ok_or_else(|| litchi_core::Error::InvalidFormat(
            "document has no styles.xml paragraph style to modify".to_string()))?;
        self.styles_xml = Some(crate::paragraph_drop_cap::set_paragraph_style_drop_cap_xml(styles, style)?);
        Ok(())
    }

    /// Replace, insert, or remove typed row properties on an existing table-row style.
    pub fn set_table_row_style_properties(&mut self, style: &crate::TableRowStyleProperties) -> Result<()> {
        let styles = self.styles_xml.as_deref().ok_or_else(|| litchi_core::Error::InvalidFormat("document has no styles.xml table-row style to modify".to_string()))?;
        self.styles_xml = Some(crate::set_table_row_style_properties_xml(styles, style)?);
        Ok(())
    }

    /// Replace, insert, or remove typed properties on an existing table style.
    pub fn set_table_style_properties(&mut self, style: &crate::TableStyleProperties) -> Result<()> {
        let styles = self.styles_xml.as_deref().ok_or_else(|| litchi_core::Error::InvalidFormat("document has no styles.xml table style to modify".to_string()))?;
        self.styles_xml = Some(crate::set_table_style_properties_xml(styles, style)?);
        Ok(())
    }

    /// Add an empty master page and its referenced page layout.
    ///
    /// A minimal page layout is created in `office:automatic-styles` when a
    /// layout with `page_layout_name` does not already exist.
    pub fn add_master_page(&mut self, name: &str, page_layout_name: &str) -> Result<()> {
        let styles = self
            .styles_xml
            .get_or_insert_with(OdfStructure::default_styles_xml);
        *styles = add_master_page(styles, name, page_layout_name)?;
        Ok(())
    }

    /// Set plain text in one header/footer region of an existing master page.
    ///
    /// Only the selected region is rewritten; all unrelated style XML is preserved.
    pub fn set_header_footer_text(
        &mut self,
        master_page_name: &str,
        kind: HeaderFooterKind,
        text: &str,
    ) -> Result<()> {
        let styles = self.styles_xml.as_deref().ok_or_else(|| {
            litchi_core::Error::InvalidFormat(
                "document has no styles.xml master page to modify".to_string(),
            )
        })?;
        self.styles_xml = Some(set_region_text(styles, master_page_name, kind, Some(text))?);
        Ok(())
    }

    /// Replace one header/footer region with a complete XML fragment.
    ///
    /// The fragment must be exactly one self-contained `style:header`,
    /// `style:footer`, or corresponding first/left variant matching `kind`.
    /// This preserves rich text, fields, tables, drawings, and extension content.
    pub fn set_header_footer_xml(
        &mut self,
        master_page_name: &str,
        kind: HeaderFooterKind,
        xml: &str,
    ) -> Result<()> {
        let styles = self.styles_xml.as_deref().ok_or_else(|| {
            litchi_core::Error::InvalidFormat(
                "document has no styles.xml master page to modify".to_string(),
            )
        })?;
        self.styles_xml = Some(set_region_xml(styles, master_page_name, kind, xml)?);
        Ok(())
    }

    /// Remove one header/footer region from an existing master page.
    pub fn clear_header_footer(
        &mut self,
        master_page_name: &str,
        kind: HeaderFooterKind,
    ) -> Result<()> {
        let styles = self.styles_xml.as_deref().ok_or_else(|| {
            litchi_core::Error::InvalidFormat(
                "document has no styles.xml master page to modify".to_string(),
            )
        })?;
        self.styles_xml = Some(set_region_text(styles, master_page_name, kind, None)?);
        Ok(())
    }

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
    /// use litchi_odf::MutableDocument;
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
        self.elements.push(DocumentElement::Paragraph(para));
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
        self.elements.push(DocumentElement::Heading(heading));
        Ok(())
    }

    /// Add an existing list element to the end of the document.
    pub fn add_list(&mut self, list: List) -> Result<()> {
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
    /// use litchi_odf::MutableDocument;
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

    /// Remove a paragraph at a specific index.
    ///
    /// # Arguments
    ///
    /// * `index` - Index of the paragraph to remove (0-based)
    ///
    /// # Examples
    ///
    /// ```no_run
    /// use litchi_odf::MutableDocument;
    ///
    /// # fn main() -> litchi_core::Result<()> {
    /// let mut doc = MutableDocument::new();
    /// doc.add_paragraph("Paragraph 1")?;
    /// doc.add_paragraph("Paragraph 2")?;
    /// doc.remove_paragraph(0)?; // Remove first paragraph
    /// # Ok(())
    /// # }
    /// ```
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
    /// use litchi_odf::MutableDocument;
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
    /// use litchi_odf::MutableDocument;
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
    /// use litchi_odf::{MutableDocument, Table};
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
    /// use litchi_odf::{MutableDocument, Table};
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
        self.elements
            .retain(|elem| !matches!(elem, DocumentElement::Table(_)));
    }

    /// Clear all content (paragraphs and tables) from the document.
    pub fn clear_content(&mut self) {
        self.elements.clear();
    }

    /// Generate content.xml from the current mutable state.
    fn generate_content_xml(&self) -> String {
        let mut estimated = 256usize;
        estimated += self.elements.len() * 96;
        estimated += self
            .elements
            .iter()
            .map(|e| match e {
                DocumentElement::Paragraph(p) => p.text().map(|t| t.len()).unwrap_or(0),
                DocumentElement::Heading(h) => h.text().map(|t| t.len()).unwrap_or(0),
                DocumentElement::Table(_) => 256,
                DocumentElement::List(_) => 256,
            })
            .sum::<usize>();
        let mut body = String::with_capacity(estimated);

        // Add elements in their insertion order (paragraphs and tables mixed)
        for element in &self.elements {
            match element {
                DocumentElement::Paragraph(para) => {
                    let elem: crate::elements::element::Element = para.clone().into();
                    body.push_str(&elem.to_xml_string());
                },
                DocumentElement::Heading(heading) => {
                    let elem: crate::elements::element::Element = heading.clone().into();
                    body.push_str(&elem.to_xml_string());
                },
                DocumentElement::Table(table) => {
                    let elem: crate::elements::element::Element = table.clone().into();
                    body.push_str(&elem.to_xml_string());
                },
                DocumentElement::List(list) => {
                    let elem: crate::elements::element::Element = list.clone().into();
                    body.push_str(&elem.to_xml_string());
                },
            }
        }

        xml_minifier::minified_xml_format!(
            r#"<?xml version="1.0" encoding="UTF-8"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0" xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" xmlns:chart="urn:oasis:names:tc:opendocument:xmlns:chart:1.0" xmlns:dr3d="urn:oasis:names:tc:opendocument:xmlns:dr3d:1.0" xmlns:math="http://www.w3.org/1998/Math/MathML" xmlns:form="urn:oasis:names:tc:opendocument:xmlns:form:1.0" xmlns:script="urn:oasis:names:tc:opendocument:xmlns:script:1.0" office:version="1.3"><office:scripts/><office:font-face-decls/><office:automatic-styles/><office:body><office:text>{}</office:text></office:body></office:document-content>"#,
            body
        )
    }

    /// Generate meta.xml with current metadata.
    fn generate_meta_xml(&self) -> String {
        let now = chrono::Utc::now().to_rfc3339();
        let mut estimated = 64usize;
        estimated += self.metadata.title.as_ref().map(|s| s.len()).unwrap_or(0);
        estimated += self.metadata.author.as_ref().map(|s| s.len()).unwrap_or(0);
        estimated += self.metadata.subject.as_ref().map(|s| s.len()).unwrap_or(0);
        estimated += self
            .metadata
            .description
            .as_ref()
            .map(|s| s.len())
            .unwrap_or(0);
        estimated += self
            .metadata
            .keywords
            .as_ref()
            .map(|s| s.len())
            .unwrap_or(0);
        let mut meta_fields = String::with_capacity(estimated);

        // Add optional metadata fields
        if let Some(ref title) = self.metadata.title {
            let escaped_title = escape_xml(title);
            meta_fields.push_str(&xml_minifier::minified_xml_format!(
                r#"<dc:title>{}</dc:title>"#,
                escaped_title
            ));
        }

        if let Some(ref author) = self.metadata.author {
            let escaped_author = escape_xml(author);
            meta_fields.push_str(&xml_minifier::minified_xml_format!(
                r#"<dc:creator>{}</dc:creator>"#,
                escaped_author
            ));
        }

        if let Some(ref subject) = self.metadata.subject {
            let escaped_subject = escape_xml(subject);
            meta_fields.push_str(&xml_minifier::minified_xml_format!(
                r#"<dc:subject>{}</dc:subject>"#,
                escaped_subject
            ));
        }

        if let Some(ref description) = self.metadata.description {
            let escaped_description = escape_xml(description);
            meta_fields.push_str(&xml_minifier::minified_xml_format!(
                r#"<dc:description>{}</dc:description>"#,
                escaped_description
            ));
        }

        if let Some(ref keywords) = self.metadata.keywords {
            let escaped_keywords = escape_xml(keywords);
            meta_fields.push_str(&xml_minifier::minified_xml_format!(
                r#"<meta:keyword>{}</meta:keyword>"#,
                escaped_keywords
            ));
        }

        xml_minifier::minified_xml_format!(
            r#"<?xml version="1.0" encoding="UTF-8"?><office:document-meta xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0" office:version="1.3"><office:meta><meta:generator>Litchi/0.0.1</meta:generator><dc:date>{}</dc:date>{}</office:meta></office:document-meta>"#,
            now,
            meta_fields
        )
    }

    /// Save the modified document to a file.
    ///
    /// # Arguments
    ///
    /// * `path` - Path where the ODT file should be saved
    ///
    /// # Examples
    ///
    /// ```no_run
    /// use litchi_odf::MutableDocument;
    ///
    /// # fn main() -> litchi_core::Result<()> {
    /// let mut doc = MutableDocument::new();
    /// doc.add_paragraph("Hello!")?;
    /// doc.save("output.odt")?;
    /// # Ok(())
    /// # }
    /// ```
    pub fn save<P: AsRef<Path>>(&self, path: P) -> Result<()> {
        let bytes = self.to_bytes()?;
        std::fs::write(path, bytes)?;
        Ok(())
    }

    /// Convert the document to bytes.
    ///
    /// # Examples
    ///
    /// ```no_run
    /// use litchi_odf::MutableDocument;
    ///
    /// # fn main() -> litchi_core::Result<()> {
    /// let mut doc = MutableDocument::new();
    /// doc.add_paragraph("Hello!")?;
    /// let bytes = doc.to_bytes()?;
    /// # Ok(())
    /// # }
    /// ```
    pub fn to_bytes(&self) -> Result<Vec<u8>> {
        let mut writer = PackageWriter::new();

        // Set MIME type
        writer.set_mimetype(&self.mimetype)?;

        // Add content.xml (regenerated from mutable state)
        let content_xml = self.generate_content_xml();
        writer.add_file("content.xml", content_xml.as_bytes())?;

        // Add styles.xml (preserved or default)
        let default_styles = OdfStructure::default_styles_xml();
        let styles_xml = self.styles_xml.as_deref().unwrap_or(&default_styles);
        writer.add_file("styles.xml", styles_xml.as_bytes())?;

        // Add meta.xml (regenerated with current metadata)
        let meta_xml = self.generate_meta_xml();
        writer.add_file("meta.xml", meta_xml.as_bytes())?;

        if let Some(package) = &self.source_package {
            writer.copy_auxiliary_files_from(package)?;
        }

        writer.finish_to_bytes()
    }
}

impl Default for MutableDocument {
    fn default() -> Self {
        Self::new()
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::elements::parser::DocumentOrderElement;
    use crate::elements::table::{TableCell, TableRow};
    use crate::elements::text::{ListItem, Paragraph};
    use crate::odt::{DocumentBuilder, PageUsage};

    const MINIMAL_CONTENT: &str = r#"<?xml version="1.0"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"><office:body><office:text><text:p>Original</text:p></office:text></office:body></office:document-content>"#;

    fn source_document() -> Document {
        let mut builder = DocumentBuilder::new();
        builder.add_paragraph("Before table").unwrap();

        let mut table = Table::new();
        table.set_name("Data");
        let mut row = TableRow::new();
        let mut cell = TableCell::new();
        cell.set_text("Cell content");
        row.add_cell(cell);
        table.add_row(row);
        builder.add_table(table).unwrap();

        builder.add_heading("After table", 2).unwrap();
        builder
            .add_bulleted_list(vec!["First item", "Second item"])
            .unwrap();
        builder.add_paragraph("Document end").unwrap();

        Document::from_bytes(builder.build().unwrap()).unwrap()
    }

    fn element_kinds(document: &Document) -> Vec<&'static str> {
        document
            .elements()
            .unwrap()
            .iter()
            .map(|element| match element {
                DocumentOrderElement::Paragraph(_) => "paragraph",
                DocumentOrderElement::Heading(_) => "heading",
                DocumentOrderElement::Table(_) => "table",
                DocumentOrderElement::List(_) => "list",
            })
            .collect()
    }

    #[test]
    fn conversion_preserves_top_level_order_without_nested_paragraph_duplicates() {
        let mutable = MutableDocument::from_document(source_document()).unwrap();

        assert_eq!(mutable.paragraphs().len(), 2);
        assert_eq!(mutable.tables().len(), 1);
        assert_eq!(mutable.headings().len(), 1);
        assert_eq!(mutable.lists().len(), 1);
        assert_eq!(mutable.headings()[0].text().unwrap(), "After table");
        assert_eq!(mutable.headings()[0].level(), Some(2));
        assert_eq!(mutable.lists()[0].items().unwrap().len(), 2);
    }

    #[test]
    fn read_modify_write_keeps_paragraph_table_heading_and_list_order() {
        let mutable = MutableDocument::from_document(source_document()).unwrap();
        let round_trip = Document::from_bytes(mutable.to_bytes().unwrap()).unwrap();

        assert_eq!(
            element_kinds(&round_trip),
            ["paragraph", "table", "heading", "list", "paragraph"]
        );

        let elements = round_trip.elements().unwrap();
        let DocumentOrderElement::Table(table) = &elements[1] else {
            panic!("second element should remain a table");
        };
        assert_eq!(table.name(), Some("Data"));
        assert_eq!(
            table
                .row(0)
                .unwrap()
                .unwrap()
                .cell(0)
                .unwrap()
                .unwrap()
                .text()
                .unwrap(),
            "Cell content"
        );

        let DocumentOrderElement::List(list) = &elements[3] else {
            panic!("fourth element should remain a list");
        };
        let items = list.items().unwrap();
        assert_eq!(items[0].text().unwrap(), "First item");
        assert_eq!(items[1].text().unwrap(), "Second item");
    }

    #[test]
    fn mutable_document_can_add_headings_and_lists() {
        let mut document = MutableDocument::new();
        assert!(document.add_heading("Invalid", 0).is_err());
        document.add_heading("Title", 1).unwrap();

        let mut list = List::new();
        let mut item = ListItem::new();
        let mut paragraph = Paragraph::new();
        paragraph.set_text("Item");
        item.add_paragraph(paragraph);
        list.add_item(item);
        document.add_list(list).unwrap();

        let round_trip = Document::from_bytes(document.to_bytes().unwrap()).unwrap();
        assert_eq!(element_kinds(&round_trip), ["heading", "list"]);
    }

    #[test]
    fn read_modify_write_preserves_auxiliary_package_parts_and_media_types() {
        let mut writer = PackageWriter::new();
        writer
            .set_mimetype("application/vnd.oasis.opendocument.text")
            .unwrap();
        writer
            .add_file("content.xml", MINIMAL_CONTENT.as_bytes())
            .unwrap();
        writer
            .add_file("settings.xml", b"document settings")
            .unwrap();
        writer
            .add_file_with_media_type("Pictures/photo.bin", b"image", "image/x-test")
            .unwrap();
        writer
            .add_manifest_entry("Object 1/", "application/vnd.oasis.opendocument.text")
            .unwrap();
        writer
            .add_file("Object 1/content.xml", b"embedded object")
            .unwrap();
        writer
            .add_file_with_media_type(
                "custom/data.bin",
                b"custom payload",
                "application/x-litchi-test",
            )
            .unwrap();
        writer
            .add_file("META-INF/documentsignatures.xml", b"stale signature")
            .unwrap();

        let source_bytes = writer.finish_to_bytes().unwrap();
        let source = Document::from_bytes(source_bytes.clone()).unwrap();
        assert_eq!(source.to_bytes().unwrap(), source_bytes);
        let mut mutable = MutableDocument::from_document(source).unwrap();
        mutable.add_paragraph("Modified").unwrap();
        let output = crate::core::OwnedPackage::from_bytes(mutable.to_bytes().unwrap()).unwrap();

        assert_eq!(
            output.get_file("settings.xml").unwrap(),
            b"document settings"
        );
        assert_eq!(output.get_file("Pictures/photo.bin").unwrap(), b"image");
        assert_eq!(
            output.get_file("Object 1/content.xml").unwrap(),
            b"embedded object"
        );
        assert_eq!(
            output.get_file("custom/data.bin").unwrap(),
            b"custom payload"
        );
        assert!(!output.has_file("META-INF/documentsignatures.xml").unwrap());

        let package = output.package().unwrap();
        assert_eq!(
            package.manifest().get_media_type("Pictures/photo.bin"),
            Some("image/x-test")
        );
        assert_eq!(
            package.manifest().get_media_type("Object 1/"),
            Some("application/vnd.oasis.opendocument.text")
        );
        assert_eq!(
            package.manifest().get_media_type("custom/data.bin"),
            Some("application/x-litchi-test")
        );
    }

    #[test]
    fn edits_master_page_regions_through_the_public_mutable_document_api() {
        const STYLES: &str = r#"<?xml version="1.0"?><office:document-styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"><office:styles><style:style style:name="preserved"/></office:styles><office:automatic-styles><style:page-layout style:name="pm1" style:page-usage="left"><style:page-layout-properties fo:page-width="21cm" fo:page-height="29.7cm"/></style:page-layout></office:automatic-styles><office:master-styles><style:master-page style:name="Standard" style:page-layout-name="pm1"><style:header><text:p>Old header</text:p></style:header><style:footer><text:p>Old footer</text:p></style:footer><style:region-left/></style:master-page></office:master-styles></office:document-styles>"#;

        let mut writer = PackageWriter::new();
        writer
            .set_mimetype("application/vnd.oasis.opendocument.text")
            .unwrap();
        writer
            .add_file("content.xml", MINIMAL_CONTENT.as_bytes())
            .unwrap();
        writer.add_file("styles.xml", STYLES.as_bytes()).unwrap();
        let source = Document::from_bytes(writer.finish_to_bytes().unwrap()).unwrap();
        let mut mutable = MutableDocument::from_document(source).unwrap();
        let layouts = mutable.page_layouts().unwrap();
        assert_eq!(layouts[0].page_usage, PageUsage::Left);
        assert_eq!(
            layouts[0].properties.as_ref().unwrap().attribute(
                Some("urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"),
                "page-width",
            ),
            Some("21cm")
        );

        mutable
            .set_header_footer_text("Standard", HeaderFooterKind::Header, "New & <header>")
            .unwrap();
        mutable
            .clear_header_footer("Standard", HeaderFooterKind::Footer)
            .unwrap();
        let pages = mutable.master_pages().unwrap();
        assert_eq!(
            pages[0].region(HeaderFooterKind::Header).unwrap().text,
            "New & <header>"
        );
        assert!(pages[0].region(HeaderFooterKind::Footer).is_none());

        let output = OwnedPackage::from_bytes(mutable.to_bytes().unwrap()).unwrap();
        let styles = String::from_utf8(output.get_file("styles.xml").unwrap()).unwrap();
        assert!(styles.contains("<style:style style:name=\"preserved\"/>"));
        assert!(styles.contains("<style:region-left/>"));
        assert!(!styles.contains("Old footer"));
        let round_trip = Document::from_bytes(output.as_bytes().to_vec()).unwrap();
        assert_eq!(round_trip.page_layouts().unwrap(), layouts);
        assert_eq!(
            round_trip.master_pages().unwrap()[0]
                .region(HeaderFooterKind::Header)
                .unwrap()
                .text,
            "New & <header>"
        );
    }

    #[test]
    fn creates_a_master_page_and_header_in_a_new_document() {
        let mut mutable = MutableDocument::new();
        mutable.add_master_page("Standard", "pm1").unwrap();
        let layout = r#"<s:page-layout xmlns:s="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:f="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" s:name="pm1" s:page-usage="mirrored"><s:page-layout-properties f:page-width="21cm" f:page-height="29.7cm"/></s:page-layout>"#;
        mutable.set_page_layout_xml("pm1", layout).unwrap();
        mutable
            .set_header_footer_text("Standard", HeaderFooterKind::Header, "Created header")
            .unwrap();
        let rich = r#"<s:header xmlns:s="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:t="urn:oasis:names:tc:opendocument:xmlns:text:1.0"><t:p>Page <t:page-number/></t:p></s:header>"#;
        mutable
            .set_header_footer_xml("Standard", HeaderFooterKind::Header, rich)
            .unwrap();

        let round_trip = Document::from_bytes(mutable.to_bytes().unwrap()).unwrap();
        let pages = round_trip.master_pages().unwrap();
        let layouts = round_trip.page_layouts().unwrap();
        assert_eq!(layouts.len(), 1);
        assert_eq!(layouts[0].xml, layout);
        assert_eq!(layouts[0].page_usage, PageUsage::Mirrored);
        assert_eq!(pages.len(), 1);
        assert_eq!(pages[0].name, "Standard");
        assert_eq!(pages[0].page_layout_name.as_deref(), Some("pm1"));
        assert_eq!(
            pages[0].region(HeaderFooterKind::Header).unwrap().text,
            "Page "
        );
        assert_eq!(pages[0].region(HeaderFooterKind::Header).unwrap().xml, rich);
        let styles = String::from_utf8(round_trip.get_file("styles.xml").unwrap()).unwrap();
        assert!(styles.contains(layout));
    }
}
