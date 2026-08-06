//! Top-level content projection and structural element operations.

use super::model::{DocumentElement, MutableDocument};
use crate::elements::table::Table;
use crate::elements::text::{Heading, Hyperlink, List, Paragraph};
use litchi_core::{Error, Result};

impl MutableDocument {
    /// Get all top-level paragraphs in document order.
    pub fn paragraphs(&self) -> Vec<&Paragraph> {
        self.elements
            .iter()
            .filter_map(|element| match element {
                DocumentElement::Paragraph(paragraph) => Some(paragraph),
                _ => None,
            })
            .collect()
    }

    /// Get all top-level paragraphs as owned values.
    pub fn paragraphs_owned(&self) -> Vec<Paragraph> {
        self.paragraphs().into_iter().cloned().collect()
    }

    /// Get all top-level headings in document order.
    pub fn headings(&self) -> Vec<&Heading> {
        self.elements
            .iter()
            .filter_map(|element| match element {
                DocumentElement::Heading(heading) => Some(heading),
                _ => None,
            })
            .collect()
    }

    /// Get all top-level lists in document order.
    pub fn lists(&self) -> Vec<&List> {
        self.elements
            .iter()
            .filter_map(|element| match element {
                DocumentElement::List(list) => Some(list),
                _ => None,
            })
            .collect()
    }

    /// Get all top-level tables in document order.
    pub fn tables(&self) -> Vec<&Table> {
        self.elements
            .iter()
            .filter_map(|element| match element {
                DocumentElement::Table(table) => Some(table),
                _ => None,
            })
            .collect()
    }

    /// Get all top-level tables as owned values.
    pub fn tables_owned(&self) -> Vec<Table> {
        self.tables().into_iter().cloned().collect()
    }

    /// Add a new plain paragraph to the end of the document.
    pub fn add_paragraph(&mut self, text: &str) -> Result<()> {
        let mut paragraph = Paragraph::new();
        paragraph.set_text(text);
        self.invalidate_content_xml();
        self.elements.push(DocumentElement::Paragraph(paragraph));
        Ok(())
    }

    /// Append a paragraph containing one simple ODF hyperlink.
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
            return Err(Error::InvalidFormat(
                "Heading level must be between 1 and 6".to_string(),
            ));
        }
        let mut heading = Heading::new(level);
        heading.set_text(text);
        self.invalidate_content_xml();
        self.elements.push(DocumentElement::Heading(heading));
        Ok(())
    }

    /// Add an existing list to the end of the document.
    pub fn add_list(&mut self, list: List) -> Result<()> {
        self.invalidate_content_xml();
        self.elements.push(DocumentElement::List(list));
        Ok(())
    }

    /// Insert a plain paragraph at a top-level structural position.
    pub fn insert_paragraph(&mut self, index: usize, text: &str) -> Result<()> {
        let mut paragraph = Paragraph::new();
        paragraph.set_text(text);
        if index > self.elements.len() {
            return Err(Error::InvalidFormat(format!(
                "Index {} out of bounds (length: {})",
                index,
                self.elements.len()
            )));
        }
        self.invalidate_content_xml();
        self.elements
            .insert(index, DocumentElement::Paragraph(paragraph));
        Ok(())
    }

    /// Remove one top-level paragraph and return its typed value.
    pub fn remove_paragraph(&mut self, index: usize) -> Result<Paragraph> {
        let element_index = nth_element_index(
            &self.elements,
            index,
            "Paragraph",
            "paragraphs",
            |element| matches!(element, DocumentElement::Paragraph(_)),
        )?;
        self.invalidate_content_xml();
        match self.elements.remove(element_index) {
            DocumentElement::Paragraph(paragraph) => Ok(paragraph),
            _ => unreachable!("paragraph index resolved to a non-paragraph"),
        }
    }

    /// Replace one top-level paragraph's plain text.
    pub fn update_paragraph(&mut self, index: usize, text: &str) -> Result<()> {
        let element_index = nth_element_index(
            &self.elements,
            index,
            "Paragraph",
            "paragraphs",
            |element| matches!(element, DocumentElement::Paragraph(_)),
        )?;
        self.invalidate_content_xml();
        match &mut self.elements[element_index] {
            DocumentElement::Paragraph(paragraph) => {
                paragraph.set_text(text);
                Ok(())
            },
            _ => unreachable!("paragraph index resolved to a non-paragraph"),
        }
    }

    /// Remove all top-level paragraphs while retaining other body elements.
    pub fn clear_paragraphs(&mut self) {
        self.invalidate_content_xml();
        self.elements
            .retain(|element| !matches!(element, DocumentElement::Paragraph(_)));
    }

    /// Add an existing table to the end of the document.
    pub fn add_table(&mut self, table: Table) -> Result<()> {
        self.invalidate_content_xml();
        self.elements.push(DocumentElement::Table(table));
        Ok(())
    }

    /// Remove one top-level table and return its typed value.
    pub fn remove_table(&mut self, index: usize) -> Result<Table> {
        let element_index =
            nth_element_index(&self.elements, index, "Table", "tables", |element| {
                matches!(element, DocumentElement::Table(_))
            })?;
        self.invalidate_content_xml();
        match self.elements.remove(element_index) {
            DocumentElement::Table(table) => Ok(table),
            _ => unreachable!("table index resolved to a non-table"),
        }
    }

    /// Remove all top-level tables while retaining other body elements.
    pub fn clear_tables(&mut self) {
        self.invalidate_content_xml();
        self.elements
            .retain(|element| !matches!(element, DocumentElement::Table(_)));
    }

    /// Remove all projected top-level body content.
    pub fn clear_content(&mut self) {
        self.invalidate_content_xml();
        self.elements.clear();
    }
}

fn nth_element_index(
    elements: &[DocumentElement],
    wanted: usize,
    label: &str,
    plural: &str,
    is_kind: impl Fn(&DocumentElement) -> bool,
) -> Result<usize> {
    let mut count = 0;
    for (element_index, element) in elements.iter().enumerate() {
        if is_kind(element) {
            if count == wanted {
                return Ok(element_index);
            }
            count += 1;
        }
    }
    Err(Error::InvalidFormat(format!(
        "{label} index {wanted} out of bounds (found {count} {plural})",
    )))
}
