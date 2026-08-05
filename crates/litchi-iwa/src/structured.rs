//! Structured data extraction from iWork documents.
//!
//! Archive traversal is kept in focused private adapters. The public result
//! remains composed entirely of the semantic values owned by the leaf crates:
//! [`litchi_numbers::Table`], [`litchi_keynote::Slide`], and
//! [`litchi_pages::Section`].

mod keynote;
mod numbers;
mod pages;

use crate::Result;
use crate::bundle::Bundle;
use crate::object_index::ObjectIndex;
use litchi_keynote::Slide;
use litchi_numbers::Table;
use litchi_pages::Section;

/// Extract Numbers tables as canonical leaf-owned semantic values.
pub(crate) fn extract_tables(bundle: &Bundle, object_index: &ObjectIndex) -> Result<Vec<Table>> {
    numbers::extract(bundle, object_index)
}

/// Extract Keynote slides as canonical leaf-owned semantic values.
pub(crate) fn extract_slides(bundle: &Bundle, object_index: &ObjectIndex) -> Result<Vec<Slide>> {
    keynote::extract(bundle, object_index)
}

/// Extract the main Pages body section as a canonical leaf-owned value.
pub(crate) fn extract_sections(
    bundle: &Bundle,
    object_index: &ObjectIndex,
) -> Result<Vec<Section>> {
    pages::extract(bundle, object_index)
}

/// Extract all supported structured data from one document snapshot.
pub(crate) fn extract_all(bundle: &Bundle, object_index: &ObjectIndex) -> Result<StructuredData> {
    Ok(StructuredData {
        tables: extract_tables(bundle, object_index)?,
        slides: extract_slides(bundle, object_index)?,
        sections: extract_sections(bundle, object_index)?,
    })
}

/// Container for all structured data extracted from a document.
#[derive(Debug, Clone)]
pub struct StructuredData {
    /// Tables (primarily from Numbers).
    pub tables: Vec<Table>,
    /// Slides (from Keynote).
    pub slides: Vec<Slide>,
    /// Sections (from Pages).
    pub sections: Vec<Section>,
}

impl StructuredData {
    /// Check if any structured data was extracted.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.tables.is_empty() && self.slides.is_empty() && self.sections.is_empty()
    }

    /// Get summary statistics.
    #[must_use]
    pub fn summary(&self) -> String {
        format!(
            "Tables: {}, Slides: {}, Sections: {}",
            self.tables.len(),
            self.slides.len(),
            self.sections.len()
        )
    }

    /// Extract all text from all structured elements.
    ///
    /// The leaf snapshots expose borrowed text views, so this method appends
    /// directly into its single result vector instead of allocating a
    /// temporary vector through each leaf's convenience `all_text` method.
    #[must_use]
    pub fn all_text(&self) -> Vec<String> {
        let mut all_text = Vec::new();

        for table in &self.tables {
            all_text.push(format!("Table: {}", table.name()));
        }
        for slide in &self.slides {
            append_slide_text(&mut all_text, slide);
        }
        for section in &self.sections {
            append_section_text(&mut all_text, section);
        }

        all_text
    }
}

fn append_slide_text(output: &mut Vec<String>, slide: &Slide) {
    if let Some(title) = slide.title() {
        output.push(title.to_owned());
    }
    output.extend(slide.text_content().iter().cloned());
    if let Some(notes) = slide.notes() {
        output.push(notes.to_owned());
    }
    output.extend(
        slide
            .text_storages()
            .iter()
            .filter(|storage| !storage.is_empty())
            .map(|storage| storage.plain_text().to_owned()),
    );
}

fn append_section_text(output: &mut Vec<String>, section: &Section) {
    if let Some(heading) = &section.heading {
        output.push(heading.clone());
    }
    output.extend(section.paragraphs.iter().cloned());
    output.extend(
        section
            .text_storages
            .iter()
            .filter(|storage| !storage.is_empty())
            .map(|storage| storage.plain_text().to_owned()),
    );
}

#[cfg(test)]
mod tests;
