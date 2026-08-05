//! Archive-free structured iWork results.
//!
//! The physical IWA adapter owns archive traversal and native decoding. This
//! crate owns only the immutable aggregation of semantic values from the
//! three concrete iWork format leaves, so it does not depend on protobufs,
//! package IDs, ZIP state, or the `litchi-iwa` facade.

#![forbid(unsafe_code)]

use litchi_keynote::Slide;
use litchi_numbers::Table;
use litchi_pages::Section;

/// Semantic values extracted from one or more iWork application families.
#[derive(Debug, Clone)]
pub struct StructuredData {
    /// Tables, primarily from Numbers documents.
    pub tables: Vec<Table>,
    /// Slides from Keynote documents.
    pub slides: Vec<Slide>,
    /// Sections from Pages documents.
    pub sections: Vec<Section>,
}

impl StructuredData {
    /// Return whether no semantic values were extracted.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.tables.is_empty() && self.slides.is_empty() && self.sections.is_empty()
    }

    /// Return deterministic summary counts for the contained values.
    #[must_use]
    pub fn summary(&self) -> String {
        format!(
            "Tables: {}, Slides: {}, Sections: {}",
            self.tables.len(),
            self.slides.len(),
            self.sections.len()
        )
    }

    /// Collect all human-readable text without materializing intermediate
    /// per-format collections.
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
            .map(|storage| storage.text().to_owned()),
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
            .map(|storage| storage.text().to_owned()),
    );
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn empty_data_is_empty_and_has_stable_summary() {
        let data = StructuredData {
            tables: Vec::new(),
            slides: Vec::new(),
            sections: Vec::new(),
        };

        assert!(data.is_empty());
        assert_eq!(data.summary(), "Tables: 0, Slides: 0, Sections: 0");
        assert!(data.all_text().is_empty());
    }

    #[test]
    fn text_aggregation_preserves_format_order() {
        let table = Table::new("Data", litchi_numbers::Dimensions::new(1, 1));
        let mut slide_builder = Slide::builder(0);
        slide_builder.set_title(Some("Title".to_owned()));
        slide_builder.push_text("Body".to_owned());
        let slide = slide_builder.build();
        let mut section = Section::new(0, litchi_pages::SectionType::Body);
        section.heading = Some("Heading".to_owned());

        let data = StructuredData {
            tables: vec![table],
            slides: vec![slide],
            sections: vec![section],
        };

        assert_eq!(data.all_text(), ["Table: Data", "Title", "Body", "Heading"]);
    }
}
