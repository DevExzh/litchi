//! Structured data model.
use super::{Section, Slide, Table};

/// Container for all structured data extracted from a document
#[derive(Debug, Clone)]
pub struct StructuredData {
    /// Tables (primarily from Numbers)
    pub tables: Vec<Table>,
    /// Slides (from Keynote)
    pub slides: Vec<Slide>,
    /// Sections (from Pages)
    pub sections: Vec<Section>,
}

impl StructuredData {
    /// Check if any structured data was extracted
    pub fn is_empty(&self) -> bool {
        self.tables.is_empty() && self.slides.is_empty() && self.sections.is_empty()
    }

    /// Get summary statistics
    pub fn summary(&self) -> String {
        format!(
            "Tables: {}, Slides: {}, Sections: {}",
            self.tables.len(),
            self.slides.len(),
            self.sections.len()
        )
    }

    /// Extract all text from all structured elements
    pub fn all_text(&self) -> Vec<String> {
        let mut all_text = Vec::new();

        // Add table names
        for table in &self.tables {
            all_text.push(format!("Table: {}", table.name));
        }

        // Add slide content
        for slide in &self.slides {
            all_text.extend(slide.all_text());
        }

        // Add section content
        for section in &self.sections {
            all_text.extend(section.all_text());
        }

        all_text
    }
}
