//! Pages Document Section Structure
//!
//! Pages documents are organized into sections, each with its own layout and content.

use crate::iwa::text::TextStorage;

/// Represents a section in a Pages document
#[derive(Debug, Clone)]
pub struct PagesSection {
    /// Section index (0-based)
    pub index: usize,
    /// Section type
    pub section_type: PagesSectionType,
    /// Section heading/title
    pub heading: Option<String>,
    /// Paragraphs in this section
    pub paragraphs: Vec<String>,
    /// Text storages in this section
    pub text_storages: Vec<TextStorage>,
    /// Page count in this section
    pub page_count: Option<usize>,
}

impl PagesSection {
    /// Create a new section
    pub fn new(index: usize, section_type: PagesSectionType) -> Self {
        Self {
            index,
            section_type,
            heading: None,
            paragraphs: Vec::new(),
            text_storages: Vec::new(),
            page_count: None,
        }
    }

    /// Get all text from the section (heading + paragraphs)
    pub fn all_text(&self) -> Vec<String> {
        let mut all = Vec::new();
        if let Some(ref heading) = self.heading {
            all.push(heading.clone());
        }
        all.extend(self.paragraphs.clone());

        // Also include text from storages
        for storage in &self.text_storages {
            let text = storage.plain_text();
            if !text.is_empty() {
                all.push(text.to_string());
            }
        }

        all
    }

    /// Get plain text content as a single string
    pub fn plain_text(&self) -> String {
        self.all_text().join("\n")
    }

    /// Check if section is empty
    pub fn is_empty(&self) -> bool {
        self.heading.is_none() && self.paragraphs.is_empty() && self.text_storages.is_empty()
    }
}

/// Types of sections in a Pages document
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum PagesSectionType {
    /// Main body content
    Body,
    /// Header section
    Header,
    /// Footer section
    Footer,
    /// Floating/anchored section
    Floating,
}

impl PagesSectionType {
    /// Get a human-readable name for the section type
    pub fn name(&self) -> &'static str {
        match self {
            Self::Body => "Body",
            Self::Header => "Header",
            Self::Footer => "Footer",
            Self::Floating => "Floating",
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_pages_section_creation() {
        let mut section = PagesSection::new(0, PagesSectionType::Body);
        assert_eq!(section.index, 0);
        assert_eq!(section.section_type, PagesSectionType::Body);
        assert!(section.is_empty());

        section.heading = Some("Introduction".to_string());
        section.paragraphs.push("First paragraph".to_string());

        assert!(!section.is_empty());
        let text = section.plain_text();
        assert!(text.contains("Introduction"));
        assert!(text.contains("First paragraph"));
    }

    #[test]
    fn test_section_type_names() {
        assert_eq!(PagesSectionType::Body.name(), "Body");
        assert_eq!(PagesSectionType::Header.name(), "Header");
        assert_eq!(PagesSectionType::Footer.name(), "Footer");
        assert_eq!(PagesSectionType::Floating.name(), "Floating");
    }

    #[test]
    fn test_all_text() {
        let mut section = PagesSection::new(0, PagesSectionType::Body);
        section.heading = Some("Title".to_string());
        section.paragraphs.push("Para 1".to_string());
        section.paragraphs.push("Para 2".to_string());
        section
            .text_storages
            .push(TextStorage::from_text("Storage text".to_string()));

        let all_text = section.all_text();
        assert_eq!(all_text.len(), 4);
        assert_eq!(all_text[0], "Title");
        assert_eq!(all_text[1], "Para 1");
        assert_eq!(all_text[2], "Para 2");
        assert_eq!(all_text[3], "Storage text");
    }
}
