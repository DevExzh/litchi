//! Pages semantic value models.
//!
//! Package parsing, object topology, and mutation remain owned by the Pages
//! implementation. This crate owns the small, allocation-bearing section
//! model so it can be reused without pulling the IWA archive substrate into
//! every incremental build.

#![forbid(unsafe_code)]

/// Pages-specific package decoding over the shared IWA archive substrate.
pub mod package;

use litchi_iwa_text::TextStorage;

/// A logical Pages document section.
#[derive(Debug, Clone)]
pub struct Section {
    /// Zero-based section index.
    pub index: usize,
    /// Semantic kind of the section.
    pub section_type: SectionType,
    /// Optional section heading.
    pub heading: Option<String>,
    /// Paragraph values extracted from the section.
    pub paragraphs: Vec<String>,
    /// Rich-text storages belonging to the section.
    pub text_storages: Vec<TextStorage>,
    /// Number of pages represented by the section, when known.
    pub page_count: Option<usize>,
}

impl Section {
    /// Creates an empty section with `index` and `section_type`.
    #[must_use]
    pub fn new(index: usize, section_type: SectionType) -> Self {
        Self {
            index,
            section_type,
            heading: None,
            paragraphs: Vec::new(),
            text_storages: Vec::new(),
            page_count: None,
        }
    }

    /// Returns all non-empty text values in document order.
    #[must_use]
    pub fn all_text(&self) -> Vec<String> {
        let mut all = Vec::with_capacity(
            usize::from(self.heading.is_some())
                .saturating_add(self.paragraphs.len())
                .saturating_add(self.text_storages.len()),
        );
        if let Some(heading) = &self.heading {
            all.push(heading.clone());
        }
        all.extend(self.paragraphs.iter().cloned());

        all.extend(
            self.text_storages
                .iter()
                .filter(|storage| !storage.is_empty())
                .map(|storage| storage.plain_text().to_owned()),
        );

        all
    }

    /// Returns all section text joined with newlines.
    #[must_use]
    pub fn plain_text(&self) -> String {
        self.all_text().join("\n")
    }

    /// Returns whether the section has no modeled content.
    #[must_use]
    pub const fn is_empty(&self) -> bool {
        self.heading.is_none() && self.paragraphs.is_empty() && self.text_storages.is_empty()
    }
}

/// Semantic section kinds used by Pages.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum SectionType {
    /// Main body content.
    Body,
    /// Header content.
    Header,
    /// Footer content.
    Footer,
    /// Floating or anchored section content.
    Floating,
}

impl SectionType {
    /// Returns a stable human-readable name.
    #[must_use]
    pub const fn name(self) -> &'static str {
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
    fn section_creation_and_text() {
        let mut section = Section::new(0, SectionType::Body);
        assert_eq!(section.index, 0);
        assert_eq!(section.section_type, SectionType::Body);
        assert!(section.is_empty());

        section.heading = Some("Introduction".to_owned());
        section.paragraphs.push("First paragraph".to_owned());
        section
            .text_storages
            .push(TextStorage::from_text("Storage text".to_owned()));

        assert!(!section.is_empty());
        assert_eq!(
            section.all_text(),
            ["Introduction", "First paragraph", "Storage text"]
        );
        assert_eq!(
            section.plain_text(),
            "Introduction\nFirst paragraph\nStorage text"
        );
    }

    #[test]
    fn section_type_names_are_stable() {
        assert_eq!(SectionType::Body.name(), "Body");
        assert_eq!(SectionType::Header.name(), "Header");
        assert_eq!(SectionType::Footer.name(), "Footer");
        assert_eq!(SectionType::Floating.name(), "Floating");
    }
}
