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
        let mut text = String::with_capacity(self.text_len());
        self.append_plain_text(&mut text);
        text
    }

    /// Returns whether the section has no modeled content.
    #[must_use]
    pub const fn is_empty(&self) -> bool {
        self.heading.is_none() && self.paragraphs.is_empty() && self.text_storages.is_empty()
    }

    /// Returns the UTF-8 byte length of the rendered section text.
    pub(crate) fn text_len(&self) -> usize {
        let mut length = 0usize;
        let mut values = 0usize;

        if let Some(heading) = &self.heading {
            length = length.saturating_add(heading.len());
            values = values.saturating_add(1);
        }
        for paragraph in &self.paragraphs {
            length = length.saturating_add(paragraph.len());
            values = values.saturating_add(1);
        }
        for storage in &self.text_storages {
            if !storage.is_empty() {
                length = length.saturating_add(storage.len());
                values = values.saturating_add(1);
            }
        }

        length.saturating_add(values.saturating_sub(1))
    }

    pub(crate) fn append_plain_text(&self, output: &mut String) {
        let mut first = true;
        if let Some(heading) = &self.heading {
            append_value(output, &mut first, heading);
        }
        for paragraph in &self.paragraphs {
            append_value(output, &mut first, paragraph);
        }
        for storage in &self.text_storages {
            if !storage.is_empty() {
                append_value(output, &mut first, storage.plain_text());
            }
        }
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

fn append_value(output: &mut String, first: &mut bool, value: &str) {
    if !*first {
        output.push('\n');
    }
    output.push_str(value);
    *first = false;
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
        assert_eq!(section.text_len(), section.plain_text().len());
    }

    #[test]
    fn section_type_names_are_stable() {
        assert_eq!(SectionType::Body.name(), "Body");
        assert_eq!(SectionType::Header.name(), "Header");
        assert_eq!(SectionType::Footer.name(), "Footer");
        assert_eq!(SectionType::Floating.name(), "Floating");
    }
}
