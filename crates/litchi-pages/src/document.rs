use thiserror::Error;

use crate::{Section, SectionType};
use litchi_iwa_text::storage::Storage;

/// Default maximum UTF-8 bytes retained by one semantic Pages document.
pub const DEFAULT_MAX_TEXT_BYTES: usize = 64 * 1024 * 1024;
/// Maximum number of ordered sections accepted by one semantic document.
pub const MAX_SECTIONS: usize = 4096;
/// Maximum number of text storages accepted by one semantic body.
pub const MAX_BODY_STORAGES: usize = 4096;

/// Errors raised while constructing a bounded Pages semantic model.
#[derive(Debug, Error, Clone, PartialEq, Eq)]
#[non_exhaustive]
pub enum Error {
    /// The semantic document contains more sections than the configured cap.
    #[error("Pages document contains {actual} sections; maximum is {limit}")]
    TooManySections {
        /// Number of supplied sections.
        actual: usize,
        /// Maximum accepted sections.
        limit: usize,
    },
    /// A section index is not the canonical zero-based position.
    #[error("Pages section index {actual} is not the expected index {expected}")]
    InvalidSectionIndex {
        /// Position in the supplied section sequence.
        expected: usize,
        /// Index stored by the section.
        actual: usize,
    },
    /// The semantic text budget would be exceeded.
    #[error("Pages document text exceeds {limit} bytes")]
    TextTooLarge {
        /// Maximum permitted UTF-8 bytes.
        limit: usize,
    },
    /// The body contains more independent text storages than the cap.
    #[error("Pages body contains {actual} text storages; maximum is {limit}")]
    TooManyBodyStorages {
        /// Number of supplied storages.
        actual: usize,
        /// Maximum accepted storages.
        limit: usize,
    },
}

/// Result type for bounded Pages semantic construction.
pub type Result<T> = std::result::Result<T, Error>;

/// Semantic body content detached from its native package object.
#[derive(Debug, Clone)]
pub struct Body {
    text_storages: Box<[Storage]>,
}

impl Body {
    /// Construct a body using [`DEFAULT_MAX_TEXT_BYTES`].
    ///
    /// # Errors
    ///
    /// Returns [`Error::TooManyBodyStorages`] or [`Error::TextTooLarge`] when
    /// the supplied values exceed the semantic bounds.
    pub fn new(text_storages: Vec<Storage>) -> Result<Self> {
        Self::with_max_text_bytes(text_storages, DEFAULT_MAX_TEXT_BYTES)
    }

    /// Construct a body under an explicit UTF-8 storage budget.
    ///
    /// # Errors
    ///
    /// Returns [`Error::TooManyBodyStorages`] or [`Error::TextTooLarge`] when
    /// the supplied values exceed the semantic bounds.
    pub fn with_max_text_bytes(text_storages: Vec<Storage>, max_text_bytes: usize) -> Result<Self> {
        if text_storages.len() > MAX_BODY_STORAGES {
            return Err(Error::TooManyBodyStorages {
                actual: text_storages.len(),
                limit: MAX_BODY_STORAGES,
            });
        }

        let text_len = text_storages.iter().try_fold(0usize, |length, storage| {
            length
                .checked_add(storage.len())
                .ok_or(Error::TextTooLarge {
                    limit: max_text_bytes,
                })
        })?;
        if text_len > max_text_bytes {
            return Err(Error::TextTooLarge {
                limit: max_text_bytes,
            });
        }

        Ok(Self {
            text_storages: text_storages.into_boxed_slice(),
        })
    }

    /// Borrow the body storages in source order.
    #[must_use]
    pub fn text_storages(&self) -> &[Storage] {
        &self.text_storages
    }

    /// Return the UTF-8 byte length before section separators are rendered.
    #[must_use]
    pub fn text_len(&self) -> usize {
        self.text_storages.iter().fold(0usize, |length, storage| {
            length.saturating_add(storage.len())
        })
    }

    /// Return whether every body storage is empty.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.text_storages.iter().all(Storage::is_empty)
    }

    fn into_section(self) -> Section {
        let mut section = Section::new(0, SectionType::Body);
        section.text_storages = self.text_storages.into_vec();
        section
    }
}

/// Semantic Pages root detached from its native protobuf representation.
#[derive(Debug, Clone, Default)]
pub struct Root {
    body: Option<Body>,
}

impl Root {
    /// Construct an empty root for a package without body text.
    #[must_use]
    pub const fn empty() -> Self {
        Self { body: None }
    }

    /// Construct a root with validated body content.
    #[must_use]
    pub fn with_body(body: Body) -> Self {
        Self { body: Some(body) }
    }

    /// Borrow the root body, if the package exposes one.
    #[must_use]
    pub fn body(&self) -> Option<&Body> {
        self.body.as_ref()
    }

    fn into_body(self) -> Option<Body> {
        self.body
    }
}

/// Immutable, archive-free Pages document snapshot.
#[derive(Debug, Clone, Default)]
pub struct Document {
    sections: Box<[Section]>,
    text_len: usize,
}

impl Document {
    /// Build an immutable document from a semantic root.
    ///
    /// # Errors
    ///
    /// Returns a typed error when the root's sections violate the canonical
    /// index or default text budget.
    pub fn from_root(root: Root) -> Result<Self> {
        Self::from_root_with_max_text_bytes(root, DEFAULT_MAX_TEXT_BYTES)
    }

    /// Build an immutable document from a root under an explicit text budget.
    ///
    /// # Errors
    ///
    /// Returns a typed error when the root's sections exceed the section or
    /// caller-selected text budget, or use a non-canonical index.
    pub fn from_root_with_max_text_bytes(root: Root, max_text_bytes: usize) -> Result<Self> {
        let sections = root
            .into_body()
            .map(Body::into_section)
            .into_iter()
            .collect();
        Self::from_sections_with_max_text_bytes(sections, max_text_bytes)
    }

    /// Build an immutable document from ordered semantic sections.
    ///
    /// # Errors
    ///
    /// Returns a typed error when the section count, indexes, or default text
    /// budget is invalid.
    pub fn from_sections(sections: Vec<Section>) -> Result<Self> {
        Self::from_sections_with_max_text_bytes(sections, DEFAULT_MAX_TEXT_BYTES)
    }

    /// Build an immutable document under an explicit text budget.
    ///
    /// # Errors
    ///
    /// Returns a typed error when the section count, indexes, or text budget is
    /// invalid.
    pub fn from_sections_with_max_text_bytes(
        sections: Vec<Section>,
        max_text_bytes: usize,
    ) -> Result<Self> {
        if sections.len() > MAX_SECTIONS {
            return Err(Error::TooManySections {
                actual: sections.len(),
                limit: MAX_SECTIONS,
            });
        }

        let mut text_len = 0usize;
        for (expected, section) in sections.iter().enumerate() {
            if section.index != expected {
                return Err(Error::InvalidSectionIndex {
                    expected,
                    actual: section.index,
                });
            }
            text_len = text_len
                .checked_add(section.text_len())
                .and_then(|length| length.checked_add(usize::from(expected != 0)))
                .ok_or(Error::TextTooLarge {
                    limit: max_text_bytes,
                })?;
        }

        if text_len > max_text_bytes {
            return Err(Error::TextTooLarge {
                limit: max_text_bytes,
            });
        }

        Ok(Self {
            sections: sections.into_boxed_slice(),
            text_len,
        })
    }

    /// Borrow all semantic sections in stable source order.
    #[must_use]
    pub fn sections(&self) -> &[Section] {
        &self.sections
    }

    /// Select one section by checked zero-based position.
    #[must_use]
    pub fn section(&self, index: usize) -> Option<&Section> {
        self.sections.get(index)
    }

    /// Return the number of semantic sections.
    #[must_use]
    pub const fn section_count(&self) -> usize {
        self.sections.len()
    }

    /// Return the UTF-8 byte length of the rendered document text.
    #[must_use]
    pub const fn text_len(&self) -> usize {
        self.text_len
    }

    /// Return whether the document has no semantic sections.
    #[must_use]
    pub const fn is_empty(&self) -> bool {
        self.sections.is_empty()
    }

    /// Render all sections in source order without an intermediate collection.
    #[must_use]
    pub fn plain_text(&self) -> String {
        let mut text = String::with_capacity(self.text_len);
        for (index, section) in self.sections.iter().enumerate() {
            if index != 0 {
                text.push('\n');
            }
            section.append_plain_text(&mut text);
        }
        text
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn body(text: &str) -> Body {
        Body::new(vec![Storage::from_text(text.to_owned())])
            .unwrap_or_else(|error| panic!("body should be valid: {error}"))
    }

    #[test]
    fn root_builds_an_immutable_document_without_wire_types() {
        let document = Document::from_root(Root::with_body(body("Pages body")))
            .unwrap_or_else(|error| panic!("document should be valid: {error}"));

        assert_eq!(document.section_count(), 1);
        assert_eq!(document.text_len(), "Pages body".len());
        assert_eq!(document.plain_text(), "Pages body");
        assert_eq!(
            document.section(0).map(Section::plain_text),
            Some("Pages body".to_owned())
        );
    }

    #[test]
    fn construction_rejects_over_budget_text_and_noncanonical_positions() {
        let oversized = Body::with_max_text_bytes(vec![Storage::from_text("12345".to_owned())], 4);
        assert!(matches!(oversized, Err(Error::TextTooLarge { limit: 4 })));

        let mut section = Section::new(1, SectionType::Body);
        section
            .text_storages
            .push(Storage::from_text("body".to_owned()));
        let invalid = Document::from_sections(vec![section]);
        assert!(matches!(
            invalid,
            Err(Error::InvalidSectionIndex {
                expected: 0,
                actual: 1
            })
        ));
    }

    #[test]
    fn empty_root_is_a_valid_empty_snapshot() {
        let document = Document::from_root(Root::empty())
            .unwrap_or_else(|error| panic!("empty root should be valid: {error}"));
        assert!(document.is_empty());
        assert_eq!(document.text_len(), 0);
        assert_eq!(document.plain_text(), "");
    }
}
