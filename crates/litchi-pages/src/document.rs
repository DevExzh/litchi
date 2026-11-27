use std::sync::Arc;

use litchi_core::Position;

use thiserror::Error;

use crate::selector::{SectionSelector, SelectorError, SelectorResult};
use crate::{Section, SectionType};
use litchi_iwa_text::storage::Storage;

/// Maximum UTF-8 bytes retained by one semantic Pages document.
///
/// Caller-selected budgets may tighten this ceiling but cannot relax it.
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
    #[error("Pages document retains {observed} text bytes; maximum is {limit}")]
    TextTooLarge {
        /// Observed UTF-8 bytes, or [`usize::MAX`] when addition overflowed.
        observed: usize,
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

/// Deterministic semantic measurements retained for a source-backed Pages
/// document.
///
/// Native object counts are intentionally omitted: they describe the
/// physical package representation rather than this archive-free projection.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct DocumentStats {
    /// Number of semantic Pages sections.
    pub section_count: usize,
}

impl DocumentStats {
    /// Return the number of semantic Pages sections.
    #[must_use]
    pub const fn section_count(self) -> usize {
        self.section_count
    }
}

/// Semantic body content detached from its package object.
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
    /// the supplied values exceed the selected budget or hard semantic
    /// ceiling.
    pub fn with_max_text_bytes(text_storages: Vec<Storage>, max_text_bytes: usize) -> Result<Self> {
        let text_limit = effective_max_text_bytes(max_text_bytes);
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
                    observed: usize::MAX,
                    limit: text_limit,
                })
        })?;
        if text_len > text_limit {
            return Err(Error::TextTooLarge {
                observed: text_len,
                limit: text_limit,
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
        let mut builder = Section::builder(0, SectionType::Body);
        for storage in self.text_storages {
            builder.push_text_storage(storage);
        }
        builder.build()
    }
}

/// Semantic Pages root detached from its encoded representation.
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
///
/// Cloning a document shares its immutable section allocation instead of
/// cloning every semantic section.
#[derive(Debug, Default)]
struct State {
    sections: Arc<[Section]>,
    text_len: usize,
    metadata: Option<litchi_core::Metadata>,
    stats: Option<DocumentStats>,
}

#[derive(Debug, Clone, Default)]
pub struct Document {
    state: Arc<State>,
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
        let text_limit = effective_max_text_bytes(max_text_bytes);
        let sections = root
            .into_body()
            .map(Body::into_section)
            .into_iter()
            .collect();
        Self::from_sections_with_max_text_bytes(sections, text_limit)
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
    /// The input vector is validated against the section, index, and text
    /// bounds before it is consumed into one shared immutable allocation.
    ///
    /// # Errors
    ///
    /// Returns a typed error when the section count, indexes, or text budget is
    /// invalid.
    pub fn from_sections_with_max_text_bytes(
        sections: Vec<Section>,
        max_text_bytes: usize,
    ) -> Result<Self> {
        let text_limit = effective_max_text_bytes(max_text_bytes);
        if sections.len() > MAX_SECTIONS {
            return Err(Error::TooManySections {
                actual: sections.len(),
                limit: MAX_SECTIONS,
            });
        }

        let mut text_len = 0usize;
        let mut retained_text_len = 0usize;
        for (expected, section) in sections.iter().enumerate() {
            if section.index() != expected {
                return Err(Error::InvalidSectionIndex {
                    expected,
                    actual: section.index(),
                });
            }
            let section_text_len = section.checked_text_len().ok_or(Error::TextTooLarge {
                observed: usize::MAX,
                limit: text_limit,
            })?;
            text_len = text_len
                .checked_add(section_text_len)
                .and_then(|length| length.checked_add(usize::from(expected != 0)))
                .ok_or(Error::TextTooLarge {
                    observed: usize::MAX,
                    limit: text_limit,
                })?;
            retained_text_len =
                retained_text_len
                    .checked_add(section.checked_retained_text_len().ok_or(
                        Error::TextTooLarge {
                            observed: usize::MAX,
                            limit: text_limit,
                        },
                    )?)
                    .ok_or(Error::TextTooLarge {
                        observed: usize::MAX,
                        limit: text_limit,
                    })?;
        }

        if retained_text_len > text_limit {
            return Err(Error::TextTooLarge {
                observed: retained_text_len,
                limit: text_limit,
            });
        }

        Ok(Self {
            state: Arc::new(State {
                sections: Arc::from(sections.into_boxed_slice()),
                text_len,
                metadata: None,
                stats: None,
            }),
        })
    }

    pub(crate) fn from_source(
        mut document: Self,
        metadata: litchi_core::Metadata,
        stats: crate::Stats,
    ) -> Self {
        if let Some(state) = Arc::get_mut(&mut document.state) {
            state.metadata = Some(metadata);
            state.stats = Some(DocumentStats {
                section_count: stats.section_count(),
            });
            return document;
        }
        Self {
            state: Arc::new(State {
                sections: Arc::clone(&document.state.sections),
                text_len: document.state.text_len,
                metadata: Some(metadata),
                stats: Some(DocumentStats {
                    section_count: stats.section_count(),
                }),
            }),
        }
    }

    /// Borrow all semantic sections in stable source order.
    #[must_use]
    pub fn sections(&self) -> &[Section] {
        &self.state.sections
    }

    /// Clone the shared immutable section allocation.
    ///
    /// This is a constant-time snapshot operation; no section or text payload
    /// is cloned.
    #[must_use]
    pub fn shared_sections(&self) -> Arc<[Section]> {
        Arc::clone(&self.state.sections)
    }

    /// Select one section by exact semantic name or checked source position.
    ///
    /// Name matching is exact and case-sensitive. A missing name or an
    /// out-of-range position returns `Ok(None)`; no selector can index or
    /// expose a source-package identifier.
    ///
    /// # Errors
    ///
    /// Returns [`SelectorError::AmbiguousSectionName`] when more than one
    /// section has the requested exact name.
    pub fn select_section<'a, S>(&self, selector: S) -> SelectorResult<Option<&Section>>
    where
        S: Into<SectionSelector<'a>>,
    {
        match selector.into() {
            SectionSelector::Name(name) => {
                let mut matches = self
                    .state
                    .sections
                    .iter()
                    .enumerate()
                    .filter(|(_, section)| section.name() == Some(name));
                let Some((first, section)) = matches.next() else {
                    return Ok(None);
                };
                if let Some((duplicate, _)) = matches.next() {
                    return Err(SelectorError::AmbiguousSectionName {
                        name: name.into(),
                        first,
                        duplicate,
                    });
                }
                Ok(Some(section))
            },
            SectionSelector::Position(position) => Ok(self.state.sections.get(position.get())),
        }
    }

    /// Select one section by its exact semantic name.
    ///
    /// # Errors
    ///
    /// Returns [`SelectorError::AmbiguousSectionName`] when more than one
    /// section has the requested exact name.
    pub fn section_named(&self, name: &str) -> SelectorResult<Option<&Section>> {
        self.select_section(SectionSelector::Name(name))
    }

    /// Select one section by checked zero-based source position.
    ///
    /// # Errors
    ///
    /// Valid positions cannot fail. The result preserves the common selector
    /// contract for callers that switch between names and source positions.
    pub fn section_at(&self, position: Position) -> SelectorResult<Option<&Section>> {
        self.select_section(SectionSelector::position(position))
    }

    /// Return the number of semantic sections.
    #[must_use]
    pub fn section_count(&self) -> usize {
        self.state.sections.len()
    }

    /// Return the UTF-8 byte length of the rendered document text.
    #[must_use]
    pub fn text_len(&self) -> usize {
        self.state.text_len
    }

    /// Return whether the document has no semantic sections.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.state.sections.is_empty()
    }

    /// Render all sections in source order without an intermediate collection.
    #[must_use]
    pub fn plain_text(&self) -> String {
        let mut text = String::with_capacity(self.state.text_len);
        for (index, section) in self.state.sections.iter().enumerate() {
            if index != 0 {
                text.push('\n');
            }
            section.append_plain_text(&mut text);
        }
        text
    }

    /// Borrow metadata captured from a validated Pages source.
    ///
    /// Documents built from semantic values have no source diagnostics and
    /// return `None`.
    #[must_use]
    pub fn metadata(&self) -> Option<&litchi_core::Metadata> {
        self.state.metadata.as_ref()
    }

    /// Return source statistics retained during bounded ingress.
    ///
    /// Documents built from semantic values have no source diagnostics and
    /// return `None`.
    #[must_use]
    pub fn stats(&self) -> Option<DocumentStats> {
        self.state.stats
    }

    /// Revalidate the immutable semantic snapshot.
    ///
    /// # Errors
    ///
    /// Returns a typed semantic error if retained values violate the document
    /// invariants. Source bytes are never needed for this check.
    pub fn validate(&self) -> Result<()> {
        validate_sections(&self.state.sections, DEFAULT_MAX_TEXT_BYTES)
    }
}

fn validate_sections(sections: &[Section], text_limit: usize) -> Result<()> {
    if sections.len() > MAX_SECTIONS {
        return Err(Error::TooManySections {
            actual: sections.len(),
            limit: MAX_SECTIONS,
        });
    }
    let mut retained_text_len = 0usize;
    for (expected, section) in sections.iter().enumerate() {
        if section.index() != expected {
            return Err(Error::InvalidSectionIndex {
                expected,
                actual: section.index(),
            });
        }
        retained_text_len = retained_text_len
            .checked_add(
                section
                    .checked_retained_text_len()
                    .ok_or(Error::TextTooLarge {
                        observed: usize::MAX,
                        limit: text_limit,
                    })?,
            )
            .ok_or(Error::TextTooLarge {
                observed: usize::MAX,
                limit: text_limit,
            })?;
    }
    if retained_text_len > text_limit {
        return Err(Error::TextTooLarge {
            observed: retained_text_len,
            limit: text_limit,
        });
    }
    Ok(())
}

fn effective_max_text_bytes(requested: usize) -> usize {
    requested.min(DEFAULT_MAX_TEXT_BYTES)
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
            document.sections().first().map(Section::plain_text),
            Some("Pages body".to_owned())
        );
    }

    #[test]
    fn construction_rejects_over_budget_text_and_noncanonical_positions() {
        let oversized = Body::with_max_text_bytes(vec![Storage::from_text("12345".to_owned())], 4);
        assert!(matches!(
            oversized,
            Err(Error::TextTooLarge {
                observed: 5,
                limit: 4,
            })
        ));

        let much_larger = Body::with_max_text_bytes(vec![Storage::from_text("x".repeat(20))], 4);
        assert!(matches!(
            much_larger,
            Err(Error::TextTooLarge {
                observed: 20,
                limit: 4,
            })
        ));

        let mut section_builder = Section::builder(1, SectionType::Body);
        section_builder.push_text_storage(Storage::from_text("body".to_owned()));
        let invalid = Document::from_sections(vec![section_builder.build()]);
        assert!(matches!(
            invalid,
            Err(Error::InvalidSectionIndex {
                expected: 0,
                actual: 1
            })
        ));
    }

    #[test]
    fn retained_text_budget_excludes_rendered_separators_and_metadata_slots() {
        let mut first = Section::builder(0, SectionType::Body);
        first.push_text_storage(Storage::from_text("first".to_owned()));
        let mut second = Section::builder(1, SectionType::Body);
        second.push_text_storage(Storage::from_text("second".to_owned()));
        let document = Document::from_sections_with_max_text_bytes(
            vec![first.build(), second.build()],
            "firstsecond".len(),
        )
        .unwrap_or_else(|error| panic!("rendered separators retain no UTF-8 bytes: {error}"));
        assert_eq!(document.text_len(), "first\nsecond".len());
        assert_eq!(document.plain_text(), "first\nsecond");

        let mut named = Section::builder(0, SectionType::Body);
        named
            .set_name(Some("X"))
            .unwrap_or_else(|error| panic!("section name should be valid: {error}"));
        let named_document = Document::from_sections_with_max_text_bytes(vec![named.build()], 1)
            .unwrap_or_else(|error| panic!("only the owned name byte should be charged: {error}"));
        assert_eq!(named_document.sections()[0].name(), Some("X"));
        assert_eq!(named_document.text_len(), 0);
    }

    #[test]
    fn caller_budget_cannot_relax_hard_text_ceiling() {
        assert_eq!(effective_max_text_bytes(usize::MAX), DEFAULT_MAX_TEXT_BYTES);
        assert_eq!(effective_max_text_bytes(8), 8);
    }

    #[test]
    fn empty_root_is_a_valid_empty_snapshot() {
        let document = Document::from_root(Root::empty())
            .unwrap_or_else(|error| panic!("empty root should be valid: {error}"));
        assert!(document.is_empty());
        assert_eq!(document.text_len(), 0);
        assert_eq!(document.plain_text(), "");
    }

    #[test]
    fn cloned_documents_share_immutable_sections() {
        let document = Document::from_root(Root::with_body(body("Pages body")))
            .unwrap_or_else(|error| panic!("document should be valid: {error}"));
        let snapshot = document.clone();

        assert!(Arc::ptr_eq(
            &document.shared_sections(),
            &snapshot.shared_sections()
        ));
        assert_eq!(snapshot.plain_text(), "Pages body");
    }

    fn named_section(index: usize, name: &str) -> Section {
        let mut builder = Section::builder(index, SectionType::Body);
        builder
            .set_name(Some(name))
            .unwrap_or_else(|error| panic!("section name should be valid: {error}"));
        builder.build()
    }

    #[test]
    fn selector_lookup_is_exact_checked_and_raw_id_free() {
        let document = Document::from_sections(vec![
            named_section(0, "Introduction"),
            named_section(1, "Appendix"),
        ])
        .unwrap_or_else(|error| panic!("document should be valid: {error}"));

        assert_eq!(
            document
                .select_section("Appendix")
                .unwrap_or_else(|error| panic!("name should resolve: {error}"))
                .map(Section::index),
            Some(1)
        );
        assert_eq!(
            document
                .select_section(0)
                .unwrap_or_else(|error| panic!("position should resolve: {error}"))
                .and_then(Section::name),
            Some("Introduction")
        );
        assert!(
            document
                .select_section("introduction")
                .unwrap_or_else(|error| panic!("missing name is not an error: {error}"))
                .is_none()
        );
        assert!(
            document
                .select_section(usize::MAX)
                .unwrap_or_else(|error| panic!("out-of-range position is not an error: {error}"))
                .is_none()
        );
    }

    #[test]
    fn duplicate_exact_names_are_typed_ambiguity_errors() {
        let document = Document::from_sections(vec![
            named_section(0, "Repeated"),
            named_section(1, "Repeated"),
        ])
        .unwrap_or_else(|error| panic!("document should be valid: {error}"));

        assert!(matches!(
            document.section_named("Repeated"),
            Err(SelectorError::AmbiguousSectionName {
                name,
                first: 0,
                duplicate: 1
            }) if name.as_ref() == "Repeated"
        ));
        assert_eq!(
            document
                .section_at(Position::new(1))
                .unwrap_or_else(|error| panic!("position should remain unambiguous: {error}"))
                .map(Section::index),
            Some(1)
        );
        assert_eq!(document.sections().get(1).map(Section::index), Some(1));
    }

    #[test]
    fn name_selection_preserves_empty_unicode_and_selector_lifetimes() {
        let unnamed = Section::new(0, SectionType::Body);
        let document = Document::from_sections(vec![
            unnamed,
            named_section(1, ""),
            named_section(2, "Résumé"),
            named_section(3, "résumé"),
        ])
        .unwrap_or_else(|error| panic!("document should be valid: {error}"));

        assert_eq!(
            document
                .select_section("")
                .unwrap_or_else(|error| panic!("empty display name should resolve: {error}"))
                .map(Section::index),
            Some(1)
        );
        assert_eq!(document.sections().first().and_then(Section::name), None);
        assert_eq!(
            document
                .select_section("Résumé")
                .unwrap_or_else(|error| panic!("Unicode name should resolve exactly: {error}"))
                .map(Section::index),
            Some(2)
        );
        assert_eq!(
            document
                .select_section("résumé")
                .unwrap_or_else(|error| panic!("case-distinct name should resolve: {error}"))
                .map(Section::index),
            Some(3)
        );

        let selected = {
            let borrowed_name = String::from("Résumé");
            document
                .select_section(borrowed_name.as_str())
                .unwrap_or_else(|error| panic!("borrowed name should resolve: {error}"))
        };
        assert_eq!(selected.map(Section::index), Some(2));
    }
}
