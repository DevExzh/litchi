use std::path::Path;
use std::sync::Arc;

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

/// A semantic Pages resource governed by [`SemanticLimits`].
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum SemanticLimitKind {
    /// Ordered sections decoded from the Pages body graph.
    Sections,
    /// Aggregate UTF-8 bytes retained by semantic section names and text.
    TextBytes,
}

impl std::fmt::Display for SemanticLimitKind {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter.write_str(match self {
            Self::Sections => "sections",
            Self::TextBytes => "text bytes",
        })
    }
}

/// An invalid caller-selected semantic Pages resource ceiling.
#[derive(Debug, Error, Clone, Copy, PartialEq, Eq)]
#[error("Pages semantic {kind} limit must be non-zero and no greater than {maximum}, got {value}")]
#[non_exhaustive]
pub struct SemanticLimitsError {
    /// Resource category whose requested limit is invalid.
    pub kind: SemanticLimitKind,
    /// Requested resource ceiling.
    pub value: usize,
    /// Format-wide hard ceiling for this resource.
    pub maximum: usize,
}

/// Checked resource ceilings for one archive-free Pages projection.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct SemanticLimits {
    sections: usize,
    text_bytes: usize,
}

/// Stable source-capture resource category for archive-free Pages ingress.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum DocumentSourceLimitKind {
    /// Encoded bytes supplied by the caller.
    InputBytes,
    /// Independently addressed items discovered in the source.
    Entries,
    /// Decoded bytes contributed by one source item.
    EntryBytes,
    /// Aggregate decoded bytes contributed by all source items.
    AggregateBytes,
    /// Bytes contributed by one document payload component.
    ComponentBytes,
}

impl std::fmt::Display for DocumentSourceLimitKind {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter.write_str(match self {
            Self::InputBytes => "input bytes",
            Self::Entries => "entries",
            Self::EntryBytes => "entry bytes",
            Self::AggregateBytes => "aggregate bytes",
            Self::ComponentBytes => "component bytes",
        })
    }
}

/// Invalid caller-selected archive-free source-capture limits.
#[derive(Debug, Error, Clone, Copy, PartialEq, Eq)]
#[error(
    "Pages document source {kind} limit must be non-zero and no greater than {maximum}, got {value}"
)]
#[non_exhaustive]
pub struct DocumentSourceLimitsError {
    /// Resource category whose requested limit is invalid.
    pub kind: DocumentSourceLimitKind,
    /// Requested resource ceiling.
    pub value: u64,
    /// Format-wide hard ceiling for this resource.
    pub maximum: u64,
}

/// Checked physical resource ceilings for archive-free Pages source capture.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct DocumentSourceLimits {
    input_bytes: u64,
    entries: usize,
    entry_bytes: u64,
    aggregate_bytes: u64,
    component_bytes: usize,
}

impl DocumentSourceLimits {
    /// Build a checked source-capture resource profile.
    ///
    /// # Errors
    ///
    /// Returns an error if any ceiling is zero or exceeds its hard maximum.
    pub fn new(
        max_input_bytes: u64,
        max_entries: usize,
        max_entry_bytes: u64,
        max_aggregate_bytes: u64,
        max_component_bytes: usize,
    ) -> std::result::Result<Self, DocumentSourceLimitsError> {
        check_source_limit(
            DocumentSourceLimitKind::InputBytes,
            max_input_bytes,
            litchi_iwa_detect::Limits::HARD_MAX_INPUT_BYTES,
        )?;
        check_source_limit(
            DocumentSourceLimitKind::Entries,
            usize_u64(max_entries),
            usize_u64(litchi_iwa_detect::Limits::HARD_MAX_FILES),
        )?;
        check_source_limit(
            DocumentSourceLimitKind::EntryBytes,
            max_entry_bytes,
            litchi_iwa_detect::Limits::HARD_MAX_ENTRY_SIZE,
        )?;
        check_source_limit(
            DocumentSourceLimitKind::AggregateBytes,
            max_aggregate_bytes,
            litchi_iwa_detect::Limits::HARD_MAX_TOTAL_SIZE,
        )?;
        check_source_limit(
            DocumentSourceLimitKind::ComponentBytes,
            usize_u64(max_component_bytes),
            usize_u64(litchi_iwa_detect::Limits::HARD_MAX_IWA_STREAM_SIZE),
        )?;
        Ok(Self {
            input_bytes: max_input_bytes,
            entries: max_entries,
            entry_bytes: max_entry_bytes,
            aggregate_bytes: max_aggregate_bytes,
            component_bytes: max_component_bytes,
        })
    }

    #[must_use]
    /// Maximum encoded bytes accepted from one source.
    pub const fn max_input_bytes(self) -> u64 {
        self.input_bytes
    }

    #[must_use]
    /// Maximum independently addressed source items accepted.
    pub const fn max_entries(self) -> usize {
        self.entries
    }

    #[must_use]
    /// Maximum decoded bytes accepted from one source item.
    pub const fn max_entry_bytes(self) -> u64 {
        self.entry_bytes
    }

    #[must_use]
    /// Maximum aggregate decoded bytes accepted from all source items.
    pub const fn max_aggregate_bytes(self) -> u64 {
        self.aggregate_bytes
    }

    #[must_use]
    /// Maximum bytes accepted from one document payload component.
    pub const fn max_component_bytes(self) -> usize {
        self.component_bytes
    }

    fn detector_limits(self) -> litchi_iwa_detect::Result<litchi_iwa_detect::Limits> {
        litchi_iwa_detect::Limits::new(
            self.input_bytes,
            self.entries,
            self.entry_bytes,
            self.aggregate_bytes,
            self.component_bytes,
        )
    }
}

impl Default for DocumentSourceLimits {
    fn default() -> Self {
        let limits = litchi_iwa_detect::Limits::default();
        Self {
            input_bytes: limits.max_input_bytes(),
            entries: limits.max_files(),
            entry_bytes: limits.max_entry_size(),
            aggregate_bytes: limits.max_total_size(),
            component_bytes: limits.max_iwa_stream_size(),
        }
    }
}

impl SemanticLimits {
    /// Hard ceiling for ordered semantic sections.
    pub const MAX_SECTIONS: usize = MAX_SECTIONS;
    /// Hard ceiling for aggregate retained semantic text.
    pub const MAX_TEXT_BYTES: usize = DEFAULT_MAX_TEXT_BYTES;

    /// Build a checked semantic resource profile.
    ///
    /// # Errors
    ///
    /// Returns an error if either ceiling is zero or exceeds its hard maximum.
    pub const fn new(
        max_sections: usize,
        max_text_bytes: usize,
    ) -> std::result::Result<Self, SemanticLimitsError> {
        if max_sections == 0 || max_sections > MAX_SECTIONS {
            return Err(SemanticLimitsError {
                kind: SemanticLimitKind::Sections,
                value: max_sections,
                maximum: MAX_SECTIONS,
            });
        }
        if max_text_bytes == 0 || max_text_bytes > DEFAULT_MAX_TEXT_BYTES {
            return Err(SemanticLimitsError {
                kind: SemanticLimitKind::TextBytes,
                value: max_text_bytes,
                maximum: DEFAULT_MAX_TEXT_BYTES,
            });
        }
        Ok(Self {
            sections: max_sections,
            text_bytes: max_text_bytes,
        })
    }

    /// Maximum ordered sections admitted by the projection.
    #[must_use]
    pub const fn max_sections(self) -> usize {
        self.sections
    }

    /// Maximum aggregate UTF-8 bytes retained by the projection.
    #[must_use]
    pub const fn max_text_bytes(self) -> usize {
        self.text_bytes
    }
}

impl Default for SemanticLimits {
    fn default() -> Self {
        Self {
            sections: MAX_SECTIONS,
            text_bytes: DEFAULT_MAX_TEXT_BYTES,
        }
    }
}

/// Physical and semantic profiles for archive-free Pages document ingress.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub struct DocumentReadOptions {
    source: DocumentSourceLimits,
    semantic: SemanticLimits,
}

/// Content-free filesystem failure category for archive-free ingress.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum IoKind {
    NotFound,
    PermissionDenied,
    AlreadyExists,
    InvalidInput,
    InvalidData,
    TimedOut,
    Interrupted,
    UnexpectedEof,
    Other,
}

impl std::fmt::Display for IoKind {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter.write_str(match self {
            Self::NotFound => "not found",
            Self::PermissionDenied => "permission denied",
            Self::AlreadyExists => "already exists",
            Self::InvalidInput => "invalid input",
            Self::InvalidData => "invalid data",
            Self::TimedOut => "timed out",
            Self::Interrupted => "interrupted",
            Self::UnexpectedEof => "unexpected end of input",
            Self::Other => "other I/O failure",
        })
    }
}

/// Content-free resource category reported by archive-free ingress.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum ReadLimitKind {
    /// Complete captured input bytes.
    InputBytes,
    /// Logical package entries.
    Entries,
    /// Logical package metadata bytes.
    MetadataBytes,
    /// Encoded or decoded semantic payload bytes.
    PayloadBytes,
    /// Fields or field-like records in semantic payloads.
    PayloadFields,
    /// Nested semantic payload traversal depth.
    PayloadNesting,
    /// Aggregate semantic payload traversal work.
    PayloadWork,
    /// Ordered Pages sections.
    Sections,
    /// Aggregate semantic UTF-8 text bytes.
    TextBytes,
    /// A future content-free resource category not known by this release.
    Other,
}

impl std::fmt::Display for ReadLimitKind {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter.write_str(match self {
            Self::InputBytes => "input bytes",
            Self::Entries => "entries",
            Self::MetadataBytes => "metadata bytes",
            Self::PayloadBytes => "payload bytes",
            Self::PayloadFields => "payload fields",
            Self::PayloadNesting => "payload nesting",
            Self::PayloadWork => "payload work",
            Self::Sections => "sections",
            Self::TextBytes => "text bytes",
            Self::Other => "other resource",
        })
    }
}

/// Errors raised while publishing an archive-free Pages document.
///
/// Display, Debug, and error-source chains contain only closed categories and
/// numeric bounds. Paths, package members, source identifiers, source content,
/// and lower-layer diagnostic strings are deliberately discarded.
#[derive(Debug, Error, Clone, PartialEq, Eq)]
#[non_exhaustive]
pub enum ReadError {
    #[error("Pages document I/O failure: {kind}")]
    Io { kind: IoKind },
    #[error("iWork source is not a Pages document")]
    NotPages,
    #[error("Pages document {kind} limit exceeded: observed {observed}, maximum {maximum}")]
    Limit {
        kind: ReadLimitKind,
        observed: u64,
        maximum: u64,
    },
    #[error("invalid Pages document source")]
    InvalidSource,
    #[error("invalid Pages document format")]
    InvalidFormat,
    #[error("Pages document allocation failed for {amount} units")]
    Allocation { amount: usize },
}

impl DocumentReadOptions {
    /// Combine checked source-capture and semantic resource profiles.
    #[must_use]
    pub const fn new(source: DocumentSourceLimits, semantic: SemanticLimits) -> Self {
        Self { source, semantic }
    }

    /// Return the bounded source-capture profile.
    #[must_use]
    pub const fn source(self) -> DocumentSourceLimits {
        self.source
    }

    /// Return the semantic projection profile.
    #[must_use]
    pub const fn semantic(self) -> SemanticLimits {
        self.semantic
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
    stats: Option<crate::Stats>,
}

#[derive(Debug, Clone, Default)]
pub struct Document {
    state: Arc<State>,
}

impl Document {
    /// Open a complete Pages package or an app-authored package directory.
    ///
    /// The path is captured once into an immutable prepared source and fully
    /// projected before this archive-free snapshot is published. Exact package
    /// bytes, source object identifiers, media, previews, and unsupported
    /// package members are not retained. Use [`crate::Package`] when exact
    /// complete-package preservation or editing is required.
    ///
    /// # Errors
    ///
    /// Returns a typed error if capture, format classification, or semantic
    /// projection fails or exceeds a checked resource ceiling.
    pub fn open(path: impl AsRef<Path>) -> std::result::Result<Self, ReadError> {
        Self::open_with_options(path, DocumentReadOptions::default())
    }

    /// Open an archive-free Pages snapshot under explicit resource profiles.
    ///
    /// # Errors
    ///
    /// Returns the same failures as [`Self::open`].
    pub fn open_with_options(
        path: impl AsRef<Path>,
        options: DocumentReadOptions,
    ) -> std::result::Result<Self, ReadError> {
        let source_limits = options
            .source()
            .detector_limits()
            .map_err(map_detection_error)?;
        let source =
            litchi_iwa_detect::PreparedSource::__from_path_with_pages_metadata(path, source_limits)
                .map_err(map_detection_error)?
                .ok_or_else(unrecognized_iwork_source)?;
        Self::from_prepared_source(source, options.semantic())
    }

    /// Decode borrowed packaged Pages bytes into an archive-free snapshot.
    ///
    /// # Errors
    ///
    /// Returns a typed capture, format, or semantic projection failure.
    pub fn from_bytes(bytes: &[u8]) -> std::result::Result<Self, ReadError> {
        Self::from_bytes_with_options(bytes, DocumentReadOptions::default())
    }

    /// Decode borrowed package bytes under explicit resource profiles.
    ///
    /// # Errors
    ///
    /// Returns the same failures as [`Self::from_bytes`].
    pub fn from_bytes_with_options(
        bytes: &[u8],
        options: DocumentReadOptions,
    ) -> std::result::Result<Self, ReadError> {
        let source_limits = options
            .source()
            .detector_limits()
            .map_err(map_detection_error)?;
        let source = litchi_iwa_detect::PreparedSource::__from_bytes_with_pages_metadata(
            bytes,
            source_limits,
        )
        .map_err(map_detection_error)?
        .ok_or_else(unrecognized_iwork_source)?;
        Self::from_prepared_source(source, options.semantic())
    }

    /// Decode already-shared immutable package bytes without copying them at
    /// the capture boundary.
    ///
    /// The input allocation is released after eager semantic projection.
    /// Use [`crate::Package`] when exact package bytes must remain authoritative.
    ///
    /// # Errors
    ///
    /// Returns a typed capture, format, or semantic projection failure.
    pub fn from_shared_bytes(bytes: Arc<[u8]>) -> std::result::Result<Self, ReadError> {
        Self::from_shared_bytes_with_options(bytes, DocumentReadOptions::default())
    }

    /// Decode shared immutable package bytes under explicit resource profiles.
    ///
    /// # Errors
    ///
    /// Returns the same failures as [`Self::from_shared_bytes`].
    pub fn from_shared_bytes_with_options(
        bytes: Arc<[u8]>,
        options: DocumentReadOptions,
    ) -> std::result::Result<Self, ReadError> {
        let source_limits = options
            .source()
            .detector_limits()
            .map_err(map_detection_error)?;
        let source = litchi_iwa_detect::PreparedSource::__from_shared_bytes_with_pages_metadata(
            bytes,
            source_limits,
        )
        .map_err(map_detection_error)?
        .ok_or_else(unrecognized_iwork_source)?;
        Self::from_prepared_source(source, options.semantic())
    }

    fn from_prepared_source(
        source: litchi_iwa_detect::PreparedSource,
        semantic: SemanticLimits,
    ) -> std::result::Result<Self, ReadError> {
        if source.format() != litchi_iwa_detect::Format::Pages {
            return Err(ReadError::NotPages);
        }
        crate::package::semantic_document_from_prepared_source(source, semantic)
            .map_err(map_package_read_error)
    }

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
            state.stats = Some(stats);
            return document;
        }
        Self {
            state: Arc::new(State {
                sections: Arc::clone(&document.state.sections),
                text_len: document.state.text_len,
                metadata: Some(metadata),
                stats: Some(stats),
            }),
        }
    }

    /// Capture another constant-time handle to the same semantic snapshot.
    #[must_use]
    pub fn snapshot(&self) -> Self {
        self.clone()
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
    pub fn section_at(&self, position: usize) -> SelectorResult<Option<&Section>> {
        self.select_section(SectionSelector::index(position))
    }

    /// Select one section by checked zero-based position.
    ///
    /// This compatibility helper delegates to [`Self::section_at`]. New code
    /// can use [`Self::select_section`] to share one name-or-position lookup
    /// boundary.
    #[must_use]
    pub fn section(&self, index: usize) -> Option<&Section> {
        self.section_at(index).ok().flatten()
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
    pub fn stats(&self) -> Option<crate::Stats> {
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

fn check_source_limit(
    kind: DocumentSourceLimitKind,
    value: u64,
    maximum: u64,
) -> std::result::Result<(), DocumentSourceLimitsError> {
    if value == 0 || value > maximum {
        return Err(DocumentSourceLimitsError {
            kind,
            value,
            maximum,
        });
    }
    Ok(())
}

fn map_package_read_error(error: crate::PackageError) -> ReadError {
    match error {
        crate::PackageError::Io(error) => ReadError::Io {
            kind: io_kind(error.kind()),
        },
        crate::PackageError::Archive(error) => map_archive_error(error),
        crate::PackageError::Detection(error) => map_detection_error(error),
        crate::PackageError::NotPages => ReadError::NotPages,
        crate::PackageError::InvalidFormat(_) => ReadError::InvalidFormat,
        crate::PackageError::Semantic(error) => map_semantic_read_error(error),
        crate::PackageError::SectionNamesTooLarge { observed, limit } => ReadError::Limit {
            kind: ReadLimitKind::TextBytes,
            observed: usize_u64(observed),
            maximum: usize_u64(limit),
        },
        crate::PackageError::PayloadLimit { observed, limit } => ReadError::Limit {
            kind: ReadLimitKind::PayloadBytes,
            observed: usize_u64(observed),
            maximum: usize_u64(limit),
        },
        crate::PackageError::ObjectLimit { observed, limit } => ReadError::Limit {
            kind: ReadLimitKind::PayloadFields,
            observed: usize_u64(observed),
            maximum: usize_u64(limit),
        },
        crate::PackageError::Allocation { amount } => ReadError::Allocation { amount },
    }
}

fn map_semantic_read_error(error: Error) -> ReadError {
    match error {
        Error::TooManySections { actual, limit } => ReadError::Limit {
            kind: ReadLimitKind::Sections,
            observed: usize_u64(actual),
            maximum: usize_u64(limit),
        },
        Error::TextTooLarge { observed, limit } => ReadError::Limit {
            kind: ReadLimitKind::TextBytes,
            observed: usize_u64(observed),
            maximum: usize_u64(limit),
        },
        Error::TooManyBodyStorages { actual, limit } => ReadError::Limit {
            kind: ReadLimitKind::PayloadFields,
            observed: usize_u64(actual),
            maximum: usize_u64(limit),
        },
        Error::InvalidSectionIndex { .. } => ReadError::InvalidFormat,
    }
}

fn map_detection_error(error: litchi_iwa_detect::Error) -> ReadError {
    match error {
        litchi_iwa_detect::Error::Io(error) => ReadError::Io {
            kind: io_kind(error.kind()),
        },
        litchi_iwa_detect::Error::IwaCore(error) => map_iwa_core_error(error),
        litchi_iwa_detect::Error::IwaCommon(error) => map_iwa_common_error(error),
        litchi_iwa_detect::Error::LimitExceeded {
            kind,
            observed,
            maximum,
        } => ReadError::Limit {
            kind: detection_limit_kind(kind),
            observed,
            maximum,
        },
        litchi_iwa_detect::Error::Allocation { amount } => ReadError::Allocation { amount },
        litchi_iwa_detect::Error::SourceChanged => ReadError::InvalidSource,
        litchi_iwa_detect::Error::InvalidFormat(_)
        | litchi_iwa_detect::Error::Archive(_)
        | litchi_iwa_detect::Error::InvalidLimits
        | litchi_iwa_detect::Error::Encrypted => ReadError::InvalidFormat,
        _ => ReadError::InvalidSource,
    }
}

fn map_archive_error(error: litchi_iwa_archive::Error) -> ReadError {
    match error {
        litchi_iwa_archive::Error::Io(error) => ReadError::Io {
            kind: io_kind(error.kind()),
        },
        litchi_iwa_archive::Error::Iwa(error) => map_iwa_core_error(error),
        litchi_iwa_archive::Error::Limit {
            kind,
            observed,
            maximum,
        } => ReadError::Limit {
            kind: archive_limit_kind(kind),
            observed,
            maximum,
        },
        litchi_iwa_archive::Error::Allocation { amount, .. } => ReadError::Allocation { amount },
        litchi_iwa_archive::Error::SourceChanged { .. }
        | litchi_iwa_archive::Error::DirectoryChanged { .. } => ReadError::InvalidSource,
        litchi_iwa_archive::Error::Zip { .. }
        | litchi_iwa_archive::Error::InvalidLimits(_)
        | litchi_iwa_archive::Error::Encrypted
        | litchi_iwa_archive::Error::Reassembly(_)
        | litchi_iwa_archive::Error::InvalidBundle(_) => ReadError::InvalidFormat,
    }
}

fn map_iwa_core_error(error: litchi_iwa_core::Error) -> ReadError {
    match error {
        litchi_iwa_core::Error::Io(error) => ReadError::Io {
            kind: io_kind(error.kind()),
        },
        litchi_iwa_core::Error::Limit {
            kind,
            observed,
            maximum,
        } => ReadError::Limit {
            kind: iwa_limit_kind(kind),
            observed: usize_u64(observed),
            maximum: usize_u64(maximum),
        },
        litchi_iwa_core::Error::Allocation { requested, .. } => {
            ReadError::Allocation { amount: requested }
        },
        litchi_iwa_core::Error::InvalidArchive { .. }
        | litchi_iwa_core::Error::InvalidLimits { .. }
        | litchi_iwa_core::Error::HeaderCodec { .. }
        | litchi_iwa_core::Error::Snappy { .. } => ReadError::InvalidFormat,
    }
}

fn map_iwa_common_error(error: litchi_iwa_common::Error) -> ReadError {
    match error {
        litchi_iwa_common::Error::LimitExceeded {
            kind,
            observed,
            limit,
        } => ReadError::Limit {
            kind: wire_limit_kind(kind),
            observed: usize_u64(observed),
            maximum: usize_u64(limit),
        },
        litchi_iwa_common::Error::Allocation { amount, .. } => ReadError::Allocation { amount },
        litchi_iwa_common::Error::InvalidFormat(_)
        | litchi_iwa_common::Error::InvalidLimit { .. } => ReadError::InvalidFormat,
    }
}

fn io_kind(kind: std::io::ErrorKind) -> IoKind {
    match kind {
        std::io::ErrorKind::NotFound => IoKind::NotFound,
        std::io::ErrorKind::PermissionDenied => IoKind::PermissionDenied,
        std::io::ErrorKind::AlreadyExists => IoKind::AlreadyExists,
        std::io::ErrorKind::InvalidInput => IoKind::InvalidInput,
        std::io::ErrorKind::InvalidData => IoKind::InvalidData,
        std::io::ErrorKind::TimedOut => IoKind::TimedOut,
        std::io::ErrorKind::Interrupted => IoKind::Interrupted,
        std::io::ErrorKind::UnexpectedEof => IoKind::UnexpectedEof,
        _ => IoKind::Other,
    }
}

fn detection_limit_kind(kind: litchi_iwa_detect::LimitKind) -> ReadLimitKind {
    match kind {
        litchi_iwa_detect::LimitKind::InputBytes => ReadLimitKind::InputBytes,
        litchi_iwa_detect::LimitKind::Entries => ReadLimitKind::Entries,
        litchi_iwa_detect::LimitKind::MetadataBytes => ReadLimitKind::MetadataBytes,
        litchi_iwa_detect::LimitKind::OutputBytes
        | litchi_iwa_detect::LimitKind::MemberNameBytes
        | litchi_iwa_detect::LimitKind::CompressedEntryBytes
        | litchi_iwa_detect::LimitKind::EntryBytes
        | litchi_iwa_detect::LimitKind::TotalBytes
        | litchi_iwa_detect::LimitKind::IwaStreamBytes
        | litchi_iwa_detect::LimitKind::IwaTotalBytes => ReadLimitKind::PayloadBytes,
        _ => ReadLimitKind::Other,
    }
}

fn archive_limit_kind(kind: litchi_iwa_archive::LimitKind) -> ReadLimitKind {
    match kind {
        litchi_iwa_archive::LimitKind::InputBytes => ReadLimitKind::InputBytes,
        litchi_iwa_archive::LimitKind::Entries => ReadLimitKind::Entries,
        litchi_iwa_archive::LimitKind::MetadataBytes => ReadLimitKind::MetadataBytes,
        litchi_iwa_archive::LimitKind::OutputBytes
        | litchi_iwa_archive::LimitKind::MemberNameBytes
        | litchi_iwa_archive::LimitKind::CompressedEntryBytes
        | litchi_iwa_archive::LimitKind::EntryBytes
        | litchi_iwa_archive::LimitKind::TotalBytes
        | litchi_iwa_archive::LimitKind::IwaStreamBytes
        | litchi_iwa_archive::LimitKind::IwaTotalBytes => ReadLimitKind::PayloadBytes,
    }
}

fn iwa_limit_kind(kind: litchi_iwa_core::LimitKind) -> ReadLimitKind {
    match kind {
        litchi_iwa_core::LimitKind::Objects
        | litchi_iwa_core::LimitKind::Messages
        | litchi_iwa_core::LimitKind::MessagesPerObject
        | litchi_iwa_core::LimitKind::HeaderFields
        | litchi_iwa_core::LimitKind::MetadataItems
        | litchi_iwa_core::LimitKind::SnappyFrames => ReadLimitKind::PayloadFields,
        litchi_iwa_core::LimitKind::HeaderNesting => ReadLimitKind::PayloadNesting,
        _ => ReadLimitKind::PayloadBytes,
    }
}

fn wire_limit_kind(kind: litchi_iwa_common::LimitKind) -> ReadLimitKind {
    match kind {
        litchi_iwa_common::LimitKind::InputBytes | litchi_iwa_common::LimitKind::OutputBytes => {
            ReadLimitKind::PayloadBytes
        },
        litchi_iwa_common::LimitKind::Fields => ReadLimitKind::PayloadFields,
        litchi_iwa_common::LimitKind::Nesting => ReadLimitKind::PayloadNesting,
        litchi_iwa_common::LimitKind::RewriteWork
        | litchi_iwa_common::LimitKind::TableRows
        | litchi_iwa_common::LimitKind::TableColumns
        | litchi_iwa_common::LimitKind::TableCells
        | litchi_iwa_common::LimitKind::MaterializedCells => ReadLimitKind::PayloadWork,
    }
}

fn usize_u64(value: usize) -> u64 {
    u64::try_from(value).unwrap_or(u64::MAX)
}

fn unrecognized_iwork_source() -> ReadError {
    ReadError::InvalidSource
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
                .section_at(1)
                .unwrap_or_else(|error| panic!("position should remain unambiguous: {error}"))
                .map(Section::index),
            Some(1)
        );
        assert_eq!(document.section(1).map(Section::index), Some(1));
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
        assert_eq!(document.section(0).and_then(Section::name), None);
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
