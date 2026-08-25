//! Pages-native package ingress and semantic projection.
//!
//! This module is the only Pages boundary that understands ZIP/IWA packages
//! and generated protobuf messages. It publishes [`Package`] snapshots whose
//! semantic content is represented by the archive-free [`crate::Document`].

mod body_footnote;
pub(crate) mod body_table_dimension;
pub(crate) mod body_table_headers;
pub(crate) mod body_table_title;
pub(crate) mod document_settings;
mod footnote_text;
mod header_footer_text;
mod page_layout;
pub(crate) mod section_background;
mod section_name;
mod section_pagination;
pub(crate) mod section_settings;
mod section_text;
mod section_transaction;
mod table_lock;
#[cfg(feature = "internal-iwork-source")]
mod text_storage;

use std::borrow::Cow;
use std::fmt;
use std::fs::{Metadata as FileMetadata, OpenOptions};
use std::io::{Read, Write};
use std::num::NonZeroU64;
use std::path::Path;
use std::sync::Arc;

use litchi_iwa_archive::package::Catalog;
use litchi_iwa_archive::{ComponentCatalog, SourceCatalog};
use litchi_iwa_common::{
    WireLimits, decode_varint_from_bytes,
    wire::{WireFieldView, WireView},
};
use litchi_iwa_detect::{Format, PreparedSource};
use litchi_iwa_protos::pages_body_codec::{self, DecodeOptions as PagesBodyDecodeOptions};
use litchi_iwa_protos::{pages_footnote_codec, pages_footnote_marker_codec};
use litchi_iwa_text::storage::{Run, Storage};
use plist::Value;
use thiserror::Error;

use crate::footnote::body::{Footnote, Position};
use crate::selector::{SectionSelector, SelectorResult};
use crate::{
    Body, DEFAULT_MAX_TEXT_BYTES, Document, Error as SemanticError, MAX_BODY_STORAGES,
    MAX_SECTIONS, Root, Section, SectionType,
};

pub use body_footnote::{
    BodyFootnoteCommit, BodyFootnoteDiagnostics, BodyFootnoteEdit, BodyFootnoteError,
    BodyFootnoteLimitKind, BodyFootnotePatch,
};
pub use body_table_dimension::{
    BodyTableDimensionCommit, BodyTableDimensionDiagnostics, BodyTableDimensionEdit,
    BodyTableDimensionError, BodyTableDimensionLimitKind, BodyTableDimensionPatch,
};
pub use body_table_headers::{
    BodyTableHeaderSettingsCommit, BodyTableHeaderSettingsDiagnostics, BodyTableHeaderSettingsEdit,
    BodyTableHeaderSettingsError, BodyTableHeaderSettingsInvalidReason,
    BodyTableHeaderSettingsLimitKind, BodyTableHeaderSettingsPatch,
};
pub use body_table_title::{
    BodyTableTitleCommit, BodyTableTitleDiagnostics, BodyTableTitleEdit, BodyTableTitleError,
    BodyTableTitleLimitKind, BodyTableTitlePatch,
};
pub use footnote_text::{
    FootnoteTextCommit, FootnoteTextDiagnostics, FootnoteTextEdit, FootnoteTextError,
    FootnoteTextLimitKind, FootnoteTextPatch,
};
pub use header_footer_text::{
    HeaderFooterTextCommit, HeaderFooterTextDiagnostics, HeaderFooterTextEdit,
    HeaderFooterTextError, HeaderFooterTextLimitKind, HeaderFooterTextPatch,
};
pub use page_layout::{
    PageLayoutCommit, PageLayoutDiagnostics, PageLayoutEdit, PageLayoutError, PageLayoutLimitKind,
    PageLayoutPatch,
};
pub use section_name::{
    SectionNameCommit, SectionNameDiagnostics, SectionNameEdit, SectionNameError,
    SectionNameLimitKind, SectionNamePatch,
};
pub use section_pagination::{
    SectionPaginationCommit, SectionPaginationDiagnostics, SectionPaginationEdit,
    SectionPaginationError, SectionPaginationLimitKind, SectionPaginationPatch,
};
pub use section_text::{
    SectionTextCommit, SectionTextDiagnostics, SectionTextEdit, SectionTextError,
    SectionTextLimitKind, SectionTextPatch,
};
pub use table_lock::{
    BodyTableLockCommit, BodyTableLockDiagnostics, BodyTableLockEdit, BodyTableLockError,
    BodyTableLockLimitKind, BodyTableLockPatch,
};
const SECTION_MESSAGE_TYPE: u32 = 10_011;
const FOOTNOTE_REFERENCE_MESSAGE_TYPE: u32 = 2_008;
const TEXTUAL_ATTACHMENT_MESSAGE_TYPE: u32 = 2_004;
const FOOTNOTE_TABLE_FIELD: u32 = 16;
const TABLE_ENTRIES_FIELD: u32 = 1;
const TABLE_ATTACHMENT_FIELD: u32 = 9;
const STORAGE_KIND_FIELD: u32 = 1;
const STORAGE_TEXT_PREFIX: &str = "\u{fffc} ";
const FOOTNOTE_ANCHOR_UNIT: u16 = 0x000e;
const FOOTNOTE_MARK_KIND: i32 = 2;
const FOOTNOTE_STORAGE_KIND: u64 = 2;
const MAX_BODY_FOOTNOTES: usize = 4096;
/// Hard package-wide ceiling for native objects inspected by Pages ingress.
pub const MAX_OBJECTS: usize = 1_000_000;

/// Aggregate semantic bytes retained while projecting one rooted body-
/// footnote collection.
///
/// [`Footnote::with_custom_mark`] protects one value at a time.  The package
/// projection also needs to bound the complete collection before publishing
/// any owned strings, otherwise many individually-valid notes could exceed
/// the effective Pages text ceiling together.  The same small coordinator is
/// used by the package reader and the footnote-text selector path.
#[derive(Debug, Clone, Copy)]
pub(super) struct FootnoteSemanticBudget {
    maximum_bytes: usize,
    retained_bytes: usize,
}

impl FootnoteSemanticBudget {
    pub(super) const fn new(maximum_bytes: usize) -> Self {
        Self {
            maximum_bytes,
            retained_bytes: 0,
        }
    }

    pub(super) fn charge(
        &mut self,
        text_bytes: usize,
        custom_mark_bytes: usize,
    ) -> PackageResult<()> {
        let amount =
            text_bytes
                .checked_add(custom_mark_bytes)
                .ok_or(PackageError::PayloadLimit {
                    observed: usize::MAX,
                    limit: self.maximum_bytes,
                })?;
        let observed =
            self.retained_bytes
                .checked_add(amount)
                .ok_or(PackageError::PayloadLimit {
                    observed: usize::MAX,
                    limit: self.maximum_bytes,
                })?;
        if observed > self.maximum_bytes {
            return Err(PackageError::PayloadLimit {
                observed,
                limit: self.maximum_bytes,
            });
        }
        self.retained_bytes = observed;
        Ok(())
    }
}

/// Validate one native Pages text-storage payload through the focused,
/// source-borrowing adapter.
///
/// This is intentionally hidden from the supported semantic API. The legacy
/// migration host uses it while its editor remains in place; the focused
/// adapter owns the finite profile and the text-wire qualification itself.
#[cfg(feature = "internal-iwork-source")]
#[doc(hidden)]
pub fn __is_valid_pages_text_storage(source: &[u8]) -> bool {
    text_storage::is_valid(source)
}

/// Bounded physical ingress limits for a Pages package.
///
/// This is the shared iWork ZIP/IWA resource profile. It remains a separate
/// type because Pages owns the application parser, while physical validation
/// is shared by the three iWork format owners.
pub type Limits = litchi_iwa_archive::Limits;

/// Errors raised while opening or validating a native Pages package.
#[derive(Debug, Error)]
#[non_exhaustive]
pub enum PackageError {
    /// A filesystem operation failed while reading a Pages package.
    #[error("Pages package I/O error: {0}")]
    Io(#[from] std::io::Error),
    /// The shared physical ZIP/IWA ingress boundary rejected the package.
    #[error(transparent)]
    Archive(#[from] litchi_iwa_archive::Error),
    /// Focused iWork source capture or classification rejected the input.
    #[error(transparent)]
    Detection(#[from] litchi_iwa_detect::Error),
    /// The input is valid iWork data but belongs to another application.
    #[error("iWork package is not a Pages document")]
    NotPages,
    /// The parsed package cannot form a valid Pages-native document.
    #[error("invalid Pages package: {0}")]
    InvalidFormat(String),
    /// The native package decoded successfully but exceeded a Pages semantic
    /// bound while being projected into an immutable document.
    #[error(transparent)]
    Semantic(SemanticError),
    /// Section names exceed the aggregate retained-text budget.
    #[error("Pages section names require at least {observed} bytes; budget is {limit}")]
    SectionNamesTooLarge {
        /// Minimum bytes required by the names decoded so far.
        observed: usize,
        /// Aggregate retained UTF-8 budget.
        limit: usize,
    },
    /// A bounded semantic payload operation exceeded its finite profile.
    #[error("Pages payload limit exceeded: observed {observed}, maximum {limit}")]
    PayloadLimit { observed: usize, limit: usize },
    /// The package-wide native object inventory exceeded its finite profile.
    #[error("Pages object limit exceeded: observed {observed}, maximum {limit}")]
    ObjectLimit { observed: usize, limit: usize },
    /// A bounded semantic allocation failed before publication.
    #[error("Pages semantic allocation failed for {amount} units")]
    Allocation { amount: usize },
}

/// Result type returned by [`Package`] operations.
pub type PackageResult<T> = Result<T, PackageError>;

/// Failure while streaming an exact Pages package artifact to a caller-owned
/// sink.
///
/// Its `Display` and `Debug` representations report only the offset reached
/// by prior conforming successful writes and the sink error kind; they never
/// include package bytes or sink error text.
#[derive(Error)]
#[error("could not write Pages package after {bytes_written} bytes ({kind:?})")]
pub struct WriteError {
    error: std::io::Error,
    kind: std::io::ErrorKind,
    bytes_written: usize,
}

impl fmt::Debug for WriteError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("WriteError")
            .field("bytes_written", &self.bytes_written)
            .field("io_error_kind", &self.kind)
            .finish_non_exhaustive()
    }
}

impl WriteError {
    /// Return the byte offset reached by prior conforming successful writes.
    ///
    /// A `WriteZero` or trait-violating over-report is detected at this offset;
    /// it does not establish how many bytes that call's sink actually accepted.
    #[must_use]
    pub const fn bytes_written(&self) -> usize {
        self.bytes_written
    }

    /// Borrow the underlying sink error.
    #[must_use]
    pub const fn io_error(&self) -> &std::io::Error {
        &self.error
    }

    /// Consume this error and return the underlying sink error.
    #[must_use]
    pub fn into_io_error(self) -> std::io::Error {
        self.error
    }
}

/// Immutable statistics captured while a Pages package is opened.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Stats {
    total_objects: usize,
    section_count: usize,
}

impl Stats {
    /// Return the number of native IWA objects observed during source capture.
    #[must_use]
    pub const fn total_objects(self) -> usize {
        self.total_objects
    }

    /// Return the number of semantic Pages sections.
    #[must_use]
    pub const fn section_count(self) -> usize {
        self.section_count
    }
}

/// An immutable, cheaply clonable parsed Pages package.
///
/// The package retains one immutable physical/component source catalog for
/// validation, metadata, and future Pages-native capabilities. Its ordinary
/// read API exposes only the immutable semantic [`Document`], never raw object
/// identifiers or protobuf messages.
#[derive(Clone)]
pub struct Package {
    state: Arc<State>,
}

impl fmt::Debug for Package {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.debug_struct("Package").finish_non_exhaustive()
    }
}

#[derive(Debug)]
struct State {
    source: SourceCatalog,
    document: Document,
    metadata: Metadata,
    object_count: usize,
}

#[derive(Debug, Default)]
struct Metadata {
    title: Option<String>,
    author: Option<String>,
    keywords: Option<String>,
    description: Option<String>,
    application: Option<String>,
    revision: Option<String>,
    format_version: Option<String>,
    build_version: Option<String>,
    identifier: Option<String>,
}

#[derive(Debug, Clone, Copy)]
struct RootReferences {
    body: Option<NonZeroU64>,
    initial_section: Option<NonZeroU64>,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
struct NativeSectionReference {
    character_index: u32,
    identifier: NonZeroU64,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
struct NativeFootnoteReference {
    character_index: u32,
    identifier: NonZeroU64,
}

#[derive(Debug)]
struct BodyPreflight {
    section_references: Vec<NativeSectionReference>,
    fragment_count: usize,
    text_bytes: usize,
}

#[derive(Debug, Clone, Copy)]
struct TextRange {
    start: usize,
    end: usize,
}

#[derive(Debug, Clone, Copy)]
struct BoundaryPoint {
    byte_offset: usize,
    preceding_character: Option<char>,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
struct FileSnapshot {
    length: u64,
    modified: Option<std::time::SystemTime>,
    #[cfg(unix)]
    device: u64,
    #[cfg(unix)]
    inode: u64,
    #[cfg(unix)]
    mode: u32,
    #[cfg(unix)]
    modified_seconds: i64,
    #[cfg(unix)]
    modified_nanoseconds: i64,
    #[cfg(unix)]
    changed_seconds: i64,
    #[cfg(unix)]
    changed_nanoseconds: i64,
}

struct StorageAccumulator {
    text: String,
    runs: Vec<Run>,
}

enum StorageWireLimitsError {
    Physical(litchi_iwa_archive::Error),
    Wire(litchi_iwa_text_wire::RewriteError),
}

enum BodyStorageDecodeError {
    Package(PackageError),
    Wire(litchi_iwa_text_wire::RewriteError),
    SemanticLimit { observed: usize, limit: usize },
}

impl From<PackageError> for BodyStorageDecodeError {
    fn from(error: PackageError) -> Self {
        Self::Package(error)
    }
}

impl FileSnapshot {
    fn from_metadata(metadata: &FileMetadata) -> Self {
        #[cfg(unix)]
        use std::os::unix::fs::MetadataExt;

        Self {
            length: metadata.len(),
            modified: metadata.modified().ok(),
            #[cfg(unix)]
            device: metadata.dev(),
            #[cfg(unix)]
            inode: metadata.ino(),
            #[cfg(unix)]
            mode: metadata.mode(),
            #[cfg(unix)]
            modified_seconds: metadata.mtime(),
            #[cfg(unix)]
            modified_nanoseconds: metadata.mtime_nsec(),
            #[cfg(unix)]
            changed_seconds: metadata.ctime(),
            #[cfg(unix)]
            changed_nanoseconds: metadata.ctime_nsec(),
        }
    }
}

impl Package {
    /// Open and parse a Pages package from a regular filesystem file.
    ///
    /// # Errors
    ///
    /// Returns [`PackageError`] when the source is not a regular file, exceeds
    /// the default physical bounds, or cannot be decoded as a valid Pages
    /// package.
    pub fn open(path: impl AsRef<Path>) -> PackageResult<Self> {
        Self::open_with_limits(path, Limits::default())
    }

    /// Open and parse a Pages package under explicit physical ingress bounds.
    ///
    /// # Errors
    ///
    /// Returns [`PackageError`] when the source is not a regular file, exceeds
    /// the selected bounds, or cannot be decoded as a valid Pages package.
    pub fn open_with_limits(path: impl AsRef<Path>, limits: Limits) -> PackageResult<Self> {
        let source_bytes = read_path(path.as_ref(), limits)?;
        let source_catalog = SourceCatalog::from_shared_bytes_with_limits(source_bytes, limits)?;
        Self::from_source_catalog(source_catalog)
    }

    /// Parse a Pages package from ZIP bytes.
    ///
    /// # Errors
    ///
    /// Returns [`PackageError`] when the bytes exceed the default physical or
    /// semantic bounds, or cannot be decoded as a valid Pages package.
    pub fn from_bytes(bytes: &[u8]) -> PackageResult<Self> {
        Self::from_bytes_with_limits(bytes, Limits::default())
    }

    /// Parse a Pages package from ZIP bytes under explicit physical bounds.
    ///
    /// The selected physical profile also caps semantic text ingress at its
    /// IWA-stream maximum, which prevents native text materialization from
    /// exceeding either layer's budget.
    ///
    /// # Errors
    ///
    /// Returns [`PackageError`] when the bytes exceed a selected limit, the
    /// package shape is invalid, or semantic projection is invalid.
    pub fn from_bytes_with_limits(bytes: &[u8], limits: Limits) -> PackageResult<Self> {
        let source = SourceCatalog::from_bytes_with_limits(bytes, limits)?;
        Self::from_source_catalog(source)
    }

    /// Consume a source prepared by the focused iWork coordinator.
    ///
    /// This cross-crate ingress is intentionally unstable and exists only so
    /// the root iWork coordinator can dispatch one already parsed immutable
    /// package without repeating ZIP, Snappy, or IWA work.
    ///
    /// # Errors
    ///
    /// Returns [`PackageError`] when the prepared source belongs to another
    /// application or its Pages semantic projection is invalid.
    #[cfg(feature = "internal-iwork-source")]
    #[doc(hidden)]
    pub fn __from_prepared_source(source: PreparedSource) -> PackageResult<Self> {
        validate_prepared_format(&source)?;
        let source_catalog = source.__into_source_catalog().ok_or_else(|| {
            PackageError::InvalidFormat(
                "directory-backed Pages sources support semantic projection only".to_owned(),
            )
        })?;
        Self::from_source_catalog(source_catalog)
    }

    fn from_source_catalog(source: SourceCatalog) -> PackageResult<Self> {
        let limits = source.limits();
        let metadata = Metadata::from_catalog(source.package())?;
        let components = source.components();
        let object_count = validate_components(components)?;
        let root_references = root_references_with_limits(components, limits)?;
        let text_limit = effective_text_limit(limits);
        let document = decode_document(
            components,
            root_references,
            MAX_SECTIONS,
            text_limit,
            limits,
        )?;
        Ok(Self {
            state: Arc::new(State {
                source,
                document,
                metadata,
                object_count,
            }),
        })
    }

    /// Capture another cheap handle to the same native and semantic snapshot.
    #[must_use]
    pub fn snapshot(&self) -> Self {
        self.clone()
    }

    /// Borrow exact package bytes for crate-internal preservation logic.
    ///
    /// The returned bytes are the exact artifact represented by this
    /// snapshot, including unsupported ZIP members and unmodeled protobuf
    /// fields. Public callers should use [`Self::write_to`] to stream the
    /// artifact to a caller-owned sink.
    #[must_use]
    pub(crate) fn source_bytes(&self) -> &[u8] {
        self.state.source.source_bytes()
    }

    /// Write this exact immutable package artifact to a caller-owned sink.
    ///
    /// Unsupported ZIP members and unmodeled protobuf fields are emitted
    /// unchanged. This method streams the retained source once without
    /// allocating another package-sized buffer and does not flush `writer`.
    /// Partial writes may leave bytes in the caller-owned sink.
    ///
    /// Callers that need durable or atomic publication must provide that
    /// policy around the sink.
    ///
    /// # Errors
    ///
    /// Returns [`WriteError`] with the byte offset reached by prior conforming
    /// successful writes. A zero-length write or over-report is detected at
    /// that offset; an over-report does not establish its actual accepted-byte
    /// count.
    pub fn write_to<W: Write + ?Sized>(&self, writer: &mut W) -> Result<(), WriteError> {
        let source = self.source_bytes();
        let mut bytes_written = 0_usize;
        while bytes_written < source.len() {
            let remaining = &source[bytes_written..];
            match writer.write(remaining) {
                Ok(0) => {
                    return Err(WriteError {
                        error: std::io::Error::new(
                            std::io::ErrorKind::WriteZero,
                            "sink accepted no package bytes",
                        ),
                        kind: std::io::ErrorKind::WriteZero,
                        bytes_written,
                    });
                },
                Ok(amount) if amount <= remaining.len() => bytes_written += amount,
                Ok(_amount) => {
                    return Err(WriteError {
                        error: std::io::Error::new(
                            std::io::ErrorKind::InvalidData,
                            "sink reported accepting more bytes than supplied",
                        ),
                        kind: std::io::ErrorKind::InvalidData,
                        bytes_written,
                    });
                },
                Err(error) if error.kind() == std::io::ErrorKind::Interrupted => {},
                Err(error) => {
                    let kind = error.kind();
                    return Err(WriteError {
                        error,
                        kind,
                        bytes_written,
                    });
                },
            }
        }
        Ok(())
    }

    /// Render all native Pages text through the immutable semantic snapshot.
    ///
    /// # Errors
    ///
    /// This infallible projection is retained in a result for uniform
    /// document-reader ergonomics across formats.
    pub fn text(&self) -> PackageResult<String> {
        Ok(self.state.document.plain_text())
    }

    /// Borrow semantic Pages sections in stable source order.
    #[must_use]
    pub fn sections(&self) -> &[Section] {
        self.state.document.sections()
    }

    /// Read every footnote attached to the rooted Pages body in native source
    /// order.
    ///
    /// The projection exposes only checked UTF-16 positions, semantic text,
    /// and optional custom markers. Native attachment, storage, and marker
    /// identities remain private to this package adapter. The retained source
    /// catalog is never changed by this read.
    ///
    /// # Errors
    ///
    /// Returns [`PackageError::InvalidFormat`] when the body table, reference,
    /// footnote storage, or marker graph is malformed, or when a referenced
    /// object is missing. Returns a bounded semantic error when the projected
    /// note text or the aggregate footnote text/custom-marker collection
    /// exceeds the package text budget.
    pub fn body_footnotes(&self) -> PackageResult<Vec<Footnote>> {
        project_body_footnotes(
            self.state.source.components(),
            self.state.source.limits(),
            effective_text_limit(self.state.source.limits()),
        )
    }

    /// Select one semantic section by exact name or checked source position.
    ///
    /// This package-level convenience delegates to the immutable semantic
    /// snapshot. Native object identifiers, archive members, and protobuf
    /// payloads are not part of this lookup boundary.
    ///
    /// # Errors
    ///
    /// Returns [`crate::SelectorError::AmbiguousSectionName`] when more than
    /// one section has the requested exact name.
    pub fn select_section<'a, S>(&self, selector: S) -> SelectorResult<Option<&Section>>
    where
        S: Into<SectionSelector<'a>>,
    {
        self.state.document.select_section(selector)
    }

    /// Select one semantic section by its exact, case-sensitive name.
    ///
    /// This is the package-level counterpart to [`Document::section_named`].
    /// It keeps callers on the semantic selector surface while the package
    /// retains its native archive state privately.
    ///
    /// # Errors
    ///
    /// Returns [`crate::SelectorError::AmbiguousSectionName`] when more than
    /// one section has the requested exact name.
    pub fn section_named(&self, name: &str) -> SelectorResult<Option<&Section>> {
        self.select_section(SectionSelector::name(name))
    }

    /// Select one semantic section by its checked zero-based source position.
    ///
    /// A missing position returns `Ok(None)`, matching [`Document::section_at`]
    /// and keeping package callers independent of native object identifiers.
    pub fn section_at(&self, position: usize) -> SelectorResult<Option<&Section>> {
        self.select_section(SectionSelector::index(position))
    }

    /// Borrow the immutable Pages semantic snapshot.
    #[must_use]
    pub fn semantic_document(&self) -> &Document {
        &self.state.document
    }

    /// Project native Pages metadata into the format-neutral core model.
    #[must_use]
    pub fn metadata(&self) -> litchi_core::Metadata {
        self.state.metadata.to_core()
    }

    /// Revalidate the retained native object inventory and semantic snapshot.
    ///
    /// # Errors
    ///
    /// Returns [`PackageError::InvalidFormat`] if retained object identities
    /// or component structure violate package invariants.
    pub fn validate(&self) -> PackageResult<()> {
        let object_count = validate_components(self.state.source.components())?;
        if object_count != self.state.object_count {
            return Err(PackageError::InvalidFormat(
                "Pages package object inventory changed after parsing".to_owned(),
            ));
        }
        Ok(())
    }

    /// Return immutable package and semantic document statistics.
    #[must_use]
    pub fn stats(&self) -> Stats {
        Stats {
            total_objects: self.state.object_count,
            section_count: self.state.document.section_count(),
        }
    }
}

impl Metadata {
    fn from_catalog(catalog: &Catalog) -> PackageResult<Self> {
        let mut metadata = Self::default();

        if let Some(data) = metadata_entry(catalog, "Metadata/Properties.plist")? {
            metadata.apply_properties(data)?;
        }
        if let Some(data) = metadata_entry(catalog, "Metadata/BuildVersionHistory.plist")? {
            metadata.build_version = parse_build_version(data)?;
        }
        if let Some(data) = metadata_entry(catalog, "Metadata/DocumentIdentifier")? {
            metadata.apply_document_identifier(data)?;
        }

        Ok(metadata)
    }

    fn from_prepared_sidecars(
        sidecars: &litchi_iwa_detect::PreparedMetadataSidecars,
    ) -> PackageResult<Self> {
        let mut metadata = Self::default();
        if let Some(data) = sidecars.properties_plist() {
            metadata.apply_properties(data)?;
        }
        if let Some(data) = sidecars.build_version_history_plist() {
            metadata.build_version = parse_build_version(data)?;
        }
        if let Some(data) = sidecars.document_identifier() {
            metadata.apply_document_identifier(data)?;
        }
        Ok(metadata)
    }

    fn apply_document_identifier(&mut self, data: &[u8]) -> PackageResult<()> {
        let identifier_text = std::str::from_utf8(data).map_err(|error| {
            PackageError::InvalidFormat(format!(
                "Pages DocumentIdentifier is not valid UTF-8: {error}"
            ))
        })?;
        let identifier = identifier_text.trim();
        if identifier.is_empty() {
            return Err(PackageError::InvalidFormat(
                "Pages DocumentIdentifier must not be empty".to_owned(),
            ));
        }
        self.identifier = Some(identifier.to_owned());
        Ok(())
    }

    fn to_core(&self) -> litchi_core::Metadata {
        let revision = self.revision.clone().or_else(|| self.build_version.clone());
        let content_status = self
            .format_version
            .as_deref()
            .map(|version| format!("Pages Format Version {version}"));
        litchi_core::Metadata {
            title: self.title.clone(),
            author: self.author.clone(),
            keywords: self.keywords.clone(),
            description: self.description.clone(),
            application: Some(
                self.application
                    .clone()
                    .unwrap_or_else(|| "Pages".to_owned()),
            ),
            revision,
            content_status,
            identifier: self.identifier.clone(),
            ..Default::default()
        }
    }

    fn apply_properties(&mut self, data: &[u8]) -> PackageResult<()> {
        let value = Value::from_reader(std::io::Cursor::new(data)).map_err(|error| {
            PackageError::InvalidFormat(format!("failed to parse Pages Properties.plist: {error}"))
        })?;
        let Value::Dictionary(properties) = value else {
            return Err(PackageError::InvalidFormat(
                "Pages Properties.plist must contain a dictionary at its root".to_owned(),
            ));
        };

        self.title = property_string(&properties, "Title")
            .or_else(|| property_string(&properties, "kDocumentTitleKey"));
        self.author = property_string(&properties, "Author")
            .or_else(|| property_string(&properties, "kDocumentAuthorKey"))
            .or_else(|| property_string(&properties, "kSFWPAuthorPropertyKey"));
        self.keywords = property_string(&properties, "Keywords");
        self.description = property_string(&properties, "Comments");
        self.revision = property_string(&properties, "revision");
        self.format_version = property_string(&properties, "fileFormatVersion");

        if let Some(application_value) = properties.get("Application") {
            let Value::String(application) = application_value else {
                return Err(PackageError::InvalidFormat(
                    "Pages Properties.plist Application must be a string".to_owned(),
                ));
            };
            self.application = Some(application.clone());
        }

        Ok(())
    }
}

impl StorageAccumulator {
    fn with_capacity(capacity: usize) -> PackageResult<Self> {
        let mut text = String::new();
        text.try_reserve_exact(capacity).map_err(|_error| {
            PackageError::InvalidFormat("could not allocate Pages section text".to_owned())
        })?;
        Ok(Self {
            text,
            runs: Vec::new(),
        })
    }

    fn push(&mut self, text: &str) -> PackageResult<()> {
        self.runs.try_reserve(1).map_err(|_error| {
            PackageError::InvalidFormat("could not allocate Pages section text runs".to_owned())
        })?;
        let start = self.text.len();
        self.text.push_str(text);
        self.runs.push(Run::new(start, text.len()));
        Ok(())
    }

    fn push_empty(&mut self) -> PackageResult<()> {
        self.runs.try_reserve(1).map_err(|_error| {
            PackageError::InvalidFormat("could not allocate Pages section text runs".to_owned())
        })?;
        self.runs.push(Run::new(self.text.len(), 0));
        Ok(())
    }

    fn finish(self, body_identifier: NonZeroU64) -> PackageResult<Storage> {
        Storage::try_from_parts(self.text, self.runs).map_err(|error| {
            PackageError::InvalidFormat(format!(
                "Pages body object {body_identifier} semantic section text is invalid: {error}"
            ))
        })
    }
}

/// Consume a prepared Pages source into an archive-free semantic document.
///
/// This unstable coordinator handoff deliberately discards the physical
/// package catalog before semantic decoding begins. The caller-selected
/// section and text limits can tighten, but never relax, Pages' hard semantic
/// caps or the physical IWA-stream limit retained by the prepared source.
///
/// # Errors
///
/// Returns [`PackageError`] when the prepared source belongs to another iWork
/// application, its Pages graph is malformed, or projection exceeds an
/// effective admission bound.
#[cfg(feature = "internal-iwork-source")]
#[doc(hidden)]
pub fn __semantic_document_from_prepared_source(
    source: PreparedSource,
    max_sections: usize,
    max_text_bytes: usize,
) -> PackageResult<Document> {
    let semantic = crate::SemanticLimits::new(max_sections, max_text_bytes)
        .map_err(|error| PackageError::InvalidFormat(error.to_string()))?;
    semantic_document_from_prepared_source(source, semantic)
}

pub(crate) fn semantic_document_from_prepared_source(
    source: PreparedSource,
    semantic: crate::SemanticLimits,
) -> PackageResult<Document> {
    validate_prepared_format(&source)?;
    let (components, limits, sidecars) = source.__into_pages_semantic_source()?.__into_parts();
    // Parse every present canonical diagnostic before publishing any source
    // metadata. Failure therefore yields no partially populated Document.
    let metadata = Metadata::from_prepared_sidecars(&sidecars)?.to_core();
    let object_count = validate_components(&components)?;
    let root_references = root_references_with_limits(&components, limits)?;
    let document = decode_document(
        &components,
        root_references,
        semantic.max_sections(),
        semantic.max_text_bytes().min(effective_text_limit(limits)),
        limits,
    )?;
    let stats = Stats {
        total_objects: object_count,
        section_count: document.section_count(),
    };
    Ok(Document::from_source(document, metadata, stats))
}

fn validate_prepared_format(source: &PreparedSource) -> PackageResult<()> {
    if source.format() != Format::Pages {
        return Err(PackageError::NotPages);
    }
    Ok(())
}

#[cfg(any(unix, windows))]
fn read_path(path: &Path, limits: Limits) -> PackageResult<Arc<[u8]>> {
    let mut options = OpenOptions::new();
    options.read(true);
    #[cfg(unix)]
    {
        use std::os::unix::fs::OpenOptionsExt;

        // Keep the final component pinned and prevent a FIFO from blocking
        // before descriptor metadata can reject it.
        options.custom_flags(libc::O_CLOEXEC | libc::O_NOFOLLOW | libc::O_NONBLOCK);
    }
    #[cfg(windows)]
    {
        use std::os::windows::fs::OpenOptionsExt;

        // Open the final reparse point itself so descriptor metadata can
        // reject symlinks and junctions without a path-check/open race.
        const FILE_FLAG_OPEN_REPARSE_POINT: u32 = 0x0020_0000;
        const FILE_FLAG_BACKUP_SEMANTICS: u32 = 0x0200_0000;
        options.custom_flags(FILE_FLAG_OPEN_REPARSE_POINT | FILE_FLAG_BACKUP_SEMANTICS);
    }

    let mut file = options.open(path).map_err(|error| {
        #[cfg(unix)]
        if error.raw_os_error() == Some(libc::ELOOP) {
            return PackageError::InvalidFormat(
                "Pages package source must not be a symbolic link".to_owned(),
            );
        }
        PackageError::Io(error)
    })?;
    let metadata = file.metadata()?;
    #[cfg(windows)]
    {
        use std::os::windows::fs::MetadataExt;

        const FILE_ATTRIBUTE_REPARSE_POINT: u32 = 0x0000_0400;
        if metadata.file_attributes() & FILE_ATTRIBUTE_REPARSE_POINT != 0 {
            return Err(PackageError::InvalidFormat(
                "Pages package source must not be a symbolic link or junction".to_owned(),
            ));
        }
    }
    if !metadata.is_file() {
        return Err(PackageError::InvalidFormat(
            "Pages package source must be a regular file".to_owned(),
        ));
    }
    let before = FileSnapshot::from_metadata(&metadata);
    let source = read_source_with_reported_length(&mut file, before.length, limits)?;
    let after = FileSnapshot::from_metadata(&file.metadata()?);
    let observed_length = u64::try_from(source.len()).map_err(|_error| {
        PackageError::InvalidFormat("Pages package input length does not fit u64".to_owned())
    })?;
    ensure_source_unchanged(before, after, observed_length)?;
    Ok(source)
}

#[cfg(not(any(unix, windows)))]
fn read_path(_path: &Path, _limits: Limits) -> PackageResult<Arc<[u8]>> {
    Err(PackageError::InvalidFormat(
        "descriptor-first Pages package opening is unsupported on this platform".to_owned(),
    ))
}

fn ensure_source_unchanged(
    before: FileSnapshot,
    after: FileSnapshot,
    observed_length: u64,
) -> PackageResult<()> {
    if before != after || observed_length != before.length {
        return Err(PackageError::InvalidFormat(
            "Pages package source changed while it was being read".to_owned(),
        ));
    }
    Ok(())
}

fn read_source_with_reported_length(
    reader: &mut impl Read,
    reported_length: u64,
    limits: Limits,
) -> PackageResult<Arc<[u8]>> {
    if reported_length > limits.max_input_bytes() {
        return Err(PackageError::InvalidFormat(format!(
            "Pages package input exceeds the {} byte limit",
            limits.max_input_bytes()
        )));
    }

    let maximum = usize::try_from(limits.max_input_bytes()).map_err(|_error| {
        PackageError::InvalidFormat("Pages package input limit does not fit usize".to_owned())
    })?;
    let capacity = usize::try_from(reported_length)
        .map_err(|_error| {
            PackageError::InvalidFormat("Pages package input length does not fit usize".to_owned())
        })?
        .min(64 * 1024);
    let mut bytes = Vec::new();
    bytes.try_reserve_exact(capacity).map_err(|_error| {
        PackageError::InvalidFormat("could not allocate Pages input".to_owned())
    })?;

    let mut buffer = [0_u8; 8 * 1024];
    loop {
        let remaining = maximum.checked_sub(bytes.len()).ok_or_else(|| {
            PackageError::InvalidFormat("Pages package input length exceeds usize".to_owned())
        })?;
        if remaining == 0 {
            let mut extra = [0_u8; 1];
            if read_retrying_interrupted(reader, &mut extra)? != 0 {
                return Err(PackageError::InvalidFormat(format!(
                    "Pages package input exceeds the {} byte limit",
                    limits.max_input_bytes()
                )));
            }
            break;
        }

        let read_limit = remaining.min(buffer.len());
        let read = read_retrying_interrupted(reader, &mut buffer[..read_limit])?;
        if read == 0 {
            break;
        }
        let required = bytes.len().checked_add(read).ok_or_else(|| {
            PackageError::InvalidFormat("Pages package input length exceeds usize".to_owned())
        })?;
        reserve_source_growth(&mut bytes, required, maximum)?;
        bytes.extend_from_slice(&buffer[..read]);
    }
    Ok(bytes.into())
}

fn read_retrying_interrupted(reader: &mut impl Read, buffer: &mut [u8]) -> std::io::Result<usize> {
    loop {
        match reader.read(buffer) {
            Err(error) if error.kind() == std::io::ErrorKind::Interrupted => {},
            result => return result,
        }
    }
}

fn reserve_source_growth(
    bytes: &mut Vec<u8>,
    required: usize,
    maximum: usize,
) -> PackageResult<()> {
    if required <= bytes.capacity() {
        return Ok(());
    }

    let doubled = bytes.capacity().checked_mul(2).unwrap_or(maximum);
    let target = required.max(doubled).min(maximum);
    let additional = target.checked_sub(bytes.len()).ok_or_else(|| {
        PackageError::InvalidFormat("Pages package input length exceeds usize".to_owned())
    })?;
    bytes
        .try_reserve_exact(additional)
        .map_err(|_error| PackageError::InvalidFormat("could not allocate Pages input".to_owned()))
}

fn metadata_entry<'a>(catalog: &'a Catalog, name: &str) -> PackageResult<Option<&'a [u8]>> {
    let Some(entry) = catalog.iter().find(|entry| entry.name() == name) else {
        return Ok(None);
    };
    if entry.is_opaque() {
        return Err(PackageError::InvalidFormat(format!(
            "Pages metadata entry {name} uses an unsupported compression method"
        )));
    }
    Ok(Some(entry.data()))
}

fn parse_build_version(data: &[u8]) -> PackageResult<Option<String>> {
    let property_list = Value::from_reader(std::io::Cursor::new(data)).map_err(|error| {
        PackageError::InvalidFormat(format!(
            "failed to parse Pages BuildVersionHistory.plist: {error}"
        ))
    })?;
    let Value::Array(versions) = property_list else {
        return Err(PackageError::InvalidFormat(
            "Pages BuildVersionHistory.plist must contain an array at its root".to_owned(),
        ));
    };

    let mut latest = None;
    for (index, version) in versions.iter().enumerate() {
        let version_string = match version {
            Value::String(text) => text.clone(),
            Value::Dictionary(values) => values
                .get("Version")
                .or_else(|| values.get("Build"))
                .and_then(value_as_string)
                .ok_or_else(|| {
                    PackageError::InvalidFormat(format!(
                        "Pages BuildVersionHistory.plist[{index}] dictionary has no string Version or Build"
                    ))
                })?,
            Value::Array(_)
            | Value::Boolean(_)
            | Value::Data(_)
            | Value::Date(_)
            | Value::Real(_)
            | Value::Integer(_)
            | Value::Uid(_)
            | _ => {
                return Err(PackageError::InvalidFormat(format!(
                    "Pages BuildVersionHistory.plist[{index}] must be a string or dictionary"
                )));
            },
        };
        latest = Some(version_string);
    }
    Ok(latest)
}

fn property_string(properties: &plist::Dictionary, key: &str) -> Option<String> {
    properties.get(key).and_then(value_as_string)
}

fn value_as_string(property: &Value) -> Option<String> {
    match property {
        Value::String(text) => Some(text.clone()),
        Value::Integer(integer) => integer.as_signed().map(|number| number.to_string()),
        Value::Real(number) => Some(number.to_string()),
        Value::Boolean(boolean) => Some(boolean.to_string()),
        Value::Date(date) => Some(format!("{date:?}")),
        Value::Data(_) | Value::Array(_) | Value::Dictionary(_) | Value::Uid(_) | _ => None,
    }
}

fn validate_components(components: &ComponentCatalog) -> PackageResult<usize> {
    if components.is_empty() {
        return Err(PackageError::InvalidFormat(
            "Pages package contains no IWA components".to_owned(),
        ));
    }

    let object_count = checked_object_count(
        components
            .iter()
            .map(|component| component.archive().objects.len()),
    )?;
    if object_count == 0 {
        return Err(PackageError::InvalidFormat(
            "Pages package contains no IWA objects".to_owned(),
        ));
    }
    let mut object_ids = Vec::new();
    object_ids
        .try_reserve_exact(object_count)
        .map_err(|_error| PackageError::Allocation {
            amount: object_count,
        })?;
    for component in components.iter() {
        for object in &component.archive().objects {
            let identifier = object.archive_info.identifier.ok_or_else(|| {
                PackageError::InvalidFormat(format!(
                    "Pages component {} contains an object without an identifier",
                    component.name()
                ))
            })?;
            if identifier == 0 {
                return Err(PackageError::InvalidFormat(format!(
                    "Pages component {} contains object identifier zero",
                    component.name()
                )));
            }
            object_ids.push(identifier);
        }
    }
    object_ids.sort_unstable();
    if object_ids.windows(2).any(|pair| pair[0] == pair[1]) {
        return Err(PackageError::InvalidFormat(
            "Pages package contains a duplicate object identifier".to_owned(),
        ));
    }
    Ok(object_count)
}

fn checked_object_count(counts: impl IntoIterator<Item = usize>) -> PackageResult<usize> {
    let object_count = counts
        .into_iter()
        .try_fold(0usize, |count, component_count| {
            count
                .checked_add(component_count)
                .ok_or(PackageError::ObjectLimit {
                    observed: usize::MAX,
                    limit: MAX_OBJECTS,
                })
        })?;
    if object_count > MAX_OBJECTS {
        return Err(PackageError::ObjectLimit {
            observed: object_count,
            limit: MAX_OBJECTS,
        });
    }
    Ok(object_count)
}

fn root_references_with_limits(
    components: &ComponentCatalog,
    limits: Limits,
) -> PackageResult<RootReferences> {
    let component = components.get("Index/Document.iwa").ok_or_else(|| {
        PackageError::InvalidFormat("Pages package does not contain Index/Document.iwa".to_owned())
    })?;
    let object = component
        .archive()
        .object(1)
        .ok_or_else(|| PackageError::InvalidFormat("Pages root object 1 is missing".to_owned()))?;
    let payload = unique_message_payload(&object.messages, 10_000, "Pages root object 1")?;
    let root = pages_body_codec::decode_document_body(payload, pages_body_options(limits)?)
        .map_err(|error| {
            PackageError::InvalidFormat(format!(
                "Pages root type-10000 payload is invalid: {error}"
            ))
        })?;

    Ok(RootReferences {
        body: root
            .body_storage()
            .map(pages_body_codec::ReferenceSnapshot::identifier),
        initial_section: root
            .initial_section()
            .map(pages_body_codec::ReferenceSnapshot::identifier),
    })
}

fn pages_body_options(limits: Limits) -> PackageResult<PagesBodyDecodeOptions> {
    let archive_limits = limits.effective_archive_limits()?;
    let wire_limits = WireLimits::default();
    let recursion_limit = u32::try_from(
        archive_limits
            .max_header_nesting()
            .min(wire_limits.max_nesting()),
    )
    .map_err(|_error| {
        PackageError::InvalidFormat("Pages projection nesting limit does not fit u32".to_owned())
    })?;
    Ok(PagesBodyDecodeOptions::new(
        archive_limits
            .max_message_bytes()
            .min(wire_limits.max_input_bytes()),
        archive_limits
            .max_header_fields()
            .min(wire_limits.max_fields()),
        wire_limits.max_rewrite_work(),
        recursion_limit,
    ))
}

fn decode_document(
    components: &ComponentCatalog,
    root_references: RootReferences,
    max_sections: usize,
    max_text_bytes: usize,
    limits: Limits,
) -> PackageResult<Document> {
    if let Some(identifier) = root_references.body {
        let object = find_object(components, identifier.get()).ok_or_else(|| {
            PackageError::InvalidFormat(format!(
                "Pages body storage object {identifier} is missing"
            ))
        })?;
        let (storage, table_references) = decode_body_storage(
            &object.messages,
            identifier,
            max_sections,
            max_text_bytes,
            limits,
        )?;
        let section_references = native_section_references(
            table_references,
            root_references.initial_section,
            max_sections,
        )?;
        if section_references.is_empty() && max_sections == 0 {
            return Err(PackageError::Semantic(SemanticError::TooManySections {
                actual: 1,
                limit: max_sections,
            }));
        }
        return project_native_body(
            components,
            storage,
            section_references,
            max_text_bytes,
            identifier,
        );
    }

    if root_references.initial_section.is_some() {
        return Err(PackageError::InvalidFormat(
            "Pages root has an initial section but no body storage".to_owned(),
        ));
    }

    let body = {
        let storages = extract_storages(components, max_sections, max_text_bytes, limits)?;
        (!storages.is_empty())
            .then(|| {
                Body::with_max_text_bytes(storages, max_text_bytes).map_err(PackageError::Semantic)
            })
            .transpose()?
    };
    let root = body.map_or_else(Root::empty, Root::with_body);
    Document::from_root_with_max_text_bytes(root, max_text_bytes).map_err(PackageError::Semantic)
}

fn find_object(
    components: &ComponentCatalog,
    identifier: u64,
) -> Option<&litchi_iwa_core::ArchiveObject> {
    components
        .iter()
        .find_map(|component| component.archive().object(identifier))
}

fn decode_body_storage(
    messages: &[litchi_iwa_core::RawMessage],
    identifier: NonZeroU64,
    max_sections: usize,
    max_text_bytes: usize,
    limits: Limits,
) -> PackageResult<(Storage, Vec<NativeSectionReference>)> {
    decode_body_storage_with_wire_error(messages, identifier, max_sections, max_text_bytes, limits)
        .map_err(map_body_storage_decode_error)
}

fn decode_body_storage_with_wire_error(
    messages: &[litchi_iwa_core::RawMessage],
    identifier: NonZeroU64,
    max_sections: usize,
    max_text_bytes: usize,
    limits: Limits,
) -> Result<(Storage, Vec<NativeSectionReference>), BodyStorageDecodeError> {
    let payload = unique_text_payload(messages, identifier)?;
    let wire_limits = storage_rewrite_limits(limits).map_err(|limit_error| match limit_error {
        StorageWireLimitsError::Physical(physical_error) => PackageError::Archive(physical_error),
        StorageWireLimitsError::Wire(wire_error) => PackageError::InvalidFormat(format!(
            "Pages body object {identifier} text validation limits are invalid: {wire_error}"
        )),
    })?;
    let decoded = litchi_iwa_text_wire::decode_storage_with_limits(payload, wire_limits)
        .map_err(BodyStorageDecodeError::Wire)?;
    // The strict text-wire pass above is authoritative for legacy balanced
    // unknown groups. `WireView` intentionally rejects group wire values, so
    // omit only those already-validated opaque root fields from its private
    // section-boundary projection. The original payload remains the source
    // for semantic decoding and every rewrite copies the group bytes exactly.
    let projection = body_wire_projection_without_unknown_groups(
        payload,
        identifier,
        wire_limits.max_nesting(),
    )?;
    let preflight = preflight_body_wire(
        projection.as_ref(),
        identifier,
        max_sections,
        max_text_bytes,
        limits,
    )?;
    let validation = decoded.validation();
    let storage = decoded.into_storage();
    let materialized_utf16 = storage.text().encode_utf16().count();
    if validation.fragments() != preflight.fragment_count
        || validation.utf8_len() != preflight.text_bytes
        || storage.runs().len() != preflight.fragment_count
        || storage.len() != preflight.text_bytes
        || materialized_utf16 != validation.utf16_len()
    {
        return Err(BodyStorageDecodeError::Package(
            PackageError::InvalidFormat(format!(
                "Pages body object {identifier} lazy text projection disagreed with strict preflight"
            )),
        ));
    }
    Ok((storage, preflight.section_references))
}

fn body_wire_projection_without_unknown_groups<'source>(
    source: &'source [u8],
    body_identifier: NonZeroU64,
    max_nesting: usize,
) -> PackageResult<Cow<'source, [u8]>> {
    let context = format!("Pages body object {body_identifier}");
    let mut offset = 0usize;
    let mut projection: Option<Vec<u8>> = None;
    while offset < source.len() {
        let start = offset;
        let (number, wire_type, key_end) = body_wire_key(source, offset, &context)?;
        offset = if wire_type == 3 {
            if is_known_body_root_field(number) {
                return Err(PackageError::InvalidFormat(format!(
                    "{context} known protobuf field {number} cannot use group wire type"
                )));
            }
            let end = skip_body_wire_group(source, key_end, number, 1, max_nesting, &context)?;
            if projection.is_none() {
                let mut filtered = Vec::new();
                filtered
                    .try_reserve_exact(source.len())
                    .map_err(|_allocation| PackageError::Allocation {
                        amount: source.len(),
                    })?;
                filtered.extend_from_slice(&source[..start]);
                projection = Some(filtered);
            }
            end
        } else if wire_type == 4 {
            return Err(PackageError::InvalidFormat(format!(
                "{context} contains an unexpected protobuf end-group field {number}"
            )));
        } else {
            let end = body_wire_value_end(source, key_end, wire_type, &context)?;
            if let Some(filtered) = projection.as_mut() {
                filtered.extend_from_slice(&source[start..end]);
            }
            end
        };
    }
    Ok(projection.map_or(Cow::Borrowed(source), Cow::Owned))
}

fn skip_body_wire_group(
    source: &[u8],
    mut offset: usize,
    expected_number: u32,
    depth: usize,
    max_nesting: usize,
    context: &str,
) -> PackageResult<usize> {
    if depth > max_nesting {
        return Err(PackageError::InvalidFormat(format!(
            "{context} protobuf group nesting exceeds {max_nesting}"
        )));
    }
    loop {
        if offset >= source.len() {
            return Err(PackageError::InvalidFormat(format!(
                "{context} contains an unterminated protobuf group {expected_number}"
            )));
        }
        let (number, wire_type, key_end) = body_wire_key(source, offset, context)?;
        offset = match wire_type {
            3 => {
                if is_known_body_root_field(number) {
                    return Err(PackageError::InvalidFormat(format!(
                        "{context} known protobuf field {number} cannot use group wire type"
                    )));
                }
                skip_body_wire_group(
                    source,
                    key_end,
                    number,
                    depth.saturating_add(1),
                    max_nesting,
                    context,
                )?
            },
            4 => {
                if number != expected_number {
                    return Err(PackageError::InvalidFormat(format!(
                        "{context} protobuf end-group field {number} does not match {expected_number}"
                    )));
                }
                return Ok(key_end);
            },
            _ => body_wire_value_end(source, key_end, wire_type, context)?,
        };
    }
}

fn body_wire_key(source: &[u8], offset: usize, context: &str) -> PackageResult<(u32, u8, usize)> {
    let (key, key_length) = decode_varint_from_bytes(source.get(offset..).ok_or_else(|| {
        PackageError::InvalidFormat(format!("{context} protobuf key offset is invalid"))
    })?)
    .map_err(|error| {
        PackageError::InvalidFormat(format!("{context} has an invalid protobuf key: {error}"))
    })?;
    let key_end = offset.checked_add(key_length).ok_or_else(|| {
        PackageError::InvalidFormat(format!("{context} protobuf key offset overflows"))
    })?;
    let number = u32::try_from(key >> 3).map_err(|_error| {
        PackageError::InvalidFormat(format!("{context} protobuf field number exceeds u32"))
    })?;
    if number == 0 || number > 0x1fff_ffff {
        return Err(PackageError::InvalidFormat(format!(
            "{context} protobuf field number {number} is invalid"
        )));
    }
    let wire_type = u8::try_from(key & 7).map_err(|_error| {
        PackageError::InvalidFormat(format!("{context} protobuf wire type exceeds u8"))
    })?;
    Ok((number, wire_type, key_end))
}

fn body_wire_value_end(
    source: &[u8],
    key_end: usize,
    wire_type: u8,
    context: &str,
) -> PackageResult<usize> {
    let end = match wire_type {
        0 => {
            let (_, width) = decode_varint_from_bytes(source.get(key_end..).ok_or_else(|| {
                PackageError::InvalidFormat(format!("{context} protobuf value offset is invalid"))
            })?)
            .map_err(|error| {
                PackageError::InvalidFormat(format!(
                    "{context} contains an invalid protobuf varint: {error}"
                ))
            })?;
            key_end.checked_add(width)
        },
        1 => key_end.checked_add(8),
        2 => {
            let (encoded_length, prefix_width) =
                decode_varint_from_bytes(source.get(key_end..).ok_or_else(|| {
                    PackageError::InvalidFormat(format!(
                        "{context} protobuf length offset is invalid"
                    ))
                })?)
                .map_err(|error| {
                    PackageError::InvalidFormat(format!(
                        "{context} contains an invalid protobuf length: {error}"
                    ))
                })?;
            let payload_start = key_end.checked_add(prefix_width).ok_or_else(|| {
                PackageError::InvalidFormat(format!("{context} protobuf length prefix overflows"))
            })?;
            let length = usize::try_from(encoded_length).map_err(|_error| {
                PackageError::InvalidFormat(format!(
                    "{context} protobuf field length exceeds usize"
                ))
            })?;
            payload_start.checked_add(length)
        },
        5 => key_end.checked_add(4),
        _ => {
            return Err(PackageError::InvalidFormat(format!(
                "{context} contains invalid protobuf wire type {wire_type}"
            )));
        },
    }
    .ok_or_else(|| {
        PackageError::InvalidFormat(format!("{context} protobuf field range overflows"))
    })?;
    if end > source.len() {
        return Err(PackageError::InvalidFormat(format!(
            "{context} contains a truncated protobuf field"
        )));
    }
    Ok(end)
}

fn is_known_body_root_field(number: u32) -> bool {
    matches!(number, 1..=12 | 14..=28)
}

fn map_body_storage_decode_error(error: BodyStorageDecodeError) -> PackageError {
    match error {
        BodyStorageDecodeError::Package(error) => error,
        BodyStorageDecodeError::Wire(error) => map_rooted_storage_decode_error(error),
        BodyStorageDecodeError::SemanticLimit { observed, limit } => {
            PackageError::PayloadLimit { observed, limit }
        },
    }
}

fn project_body_footnotes(
    components: &ComponentCatalog,
    limits: Limits,
    max_text_bytes: usize,
) -> PackageResult<Vec<Footnote>> {
    let mut semantic_budget = FootnoteSemanticBudget::new(max_text_bytes);
    project_body_footnotes_with_budget(components, limits, max_text_bytes, &mut semantic_budget)
        .map_err(map_body_storage_decode_error)
}

fn project_body_footnotes_with_budget(
    components: &ComponentCatalog,
    limits: Limits,
    max_text_bytes: usize,
    semantic_budget: &mut FootnoteSemanticBudget,
) -> Result<Vec<Footnote>, BodyStorageDecodeError> {
    let root_references = root_references_with_limits(components, limits)?;
    let Some(body_identifier) = root_references.body else {
        return Ok(Vec::new());
    };
    let body_object = find_object(components, body_identifier.get()).ok_or_else(|| {
        PackageError::InvalidFormat(format!(
            "Pages body storage object {body_identifier} is missing"
        ))
    })?;
    let body_payload = unique_text_payload(&body_object.messages, body_identifier)?;
    let (body_storage, _) = decode_body_storage_with_wire_error(
        &body_object.messages,
        body_identifier,
        MAX_SECTIONS,
        max_text_bytes,
        limits,
    )?;
    let entries = footnote_table_entries(body_payload, body_identifier, limits)?;
    if entries.len() > MAX_BODY_FOOTNOTES {
        return Err(PackageError::PayloadLimit {
            observed: entries.len(),
            limit: MAX_BODY_FOOTNOTES,
        }
        .into());
    }

    let mut footnotes = Vec::new();
    footnotes
        .try_reserve_exact(entries.len())
        .map_err(|_error| PackageError::Allocation {
            amount: entries.len(),
        })?;
    let mut seen_references = Vec::new();
    seen_references
        .try_reserve_exact(entries.len())
        .map_err(|_error| PackageError::Allocation {
            amount: entries.len(),
        })?;
    let mut seen_storages = Vec::new();
    seen_storages
        .try_reserve_exact(entries.len())
        .map_err(|_error| PackageError::Allocation {
            amount: entries.len(),
        })?;
    let mut seen_markers = Vec::new();
    seen_markers
        .try_reserve_exact(entries.len())
        .map_err(|_error| PackageError::Allocation {
            amount: entries.len(),
        })?;
    let mut previous_position = None;
    for entry in entries {
        if previous_position.is_some_and(|previous| previous >= entry.character_index) {
            return Err(PackageError::InvalidFormat(
                "Pages body footnote positions are not strictly increasing".to_owned(),
            )
            .into());
        }
        previous_position = Some(entry.character_index);
        if seen_references.contains(&entry.identifier) {
            return Err(PackageError::InvalidFormat(format!(
                "Pages body references footnote object {} more than once",
                entry.identifier
            ))
            .into());
        }
        seen_references.push(entry.identifier);
        validate_body_footnote_anchor(body_storage.text(), body_identifier, entry.character_index)?;
        let (footnote, storage_identifier, marker_identifier) =
            project_one_body_footnote(components, limits, max_text_bytes, entry, semantic_budget)?;
        if entry.identifier == storage_identifier
            || entry.identifier == marker_identifier
            || seen_storages.contains(&storage_identifier)
            || seen_markers.contains(&marker_identifier)
        {
            return Err(PackageError::InvalidFormat(
                "Pages body footnote graph reuses a native object".to_owned(),
            )
            .into());
        }
        seen_storages.push(storage_identifier);
        seen_markers.push(marker_identifier);
        footnotes.push(footnote);
    }
    Ok(footnotes)
}

fn footnote_table_entries(
    body_payload: &[u8],
    body_identifier: NonZeroU64,
    limits: Limits,
) -> PackageResult<Vec<NativeFootnoteReference>> {
    let wire_limits = storage_wire_limits(limits)?;
    let context = format!("Pages body object {body_identifier}");
    let body_view = WireView::parse_with_limits(body_payload, wire_limits).map_err(|error| {
        PackageError::InvalidFormat(format!("{context} has invalid protobuf wire data: {error}"))
    })?;
    let Some(table_field) =
        unique_wire_field(&body_view, FOOTNOTE_TABLE_FIELD, 2, false, &context)?
    else {
        return Ok(Vec::new());
    };
    let table_view =
        WireView::parse_with_limits(table_field.payload(), wire_limits).map_err(|error| {
            PackageError::InvalidFormat(format!(
                "{context} footnote table has invalid protobuf wire data: {error}"
            ))
        })?;
    let boundary_options = pages_body_options(limits)?;
    let mut entries = Vec::new();
    for field in table_view
        .fields()
        .filter(|field| field.number() == TABLE_ENTRIES_FIELD)
    {
        validate_wire_field(field, 2, &context)?;
        let next_count = entries.len().saturating_add(1);
        if next_count > MAX_BODY_FOOTNOTES {
            return Err(PackageError::PayloadLimit {
                observed: next_count,
                limit: MAX_BODY_FOOTNOTES,
            });
        }
        let boundary = pages_body_codec::decode_section_boundary(field.payload(), boundary_options)
            .map_err(|error| {
                PackageError::InvalidFormat(format!(
                    "{context} footnote table entry is invalid: {error}"
                ))
            })?;
        let identifier = boundary.section().ok_or_else(|| {
            PackageError::InvalidFormat(format!(
                "{context} footnote table entry has no footnote reference"
            ))
        })?;
        entries
            .try_reserve(1)
            .map_err(|_error| PackageError::Allocation { amount: next_count })?;
        entries.push(NativeFootnoteReference {
            character_index: boundary.character_index(),
            identifier: identifier.identifier(),
        });
    }
    Ok(entries)
}

fn project_one_body_footnote(
    components: &ComponentCatalog,
    limits: Limits,
    max_text_bytes: usize,
    entry: NativeFootnoteReference,
    semantic_budget: &mut FootnoteSemanticBudget,
) -> Result<(Footnote, NonZeroU64, NonZeroU64), BodyStorageDecodeError> {
    let reference_object = find_object(components, entry.identifier.get()).ok_or_else(|| {
        PackageError::InvalidFormat(format!(
            "Pages footnote reference object {} is missing",
            entry.identifier
        ))
    })?;
    let reference_payload = unique_message_payload(
        &reference_object.messages,
        FOOTNOTE_REFERENCE_MESSAGE_TYPE,
        &format!("Pages footnote reference object {}", entry.identifier),
    )?;
    let reference = pages_footnote_codec::decode_footnote_reference(
        reference_payload,
        footnote_decode_options(reference_payload, limits)?,
    )
    .map_err(|error| {
        PackageError::InvalidFormat(format!(
            "Pages footnote reference object {} failed strict validation: {error}",
            entry.identifier
        ))
    })?;
    if reference
        .super_kind()
        .is_some_and(|kind| kind != FOOTNOTE_MARK_KIND)
    {
        return Err(PackageError::InvalidFormat(format!(
            "Pages footnote object {} has the wrong attachment kind",
            entry.identifier
        ))
        .into());
    }
    let storage_identifier = reference
        .contained_storage()
        .map(|reference| reference.identifier())
        .ok_or_else(|| {
            PackageError::InvalidFormat(format!(
                "Pages footnote object {} has no contained storage",
                entry.identifier
            ))
        })?;
    let storage_object = find_object(components, storage_identifier.get()).ok_or_else(|| {
        PackageError::InvalidFormat(format!(
            "Pages footnote storage object {} is missing",
            storage_identifier
        ))
    })?;
    let (storage, _) = decode_body_storage_with_wire_error(
        &storage_object.messages,
        storage_identifier,
        MAX_SECTIONS,
        max_text_bytes,
        limits,
    )?;
    let storage_payload = unique_text_payload(&storage_object.messages, storage_identifier)?;
    let marker_identifier =
        footnote_marker_identifier(storage_payload, storage_identifier, limits)?;
    let marker_object = find_object(components, marker_identifier.get()).ok_or_else(|| {
        PackageError::InvalidFormat(format!(
            "Pages footnote marker object {} is missing",
            marker_identifier
        ))
    })?;
    let marker_payload = unique_message_payload(
        &marker_object.messages,
        TEXTUAL_ATTACHMENT_MESSAGE_TYPE,
        &format!("Pages footnote marker object {}", marker_identifier),
    )?;
    let marker = pages_footnote_marker_codec::decode_textual_attachment(
        marker_payload,
        footnote_marker_decode_options(marker_payload, limits)?,
    )
    .map_err(|error| {
        PackageError::InvalidFormat(format!(
            "Pages footnote marker object {} failed strict validation: {error}",
            marker_identifier
        ))
    })?;
    if marker.kind() != Some(FOOTNOTE_MARK_KIND) {
        return Err(PackageError::InvalidFormat(format!(
            "Pages footnote marker object {} has the wrong attachment kind",
            marker_identifier
        ))
        .into());
    }

    let text = storage
        .text()
        .strip_prefix(STORAGE_TEXT_PREFIX)
        .ok_or_else(|| {
            PackageError::InvalidFormat(format!(
                "Pages footnote storage {} lacks its native marker prefix",
                storage_identifier
            ))
        })?;
    if text.len() > crate::footnote::body::MAX_TEXT_BYTES {
        return Err(PackageError::InvalidFormat(format!(
            "Pages footnote storage {storage_identifier} exceeds its semantic text budget"
        ))
        .into());
    }
    let custom_mark = reference
        .custom_mark_string()
        .map(|value| {
            if value.len() > crate::footnote::body::MAX_CUSTOM_MARK_BYTES {
                return Err(PackageError::InvalidFormat(format!(
                    "Pages footnote object {} exceeds its semantic custom-marker budget",
                    entry.identifier
                )));
            }
            Ok(value)
        })
        .transpose()?;
    semantic_budget
        .charge(text.len(), custom_mark.map_or(0, str::len))
        .map_err(|error| match error {
            PackageError::PayloadLimit { observed, limit } => {
                BodyStorageDecodeError::SemanticLimit { observed, limit }
            },
            error => BodyStorageDecodeError::Package(error),
        })?;
    let text = try_owned_footnote_string(text)?;
    let custom_mark = custom_mark.map(try_owned_footnote_string).transpose()?;
    let footnote = Footnote::with_custom_mark(
        Position::from_utf16_index(entry.character_index as usize).map_err(|error| {
            PackageError::InvalidFormat(format!(
                "Pages footnote position {} is invalid: {error}",
                entry.character_index
            ))
        })?,
        text,
        custom_mark,
    )
    .map_err(|error| {
        PackageError::InvalidFormat(format!(
            "Pages footnote object {} has an invalid semantic value: {error}",
            entry.identifier
        ))
    })?;
    Ok((footnote, storage_identifier, marker_identifier))
}

fn try_owned_footnote_string(value: &str) -> PackageResult<Box<str>> {
    let mut owned = String::new();
    owned
        .try_reserve_exact(value.len())
        .map_err(|_error| PackageError::Allocation {
            amount: value.len(),
        })?;
    owned.push_str(value);
    Ok(owned.into_boxed_str())
}

fn validate_body_footnote_anchor(
    text: &str,
    body_identifier: NonZeroU64,
    requested: u32,
) -> PackageResult<()> {
    let requested = usize::try_from(requested).map_err(|_error| {
        PackageError::InvalidFormat(format!(
            "Pages body object {body_identifier} footnote position exceeds the platform index range"
        ))
    })?;
    let mut offset = 0usize;
    for character in text.chars() {
        if offset == requested {
            if character as u32 == u32::from(FOOTNOTE_ANCHOR_UNIT) {
                return Ok(());
            }
            return Err(PackageError::InvalidFormat(format!(
                "Pages body object {body_identifier} has no U+000E footnote anchor at UTF-16 index {requested}"
            )));
        }
        offset = offset.checked_add(character.len_utf16()).ok_or_else(|| {
            PackageError::InvalidFormat(format!(
                "Pages body object {body_identifier} UTF-16 length overflows usize"
            ))
        })?;
        if offset > requested {
            return Err(PackageError::InvalidFormat(format!(
                "Pages body object {body_identifier} footnote position {requested} splits a UTF-16 surrogate pair"
            )));
        }
    }
    Err(PackageError::InvalidFormat(format!(
        "Pages body object {body_identifier} footnote position {requested} exceeds body UTF-16 length {offset}"
    )))
}

fn footnote_marker_identifier(
    storage_payload: &[u8],
    storage_identifier: NonZeroU64,
    limits: Limits,
) -> PackageResult<NonZeroU64> {
    let wire_limits = storage_wire_limits(limits)?;
    let context = format!("Pages footnote storage object {storage_identifier}");
    let view = WireView::parse_with_limits(storage_payload, wire_limits).map_err(|error| {
        PackageError::InvalidFormat(format!("{context} has invalid protobuf wire data: {error}"))
    })?;
    let kind = unique_wire_field(&view, STORAGE_KIND_FIELD, 0, true, &context)?
        .ok_or_else(|| PackageError::InvalidFormat(format!("{context} has no native kind")))?;
    let kind = decode_canonical_varint(kind, &context)?;
    if kind != FOOTNOTE_STORAGE_KIND {
        return Err(PackageError::InvalidFormat(format!(
            "{context} is not a native footnote storage"
        )));
    }
    let table =
        unique_wire_field(&view, TABLE_ATTACHMENT_FIELD, 2, true, &context)?.ok_or_else(|| {
            PackageError::InvalidFormat(format!("{context} has no marker attachment table"))
        })?;
    let table_view =
        WireView::parse_with_limits(table.payload(), wire_limits).map_err(|error| {
            PackageError::InvalidFormat(format!(
                "{context} marker attachment table has invalid protobuf wire data: {error}"
            ))
        })?;
    let boundary_options = pages_body_options(limits)?;
    let mut marker = None;
    for field in table_view
        .fields()
        .filter(|field| field.number() == TABLE_ENTRIES_FIELD)
    {
        validate_wire_field(field, 2, &context)?;
        let entry = pages_body_codec::decode_section_boundary(field.payload(), boundary_options)
            .map_err(|error| {
                PackageError::InvalidFormat(format!(
                    "{context} marker attachment entry is invalid: {error}"
                ))
            })?;
        if entry.character_index() != 0 {
            continue;
        }
        let identifier = entry.section().ok_or_else(|| {
            PackageError::InvalidFormat(format!(
                "{context} marker attachment at index zero has no object reference"
            ))
        })?;
        if marker.replace(identifier.identifier()).is_some() {
            return Err(PackageError::InvalidFormat(format!(
                "{context} has more than one marker attachment at index zero"
            )));
        }
    }
    marker.ok_or_else(|| {
        PackageError::InvalidFormat(format!("{context} has no marker attachment at index zero"))
    })
}

fn decode_canonical_varint(field: WireFieldView<'_>, context: &str) -> PackageResult<u64> {
    let payload = field.payload();
    let (value, length) = decode_varint_from_bytes(payload).map_err(|error| {
        PackageError::InvalidFormat(format!("{context} protobuf varint is invalid: {error}"))
    })?;
    if length != payload.len() || litchi_iwa_common::varint::encoded_len(value) != length {
        return Err(PackageError::InvalidFormat(format!(
            "{context} protobuf varint is not canonical"
        )));
    }
    Ok(value)
}

fn footnote_decode_options(
    _source: &[u8],
    limits: Limits,
) -> PackageResult<pages_footnote_codec::DecodeOptions> {
    let rewrite_limits = storage_rewrite_limits(limits).map_err(map_storage_wire_limits_error)?;
    let recursion_limit = u32::try_from(rewrite_limits.max_nesting()).map_err(|_error| {
        PackageError::InvalidFormat("Pages footnote nesting limit does not fit u32".to_owned())
    })?;
    Ok(pages_footnote_codec::DecodeOptions::new(
        rewrite_limits.max_message_bytes(),
        rewrite_limits.max_fields(),
        rewrite_limits.max_rewrite_work(),
        recursion_limit,
    ))
}

fn footnote_marker_decode_options(
    _source: &[u8],
    limits: Limits,
) -> PackageResult<pages_footnote_marker_codec::DecodeOptions> {
    let rewrite_limits = storage_rewrite_limits(limits).map_err(map_storage_wire_limits_error)?;
    let recursion_limit = u32::try_from(rewrite_limits.max_nesting()).map_err(|_error| {
        PackageError::InvalidFormat("Pages footnote nesting limit does not fit u32".to_owned())
    })?;
    Ok(pages_footnote_marker_codec::DecodeOptions::new(
        rewrite_limits.max_message_bytes(),
        rewrite_limits.max_fields(),
        rewrite_limits.max_rewrite_work(),
        recursion_limit,
    ))
}

fn storage_wire_limits(limits: Limits) -> PackageResult<WireLimits> {
    let rewrite_limits = storage_rewrite_limits(limits).map_err(map_storage_wire_limits_error)?;
    WireLimits::default()
        .with_input_bytes(rewrite_limits.max_message_bytes())
        .and_then(|limits| limits.with_fields(rewrite_limits.max_fields()))
        .and_then(|limits| limits.with_nesting(rewrite_limits.max_nesting()))
        .and_then(|limits| limits.with_rewrite_work(rewrite_limits.max_rewrite_work()))
        .map_err(|error| {
            PackageError::InvalidFormat(format!("Pages footnote wire limits are invalid: {error}"))
        })
}

fn map_storage_wire_limits_error(error: StorageWireLimitsError) -> PackageError {
    match error {
        StorageWireLimitsError::Physical(error) => PackageError::Archive(error),
        StorageWireLimitsError::Wire(error) => {
            PackageError::InvalidFormat(format!("Pages footnote wire limits are invalid: {error}"))
        },
    }
}

fn map_rooted_storage_decode_error(error: litchi_iwa_text_wire::RewriteError) -> PackageError {
    match error {
        litchi_iwa_text_wire::RewriteError::Allocation { amount, .. } => {
            PackageError::Allocation { amount }
        },
        litchi_iwa_text_wire::RewriteError::LimitExceeded {
            observed, limit, ..
        } => PackageError::PayloadLimit { observed, limit },
        litchi_iwa_text_wire::RewriteError::InvalidFormat(reason) => PackageError::InvalidFormat(
            format!("Pages body text payload failed bounded validation: {reason}"),
        ),
        litchi_iwa_text_wire::RewriteError::Projection(_) => PackageError::InvalidFormat(
            "Pages body lazy text projection disagreed with validated input".to_owned(),
        ),
        _ => {
            PackageError::InvalidFormat("Pages body text validation could not complete".to_owned())
        },
    }
}

fn preflight_body_wire(
    payload: &[u8],
    body_identifier: NonZeroU64,
    max_sections: usize,
    max_text_bytes: usize,
    limits: Limits,
) -> PackageResult<BodyPreflight> {
    let context = format!("Pages body object {body_identifier}");
    let view = parse_wire(payload, &context)?;
    let mut fragment_count = 0usize;
    let mut text_bytes = 0usize;
    for field in view.fields() {
        match field.number() {
            3 => {
                validate_wire_field(field, 2, &context)?;
                fragment_count = fragment_count.checked_add(1).ok_or_else(|| {
                    PackageError::InvalidFormat(format!(
                        "{context} text fragment count overflows usize"
                    ))
                })?;
                if fragment_count > litchi_iwa_text_wire::MAX_FRAGMENTS {
                    return Err(PackageError::InvalidFormat(format!(
                        "{context} contains {fragment_count} text fragments; maximum is {}",
                        litchi_iwa_text_wire::MAX_FRAGMENTS
                    )));
                }
                std::str::from_utf8(field.payload()).map_err(|error| {
                    PackageError::InvalidFormat(format!(
                        "{context} text fragment {} is not valid UTF-8: {error}",
                        fragment_count - 1
                    ))
                })?;
                text_bytes = text_bytes
                    .checked_add(field.payload().len())
                    .ok_or_else(|| {
                        PackageError::InvalidFormat(format!(
                            "{context} text length overflows usize"
                        ))
                    })?;
                if text_bytes > max_text_bytes {
                    return Err(PackageError::Semantic(SemanticError::TextTooLarge {
                        observed: text_bytes,
                        limit: max_text_bytes,
                    }));
                }
            },
            17 => validate_wire_field(field, 2, &context)?,
            _ => {},
        }
    }

    let optional_table_field = unique_wire_field(&view, 17, 2, false, &context)?;
    let Some(wire_table_field) = optional_table_field else {
        return Ok(BodyPreflight {
            section_references: Vec::new(),
            fragment_count,
            text_bytes,
        });
    };
    let table_view = parse_wire(wire_table_field.payload(), &context)?;
    let mut entry_count = 0usize;
    let mut section_references = Vec::new();
    let boundary_options = pages_body_options(limits)?;
    for field in table_view.fields().filter(|field| field.number() == 1) {
        validate_wire_field(field, 2, &context)?;
        entry_count = entry_count.checked_add(1).ok_or_else(|| {
            PackageError::InvalidFormat(format!("{context} section count overflows usize"))
        })?;
        if entry_count > max_sections {
            return Err(PackageError::Semantic(SemanticError::TooManySections {
                actual: entry_count,
                limit: max_sections,
            }));
        }
        let reference = preflight_section_table_entry(
            field.payload(),
            entry_count - 1,
            body_identifier,
            boundary_options,
        )?;
        section_references
            .try_reserve(1)
            .map_err(|_allocation| PackageError::Allocation {
                amount: entry_count,
            })?;
        section_references.push(reference);
    }
    Ok(BodyPreflight {
        section_references,
        fragment_count,
        text_bytes,
    })
}

fn preflight_section_table_entry(
    payload: &[u8],
    entry_index: usize,
    body_identifier: NonZeroU64,
    options: PagesBodyDecodeOptions,
) -> PackageResult<NativeSectionReference> {
    let context = format!("Pages body object {body_identifier} section table entry {entry_index}");
    let boundary = pages_body_codec::decode_section_boundary(payload, options)
        .map_err(|error| PackageError::InvalidFormat(format!("{context} is invalid: {error}")))?;
    let section = boundary.section().ok_or_else(|| {
        PackageError::InvalidFormat(format!("{context} has no section reference"))
    })?;
    Ok(NativeSectionReference {
        character_index: boundary.character_index(),
        identifier: section.identifier(),
    })
}

fn parse_wire<'a>(payload: &'a [u8], context: &str) -> PackageResult<WireView<'a>> {
    WireView::parse(payload).map_err(|error| {
        PackageError::InvalidFormat(format!("{context} has invalid protobuf wire data: {error}"))
    })
}

fn unique_wire_field<'a>(
    view: &WireView<'a>,
    field_number: u32,
    wire_type: u8,
    required: bool,
    context: &str,
) -> PackageResult<Option<WireFieldView<'a>>> {
    let mut matching = None;
    for field in view.fields().filter(|field| field.number() == field_number) {
        validate_wire_field(field, wire_type, context)?;
        if matching.replace(field).is_some() {
            return Err(PackageError::InvalidFormat(format!(
                "{context} contains duplicate protobuf field {field_number}"
            )));
        }
    }
    if required && matching.is_none() {
        return Err(PackageError::InvalidFormat(format!(
            "{context} has no protobuf field {field_number}"
        )));
    }
    Ok(matching)
}

fn validate_wire_field(
    field: WireFieldView<'_>,
    wire_type: u8,
    context: &str,
) -> PackageResult<()> {
    if field.wire_type() != wire_type {
        return Err(PackageError::InvalidFormat(format!(
            "{context} protobuf field {} has wire type {} instead of {wire_type}",
            field.number(),
            field.wire_type()
        )));
    }
    field.validate_canonical_framing().map_err(|error| {
        PackageError::InvalidFormat(format!(
            "{context} protobuf field {} has invalid framing: {error}",
            field.number()
        ))
    })
}

fn native_section_references(
    mut references: Vec<NativeSectionReference>,
    initial_section: Option<NonZeroU64>,
    max_sections: usize,
) -> PackageResult<Vec<NativeSectionReference>> {
    let maximum_count = references
        .len()
        .checked_add(usize::from(initial_section.is_some()))
        .ok_or_else(|| {
            PackageError::InvalidFormat("Pages section count overflows usize".to_owned())
        })?;
    if maximum_count > max_sections.saturating_add(1) {
        return Err(PackageError::Semantic(SemanticError::TooManySections {
            actual: maximum_count,
            limit: max_sections,
        }));
    }

    references.sort_unstable_by_key(|reference| reference.character_index);
    if let Some(duplicates) = references
        .windows(2)
        .find(|pair| pair[0].character_index == pair[1].character_index)
    {
        return Err(PackageError::InvalidFormat(format!(
            "Pages has multiple section boundaries at UTF-16 index {}",
            duplicates[0].character_index
        )));
    }

    if let Some(identifier) = initial_section {
        if let Some(existing) = references
            .iter()
            .find(|reference| reference.character_index == 0)
        {
            if existing.identifier != identifier {
                return Err(PackageError::InvalidFormat(format!(
                    "Pages root section {identifier} conflicts with section {} at UTF-16 index zero",
                    existing.identifier
                )));
            }
        } else {
            references
                .try_reserve(1)
                .map_err(|_allocation| PackageError::Allocation {
                    amount: maximum_count,
                })?;
            references.push(NativeSectionReference {
                character_index: 0,
                identifier,
            });
            references.sort_unstable_by_key(|reference| reference.character_index);
        }
    }

    if references
        .first()
        .is_some_and(|reference| reference.character_index != 0)
    {
        return Err(PackageError::InvalidFormat(format!(
            "Pages initial section boundary starts at UTF-16 index {} instead of zero",
            references[0].character_index
        )));
    }
    if references.len() > max_sections {
        return Err(PackageError::Semantic(SemanticError::TooManySections {
            actual: references.len(),
            limit: max_sections,
        }));
    }

    let mut identifiers = Vec::new();
    identifiers
        .try_reserve_exact(references.len())
        .map_err(|_error| {
            PackageError::InvalidFormat("could not allocate Pages section identities".to_owned())
        })?;
    identifiers.extend(references.iter().map(|reference| reference.identifier));
    identifiers.sort_unstable();
    if let Some(identifier) = identifiers
        .windows(2)
        .find(|pair| pair[0] == pair[1])
        .map(|pair| pair[0])
    {
        return Err(PackageError::InvalidFormat(format!(
            "Pages section object {identifier} is attached at multiple boundaries"
        )));
    }
    Ok(references)
}

fn project_native_body(
    components: &ComponentCatalog,
    storage: Storage,
    section_references: Vec<NativeSectionReference>,
    max_text_bytes: usize,
    body_identifier: NonZeroU64,
) -> PackageResult<Document> {
    if section_references.is_empty() {
        let body = Body::with_max_text_bytes(vec![storage], max_text_bytes)
            .map_err(PackageError::Semantic)?;
        return Document::from_root_with_max_text_bytes(Root::with_body(body), max_text_bytes)
            .map_err(PackageError::Semantic);
    }

    let ranges = section_text_ranges(storage.text(), &section_references, body_identifier)?;
    let storages = split_native_text(&storage, &ranges, body_identifier)?;
    let names = decode_section_names(components, &section_references, max_text_bytes)?;
    let mut sections = Vec::new();
    sections
        .try_reserve_exact(section_references.len())
        .map_err(|_error| {
            PackageError::InvalidFormat("could not allocate Pages semantic sections".to_owned())
        })?;

    for (((index, reference), storage), name) in section_references
        .into_iter()
        .enumerate()
        .zip(storages)
        .zip(names)
    {
        let mut builder = Section::builder(index, SectionType::Body);
        builder.set_name(name).map_err(|error| {
            PackageError::InvalidFormat(format!(
                "Pages section object {} has an invalid name: {error}",
                reference.identifier
            ))
        })?;
        builder.push_text_storage(storage);
        sections.push(builder.build());
    }

    Document::from_sections_with_max_text_bytes(sections, max_text_bytes)
        .map_err(PackageError::Semantic)
}

fn decode_section_names(
    components: &ComponentCatalog,
    references: &[NativeSectionReference],
    max_name_text_bytes: usize,
) -> PackageResult<Vec<Option<Box<str>>>> {
    let mut requested = Vec::new();
    requested
        .try_reserve_exact(references.len())
        .map_err(|_error| {
            PackageError::InvalidFormat("could not allocate Pages section lookup".to_owned())
        })?;
    requested.extend(
        references
            .iter()
            .enumerate()
            .map(|(index, reference)| (reference.identifier.get(), index)),
    );
    requested.sort_unstable_by_key(|(identifier, _index)| *identifier);

    let mut names: Vec<Option<Box<str>>> = Vec::new();
    names
        .try_reserve_exact(references.len())
        .map_err(|_error| {
            PackageError::InvalidFormat("could not allocate Pages section names".to_owned())
        })?;
    names.resize_with(references.len(), || None);
    let mut found = Vec::new();
    found
        .try_reserve_exact(references.len())
        .map_err(|_error| {
            PackageError::InvalidFormat("could not allocate Pages section-name state".to_owned())
        })?;
    found.resize(references.len(), false);
    let mut retained_bytes = 0usize;

    for component in components.iter() {
        for object in &component.archive().objects {
            let Some(identifier) = object.archive_info.identifier else {
                continue;
            };
            let Ok(request_index) =
                requested.binary_search_by_key(&identifier, |(candidate, _index)| *candidate)
            else {
                continue;
            };
            let destination = requested[request_index].1;
            let (name, next_retained_bytes) = decode_section_name(
                object,
                references[destination].identifier,
                retained_bytes,
                max_name_text_bytes,
            )?;
            retained_bytes = next_retained_bytes;
            names[destination] = name;
            found[destination] = true;
        }
    }

    if let Some(index) = found.iter().position(|is_found| !is_found) {
        return Err(PackageError::InvalidFormat(format!(
            "Pages section object {} is missing",
            references[index].identifier
        )));
    }
    Ok(names)
}

fn decode_section_name(
    object: &litchi_iwa_core::ArchiveObject,
    identifier: NonZeroU64,
    retained_bytes: usize,
    max_retained_bytes: usize,
) -> PackageResult<(Option<Box<str>>, usize)> {
    let context = format!("Pages section object {identifier}");
    let payload = unique_message_payload(&object.messages, SECTION_MESSAGE_TYPE, &context)?;
    let view = parse_wire(payload, &context)?;
    let Some(field) = unique_wire_field(&view, 26, 2, false, &context)? else {
        return Ok((None, retained_bytes));
    };
    let name = std::str::from_utf8(field.payload()).map_err(|error| {
        PackageError::InvalidFormat(format!(
            "Pages section object {identifier} name is not valid UTF-8: {error}"
        ))
    })?;
    let next_retained_bytes =
        retained_bytes
            .checked_add(name.len())
            .ok_or(PackageError::SectionNamesTooLarge {
                observed: usize::MAX,
                limit: max_retained_bytes,
            })?;
    if next_retained_bytes > max_retained_bytes {
        return Err(PackageError::SectionNamesTooLarge {
            observed: next_retained_bytes,
            limit: max_retained_bytes,
        });
    }
    let mut owned = String::new();
    owned
        .try_reserve_exact(name.len())
        .map_err(|_error| PackageError::Allocation { amount: name.len() })?;
    owned.push_str(name);
    Ok((Some(owned.into_boxed_str()), next_retained_bytes))
}

fn section_text_ranges(
    text: &str,
    references: &[NativeSectionReference],
    body_identifier: NonZeroU64,
) -> PackageResult<Vec<TextRange>> {
    let mut points = Vec::new();
    points
        .try_reserve_exact(references.len())
        .map_err(|_error| {
            PackageError::InvalidFormat("could not allocate Pages section boundaries".to_owned())
        })?;
    let mut reference_index = 0usize;
    let mut utf16_offset = 0usize;
    let mut byte_offset = 0usize;
    let mut preceding_character = None;

    for character in text.chars() {
        capture_section_boundary(
            references,
            &mut reference_index,
            utf16_offset,
            byte_offset,
            preceding_character,
            &mut points,
        )?;
        let next_utf16 = utf16_offset
            .checked_add(character.len_utf16())
            .ok_or_else(|| {
                PackageError::InvalidFormat(format!(
                    "Pages body object {body_identifier} UTF-16 length overflows usize"
                ))
            })?;
        if references.get(reference_index).is_some_and(|reference| {
            let target = reference.character_index as usize;
            target > utf16_offset && target < next_utf16
        }) {
            return Err(PackageError::InvalidFormat(format!(
                "Pages section {} boundary {} splits a UTF-16 surrogate pair",
                references[reference_index].identifier, references[reference_index].character_index
            )));
        }
        utf16_offset = next_utf16;
        byte_offset = byte_offset
            .checked_add(character.len_utf8())
            .ok_or_else(|| {
                PackageError::InvalidFormat(format!(
                    "Pages body object {body_identifier} UTF-8 length overflows usize"
                ))
            })?;
        preceding_character = Some(character);
    }
    capture_section_boundary(
        references,
        &mut reference_index,
        utf16_offset,
        byte_offset,
        preceding_character,
        &mut points,
    )?;
    if reference_index != references.len() {
        let reference = references[reference_index];
        return Err(PackageError::InvalidFormat(format!(
            "Pages section {} boundary {} exceeds body UTF-16 length {utf16_offset}",
            reference.identifier, reference.character_index
        )));
    }

    let mut ranges = Vec::new();
    ranges.try_reserve_exact(points.len()).map_err(|_error| {
        PackageError::InvalidFormat("could not allocate Pages section text ranges".to_owned())
    })?;
    for (index, point) in points.iter().copied().enumerate() {
        let end = if let Some(next) = points.get(index + 1).copied() {
            if next.preceding_character != Some('\u{0004}') {
                return Err(PackageError::InvalidFormat(format!(
                    "Pages section {} boundary {} is not preceded by a native section-break marker",
                    references[index + 1].identifier,
                    references[index + 1].character_index
                )));
            }
            next.byte_offset.checked_sub(1).ok_or_else(|| {
                PackageError::InvalidFormat(
                    "Pages section boundary underflows UTF-8 text".to_owned(),
                )
            })?
        } else {
            byte_offset
        };
        if end < point.byte_offset {
            return Err(PackageError::InvalidFormat(format!(
                "Pages section {} has an invalid text range",
                references[index].identifier
            )));
        }
        ranges.push(TextRange {
            start: point.byte_offset,
            end,
        });
    }
    Ok(ranges)
}

fn capture_section_boundary(
    references: &[NativeSectionReference],
    reference_index: &mut usize,
    utf16_offset: usize,
    byte_offset: usize,
    preceding_character: Option<char>,
    points: &mut Vec<BoundaryPoint>,
) -> PackageResult<()> {
    let Some(reference) = references.get(*reference_index) else {
        return Ok(());
    };
    let target = reference.character_index as usize;
    if target < utf16_offset {
        return Err(PackageError::InvalidFormat(format!(
            "Pages section {} boundary {} is not on a UTF-16 character boundary",
            reference.identifier, reference.character_index
        )));
    }
    if target == utf16_offset {
        points.push(BoundaryPoint {
            byte_offset,
            preceding_character,
        });
        *reference_index += 1;
    }
    Ok(())
}

fn split_native_text(
    storage: &Storage,
    ranges: &[TextRange],
    body_identifier: NonZeroU64,
) -> PackageResult<Vec<Storage>> {
    if storage.runs().len() > litchi_iwa_text_wire::MAX_FRAGMENTS {
        return Err(PackageError::InvalidFormat(format!(
            "Pages body object {body_identifier} contains {} text fragments; maximum is {}",
            storage.runs().len(),
            litchi_iwa_text_wire::MAX_FRAGMENTS
        )));
    }

    let mut accumulators = Vec::new();
    accumulators
        .try_reserve_exact(ranges.len())
        .map_err(|_error| {
            PackageError::InvalidFormat("could not allocate Pages section text storage".to_owned())
        })?;
    for range in ranges {
        accumulators.push(StorageAccumulator::with_capacity(range.end - range.start)?);
    }

    let mut expected_start = 0usize;
    let mut section_index = 0usize;
    for run in storage.runs().iter().copied() {
        if run.start() != expected_start {
            return Err(PackageError::InvalidFormat(format!(
                "Pages body object {body_identifier} lazy text runs are not contiguous"
            )));
        }
        let global_start = run.start();
        let global_end = run.end().ok_or_else(|| {
            PackageError::InvalidFormat(format!(
                "Pages body object {body_identifier} UTF-8 length overflows usize"
            ))
        })?;
        let fragment = storage
            .text()
            .get(global_start..global_end)
            .ok_or_else(|| {
                PackageError::InvalidFormat(format!(
                    "Pages body object {body_identifier} lazy text run is invalid"
                ))
            })?;
        if run.is_empty() {
            if let Some(index) = range_containing_empty_offset(ranges, global_start) {
                accumulators[index].push_empty()?;
            }
            expected_start = global_end;
            continue;
        }

        let mut cursor = global_start;
        while cursor < global_end && section_index < ranges.len() {
            let range = ranges[section_index];
            if cursor >= range.end {
                section_index += 1;
                continue;
            }
            if cursor < range.start {
                cursor = global_end.min(range.start);
                continue;
            }
            let overlap_end = global_end.min(range.end);
            if cursor < overlap_end {
                let local_start = cursor - global_start;
                let local_end = overlap_end - global_start;
                let slice = fragment.get(local_start..local_end).ok_or_else(|| {
                    PackageError::InvalidFormat(format!(
                        "Pages body object {body_identifier} section boundary is not on a UTF-8 character boundary"
                    ))
                })?;
                accumulators[section_index].push(slice)?;
                cursor = overlap_end;
            }
        }
        expected_start = global_end;
    }
    if expected_start != storage.len() {
        return Err(PackageError::InvalidFormat(format!(
            "Pages body object {body_identifier} lazy text runs do not cover the materialized text"
        )));
    }

    let mut storages = Vec::new();
    storages
        .try_reserve_exact(accumulators.len())
        .map_err(|_error| {
            PackageError::InvalidFormat("could not allocate Pages section storages".to_owned())
        })?;
    for accumulator in accumulators {
        storages.push(accumulator.finish(body_identifier)?);
    }
    Ok(storages)
}

fn range_containing_empty_offset(ranges: &[TextRange], offset: usize) -> Option<usize> {
    let insertion = ranges.partition_point(|range| range.start <= offset);
    let index = insertion.checked_sub(1)?;
    (offset <= ranges[index].end).then_some(index)
}

fn extract_storages(
    components: &ComponentCatalog,
    max_sections: usize,
    max_text_bytes: usize,
    limits: Limits,
) -> PackageResult<Vec<Storage>> {
    let mut storages = Vec::new();
    let mut text_bytes = 0usize;
    let validation_limits =
        storage_rewrite_limits(limits).map_err(|limit_error| match limit_error {
            StorageWireLimitsError::Physical(physical_error) => {
                PackageError::Archive(physical_error)
            },
            StorageWireLimitsError::Wire(wire_error) => PackageError::InvalidFormat(format!(
                "Pages fallback text validation limits are invalid: {wire_error}"
            )),
        })?;

    for component in components.iter() {
        for object in &component.archive().objects {
            if !object
                .messages
                .iter()
                .any(|message| is_fallback_trigger_type(message.type_))
            {
                continue;
            }
            let Some(identifier) = object.archive_info.identifier.and_then(NonZeroU64::new) else {
                continue;
            };
            let mut object_storages = Vec::new();
            let mut object_text_bytes = 0usize;
            let mut object_fragments = 0usize;
            for message in &object.messages {
                if !is_storage_message_type(message.type_) {
                    continue;
                }
                let validation_result = litchi_iwa_text_wire::validate_storage_with_limits(
                    message.data.as_slice(),
                    validation_limits,
                );
                let validation = match validation_result {
                    Ok(validation) => validation,
                    Err(error) => {
                        skip_or_reject_fallback_validation(error)?;
                        continue;
                    },
                };
                let next_fragments = object_fragments.checked_add(validation.fragments()).ok_or(
                    PackageError::Semantic(SemanticError::TextTooLarge {
                        observed: usize::MAX,
                        limit: max_text_bytes,
                    }),
                )?;
                let separator_bytes = next_fragments.saturating_sub(1);
                let next_object_text = object_text_bytes.checked_add(validation.utf8_len()).ok_or(
                    PackageError::Semantic(SemanticError::TextTooLarge {
                        observed: usize::MAX,
                        limit: max_text_bytes,
                    }),
                )?;
                let charged_object_text =
                    next_object_text
                        .checked_add(separator_bytes)
                        .ok_or(PackageError::Semantic(SemanticError::TextTooLarge {
                            observed: usize::MAX,
                            limit: max_text_bytes,
                        }))?;
                let charged_total =
                    text_bytes
                        .checked_add(charged_object_text)
                        .ok_or(PackageError::Semantic(SemanticError::TextTooLarge {
                            observed: usize::MAX,
                            limit: max_text_bytes,
                        }))?;
                if charged_total > max_text_bytes {
                    return Err(PackageError::Semantic(SemanticError::TextTooLarge {
                        observed: charged_total,
                        limit: max_text_bytes,
                    }));
                }
                let remaining_text = max_text_bytes
                    .saturating_sub(text_bytes)
                    .saturating_sub(object_text_bytes)
                    .max(1);
                let text_limits = litchi_iwa_text_wire::Limits::new(
                    message.data.len().max(1),
                    litchi_iwa_text_wire::DEFAULT_MAX_FIELDS,
                    litchi_iwa_text_wire::DEFAULT_MAX_WIRE_FRAGMENTS,
                    remaining_text,
                )
                .map_err(|error| {
                    PackageError::InvalidFormat(format!(
                        "Pages fallback text limits are invalid: {error}"
                    ))
                })?;
                // Fallback discovery is deliberately speculative, matching
                // the legacy archive-wide projection: a recognized type whose
                // payload is not a strict TSWP storage is not promoted into
                // semantic state. Valid candidates pass raw-wire preflight and
                // the bounded Buffa projection before any text is published.
                let storage_result = litchi_iwa_text_wire::from_bytes_with_limits(
                    message.data.as_slice(),
                    text_limits,
                );
                let storage = match storage_result {
                    Ok(storage) => storage,
                    Err(error) => {
                        skip_or_reject_fallback_materialization(error, text_bytes, max_text_bytes)?;
                        continue;
                    },
                };
                object_storages
                    .try_reserve(1)
                    .map_err(|_error| PackageError::Allocation { amount: 1 })?;
                object_storages.push(storage);
                object_text_bytes = next_object_text;
                object_fragments = next_fragments;
            }
            let storage = legacy_fallback_object_storage(
                &object_storages,
                identifier,
                max_text_bytes.saturating_sub(text_bytes),
            )?;
            if storage.is_empty() {
                continue;
            }
            if max_sections == 0 {
                return Err(PackageError::Semantic(SemanticError::TooManySections {
                    actual: 1,
                    limit: max_sections,
                }));
            }
            if storages.len() == MAX_BODY_STORAGES {
                return Err(PackageError::Semantic(SemanticError::TooManyBodyStorages {
                    actual: storages.len().saturating_add(1),
                    limit: MAX_BODY_STORAGES,
                }));
            }
            let next_text_bytes = text_bytes.checked_add(storage.len()).ok_or({
                PackageError::Semantic(SemanticError::TextTooLarge {
                    observed: usize::MAX,
                    limit: max_text_bytes,
                })
            })?;
            if next_text_bytes > max_text_bytes {
                return Err(PackageError::Semantic(SemanticError::TextTooLarge {
                    observed: next_text_bytes,
                    limit: max_text_bytes,
                }));
            }
            text_bytes = next_text_bytes;
            storages.push(storage);
        }
    }
    Ok(storages)
}

fn skip_or_reject_fallback_validation(
    error: litchi_iwa_text_wire::RewriteError,
) -> PackageResult<()> {
    match error {
        litchi_iwa_text_wire::RewriteError::InvalidFormat(_) => Ok(()),
        litchi_iwa_text_wire::RewriteError::Allocation { amount, .. } => {
            Err(PackageError::Allocation { amount })
        },
        litchi_iwa_text_wire::RewriteError::LimitExceeded {
            observed, limit, ..
        } => Err(PackageError::PayloadLimit { observed, limit }),
        litchi_iwa_text_wire::RewriteError::Projection(_) => Err(PackageError::InvalidFormat(
            "Pages fallback text projection disagreed with validated input".to_owned(),
        )),
        _ => Err(PackageError::InvalidFormat(
            "Pages fallback text validation could not complete".to_owned(),
        )),
    }
}

fn skip_or_reject_fallback_materialization(
    error: litchi_iwa_text_wire::Error,
    retained_text_bytes: usize,
    max_text_bytes: usize,
) -> PackageResult<()> {
    match error {
        litchi_iwa_text_wire::Error::InvalidUtf8 { .. }
        | litchi_iwa_text_wire::Error::WrongTextWireType { .. }
        | litchi_iwa_text_wire::Error::Storage(_) => Ok(()),
        litchi_iwa_text_wire::Error::TooManyTextBytes { actual, .. } => {
            Err(PackageError::Semantic(SemanticError::TextTooLarge {
                observed: retained_text_bytes.saturating_add(actual),
                limit: max_text_bytes,
            }))
        },
        litchi_iwa_text_wire::Error::TextLengthOverflow => {
            Err(PackageError::Semantic(SemanticError::TextTooLarge {
                observed: usize::MAX,
                limit: max_text_bytes,
            }))
        },
        litchi_iwa_text_wire::Error::TooManyFragments { actual, limit } => {
            Err(PackageError::PayloadLimit {
                observed: actual,
                limit,
            })
        },
        litchi_iwa_text_wire::Error::Common(litchi_iwa_common::Error::Allocation {
            amount,
            ..
        }) => Err(PackageError::Allocation { amount }),
        litchi_iwa_text_wire::Error::Common(litchi_iwa_common::Error::LimitExceeded {
            observed,
            limit,
            ..
        }) => Err(PackageError::PayloadLimit { observed, limit }),
        litchi_iwa_text_wire::Error::ProjectionDecode { .. }
        | litchi_iwa_text_wire::Error::ProjectionMismatch { .. }
        | litchi_iwa_text_wire::Error::ProjectionTextLengthMismatch { .. } => {
            Err(PackageError::InvalidFormat(
                "Pages fallback text projection disagreed with validated input".to_owned(),
            ))
        },
        _ => Err(PackageError::InvalidFormat(
            "Pages fallback text materialization could not complete".to_owned(),
        )),
    }
}

fn legacy_fallback_object_storage(
    storages: &[Storage],
    identifier: NonZeroU64,
    max_text_bytes: usize,
) -> PackageResult<Storage> {
    let fragment_count = storages
        .iter()
        .try_fold(0usize, |count, storage| {
            count.checked_add(storage.runs().len())
        })
        .ok_or(PackageError::Semantic(SemanticError::TextTooLarge {
            observed: usize::MAX,
            limit: max_text_bytes,
        }))?;
    if fragment_count == 0 {
        return Ok(Storage::new());
    }
    let joined_len = storages
        .iter()
        .try_fold(fragment_count - 1, |length, storage| {
            length.checked_add(storage.len())
        })
        .ok_or(PackageError::Semantic(SemanticError::TextTooLarge {
            observed: usize::MAX,
            limit: max_text_bytes,
        }))?;
    if joined_len > max_text_bytes {
        return Err(PackageError::Semantic(SemanticError::TextTooLarge {
            observed: joined_len,
            limit: max_text_bytes,
        }));
    }

    let mut joined = String::new();
    joined
        .try_reserve_exact(joined_len)
        .map_err(|_error| PackageError::Allocation { amount: joined_len })?;
    let mut fragment_index = 0usize;
    for storage in storages {
        for run in storage.runs().iter().copied() {
            if fragment_index != 0 {
                joined.push('\n');
            }
            fragment_index += 1;
            let end = run.end().ok_or_else(|| {
                PackageError::InvalidFormat(format!(
                    "Pages fallback storage object {identifier} contains an overflowing text range"
                ))
            })?;
            let fragment = storage.text().get(run.start()..end).ok_or_else(|| {
                PackageError::InvalidFormat(format!(
                    "Pages fallback storage object {identifier} contains an invalid text range"
                ))
            })?;
            joined.push_str(fragment);
        }
    }
    Ok(Storage::from_text(joined))
}

fn unique_message_payload<'a>(
    messages: &'a [litchi_iwa_core::RawMessage],
    message_type: u32,
    context: &str,
) -> PackageResult<&'a [u8]> {
    let mut payload = None;
    for message in messages {
        if message.type_ == message_type && payload.replace(message.data.as_slice()).is_some() {
            return Err(PackageError::InvalidFormat(format!(
                "{context} contains duplicate type-{message_type} payloads"
            )));
        }
    }
    payload.ok_or_else(|| {
        PackageError::InvalidFormat(format!("{context} has no type-{message_type} payload"))
    })
}

fn unique_text_payload(
    messages: &[litchi_iwa_core::RawMessage],
    identifier: NonZeroU64,
) -> PackageResult<&[u8]> {
    let mut payload = None;
    for message in messages {
        if is_body_text_message_type(message.type_)
            && payload.replace(message.data.as_slice()).is_some()
        {
            return Err(PackageError::InvalidFormat(format!(
                "Pages body storage object {identifier} contains duplicate text payloads"
            )));
        }
    }
    payload.ok_or_else(|| {
        PackageError::InvalidFormat(format!(
            "Pages body object {identifier} has no type-2001/type-2022 text payload"
        ))
    })
}

fn effective_text_limit(limits: Limits) -> usize {
    limits.max_iwa_stream_bytes().min(DEFAULT_MAX_TEXT_BYTES)
}

fn storage_rewrite_limits(
    limits: Limits,
) -> Result<litchi_iwa_text_wire::RewriteLimits, StorageWireLimitsError> {
    let archive_limits = limits
        .effective_archive_limits()
        .map_err(StorageWireLimitsError::Physical)?;
    let maximum = archive_limits.max_message_bytes();
    let common = WireLimits::default();
    let fields = common.max_fields().min(archive_limits.max_header_fields());
    let nesting = common
        .max_nesting()
        .min(archive_limits.max_header_nesting());
    let fragments = fields.min(litchi_iwa_text_wire::MAX_FRAGMENTS);
    let table_entries = fields;
    let object_references = fields.min(archive_limits.max_metadata_items());
    litchi_iwa_text_wire::RewriteLimits::new(
        maximum,
        fields,
        nesting,
        fragments,
        effective_text_limit(limits),
        table_entries,
        object_references,
        maximum,
        common.max_rewrite_work(),
    )
    .map_err(StorageWireLimitsError::Wire)
}

const fn is_body_text_message_type(type_id: u32) -> bool {
    matches!(type_id, 2001 | 2022)
}

const fn is_storage_message_type(type_id: u32) -> bool {
    matches!(type_id, 2001..=2014 | 2022)
}

const fn is_fallback_trigger_type(type_id: u32) -> bool {
    matches!(
        type_id,
        200 | 201 | 202 | 203 | 204 | 205 | 2001 | 2002 | 2003 | 2004 | 2005 | 2011 | 2012 | 2022
    )
}

#[cfg(test)]
mod tests {
    use std::fs;
    use std::io::{self, Cursor, Write};

    use super::*;
    use litchi_iwa_core::{Archive, ArchiveObject, RawMessage, SnappyStream};
    use litchi_iwa_protos::tswp::{
        self, ObjectAttributeTable, object_attribute_table::ObjectAttribute,
    };
    use litchi_iwa_protos::{tp, tsp::Reference};
    use prost::Message;

    struct InterruptedOnce<R> {
        inner: R,
        pending: bool,
    }

    impl<R> InterruptedOnce<R> {
        const fn new(inner: R) -> Self {
            Self {
                inner,
                pending: true,
            }
        }
    }

    impl<R: Read> Read for InterruptedOnce<R> {
        fn read(&mut self, buffer: &mut [u8]) -> io::Result<usize> {
            if self.pending {
                self.pending = false;
                return Err(io::ErrorKind::Interrupted.into());
            }
            self.inner.read(buffer)
        }
    }

    #[derive(Clone, Copy)]
    enum PrefixWriteBehavior {
        Failure,
        Zero,
        OverReport,
    }

    struct PrefixWriter {
        prefix_remaining: usize,
        behavior: PrefixWriteBehavior,
        output: Vec<u8>,
        flushes: usize,
    }

    impl PrefixWriter {
        const fn new(behavior: PrefixWriteBehavior) -> Self {
            Self {
                prefix_remaining: 4,
                behavior,
                output: Vec::new(),
                flushes: 0,
            }
        }
    }

    impl Write for PrefixWriter {
        fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
            if self.prefix_remaining != 0 {
                let amount = bytes.len().min(self.prefix_remaining);
                self.output.extend_from_slice(&bytes[..amount]);
                self.prefix_remaining -= amount;
                return Ok(amount);
            }

            match self.behavior {
                PrefixWriteBehavior::Failure => Err(io::Error::other("pages sink authored secret")),
                PrefixWriteBehavior::Zero => Ok(0),
                PrefixWriteBehavior::OverReport => Ok(bytes.len().saturating_add(1)),
            }
        }

        fn flush(&mut self) -> io::Result<()> {
            self.flushes += 1;
            Ok(())
        }
    }

    struct PrefixInterruptedWriter {
        interruptions_remaining: usize,
        interruptions_observed: usize,
        output: Vec<u8>,
        flushes: usize,
    }

    impl PrefixInterruptedWriter {
        const fn new() -> Self {
            Self {
                interruptions_remaining: 3,
                interruptions_observed: 0,
                output: Vec::new(),
                flushes: 0,
            }
        }
    }

    impl Write for PrefixInterruptedWriter {
        fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
            if self.interruptions_remaining != 0 {
                self.interruptions_remaining -= 1;
                self.interruptions_observed += 1;
                return Err(io::ErrorKind::Interrupted.into());
            }
            let amount = bytes.len().min(1);
            self.output.extend_from_slice(&bytes[..amount]);
            Ok(amount)
        }

        fn flush(&mut self) -> io::Result<()> {
            self.flushes += 1;
            Ok(())
        }
    }

    fn limits_with_input_bytes(max_input_bytes: u64) -> PackageResult<Limits> {
        let defaults = Limits::default();
        Ok(Limits::new(
            max_input_bytes,
            defaults.max_entries(),
            defaults.max_entry_bytes(),
            defaults.max_total_bytes(),
            defaults.max_iwa_stream_bytes(),
        )?)
    }

    fn package_bytes(
        body: Option<&str>,
        root_references_body: bool,
        metadata: bool,
    ) -> PackageResult<Vec<u8>> {
        let body_identifier = 42;
        let root = tp::DocumentArchive {
            body_storage: (root_references_body && body.is_some()).then(|| Reference {
                identifier: body_identifier,
                ..Reference::default()
            }),
            ..tp::DocumentArchive::default()
        };
        let mut objects = vec![
            ArchiveObject::new(
                1,
                vec![RawMessage {
                    type_: 10_000,
                    data: root.encode_to_vec(),
                }],
            )
            .map_err(|error| PackageError::InvalidFormat(error.to_string()))?,
        ];
        if let Some(body_text) = body {
            let storage = tswp::StorageArchive {
                text: vec![body_text.to_owned()],
                ..tswp::StorageArchive::default()
            };
            objects.push(
                ArchiveObject::new(
                    body_identifier,
                    vec![RawMessage {
                        type_: 2001,
                        data: storage.encode_to_vec(),
                    }],
                )
                .map_err(|error| PackageError::InvalidFormat(error.to_string()))?,
            );
        }
        archive_package_bytes(objects, metadata)
    }

    fn sectioned_package_bytes(
        text: Vec<String>,
        boundaries: &[(u32, u64)],
        root_section: Option<u64>,
        section_objects: Vec<(u64, Vec<RawMessage>)>,
    ) -> PackageResult<Vec<u8>> {
        let body_identifier = 42;
        let root = tp::DocumentArchive {
            body_storage: Some(Reference {
                identifier: body_identifier,
                ..Reference::default()
            }),
            section: root_section.map(|identifier| Reference {
                identifier,
                ..Reference::default()
            }),
            ..tp::DocumentArchive::default()
        };
        let table_section = (!boundaries.is_empty()).then(|| ObjectAttributeTable {
            entries: boundaries
                .iter()
                .map(|&(character_index, identifier)| ObjectAttribute {
                    character_index,
                    object: Some(Reference {
                        identifier,
                        ..Reference::default()
                    }),
                })
                .collect(),
        });
        let storage = tswp::StorageArchive {
            text,
            table_section,
            ..tswp::StorageArchive::default()
        };
        let mut objects = Vec::new();
        objects
            .try_reserve_exact(section_objects.len().saturating_add(2))
            .map_err(|_error| {
                PackageError::InvalidFormat("could not allocate test objects".to_owned())
            })?;
        objects.push(
            ArchiveObject::new(
                1,
                vec![RawMessage {
                    type_: 10_000,
                    data: root.encode_to_vec(),
                }],
            )
            .map_err(|error| PackageError::InvalidFormat(error.to_string()))?,
        );
        objects.push(
            ArchiveObject::new(
                body_identifier,
                vec![RawMessage {
                    type_: 2001,
                    data: storage.encode_to_vec(),
                }],
            )
            .map_err(|error| PackageError::InvalidFormat(error.to_string()))?,
        );
        for (identifier, messages) in section_objects {
            objects.push(
                ArchiveObject::new(identifier, messages)
                    .map_err(|error| PackageError::InvalidFormat(error.to_string()))?,
            );
        }
        archive_package_bytes(objects, false)
    }

    fn body_footnote_package_bytes() -> PackageResult<Vec<u8>> {
        let body_identifier = 42;
        let root = tp::DocumentArchive {
            body_storage: Some(Reference {
                identifier: body_identifier,
                ..Reference::default()
            }),
            ..tp::DocumentArchive::default()
        };
        let body_text = "A😀\u{e}B\u{e}C";
        let body = tswp::StorageArchive {
            kind: Some(tswp::storage_archive::KindType::Body as i32),
            text: vec![body_text.to_owned()],
            table_footnote: Some(ObjectAttributeTable {
                entries: vec![
                    ObjectAttribute {
                        character_index: 3,
                        object: Some(Reference {
                            identifier: 100,
                            ..Reference::default()
                        }),
                    },
                    ObjectAttribute {
                        character_index: 5,
                        object: Some(Reference {
                            identifier: 110,
                            ..Reference::default()
                        }),
                    },
                ],
            }),
            ..tswp::StorageArchive::default()
        };
        let footnote = |reference_id: u64,
                        storage_id: u64,
                        marker_id: u64,
                        text: &str,
                        custom_mark: Option<&str>| {
            let mut reference = tswp::FootnoteReferenceAttachmentArchive {
                super_: Some(tswp::TextualAttachmentArchive {
                    string_equivalent: Some("*".to_owned()),
                    kind: Some(tswp::textual_attachment_archive::Kind::KKindFootnoteMark as i32),
                }),
                contained_storage: Some(Reference {
                    identifier: storage_id,
                    ..Reference::default()
                }),
                custom_mark_string: custom_mark.map(str::to_owned),
            }
            .encode_to_vec();
            reference.extend_from_slice(&[0xa0, 0x06, 0x01, 0xaa, 0x06, 0x03, b'r', b'e', b'f']);

            let storage = tswp::StorageArchive {
                kind: Some(tswp::storage_archive::KindType::Footnote as i32),
                text: vec![format!("{STORAGE_TEXT_PREFIX}{text}")],
                table_attachment: Some(ObjectAttributeTable {
                    entries: vec![ObjectAttribute {
                        character_index: 0,
                        object: Some(Reference {
                            identifier: marker_id,
                            ..Reference::default()
                        }),
                    }],
                }),
                ..tswp::StorageArchive::default()
            };
            let mut marker = tswp::TextualAttachmentArchive {
                string_equivalent: Some("*".to_owned()),
                kind: Some(tswp::textual_attachment_archive::Kind::KKindFootnoteMark as i32),
            }
            .encode_to_vec();
            marker.extend_from_slice(&[0xa0, 0x06, 0x01, 0xaa, 0x06, 0x03, b'm', b'a', b'r']);
            [
                ArchiveObject::new(
                    reference_id,
                    vec![RawMessage {
                        type_: FOOTNOTE_REFERENCE_MESSAGE_TYPE,
                        data: reference,
                    }],
                )
                .map_err(|error| PackageError::InvalidFormat(error.to_string())),
                ArchiveObject::new(
                    storage_id,
                    vec![RawMessage {
                        type_: 2_001,
                        data: storage.encode_to_vec(),
                    }],
                )
                .map_err(|error| PackageError::InvalidFormat(error.to_string())),
                ArchiveObject::new(
                    marker_id,
                    vec![RawMessage {
                        type_: TEXTUAL_ATTACHMENT_MESSAGE_TYPE,
                        data: marker,
                    }],
                )
                .map_err(|error| PackageError::InvalidFormat(error.to_string())),
            ]
        };
        let first = footnote(100, 101, 102, "First", None);
        let second = footnote(110, 111, 112, "Second", Some("†"));
        let mut objects = vec![
            ArchiveObject::new(
                1,
                vec![RawMessage {
                    type_: 10_000,
                    data: root.encode_to_vec(),
                }],
            )
            .map_err(|error| PackageError::InvalidFormat(error.to_string()))?,
            ArchiveObject::new(
                body_identifier,
                vec![RawMessage {
                    type_: 2_001,
                    data: body.encode_to_vec(),
                }],
            )
            .map_err(|error| PackageError::InvalidFormat(error.to_string()))?,
        ];
        objects.extend(first.into_iter().collect::<PackageResult<Vec<_>>>()?);
        objects.extend(second.into_iter().collect::<PackageResult<Vec<_>>>()?);
        archive_package_bytes(objects, false)
    }

    fn archive_package_bytes(
        objects: Vec<ArchiveObject>,
        metadata: bool,
    ) -> PackageResult<Vec<u8>> {
        let archive = Archive { objects };
        let compressed = SnappyStream::compress(
            archive
                .to_bytes()
                .map_err(|error| PackageError::InvalidFormat(error.to_string()))?
                .as_slice(),
        )
        .map_err(|error| PackageError::InvalidFormat(error.to_string()))?;

        let mut entries = vec![("Index/Document.iwa", compressed.as_slice())];
        if metadata {
            entries.push((
                "Metadata/Properties.plist",
                br#"<?xml version="1.0" encoding="UTF-8"?><plist version="1.0"><dict><key>Title</key><string>Report</string><key>Author</key><string>Ada</string><key>Application</key><string>Pages</string><key>revision</key><string>3</string><key>fileFormatVersion</key><string>7</string></dict></plist>"#
                    .as_slice(),
            ));
            entries.push(("Metadata/DocumentIdentifier", b"pages-id\n".as_slice()));
        }
        Ok(litchi_iwa_archive::package::to_bytes(
            entries,
            Limits::default(),
        )?)
    }

    fn package_bytes_with_opaque_member() -> PackageResult<Vec<u8>> {
        let source = package_bytes(Some("Opaque body"), true, true)?;
        let catalog = Catalog::from_bytes(&source)?;
        let mut entries: Vec<(&str, &[u8])> = catalog
            .iter()
            .map(|entry| (entry.name(), entry.data()))
            .collect();
        entries.push(("Data/opaque.bin", b"opaque Pages payload"));
        let mut bytes = litchi_iwa_archive::package::to_bytes(entries, Limits::default())?;

        const NAME: &[u8] = b"Data/opaque.bin";
        const UNSUPPORTED_METHOD: [u8; 2] = 99_u16.to_le_bytes();
        let mut cursor = 0usize;
        let mut changed = 0usize;
        while let Some(relative) = bytes[cursor..]
            .windows(NAME.len())
            .position(|candidate| candidate == NAME)
        {
            let position = cursor + relative;
            if position >= 30 && bytes[position - 30..position - 26] == [0x50, 0x4b, 0x03, 0x04] {
                bytes[position - 22..position - 20].copy_from_slice(&UNSUPPORTED_METHOD);
                changed += 1;
            } else if position >= 46
                && bytes[position - 46..position - 42] == [0x50, 0x4b, 0x01, 0x02]
            {
                bytes[position - 36..position - 34].copy_from_slice(&UNSUPPORTED_METHOD);
                changed += 1;
            }
            cursor = position.saturating_add(NAME.len());
        }
        if changed != 2 {
            return Err(PackageError::InvalidFormat(
                "opaque Pages test member did not have two ZIP records".to_owned(),
            ));
        }
        Ok(bytes)
    }

    fn section_payload(name: Option<&str>) -> RawMessage {
        RawMessage {
            type_: SECTION_MESSAGE_TYPE,
            data: tp::SectionArchive {
                name: name.map(str::to_owned),
                ..tp::SectionArchive::default()
            }
            .encode_to_vec(),
        }
    }

    #[test]
    fn bounded_reader_accepts_exact_limit_and_detects_growth() -> PackageResult<()> {
        let limits = limits_with_input_bytes(4)?;

        let mut exact = Cursor::new([1_u8, 2, 3, 4]);
        assert_eq!(
            read_source_with_reported_length(&mut exact, 4, limits)?.as_ref(),
            &[1, 2, 3, 4]
        );
        assert_eq!(exact.position(), 4);

        let mut initially_empty = Cursor::new([1_u8, 2, 3, 4]);
        assert_eq!(
            read_source_with_reported_length(&mut initially_empty, 0, limits)?.as_ref(),
            &[1, 2, 3, 4]
        );
        assert_eq!(initially_empty.position(), 4);

        let growing_bytes = vec![0x5a_u8; 20 * 1024];
        let growing_limits = limits_with_input_bytes(20 * 1024)?;
        let mut growing = Cursor::new(growing_bytes.as_slice());
        assert_eq!(
            read_source_with_reported_length(&mut growing, 0, growing_limits)?.as_ref(),
            growing_bytes
        );

        let mut grew_over_limit = Cursor::new([1_u8, 2, 3, 4, 5]);
        let growth_error = read_source_with_reported_length(&mut grew_over_limit, 0, limits)
            .err()
            .unwrap_or_else(|| panic!("one byte beyond the input limit must fail"));
        assert!(
            growth_error
                .to_string()
                .contains("Pages package input exceeds the 4 byte limit")
        );
        assert_eq!(grew_over_limit.position(), 5);

        let mut overreported = Cursor::new([1_u8, 2, 3, 4, 5]);
        let reported_length_error = read_source_with_reported_length(&mut overreported, 5, limits)
            .err()
            .unwrap_or_else(|| panic!("an oversized reported length must fail"));
        assert!(reported_length_error.to_string().contains("4 byte limit"));
        assert_eq!(overreported.position(), 0);
        Ok(())
    }

    #[test]
    fn bounded_reader_retries_interrupted_reads() -> PackageResult<()> {
        let limits = limits_with_input_bytes(4)?;
        let mut reader = InterruptedOnce::new(Cursor::new([1_u8, 2, 3, 4]));

        assert_eq!(
            read_source_with_reported_length(&mut reader, 0, limits)?.as_ref(),
            &[1, 2, 3, 4]
        );
        assert!(!reader.pending);
        Ok(())
    }

    #[test]
    fn descriptor_stability_rejects_length_mismatch_and_mutation() -> PackageResult<()> {
        let directory = tempfile::tempdir()?;
        let path = directory.path().join("stable.pages");
        fs::write(&path, [1_u8, 2, 3, 4])?;
        let file = OpenOptions::new().read(true).write(true).open(&path)?;
        let before = FileSnapshot::from_metadata(&file.metadata()?);

        ensure_source_unchanged(before, before, 4)?;
        let length_error = ensure_source_unchanged(before, before, 3)
            .err()
            .unwrap_or_else(|| panic!("an observed-length mismatch must fail"));
        assert!(length_error.to_string().contains("changed while"));

        file.set_len(5)?;
        let after = FileSnapshot::from_metadata(&file.metadata()?);
        let mutation_error = ensure_source_unchanged(before, after, 4)
            .err()
            .unwrap_or_else(|| panic!("a descriptor mutation must fail"));
        assert!(mutation_error.to_string().contains("changed while"));
        assert!(
            !mutation_error
                .to_string()
                .contains(path.to_string_lossy().as_ref())
        );
        Ok(())
    }

    #[test]
    fn path_reader_retains_one_shared_source_allocation() -> PackageResult<()> {
        let directory = tempfile::tempdir()?;
        let path = directory.path().join("valid-pages-package.pages");
        let expected = package_bytes(Some("Shared source"), true, false)?;
        let mut file = fs::File::create(&path)?;
        file.write_all(&expected)?;
        file.sync_all()?;

        let limits = Limits::default();
        let source = read_path(&path, limits)?;
        let source_pointer = source.as_ptr();
        let catalog = SourceCatalog::from_shared_bytes_with_limits(Arc::clone(&source), limits)?;
        assert_eq!(catalog.source_bytes().as_ptr(), source_pointer);
        assert_eq!(catalog.source_bytes(), expected);

        let package = Package::open(&path)?;
        assert_eq!(package.text()?, "Shared source");
        Ok(())
    }

    #[test]
    fn path_errors_reject_non_files_without_disclosing_paths() -> PackageResult<()> {
        let directory = tempfile::tempdir()?;
        let secret_directory = directory.path().join("private-pages-path-do-not-leak");
        fs::create_dir(&secret_directory)?;

        let directory_error = Package::open(&secret_directory)
            .err()
            .unwrap_or_else(|| panic!("a directory must not be accepted as a Pages package"));
        assert!(directory_error.to_string().contains("regular file"));
        assert!(
            !directory_error
                .to_string()
                .contains("private-pages-path-do-not-leak")
        );
        assert!(
            !directory_error
                .to_string()
                .contains(secret_directory.to_string_lossy().as_ref())
        );

        let missing = directory.path().join("private-missing-path-do-not-leak");
        let missing_error = Package::open(&missing)
            .err()
            .unwrap_or_else(|| panic!("a missing path must fail"));
        assert!(
            !missing_error
                .to_string()
                .contains("private-missing-path-do-not-leak")
        );
        assert!(
            !missing_error
                .to_string()
                .contains(missing.to_string_lossy().as_ref())
        );
        Ok(())
    }

    #[cfg(unix)]
    #[test]
    fn unix_path_reader_rejects_symlinks_and_fifos_without_disclosure() -> PackageResult<()> {
        use std::os::unix::fs::symlink;
        use std::process::Command;

        let directory = tempfile::tempdir()?;
        let target = directory.path().join("target.pages");
        fs::write(&target, [0_u8])?;

        let symlink_path = directory.path().join("private-pages-symlink-do-not-leak");
        symlink(&target, &symlink_path)?;
        let symlink_error = Package::open(&symlink_path)
            .err()
            .unwrap_or_else(|| panic!("a symbolic link must not be followed"));
        assert!(matches!(&symlink_error, PackageError::InvalidFormat(_)));
        assert!(
            !symlink_error
                .to_string()
                .contains("private-pages-symlink-do-not-leak")
        );
        assert!(
            !symlink_error
                .to_string()
                .contains(symlink_path.to_string_lossy().as_ref())
        );

        let fifo_path = directory.path().join("private-pages-fifo-do-not-leak");
        let status = Command::new("mkfifo").arg(&fifo_path).status()?;
        if !status.success() {
            return Err(PackageError::InvalidFormat(
                "test could not create a Pages FIFO".to_owned(),
            ));
        }
        let fifo_error = Package::open(&fifo_path)
            .err()
            .unwrap_or_else(|| panic!("a FIFO must not be accepted as a Pages package"));
        assert!(fifo_error.to_string().contains("regular file"));
        assert!(
            !fifo_error
                .to_string()
                .contains("private-pages-fifo-do-not-leak")
        );
        assert!(
            !fifo_error
                .to_string()
                .contains(fifo_path.to_string_lossy().as_ref())
        );
        Ok(())
    }

    #[test]
    fn package_decodes_root_text_metadata_and_shared_snapshots() -> PackageResult<()> {
        let package = Package::from_bytes(&package_bytes(Some("Pages body"), true, true)?)?;

        assert_eq!(package.text()?, "Pages body");
        assert_eq!(package.sections().len(), 1);
        assert_eq!(package.sections()[0].plain_text(), "Pages body");
        assert_eq!(package.stats().total_objects(), 2);
        assert_eq!(package.stats().section_count(), 1);
        assert_eq!(package.metadata().title.as_deref(), Some("Report"));
        assert_eq!(package.metadata().author.as_deref(), Some("Ada"));
        assert_eq!(package.metadata().revision.as_deref(), Some("3"));
        assert_eq!(
            package.metadata().content_status.as_deref(),
            Some("Pages Format Version 7")
        );
        assert_eq!(package.metadata().identifier.as_deref(), Some("pages-id"));
        package.validate()?;

        let snapshot = package.snapshot();
        assert!(std::ptr::eq(
            package.semantic_document(),
            snapshot.semantic_document()
        ));
        Ok(())
    }

    #[test]
    fn package_handles_and_write_errors_are_send_sync() {
        fn assert_send_sync<T: Send + Sync>() {}

        assert_send_sync::<Package>();
        assert_send_sync::<WriteError>();
    }

    #[test]
    fn write_to_accepts_dynamic_dispatch_and_preserves_exact_opaque_bytes()
    -> Result<(), Box<dyn std::error::Error>> {
        let bytes = package_bytes_with_opaque_member()?;
        let package = Package::from_bytes(&bytes)?;
        let snapshot = package.snapshot();

        assert!(Arc::ptr_eq(&package.state, &snapshot.state));
        assert_eq!(
            package.source_bytes().as_ptr(),
            snapshot.source_bytes().as_ptr()
        );
        assert!(
            package
                .state
                .source
                .package()
                .iter()
                .any(|entry| entry.name() == "Data/opaque.bin" && entry.is_opaque())
        );

        let mut output = Vec::new();
        let sink: &mut dyn Write = &mut output;
        package.write_to(sink)?;
        assert_eq!(output, bytes);
        Ok(())
    }

    #[test]
    fn write_to_retries_interruptions_without_flushing() -> PackageResult<()> {
        let bytes = package_bytes(Some("Interrupted Pages output"), true, false)?;
        let package = Package::from_bytes(&bytes)?;
        let mut writer = PrefixInterruptedWriter::new();

        package
            .write_to(&mut writer)
            .map_err(|error| PackageError::InvalidFormat(error.to_string()))?;

        assert_eq!(writer.interruptions_observed, 3);
        assert_eq!(writer.output, bytes);
        assert_eq!(writer.flushes, 0);
        Ok(())
    }

    #[test]
    fn write_to_reports_prefix_progress_for_zero_overreport_and_failure()
    -> Result<(), Box<dyn std::error::Error>> {
        let bytes = package_bytes(Some("Adversarial Pages output"), true, false)?;
        let package = Package::from_bytes(&bytes)?;

        let mut failing = PrefixWriter::new(PrefixWriteBehavior::Failure);
        let failure = package.write_to(&mut failing).unwrap_err();
        assert_eq!(failure.bytes_written(), 4);
        assert_eq!(failing.output, bytes[..4]);
        assert_eq!(failing.flushes, 0);
        assert_eq!(failure.io_error().kind(), io::ErrorKind::Other);
        let display = failure.to_string();
        let debug = format!("{failure:?}");
        assert!(!display.contains("pages sink authored secret"));
        assert!(!debug.contains("pages sink authored secret"));
        assert!(std::error::Error::source(&failure).is_none());
        let underlying = failure.into_io_error();
        assert_eq!(underlying.kind(), io::ErrorKind::Other);
        assert_eq!(underlying.to_string(), "pages sink authored secret");

        let mut zero = PrefixWriter::new(PrefixWriteBehavior::Zero);
        let zero_error = package.write_to(&mut zero).unwrap_err();
        assert_eq!(zero_error.bytes_written(), 4);
        assert_eq!(zero.output, bytes[..4]);
        assert_eq!(zero.flushes, 0);
        assert_eq!(zero_error.io_error().kind(), io::ErrorKind::WriteZero);

        let mut over_report = PrefixWriter::new(PrefixWriteBehavior::OverReport);
        let over_report_error = package.write_to(&mut over_report).unwrap_err();
        assert_eq!(over_report_error.bytes_written(), 4);
        assert_eq!(over_report.output, bytes[..4]);
        assert_eq!(over_report.flushes, 0);
        assert_eq!(
            over_report_error.io_error().kind(),
            io::ErrorKind::InvalidData
        );
        Ok(())
    }

    #[test]
    fn package_projects_body_footnotes_in_source_order_without_mutating_source() -> PackageResult<()>
    {
        let package_bytes = body_footnote_package_bytes()?;
        let package = Package::from_bytes(&package_bytes)?;
        let source_before = package.source_bytes().to_vec();

        let footnotes = package.body_footnotes()?;

        assert_eq!(
            footnotes,
            vec![
                Footnote {
                    position: Position::from_utf16_index(3)
                        .map_err(|error| PackageError::InvalidFormat(error.to_string()))?,
                    text: "First".into(),
                    custom_mark: None,
                },
                Footnote {
                    position: Position::from_utf16_index(5)
                        .map_err(|error| PackageError::InvalidFormat(error.to_string()))?,
                    text: "Second".into(),
                    custom_mark: Some("†".into()),
                },
            ]
        );
        assert_eq!(package.source_bytes(), source_before.as_slice());
        assert_eq!(package.body_footnotes()?, footnotes);
        Ok(())
    }

    #[test]
    fn footnote_semantic_budget_accepts_cap_and_rejects_next() -> PackageResult<()> {
        let mut budget = FootnoteSemanticBudget::new(7);
        budget.charge(4, 3)?;
        assert!(matches!(
            budget.charge(0, 1),
            Err(PackageError::PayloadLimit {
                observed: 8,
                limit: 7,
            })
        ));
        // A rejected charge must not consume the exact-cap allowance.
        budget.charge(0, 0)?;
        Ok(())
    }

    #[test]
    fn footnote_projection_budget_spans_source_and_candidate() -> PackageResult<()> {
        let package = Package::from_bytes(&body_footnote_package_bytes()?)?;
        let source_bytes = package.source_bytes().to_vec();
        let maximum = 14;
        let mut budget = FootnoteSemanticBudget::new(maximum);

        let source = project_body_footnotes_with_budget(
            package.state.source.components(),
            package.state.source.limits(),
            maximum,
            &mut budget,
        )
        .map_err(map_body_storage_decode_error)?;
        assert_eq!(source.len(), 2);
        assert_eq!(budget.retained_bytes, maximum);

        let error = project_body_footnotes_with_budget(
            package.state.source.components(),
            package.state.source.limits(),
            maximum,
            &mut budget,
        )
        .err()
        .unwrap_or_else(|| panic!("a second exact-cap projection must be rejected"));
        assert!(matches!(
            error,
            BodyStorageDecodeError::SemanticLimit {
                observed: 19,
                limit: 14,
            }
        ));
        // The rejected next-byte charge does not consume the exact-cap
        // allowance, and the immutable source was never modified.
        budget.charge(0, 0)?;
        assert_eq!(package.source_bytes(), source_bytes.as_slice());
        assert_eq!(source[0].text.as_ref(), "First");
        assert_eq!(package.body_footnotes()?[0].text.as_ref(), "First");
        Ok(())
    }

    #[test]
    fn package_rewrites_existing_body_footnote_text_as_an_exact_reversible_edit()
    -> PackageResult<()> {
        let package_bytes = body_footnote_package_bytes()?;
        let package = Package::from_bytes(&package_bytes)?;
        let source_footnotes = package.body_footnotes()?;

        let mut edit = package
            .edit_body_footnote_text(crate::footnote::body::Selector::Index(1))
            .map_err(|error| PackageError::InvalidFormat(error.to_string()))?;
        assert_eq!(edit.position(), Position::from_utf16_index(5).unwrap());
        assert_eq!(edit.before(), &source_footnotes[1]);
        edit.set("Updated😀")
            .map_err(|error| PackageError::InvalidFormat(error.to_string()))?;
        let commit = edit
            .commit()
            .map_err(|error| PackageError::InvalidFormat(error.to_string()))?;
        assert!(commit.diagnostics().changed());
        assert_eq!(commit.diagnostics().touched_components(), 1);
        assert!(commit.diagnostics().full_reparse_performed());
        assert_eq!(commit.package().body_footnotes()?[0], source_footnotes[0]);
        assert_eq!(
            commit.package().body_footnotes()?[1],
            Footnote {
                position: Position::from_utf16_index(5).unwrap(),
                text: "Updated😀".into(),
                custom_mark: Some("†".into()),
            }
        );

        let inverse = commit.patch().inverse();
        let restored = commit
            .package()
            .apply_body_footnote_text(&inverse)
            .map_err(|error| PackageError::InvalidFormat(error.to_string()))?;
        assert_eq!(restored.package().source_bytes(), package_bytes.as_slice());
        assert_eq!(restored.package().body_footnotes()?, source_footnotes);
        assert!(!format!("{:?}", commit.patch()).contains("Index/Document.iwa"));
        Ok(())
    }

    #[test]
    fn package_rewrites_existing_body_footnote_marker_without_touching_graph_or_unknown_bytes()
    -> PackageResult<()> {
        let package_bytes = body_footnote_package_bytes()?;
        let package = Package::from_bytes(&package_bytes)?;
        let source_footnotes = package.body_footnotes()?;
        let unknown_reference_bytes = [0xa0, 0x06, 0x01, 0xaa, 0x06, 0x03, b'r', b'e', b'f'];
        let source_reference_payload = package
            .state
            .source
            .components()
            .iter()
            .find_map(|component| component.archive().object(110))
            .and_then(|object| {
                object
                    .messages
                    .iter()
                    .find(|message| message.type_ == FOOTNOTE_REFERENCE_MESSAGE_TYPE)
                    .map(|message| message.data.as_slice())
            })
            .ok_or_else(|| PackageError::InvalidFormat("test reference is missing".to_owned()))?;
        assert!(
            source_reference_payload
                .windows(unknown_reference_bytes.len())
                .any(|window| { window == unknown_reference_bytes })
        );

        let mut edit = package
            .edit_body_footnote_text(crate::footnote::body::Selector::Index(1))
            .map_err(|error| PackageError::InvalidFormat(error.to_string()))?;
        assert_eq!(edit.custom_mark(), Some("†"));
        edit.set("UpdatedSecond")
            .map_err(|error| PackageError::InvalidFormat(error.to_string()))?;
        edit.set_custom_mark(Some("‡"))
            .map_err(|error| PackageError::InvalidFormat(error.to_string()))?;
        let commit = edit
            .commit()
            .map_err(|error| PackageError::InvalidFormat(error.to_string()))?;
        assert!(commit.diagnostics().changed());
        assert_eq!(commit.diagnostics().touched_components(), 1);
        assert_eq!(commit.package().body_footnotes()?[0], source_footnotes[0]);
        assert_eq!(
            commit.package().body_footnotes()?[1],
            Footnote {
                position: Position::from_utf16_index(5).unwrap(),
                text: "UpdatedSecond".into(),
                custom_mark: Some("‡".into()),
            }
        );
        assert!(
            commit
                .package()
                .state
                .source
                .components()
                .iter()
                .find_map(|component| component.archive().object(110))
                .and_then(|object| {
                    object
                        .messages
                        .iter()
                        .find(|message| message.type_ == FOOTNOTE_REFERENCE_MESSAGE_TYPE)
                        .map(|message| message.data.as_slice())
                })
                .is_some_and(|payload| {
                    payload
                        .windows(unknown_reference_bytes.len())
                        .any(|window| window == unknown_reference_bytes)
                })
        );

        let inverse = commit.patch().inverse();
        let restored = commit
            .package()
            .apply_body_footnote_text(&inverse)
            .map_err(|error| PackageError::InvalidFormat(error.to_string()))?;
        assert_eq!(restored.package().source_bytes(), package_bytes.as_slice());
        assert_eq!(restored.package().body_footnotes()?, source_footnotes);
        Ok(())
    }

    #[test]
    fn package_empty_root_has_no_synthetic_section() -> PackageResult<()> {
        let package = Package::from_bytes(&package_bytes(None, false, false)?)?;

        assert!(package.sections().is_empty());
        assert!(package.semantic_document().is_empty());
        assert_eq!(package.text()?, "");
        assert_eq!(package.stats().section_count(), 0);
        Ok(())
    }

    #[test]
    fn package_preserves_rooted_text_fragments_as_semantic_runs() -> PackageResult<()> {
        let root = tp::DocumentArchive {
            body_storage: Some(Reference {
                identifier: 42,
                ..Reference::default()
            }),
            ..tp::DocumentArchive::default()
        };
        let storage = tswp::StorageArchive {
            text: vec!["first".to_owned(), " second".to_owned()],
            ..tswp::StorageArchive::default()
        };
        let bytes = archive_package_bytes(
            vec![
                ArchiveObject::new(
                    1,
                    vec![RawMessage {
                        type_: 10_000,
                        data: root.encode_to_vec(),
                    }],
                )
                .map_err(|error| PackageError::InvalidFormat(error.to_string()))?,
                ArchiveObject::new(
                    42,
                    vec![RawMessage {
                        type_: 2_001,
                        data: storage.encode_to_vec(),
                    }],
                )
                .map_err(|error| PackageError::InvalidFormat(error.to_string()))?,
            ],
            false,
        )?;

        let package = Package::from_bytes(&bytes)?;
        assert_eq!(package.sections().len(), 1);
        let storages = package.sections()[0].text_storages();
        assert_eq!(storages.len(), 1);
        assert_eq!(storages[0].text(), "first second");
        assert_eq!(storages[0].runs(), [Run::new(0, 5), Run::new(5, 7)]);
        Ok(())
    }

    #[test]
    fn package_root_projection_requires_unique_base_envelope() -> PackageResult<()> {
        let missing_super = archive_package_bytes(
            vec![
                ArchiveObject::new(
                    1,
                    vec![RawMessage {
                        type_: 10_000,
                        data: vec![0x22, 0x02, 0x08, 0x2a],
                    }],
                )
                .map_err(|error| PackageError::InvalidFormat(error.to_string()))?,
            ],
            false,
        )?;
        let missing_error = Package::from_bytes(&missing_super)
            .err()
            .unwrap_or_else(|| panic!("a Pages root without its base envelope must fail"));
        assert!(
            missing_error
                .to_string()
                .contains("TP.DocumentArchive.super")
        );

        let duplicate_body = archive_package_bytes(
            vec![
                ArchiveObject::new(
                    1,
                    vec![RawMessage {
                        type_: 10_000,
                        data: vec![0x22, 0x02, 0x08, 0x2a, 0x22, 0x02, 0x08, 0x2b, 0x7a, 0x00],
                    }],
                )
                .map_err(|error| PackageError::InvalidFormat(error.to_string()))?,
            ],
            false,
        )?;
        let duplicate_error = Package::from_bytes(&duplicate_body)
            .err()
            .unwrap_or_else(|| panic!("duplicate Pages body references must fail"));
        assert!(
            duplicate_error
                .to_string()
                .contains("TP.DocumentArchive.body_storage")
        );
        Ok(())
    }

    #[test]
    fn section_boundary_projection_rejects_ambiguous_references() -> PackageResult<()> {
        let duplicate_reference = [0x08, 0x00, 0x12, 0x02, 0x08, 0x2a, 0x12, 0x02, 0x08, 0x2b];
        let error = preflight_section_table_entry(
            &duplicate_reference,
            0,
            NonZeroU64::MIN,
            pages_body_options(Limits::default())?,
        )
        .err()
        .unwrap_or_else(|| panic!("duplicate section references must fail strict projection"));
        assert!(error.to_string().contains("ObjectAttribute.object"));
        Ok(())
    }

    #[cfg(feature = "internal-iwork-source")]
    #[test]
    fn prepared_source_handoff_matches_direct_ingress() -> PackageResult<()> {
        let bytes = package_bytes(Some("Prepared Pages body"), true, true)?;
        let direct = Package::from_bytes(&bytes)?;
        let prepared = PreparedSource::from_bytes(&bytes)
            .map_err(|error| PackageError::InvalidFormat(error.to_string()))?
            .ok_or_else(|| {
                PackageError::InvalidFormat("Pages fixture was not detected".to_owned())
            })?;

        let handed_off = Package::__from_prepared_source(prepared)?;

        assert_eq!(handed_off.text()?, direct.text()?);
        assert_eq!(handed_off.sections().len(), direct.sections().len());
        assert_eq!(
            handed_off.sections()[0].plain_text(),
            direct.sections()[0].plain_text()
        );
        assert_eq!(handed_off.metadata().title, direct.metadata().title);
        Ok(())
    }

    #[cfg(feature = "internal-iwork-source")]
    #[test]
    fn semantic_prepared_source_handoff_matches_package_projection() -> PackageResult<()> {
        let bytes = sectioned_package_bytes(
            vec!["A\u{0004}B".to_owned()],
            &[(0, 43), (2, 44)],
            None,
            vec![
                (43, vec![section_payload(Some("One"))]),
                (44, vec![section_payload(Some("Two"))]),
            ],
        )?;
        let direct = Package::from_bytes(&bytes)?;
        let prepared = PreparedSource::from_bytes(&bytes)
            .map_err(|error| PackageError::InvalidFormat(error.to_string()))?
            .ok_or_else(|| {
                PackageError::InvalidFormat("Pages fixture was not detected".to_owned())
            })?;

        let document = __semantic_document_from_prepared_source(
            prepared,
            MAX_SECTIONS,
            DEFAULT_MAX_TEXT_BYTES,
        )?;

        assert_eq!(document.plain_text(), direct.text()?);
        assert_eq!(document.sections().len(), direct.sections().len());
        for (semantic, packaged) in document.sections().iter().zip(direct.sections()) {
            assert_eq!(semantic.index(), packaged.index());
            assert_eq!(semantic.name(), packaged.name());
            assert_eq!(semantic.plain_text(), packaged.plain_text());
        }
        Ok(())
    }

    #[cfg(feature = "internal-iwork-source")]
    #[test]
    fn semantic_prepared_source_releases_physical_source() -> PackageResult<()> {
        let source: Arc<[u8]> = package_bytes(Some("Detached"), true, true)?.into();
        let weak_source = Arc::downgrade(&source);
        let prepared = PreparedSource::from_shared_bytes(Arc::clone(&source))
            .map_err(|error| PackageError::InvalidFormat(error.to_string()))?
            .ok_or_else(|| {
                PackageError::InvalidFormat("Pages fixture was not detected".to_owned())
            })?;
        drop(source);
        assert!(weak_source.upgrade().is_some());

        let document = __semantic_document_from_prepared_source(
            prepared,
            MAX_SECTIONS,
            DEFAULT_MAX_TEXT_BYTES,
        )?;

        assert!(weak_source.upgrade().is_none());
        assert_eq!(document.plain_text(), "Detached");
        Ok(())
    }

    #[cfg(feature = "internal-iwork-source")]
    #[test]
    fn semantic_prepared_source_rejects_malformed_pages_graph() -> PackageResult<()> {
        let root = tp::DocumentArchive {
            body_storage: Some(Reference {
                identifier: 42,
                ..Reference::default()
            }),
            ..tp::DocumentArchive::default()
        };
        let bytes = archive_package_bytes(
            vec![
                ArchiveObject::new(
                    1,
                    vec![RawMessage {
                        type_: 10_000,
                        data: root.encode_to_vec(),
                    }],
                )
                .map_err(|error| PackageError::InvalidFormat(error.to_string()))?,
            ],
            false,
        )?;
        let prepared = PreparedSource::from_bytes(&bytes)
            .map_err(|error| PackageError::InvalidFormat(error.to_string()))?
            .ok_or_else(|| {
                PackageError::InvalidFormat("Pages fixture was not detected".to_owned())
            })?;

        let error = __semantic_document_from_prepared_source(
            prepared,
            MAX_SECTIONS,
            DEFAULT_MAX_TEXT_BYTES,
        )
        .err()
        .unwrap_or_else(|| panic!("missing Pages body should fail"));

        assert!(
            error
                .to_string()
                .contains("body storage object 42 is missing")
        );
        Ok(())
    }

    #[cfg(feature = "internal-iwork-source")]
    #[test]
    fn semantic_prepared_source_enforces_exact_and_one_over_text_limits() -> PackageResult<()> {
        let bytes = package_bytes(Some("12345"), true, false)?;
        let exact = PreparedSource::from_bytes(&bytes)
            .map_err(|error| PackageError::InvalidFormat(error.to_string()))?
            .ok_or_else(|| {
                PackageError::InvalidFormat("Pages fixture was not detected".to_owned())
            })?;
        let document = __semantic_document_from_prepared_source(exact, MAX_SECTIONS, 5)?;
        assert_eq!(document.plain_text(), "12345");

        let one_over = PreparedSource::from_bytes(&bytes)
            .map_err(|error| PackageError::InvalidFormat(error.to_string()))?
            .ok_or_else(|| {
                PackageError::InvalidFormat("Pages fixture was not detected".to_owned())
            })?;
        let error = __semantic_document_from_prepared_source(one_over, MAX_SECTIONS, 4)
            .err()
            .unwrap_or_else(|| panic!("one-over text should fail during wire preflight"));
        assert!(matches!(
            error,
            PackageError::Semantic(SemanticError::TextTooLarge {
                observed: 5,
                limit: 4,
            })
        ));
        Ok(())
    }

    #[cfg(feature = "internal-iwork-source")]
    #[test]
    fn semantic_prepared_source_charges_section_name_bytes_not_scratch_slots() -> PackageResult<()>
    {
        let bytes = sectioned_package_bytes(
            vec![String::new()],
            &[(0, 43)],
            None,
            vec![(43, vec![section_payload(Some("X"))])],
        )?;
        let exact = PreparedSource::from_bytes(&bytes)
            .map_err(|error| PackageError::InvalidFormat(error.to_string()))?
            .ok_or_else(|| {
                PackageError::InvalidFormat("Pages fixture was not detected".to_owned())
            })?;

        let document = __semantic_document_from_prepared_source(exact, MAX_SECTIONS, "X".len())?;
        assert_eq!(document.sections().len(), 1);
        assert_eq!(document.sections()[0].name(), Some("X"));
        assert_eq!(document.text_len(), 0);
        Ok(())
    }

    #[cfg(feature = "internal-iwork-source")]
    #[test]
    fn semantic_prepared_source_enforces_exact_and_one_over_section_limits() -> PackageResult<()> {
        let bytes = sectioned_package_bytes(
            vec!["A\u{0004}B".to_owned()],
            &[(0, 43), (2, 44)],
            None,
            vec![
                (43, vec![section_payload(Some("One"))]),
                (44, vec![section_payload(Some("Two"))]),
            ],
        )?;
        let exact = PreparedSource::from_bytes(&bytes)
            .map_err(|error| PackageError::InvalidFormat(error.to_string()))?
            .ok_or_else(|| {
                PackageError::InvalidFormat("Pages fixture was not detected".to_owned())
            })?;
        let document = __semantic_document_from_prepared_source(exact, 2, DEFAULT_MAX_TEXT_BYTES)?;
        assert_eq!(document.sections().len(), 2);

        let one_over = PreparedSource::from_bytes(&bytes)
            .map_err(|error| PackageError::InvalidFormat(error.to_string()))?
            .ok_or_else(|| {
                PackageError::InvalidFormat("Pages fixture was not detected".to_owned())
            })?;
        let error = __semantic_document_from_prepared_source(one_over, 1, DEFAULT_MAX_TEXT_BYTES)
            .err()
            .unwrap_or_else(|| panic!("one-over section should fail during wire preflight"));
        assert!(matches!(
            error,
            PackageError::Semantic(SemanticError::TooManySections {
                actual: 2,
                limit: 1
            })
        ));
        Ok(())
    }

    #[test]
    fn package_without_body_uses_native_storage_fallback() -> PackageResult<()> {
        let package = Package::from_bytes(&package_bytes(Some("Fallback text"), false, false)?)?;
        assert_eq!(package.text()?, "Fallback text");
        assert_eq!(package.sections().len(), 1);
        Ok(())
    }

    #[test]
    fn native_section_names_and_utf16_boundaries_are_projected() -> PackageResult<()> {
        let package = Package::from_bytes(&sectioned_package_bytes(
            vec!["A🚀\u{0004}B".to_owned()],
            &[(0, 43), (4, 44)],
            None,
            vec![
                (43, vec![section_payload(Some("Introduction"))]),
                (44, vec![section_payload(Some("Appendix"))]),
            ],
        )?)?;

        assert_eq!(package.text()?, "A🚀\nB");
        assert_eq!(package.stats().section_count(), 2);
        assert_eq!(package.sections()[0].name(), Some("Introduction"));
        assert_eq!(package.sections()[0].plain_text(), "A🚀");
        assert_eq!(
            package
                .section_named("Introduction")
                .unwrap_or_else(|error| panic!("unique section name should resolve: {error}"))
                .map(Section::index),
            Some(0)
        );
        assert_eq!(
            package.sections()[0].text_storages()[0].runs(),
            [Run::new(0, 5)]
        );
        assert_eq!(package.sections()[1].name(), Some("Appendix"));
        assert_eq!(package.sections()[1].plain_text(), "B");
        assert_eq!(
            package
                .section_at(1)
                .unwrap_or_else(|error| panic!("checked native position should resolve: {error}"))
                .map(Section::index),
            Some(1)
        );
        Ok(())
    }

    #[test]
    fn root_initial_section_and_empty_name_presence_are_preserved() -> PackageResult<()> {
        let package = Package::from_bytes(&sectioned_package_bytes(
            vec!["Body".to_owned()],
            &[],
            Some(43),
            vec![(43, vec![section_payload(Some(""))])],
        )?)?;

        assert_eq!(package.sections().len(), 1);
        assert_eq!(package.sections()[0].name(), Some(""));
        assert_eq!(package.sections()[0].plain_text(), "Body");
        Ok(())
    }

    #[test]
    fn duplicate_native_names_remain_a_typed_selector_ambiguity() -> PackageResult<()> {
        let package = Package::from_bytes(&sectioned_package_bytes(
            vec!["One\u{0004}Two".to_owned()],
            &[(0, 43), (4, 44)],
            None,
            vec![
                (43, vec![section_payload(Some("Repeated"))]),
                (44, vec![section_payload(Some("Repeated"))]),
            ],
        )?)?;

        assert_eq!(
            package.section_named("Repeated").err(),
            Some(crate::SelectorError::AmbiguousSectionName {
                name: "Repeated".into(),
                first: 0,
                duplicate: 1,
            })
        );
        assert_eq!(
            package
                .section_at(1)
                .unwrap_or_else(|error| panic!("typed native position should resolve: {error}"))
                .and_then(Section::name),
            Some("Repeated")
        );
        Ok(())
    }

    #[test]
    fn native_section_graph_rejects_duplicate_boundaries_and_missing_breaks() {
        let duplicate_boundary = sectioned_package_bytes(
            vec!["Body".to_owned()],
            &[(0, 43), (0, 44)],
            None,
            vec![
                (43, vec![section_payload(Some("One"))]),
                (44, vec![section_payload(Some("Two"))]),
            ],
        )
        .and_then(|bytes| Package::from_bytes(&bytes))
        .err()
        .unwrap_or_else(|| panic!("duplicate boundaries should fail"));
        assert!(
            duplicate_boundary
                .to_string()
                .contains("duplicate or unsorted"),
            "unexpected duplicate-boundary error: {duplicate_boundary}"
        );

        let missing_break = sectioned_package_bytes(
            vec!["A🚀B".to_owned()],
            &[(0, 43), (3, 44)],
            None,
            vec![
                (43, vec![section_payload(Some("One"))]),
                (44, vec![section_payload(Some("Two"))]),
            ],
        )
        .and_then(|bytes| Package::from_bytes(&bytes))
        .err()
        .unwrap_or_else(|| panic!("missing native section break should fail"));
        assert!(missing_break.to_string().contains("section-break marker"));
    }

    #[test]
    fn referenced_section_requires_one_exact_typed_payload() -> PackageResult<()> {
        let cases = [
            (
                vec![RawMessage {
                    type_: SECTION_MESSAGE_TYPE + 1,
                    data: tp::SectionArchive::default().encode_to_vec(),
                }],
                "has no type-10011 payload",
            ),
            (
                vec![section_payload(Some("One")), section_payload(Some("Two"))],
                "duplicate type-10011 payloads",
            ),
        ];

        for (messages, expected) in cases {
            let error = Package::from_bytes(&sectioned_package_bytes(
                vec!["Body".to_owned()],
                &[(0, 43)],
                None,
                vec![(43, messages)],
            )?)
            .err()
            .unwrap_or_else(|| panic!("invalid section payload should fail"));
            assert!(error.to_string().contains(expected), "{error}");
        }
        Ok(())
    }

    #[test]
    fn duplicate_section_name_wire_field_is_rejected_before_publication() -> PackageResult<()> {
        let mut message = section_payload(Some("First"));
        message
            .data
            .extend(litchi_iwa_common::varint::encode_varint((26 << 3) | 2));
        message
            .data
            .extend(litchi_iwa_common::varint::encode_varint(4));
        message.data.extend_from_slice(b"Last");

        let error = Package::from_bytes(&sectioned_package_bytes(
            vec!["Body".to_owned()],
            &[(0, 43)],
            None,
            vec![(43, vec![message])],
        )?)
        .err()
        .unwrap_or_else(|| panic!("duplicate section name field should fail"));
        assert!(error.to_string().contains("duplicate protobuf field 26"));
        Ok(())
    }

    #[test]
    fn section_name_budget_is_charged_before_owned_materialization() -> PackageResult<()> {
        let object = ArchiveObject::new(43, vec![section_payload(Some("oversized"))])
            .map_err(|error| PackageError::InvalidFormat(error.to_string()))?;
        let error = decode_section_name(&object, NonZeroU64::MIN, 2, 4)
            .err()
            .unwrap_or_else(|| panic!("an over-budget section name should fail"));

        assert!(matches!(
            error,
            PackageError::SectionNamesTooLarge {
                observed: 11,
                limit: 4,
            }
        ));
        Ok(())
    }

    #[test]
    fn input_limit_is_enforced_before_native_projection() {
        let limits = Limits::new(1, 10, 100, 100, 100)
            .unwrap_or_else(|error| panic!("test limits should be valid: {error}"));
        let error = Package::from_bytes_with_limits(&[0, 1], limits)
            .err()
            .unwrap_or_else(|| panic!("oversized input should fail"));
        assert!(error.to_string().contains("limit"));
    }

    #[test]
    fn aggregate_object_count_accepts_exact_cap_and_refuses_one_over() {
        assert_eq!(
            checked_object_count([MAX_OBJECTS])
                .unwrap_or_else(|error| panic!("exact object cap must pass: {error}")),
            MAX_OBJECTS
        );
        assert!(matches!(
            checked_object_count([MAX_OBJECTS - 1, 2]),
            Err(PackageError::ObjectLimit {
                observed,
                limit: MAX_OBJECTS,
            }) if observed == MAX_OBJECTS + 1
        ));
        assert!(matches!(
            checked_object_count([usize::MAX, 1]),
            Err(PackageError::ObjectLimit {
                observed: usize::MAX,
                limit: MAX_OBJECTS,
            })
        ));
    }

    #[test]
    fn object_inventory_rejects_duplicate_identifiers_across_components() -> PackageResult<()> {
        let root = Archive {
            objects: vec![
                ArchiveObject::new(
                    1,
                    vec![RawMessage {
                        type_: 10_000,
                        data: tp::DocumentArchive::default().encode_to_vec(),
                    }],
                )
                .map_err(|error| PackageError::InvalidFormat(error.to_string()))?,
            ],
        };
        let other = Archive {
            objects: vec![
                ArchiveObject::new(
                    1,
                    vec![RawMessage {
                        type_: 999,
                        data: Vec::new(),
                    }],
                )
                .map_err(|error| PackageError::InvalidFormat(error.to_string()))?,
            ],
        };
        let root_stream = SnappyStream::compress(
            &root
                .to_bytes()
                .map_err(|error| PackageError::InvalidFormat(error.to_string()))?,
        )
        .map_err(|error| PackageError::InvalidFormat(error.to_string()))?;
        let other_stream = SnappyStream::compress(
            &other
                .to_bytes()
                .map_err(|error| PackageError::InvalidFormat(error.to_string()))?,
        )
        .map_err(|error| PackageError::InvalidFormat(error.to_string()))?;
        let bytes = litchi_iwa_archive::package::to_bytes(
            [
                ("Index/Document.iwa", root_stream.as_slice()),
                ("Index/Other.iwa", other_stream.as_slice()),
            ],
            Limits::default(),
        )?;

        let error = Package::from_bytes(&bytes)
            .err()
            .unwrap_or_else(|| panic!("duplicate cross-component identifiers must fail"));
        assert!(matches!(error, PackageError::InvalidFormat(_)));
        Ok(())
    }

    #[test]
    fn fallback_speculation_skips_only_content_invalid_payloads() {
        assert!(
            skip_or_reject_fallback_validation(litchi_iwa_text_wire::RewriteError::InvalidFormat(
                "speculative".to_owned()
            ))
            .is_ok()
        );
        assert!(matches!(
            skip_or_reject_fallback_validation(litchi_iwa_text_wire::RewriteError::Allocation {
                resource: "test",
                amount: 7,
            }),
            Err(PackageError::Allocation { amount: 7 })
        ));
        assert!(matches!(
            skip_or_reject_fallback_validation(litchi_iwa_text_wire::RewriteError::Projection(
                "parity".to_owned()
            )),
            Err(PackageError::InvalidFormat(_))
        ));

        assert!(
            skip_or_reject_fallback_materialization(
                litchi_iwa_text_wire::Error::WrongTextWireType { actual: 0 },
                0,
                16,
            )
            .is_ok()
        );
        assert!(matches!(
            skip_or_reject_fallback_materialization(
                litchi_iwa_text_wire::Error::Common(litchi_iwa_common::Error::Allocation {
                    resource: "test",
                    amount: 11,
                }),
                0,
                16,
            ),
            Err(PackageError::Allocation { amount: 11 })
        ));
        assert!(matches!(
            skip_or_reject_fallback_materialization(
                litchi_iwa_text_wire::Error::ProjectionMismatch {
                    preflight: 1,
                    decoded: 0,
                },
                0,
                16,
            ),
            Err(PackageError::InvalidFormat(_))
        ));
    }
}
