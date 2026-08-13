//! Pages-native package ingress and semantic projection.
//!
//! This module is the only Pages boundary that understands ZIP/IWA packages
//! and generated protobuf messages. It publishes [`Package`] snapshots whose
//! semantic content is represented by the archive-free [`crate::Document`].

pub(crate) mod document_settings;
mod page_layout;
pub(crate) mod section_background;
mod section_name;
mod section_pagination;
pub(crate) mod section_settings;
mod section_text;
mod section_transaction;

use std::collections::BTreeSet;
use std::fmt;
use std::fs::{Metadata as FileMetadata, OpenOptions};
use std::io::Read;
use std::num::NonZeroU64;
use std::path::Path;
use std::sync::Arc;

use litchi_iwa_archive::package::Catalog;
use litchi_iwa_archive::{ComponentCatalog, SourceCatalog};
use litchi_iwa_common::{
    WireLimits,
    wire::{WireFieldView, WireView},
};
#[cfg(feature = "internal-iwork-source")]
use litchi_iwa_detect::{Format, PreparedSource};
use litchi_iwa_protos::{
    pages_body_codec::{self, DecodeOptions as PagesBodyDecodeOptions},
    tswp,
};
use litchi_iwa_text::storage::{Run, Storage};
use plist::Value;
use prost::Message;
use thiserror::Error;

use crate::{
    Body, DEFAULT_MAX_TEXT_BYTES, Document, Error as SemanticError, MAX_BODY_STORAGES,
    MAX_SECTIONS, Root, Section, SectionType,
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

const SECTION_MESSAGE_TYPE: u32 = 10_011;

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
    /// The parsed package cannot form a valid Pages-native document.
    #[error("invalid Pages package: {0}")]
    InvalidFormat(String),
    /// The native package decoded successfully but exceeded a Pages semantic
    /// bound while being projected into an immutable document.
    #[error(transparent)]
    Semantic(#[from] SemanticError),
    /// Section names exceed the aggregate retained-text budget.
    #[error("Pages section names require at least {observed} bytes; budget is {limit}")]
    SectionNamesTooLarge {
        /// Minimum bytes required by the names decoded so far.
        observed: usize,
        /// Aggregate retained UTF-8 budget.
        limit: usize,
    },
}

/// Result type returned by [`Package`] operations.
pub type PackageResult<T> = Result<T, PackageError>;

/// Immutable statistics captured while a Pages package is opened.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Stats {
    total_objects: usize,
    section_count: usize,
}

impl Stats {
    /// Return the number of native IWA objects retained by the package.
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

    /// Parse a Pages package from archive bytes.
    ///
    /// This is an explicit alias for [`Self::from_bytes`] for callers whose
    /// input is already known to be a ZIP archive.
    ///
    /// # Errors
    ///
    /// Returns the same errors as [`Self::from_bytes`].
    pub fn from_archive_bytes(bytes: &[u8]) -> PackageResult<Self> {
        Self::from_bytes(bytes)
    }

    /// Capture another cheap handle to the same native and semantic snapshot.
    #[must_use]
    pub fn snapshot(&self) -> Self {
        self.clone()
    }

    /// Borrow the authoritative immutable package bytes.
    ///
    /// The returned bytes are the exact artifact represented by this
    /// snapshot. They can be written directly to publish a committed edit;
    /// callers never need access to native object identifiers or package
    /// members.
    #[must_use]
    pub fn source_bytes(&self) -> &[u8] {
        self.state.source.source_bytes()
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

    /// Borrow the immutable Pages semantic snapshot.
    #[must_use]
    pub fn semantic_document(&self) -> &Document {
        &self.state.document
    }

    /// Project native Pages metadata into the format-neutral core model.
    #[must_use]
    pub fn metadata(&self) -> litchi_core::Metadata {
        let metadata = &self.state.metadata;
        let revision = metadata
            .revision
            .clone()
            .or_else(|| metadata.build_version.clone());
        let content_status = metadata
            .format_version
            .as_deref()
            .map(|version| format!("Pages Format Version {version}"));

        litchi_core::Metadata {
            title: metadata.title.clone(),
            author: metadata.author.clone(),
            keywords: metadata.keywords.clone(),
            description: metadata.description.clone(),
            application: Some(
                metadata
                    .application
                    .clone()
                    .unwrap_or_else(|| "Pages".to_owned()),
            ),
            revision,
            content_status,
            identifier: metadata.identifier.clone(),
            ..Default::default()
        }
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
            metadata.identifier = Some(identifier.to_owned());
        }

        Ok(metadata)
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
    let (components, limits) = semantic_components_from_prepared_source(source)?;
    decode_semantic_components(
        &components,
        max_sections.min(MAX_SECTIONS),
        max_text_bytes.min(effective_text_limit(limits)),
        limits,
    )
}

#[cfg(feature = "internal-iwork-source")]
fn validate_prepared_format(source: &PreparedSource) -> PackageResult<()> {
    if source.format() != Format::Pages {
        return Err(PackageError::InvalidFormat(
            "prepared iWork source is not a Pages document".to_owned(),
        ));
    }
    Ok(())
}

#[cfg(feature = "internal-iwork-source")]
fn semantic_components_from_prepared_source(
    source: PreparedSource,
) -> PackageResult<(Arc<ComponentCatalog>, Limits)> {
    validate_prepared_format(&source)?;
    Ok(source.__into_components())
}

#[cfg(feature = "internal-iwork-source")]
fn decode_semantic_components(
    components: &ComponentCatalog,
    max_sections: usize,
    max_text_bytes: usize,
    limits: Limits,
) -> PackageResult<Document> {
    validate_components(components)?;
    let root_references = root_references_with_limits(components, limits)?;
    decode_document(
        components,
        root_references,
        max_sections,
        max_text_bytes,
        limits,
    )
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

    let mut object_ids = BTreeSet::new();
    let mut object_count = 0usize;
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
            if !object_ids.insert(identifier) {
                return Err(PackageError::InvalidFormat(format!(
                    "Pages package contains duplicate object identifier {identifier}"
                )));
            }
            object_count = object_count.checked_add(1).ok_or_else(|| {
                PackageError::InvalidFormat("Pages package object count overflows usize".to_owned())
            })?;
        }
    }

    if object_count == 0 {
        return Err(PackageError::InvalidFormat(
            "Pages package contains no IWA objects".to_owned(),
        ));
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
        let (native, payload) = decode_body_storage(
            &object.messages,
            identifier,
            max_sections,
            max_text_bytes,
            limits,
        )?;
        validate_section_table_wire_with_limits(payload, &native, identifier, limits)?;
        let section_references =
            native_section_references(&native, root_references.initial_section, max_sections)?;
        if section_references.is_empty() && max_sections == 0 {
            return Err(SemanticError::TooManySections {
                actual: 1,
                limit: max_sections,
            }
            .into());
        }
        return project_native_body(
            components,
            native,
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
        let storages = extract_storages(components, max_sections, max_text_bytes)?;
        (!storages.is_empty())
            .then(|| Body::with_max_text_bytes(storages, max_text_bytes))
            .transpose()?
    };
    let root = body.map_or_else(Root::empty, Root::with_body);
    Document::from_root_with_max_text_bytes(root, max_text_bytes).map_err(Into::into)
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
) -> PackageResult<(tswp::StorageArchive, &[u8])> {
    let payload = unique_text_payload(messages, identifier)?;
    let wire_limits = storage_rewrite_limits(limits).map_err(|limit_error| match limit_error {
        StorageWireLimitsError::Physical(physical_error) => PackageError::Archive(physical_error),
        StorageWireLimitsError::Wire(wire_error) => PackageError::InvalidFormat(format!(
            "Pages body object {identifier} text validation limits are invalid: {wire_error}"
        )),
    })?;
    litchi_iwa_text_wire::validate_storage_with_limits(payload, wire_limits).map_err(|error| {
        PackageError::InvalidFormat(format!(
            "Pages body object {identifier} text payload failed bounded validation: {error}"
        ))
    })?;
    preflight_body_wire(payload, identifier, max_sections, max_text_bytes, limits)?;
    let native = tswp::StorageArchive::decode(payload).map_err(|error| {
        PackageError::InvalidFormat(format!(
            "Pages body object {identifier} text payload is invalid: {error}"
        ))
    })?;
    ensure_text_size(&native, 0, max_text_bytes, identifier)?;
    Ok((native, payload))
}

fn preflight_body_wire(
    payload: &[u8],
    body_identifier: NonZeroU64,
    max_sections: usize,
    max_text_bytes: usize,
    limits: Limits,
) -> PackageResult<()> {
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
        return Ok(());
    };
    let table_view = parse_wire(wire_table_field.payload(), &context)?;
    let mut entry_count = 0usize;
    let boundary_options = pages_body_options(limits)?;
    for field in table_view.fields().filter(|field| field.number() == 1) {
        validate_wire_field(field, 2, &context)?;
        entry_count = entry_count.checked_add(1).ok_or_else(|| {
            PackageError::InvalidFormat(format!("{context} section count overflows usize"))
        })?;
        if entry_count > max_sections {
            return Err(SemanticError::TooManySections {
                actual: entry_count,
                limit: max_sections,
            }
            .into());
        }
        preflight_section_table_entry(
            field.payload(),
            entry_count - 1,
            body_identifier,
            boundary_options,
        )?;
    }
    Ok(())
}

fn preflight_section_table_entry(
    payload: &[u8],
    entry_index: usize,
    body_identifier: NonZeroU64,
    options: PagesBodyDecodeOptions,
) -> PackageResult<()> {
    let context = format!("Pages body object {body_identifier} section table entry {entry_index}");
    let boundary = pages_body_codec::decode_section_boundary(payload, options)
        .map_err(|error| PackageError::InvalidFormat(format!("{context} is invalid: {error}")))?;
    boundary.section().ok_or_else(|| {
        PackageError::InvalidFormat(format!("{context} has no section reference"))
    })?;
    Ok(())
}

fn validate_section_table_wire_with_limits(
    payload: &[u8],
    decoded: &tswp::StorageArchive,
    body_identifier: NonZeroU64,
    limits: Limits,
) -> PackageResult<()> {
    let context = format!("Pages body object {body_identifier} section table");
    let view = parse_wire(payload, &context)?;
    let optional_field = unique_wire_field(&view, 17, 2, false, &context)?;
    let optional_decoded_table = decoded.table_section.as_ref();
    let (Some(wire_field), Some(native_table)) = (optional_field, optional_decoded_table) else {
        if optional_field.is_none() && optional_decoded_table.is_none() {
            return Ok(());
        }
        return Err(PackageError::InvalidFormat(format!(
            "{context} presence changed while decoding"
        )));
    };

    let table_view = parse_wire(wire_field.payload(), &context)?;
    let mut decoded_entries = native_table.entries.iter();
    let mut entry_index = 0usize;
    let boundary_options = pages_body_options(limits)?;
    for entry_field in table_view
        .fields()
        .filter(|candidate| candidate.number() == 1)
    {
        validate_wire_field(entry_field, 2, &context)?;
        let decoded_entry = decoded_entries.next().ok_or_else(|| {
            PackageError::InvalidFormat(format!("{context} gained an entry while decoding"))
        })?;
        validate_section_table_entry_wire(
            entry_field.payload(),
            decoded_entry,
            entry_index,
            body_identifier,
            boundary_options,
        )?;
        entry_index = entry_index.checked_add(1).ok_or_else(|| {
            PackageError::InvalidFormat("Pages section entry count overflows usize".to_owned())
        })?;
    }
    if decoded_entries.next().is_some() {
        return Err(PackageError::InvalidFormat(format!(
            "{context} lost an entry while decoding"
        )));
    }
    Ok(())
}

fn validate_section_table_entry_wire(
    payload: &[u8],
    decoded: &tswp::object_attribute_table::ObjectAttribute,
    entry_index: usize,
    body_identifier: NonZeroU64,
    options: PagesBodyDecodeOptions,
) -> PackageResult<()> {
    let context = format!("Pages body object {body_identifier} section table entry {entry_index}");
    let boundary = pages_body_codec::decode_section_boundary(payload, options)
        .map_err(|error| PackageError::InvalidFormat(format!("{context} is invalid: {error}")))?;
    if boundary.character_index() != decoded.character_index {
        return Err(PackageError::InvalidFormat(format!(
            "{context} character index changed while decoding"
        )));
    }

    match (boundary.section(), decoded.object.as_ref()) {
        (Some(reference), Some(decoded_reference)) => {
            if reference.identifier().get() != decoded_reference.identifier {
                return Err(PackageError::InvalidFormat(format!(
                    "{context} section reference changed while decoding"
                )));
            }
        },
        (None, None) => {},
        _ => {
            return Err(PackageError::InvalidFormat(format!(
                "{context} section-reference presence changed while decoding"
            )));
        },
    }
    Ok(())
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
    body: &tswp::StorageArchive,
    initial_section: Option<NonZeroU64>,
    max_sections: usize,
) -> PackageResult<Vec<NativeSectionReference>> {
    let table_entries = body
        .table_section
        .as_ref()
        .map_or(&[][..], |table| table.entries.as_slice());
    let maximum_count = table_entries
        .len()
        .checked_add(usize::from(initial_section.is_some()))
        .ok_or_else(|| {
            PackageError::InvalidFormat("Pages section count overflows usize".to_owned())
        })?;
    if maximum_count > max_sections.saturating_add(1) {
        return Err(SemanticError::TooManySections {
            actual: maximum_count,
            limit: max_sections,
        }
        .into());
    }

    let mut references = Vec::new();
    references
        .try_reserve_exact(maximum_count)
        .map_err(|_error| {
            PackageError::InvalidFormat("could not allocate Pages section references".to_owned())
        })?;
    for (entry_index, entry) in table_entries.iter().enumerate() {
        let reference = entry.object.as_ref().ok_or_else(|| {
            PackageError::InvalidFormat(format!(
                "Pages section table entry {entry_index} has no section reference"
            ))
        })?;
        let identifier = NonZeroU64::new(reference.identifier).ok_or_else(|| {
            PackageError::InvalidFormat(format!(
                "Pages section table entry {entry_index} has a zero section reference"
            ))
        })?;
        references.push(NativeSectionReference {
            character_index: entry.character_index,
            identifier,
        });
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
        return Err(SemanticError::TooManySections {
            actual: references.len(),
            limit: max_sections,
        }
        .into());
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
    native: tswp::StorageArchive,
    section_references: Vec<NativeSectionReference>,
    max_text_bytes: usize,
    body_identifier: NonZeroU64,
) -> PackageResult<Document> {
    if section_references.is_empty() {
        let storage = litchi_iwa_text_wire::from_archive(native).map_err(|error| {
            PackageError::InvalidFormat(format!(
                "Pages body object {body_identifier} text payload is invalid: {error}"
            ))
        })?;
        let body = Body::with_max_text_bytes(vec![storage], max_text_bytes)?;
        return Document::from_root_with_max_text_bytes(Root::with_body(body), max_text_bytes)
            .map_err(Into::into);
    }

    let ranges = section_text_ranges(&native.text, &section_references, body_identifier)?;
    let storages = split_native_text(&native.text, &ranges, body_identifier)?;
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

    Document::from_sections_with_max_text_bytes(sections, max_text_bytes).map_err(Into::into)
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
            let name = decode_section_name(object, references[destination].identifier)?;
            retained_bytes = retained_bytes
                .checked_add(name.as_deref().map_or(0, str::len))
                .ok_or(PackageError::SectionNamesTooLarge {
                    observed: usize::MAX,
                    limit: max_name_text_bytes,
                })?;
            if retained_bytes > max_name_text_bytes {
                return Err(PackageError::SectionNamesTooLarge {
                    observed: retained_bytes,
                    limit: max_name_text_bytes,
                });
            }
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
) -> PackageResult<Option<Box<str>>> {
    let context = format!("Pages section object {identifier}");
    let payload = unique_message_payload(&object.messages, SECTION_MESSAGE_TYPE, &context)?;
    let view = parse_wire(payload, &context)?;
    let Some(field) = unique_wire_field(&view, 26, 2, false, &context)? else {
        return Ok(None);
    };
    let name = std::str::from_utf8(field.payload()).map_err(|error| {
        PackageError::InvalidFormat(format!(
            "Pages section object {identifier} name is not valid UTF-8: {error}"
        ))
    })?;
    let mut owned = String::new();
    owned.try_reserve_exact(name.len()).map_err(|_error| {
        PackageError::InvalidFormat(format!(
            "could not allocate Pages section object {identifier} name"
        ))
    })?;
    owned.push_str(name);
    Ok(Some(owned.into_boxed_str()))
}

fn section_text_ranges(
    fragments: &[String],
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

    for fragment in fragments {
        for character in fragment.chars() {
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
                    references[reference_index].identifier,
                    references[reference_index].character_index
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
    fragments: &[String],
    ranges: &[TextRange],
    body_identifier: NonZeroU64,
) -> PackageResult<Vec<Storage>> {
    if fragments.len() > litchi_iwa_text_wire::MAX_FRAGMENTS {
        return Err(PackageError::InvalidFormat(format!(
            "Pages body object {body_identifier} contains {} text fragments; maximum is {}",
            fragments.len(),
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

    let mut global_start = 0usize;
    let mut section_index = 0usize;
    for fragment in fragments {
        let global_end = global_start.checked_add(fragment.len()).ok_or_else(|| {
            PackageError::InvalidFormat(format!(
                "Pages body object {body_identifier} UTF-8 length overflows usize"
            ))
        })?;
        if fragment.is_empty() {
            if let Some(index) = range_containing_empty_offset(ranges, global_start) {
                accumulators[index].push_empty()?;
            }
            global_start = global_end;
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
        global_start = global_end;
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
) -> PackageResult<Vec<Storage>> {
    let mut storages = Vec::new();
    let mut text_bytes = 0usize;

    for component in components.iter() {
        for object in &component.archive().objects {
            for message in &object.messages {
                if !is_storage_message_type(message.type_) {
                    continue;
                }
                let Ok(native) = tswp::StorageArchive::decode(message.data.as_slice()) else {
                    continue;
                };
                let Some(identifier) = object.archive_info.identifier.and_then(NonZeroU64::new)
                else {
                    continue;
                };
                ensure_text_size(&native, text_bytes, max_text_bytes, identifier)?;
                let Ok(storage) = litchi_iwa_text_wire::from_archive(native) else {
                    continue;
                };
                if storage.is_empty() {
                    continue;
                }
                if max_sections == 0 {
                    return Err(SemanticError::TooManySections {
                        actual: 1,
                        limit: max_sections,
                    }
                    .into());
                }
                if storages.len() == MAX_BODY_STORAGES {
                    return Err(SemanticError::TooManyBodyStorages {
                        actual: storages.len().saturating_add(1),
                        limit: MAX_BODY_STORAGES,
                    }
                    .into());
                }
                text_bytes = text_bytes.checked_add(storage.len()).ok_or({
                    PackageError::Semantic(SemanticError::TextTooLarge {
                        observed: usize::MAX,
                        limit: max_text_bytes,
                    })
                })?;
                storages.push(storage);
            }
        }
    }
    Ok(storages)
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

fn ensure_text_size(
    storage: &tswp::StorageArchive,
    initial: usize,
    max_text_bytes: usize,
    identifier: NonZeroU64,
) -> PackageResult<()> {
    let text_len = storage.text.iter().try_fold(initial, |length, fragment| {
        length.checked_add(fragment.len()).ok_or_else(|| {
            PackageError::InvalidFormat(format!(
                "Pages body object {identifier} text length overflows usize"
            ))
        })
    })?;
    if text_len > max_text_bytes {
        return Err(PackageError::Semantic(SemanticError::TextTooLarge {
            observed: text_len,
            limit: max_text_bytes,
        }));
    }
    Ok(())
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

#[cfg(test)]
mod tests {
    use std::fs;
    use std::io::{self, Cursor, Write};

    use super::*;
    use litchi_iwa_core::{Archive, ArchiveObject, RawMessage, SnappyStream};
    use litchi_iwa_protos::tswp::{ObjectAttributeTable, object_attribute_table::ObjectAttribute};
    use litchi_iwa_protos::{tp, tsp::Reference};

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
        let decoded = ObjectAttribute {
            character_index: 0,
            object: Some(Reference {
                identifier: 42,
                ..Reference::default()
            }),
        };
        let duplicate_reference = [0x08, 0x00, 0x12, 0x02, 0x08, 0x2a, 0x12, 0x02, 0x08, 0x2b];
        let error = validate_section_table_entry_wire(
            &duplicate_reference,
            &decoded,
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
    fn semantic_component_boundary_has_no_physical_source_to_parse_again() -> PackageResult<()> {
        let source: Arc<[u8]> = package_bytes(Some("Single parse"), true, false)?.into();
        let weak_source = Arc::downgrade(&source);
        let prepared = PreparedSource::from_shared_bytes(Arc::clone(&source))
            .map_err(|error| PackageError::InvalidFormat(error.to_string()))?
            .ok_or_else(|| {
                PackageError::InvalidFormat("Pages fixture was not detected".to_owned())
            })?;
        drop(source);

        let (components, limits) = semantic_components_from_prepared_source(prepared)?;
        assert!(weak_source.upgrade().is_none());

        let document = decode_semantic_components(
            &components,
            MAX_SECTIONS,
            effective_text_limit(limits),
            limits,
        )?;
        assert_eq!(document.plain_text(), "Single parse");
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
            package.sections()[0].text_storages()[0].runs(),
            [Run::new(0, 5)]
        );
        assert_eq!(package.sections()[1].name(), Some("Appendix"));
        assert_eq!(package.sections()[1].plain_text(), "B");
        assert_eq!(
            package
                .semantic_document()
                .section_named("Appendix")
                .unwrap_or_else(|error| panic!("unique native name should resolve: {error}"))
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
            package.semantic_document().section_named("Repeated").err(),
            Some(crate::SelectorError::AmbiguousSectionName {
                name: "Repeated".into(),
                first: 0,
                duplicate: 1,
            })
        );
        assert_eq!(
            package
                .semantic_document()
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
                .contains("duplicate or unsorted")
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
    fn input_limit_is_enforced_before_native_projection() {
        let limits = Limits::new(1, 10, 100, 100, 100)
            .unwrap_or_else(|error| panic!("test limits should be valid: {error}"));
        let error = Package::from_bytes_with_limits(&[0, 1], limits)
            .err()
            .unwrap_or_else(|| panic!("oversized input should fail"));
        assert!(error.to_string().contains("limit"));
    }
}
