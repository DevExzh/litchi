//! Native Numbers package parsing.
//!
//! This adapter owns physical package ingress, IWA-object lookup, and
//! protobuf-to-semantic conversion. The archive-free [`crate::Document`]
//! remains the public semantic owner; this module retains no dependency on the
//! historical umbrella facade.

#[allow(
    dead_code,
    reason = "The parser keeps private native table helpers together so all IWA table variants share one bounded decoder."
)]
mod extractor;
#[allow(
    dead_code,
    reason = "Formula-name reverse lookup is retained with the native token registry for future write support."
)]
mod function_map;
#[allow(
    dead_code,
    reason = "The compact index retains type probes used by the complete native table decoder."
)]
mod index;
mod limits;
/// Exact-source sheet and table name transactions.
pub(crate) mod names;
#[allow(
    dead_code,
    reason = "Decoded sheets expose only the construction path used at package ingress."
)]
mod sheet;
pub(crate) mod sheet_order;
#[allow(
    dead_code,
    reason = "Private native tables retain sidecar helpers while the public surface exposes only semantic tables."
)]
mod table;
pub(crate) mod table_cell_edit;
pub(crate) mod table_cells;
pub(crate) mod table_headers;
mod table_lock;
pub(crate) mod table_title;

use std::collections::HashSet;
use std::fmt;
use std::fs::{Metadata, OpenOptions};
use std::io::{self, Read, Write};
use std::path::Path;
use std::sync::Arc;

use litchi_iwa_archive::{ComponentCatalog, SourceCatalog};
use litchi_iwa_common::WireLimits;
use litchi_iwa_common::wire::{WireDescent, preflight_wire_tree_with_limits};
use litchi_iwa_core::{Archive, RawMessage};
#[cfg(feature = "internal-iwork-source")]
use litchi_iwa_detect::PreparedSource;
use litchi_iwa_detect::{Format, detect_application_from_document};
use litchi_iwa_protos::table_info_codec;
use litchi_iwa_protos::{tn, tswp};
use prost::Message;
use thiserror::Error;

use crate::{Document, DocumentError, DocumentLimits, Sheet};
use extractor::TableDataExtractor;
use index::{Index, Resolved};
use sheet::DecodedSheet;

pub use limits::{
    MAX_OBJECTS, MAX_REFERENCES, ReadOptions, SemanticLimitKind, SemanticLimits,
    SemanticLimitsError,
};
/// Physical ingress ceilings for a parsed Numbers package.
pub use litchi_iwa_archive::Limits;
pub use table_lock::{
    TableLockCommit, TableLockDiagnostics, TableLockEdit, TableLockError, TableLockLimitKind,
    TableLockPatch,
};

const DOCUMENT_MESSAGE_TYPE: u32 = 1;
const SHEET_MESSAGE_TYPE: u32 = 2;
const FORM_BASED_SHEET_MESSAGE_TYPE: u32 = 3;
const TABLE_INFO_MESSAGE_TYPE: u32 = 6_000;
const TABLE_MODEL_MESSAGE_TYPE: u32 = 6_001;
const LEGACY_TABLE_INFO_MESSAGE_TYPE: u32 = 6_003;
const TABLE_INFO_PROJECTION_RECURSION_LIMIT: u32 = 2;

/// Content-free semantic location associated with a Numbers read failure.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum SemanticPath {
    /// Whole-package ingress or indexing.
    Package,
    /// The Numbers document root.
    Document,
    /// One rooted sheet at its zero-based source position.
    Sheet { index: usize },
    /// One rooted drawable at its zero-based position within a sheet.
    Drawable { sheet: usize, index: usize },
    /// The global compatibility table projection.
    StructuredTables,
}

/// Bounded payload resource reported by the native Numbers decoder.
///
/// This Numbers-owned vocabulary prevents callers from depending on the
/// format-neutral IWA implementation crate merely to classify package read
/// failures.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum PayloadLimitKind {
    /// Bytes inspected in one encoded payload.
    InputBytes,
    /// Protobuf fields inspected in one encoded payload.
    Fields,
    /// Bytes produced while rewriting one encoded payload.
    OutputBytes,
    /// Nested length-delimited traversal depth.
    Nesting,
    /// Aggregate payload traversal or rewrite work.
    RewriteWork,
    /// Addressable rows declared by a table payload.
    TableRows,
    /// Addressable columns declared by a table payload.
    TableColumns,
    /// Addressable cells implied by table dimensions.
    TableCells,
    /// Sparse cells materialized from table payloads.
    MaterializedCells,
}

/// Content-free resource failure normalized at the Numbers package boundary.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum ResourceError {
    /// A finite native-payload resource ceiling was exceeded.
    LimitExceeded {
        /// Resource that exceeded its ceiling.
        kind: PayloadLimitKind,
        /// Observed or requested amount.
        observed: usize,
        /// Configured maximum.
        maximum: usize,
    },
    /// A fallible allocation could not reserve the requested capacity.
    Allocation {
        /// Elements or bytes requested by the failed reservation.
        amount: usize,
    },
}

impl fmt::Display for SemanticPath {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Package => formatter.write_str("package"),
            Self::Document => formatter.write_str("document"),
            Self::Sheet { index } => write!(formatter, "sheet {index}"),
            Self::Drawable { sheet, index } => {
                write!(formatter, "sheet {sheet} drawable {index}")
            },
            Self::StructuredTables => formatter.write_str("structured tables"),
        }
    }
}

/// Errors returned while parsing a native Numbers package.
#[derive(Debug, Error)]
#[non_exhaustive]
pub enum Error {
    /// Reading a package from the filesystem failed.
    #[error("could not read Numbers package: {0}")]
    Io(#[from] io::Error),
    /// ZIP or IWA package ingress failed.
    #[error(transparent)]
    Archive(#[from] litchi_iwa_archive::Error),
    /// A native payload could not be decoded at a stable semantic location.
    #[error("malformed Numbers payload at {path}")]
    MalformedPayload {
        /// Content-free semantic location of the malformed payload.
        path: SemanticPath,
    },
    /// The package is a recognized iWork document owned by another application.
    #[error("package is not a Numbers document")]
    NotNumbers,
    /// A native IWA value could not be decoded or validated.
    #[error(transparent)]
    Common(#[from] litchi_iwa_common::Error),
    /// The package does not contain a valid Numbers document structure.
    #[error("invalid Numbers package: {0}")]
    InvalidFormat(String),
    /// A table-cell record cannot be interpreted safely.
    #[error("could not parse Numbers table data: {0}")]
    ParseError(String),
    /// Semantic ingress rejected the decoded sheet sequence.
    #[error("invalid Numbers semantic document: {0}")]
    Semantic(#[from] DocumentError),
    /// A package-wide semantic resource ceiling was exceeded.
    #[error(
        "Numbers semantic {kind} limit exceeded at {path}: observed {observed}, maximum {maximum}"
    )]
    SemanticLimit {
        /// Resource category that exceeded its ceiling.
        kind: SemanticLimitKind,
        /// Observed or requested amount.
        observed: usize,
        /// Caller-selected or fixed adapter-owned maximum.
        maximum: usize,
        /// Content-free semantic location where the limit was encountered.
        path: SemanticPath,
    },
    /// The source file exceeds the selected physical input ceiling.
    #[error("Numbers package is {observed} bytes; maximum is {maximum}")]
    InputTooLarge {
        /// Source size observed before allocating the package buffer.
        observed: u64,
        /// Maximum input size selected by the caller.
        maximum: u64,
    },
}

/// Failure while streaming an exact Numbers package artifact to a caller-owned sink.
///
/// Its `Display` and `Debug` representations report only the offset reached
/// by prior conforming successful writes and the sink error kind; they never
/// include package bytes or sink error text.
#[derive(Error)]
#[error("could not write Numbers package after {bytes_written} bytes ({kind:?})")]
pub struct WriteError {
    source: io::Error,
    kind: io::ErrorKind,
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
    pub const fn io_error(&self) -> &io::Error {
        &self.source
    }

    /// Consume this error and return the underlying sink error.
    #[must_use]
    pub fn into_io_error(self) -> io::Error {
        self.source
    }
}

impl Error {
    /// Return the content-free bounded-resource classification, when present.
    ///
    /// The returned value carries neither authored content nor an internal
    /// allocation-site label. Callers can therefore log or translate it
    /// without exposing implementation details from the decoded document.
    #[must_use]
    pub const fn resource_error(&self) -> Option<ResourceError> {
        match self {
            Self::Common(error) => match error {
                litchi_iwa_common::Error::LimitExceeded {
                    kind,
                    observed,
                    limit,
                } => Some(ResourceError::LimitExceeded {
                    kind: match kind {
                        litchi_iwa_common::LimitKind::InputBytes => PayloadLimitKind::InputBytes,
                        litchi_iwa_common::LimitKind::Fields => PayloadLimitKind::Fields,
                        litchi_iwa_common::LimitKind::OutputBytes => PayloadLimitKind::OutputBytes,
                        litchi_iwa_common::LimitKind::Nesting => PayloadLimitKind::Nesting,
                        litchi_iwa_common::LimitKind::RewriteWork => PayloadLimitKind::RewriteWork,
                        litchi_iwa_common::LimitKind::TableRows => PayloadLimitKind::TableRows,
                        litchi_iwa_common::LimitKind::TableColumns => {
                            PayloadLimitKind::TableColumns
                        },
                        litchi_iwa_common::LimitKind::TableCells => PayloadLimitKind::TableCells,
                        litchi_iwa_common::LimitKind::MaterializedCells => {
                            PayloadLimitKind::MaterializedCells
                        },
                    },
                    observed: *observed,
                    maximum: *limit,
                }),
                litchi_iwa_common::Error::Allocation { amount, .. } => {
                    Some(ResourceError::Allocation { amount: *amount })
                },
                litchi_iwa_common::Error::InvalidFormat(_)
                | litchi_iwa_common::Error::InvalidLimit { .. } => None,
            },
            Self::Io(_)
            | Self::Archive(_)
            | Self::MalformedPayload { .. }
            | Self::NotNumbers
            | Self::InvalidFormat(_)
            | Self::ParseError(_)
            | Self::Semantic(_)
            | Self::SemanticLimit { .. }
            | Self::InputTooLarge { .. } => None,
        }
    }

    fn protobuf(_error: prost::DecodeError) -> Self {
        Self::MalformedPayload {
            path: SemanticPath::StructuredTables,
        }
    }

    fn malformed_payload(_error: prost::DecodeError, path: SemanticPath) -> Self {
        Self::MalformedPayload { path }
    }
}

/// Result returned by native Numbers package operations.
pub type Result<T> = std::result::Result<T, Error>;

#[derive(Debug)]
enum Components {
    Physical(Box<SourceCatalog>),
    #[cfg(feature = "internal-iwork-source")]
    Semantic(Arc<ComponentCatalog>),
}

impl Components {
    fn from_bytes(bytes: &[u8], limits: Limits) -> Result<Self> {
        Ok(Self::Physical(Box::new(
            SourceCatalog::from_bytes_with_limits(bytes, limits)?,
        )))
    }

    fn from_shared_bytes(source: Arc<[u8]>, limits: Limits) -> Result<Self> {
        Ok(Self::Physical(Box::new(
            SourceCatalog::from_shared_bytes_with_limits(source, limits)?,
        )))
    }

    #[cfg(feature = "internal-iwork-source")]
    fn from_catalog(catalog: Arc<ComponentCatalog>) -> Self {
        Self::Semantic(catalog)
    }

    fn catalog(&self) -> &ComponentCatalog {
        match self {
            Self::Physical(source) => source.components(),
            #[cfg(feature = "internal-iwork-source")]
            Self::Semantic(catalog) => catalog,
        }
    }

    fn physical(&self) -> Option<&SourceCatalog> {
        match self {
            Self::Physical(source) => Some(source),
            #[cfg(feature = "internal-iwork-source")]
            Self::Semantic(_) => None,
        }
    }

    fn get_archive(&self, name: &str) -> Option<&Archive> {
        self.catalog()
            .get(name)
            .map(litchi_iwa_archive::Component::archive)
    }

    fn iter_archives(&self) -> impl Iterator<Item = (&str, &Archive)> {
        self.catalog()
            .iter()
            .map(|component| (component.name(), component.archive()))
    }

    fn iter_objects(&self) -> impl Iterator<Item = &litchi_iwa_core::ArchiveObject> {
        self.iter_archives()
            .flat_map(|(_name, archive)| archive.objects.iter())
    }
}

/// A parsed native Numbers package and its immutable semantic projection.
///
/// Cloning this value or calling [`Self::snapshot`] shares the physical IWA
/// catalog, object index, and semantic sheet allocation without copying any
/// ZIP member, protobuf payload, table, or cell value.
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
    source: Arc<[u8]>,
    components: Components,
    index: Index,
    document: Document,
    options: ReadOptions,
}

#[derive(Debug)]
struct SemanticBudget {
    limits: SemanticLimits,
    references: usize,
    tables: usize,
}

impl SemanticBudget {
    const fn new(limits: SemanticLimits) -> Self {
        Self {
            limits,
            references: 0,
            tables: 0,
        }
    }

    fn charge_references(&mut self, amount: usize, path: SemanticPath) -> Result<()> {
        self.references = checked_charge(
            self.references,
            amount,
            self.limits.max_references(),
            SemanticLimitKind::References,
            path,
        )?;
        Ok(())
    }

    fn charge_table(&mut self, path: SemanticPath) -> Result<()> {
        self.tables = checked_charge(
            self.tables,
            1,
            self.limits.max_tables(),
            SemanticLimitKind::Tables,
            path,
        )?;
        Ok(())
    }
}

impl Package {
    /// Open a Numbers package from a filesystem path using default limits.
    ///
    /// # Errors
    ///
    /// Returns a typed error when the path cannot be read, the package exceeds
    /// a physical ceiling, or its IWA/Numbers contents are malformed.
    pub fn open(path: impl AsRef<Path>) -> Result<Self> {
        Self::open_with_options(path, ReadOptions::default())
    }

    /// Open a Numbers package from a filesystem path under explicit limits.
    ///
    /// # Errors
    ///
    /// Returns a typed error before allocating more than the selected input
    /// ceiling, or when the package cannot become a semantic document.
    pub fn open_with_limits(path: impl AsRef<Path>, limits: Limits) -> Result<Self> {
        Self::open_with_options(path, ReadOptions::new(limits, SemanticLimits::default()))
    }

    /// Open a Numbers package under explicit physical and semantic limits.
    ///
    /// # Errors
    ///
    /// Returns a typed error when either selected profile is exceeded, or when
    /// the package cannot become a strict rooted semantic document.
    pub fn open_with_options(path: impl AsRef<Path>, options: ReadOptions) -> Result<Self> {
        let bytes = read_source(path.as_ref(), options.archive())?;
        Self::from_shared_bytes_with_options(bytes.into(), options)
    }

    /// Parse a Numbers package from an in-memory ZIP payload using defaults.
    ///
    /// # Errors
    ///
    /// Returns a typed error when physical ingress, IWA framing, protobuf
    /// decoding, or semantic construction fails.
    pub fn from_bytes(bytes: &[u8]) -> Result<Self> {
        Self::from_bytes_with_options(bytes, ReadOptions::default())
    }

    /// Parse a Numbers package from an in-memory ZIP payload under limits.
    ///
    /// # Errors
    ///
    /// Returns a typed error when the package exceeds a selected resource
    /// ceiling or cannot be decoded as a Numbers document.
    pub fn from_bytes_with_limits(bytes: &[u8], limits: Limits) -> Result<Self> {
        Self::from_bytes_with_options(bytes, ReadOptions::new(limits, SemanticLimits::default()))
    }

    /// Parse a Numbers package under explicit physical and semantic limits.
    ///
    /// Rooted document construction is failure-atomic: object, sheet,
    /// reference, and table ceilings are checked before the semantic snapshot
    /// is published.
    ///
    /// # Errors
    ///
    /// Returns a typed error when the package exceeds either selected profile,
    /// contains ambiguous rooted ownership, or cannot be decoded as Numbers.
    pub fn from_bytes_with_options(bytes: &[u8], options: ReadOptions) -> Result<Self> {
        let components = Components::from_bytes(bytes, options.archive())?;
        Self::from_components_with_options(components, options)
    }

    fn from_shared_bytes_with_options(source: Arc<[u8]>, options: ReadOptions) -> Result<Self> {
        let components = Components::from_shared_bytes(source, options.archive())?;
        Self::from_components_with_options(components, options)
    }

    fn from_components_with_options(components: Components, options: ReadOptions) -> Result<Self> {
        let source = components
            .physical()
            .ok_or_else(|| {
                Error::InvalidFormat(
                    "semantic-only Numbers components cannot construct a package".to_owned(),
                )
            })?
            .shared_source();
        validate_numbers_application(&components, options.archive())?;
        let semantic = options.semantic();
        let index = Index::from_components(&components, semantic.max_objects())?;
        let root = Self::root_document(&components)?;
        let sheets = Self::decode_sheets(&components, &index, &root, semantic)?;
        let document = Document::from_sheets_with_limits(
            sheets,
            DocumentLimits::new(
                semantic.max_sheets(),
                semantic.max_tables(),
                semantic.max_materialized_cells(),
                semantic.max_output_text_bytes(),
            ),
        )?;
        Ok(Self {
            state: Arc::new(State {
                source,
                components,
                index,
                document,
                options,
            }),
        })
    }

    /// Capture a cheap immutable handle to the same parsed package.
    #[must_use]
    pub fn snapshot(&self) -> Self {
        self.clone()
    }

    /// Borrow the exact immutable package source retained by this snapshot.
    ///
    /// Unsupported ZIP members and unmodeled protobuf fields remain in this
    /// source even though the ordinary package API exposes semantic values.
    #[must_use]
    pub(crate) fn source_bytes(&self) -> &[u8] {
        &self.state.source
    }

    /// Write this validated package artifact to a byte sink.
    ///
    /// The emitted bytes are the exact source retained by this immutable
    /// snapshot, including unsupported ZIP members and unmodeled fields.
    /// This streaming primitive does not flush or synchronize `writer`, and it
    /// does not atomically publish a destination. Callers that need durable or
    /// atomic replacement must provide that policy around the sink.
    ///
    /// # Costs
    ///
    /// Streams `O(package size)` bytes in one forward pass without constructing
    /// another package-sized buffer. Partial or interrupted writes may call the
    /// sink more than once.
    ///
    /// # Errors
    ///
    /// Returns [`WriteError`] with the byte offset reached by prior conforming
    /// successful writes. A zero-length write or over-report is detected at
    /// that offset; an over-report does not establish its actual accepted-byte
    /// count. Bytes accepted before an error remain in the caller-owned sink.
    pub fn write_to<W: Write + ?Sized>(
        &self,
        writer: &mut W,
    ) -> std::result::Result<(), WriteError> {
        let source = self.source_bytes();
        let mut bytes_written = 0_usize;
        while bytes_written < source.len() {
            let remaining = &source[bytes_written..];
            match writer.write(remaining) {
                Ok(0) => {
                    return Err(WriteError {
                        source: io::Error::new(
                            io::ErrorKind::WriteZero,
                            "sink accepted no package bytes",
                        ),
                        kind: io::ErrorKind::WriteZero,
                        bytes_written,
                    });
                },
                Ok(amount) if amount <= remaining.len() => bytes_written += amount,
                Ok(_amount) => {
                    return Err(WriteError {
                        source: io::Error::new(
                            io::ErrorKind::InvalidData,
                            "sink reported accepting more bytes than supplied",
                        ),
                        kind: io::ErrorKind::InvalidData,
                        bytes_written,
                    });
                },
                Err(error) if error.kind() == io::ErrorKind::Interrupted => {},
                Err(write_error) => {
                    let kind = write_error.kind();
                    return Err(WriteError {
                        source: write_error,
                        kind,
                        bytes_written,
                    });
                },
            }
        }
        Ok(())
    }

    /// Borrow decoded semantic sheets in stable source order.
    #[must_use]
    pub fn sheets(&self) -> &[Sheet] {
        self.state.document.sheets()
    }

    /// Clone the shared semantic sheet allocation without cloning sheet data.
    #[must_use]
    pub fn shared_sheets(&self) -> Arc<[Sheet]> {
        self.state.document.shared_sheets()
    }

    /// Borrow the archive-free semantic Numbers document.
    #[must_use]
    pub fn document(&self) -> &Document {
        &self.state.document
    }

    /// Capture the archive-free semantic document snapshot.
    #[must_use]
    pub fn document_snapshot(&self) -> Document {
        self.state.document.snapshot()
    }

    /// Return the physical and semantic profiles retained by this snapshot.
    #[must_use]
    pub fn read_options(&self) -> ReadOptions {
        self.state.options
    }

    /// Extract the legacy-compatible global table projection.
    ///
    /// Unlike [`Self::document`], this allocating compatibility view scans
    /// package objects by their primary native type. It preserves the former
    /// structured extractor's deterministic order and includes valid detached
    /// table models. Type-6001 models precede legacy type-6000 models; object
    /// identity orders each group and duplicate candidates are emitted once.
    /// Ordinary sheets never expose those detached models.
    ///
    /// # Errors
    ///
    /// Returns a typed error when a preferred table-model payload is malformed,
    /// a referenced table sidecar is invalid, allocation fails, or the selected
    /// semantic table ceiling would be exceeded.
    pub fn extract_structured_tables(&self) -> Result<Vec<crate::Table>> {
        project_compatibility_tables(
            &self.state.components,
            &self.state.index,
            self.state.options.semantic(),
        )
    }

    /// Return the count of indexed IWA objects retained by this package.
    #[must_use]
    pub fn object_count(&self) -> usize {
        self.state.index.object_count()
    }

    /// Extract all native rich-text storages in deterministic archive order.
    ///
    /// Storage objects are preserved separately from semantic tables because
    /// Numbers may retain text for shapes and auxiliary objects. Each decoded
    /// storage is separated with one newline, matching the former IWA reader.
    ///
    /// # Errors
    ///
    /// Returns a typed error if decoded text cannot be represented safely.
    pub fn text(&self) -> Result<String> {
        const STORAGE_TYPES: [u32; 14] = [
            200, 201, 202, 203, 204, 205, 2001, 2002, 2003, 2004, 2005, 2011, 2012, 2022,
        ];
        let mut text = String::new();
        for object in self.state.components.iter_objects() {
            for message in &object.messages {
                if !STORAGE_TYPES.contains(&message.type_) {
                    continue;
                }
                let Ok(storage) = tswp::StorageArchive::decode(message.data.as_slice()) else {
                    continue;
                };
                if storage.text.is_empty() {
                    continue;
                }
                if !text.is_empty() {
                    text.push('\n');
                }
                text.push_str(&storage.text.join("\n"));
            }
        }
        Ok(text)
    }

    fn root_document(components: &Components) -> Result<tn::DocumentArchive> {
        let object = components
            .get_archive("Index/Document.iwa")
            .and_then(|archive| archive.object(1))
            .ok_or_else(|| {
                Error::InvalidFormat("package does not contain a Numbers root document".to_owned())
            })?;
        let message = unique_message(
            &object.messages,
            DOCUMENT_MESSAGE_TYPE,
            SemanticPath::Document,
            "document",
        )?
        .ok_or_else(|| {
            Error::InvalidFormat(
                "Numbers document root has no canonical document payload".to_owned(),
            )
        })?;
        tn::DocumentArchive::decode(message.data.as_slice())
            .map_err(|error| Error::malformed_payload(error, SemanticPath::Document))
    }

    fn decode_sheets(
        components: &Components,
        index: &Index,
        document: &tn::DocumentArchive,
        limits: SemanticLimits,
    ) -> Result<Vec<Sheet>> {
        if document.sheets.len() > limits.max_sheets() {
            return Err(Error::SemanticLimit {
                kind: SemanticLimitKind::Sheets,
                observed: document.sheets.len(),
                maximum: limits.max_sheets(),
                path: SemanticPath::Document,
            });
        }

        let mut budget = SemanticBudget::new(limits);
        budget.charge_references(document.sheets.len(), SemanticPath::Document)?;
        let extractor = TableDataExtractor::new(components, index, limits);
        let mut sheets = Vec::new();
        sheets
            .try_reserve_exact(document.sheets.len())
            .map_err(|_error| {
                Error::Common(litchi_iwa_common::Error::Allocation {
                    resource: "Numbers semantic sheets",
                    amount: document.sheets.len(),
                })
            })?;

        let mut seen_sheets = HashSet::new();
        seen_sheets
            .try_reserve(document.sheets.len())
            .map_err(|_error| {
                allocation_error("Numbers rooted sheet identities", document.sheets.len())
            })?;
        let mut seen_drawables = HashSet::new();
        let mut seen_models = HashSet::new();
        for (position, reference) in document.sheets.iter().enumerate() {
            let path = SemanticPath::Sheet { index: position };
            if !seen_sheets.insert(reference.identifier) {
                return Err(Error::InvalidFormat(format!(
                    "Numbers {path} repeats an earlier rooted sheet"
                )));
            }
            let object = index
                .resolve_ref_id(components, reference.identifier)?
                .ok_or_else(|| Error::InvalidFormat(format!("Numbers {path} is missing")))?;
            let archive = decode_sheet_payload(object.messages, path)?;
            budget.charge_references(archive.drawable_infos.len(), path)?;
            seen_drawables
                .try_reserve(archive.drawable_infos.len())
                .map_err(|_error| {
                    allocation_error(
                        "Numbers rooted drawable identities",
                        archive.drawable_infos.len(),
                    )
                })?;
            extractor.charge_output_text(archive.name.len())?;
            let mut sheet = DecodedSheet::new(archive.name, position);
            for (drawable_position, drawable) in archive.drawable_infos.into_iter().enumerate() {
                let drawable_path = SemanticPath::Drawable {
                    sheet: position,
                    index: drawable_position,
                };
                if !seen_drawables.insert(drawable.identifier) {
                    return Err(Error::InvalidFormat(format!(
                        "Numbers {drawable_path} repeats an earlier rooted drawable"
                    )));
                }
                let Some(table) = Self::extract_table(
                    components,
                    index,
                    drawable.identifier,
                    &extractor,
                    drawable_path,
                    &mut budget,
                    &mut seen_models,
                )?
                else {
                    continue;
                };
                sheet.try_add_table(table)?;
            }
            sheets.push(sheet.into_semantic()?);
        }
        Ok(sheets)
    }

    fn extract_table(
        components: &Components,
        index: &Index,
        drawable_id: u64,
        extractor: &TableDataExtractor<'_>,
        path: SemanticPath,
        budget: &mut SemanticBudget,
        seen_models: &mut HashSet<u64>,
    ) -> Result<Option<table::Table>> {
        let resolved = index
            .resolve_ref_id(components, drawable_id)?
            .ok_or_else(|| Error::InvalidFormat(format!("Numbers {path} is missing")))?;
        let canonical = unique_message(
            resolved.messages,
            TABLE_INFO_MESSAGE_TYPE,
            path,
            "table-info",
        )?;
        let legacy = unique_message(
            resolved.messages,
            LEGACY_TABLE_INFO_MESSAGE_TYPE,
            path,
            "legacy table-info",
        )?;
        let message = match (canonical, legacy) {
            (Some(_), Some(_)) => {
                return Err(Error::InvalidFormat(format!(
                    "Numbers {path} has ambiguous table-info payload ownership"
                )));
            },
            (Some(message), None) | (None, Some(message)) => message,
            (None, None) => return Ok(None),
        };
        let model_reference = table_info_codec::decode_table_model_reference(
            message.data.as_slice(),
            table_info_decode_options(message.data.as_slice()),
        )
        .map_err(|_error| Error::MalformedPayload { path })?;
        budget.charge_references(1, path)?;
        let model_id = model_reference.identifier().get();
        let model = index.resolve_ref_id(components, model_id)?.ok_or_else(|| {
            Error::InvalidFormat(format!("Numbers {path} table model is missing"))
        })?;
        if seen_models.contains(&model_id) {
            return Err(Error::InvalidFormat(format!(
                "Numbers {path} reuses a table model owned by an earlier drawable"
            )));
        }
        budget.charge_table(path)?;
        seen_models.try_reserve(1).map_err(|_error| {
            allocation_error(
                "Numbers rooted table model identities",
                seen_models.len().saturating_add(1),
            )
        })?;
        let inserted = seen_models.insert(model_id);
        debug_assert!(
            inserted,
            "duplicate table models were rejected before admission"
        );
        extractor
            .extract_reachable_table_from_object(&model, path)
            .map(Some)
    }
}

/// Bounded policy for the focused table-info reference projection.
///
/// The raw message was already admitted under the caller's physical IWA
/// message ceiling. Its exact source width therefore bounds every projection
/// budget, while the codec independently bounds strict and Buffa scans.
pub(super) fn table_info_decode_options(source: &[u8]) -> table_info_codec::DecodeOptions {
    table_info_codec::DecodeOptions::new(
        source.len().max(1),
        source.len().max(1),
        source.len().saturating_mul(4).max(1),
        TABLE_INFO_PROJECTION_RECURSION_LIMIT,
    )
}

/// Extract the allocating, legacy-compatible global table projection from
/// complete Numbers package bytes under default limits.
///
/// This entry point validates Numbers application ownership and indexes the
/// physical package, but deliberately does not construct the strict rooted
/// [`Document`]. It therefore remains usable for migration callers that need
/// detached historical table models even when unrelated rooted topology is
/// incomplete. New application code should normally use [`Package::document`].
///
/// # Errors
///
/// Returns a typed error when the bytes are not an unambiguous Numbers package,
/// physical or semantic limits are exceeded, or a selected table is malformed.
pub fn compatibility_tables_from_bytes(bytes: &[u8]) -> Result<Vec<crate::Table>> {
    compatibility_tables_from_bytes_with_options(bytes, ReadOptions::default())
}

/// Extract the allocating global compatibility projection under explicit
/// physical and semantic limits.
///
/// This path does not construct the strict rooted [`Document`]. It may lazily
/// resolve source sheet and drawable names when enriching a non-empty formula
/// sidecar. The object and table portions of [`SemanticLimits`] are consumed
/// directly; its reference ceiling applies to unique source-derived
/// formula-enrichment entries. Formula discovery separately has fixed
/// aggregate category-wire, work, and text ceilings plus a fixed
/// category-depth ceiling. The sheet ceiling applies only when constructing a
/// strict [`Package`].
///
/// # Errors
///
/// Returns a typed error when application ownership, package ingress, object
/// indexing, table decoding, or a selected resource ceiling fails.
pub fn compatibility_tables_from_bytes_with_options(
    bytes: &[u8],
    options: ReadOptions,
) -> Result<Vec<crate::Table>> {
    let components = Components::from_bytes(bytes, options.archive())?;
    compatibility_tables_from_components(&components, options.archive(), options.semantic())
}

/// Consume one prepared source into the global Numbers compatibility view.
///
/// This explicitly unstable handoff is reserved for the root iWork
/// coordinator. It retains detached/orphan table behavior while avoiding a
/// second ZIP traversal and IWA component decode. The source's original
/// physical profile remains authoritative; callers select only semantic
/// projection limits here.
///
/// # Errors
///
/// Returns [`Error`] when the source belongs to another application or the
/// Numbers compatibility projection is malformed or over budget.
#[cfg(feature = "internal-iwork-source")]
#[doc(hidden)]
pub fn __compatibility_tables_from_prepared_source(
    source: PreparedSource,
    semantic: SemanticLimits,
) -> Result<Vec<crate::Table>> {
    if source.format() != Format::Numbers {
        return Err(Error::NotNumbers);
    }
    let (catalog, archive_limits) = source.__into_components();
    let components = Components::from_catalog(catalog);
    compatibility_tables_from_components(&components, archive_limits, semantic)
}

fn compatibility_tables_from_components(
    components: &Components,
    archive_limits: Limits,
    semantic: SemanticLimits,
) -> Result<Vec<crate::Table>> {
    validate_numbers_application(components, archive_limits)?;
    let index = Index::from_components(components, semantic.max_objects())?;
    project_compatibility_tables(components, &index, semantic)
}

fn validate_numbers_application(components: &Components, archive_limits: Limits) -> Result<()> {
    let root = components
        .get_archive("Index/Document.iwa")
        .and_then(|archive| archive.object(1))
        .ok_or_else(|| {
            Error::InvalidFormat("package has no canonical iWork document root".to_owned())
        })?;
    let canonical_message = unique_message(
        &root.messages,
        DOCUMENT_MESSAGE_TYPE,
        SemanticPath::Document,
        "document",
    )?;
    let Some(canonical) = canonical_message else {
        let mut foreign = root
            .messages
            .iter()
            .filter_map(|message| detect_application_from_document(&message.data))
            .filter(|format| matches!(format, Format::Pages | Format::Keynote));
        if foreign.next().is_some() && foreign.next().is_none() {
            return Err(Error::NotNumbers);
        }
        return Err(Error::InvalidFormat(
            "iWork document root has no canonical Numbers payload".to_owned(),
        ));
    };
    preflight_application_payload(&canonical.data, archive_limits)?;
    let application = detect_application_from_document(&canonical.data).ok_or_else(|| {
        Error::InvalidFormat(
            "canonical iWork document payload has no unambiguous application shape".to_owned(),
        )
    })?;
    match application {
        Format::Numbers => Ok(()),
        Format::Pages | Format::Keynote => Err(Error::NotNumbers),
    }
}

fn preflight_application_payload(payload: &[u8], archive_limits: Limits) -> Result<()> {
    let source_ceiling = payload.len().min(archive_limits.max_iwa_stream_bytes());
    let input_bytes = source_ceiling
        .saturating_mul(3)
        .clamp(1, WireLimits::MAX_INPUT_BYTES);
    let fields = source_ceiling
        .saturating_mul(2)
        .clamp(1, WireLimits::default().max_fields());
    let wire_limits = WireLimits::default()
        .with_input_bytes(input_bytes)?
        .with_fields(fields)?;
    preflight_wire_tree_with_limits(payload, wire_limits, |visit| {
        let field = visit.field().number();
        let descend = match visit.path() {
            [] => matches!(field, 2 | 3 | 4 | 5 | 6 | 8 | 15) && visit.field().wire_type() == 2,
            [3 | 8 | 15] => field == 1 && visit.field().wire_type() == 2,
            _ => false,
        };
        Ok(if descend {
            WireDescent::Descend
        } else {
            WireDescent::Skip
        })
    })?;
    Ok(())
}

fn project_compatibility_tables(
    components: &Components,
    index: &Index,
    limits: SemanticLimits,
) -> Result<Vec<crate::Table>> {
    if !TableDataExtractor::has_table_models(index) {
        return Ok(Vec::new());
    }
    TableDataExtractor::new(components, index, limits)
        .extract_all_semantic_tables(limits.max_tables())
}

fn checked_charge(
    current: usize,
    amount: usize,
    maximum: usize,
    kind: SemanticLimitKind,
    path: SemanticPath,
) -> Result<usize> {
    let observed = current.checked_add(amount).ok_or(Error::SemanticLimit {
        kind,
        observed: usize::MAX,
        maximum,
        path,
    })?;
    if observed > maximum {
        return Err(Error::SemanticLimit {
            kind,
            observed,
            maximum,
            path,
        });
    }
    Ok(observed)
}

fn unique_message<'a>(
    messages: &'a [RawMessage],
    message_type: u32,
    path: SemanticPath,
    label: &str,
) -> Result<Option<&'a RawMessage>> {
    let mut matches = messages
        .iter()
        .filter(|message| message.type_ == message_type);
    let Some(message) = matches.next() else {
        return Ok(None);
    };
    if matches.next().is_some() {
        return Err(Error::InvalidFormat(format!(
            "Numbers {path} contains duplicate {label} payloads"
        )));
    }
    Ok(Some(message))
}

fn decode_sheet_payload(messages: &[RawMessage], path: SemanticPath) -> Result<tn::SheetArchive> {
    let sheet = unique_message(messages, SHEET_MESSAGE_TYPE, path, "sheet")?;
    let form = unique_message(
        messages,
        FORM_BASED_SHEET_MESSAGE_TYPE,
        path,
        "form-based sheet",
    )?;
    match (sheet, form) {
        (Some(_), Some(_)) => Err(Error::InvalidFormat(format!(
            "Numbers {path} has ambiguous sheet payload ownership"
        ))),
        (Some(message), None) => tn::SheetArchive::decode(message.data.as_slice())
            .map_err(|error| Error::malformed_payload(error, path)),
        (None, Some(message)) => tn::FormBasedSheetArchive::decode(message.data.as_slice())
            .map(|form_archive| form_archive.super_)
            .map_err(|error| Error::malformed_payload(error, path)),
        (None, None) => Err(Error::InvalidFormat(format!(
            "Numbers {path} has no canonical sheet payload"
        ))),
    }
}

fn allocation_error(resource: &'static str, amount: usize) -> Error {
    Error::Common(litchi_iwa_common::Error::Allocation { resource, amount })
}

fn read_source(path: &Path, limits: Limits) -> Result<Vec<u8>> {
    let mut open_options = OpenOptions::new();
    open_options.read(true);
    #[cfg(unix)]
    {
        use std::os::unix::fs::OpenOptionsExt;
        open_options.custom_flags(libc::O_NONBLOCK);
    }
    let mut file = open_options.open(path)?;
    let metadata = file.metadata()?;
    if !metadata.is_file() {
        return Err(Error::InvalidFormat(
            "Numbers package source is not a regular file".to_owned(),
        ));
    }
    let reported_length = metadata.len();
    let bytes = read_source_with_reported_length(&mut file, reported_length, limits)?;
    ensure_source_unchanged(&metadata, &file.metadata()?)?;
    Ok(bytes)
}

fn ensure_source_unchanged(before: &Metadata, after: &Metadata) -> Result<()> {
    let modified = matches!(
        (before.modified(), after.modified()),
        (Ok(before_modified), Ok(after_modified)) if before_modified != after_modified
    );
    if before.len() != after.len() || modified {
        return Err(Error::InvalidFormat(
            "Numbers package source changed while it was being read".to_owned(),
        ));
    }
    Ok(())
}

fn read_source_with_reported_length(
    reader: &mut impl Read,
    reported_length: u64,
    limits: Limits,
) -> Result<Vec<u8>> {
    if reported_length > limits.max_input_bytes() {
        return Err(input_too_large(reported_length, limits));
    }

    let maximum = usize::try_from(limits.max_input_bytes()).map_err(|_error| {
        Error::InvalidFormat("Numbers input ceiling does not fit usize".to_owned())
    })?;
    let reported_capacity = usize::try_from(reported_length).map_err(|_error| {
        Error::InvalidFormat("Numbers input length does not fit usize".to_owned())
    })?;
    // File metadata is advisory: a sparse or concurrently truncated source
    // must not force one allocation proportional to its reported logical size.
    let capacity = reported_capacity.min(64 * 1024);
    let mut bytes = Vec::new();
    bytes
        .try_reserve_exact(capacity)
        .map_err(|_error| allocation_error("Numbers package input", capacity))?;
    loop {
        let remaining = maximum
            .checked_sub(bytes.len())
            .ok_or_else(|| Error::InvalidFormat("Numbers input length exceeds usize".to_owned()))?;
        if remaining == 0 {
            let mut extra = [0u8; 1];
            if reader.read(&mut extra)? != 0 {
                return Err(input_too_large(
                    limits.max_input_bytes().saturating_add(1),
                    limits,
                ));
            }
            break;
        }

        if bytes.len() == bytes.capacity() {
            // Probe EOF before growing beyond the descriptor's advisory
            // length. If one byte exists, retain it and grow geometrically.
            let mut extra = [0u8; 1];
            if reader.read(&mut extra)? == 0 {
                break;
            }
            let growth = bytes.capacity().max(8 * 1024).min(remaining);
            let target = bytes.len().checked_add(growth).ok_or_else(|| {
                Error::InvalidFormat("Numbers input length exceeds usize".to_owned())
            })?;
            reserve_source_growth(&mut bytes, target, maximum)?;
            bytes.push(extra[0]);
            continue;
        }

        let writable = bytes.capacity().saturating_sub(bytes.len()).min(remaining);
        let read_limit = u64::try_from(writable).map_err(|_error| {
            Error::InvalidFormat("Numbers input read size does not fit u64".to_owned())
        })?;
        // The standard `read_to_end` implementation fills `Vec` spare capacity
        // directly. Limiting it to the already-reserved region avoids both a
        // zero-fill pass and an infallible growth step.
        let read = reader.by_ref().take(read_limit).read_to_end(&mut bytes)?;
        if read < writable {
            break;
        }
    }
    Ok(bytes)
}

fn reserve_source_growth(bytes: &mut Vec<u8>, required: usize, maximum: usize) -> Result<()> {
    if required <= bytes.capacity() {
        return Ok(());
    }
    let doubled = bytes.capacity().checked_mul(2).unwrap_or(maximum);
    let target = required.max(doubled).min(maximum);
    let additional = target
        .checked_sub(bytes.len())
        .ok_or_else(|| Error::InvalidFormat("Numbers input length exceeds usize".to_owned()))?;
    bytes
        .try_reserve_exact(additional)
        .map_err(|_error| allocation_error("Numbers package input", target))
}

const fn input_too_large(observed: u64, limits: Limits) -> Error {
    Error::InputTooLarge {
        observed,
        maximum: limits.max_input_bytes(),
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use litchi_iwa_core::{ArchiveObject, RawMessage, SnappyStream};
    use litchi_iwa_protos::{kn, tp, tsa, tsce, tsk, tst};
    use std::io::Write;

    #[test]
    fn write_to_accepts_a_dynamically_dispatched_sink() -> Result<()> {
        let bytes = package_bytes(&tn::DocumentArchive::default())?;
        let package = Package::from_bytes(&bytes)?;
        let mut written = Vec::new();
        let sink: &mut dyn Write = &mut written;

        package
            .write_to(sink)
            .map_err(|error| Error::Io(error.into_io_error()))?;

        assert_eq!(written, bytes);
        Ok(())
    }

    #[test]
    fn write_to_reports_progress_and_handles_adversarial_sinks() -> Result<()> {
        struct FailingWriter {
            remaining: usize,
            output: Vec<u8>,
            flushes: usize,
        }

        impl Write for FailingWriter {
            fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
                if self.remaining == 0 {
                    return Err(io::Error::other("authored secret"));
                }
                let amount = bytes.len().min(self.remaining);
                self.output.extend_from_slice(&bytes[..amount]);
                self.remaining -= amount;
                Ok(amount)
            }

            fn flush(&mut self) -> io::Result<()> {
                self.flushes += 1;
                Ok(())
            }
        }

        struct InterruptedWriter {
            interrupted: bool,
            output: Vec<u8>,
            flushes: usize,
        }

        impl Write for InterruptedWriter {
            fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
                if !self.interrupted {
                    self.interrupted = true;
                    return Err(io::Error::from(io::ErrorKind::Interrupted));
                }
                self.output.extend_from_slice(bytes);
                Ok(bytes.len())
            }

            fn flush(&mut self) -> io::Result<()> {
                self.flushes += 1;
                Ok(())
            }
        }

        struct ZeroWriter;

        impl Write for ZeroWriter {
            fn write(&mut self, _bytes: &[u8]) -> io::Result<usize> {
                Ok(0)
            }

            fn flush(&mut self) -> io::Result<()> {
                Ok(())
            }
        }

        struct OverReportingWriter;

        impl Write for OverReportingWriter {
            fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
                Ok(bytes.len().saturating_add(1))
            }

            fn flush(&mut self) -> io::Result<()> {
                Ok(())
            }
        }

        let bytes = package_bytes(&tn::DocumentArchive::default())?;
        let package = Package::from_bytes(&bytes)?;

        let mut failing = FailingWriter {
            remaining: 7,
            output: Vec::new(),
            flushes: 0,
        };
        let error = package.write_to(&mut failing).unwrap_err();
        assert_eq!(error.bytes_written(), 7);
        assert_eq!(error.io_error().kind(), io::ErrorKind::Other);
        assert_eq!(failing.output, bytes[..7]);
        assert_eq!(failing.flushes, 0);
        assert!(!format!("{error:?}").contains("authored secret"));
        assert!(!error.to_string().contains("authored secret"));

        let mut interrupted = InterruptedWriter {
            interrupted: false,
            output: Vec::new(),
            flushes: 0,
        };
        package
            .write_to(&mut interrupted)
            .map_err(|error| Error::Io(error.into_io_error()))?;
        assert!(interrupted.interrupted);
        assert_eq!(interrupted.output, bytes);
        assert_eq!(interrupted.flushes, 0);

        let zero_error = package.write_to(&mut ZeroWriter).unwrap_err();
        assert_eq!(zero_error.bytes_written(), 0);
        assert_eq!(zero_error.io_error().kind(), io::ErrorKind::WriteZero);

        let over_report_error = package.write_to(&mut OverReportingWriter).unwrap_err();
        assert_eq!(over_report_error.bytes_written(), 0);
        assert_eq!(
            over_report_error.io_error().kind(),
            io::ErrorKind::InvalidData
        );

        Ok(())
    }

    #[test]
    fn common_limits_normalize_to_every_numbers_payload_resource() {
        let kinds = [
            (
                litchi_iwa_common::LimitKind::InputBytes,
                PayloadLimitKind::InputBytes,
            ),
            (
                litchi_iwa_common::LimitKind::Fields,
                PayloadLimitKind::Fields,
            ),
            (
                litchi_iwa_common::LimitKind::OutputBytes,
                PayloadLimitKind::OutputBytes,
            ),
            (
                litchi_iwa_common::LimitKind::Nesting,
                PayloadLimitKind::Nesting,
            ),
            (
                litchi_iwa_common::LimitKind::RewriteWork,
                PayloadLimitKind::RewriteWork,
            ),
            (
                litchi_iwa_common::LimitKind::TableRows,
                PayloadLimitKind::TableRows,
            ),
            (
                litchi_iwa_common::LimitKind::TableColumns,
                PayloadLimitKind::TableColumns,
            ),
            (
                litchi_iwa_common::LimitKind::TableCells,
                PayloadLimitKind::TableCells,
            ),
            (
                litchi_iwa_common::LimitKind::MaterializedCells,
                PayloadLimitKind::MaterializedCells,
            ),
        ];

        for (common, normalized) in kinds {
            let error = Error::from(litchi_iwa_common::Error::LimitExceeded {
                kind: common,
                observed: 9,
                limit: 8,
            });
            assert_eq!(
                error.resource_error(),
                Some(ResourceError::LimitExceeded {
                    kind: normalized,
                    observed: 9,
                    maximum: 8,
                })
            );
        }
    }

    #[test]
    fn common_allocation_normalization_preserves_amount_and_discards_label() {
        const PRIVATE: &str = "private-path/member/authored-value";
        let error = Error::from(litchi_iwa_common::Error::Allocation {
            resource: PRIVATE,
            amount: 17,
        });
        let normalized = error.resource_error();
        assert_eq!(normalized, Some(ResourceError::Allocation { amount: 17 }));
        assert!(!format!("{normalized:?}").contains(PRIVATE));
    }

    #[test]
    fn non_resource_common_failures_do_not_acquire_resource_metadata() {
        for error in [
            litchi_iwa_common::Error::InvalidFormat("private".to_owned()),
            litchi_iwa_common::Error::InvalidLimit {
                field: "private",
                value: 9,
                maximum: 8,
            },
        ] {
            assert_eq!(Error::from(error).resource_error(), None);
        }
    }

    fn reference(identifier: u64) -> litchi_iwa_protos::tsp::Reference {
        litchi_iwa_protos::tsp::Reference {
            identifier,
            ..Default::default()
        }
    }

    fn shared_document() -> tsa::DocumentArchive {
        tsa::DocumentArchive {
            super_: tsk::DocumentArchive::default(),
            ..Default::default()
        }
    }

    fn package_bytes(root: &tn::DocumentArchive) -> Result<Vec<u8>> {
        package_bytes_from_archive(Archive {
            objects: vec![object(1, DOCUMENT_MESSAGE_TYPE, root.encode_to_vec())?],
        })
    }

    fn package_with_two_materialized_empty_tables() -> Result<Vec<u8>> {
        let sidecars = ArchiveObject::new(
            20,
            [
                tst::table_data_list::ListType::String,
                tst::table_data_list::ListType::Formula,
            ]
            .into_iter()
            .map(|list_type| RawMessage {
                type_: 6_005,
                data: tst::TableDataList {
                    list_type: list_type as i32,
                    next_list_id: 1,
                    ..Default::default()
                }
                .encode_to_vec(),
            })
            .collect(),
        )
        .map_err(|error| Error::InvalidFormat(error.to_string()))?;
        let empty_cell = crate::cell::wire::BncCell::minimal().encode();
        let mut objects = vec![
            object(
                1,
                DOCUMENT_MESSAGE_TYPE,
                tn::DocumentArchive::default().encode_to_vec(),
            )?,
            sidecars,
        ];
        for (model_id, tile_id, name) in [(10, 30, "A"), (11, 31, "B")] {
            let model = tst::TableModelArchive {
                table_name: name.to_owned(),
                number_of_rows: 1,
                number_of_columns: 1,
                base_data_store: tst::DataStore {
                    tiles: tst::TileStorage {
                        tiles: vec![tst::tile_storage::Tile {
                            tileid: 0,
                            tile: reference(tile_id),
                        }],
                        tile_size: Some(256),
                        ..Default::default()
                    },
                    string_table: reference(20),
                    formula_table: reference(20),
                    ..Default::default()
                },
                ..Default::default()
            };
            let tile = tst::Tile {
                numrows: 1,
                row_infos: vec![tst::TileRowInfo {
                    tile_row_index: 0,
                    cell_count: 1,
                    storage_version: Some(5),
                    cell_storage_buffer: Some(empty_cell.clone()),
                    cell_offsets: Some(vec![0, 0]),
                    ..Default::default()
                }],
                ..Default::default()
            };
            objects.push(object(
                model_id,
                TABLE_MODEL_MESSAGE_TYPE,
                model.encode_to_vec(),
            )?);
            objects.push(object(tile_id, 6_002, tile.encode_to_vec())?);
        }
        package_bytes_from_archive(Archive { objects })
    }

    fn rooted_two_table_package(
        second_model_id: u64,
        second_model_payload: Vec<u8>,
    ) -> Result<Vec<u8>> {
        let sidecars = ArchiveObject::new(
            7,
            [
                tst::table_data_list::ListType::String,
                tst::table_data_list::ListType::Formula,
            ]
            .into_iter()
            .map(|list_type| RawMessage {
                type_: 6_005,
                data: tst::TableDataList {
                    list_type: list_type as i32,
                    next_list_id: 1,
                    ..Default::default()
                }
                .encode_to_vec(),
            })
            .collect(),
        )
        .map_err(|error| Error::InvalidFormat(error.to_string()))?;
        let first_model = tst::TableModelArchive {
            table_name: "First".to_owned(),
            base_data_store: tst::DataStore {
                string_table: reference(7),
                formula_table: reference(7),
                ..Default::default()
            },
            ..Default::default()
        };
        let table_info = |model_id| {
            tst::TableInfoArchive {
                table_model: reference(model_id),
                ..Default::default()
            }
            .encode_to_vec()
        };
        let mut objects = vec![
            object(
                1,
                DOCUMENT_MESSAGE_TYPE,
                tn::DocumentArchive {
                    sheets: vec![reference(2)],
                    ..Default::default()
                }
                .encode_to_vec(),
            )?,
            object(
                2,
                SHEET_MESSAGE_TYPE,
                tn::SheetArchive {
                    name: "Admission".to_owned(),
                    drawable_infos: vec![reference(3), reference(4)],
                    ..Default::default()
                }
                .encode_to_vec(),
            )?,
            object(3, TABLE_INFO_MESSAGE_TYPE, table_info(5))?,
            object(4, TABLE_INFO_MESSAGE_TYPE, table_info(second_model_id))?,
            object(5, TABLE_MODEL_MESSAGE_TYPE, first_model.encode_to_vec())?,
            sidecars,
        ];
        if second_model_id != 5 {
            objects.push(object(
                second_model_id,
                TABLE_MODEL_MESSAGE_TYPE,
                second_model_payload,
            )?);
        }
        package_bytes_from_archive(Archive { objects })
    }

    fn package_with_formula_reference(table_info_data: Vec<u8>) -> Result<Vec<u8>> {
        const OWNER_WORDS: [u32; 4] = [0x89ab_cdef, 0x0123_4567, 0x7654_3210, 0xfedc_ba98];
        let owner_uuid = litchi_iwa_protos::tsp::Uuid {
            lower: u64::from(OWNER_WORDS[0]) | (u64::from(OWNER_WORDS[1]) << 32),
            upper: u64::from(OWNER_WORDS[2]) | (u64::from(OWNER_WORDS[3]) << 32),
        };
        let cross_table_id = litchi_iwa_protos::tsp::CfuuidArchive {
            uuid_w0: Some(OWNER_WORDS[0]),
            uuid_w1: Some(OWNER_WORDS[1]),
            uuid_w2: Some(OWNER_WORDS[2]),
            uuid_w3: Some(OWNER_WORDS[3]),
            ..Default::default()
        };
        let formula = tsce::FormulaArchive {
            ast_node_array: tsce::AstNodeArrayArchive {
                ast_node: vec![tsce::ast_node_array_archive::AstNodeArchive {
                    ast_node_type: tsce::ast_node_array_archive::AstNodeType::CellReferenceNode
                        as i32,
                    ast_column: Some(tsce::ast_node_array_archive::AstColumnCoordinateArchive {
                        column: 0,
                        absolute: Some(false),
                    }),
                    ast_row: Some(tsce::ast_node_array_archive::AstRowCoordinateArchive {
                        row: 0,
                        absolute: Some(false),
                    }),
                    ast_cross_table_reference_extra_info: Some(
                        tsce::ast_node_array_archive::AstCrossTableReferenceExtraInfoArchive {
                            table_id: cross_table_id,
                            ..Default::default()
                        },
                    ),
                    ..Default::default()
                }],
            },
            ..Default::default()
        };
        let sidecars = ArchiveObject::new(
            5,
            [
                tst::TableDataList {
                    list_type: tst::table_data_list::ListType::String as i32,
                    next_list_id: 1,
                    ..Default::default()
                },
                tst::TableDataList {
                    list_type: tst::table_data_list::ListType::Formula as i32,
                    next_list_id: 2,
                    entries: vec![tst::table_data_list::ListEntry {
                        key: 1,
                        refcount: 1,
                        formula: Some(formula),
                        ..Default::default()
                    }],
                    ..Default::default()
                },
            ]
            .into_iter()
            .map(|list| RawMessage {
                type_: 6_005,
                data: list.encode_to_vec(),
            })
            .collect(),
        )
        .map_err(|error| Error::InvalidFormat(error.to_string()))?;
        let mut formula_cell = crate::cell::wire::BncCell::minimal();
        formula_cell.set_formula_reference(1);
        let model = tst::TableModelArchive {
            table_name: "Table".to_owned(),
            number_of_rows: 1,
            number_of_columns: 1,
            base_data_store: tst::DataStore {
                tiles: tst::TileStorage {
                    tiles: vec![tst::tile_storage::Tile {
                        tileid: 0,
                        tile: reference(6),
                    }],
                    tile_size: Some(256),
                    ..Default::default()
                },
                string_table: reference(5),
                formula_table: reference(5),
                ..Default::default()
            },
            ..Default::default()
        };
        let tile = tst::Tile {
            numrows: 1,
            row_infos: vec![tst::TileRowInfo {
                tile_row_index: 0,
                cell_count: 1,
                storage_version: Some(5),
                cell_storage_buffer: Some(formula_cell.encode()),
                cell_offsets: Some(vec![0, 0]),
                ..Default::default()
            }],
            ..Default::default()
        };
        let formula_owner = tsce::FormulaOwnerDependenciesArchive {
            formula_owner_uid: owner_uuid,
            internal_formula_owner_id: 1,
            formula_owner: Some(reference(3)),
            ..Default::default()
        };

        package_bytes_from_archive(Archive {
            objects: vec![
                object(
                    1,
                    DOCUMENT_MESSAGE_TYPE,
                    tn::DocumentArchive {
                        sheets: vec![reference(2)],
                        ..Default::default()
                    }
                    .encode_to_vec(),
                )?,
                object(
                    2,
                    SHEET_MESSAGE_TYPE,
                    tn::SheetArchive {
                        name: "Sheet".to_owned(),
                        drawable_infos: vec![reference(3)],
                        ..Default::default()
                    }
                    .encode_to_vec(),
                )?,
                object(3, TABLE_INFO_MESSAGE_TYPE, table_info_data)?,
                object(4, TABLE_MODEL_MESSAGE_TYPE, model.encode_to_vec())?,
                sidecars,
                object(6, 6_002, tile.encode_to_vec())?,
                object(7, 4_008, formula_owner.encode_to_vec())?,
            ],
        })
    }

    fn package_bytes_from_archive(archive: Archive) -> Result<Vec<u8>> {
        package_bytes_from_archives([("Index/Document.iwa", archive)])
    }

    fn package_bytes_from_archives(
        archives: impl IntoIterator<Item = (&'static str, Archive)>,
    ) -> Result<Vec<u8>> {
        let mut entries = Vec::new();
        for (name, archive) in archives {
            let iwa = SnappyStream::compress(
                &archive
                    .to_bytes()
                    .map_err(|error| Error::InvalidFormat(error.to_string()))?,
            )
            .map_err(|error| Error::InvalidFormat(error.to_string()))?;
            entries.push((name, iwa));
        }
        litchi_iwa_archive::package::to_bytes(
            entries.iter().map(|(name, data)| (*name, data.as_slice())),
            Limits::default(),
        )
        .map_err(|error| Error::InvalidFormat(error.to_string()))
    }

    fn object(identifier: u64, message_type: u32, data: Vec<u8>) -> Result<ArchiveObject> {
        ArchiveObject::new(
            identifier,
            vec![RawMessage {
                type_: message_type,
                data,
            }],
        )
        .map_err(|error| Error::InvalidFormat(error.to_string()))
    }

    #[test]
    fn parses_a_minimal_package_into_shared_empty_semantics() -> Result<()> {
        fn assert_send_sync<T: Send + Sync>() {}
        assert_send_sync::<Package>();

        let package = Package::from_bytes(&package_bytes(&tn::DocumentArchive::default())?)?;
        let snapshot = package.snapshot();

        assert_eq!(package.object_count(), 1);
        assert!(package.sheets().is_empty());
        assert!(package.shared_sheets().is_empty());
        assert!(package.document().is_empty());
        assert!(package.document_snapshot().is_empty());
        assert_eq!(package.read_options(), ReadOptions::default());
        assert_eq!(package.text()?, "");
        assert!(Arc::ptr_eq(&package.state, &snapshot.state));
        Ok(())
    }

    #[test]
    fn direct_package_requires_unambiguous_numbers_application_ownership() -> Result<()> {
        let mut pages = tp::DocumentArchive {
            super_: shared_document(),
            ..Default::default()
        };
        pages.floating_drawables = Some(reference(7));
        let keynote = kn::DocumentArchive {
            show: reference(8),
            super_: shared_document(),
            ..Default::default()
        };
        for foreign_payload in [pages.encode_to_vec(), keynote.encode_to_vec()] {
            let bytes = package_bytes_from_archive(Archive {
                objects: vec![object(1, DOCUMENT_MESSAGE_TYPE, foreign_payload)?],
            })?;
            assert!(matches!(
                Package::from_bytes(&bytes),
                Err(Error::NotNumbers)
            ));
            assert!(matches!(
                compatibility_tables_from_bytes(&bytes),
                Err(Error::NotNumbers)
            ));
        }

        let native_pages = package_bytes_from_archive(Archive {
            objects: vec![object(1, 10_000, pages.encode_to_vec())?],
        })?;
        assert!(matches!(
            Package::from_bytes(&native_pages),
            Err(Error::NotNumbers)
        ));
        assert!(matches!(
            compatibility_tables_from_bytes(&native_pages),
            Err(Error::NotNumbers)
        ));

        let unknown = package_bytes_from_archive(Archive {
            objects: vec![object(1, DOCUMENT_MESSAGE_TYPE, vec![0x08, 0x01])?],
        })?;
        assert!(matches!(
            Package::from_bytes(&unknown),
            Err(Error::InvalidFormat(_))
        ));

        let mut mixed_payload = tn::DocumentArchive::default().encode_to_vec();
        mixed_payload.extend_from_slice(&pages.encode_to_vec());
        let mixed_root = object(1, DOCUMENT_MESSAGE_TYPE, mixed_payload)?;
        let mixed = package_bytes_from_archive(Archive {
            objects: vec![mixed_root],
        })?;
        assert!(matches!(
            Package::from_bytes(&mixed),
            Err(Error::InvalidFormat(_))
        ));

        let unrelated_sibling = ArchiveObject::new(
            1,
            vec![
                RawMessage {
                    type_: DOCUMENT_MESSAGE_TYPE,
                    data: tn::DocumentArchive::default().encode_to_vec(),
                },
                RawMessage {
                    type_: 10_000,
                    data: pages.encode_to_vec(),
                },
            ],
        )
        .map_err(|error| Error::InvalidFormat(error.to_string()))?;
        let unrelated = package_bytes_from_archive(Archive {
            objects: vec![unrelated_sibling],
        })?;
        assert!(Package::from_bytes(&unrelated).is_ok());

        #[allow(deprecated, reason = "Regression for a supported native field")]
        let numbers_with_calculation_engine = tn::DocumentArchive {
            calculation_engine: Some(reference(7)),
            ..Default::default()
        };
        let calculation_engine = package_bytes(&numbers_with_calculation_engine)?;
        assert!(Package::from_bytes(&calculation_engine).is_ok());
        assert!(compatibility_tables_from_bytes(&calculation_engine)?.is_empty());

        let mut numbers_with_scalar_collisions = tn::DocumentArchive::default().encode_to_vec();
        numbers_with_scalar_collisions.extend_from_slice(&[0x10, 0x01, 0x78, 0x01]);
        let scalar_collisions = package_bytes_from_archive(Archive {
            objects: vec![object(
                1,
                DOCUMENT_MESSAGE_TYPE,
                numbers_with_scalar_collisions,
            )?],
        })?;
        assert!(Package::from_bytes(&scalar_collisions).is_ok());

        let noncanonical_numbers = package_bytes_from_archive(Archive {
            objects: vec![object(
                1,
                TABLE_INFO_MESSAGE_TYPE,
                tn::DocumentArchive::default().encode_to_vec(),
            )?],
        })?;
        assert!(matches!(
            compatibility_tables_from_bytes(&noncanonical_numbers),
            Err(Error::InvalidFormat(_))
        ));

        let masked_unknown = ArchiveObject::new(
            1,
            vec![
                RawMessage {
                    type_: DOCUMENT_MESSAGE_TYPE,
                    data: vec![0x08, 0x01],
                },
                RawMessage {
                    type_: TABLE_INFO_MESSAGE_TYPE,
                    data: tn::DocumentArchive::default().encode_to_vec(),
                },
            ],
        )
        .map_err(|error| Error::InvalidFormat(error.to_string()))?;
        let masked = package_bytes_from_archive(Archive {
            objects: vec![masked_unknown],
        })?;
        assert!(matches!(
            compatibility_tables_from_bytes(&masked),
            Err(Error::InvalidFormat(_))
        ));
        Ok(())
    }

    #[test]
    fn compatibility_projection_does_not_require_rooted_sheet_construction() -> Result<()> {
        let root = tn::DocumentArchive {
            sheets: vec![reference(2)],
            ..Default::default()
        };
        let bytes = package_bytes(&root)?;

        assert!(matches!(
            Package::from_bytes(&bytes),
            Err(Error::InvalidFormat(_))
        ));
        assert!(compatibility_tables_from_bytes(&bytes)?.is_empty());
        Ok(())
    }

    #[test]
    fn formula_reference_enrichment_is_not_built_without_formula_tables() -> Result<()> {
        let group = tst::group_by_archive::GroupNodeArchive {
            child: vec![tst::group_by_archive::GroupNodeArchive::default()],
            ..Default::default()
        };
        let bytes = package_bytes_from_archive(Archive {
            objects: vec![
                object(
                    1,
                    DOCUMENT_MESSAGE_TYPE,
                    tn::DocumentArchive::default().encode_to_vec(),
                )?,
                object(2, 6_383, group.encode_to_vec())?,
            ],
        })?;
        let semantic = SemanticLimits::new(2, crate::MAX_SHEETS, crate::MAX_TABLES, 1)
            .map_err(|error| Error::InvalidFormat(error.to_string()))?;

        let package = Package::from_bytes_with_options(
            &bytes,
            ReadOptions::new(Limits::default(), semantic),
        )?;
        assert!(package.document().is_empty());
        Ok(())
    }

    #[test]
    fn formula_reference_enrichment_uses_table_info_projection_and_is_best_effort() -> Result<()> {
        let valid_table_info = tst::TableInfoArchive {
            table_model: reference(4),
            ..Default::default()
        }
        .encode_to_vec();
        let valid =
            compatibility_tables_from_bytes(&package_with_formula_reference(valid_table_info)?)?;
        assert!(matches!(
            valid.first().and_then(|table| table.get_a1("A1").ok().flatten()),
            Some(crate::cell::Value::Formula(formula)) if formula == "=Sheet::Table::A1"
        ));

        // Canonical field 2 still points at the model, but required field 1
        // is absent. Formula-name discovery must ignore this malformed
        // best-effort metadata rather than aborting the table projection.
        let malformed_table_info = vec![0x12, 0x02, 0x08, 0x04];
        let malformed = compatibility_tables_from_bytes(&package_with_formula_reference(
            malformed_table_info,
        )?)?;
        assert!(matches!(
            malformed
                .first()
                .and_then(|table| table.get_a1("A1").ok().flatten()),
            Some(crate::cell::Value::Formula(formula)) if formula == "=Table::A1"
        ));
        Ok(())
    }

    #[test]
    fn compact_index_rejects_the_null_object_identifier() -> Result<()> {
        let bytes = package_bytes_from_archive(Archive {
            objects: vec![
                object(
                    1,
                    DOCUMENT_MESSAGE_TYPE,
                    tn::DocumentArchive::default().encode_to_vec(),
                )?,
                object(0, 99_999, Vec::new())?,
            ],
        })?;
        assert!(matches!(
            Package::from_bytes(&bytes),
            Err(Error::InvalidFormat(_))
        ));
        assert!(matches!(
            compatibility_tables_from_bytes(&bytes),
            Err(Error::InvalidFormat(_))
        ));
        Ok(())
    }

    #[test]
    fn legacy_primary_candidates_cannot_promote_secondary_canonical_payloads() -> Result<()> {
        let candidate = ArchiveObject::new(
            2,
            vec![
                RawMessage {
                    type_: 6_000,
                    data: tst::TableInfoArchive::default().encode_to_vec(),
                },
                RawMessage {
                    type_: TABLE_MODEL_MESSAGE_TYPE,
                    data: tst::TableModelArchive::default().encode_to_vec(),
                },
            ],
        )
        .map_err(|error| Error::InvalidFormat(error.to_string()))?;
        let bytes = package_bytes_from_archive(Archive {
            objects: vec![
                object(
                    1,
                    DOCUMENT_MESSAGE_TYPE,
                    tn::DocumentArchive::default().encode_to_vec(),
                )?,
                candidate,
            ],
        })?;

        assert!(compatibility_tables_from_bytes(&bytes)?.is_empty());
        Ok(())
    }

    #[test]
    fn duplicate_canonical_candidate_payloads_fail_closed() -> Result<()> {
        let candidate = ArchiveObject::new(
            2,
            vec![
                RawMessage {
                    type_: TABLE_MODEL_MESSAGE_TYPE,
                    data: tst::TableModelArchive::default().encode_to_vec(),
                },
                RawMessage {
                    type_: TABLE_MODEL_MESSAGE_TYPE,
                    data: tst::TableModelArchive::default().encode_to_vec(),
                },
            ],
        )
        .map_err(|error| Error::InvalidFormat(error.to_string()))?;
        let bytes = package_bytes_from_archive(Archive {
            objects: vec![
                object(
                    1,
                    DOCUMENT_MESSAGE_TYPE,
                    tn::DocumentArchive::default().encode_to_vec(),
                )?,
                candidate,
            ],
        })?;

        assert!(matches!(
            compatibility_tables_from_bytes(&bytes),
            Err(Error::InvalidFormat(_))
        ));
        Ok(())
    }

    #[test]
    fn duplicate_legacy_candidate_payloads_fail_closed() -> Result<()> {
        let candidate = ArchiveObject::new(
            2,
            vec![
                RawMessage {
                    type_: TABLE_INFO_MESSAGE_TYPE,
                    data: tst::TableInfoArchive::default().encode_to_vec(),
                },
                RawMessage {
                    type_: TABLE_INFO_MESSAGE_TYPE,
                    data: tst::TableInfoArchive::default().encode_to_vec(),
                },
            ],
        )
        .map_err(|error| Error::InvalidFormat(error.to_string()))?;
        let bytes = package_bytes_from_archive(Archive {
            objects: vec![
                object(
                    1,
                    DOCUMENT_MESSAGE_TYPE,
                    tn::DocumentArchive::default().encode_to_vec(),
                )?,
                candidate,
            ],
        })?;

        assert!(matches!(
            compatibility_tables_from_bytes(&bytes),
            Err(Error::InvalidFormat(_))
        ));
        Ok(())
    }

    #[test]
    fn path_reader_bounds_descriptor_growth_and_reported_length() -> Result<()> {
        let defaults = Limits::default();
        let limits = Limits::new(
            4,
            defaults.max_entries(),
            defaults.max_entry_bytes(),
            defaults.max_total_bytes(),
            defaults.max_iwa_stream_bytes(),
        )?;

        let mut exact = io::Cursor::new([1, 2, 3, 4]);
        assert_eq!(
            read_source_with_reported_length(&mut exact, 4, limits)?,
            vec![1, 2, 3, 4]
        );

        let mut grew = io::Cursor::new([1, 2, 3, 4, 5]);
        assert!(matches!(
            read_source_with_reported_length(&mut grew, 4, limits),
            Err(Error::InputTooLarge {
                observed: 5,
                maximum: 4,
            })
        ));

        let mut oversized = io::Cursor::new([1, 2, 3, 4, 5]);
        assert!(matches!(
            read_source_with_reported_length(&mut oversized, 5, limits),
            Err(Error::InputTooLarge {
                observed: 5,
                maximum: 4,
            })
        ));

        #[cfg(unix)]
        assert!(matches!(
            read_source(Path::new(env!("CARGO_MANIFEST_DIR")), limits),
            Err(Error::InvalidFormat(_))
        ));

        let mut changed = tempfile::NamedTempFile::new()?;
        changed.write_all(&[1])?;
        changed.flush()?;
        let before = changed.as_file().metadata()?;
        changed.write_all(&[2])?;
        changed.flush()?;
        let after = changed.as_file().metadata()?;
        assert!(matches!(
            ensure_source_unchanged(&before, &after),
            Err(Error::InvalidFormat(_))
        ));

        let version_before = changed.as_file().metadata()?;
        let changed_time = version_before
            .modified()?
            .checked_add(std::time::Duration::from_secs(1))
            .ok_or_else(|| Error::InvalidFormat("test timestamp overflow".to_owned()))?;
        changed
            .as_file()
            .set_times(std::fs::FileTimes::new().set_modified(changed_time))?;
        let version_after = changed.as_file().metadata()?;
        assert_eq!(version_before.len(), version_after.len());
        assert!(matches!(
            ensure_source_unchanged(&version_before, &version_after),
            Err(Error::InvalidFormat(_))
        ));
        Ok(())
    }

    #[test]
    fn root_requires_one_canonical_document_payload() -> Result<()> {
        let wrong_type = package_bytes_from_archive(Archive {
            objects: vec![object(
                1,
                TABLE_INFO_MESSAGE_TYPE,
                tn::DocumentArchive::default().encode_to_vec(),
            )?],
        })?;
        assert!(matches!(
            Package::from_bytes(&wrong_type),
            Err(Error::InvalidFormat(_))
        ));

        let duplicate_object = ArchiveObject::new(
            1,
            vec![
                RawMessage {
                    type_: DOCUMENT_MESSAGE_TYPE,
                    data: tn::DocumentArchive::default().encode_to_vec(),
                },
                RawMessage {
                    type_: DOCUMENT_MESSAGE_TYPE,
                    data: tn::DocumentArchive::default().encode_to_vec(),
                },
            ],
        )
        .map_err(|error| Error::InvalidFormat(error.to_string()))?;
        let duplicate_bytes = package_bytes_from_archive(Archive {
            objects: vec![duplicate_object],
        })?;
        assert!(matches!(
            Package::from_bytes(&duplicate_bytes),
            Err(Error::InvalidFormat(_))
        ));
        assert!(matches!(
            compatibility_tables_from_bytes(&duplicate_bytes),
            Err(Error::InvalidFormat(_))
        ));
        Ok(())
    }

    #[test]
    fn object_budget_is_checked_before_index_allocation() -> Result<()> {
        let bytes = package_bytes_from_archive(Archive {
            objects: vec![
                object(
                    1,
                    DOCUMENT_MESSAGE_TYPE,
                    tn::DocumentArchive::default().encode_to_vec(),
                )?,
                object(2, 99_999, Vec::new())?,
            ],
        })?;
        let exact = SemanticLimits::new(2, crate::MAX_SHEETS, crate::MAX_TABLES, MAX_REFERENCES)
            .map_err(|error| Error::InvalidFormat(error.to_string()))?;
        let exact_options = ReadOptions::new(Limits::default(), exact);
        let exact_package = Package::from_bytes_with_options(&bytes, exact_options)?;
        assert_eq!(exact_package.object_count(), 2);
        assert_eq!(exact_package.read_options(), exact_options);

        let exceeded = SemanticLimits::new(1, crate::MAX_SHEETS, crate::MAX_TABLES, MAX_REFERENCES)
            .map_err(|error| Error::InvalidFormat(error.to_string()))?;
        assert!(matches!(
            Package::from_bytes_with_options(&bytes, ReadOptions::new(Limits::default(), exceeded)),
            Err(Error::SemanticLimit {
                kind: SemanticLimitKind::Objects,
                observed: 2,
                maximum: 1,
                path: SemanticPath::Package,
            })
        ));
        Ok(())
    }

    #[test]
    fn compatibility_cell_and_text_budgets_span_all_tables() -> Result<()> {
        let bytes = package_with_two_materialized_empty_tables()?;
        let limits = |cells, text| {
            SemanticLimits::default()
                .with_projection_limits(cells, text)
                .map_err(|error| Error::InvalidFormat(error.to_string()))
        };

        let exact = compatibility_tables_from_bytes_with_options(
            &bytes,
            ReadOptions::new(Limits::default(), limits(2, 2)?),
        )?;
        assert_eq!(exact.len(), 2);
        assert_eq!(exact.iter().map(crate::Table::cell_count).sum::<usize>(), 2);

        assert!(matches!(
            compatibility_tables_from_bytes_with_options(
                &bytes,
                ReadOptions::new(Limits::default(), limits(1, 2)?),
            ),
            Err(Error::SemanticLimit {
                kind: SemanticLimitKind::MaterializedCells,
                observed: 2,
                maximum: 1,
                ..
            })
        ));
        assert!(matches!(
            compatibility_tables_from_bytes_with_options(
                &bytes,
                ReadOptions::new(Limits::default(), limits(2, 1)?),
            ),
            Err(Error::SemanticLimit {
                kind: SemanticLimitKind::OutputTextBytes,
                observed: 2,
                maximum: 1,
                ..
            })
        ));
        Ok(())
    }

    #[test]
    fn rooted_reference_budget_is_inclusive() -> Result<()> {
        let root = tn::DocumentArchive {
            sheets: vec![litchi_iwa_protos::tsp::Reference {
                identifier: 2,
                ..Default::default()
            }],
            ..Default::default()
        };
        let sheet = tn::SheetArchive {
            name: "Budgeted".to_owned(),
            drawable_infos: vec![
                litchi_iwa_protos::tsp::Reference {
                    identifier: 3,
                    ..Default::default()
                },
                litchi_iwa_protos::tsp::Reference {
                    identifier: 4,
                    ..Default::default()
                },
            ],
            ..Default::default()
        };
        let bytes = package_bytes_from_archive(Archive {
            objects: vec![
                object(1, DOCUMENT_MESSAGE_TYPE, root.encode_to_vec())?,
                object(2, SHEET_MESSAGE_TYPE, sheet.encode_to_vec())?,
                object(3, 99_998, Vec::new())?,
                object(4, 99_999, Vec::new())?,
            ],
        })?;
        let limits = |max_references| {
            SemanticLimits::new(4, 1, crate::MAX_TABLES, max_references)
                .map_err(|error| Error::InvalidFormat(error.to_string()))
        };

        let exact = Package::from_bytes_with_options(
            &bytes,
            ReadOptions::new(Limits::default(), limits(3)?),
        )?;
        assert_eq!(exact.sheets().len(), 1);
        assert_eq!(exact.sheets()[0].tables().len(), 0);

        let exceeded = Package::from_bytes_with_options(
            &bytes,
            ReadOptions::new(Limits::default(), limits(2)?),
        );
        assert!(matches!(
            exceeded,
            Err(Error::SemanticLimit {
                kind: SemanticLimitKind::References,
                observed: 3,
                maximum: 2,
                path: SemanticPath::Sheet { index: 0 },
            })
        ));
        Ok(())
    }

    #[test]
    fn rooted_sheet_budget_is_inclusive() -> Result<()> {
        let root = tn::DocumentArchive {
            sheets: vec![
                litchi_iwa_protos::tsp::Reference {
                    identifier: 2,
                    ..Default::default()
                },
                litchi_iwa_protos::tsp::Reference {
                    identifier: 3,
                    ..Default::default()
                },
            ],
            ..Default::default()
        };
        let bytes = package_bytes_from_archive(Archive {
            objects: vec![
                object(1, DOCUMENT_MESSAGE_TYPE, root.encode_to_vec())?,
                object(
                    2,
                    SHEET_MESSAGE_TYPE,
                    tn::SheetArchive {
                        name: "First".to_owned(),
                        ..Default::default()
                    }
                    .encode_to_vec(),
                )?,
                object(
                    3,
                    SHEET_MESSAGE_TYPE,
                    tn::SheetArchive {
                        name: "Second".to_owned(),
                        ..Default::default()
                    }
                    .encode_to_vec(),
                )?,
            ],
        })?;
        let limits = |max_sheets| {
            SemanticLimits::new(3, max_sheets, crate::MAX_TABLES, 2)
                .map_err(|error| Error::InvalidFormat(error.to_string()))
        };

        let exact = Package::from_bytes_with_options(
            &bytes,
            ReadOptions::new(Limits::default(), limits(2)?),
        )?;
        assert_eq!(exact.sheets().len(), 2);

        let exceeded = Package::from_bytes_with_options(
            &bytes,
            ReadOptions::new(Limits::default(), limits(1)?),
        );
        assert!(matches!(
            exceeded,
            Err(Error::SemanticLimit {
                kind: SemanticLimitKind::Sheets,
                observed: 2,
                maximum: 1,
                path: SemanticPath::Document,
            })
        ));
        Ok(())
    }

    #[test]
    fn rooted_duplicate_model_precedes_an_exhausted_table_budget() -> Result<()> {
        let bytes = rooted_two_table_package(5, Vec::new())?;
        let semantic = SemanticLimits::new(7, 1, 1, 5)
            .map_err(|error| Error::InvalidFormat(error.to_string()))?;

        let error =
            Package::from_bytes_with_options(&bytes, ReadOptions::new(Limits::default(), semantic))
                .expect_err("the second drawable must not reuse the first table model");
        assert!(matches!(
            error,
            Error::InvalidFormat(message)
                if message
                    == "Numbers sheet 0 drawable 1 reuses a table model owned by an earlier drawable"
        ));
        Ok(())
    }

    #[test]
    fn rooted_table_budget_is_charged_before_model_decoding() -> Result<()> {
        let bytes = rooted_two_table_package(6, vec![0xff])?;
        let semantic = SemanticLimits::new(7, 1, 1, 5)
            .map_err(|error| Error::InvalidFormat(error.to_string()))?;

        assert!(matches!(
            Package::from_bytes_with_options(&bytes, ReadOptions::new(Limits::default(), semantic),),
            Err(Error::SemanticLimit {
                kind: SemanticLimitKind::Tables,
                observed: 2,
                maximum: 1,
                path: SemanticPath::Drawable { sheet: 0, index: 1 },
            })
        ));
        Ok(())
    }

    #[test]
    fn structured_candidates_use_only_the_primary_message_type() -> Result<()> {
        let secondary = ArchiveObject::new(
            2,
            vec![
                RawMessage {
                    type_: 99_999,
                    data: Vec::new(),
                },
                RawMessage {
                    type_: 6_001,
                    data: vec![0xff],
                },
            ],
        )
        .map_err(|error| Error::InvalidFormat(error.to_string()))?;
        let bytes = package_bytes_from_archive(Archive {
            objects: vec![
                object(
                    1,
                    DOCUMENT_MESSAGE_TYPE,
                    tn::DocumentArchive::default().encode_to_vec(),
                )?,
                secondary,
            ],
        })?;
        let package = Package::from_bytes(&bytes)?;
        assert!(package.extract_structured_tables()?.is_empty());
        Ok(())
    }

    #[test]
    fn malformed_primary_table_model_is_not_silently_skipped() -> Result<()> {
        let bytes = package_bytes_from_archive(Archive {
            objects: vec![
                object(
                    1,
                    DOCUMENT_MESSAGE_TYPE,
                    tn::DocumentArchive::default().encode_to_vec(),
                )?,
                object(2, 6_001, vec![0xff])?,
            ],
        })?;
        let package = Package::from_bytes(&bytes)?;
        assert!(matches!(
            package.extract_structured_tables(),
            Err(Error::InvalidFormat(_))
        ));
        Ok(())
    }

    #[test]
    fn valid_table_info_bytes_under_an_unrelated_type_are_not_tables() -> Result<()> {
        let root = tn::DocumentArchive {
            sheets: vec![litchi_iwa_protos::tsp::Reference {
                identifier: 2,
                ..Default::default()
            }],
            ..Default::default()
        };
        let sheet = tn::SheetArchive {
            name: "Strict".to_owned(),
            drawable_infos: vec![litchi_iwa_protos::tsp::Reference {
                identifier: 3,
                ..Default::default()
            }],
            ..Default::default()
        };
        let bytes = package_bytes_from_archive(Archive {
            objects: vec![
                object(1, DOCUMENT_MESSAGE_TYPE, root.encode_to_vec())?,
                object(2, SHEET_MESSAGE_TYPE, sheet.encode_to_vec())?,
                object(3, 99_999, tst::TableInfoArchive::default().encode_to_vec())?,
            ],
        })?;
        let package = Package::from_bytes(&bytes)?;
        assert_eq!(package.sheets().len(), 1);
        assert_eq!(package.sheets()[0].tables().len(), 0);
        Ok(())
    }

    #[test]
    fn rooted_malformed_table_info_reports_a_content_free_location() -> Result<()> {
        let root = tn::DocumentArchive {
            sheets: vec![reference(2)],
            ..Default::default()
        };
        let sheet = tn::SheetArchive {
            name: "Strict".to_owned(),
            drawable_infos: vec![reference(3)],
            ..Default::default()
        };
        let bytes = package_bytes_from_archive(Archive {
            objects: vec![
                object(1, DOCUMENT_MESSAGE_TYPE, root.encode_to_vec())?,
                object(2, SHEET_MESSAGE_TYPE, sheet.encode_to_vec())?,
                // `super` is valid opaque framing, but field 2 is absent.
                object(3, TABLE_INFO_MESSAGE_TYPE, vec![0x0a, 0x00])?,
            ],
        })?;
        assert!(matches!(
            Package::from_bytes(&bytes),
            Err(Error::MalformedPayload {
                path: SemanticPath::Drawable { sheet: 0, index: 0 },
            })
        ));
        Ok(())
    }

    #[test]
    fn rooted_sheet_ownership_rejects_duplicates_and_ambiguous_table_info() -> Result<()> {
        let duplicated_root = tn::DocumentArchive {
            sheets: vec![
                litchi_iwa_protos::tsp::Reference {
                    identifier: 2,
                    ..Default::default()
                },
                litchi_iwa_protos::tsp::Reference {
                    identifier: 2,
                    ..Default::default()
                },
            ],
            ..Default::default()
        };
        let empty_sheet = tn::SheetArchive {
            name: "Strict".to_owned(),
            ..Default::default()
        };
        let duplicate = package_bytes_from_archive(Archive {
            objects: vec![
                object(1, DOCUMENT_MESSAGE_TYPE, duplicated_root.encode_to_vec())?,
                object(2, SHEET_MESSAGE_TYPE, empty_sheet.encode_to_vec())?,
            ],
        })?;
        assert!(matches!(
            Package::from_bytes(&duplicate),
            Err(Error::InvalidFormat(_))
        ));

        let root = tn::DocumentArchive {
            sheets: vec![litchi_iwa_protos::tsp::Reference {
                identifier: 2,
                ..Default::default()
            }],
            ..Default::default()
        };
        let sheet = tn::SheetArchive {
            name: "Strict".to_owned(),
            drawable_infos: vec![litchi_iwa_protos::tsp::Reference {
                identifier: 3,
                ..Default::default()
            }],
            ..Default::default()
        };
        let info = tst::TableInfoArchive::default().encode_to_vec();
        let ambiguous_object = ArchiveObject::new(
            3,
            vec![
                RawMessage {
                    type_: TABLE_INFO_MESSAGE_TYPE,
                    data: info.clone(),
                },
                RawMessage {
                    type_: LEGACY_TABLE_INFO_MESSAGE_TYPE,
                    data: info,
                },
            ],
        )
        .map_err(|error| Error::InvalidFormat(error.to_string()))?;
        let ambiguous_bytes = package_bytes_from_archive(Archive {
            objects: vec![
                object(1, DOCUMENT_MESSAGE_TYPE, root.encode_to_vec())?,
                object(2, SHEET_MESSAGE_TYPE, sheet.encode_to_vec())?,
                ambiguous_object,
            ],
        })?;
        assert!(matches!(
            Package::from_bytes(&ambiguous_bytes),
            Err(Error::InvalidFormat(_))
        ));
        Ok(())
    }

    #[test]
    fn duplicate_identities_across_components_fail_before_lookup() -> Result<()> {
        let bytes = package_bytes_from_archives([
            (
                "Index/Document.iwa",
                Archive {
                    objects: vec![object(
                        1,
                        DOCUMENT_MESSAGE_TYPE,
                        tn::DocumentArchive::default().encode_to_vec(),
                    )?],
                },
            ),
            (
                "Index/Other.iwa",
                Archive {
                    objects: vec![object(1, 99_999, Vec::new())?],
                },
            ),
        ])?;
        assert!(matches!(
            Package::from_bytes(&bytes),
            Err(Error::InvalidFormat(_))
        ));
        Ok(())
    }

    #[test]
    fn malformed_legacy_table_candidates_remain_ignorable_false_positives() -> Result<()> {
        let bytes = package_bytes_from_archive(Archive {
            objects: vec![
                object(
                    1,
                    DOCUMENT_MESSAGE_TYPE,
                    tn::DocumentArchive::default().encode_to_vec(),
                )?,
                object(2, 6_000, vec![0xff])?,
            ],
        })?;
        let package = Package::from_bytes(&bytes)?;
        assert!(package.extract_structured_tables()?.is_empty());
        Ok(())
    }
}
