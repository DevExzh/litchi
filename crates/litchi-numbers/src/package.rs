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
#[allow(
    dead_code,
    reason = "Decoded sheets expose only the construction path used at package ingress."
)]
mod sheet;
#[allow(
    dead_code,
    reason = "Private native tables retain sidecar helpers while the public surface exposes only semantic tables."
)]
mod table;

use std::collections::HashSet;
use std::fmt;
use std::fs::File;
use std::io::{self, Read};
use std::path::Path;
use std::sync::Arc;

use litchi_iwa_archive::ComponentCatalog;
use litchi_iwa_core::{Archive, RawMessage};
use litchi_iwa_protos::{tn, tst, tswp};
use prost::Message;
use thiserror::Error;

use crate::{
    DEFAULT_MAX_TEXT_BYTES, Document, DocumentError, DocumentLimits, MAX_MATERIALIZED_CELLS, Sheet,
};
use extractor::TableDataExtractor;
use index::{Index, Resolved};
use sheet::DecodedSheet;

pub use limits::{
    MAX_OBJECTS, MAX_REFERENCES, ReadOptions, SemanticLimitKind, SemanticLimits,
    SemanticLimitsError,
};
/// Physical ingress ceilings for a parsed Numbers package.
pub use litchi_iwa_archive::Limits;

const DOCUMENT_MESSAGE_TYPE: u32 = 1;
const SHEET_MESSAGE_TYPE: u32 = 2;
const FORM_BASED_SHEET_MESSAGE_TYPE: u32 = 3;
const TABLE_INFO_MESSAGE_TYPE: u32 = 6_000;
const LEGACY_TABLE_INFO_MESSAGE_TYPE: u32 = 6_003;

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
    /// A native protobuf payload could not be decoded.
    #[error("could not decode Numbers protobuf payload: {0}")]
    Protobuf(String),
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
        /// Configured maximum.
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

impl Error {
    fn protobuf(error: &prost::DecodeError) -> Self {
        Self::Protobuf(error.to_string())
    }
}

impl From<crate::cell::wire::Error> for Error {
    fn from(error: crate::cell::wire::Error) -> Self {
        match error {
            crate::cell::wire::Error::InvalidFormat(message) => Self::InvalidFormat(message),
            crate::cell::wire::Error::ParseError(message) => Self::ParseError(message),
        }
    }
}

impl From<litchi_iwa_common::comment::Error> for Error {
    fn from(error: litchi_iwa_common::comment::Error) -> Self {
        Self::InvalidFormat(error.to_string())
    }
}

/// Result returned by native Numbers package operations.
pub type Result<T> = std::result::Result<T, Error>;

#[derive(Debug)]
struct Components {
    catalog: ComponentCatalog,
}

impl Components {
    fn from_bytes(bytes: &[u8], limits: Limits) -> Result<Self> {
        Ok(Self {
            catalog: ComponentCatalog::from_bytes_with_limits(bytes, limits)?,
        })
    }

    fn get_archive(&self, name: &str) -> Option<&Archive> {
        self.catalog
            .get(name)
            .map(litchi_iwa_archive::Component::archive)
    }

    fn iter_archives(&self) -> impl Iterator<Item = (&str, &Archive)> {
        self.catalog
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
    /// Returns a typed error before allocating beyond either selected profile,
    /// or when the package cannot become a strict rooted semantic document.
    pub fn open_with_options(path: impl AsRef<Path>, options: ReadOptions) -> Result<Self> {
        let source_path = path.as_ref();
        let limits = options.archive();
        let metadata = std::fs::metadata(source_path)?;
        if metadata.len() > limits.max_input_bytes() {
            return Err(Error::InputTooLarge {
                observed: metadata.len(),
                maximum: limits.max_input_bytes(),
            });
        }

        let maximum = usize::try_from(limits.max_input_bytes()).map_err(|_error| {
            Error::InvalidFormat("Numbers input ceiling does not fit usize".to_owned())
        })?;
        let file = File::open(source_path)?;
        let mut bytes = Vec::new();
        bytes
            .try_reserve_exact(usize::try_from(metadata.len()).unwrap_or(maximum))
            .map_err(|_error| {
                Error::Common(litchi_iwa_common::Error::Allocation {
                    resource: "Numbers package input",
                    amount: usize::try_from(metadata.len()).unwrap_or(maximum),
                })
            })?;
        file.take(limits.max_input_bytes().saturating_add(1))
            .read_to_end(&mut bytes)?;
        if bytes.len() > maximum {
            return Err(Error::InputTooLarge {
                observed: u64::try_from(bytes.len()).unwrap_or(u64::MAX),
                maximum: limits.max_input_bytes(),
            });
        }
        Self::from_bytes_with_options(&bytes, options)
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
        let semantic = options.semantic();
        let index = Index::from_components(&components, semantic.max_objects())?;
        let root = Self::root_document(&components)?;
        let sheets = Self::decode_sheets(&components, &index, &root, semantic)?;
        let document = Document::from_sheets_with_limits(
            sheets,
            DocumentLimits::new(
                semantic.max_sheets(),
                semantic.max_tables(),
                MAX_MATERIALIZED_CELLS,
                DEFAULT_MAX_TEXT_BYTES,
            ),
        )?;
        Ok(Self {
            state: Arc::new(State {
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
        if !TableDataExtractor::has_table_models(&self.state.index) {
            return Ok(Vec::new());
        }
        TableDataExtractor::new(&self.state.components, &self.state.index)
            .extract_all_semantic_tables(self.state.options.semantic().max_tables())
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
            .map_err(|error| Error::protobuf(&error))
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
        let extractor = TableDataExtractor::new(components, index);
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
        let info = tst::TableInfoArchive::decode(message.data.as_slice()).map_err(|error| {
            Error::InvalidFormat(format!(
                "Numbers {path} table-info payload is malformed: {error}"
            ))
        })?;
        budget.charge_references(1, path)?;
        let model_id = info.table_model.identifier;
        let model = index.resolve_ref_id(components, model_id)?.ok_or_else(|| {
            Error::InvalidFormat(format!("Numbers {path} table model is missing"))
        })?;
        if !seen_models.insert(model_id) {
            return Err(Error::InvalidFormat(format!(
                "Numbers {path} reuses a table model owned by an earlier drawable"
            )));
        }
        budget.charge_table(path)?;
        extractor
            .extract_reachable_table_from_object(&model, path)
            .map(Some)
    }
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
            .map_err(|error| Error::protobuf(&error)),
        (None, Some(message)) => tn::FormBasedSheetArchive::decode(message.data.as_slice())
            .map(|form_archive| form_archive.super_)
            .map_err(|error| Error::protobuf(&error)),
        (None, None) => Err(Error::InvalidFormat(format!(
            "Numbers {path} has no canonical sheet payload"
        ))),
    }
}

fn allocation_error(resource: &'static str, amount: usize) -> Error {
    Error::Common(litchi_iwa_common::Error::Allocation { resource, amount })
}

#[cfg(test)]
mod tests {
    use super::*;
    use litchi_iwa_core::{ArchiveObject, RawMessage, SnappyStream};
    use soapberry_zip::office::StreamingArchiveWriter;

    fn package_bytes(root: &tn::DocumentArchive) -> Result<Vec<u8>> {
        package_bytes_from_archive(Archive {
            objects: vec![object(1, DOCUMENT_MESSAGE_TYPE, root.encode_to_vec())?],
        })
    }

    fn package_bytes_from_archive(archive: Archive) -> Result<Vec<u8>> {
        package_bytes_from_archives([("Index/Document.iwa", archive)])
    }

    fn package_bytes_from_archives(
        archives: impl IntoIterator<Item = (&'static str, Archive)>,
    ) -> Result<Vec<u8>> {
        let mut writer = StreamingArchiveWriter::new();
        for (name, archive) in archives {
            let iwa = SnappyStream::compress(
                &archive
                    .to_bytes()
                    .map_err(|error| Error::InvalidFormat(error.to_string()))?,
            )
            .map_err(|error| Error::InvalidFormat(error.to_string()))?;
            writer
                .write_stored(name, &iwa)
                .map_err(|error| Error::InvalidFormat(error.to_string()))?;
        }
        writer
            .finish_to_bytes()
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
