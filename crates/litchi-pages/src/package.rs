//! Pages-native package ingress and semantic projection.
//!
//! This module is the only Pages boundary that understands ZIP/IWA packages
//! and generated protobuf messages. It publishes [`Package`] snapshots whose
//! semantic content is represented by the archive-free [`crate::Document`].

use std::collections::BTreeSet;
use std::fmt;
use std::fs;
use std::io::Read;
use std::mem::size_of;
use std::num::NonZeroU64;
use std::path::Path;
use std::sync::Arc;

use litchi_iwa_archive::ComponentCatalog;
use litchi_iwa_archive::package::Catalog;
use litchi_iwa_common::wire::{WireFieldView, WireView};
use litchi_iwa_protos::{tp, tswp};
use litchi_iwa_text::storage::{Run, Storage};
use plist::Value;
use prost::Message;
use thiserror::Error;

use crate::{
    Body, DEFAULT_MAX_TEXT_BYTES, Document, Error as SemanticError, MAX_BODY_STORAGES,
    MAX_SECTIONS, Root, Section, SectionType,
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
    /// Section names exceed the aggregate retained-memory budget.
    #[error("Pages section names require at least {observed} bytes; budget is {limit}")]
    SectionNamesTooLarge {
        /// Minimum bytes required by the names decoded so far.
        observed: usize,
        /// Aggregate semantic metadata budget.
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
/// The package retains native IWA components for validation and future
/// Pages-native capabilities. Its ordinary read API exposes only the
/// immutable semantic [`Document`], never raw object identifiers or protobuf
/// messages.
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
    components: ComponentCatalog,
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

struct StorageAccumulator {
    text: String,
    runs: Vec<Run>,
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
        let bytes = read_path(path.as_ref(), limits)?;
        Self::from_bytes_with_limits(&bytes, limits)
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
        let metadata = Metadata::from_package(bytes, limits)?;
        let components = ComponentCatalog::from_bytes_with_limits(bytes, limits)?;
        let object_count = validate_components(&components)?;
        let root_references = root_references(&components)?;
        let text_limit = effective_text_limit(limits);
        let document = decode_document(&components, root_references, text_limit)?;

        Ok(Self {
            state: Arc::new(State {
                components,
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
        let object_count = validate_components(&self.state.components)?;
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
    fn from_package(bytes: &[u8], limits: Limits) -> PackageResult<Self> {
        let catalog = Catalog::from_bytes_with_limits(bytes, limits)?;
        let mut metadata = Self::default();

        if let Some(data) = metadata_entry(&catalog, "Metadata/Properties.plist")? {
            metadata.apply_properties(data)?;
        }
        if let Some(data) = metadata_entry(&catalog, "Metadata/BuildVersionHistory.plist")? {
            metadata.build_version = parse_build_version(data)?;
        }
        if let Some(data) = metadata_entry(&catalog, "Metadata/DocumentIdentifier")? {
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

fn read_path(path: &Path, limits: Limits) -> PackageResult<Vec<u8>> {
    let metadata = fs::symlink_metadata(path)?;
    if metadata.file_type().is_symlink() {
        return Err(PackageError::InvalidFormat(format!(
            "Pages package source must not be a symbolic link: {}",
            path.display()
        )));
    }
    if !metadata.is_file() {
        return Err(PackageError::InvalidFormat(format!(
            "Pages package source must be a regular file: {}",
            path.display()
        )));
    }
    if metadata.len() > limits.max_input_bytes() {
        return Err(PackageError::InvalidFormat(format!(
            "Pages package input exceeds the {} byte limit",
            limits.max_input_bytes()
        )));
    }

    let capacity = usize::try_from(metadata.len()).map_err(|_error| {
        PackageError::InvalidFormat("Pages package input length does not fit usize".to_owned())
    })?;
    let mut bytes = Vec::new();
    bytes.try_reserve_exact(capacity).map_err(|_error| {
        PackageError::InvalidFormat("could not allocate Pages input".to_owned())
    })?;

    let maximum = limits.max_input_bytes().saturating_add(1);
    fs::File::open(path)?
        .take(maximum)
        .read_to_end(&mut bytes)?;
    let length = u64::try_from(bytes.len()).map_err(|_error| {
        PackageError::InvalidFormat("Pages package input length does not fit u64".to_owned())
    })?;
    if length > limits.max_input_bytes() {
        return Err(PackageError::InvalidFormat(format!(
            "Pages package input exceeds the {} byte limit",
            limits.max_input_bytes()
        )));
    }
    Ok(bytes)
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

fn root_references(components: &ComponentCatalog) -> PackageResult<RootReferences> {
    let component = components.get("Index/Document.iwa").ok_or_else(|| {
        PackageError::InvalidFormat("Pages package does not contain Index/Document.iwa".to_owned())
    })?;
    let object = component
        .archive()
        .object(1)
        .ok_or_else(|| PackageError::InvalidFormat("Pages root object 1 is missing".to_owned()))?;
    let payload = unique_message_payload(&object.messages, 10_000, "Pages root object 1")?;
    let root = tp::DocumentArchive::decode(payload).map_err(|error| {
        PackageError::InvalidFormat(format!("Pages root type-10000 payload is invalid: {error}"))
    })?;
    validate_optional_reference_field(
        payload,
        4,
        root.body_storage.as_ref(),
        "Pages root body-storage reference",
    )?;
    validate_optional_reference_field(
        payload,
        5,
        root.section.as_ref(),
        "Pages root initial-section reference",
    )?;

    Ok(RootReferences {
        body: nonzero_reference(root.body_storage, "Pages root body-storage reference")?,
        initial_section: nonzero_reference(root.section, "Pages root initial-section reference")?,
    })
}

fn decode_document(
    components: &ComponentCatalog,
    root_references: RootReferences,
    max_text_bytes: usize,
) -> PackageResult<Document> {
    if let Some(identifier) = root_references.body {
        let object = find_object(components, identifier.get()).ok_or_else(|| {
            PackageError::InvalidFormat(format!(
                "Pages body storage object {identifier} is missing"
            ))
        })?;
        let (native, payload) = decode_body_storage(&object.messages, identifier, max_text_bytes)?;
        validate_section_table_wire(payload, &native, identifier)?;
        let section_references =
            native_section_references(&native, root_references.initial_section)?;
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
        let storages = extract_storages(components, max_text_bytes)?;
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
    max_text_bytes: usize,
) -> PackageResult<(tswp::StorageArchive, &[u8])> {
    let payload = unique_text_payload(messages, identifier)?;
    preflight_body_wire(payload, identifier, max_text_bytes)?;
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
    max_text_bytes: usize,
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
    for field in table_view.fields().filter(|field| field.number() == 1) {
        validate_wire_field(field, 2, &context)?;
        entry_count = entry_count.checked_add(1).ok_or_else(|| {
            PackageError::InvalidFormat(format!("{context} section count overflows usize"))
        })?;
        if entry_count > MAX_SECTIONS {
            return Err(SemanticError::TooManySections {
                actual: entry_count,
                limit: MAX_SECTIONS,
            }
            .into());
        }
        preflight_section_table_entry(field.payload(), entry_count - 1, body_identifier)?;
    }
    Ok(())
}

fn preflight_section_table_entry(
    payload: &[u8],
    entry_index: usize,
    body_identifier: NonZeroU64,
) -> PackageResult<()> {
    let context = format!("Pages body object {body_identifier} section table entry {entry_index}");
    let view = parse_wire(payload, &context)?;
    let character_index = unique_wire_field(&view, 1, 0, true, &context)?
        .ok_or_else(|| PackageError::InvalidFormat(format!("{context} has no character index")))?;
    let index = decode_canonical_varint(character_index.payload(), &context)?;
    if u32::try_from(index).is_err() {
        return Err(PackageError::InvalidFormat(format!(
            "{context} character index exceeds u32"
        )));
    }
    let section = unique_wire_field(&view, 2, 2, true, &context)?.ok_or_else(|| {
        PackageError::InvalidFormat(format!("{context} has no section reference"))
    })?;
    let identifier = decode_reference_identifier(section.payload(), &context)?;
    if identifier == 0 {
        return Err(PackageError::InvalidFormat(format!(
            "{context} has a zero section reference"
        )));
    }
    Ok(())
}

fn nonzero_reference(
    reference: Option<litchi_iwa_protos::tsp::Reference>,
    context: &str,
) -> PackageResult<Option<NonZeroU64>> {
    reference
        .map(|native_reference| {
            NonZeroU64::new(native_reference.identifier)
                .ok_or_else(|| PackageError::InvalidFormat(format!("{context} is zero")))
        })
        .transpose()
}

fn validate_optional_reference_field(
    payload: &[u8],
    field_number: u32,
    decoded: Option<&litchi_iwa_protos::tsp::Reference>,
    context: &str,
) -> PackageResult<()> {
    let view = parse_wire(payload, context)?;
    let optional_field = unique_wire_field(&view, field_number, 2, false, context)?;
    match (optional_field, decoded) {
        (Some(wire_field), Some(reference)) => {
            let identifier = decode_reference_identifier(wire_field.payload(), context)?;
            if identifier != reference.identifier {
                return Err(PackageError::InvalidFormat(format!(
                    "{context} changed while decoding"
                )));
            }
        },
        (None, None) => {},
        _ => {
            return Err(PackageError::InvalidFormat(format!(
                "{context} presence changed while decoding"
            )));
        },
    }
    Ok(())
}

fn validate_section_table_wire(
    payload: &[u8],
    decoded: &tswp::StorageArchive,
    body_identifier: NonZeroU64,
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
) -> PackageResult<()> {
    let context = format!("Pages body object {body_identifier} section table entry {entry_index}");
    let view = parse_wire(payload, &context)?;
    let character_index = unique_wire_field(&view, 1, 0, true, &context)?
        .ok_or_else(|| PackageError::InvalidFormat(format!("{context} has no character index")))?;
    let wire_index = decode_canonical_varint(character_index.payload(), &context)?;
    if wire_index != u64::from(decoded.character_index) {
        return Err(PackageError::InvalidFormat(format!(
            "{context} character index changed while decoding"
        )));
    }

    let object = unique_wire_field(&view, 2, 2, false, &context)?;
    match (object, decoded.object.as_ref()) {
        (Some(field), Some(reference)) => {
            let identifier = decode_reference_identifier(field.payload(), &context)?;
            if identifier != reference.identifier {
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

fn decode_reference_identifier(payload: &[u8], context: &str) -> PackageResult<u64> {
    let view = parse_wire(payload, context)?;
    let field = unique_wire_field(&view, 1, 0, true, context)?
        .ok_or_else(|| PackageError::InvalidFormat(format!("{context} has no identifier")))?;
    decode_canonical_varint(field.payload(), context)
}

fn decode_canonical_varint(payload: &[u8], context: &str) -> PackageResult<u64> {
    let (value, length) =
        litchi_iwa_common::varint::decode_varint_from_bytes(payload).map_err(|error| {
            PackageError::InvalidFormat(format!("{context} has invalid varint: {error}"))
        })?;
    if length != payload.len() || length != litchi_iwa_common::varint::encoded_len(value) {
        return Err(PackageError::InvalidFormat(format!(
            "{context} has a noncanonical varint"
        )));
    }
    Ok(value)
}

fn native_section_references(
    body: &tswp::StorageArchive,
    initial_section: Option<NonZeroU64>,
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
    if maximum_count > MAX_SECTIONS.saturating_add(1) {
        return Err(SemanticError::TooManySections {
            actual: maximum_count,
            limit: MAX_SECTIONS,
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
    if references.len() > MAX_SECTIONS {
        return Err(SemanticError::TooManySections {
            actual: references.len(),
            limit: MAX_SECTIONS,
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
    max_name_memory_bytes: usize,
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
    let mut retained_bytes = references
        .len()
        .checked_mul(size_of::<Option<Box<str>>>())
        .ok_or(PackageError::SectionNamesTooLarge {
            observed: usize::MAX,
            limit: max_name_memory_bytes,
        })?;
    if retained_bytes > max_name_memory_bytes {
        return Err(PackageError::SectionNamesTooLarge {
            observed: retained_bytes,
            limit: max_name_memory_bytes,
        });
    }

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
                    limit: max_name_memory_bytes,
                })?;
            if retained_bytes > max_name_memory_bytes {
                return Err(PackageError::SectionNamesTooLarge {
                    observed: retained_bytes,
                    limit: max_name_memory_bytes,
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
                if storages.len() == MAX_BODY_STORAGES {
                    return Err(SemanticError::TooManyBodyStorages {
                        actual: storages.len().saturating_add(1),
                        limit: MAX_BODY_STORAGES,
                    }
                    .into());
                }
                text_bytes = text_bytes.checked_add(storage.len()).ok_or({
                    PackageError::Semantic(SemanticError::TextTooLarge {
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
        if matches!(message.type_, 2001 | 2022)
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
            limit: max_text_bytes,
        }));
    }
    Ok(())
}

fn effective_text_limit(limits: Limits) -> usize {
    limits.max_iwa_stream_bytes().min(DEFAULT_MAX_TEXT_BYTES)
}

const fn is_storage_message_type(type_id: u32) -> bool {
    matches!(type_id, 2001..=2014 | 2022)
}

#[cfg(test)]
mod tests {
    use litchi_iwa_core::{Archive, ArchiveObject, RawMessage, SnappyStream};
    use litchi_iwa_protos::tsp::Reference;
    use litchi_iwa_protos::tswp::{ObjectAttributeTable, object_attribute_table::ObjectAttribute};
    use soapberry_zip::office::StreamingArchiveWriter;

    use super::*;

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

        let mut writer = StreamingArchiveWriter::new();
        writer
            .write_stored("Index/Document.iwa", &compressed)
            .map_err(|error| PackageError::InvalidFormat(error.to_string()))?;
        if metadata {
            writer
                .write_stored(
                    "Metadata/Properties.plist",
                    br#"<?xml version="1.0" encoding="UTF-8"?><plist version="1.0"><dict><key>Title</key><string>Report</string><key>Author</key><string>Ada</string><key>Application</key><string>Pages</string><key>revision</key><string>3</string><key>fileFormatVersion</key><string>7</string></dict></plist>"#,
                )
                .map_err(|error| PackageError::InvalidFormat(error.to_string()))?;
            writer
                .write_stored("Metadata/DocumentIdentifier", b"pages-id\n")
                .map_err(|error| PackageError::InvalidFormat(error.to_string()))?;
        }
        writer
            .finish_to_bytes()
            .map_err(|error| PackageError::InvalidFormat(error.to_string()))
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
                .contains("multiple section boundaries")
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
