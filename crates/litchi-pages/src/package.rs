//! Pages-native package ingress and semantic projection.
//!
//! This module is the only Pages boundary that understands ZIP/IWA packages
//! and generated protobuf messages. It publishes [`Package`] snapshots whose
//! semantic content is represented by the archive-free [`crate::Document`].

use std::collections::BTreeSet;
use std::fs;
use std::io::Read;
use std::num::NonZeroU64;
use std::path::Path;
use std::sync::Arc;

use litchi_iwa_archive::ComponentCatalog;
use litchi_iwa_archive::package::Catalog;
use litchi_iwa_protos::{tp, tswp};
use litchi_iwa_text::storage::Storage;
use plist::Value;
use prost::Message;
use thiserror::Error;

use crate::{
    Body, DEFAULT_MAX_TEXT_BYTES, Document, Error as SemanticError, MAX_BODY_STORAGES, Root,
    Section,
};

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
#[derive(Debug, Clone)]
pub struct Package {
    state: Arc<State>,
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
        let body_identifier = root_body_identifier(&components)?;
        let text_limit = effective_text_limit(limits);
        let document = decode_document(&components, body_identifier, text_limit)?;

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

fn root_body_identifier(components: &ComponentCatalog) -> PackageResult<Option<NonZeroU64>> {
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

    root.body_storage
        .map(|reference| {
            NonZeroU64::new(reference.identifier).ok_or_else(|| {
                PackageError::InvalidFormat("Pages root body-storage reference is zero".to_owned())
            })
        })
        .transpose()
}

fn decode_document(
    components: &ComponentCatalog,
    body_identifier: Option<NonZeroU64>,
    max_text_bytes: usize,
) -> PackageResult<Document> {
    let body = if let Some(identifier) = body_identifier {
        let object = find_object(components, identifier.get()).ok_or_else(|| {
            PackageError::InvalidFormat(format!(
                "Pages body storage object {identifier} is missing"
            ))
        })?;
        let storage = decode_body_storage(&object.messages, identifier, max_text_bytes)?;
        Some(Body::with_max_text_bytes(vec![storage], max_text_bytes)?)
    } else {
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
) -> PackageResult<Storage> {
    let payload = unique_text_payload(messages, identifier)?;
    let native = tswp::StorageArchive::decode(payload).map_err(|error| {
        PackageError::InvalidFormat(format!(
            "Pages body object {identifier} text payload is invalid: {error}"
        ))
    })?;
    ensure_text_size(&native, 0, max_text_bytes, identifier)?;
    litchi_iwa_text_wire::from_archive(native).map_err(|error| {
        PackageError::InvalidFormat(format!(
            "Pages body object {identifier} text payload is invalid: {error}"
        ))
    })
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
        if let Some(body) = body {
            let storage = tswp::StorageArchive {
                text: vec![body.to_owned()],
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
    fn input_limit_is_enforced_before_native_projection() {
        let limits = Limits::new(1, 10, 100, 100, 100)
            .unwrap_or_else(|error| panic!("test limits should be valid: {error}"));
        let error = Package::from_bytes_with_limits(&[0, 1], limits)
            .err()
            .unwrap_or_else(|| panic!("oversized input should fail"));
        assert!(error.to_string().contains("limit"));
    }
}
