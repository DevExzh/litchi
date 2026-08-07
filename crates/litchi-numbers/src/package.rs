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

use std::fmt;
use std::fs::File;
use std::io::{self, Read};
use std::path::Path;
use std::sync::Arc;

use litchi_iwa_archive::ComponentCatalog;
use litchi_iwa_core::{Archive, ArchiveObject};
use litchi_iwa_protos::{tn, tst, tswp};
use prost::Message;
use thiserror::Error;

use crate::{Document, DocumentError, Sheet};
use extractor::TableDataExtractor;
use index::{Index, Resolved};
use sheet::DecodedSheet;

/// Physical ingress ceilings for a parsed Numbers package.
pub use litchi_iwa_archive::Limits;

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
    fn protobuf(error: prost::DecodeError) -> Self {
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

    fn find_object(&self, identifier: u64) -> Option<&ArchiveObject> {
        self.iter_archives()
            .find_map(|(_name, archive)| archive.object(identifier))
    }

    fn iter_objects(&self) -> impl Iterator<Item = &ArchiveObject> {
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
}

impl Package {
    /// Open a Numbers package from a filesystem path using default limits.
    ///
    /// # Errors
    ///
    /// Returns a typed error when the path cannot be read, the package exceeds
    /// a physical ceiling, or its IWA/Numbers contents are malformed.
    pub fn open(path: impl AsRef<Path>) -> Result<Self> {
        Self::open_with_limits(path, Limits::default())
    }

    /// Open a Numbers package from a filesystem path under explicit limits.
    ///
    /// # Errors
    ///
    /// Returns a typed error before allocating more than the selected input
    /// ceiling, or when the package cannot become a semantic document.
    pub fn open_with_limits(path: impl AsRef<Path>, limits: Limits) -> Result<Self> {
        let path = path.as_ref();
        let metadata = std::fs::metadata(path)?;
        if metadata.len() > limits.max_input_bytes() {
            return Err(Error::InputTooLarge {
                observed: metadata.len(),
                maximum: limits.max_input_bytes(),
            });
        }

        let maximum = usize::try_from(limits.max_input_bytes()).map_err(|_error| {
            Error::InvalidFormat("Numbers input ceiling does not fit usize".to_owned())
        })?;
        let file = File::open(path)?;
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
        Self::from_bytes_with_limits(&bytes, limits)
    }

    /// Parse a Numbers package from an in-memory ZIP payload using defaults.
    ///
    /// # Errors
    ///
    /// Returns a typed error when physical ingress, IWA framing, protobuf
    /// decoding, or semantic construction fails.
    pub fn from_bytes(bytes: &[u8]) -> Result<Self> {
        Self::from_bytes_with_limits(bytes, Limits::default())
    }

    /// Parse a Numbers package from an in-memory ZIP payload under limits.
    ///
    /// # Errors
    ///
    /// Returns a typed error when the package exceeds a selected resource
    /// ceiling or cannot be decoded as a Numbers document.
    pub fn from_bytes_with_limits(bytes: &[u8], limits: Limits) -> Result<Self> {
        let components = Components::from_bytes(bytes, limits)?;
        let index = Index::from_components(&components)?;
        let root = Self::root_document(&components)?;
        let sheets = Self::decode_sheets(&components, &index, &root)?;
        let document = Document::from_sheets(sheets)?;
        Ok(Self {
            state: Arc::new(State {
                components,
                index,
                document,
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
        components
            .get_archive("Index/Document.iwa")
            .and_then(|archive| archive.object(1))
            .and_then(|object| {
                object
                    .messages
                    .iter()
                    .find_map(|message| tn::DocumentArchive::decode(message.data.as_slice()).ok())
            })
            .ok_or_else(|| {
                Error::InvalidFormat("package does not contain a Numbers root document".to_owned())
            })
    }

    fn decode_sheets(
        components: &Components,
        index: &Index,
        document: &tn::DocumentArchive,
    ) -> Result<Vec<Sheet>> {
        let extractor = TableDataExtractor::new(components, index);
        let mut sheets = Vec::new();
        sheets
            .try_reserve(document.sheets.len())
            .map_err(|_error| {
                Error::Common(litchi_iwa_common::Error::Allocation {
                    resource: "Numbers semantic sheets",
                    amount: document.sheets.len(),
                })
            })?;
        for (position, reference) in document.sheets.iter().enumerate() {
            let object = components
                .find_object(reference.identifier)
                .ok_or_else(|| {
                    Error::InvalidFormat(format!(
                        "Numbers document references missing sheet object {}",
                        reference.identifier
                    ))
                })?;
            let archive = object
                .messages
                .iter()
                .find_map(|message| {
                    tn::SheetArchive::decode(message.data.as_slice())
                        .ok()
                        .or_else(|| {
                            tn::FormBasedSheetArchive::decode(message.data.as_slice())
                                .ok()
                                .map(|form| form.super_)
                        })
                })
                .ok_or_else(|| {
                    Error::InvalidFormat(format!(
                        "Numbers sheet object {} has no TN.SheetArchive payload",
                        reference.identifier
                    ))
                })?;
            let mut sheet = DecodedSheet::new(archive.name, position);
            for drawable in archive.drawable_infos {
                let Some(table) =
                    Self::extract_table(components, index, drawable.identifier, &extractor)?
                else {
                    continue;
                };
                sheet.add_table(table);
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
    ) -> Result<Option<table::Table>> {
        let resolved = index
            .resolve_ref_id(components, drawable_id)?
            .ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "Numbers sheet references missing drawable object {drawable_id}"
                ))
            })?;
        for message in resolved.messages {
            let Ok(info) = tst::TableInfoArchive::decode(message.data.as_slice()) else {
                continue;
            };
            let model_id = info.table_model.identifier;
            let Some(model) = index.resolve_ref_id(components, model_id)? else {
                continue;
            };
            if !model.messages.iter().any(|message| {
                (message.type_ == 6000 || message.type_ == 6001)
                    && tst::TableModelArchive::decode(message.data.as_slice()).is_ok()
            }) {
                continue;
            }
            return extractor.extract_table_from_object(&model);
        }
        Ok(None)
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use litchi_iwa_core::{ArchiveObject, RawMessage, SnappyStream};
    use soapberry_zip::office::StreamingArchiveWriter;

    fn package_bytes(root: tn::DocumentArchive) -> Result<Vec<u8>> {
        let archive = Archive {
            objects: vec![
                ArchiveObject::new(
                    1,
                    vec![RawMessage {
                        type_: 6_000,
                        data: root.encode_to_vec(),
                    }],
                )
                .map_err(|error| Error::InvalidFormat(error.to_string()))?,
            ],
        };
        let iwa = SnappyStream::compress(
            &archive
                .to_bytes()
                .map_err(|error| Error::InvalidFormat(error.to_string()))?,
        )
        .map_err(|error| Error::InvalidFormat(error.to_string()))?;
        let mut writer = StreamingArchiveWriter::new();
        writer
            .write_stored("Index/Document.iwa", &iwa)
            .map_err(|error| Error::InvalidFormat(error.to_string()))?;
        writer
            .finish_to_bytes()
            .map_err(|error| Error::InvalidFormat(error.to_string()))
    }

    #[test]
    fn parses_a_minimal_package_into_shared_empty_semantics() -> Result<()> {
        let package = Package::from_bytes(&package_bytes(tn::DocumentArchive::default())?)?;
        let snapshot = package.snapshot();

        assert_eq!(package.object_count(), 1);
        assert!(package.sheets().is_empty());
        assert!(package.shared_sheets().is_empty());
        assert!(package.document().is_empty());
        assert!(package.document_snapshot().is_empty());
        assert_eq!(package.text()?, "");
        assert!(Arc::ptr_eq(&package.state, &snapshot.state));
        Ok(())
    }
}
