//! Typed `MS-OFFCRYPTO` `DataSpaces` and IRM metadata.
//!
//! This module validates the structural graph only. `XrML` licenses and
//! protected content remain inert and are never activated, fetched, or
//! decrypted.
//!
//! ```no_run
//! use litchi_crypto::spaces::inspect_bytes;
//!
//! let bytes = std::fs::read("protected.docx")?;
//! if let Some(graph) = inspect_bytes(&bytes)? {
//!     println!("IRM profile: {:?}", graph.irm);
//! }
//! # Ok::<(), Box<dyn std::error::Error>>(())
//! ```

use std::fmt;
use std::io::{Read, Seek, Write};
use std::sync::Arc;

use litchi_cfb::consts::{STGTY_STORAGE, STGTY_STREAM};
use litchi_cfb::{OleError, OleFile, OleWriter};

use super::integrity::{
    self, DOCUMENT_SUMMARY_HASH_STREAM, DOCUMENT_SUMMARY_STREAM, Info, SUMMARY_HASH_STREAM,
    SUMMARY_STREAM,
};
use super::labels::{self, List};
use litchi_ole_common::custom_xml::{Promotion, Store, inspect as inspect_custom_xml};

const HEADER_LENGTH: u32 = 8;
const TRANSFORM_TYPE: u32 = 1;
const EXTENSIBILITY_HEADER_LENGTH: u32 = 4;
const MAX_STREAM_BYTES: usize = 16 * 1024 * 1024;
const MAX_ENTRIES: usize = 65_536;
const MAX_COMPONENTS: usize = MAX_ENTRIES * 8;
const MAX_STRING_BYTES: usize = 1_048_576;
const MAX_XML_DEPTH: usize = 256;

pub const STORAGE: &str = "\u{0006}DataSpaces";
pub const PRIMARY: &str = "\u{0006}Primary";
pub const FEATURE: &str = "Microsoft.Container.DataSpaces";
pub const DRM_ID: &str = "{C73DFACD-061F-43B0-8B64-0C620D2A8B50}";
pub const DRM_NAME: &str = "Microsoft.Metadata.DRMTransform";
pub const LZX_ID: &str = "{86DE7F2B-DDCE-486d-B016-405BBE82B8BC}";
pub const LZX_NAME: &str = "Microsoft.Metadata.CompressionTransform";
pub const ENCRYPTION_ID: &str = "{FF9A3F03-56EF-4613-BDD5-5A41C1D07246}";
pub const ENCRYPTION_NAME: &str = "Microsoft.Container.EncryptionTransform";

#[derive(Debug)]
pub enum Error {
    Invalid(String),
    Ole(OleError),
}

impl fmt::Display for Error {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Invalid(message) => write!(formatter, "invalid DataSpaces structure: {message}"),
            Self::Ole(error) => write!(formatter, "OLE DataSpaces error: {error}"),
        }
    }
}

impl std::error::Error for Error {
    fn source(&self) -> Option<&(dyn std::error::Error + 'static)> {
        match self {
            Self::Invalid(_) => None,
            Self::Ole(error) => Some(error),
        }
    }
}

impl From<OleError> for Error {
    fn from(error: OleError) -> Self {
        Self::Ole(error)
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Version {
    pub major: u16,
    pub minor: u16,
}

impl Version {
    pub const V1_0: Self = Self { major: 1, minor: 0 };
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct VersionInfo {
    pub feature_identifier: String,
    pub reader: Version,
    pub updater: Version,
    pub writer: Version,
}

impl Default for VersionInfo {
    fn default() -> Self {
        Self {
            feature_identifier: FEATURE.to_string(),
            reader: Version::V1_0,
            updater: Version::V1_0,
            writer: Version::V1_0,
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ReferenceKind {
    Stream,
    Storage,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Reference {
    pub kind: ReferenceKind,
    pub component: String,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct MapEntry {
    pub references: Vec<Reference>,
    pub data_space_name: String,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Map {
    pub entries: Vec<MapEntry>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Definition {
    pub transforms: Vec<String>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Header {
    pub transform_id: String,
    pub transform_name: String,
    pub reader: Version,
    pub updater: Version,
    pub writer: Version,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct IrmTransform {
    pub header: Header,
    /// Signed issuance license XML retained verbatim and never interpreted.
    pub publishing_license: Option<String>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct EncryptionTransform {
    pub header: Header,
    /// Null when `EncryptionInfo` is authoritative, as with Agile encryption.
    pub encryption_name: Option<String>,
    pub encryption_block_size: u32,
    pub cipher_mode: u32,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct License {
    pub stream_name: String,
    /// Base64-encoded Unicode `LicenseID` retained verbatim.
    pub encoded_license_id: String,
    /// Certificate-chain XML retained verbatim and never interpreted.
    pub certificate_chain: Option<String>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct NamedDefinition {
    pub name: String,
    pub definition: Definition,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Transform {
    pub name: String,
    pub header: Header,
    pub irm: Option<IrmTransform>,
    pub encryption: Option<EncryptionTransform>,
    pub end_user_licenses: Vec<License>,
    /// Non-IRM bytes following the transform header.
    pub opaque_tail: Vec<u8>,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum DocumentKind {
    Ooxml,
    LegacyBinary,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Irm {
    pub document_kind: DocumentKind,
    pub protected_stream: String,
    pub viewer_content_stream: Option<String>,
    pub transform_name: String,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Graph {
    pub version: VersionInfo,
    pub map: Map,
    pub definitions: Vec<NamedDefinition>,
    pub transforms: Vec<Transform>,
    pub irm: Option<Irm>,
    /// Exact sensitivity-label XML bytes, retained inert when present.
    pub label_info: Option<Vec<u8>>,
    /// Validated typed view of `label_info`.
    pub labels: Option<List>,
    /// Integrity metadata for the public `SummaryInformation` property stream.
    pub summary_information_integrity: Option<Integrity>,
    /// Integrity metadata for the public `DocumentSummaryInformation` property stream.
    pub document_summary_information_integrity: Option<Integrity>,
    /// Public legacy Custom XML mirror and its IRM promotion semantics.
    pub custom_xml_data_store: Option<Store>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Integrity {
    pub info: Info,
    /// `None` means a future info-stream version that readers must ignore.
    pub valid: Option<bool>,
}

/// A deterministic identity for the `DataSpaces` streams captured by a snapshot.
///
/// This is intentionally scoped to the `DataSpaces` graph rather than to the
/// physical CFB allocation. Package patches therefore remain applicable when
/// unrelated streams are rewritten, while exact `DataSpaces` source bytes are
/// still required before a patch can be applied.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct Revision(u64);

impl Revision {
    /// Returns the compact source fingerprint.
    #[must_use]
    pub const fn value(self) -> u64 {
        self.0
    }

    /// Alias for [`Self::value`].
    #[must_use]
    pub const fn fingerprint(self) -> u64 {
        self.value()
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
struct Component {
    path: Vec<String>,
    kind: ReferenceKind,
}

#[derive(Debug, Clone, PartialEq, Eq)]
struct SourceStream {
    path: Vec<String>,
    bytes: Arc<[u8]>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
struct Source {
    streams: Vec<SourceStream>,
    components: Vec<Component>,
    revision: Revision,
}

/// An immutable, source-preserving `DataSpaces` graph snapshot.
///
/// The snapshot owns the exact bytes of every `DataSpaces` stream that it
/// exposes. An unchanged transaction therefore replays producer bytes exactly
/// and never canonicalizes IRM, license, encryption, label, or opaque payloads.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Snapshot {
    graph: Graph,
    source: Arc<Source>,
}

impl Snapshot {
    /// Parses and captures a validated `DataSpaces` graph from an OLE package.
    ///
    /// `None` means the package has no `\x06DataSpaces` storage.
    ///
    /// # Errors
    ///
    /// Returns [`Error`] when the OLE container or a `DataSpaces` stream cannot be read or validated.
    pub fn from_ole<R: Read + Seek>(ole: &mut OleFile<R>) -> Result<Option<Self>, Error> {
        let Some(graph) = inspect(ole)? else {
            return Ok(None);
        };
        let source = Source::capture(ole, &graph)?;
        Ok(Some(Self {
            graph,
            source: Arc::new(source),
        }))
    }

    /// Borrows the validated semantic graph.
    #[must_use]
    pub const fn graph(&self) -> &Graph {
        &self.graph
    }

    /// Returns the exact `DataSpaces` source revision.
    #[must_use]
    pub fn revision(&self) -> Revision {
        self.source.revision
    }

    /// Returns the exact `DataSpaces` source fingerprint.
    #[must_use]
    pub fn fingerprint(&self) -> u64 {
        self.revision().value()
    }

    /// Starts an isolated typed edit.
    #[must_use]
    pub fn edit(&self) -> Transaction {
        Transaction {
            source: self.clone(),
            candidate: self.graph.clone(),
        }
    }

    fn patch_to(&self, after: &Snapshot) -> Result<Patch, Error> {
        if self.source.components != after.source.components
            || self
                .source
                .streams
                .iter()
                .map(|stream| &stream.path)
                .ne(after.source.streams.iter().map(|stream| &stream.path))
        {
            return Err(invalid(
                "DataSpaces edit changed the package graph or stream paths",
            ));
        }
        let changes = self
            .source
            .streams
            .iter()
            .zip(&after.source.streams)
            .filter(|(before, changed)| before.bytes != changed.bytes)
            .map(|(before, changed)| StreamChange {
                path: before.path.clone(),
                before: Arc::clone(&before.bytes),
                after: Arc::clone(&changed.bytes),
            })
            .collect();
        Ok(Patch {
            base: self.revision(),
            target: after.revision(),
            before: Arc::clone(&self.source),
            after: Arc::clone(&after.source),
            before_graph: self.graph.clone(),
            after_graph: after.graph.clone(),
            changes,
        })
    }
}

/// A failure-atomic typed `DataSpaces` edit.
#[derive(Debug, Clone)]
pub struct Transaction {
    source: Snapshot,
    candidate: Graph,
}

impl Transaction {
    /// Borrows the source snapshot used by this transaction.
    #[must_use]
    pub const fn source(&self) -> &Snapshot {
        &self.source
    }

    /// Borrows the working semantic graph.
    #[must_use]
    pub const fn graph(&self) -> &Graph {
        &self.candidate
    }

    /// Replaces the checked `DataSpaceVersionInfo` value.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Invalid`] when `value` fails `DataSpaceVersionInfo` validation.
    pub fn set_version_info(&mut self, value: VersionInfo) -> Result<&mut Self, Error> {
        validate_version_info(&value)?;
        self.candidate.version = value;
        Ok(self)
    }

    /// Replaces and checks the `DataSpaceMap` value.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Invalid`] when `value` fails `DataSpaceMap` validation.
    pub fn set_map(&mut self, value: Map) -> Result<&mut Self, Error> {
        validate_map(&value)?;
        self.candidate.map = value;
        Ok(self)
    }

    /// Replaces an existing named `DataSpaceDefinition`.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Invalid`] when `name` or `value` is invalid, or no definition named `name` exists.
    pub fn set_definition(&mut self, name: &str, value: Definition) -> Result<&mut Self, Error> {
        validate_name(name, "data space name")?;
        validate_definition(&value)?;
        let definition = self
            .candidate
            .definitions
            .iter_mut()
            .find(|definition| definition.name == name)
            .ok_or_else(|| invalid(format!("unknown data space definition '{name}'")))?;
        definition.definition = value;
        Ok(self)
    }

    /// Replaces the non-payload header of an existing transform.
    ///
    /// The identity of a known IRM or encryption transform is immutable here;
    /// changing it would reinterpret its inert payload. Version fields are
    /// still validated by the transform-specific `MS-OFFCRYPTO` rules.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Invalid`] when `name` or `value` is invalid, the transform is unknown, or the edit would change a known transform's identity.
    pub fn set_transform_header(&mut self, name: &str, value: Header) -> Result<&mut Self, Error> {
        validate_name(name, "transform name")?;
        validate_transform_header(&value)?;
        let transform = self
            .candidate
            .transforms
            .iter_mut()
            .find(|transform| transform.name == name)
            .ok_or_else(|| invalid(format!("unknown transform '{name}'")))?;
        let mut candidate = transform.clone();
        if (candidate.irm.is_some() || candidate.encryption.is_some())
            && (candidate.header.transform_id != value.transform_id
                || candidate.header.transform_name != value.transform_name)
        {
            return Err(invalid(
                "known transform identity cannot be changed while its payload is inert",
            ));
        }
        if let Some(irm) = candidate.irm.as_mut() {
            irm.header = value.clone();
        }
        if let Some(encryption) = candidate.encryption.as_mut() {
            encryption.header = value.clone();
        }
        candidate.header = value;
        validate_transform_model(&candidate)?;
        *transform = candidate;
        Ok(self)
    }

    /// Replaces checked algorithm metadata for a known encryption transform.
    ///
    /// This never touches `EncryptedPackage`, `EncryptionInfo`, or any other
    /// encrypted payload. The transform metadata remains advisory when the
    /// authoritative encryption information says otherwise.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Invalid`] when the transform is unknown or is not an encryption transform, or the resulting metadata is invalid.
    pub fn set_encryption_info(
        &mut self,
        name: &str,
        encryption_name: Option<String>,
        encryption_block_size: u32,
        cipher_mode: u32,
    ) -> Result<&mut Self, Error> {
        let transform = self
            .candidate
            .transforms
            .iter_mut()
            .find(|transform| transform.name == name)
            .ok_or_else(|| invalid(format!("unknown transform '{name}'")))?;
        let mut candidate = transform.clone();
        let encryption = candidate
            .encryption
            .as_mut()
            .ok_or_else(|| invalid(format!("transform '{name}' is not an encryption transform")))?;
        encryption.encryption_name = encryption_name;
        encryption.encryption_block_size = encryption_block_size;
        encryption.cipher_mode = cipher_mode;
        validate_transform_model(&candidate)?;
        *transform = candidate;
        Ok(self)
    }

    /// Replaces the opaque tail of an unknown transform header.
    ///
    /// Known IRM and encryption payloads cannot be edited by this owner.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Invalid`] when `value` exceeds the parser limit, the transform is unknown, or its payload is inert.
    pub fn set_transform_opaque_tail(
        &mut self,
        name: &str,
        value: Vec<u8>,
    ) -> Result<&mut Self, Error> {
        if value.len() > MAX_STREAM_BYTES {
            return Err(invalid("transform opaque tail exceeds parser limit"));
        }
        let transform = self
            .candidate
            .transforms
            .iter_mut()
            .find(|transform| transform.name == name)
            .ok_or_else(|| invalid(format!("unknown transform '{name}'")))?;
        if transform.irm.is_some() || transform.encryption.is_some() {
            return Err(invalid(
                "known transform payloads are inert and cannot be replaced",
            ));
        }
        transform.opaque_tail = value;
        Ok(self)
    }

    /// Whether a semantic field differs from the source graph.
    #[must_use]
    pub fn is_changed(&self) -> bool {
        self.candidate != self.source.graph
    }

    /// Validates and materializes the current candidate snapshot.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Invalid`] when the candidate graph fails validation.
    pub fn snapshot(&self) -> Result<Snapshot, Error> {
        self.materialize()
    }

    /// Abandons the edit and returns the original snapshot.
    #[must_use]
    pub fn rollback(self) -> Snapshot {
        self.source
    }

    /// Validates and publishes the candidate with a reversible patch.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Invalid`] when the candidate graph fails validation.
    pub fn commit(self) -> Result<Commit, Error> {
        let snapshot = self.materialize()?;
        let patch = self.source.patch_to(&snapshot)?;
        Ok(Commit { snapshot, patch })
    }

    fn materialize(&self) -> Result<Snapshot, Error> {
        if !self.is_changed() {
            return Ok(self.source.clone());
        }
        let mut graph = self.candidate.clone();
        validate_graph_model(&graph.map, &graph.definitions, &graph.transforms)?;
        validate_component_references(&graph.map, &self.source.source.components)?;
        graph.irm = classify_irm(&graph.map, &graph.definitions, &graph.transforms)?;
        validate_derived_graph(&graph)?;
        let streams = encode_graph_streams(&graph, &self.source.source)?;
        let source = Source::new(streams, self.source.source.components.clone())?;
        Ok(Snapshot {
            graph,
            source: Arc::new(source),
        })
    }
}

/// A successful `DataSpaces` publication.
#[derive(Debug, Clone)]
pub struct Commit {
    snapshot: Snapshot,
    patch: Patch,
}

impl Commit {
    /// Whether any `DataSpaces` stream changed.
    #[must_use]
    pub fn changed(&self) -> bool {
        !self.patch.is_noop()
    }

    /// Borrows the published snapshot.
    #[must_use]
    pub const fn snapshot(&self) -> &Snapshot {
        &self.snapshot
    }

    /// Borrows the reversible patch.
    #[must_use]
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Moves the published snapshot out of the commit.
    #[must_use]
    pub fn into_snapshot(self) -> Snapshot {
        self.snapshot
    }

    /// Moves the patch out of the commit.
    #[must_use]
    pub fn into_patch(self) -> Patch {
        self.patch
    }
}

/// One source-checked replacement of a `DataSpaces` stream.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct StreamChange {
    path: Vec<String>,
    before: Arc<[u8]>,
    after: Arc<[u8]>,
}

impl StreamChange {
    /// Borrows the changed stream path.
    #[must_use]
    pub fn path(&self) -> &[String] {
        &self.path
    }

    /// Borrows the exact source bytes required by this replacement.
    #[must_use]
    pub fn before(&self) -> &[u8] {
        &self.before
    }

    /// Borrows the exact bytes produced by this replacement.
    #[must_use]
    pub fn after(&self) -> &[u8] {
        &self.after
    }
}

/// A reversible, source-checked `DataSpaces` graph patch.
#[derive(Debug, Clone)]
pub struct Patch {
    base: Revision,
    target: Revision,
    before: Arc<Source>,
    after: Arc<Source>,
    before_graph: Graph,
    after_graph: Graph,
    changes: Vec<StreamChange>,
}

impl Patch {
    /// Returns the expected source revision.
    #[must_use]
    pub const fn base(&self) -> Revision {
        self.base
    }

    /// Returns the revision produced by this patch.
    #[must_use]
    pub const fn target(&self) -> Revision {
        self.target
    }

    /// Returns the source fingerprint required by this patch.
    #[must_use]
    pub const fn source_fingerprint(&self) -> u64 {
        self.base.value()
    }

    /// Returns the resulting source fingerprint.
    #[must_use]
    pub const fn target_fingerprint(&self) -> u64 {
        self.target.value()
    }

    /// Borrows the semantic graph required before the patch.
    #[must_use]
    pub const fn before(&self) -> &Graph {
        &self.before_graph
    }

    /// Borrows the semantic graph produced by the patch.
    #[must_use]
    pub const fn after(&self) -> &Graph {
        &self.after_graph
    }

    /// Borrows the changed `DataSpaces` stream replacements.
    #[must_use]
    pub fn changes(&self) -> &[StreamChange] {
        &self.changes
    }

    /// Whether this patch is an exact `DataSpaces` no-op.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.changes.is_empty()
    }

    /// Alias for [`Self::is_noop`].
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.is_noop()
    }

    /// Applies the patch only to its exact source snapshot.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Invalid`] when `source` is not this patch's base snapshot.
    pub fn apply(&self, source: &Snapshot) -> Result<Snapshot, Error> {
        if source.revision() != self.base || source.source.as_ref() != self.before.as_ref() {
            return Err(invalid(
                "DataSpaces patch source does not match its base snapshot",
            ));
        }
        Ok(Snapshot {
            graph: self.after_graph.clone(),
            source: Arc::clone(&self.after),
        })
    }

    /// Returns the exact inverse replacement.
    #[must_use]
    pub fn inverse(&self) -> Self {
        let changes = self
            .changes
            .iter()
            .map(|change| StreamChange {
                path: change.path.clone(),
                before: Arc::clone(&change.after),
                after: Arc::clone(&change.before),
            })
            .collect();
        Self {
            base: self.target,
            target: self.base,
            before: Arc::clone(&self.after),
            after: Arc::clone(&self.before),
            before_graph: self.after_graph.clone(),
            after_graph: self.before_graph.clone(),
            changes,
        }
    }

    /// Rebuilds an OLE package after validating the exact `DataSpaces` source.
    ///
    /// All current streams and storages are copied into a fresh CFB writer;
    /// only the streams listed by [`Self::changes`] are replaced. IRM/license,
    /// encryption, and unrelated payloads are copied as inert bytes. The
    /// physical CFB allocation is intentionally rebuilt, while logical
    /// storage names, stream contents, and storage CLSIDs are retained.
    ///
    /// # Errors
    ///
    /// Returns [`Error`] when the package has no `DataSpaces` graph, its source does not match this patch's base, or the output cannot be written.
    pub fn write_to<R: Read + Seek, W: Write + Seek>(
        &self,
        ole: &mut OleFile<R>,
        output: &mut W,
    ) -> Result<(), Error> {
        let current =
            Snapshot::from_ole(ole)?.ok_or_else(|| invalid("package has no DataSpaces graph"))?;
        self.apply(&current)?;
        rebuild_ole(ole, &self.changes, output)
    }
}

impl Source {
    fn capture<R: Read + Seek>(ole: &mut OleFile<R>, graph: &Graph) -> Result<Self, Error> {
        let components = collect_components(ole)?;
        let known_paths = graph_stream_paths(graph)?;
        let mut paths = ole
            .list_streams()
            .into_iter()
            .filter(|path| path.first().is_some_and(|name| name == STORAGE))
            .collect::<Vec<_>>();
        paths.sort();
        for path in &known_paths {
            let references = path.iter().map(String::as_str).collect::<Vec<_>>();
            if !ole.exists(&references) {
                return Err(invalid(format!(
                    "DataSpaces source stream '{}' is missing",
                    path.join("/")
                )));
            }
        }
        let mut streams = Vec::new();
        for path in paths {
            let references = path.iter().map(String::as_str).collect::<Vec<_>>();
            let bytes = read_stream(ole, &references)?;
            streams.push(SourceStream {
                path,
                bytes: Arc::from(bytes.into_boxed_slice()),
            });
        }
        Self::new(streams, components)
    }

    fn new(streams: Vec<SourceStream>, components: Vec<Component>) -> Result<Self, Error> {
        if streams.is_empty() {
            return Err(invalid("DataSpaces source contains no streams"));
        }
        if streams.len() > MAX_ENTRIES {
            return Err(invalid("DataSpaces source contains too many streams"));
        }
        if components.len() > MAX_COMPONENTS {
            return Err(invalid(
                "OLE package contains too many directory components",
            ));
        }
        let mut stream_paths = std::collections::HashSet::with_capacity(streams.len());
        for stream in &streams {
            if stream.path.is_empty() || !stream_paths.insert(stream.path.clone()) {
                return Err(invalid("duplicate DataSpaces source stream path"));
            }
            if stream.bytes.len() > MAX_STREAM_BYTES {
                return Err(invalid("DataSpaces source stream exceeds parser limit"));
            }
        }
        let mut component_paths = std::collections::HashSet::with_capacity(components.len());
        for component in &components {
            if component.path.is_empty() || !component_paths.insert(component.path.clone()) {
                return Err(invalid("duplicate OLE directory component path"));
            }
        }
        Ok(Self {
            revision: Revision::of_streams(&streams),
            streams,
            components,
        })
    }
}

impl Revision {
    fn of_streams(streams: &[SourceStream]) -> Self {
        let mut value = 0xcbf2_9ce4_8422_2325u64;
        for stream in streams {
            for component in &stream.path {
                for byte in component.as_bytes() {
                    value ^= u64::from(*byte);
                    value = value.wrapping_mul(0x0000_0100_0000_01b3);
                }
                value ^= 0xff;
                value = value.wrapping_mul(0x0000_0100_0000_01b3);
            }
            value ^= 0xfe;
            value = value.wrapping_mul(0x0000_0100_0000_01b3);
            for byte in stream.bytes.iter() {
                value ^= u64::from(*byte);
                value = value.wrapping_mul(0x0000_0100_0000_01b3);
            }
            value ^= 0xfd;
            value = value.wrapping_mul(0x0000_0100_0000_01b3);
        }
        Self(value)
    }
}

impl Graph {
    fn end_user_license_count(&self) -> usize {
        self.transforms
            .iter()
            .map(|transform| transform.end_user_licenses.len())
            .sum()
    }
}

#[derive(Debug)]
struct StorageCopy {
    path: Vec<String>,
    clsid: Option<[u8; 16]>,
}

#[derive(Debug)]
struct StreamCopy {
    path: Vec<String>,
}

struct SliceReader<'a> {
    data: &'a [u8],
    offset: usize,
}

impl<'a> SliceReader<'a> {
    fn new(data: &'a [u8]) -> Result<Self, Error> {
        if data.len() > MAX_STREAM_BYTES {
            return Err(invalid("DataSpaces stream exceeds parser limit"));
        }
        Ok(Self { data, offset: 0 })
    }

    fn at(data: &'a [u8], offset: usize) -> Result<Self, Error> {
        let mut reader = Self::new(data)?;
        if offset > data.len() {
            return Err(invalid("parser offset exceeds stream"));
        }
        reader.offset = offset;
        Ok(reader)
    }

    fn position(&self) -> usize {
        self.offset
    }

    fn take(&mut self, count: usize) -> Result<&'a [u8], Error> {
        let end = self
            .offset
            .checked_add(count)
            .ok_or_else(|| invalid("stream offset overflow"))?;
        let bytes = self
            .data
            .get(self.offset..end)
            .ok_or_else(|| invalid("truncated DataSpaces stream"))?;
        self.offset = end;
        Ok(bytes)
    }

    fn u16(&mut self) -> Result<u16, Error> {
        let bytes = self
            .take(2)?
            .try_into()
            .map_err(|_err| invalid("invalid two-byte DataSpaces field"))?;
        Ok(u16::from_le_bytes(bytes))
    }

    fn u32(&mut self) -> Result<u32, Error> {
        let bytes = self
            .take(4)?
            .try_into()
            .map_err(|_err| invalid("invalid four-byte DataSpaces field"))?;
        Ok(u32::from_le_bytes(bytes))
    }

    fn version(&mut self) -> Result<Version, Error> {
        Ok(Version {
            major: self.u16()?,
            minor: self.u16()?,
        })
    }

    fn unicode_lpp4(&mut self) -> Result<String, Error> {
        let byte_len = usize::try_from(self.u32()?)
            .map_err(|_err| invalid("UNICODE-LP-P4 length overflows usize"))?;
        if byte_len == 0 || byte_len > MAX_STRING_BYTES || byte_len % 2 != 0 {
            return Err(invalid("invalid UNICODE-LP-P4 length"));
        }
        let bytes = self.take(byte_len)?;
        let mut units = Vec::with_capacity(byte_len / 2);
        for pair in bytes.chunks_exact(2) {
            units.push(u16::from_le_bytes([pair[0], pair[1]]));
        }
        let value =
            String::from_utf16(&units).map_err(|_err| invalid("invalid UNICODE-LP-P4 UTF-16"))?;
        if value.contains('\0') {
            return Err(invalid("UNICODE-LP-P4 string contains NUL"));
        }
        let padding = (4 - (byte_len % 4)) % 4;
        if self.take(padding)?.iter().any(|byte| *byte != 0) {
            return Err(invalid("UNICODE-LP-P4 padding is nonzero"));
        }
        Ok(value)
    }

    fn utf8_lpp4(&mut self) -> Result<Option<String>, Error> {
        let byte_len = usize::try_from(self.u32()?)
            .map_err(|_err| invalid("UTF-8-LP-P4 length overflows usize"))?;
        if byte_len == 0 {
            return Ok(None);
        }
        if byte_len > MAX_STRING_BYTES {
            return Err(invalid("UTF-8-LP-P4 length exceeds parser limit"));
        }
        let bytes = self.take(byte_len)?;
        let value =
            std::str::from_utf8(bytes).map_err(|_err| invalid("invalid UTF-8-LP-P4 UTF-8"))?;
        if value.contains('\0') {
            return Err(invalid("UTF-8-LP-P4 string contains NUL"));
        }
        let padding = (4 - (byte_len % 4)) % 4;
        if self.take(padding)?.iter().any(|byte| *byte != 0) {
            return Err(invalid("UTF-8-LP-P4 padding is nonzero"));
        }
        Ok(Some(value.to_string()))
    }

    fn finish(self) -> Result<(), Error> {
        if self.offset == self.data.len() {
            Ok(())
        } else {
            Err(invalid("trailing bytes in DataSpaces stream"))
        }
    }
}

/// Runs one checked `DataSpaces` edit and publishes it atomically.
///
/// # Errors
///
/// Returns [`Error`] when `edit` fails or the edited graph cannot be validated and published.
pub fn update<F>(snapshot: &Snapshot, edit: F) -> Result<Commit, Error>
where
    F: FnOnce(&mut Transaction) -> Result<(), Error>,
{
    let mut transaction = snapshot.edit();
    edit(&mut transaction)?;
    transaction.commit()
}

fn graph_stream_paths(graph: &Graph) -> Result<Vec<Vec<String>>, Error> {
    let mut paths = Vec::with_capacity(
        2 + graph.definitions.len() + graph.transforms.len() + graph.end_user_license_count(),
    );
    paths.push(vec![STORAGE.to_string(), "Version".to_string()]);
    paths.push(vec![STORAGE.to_string(), "DataSpaceMap".to_string()]);
    for definition in &graph.definitions {
        validate_name(&definition.name, "data space name")?;
        paths.push(vec![
            STORAGE.to_string(),
            "DataSpaceInfo".to_string(),
            definition.name.clone(),
        ]);
    }
    for transform in &graph.transforms {
        validate_name(&transform.name, "transform name")?;
        paths.push(vec![
            STORAGE.to_string(),
            "TransformInfo".to_string(),
            transform.name.clone(),
            PRIMARY.to_string(),
        ]);
        for license in &transform.end_user_licenses {
            validate_eul_stream_name(&license.stream_name)?;
            paths.push(vec![
                STORAGE.to_string(),
                "TransformInfo".to_string(),
                transform.name.clone(),
                license.stream_name.clone(),
            ]);
        }
    }
    if graph.label_info.is_some() {
        paths.push(vec![
            STORAGE.to_string(),
            "TransformInfo".to_string(),
            "LabelInfo".to_string(),
        ]);
    }
    Ok(paths)
}

fn collect_components<R: Read + Seek>(ole: &OleFile<R>) -> Result<Vec<Component>, Error> {
    let mut components = Vec::new();
    let mut path = Vec::new();
    collect_directory_components(ole, &mut path, &mut components, 0)?;
    Ok(components)
}

fn collect_directory_components<R: Read + Seek>(
    ole: &OleFile<R>,
    path: &mut Vec<String>,
    components: &mut Vec<Component>,
    depth: usize,
) -> Result<(), Error> {
    if depth > MAX_XML_DEPTH {
        return Err(invalid("OLE directory nesting exceeds parser limit"));
    }
    let references = path.iter().map(String::as_str).collect::<Vec<_>>();
    let raw_entries = ole.list_directory_entries(&references)?;
    let entries = raw_entries
        .into_iter()
        .map(|entry| (entry.name.clone(), entry.entry_type))
        .collect::<Vec<_>>();
    for (name, kind) in entries {
        if name.is_empty() {
            return Err(invalid("OLE directory entry has an empty name"));
        }
        path.push(name);
        match kind {
            STGTY_STORAGE => {
                components.push(Component {
                    path: path.clone(),
                    kind: ReferenceKind::Storage,
                });
                if components.len() > MAX_COMPONENTS {
                    return Err(invalid(
                        "OLE package contains too many directory components",
                    ));
                }
                collect_directory_components(ole, path, components, depth + 1)?;
            },
            STGTY_STREAM => {
                components.push(Component {
                    path: path.clone(),
                    kind: ReferenceKind::Stream,
                });
                if components.len() > MAX_COMPONENTS {
                    return Err(invalid(
                        "OLE package contains too many directory components",
                    ));
                }
            },
            _ => {
                return Err(invalid(format!(
                    "unsupported OLE directory entry type {kind} at '{}'",
                    path.join("/")
                )));
            },
        }
        path.pop();
    }
    Ok(())
}

fn validate_graph_model(
    map: &Map,
    definitions: &[NamedDefinition],
    transforms: &[Transform],
) -> Result<(), Error> {
    validate_map(map)?;
    if definitions.len() != map.entries.len() {
        return Err(invalid(
            "DataSpaceInfo streams do not correspond one-to-one with map entries",
        ));
    }
    let mut definition_names = std::collections::HashSet::with_capacity(definitions.len());
    for definition in definitions {
        validate_name(&definition.name, "data space name")?;
        if !definition_names.insert(definition.name.as_str()) {
            return Err(invalid("duplicate data space definition name"));
        }
        validate_definition(&definition.definition)?;
        if !map
            .entries
            .iter()
            .any(|entry| entry.data_space_name == definition.name)
        {
            return Err(invalid(format!(
                "data space definition '{}' is not present in the map",
                definition.name
            )));
        }
    }

    let mut transform_names = std::collections::HashSet::with_capacity(transforms.len());
    for transform in transforms {
        validate_name(&transform.name, "transform name")?;
        if !transform_names.insert(transform.name.as_str()) {
            return Err(invalid("duplicate transform name"));
        }
        validate_transform_model(transform)?;
    }
    for definition in definitions {
        for transform_name in &definition.definition.transforms {
            if !transform_names.contains(transform_name.as_str()) {
                return Err(invalid(format!(
                    "data space '{}' references missing transform '{}'",
                    definition.name, transform_name
                )));
            }
        }
    }
    Ok(())
}

fn validate_transform_model(transform: &Transform) -> Result<(), Error> {
    validate_transform_header(&transform.header)?;
    if transform.opaque_tail.len() > MAX_STREAM_BYTES {
        return Err(invalid("transform opaque tail exceeds parser limit"));
    }
    match (&transform.irm, &transform.encryption) {
        (Some(_), Some(_)) => {
            return Err(invalid(
                "a transform cannot contain both IRM and encryption metadata",
            ));
        },
        (Some(irm), None) => {
            if irm.header != transform.header || !transform.opaque_tail.is_empty() {
                return Err(invalid("IRM transform metadata is inconsistent"));
            }
            validate_drm_header(&irm.header)?;
            let license = irm
                .publishing_license
                .as_deref()
                .filter(|license| !license.is_empty())
                .ok_or_else(|| invalid("IRM publishing license cannot be null or empty"))?;
            validate_inert_xml(license, "IRM publishing license")?;
        },
        (None, Some(encryption)) => {
            if encryption.header != transform.header || !transform.opaque_tail.is_empty() {
                return Err(invalid("encryption transform metadata is inconsistent"));
            }
            validate_encryption_transform(encryption)?;
        },
        (None, None) => {},
    }
    for license in &transform.end_user_licenses {
        validate_eul_stream_name(&license.stream_name)?;
        if license.encoded_license_id.is_empty() {
            return Err(invalid("EndUserLicenseHeader.ID_String cannot be empty"));
        }
        let chain = license
            .certificate_chain
            .as_deref()
            .filter(|chain| !chain.is_empty())
            .ok_or_else(|| invalid("end-user license certificate chain cannot be null or empty"))?;
        validate_inert_xml(chain, "end-user license certificate chain")?;
    }
    if transform.irm.is_some() && transform.end_user_licenses.is_empty() {
        return Err(invalid(format!(
            "IRM transform '{}' has no end-user license stream",
            transform.name
        )));
    }
    Ok(())
}

fn validate_component_references(map: &Map, components: &[Component]) -> Result<(), Error> {
    let mut protected_paths = std::collections::HashSet::with_capacity(map.entries.len());
    for entry in &map.entries {
        let mut path = Vec::with_capacity(entry.references.len());
        for reference in &entry.references {
            path.push(reference.component.clone());
            let component = components
                .iter()
                .find(|component| same_path(&component.path, &path))
                .ok_or_else(|| {
                    invalid(format!(
                        "map references missing component '{}'",
                        path.join("/")
                    ))
                })?;
            if component.kind != reference.kind {
                return Err(invalid(format!(
                    "map component '{}' has the wrong reference kind",
                    path.join("/")
                )));
            }
        }
        let protected_path = path
            .iter()
            .map(|component| component.to_ascii_lowercase())
            .collect::<Vec<_>>();
        if !protected_paths.insert(protected_path) {
            return Err(invalid(
                "a protected content component is mapped by more than one data space",
            ));
        }
    }
    Ok(())
}

fn same_path(left: &[String], right: &[String]) -> bool {
    left.len() == right.len()
        && left
            .iter()
            .zip(right)
            .all(|(lhs, rhs)| lhs.eq_ignore_ascii_case(rhs))
}

fn validate_derived_graph(graph: &Graph) -> Result<(), Error> {
    let expected_irm = classify_irm(&graph.map, &graph.definitions, &graph.transforms)?;
    if graph.irm != expected_irm {
        return Err(invalid(
            "derived IRM profile does not match the DataSpaces graph",
        ));
    }
    if graph.label_info.is_some() != graph.labels.is_some() {
        return Err(invalid(
            "LabelInfo bytes and its typed sensitivity-label view disagree",
        ));
    }
    if graph.labels.is_some()
        && !graph.irm.as_ref().is_some_and(|profile| {
            graph.transforms.iter().any(|transform| {
                transform.name == profile.transform_name
                    && transform
                        .irm
                        .as_ref()
                        .is_some_and(|metadata| metadata.publishing_license.is_some())
            })
        })
    {
        return Err(invalid(
            "LabelInfo requires an IRM transform with a publishing license",
        ));
    }
    if (graph.summary_information_integrity.is_some()
        || graph.document_summary_information_integrity.is_some())
        && graph.irm.is_none()
        && !graph
            .transforms
            .iter()
            .any(|transform| transform.encryption.is_some())
    {
        return Err(invalid(
            "encrypted property hash stream is present without an encryption or IRM transform",
        ));
    }
    validate_custom_xml_promotion(graph.custom_xml_data_store.as_ref(), graph.irm.as_ref())
}

fn encode_graph_streams(graph: &Graph, source: &Source) -> Result<Vec<SourceStream>, Error> {
    validate_version_info(&graph.version)?;
    validate_graph_model(&graph.map, &graph.definitions, &graph.transforms)?;
    let mut known = Vec::new();
    known.push(SourceStream {
        path: vec![STORAGE.to_string(), "Version".to_string()],
        bytes: Arc::from(write_version_info(&graph.version)?.into_boxed_slice()),
    });
    known.push(SourceStream {
        path: vec![STORAGE.to_string(), "DataSpaceMap".to_string()],
        bytes: Arc::from(write_map(&graph.map)?.into_boxed_slice()),
    });
    for definition in &graph.definitions {
        known.push(SourceStream {
            path: vec![
                STORAGE.to_string(),
                "DataSpaceInfo".to_string(),
                definition.name.clone(),
            ],
            bytes: Arc::from(write_definition(&definition.definition)?.into_boxed_slice()),
        });
    }
    for transform in &graph.transforms {
        known.push(SourceStream {
            path: vec![
                STORAGE.to_string(),
                "TransformInfo".to_string(),
                transform.name.clone(),
                PRIMARY.to_string(),
            ],
            bytes: Arc::from(encode_transform_primary(transform)?.into_boxed_slice()),
        });
        for license in &transform.end_user_licenses {
            known.push(SourceStream {
                path: vec![
                    STORAGE.to_string(),
                    "TransformInfo".to_string(),
                    transform.name.clone(),
                    license.stream_name.clone(),
                ],
                bytes: Arc::from(write_license(license)?.into_boxed_slice()),
            });
        }
    }
    if let Some(label_info) = &graph.label_info {
        known.push(SourceStream {
            path: vec![
                STORAGE.to_string(),
                "TransformInfo".to_string(),
                "LabelInfo".to_string(),
            ],
            bytes: Arc::from(label_info.clone().into_boxed_slice()),
        });
    }
    let mut encoded = known
        .into_iter()
        .map(|stream| (stream.path.clone(), stream))
        .collect::<std::collections::HashMap<_, _>>();
    let mut streams = Vec::with_capacity(source.streams.len());
    for original in &source.streams {
        if let Some(replacement) = encoded.remove(&original.path) {
            streams.push(replacement);
        } else {
            streams.push(original.clone());
        }
    }
    if !encoded.is_empty() {
        return Err(invalid(
            "DataSpaces edit changed the set of known metadata streams",
        ));
    }
    Ok(streams)
}

fn encode_transform_primary(transform: &Transform) -> Result<Vec<u8>, Error> {
    match (&transform.irm, &transform.encryption) {
        (Some(irm), None) => write_irm_transform(irm),
        (None, Some(encryption)) => write_encryption_transform(encryption),
        (None, None) => {
            let mut bytes = write_transform_header(&transform.header)?;
            bytes.extend_from_slice(&transform.opaque_tail);
            Ok(bytes)
        },
        (Some(_), Some(_)) => Err(invalid(
            "a transform cannot contain both IRM and encryption metadata",
        )),
    }
}

fn rebuild_ole<R: Read + Seek, W: Write + Seek>(
    ole: &mut OleFile<R>,
    changes: &[StreamChange],
    output: &mut W,
) -> Result<(), Error> {
    let (storages, mut streams) = collect_ole_layout(ole)?;
    let mut writer = OleWriter::with_sector_size(ole.sector_size())?;
    if let Some(root) = ole.root_entry()
        && let Some(clsid) = parse_clsid(&root.clsid)?
    {
        writer.set_root_clsid(clsid);
    }
    for storage in &storages {
        let path = storage.path.iter().map(String::as_str).collect::<Vec<_>>();
        writer.create_storage(&path)?;
        if let Some(clsid) = storage.clsid {
            writer.set_storage_clsid(&path, clsid)?;
        }
    }

    // Word's legacy writer requires WordDocument to receive the first large
    // stream allocation. Keep that format-specific invariant while preserving
    // the source traversal order for every other stream.
    streams.sort_by_key(|stream| u8::from(stream.path.as_slice() != ["WordDocument"]));
    for stream in streams {
        let path = stream.path.iter().map(String::as_str).collect::<Vec<_>>();
        let data = if let Some(change) = changes
            .iter()
            .find(|change| same_path(&change.path, &stream.path))
        {
            change.after.to_vec()
        } else {
            ole.open_stream(&path)?
        };
        writer.create_stream_owned(&path, data)?;
    }
    writer.write_to(output)?;
    Ok(())
}

fn collect_ole_layout<R: Read + Seek>(
    ole: &OleFile<R>,
) -> Result<(Vec<StorageCopy>, Vec<StreamCopy>), Error> {
    let mut storages = Vec::new();
    let mut streams = Vec::new();
    let mut path = Vec::new();
    collect_ole_layout_directory(ole, &mut path, &mut storages, &mut streams, 0)?;
    Ok((storages, streams))
}

fn collect_ole_layout_directory<R: Read + Seek>(
    ole: &OleFile<R>,
    path: &mut Vec<String>,
    storages: &mut Vec<StorageCopy>,
    streams: &mut Vec<StreamCopy>,
    depth: usize,
) -> Result<(), Error> {
    if depth > MAX_XML_DEPTH {
        return Err(invalid("OLE directory nesting exceeds parser limit"));
    }
    let references = path.iter().map(String::as_str).collect::<Vec<_>>();
    let raw_entries = ole.list_directory_entries(&references)?;
    let entries = raw_entries
        .into_iter()
        .map(|entry| (entry.name.clone(), entry.entry_type, entry.clsid.clone()))
        .collect::<Vec<_>>();
    for (name, entry_type, clsid) in entries {
        if name.is_empty() {
            return Err(invalid("OLE directory entry has an empty name"));
        }
        path.push(name);
        match entry_type {
            STGTY_STORAGE => {
                storages.push(StorageCopy {
                    path: path.clone(),
                    clsid: parse_clsid(&clsid)?,
                });
                collect_ole_layout_directory(ole, path, storages, streams, depth + 1)?;
            },
            STGTY_STREAM => streams.push(StreamCopy { path: path.clone() }),
            _ => {
                return Err(invalid(format!(
                    "unsupported OLE directory entry type {entry_type} at '{}'",
                    path.join("/")
                )));
            },
        }
        path.pop();
    }
    Ok(())
}

fn parse_clsid(value: &str) -> Result<Option<[u8; 16]>, Error> {
    if value.is_empty() {
        return Ok(None);
    }
    let fields = value.split('-').collect::<Vec<_>>();
    if fields.len() != 5
        || fields[0].len() != 8
        || fields[1].len() != 4
        || fields[2].len() != 4
        || fields[3].len() != 4
        || fields[4].len() != 12
    {
        return Err(invalid(format!("invalid CFB CLSID '{value}'")));
    }
    let data1 = u32::from_str_radix(fields[0], 16)
        .map_err(|_err| invalid(format!("invalid CFB CLSID '{value}'")))?;
    let data2 = u16::from_str_radix(fields[1], 16)
        .map_err(|_err| invalid(format!("invalid CFB CLSID '{value}'")))?;
    let data3 = u16::from_str_radix(fields[2], 16)
        .map_err(|_err| invalid(format!("invalid CFB CLSID '{value}'")))?;
    let mut bytes = [0u8; 16];
    bytes[..4].copy_from_slice(&data1.to_le_bytes());
    bytes[4..6].copy_from_slice(&data2.to_le_bytes());
    bytes[6..8].copy_from_slice(&data3.to_le_bytes());
    for (index, pair) in fields[3].as_bytes().chunks_exact(2).enumerate() {
        bytes[8 + index] = u8::from_str_radix(
            std::str::from_utf8(pair)
                .map_err(|_err| invalid(format!("invalid CFB CLSID '{value}'")))?,
            16,
        )
        .map_err(|_err| invalid(format!("invalid CFB CLSID '{value}'")))?;
    }
    for (index, pair) in fields[4].as_bytes().chunks_exact(2).enumerate() {
        bytes[10 + index] = u8::from_str_radix(
            std::str::from_utf8(pair)
                .map_err(|_err| invalid(format!("invalid CFB CLSID '{value}'")))?,
            16,
        )
        .map_err(|_err| invalid(format!("invalid CFB CLSID '{value}'")))?;
    }
    Ok(Some(bytes))
}

/// Parses a `DataSpaceVersionInfo` stream.
///
/// # Errors
///
/// Returns [`Error::Invalid`] when the input is truncated, malformed, or fails validation.
pub fn parse_version_info(data: &[u8]) -> Result<VersionInfo, Error> {
    let mut reader = SliceReader::new(data)?;
    let value = VersionInfo {
        feature_identifier: reader.unicode_lpp4()?,
        reader: reader.version()?,
        updater: reader.version()?,
        writer: reader.version()?,
    };
    reader.finish()?;
    validate_version_info(&value)?;
    Ok(value)
}

/// Serializes a `DataSpaceVersionInfo` value.
///
/// # Errors
///
/// Returns [`Error::Invalid`] when the value fails validation.
pub fn write_version_info(value: &VersionInfo) -> Result<Vec<u8>, Error> {
    validate_version_info(value)?;
    let mut output = Vec::new();
    write_unicode_lpp4(&mut output, &value.feature_identifier)?;
    write_version(&mut output, value.reader);
    write_version(&mut output, value.updater);
    write_version(&mut output, value.writer);
    Ok(output)
}

/// Parses a `DataSpaceMap` stream.
///
/// # Errors
///
/// Returns [`Error::Invalid`] when the input is truncated, malformed, or fails validation.
pub fn parse_map(data: &[u8]) -> Result<Map, Error> {
    let mut reader = SliceReader::new(data)?;
    require_u32(reader.u32()?, HEADER_LENGTH, "DataSpaceMap.HeaderLength")?;
    let count = bounded_count(reader.u32()?, "DataSpaceMap.EntryCount")?;
    if count == 0 {
        return Err(invalid("DataSpaceMap requires at least one entry"));
    }
    let mut entries = Vec::with_capacity(count);
    for _ in 0..count {
        let start = reader.position();
        let length = usize::try_from(reader.u32()?)
            .map_err(|_err| invalid("DataSpaceMapEntry.Length overflows usize"))?;
        if length < 12 || start.checked_add(length).is_none_or(|end| end > data.len()) {
            return Err(invalid("DataSpaceMapEntry.Length exceeds its stream"));
        }
        let reference_count =
            bounded_count(reader.u32()?, "DataSpaceMapEntry.ReferenceComponentCount")?;
        if reference_count == 0 {
            return Err(invalid("DataSpaceMapEntry requires a reference component"));
        }
        let mut references = Vec::with_capacity(reference_count);
        for _ in 0..reference_count {
            let kind = match reader.u32()? {
                0 => ReferenceKind::Stream,
                1 => ReferenceKind::Storage,
                value => {
                    return Err(invalid(format!(
                        "unknown DataSpaceReferenceComponent type {value}"
                    )));
                },
            };
            references.push(Reference {
                kind,
                component: reader.unicode_lpp4()?,
            });
        }
        let data_space_name = reader.unicode_lpp4()?;
        if reader.position() != start + length {
            return Err(invalid(
                "DataSpaceMapEntry.Length does not match its fields",
            ));
        }
        entries.push(MapEntry {
            references,
            data_space_name,
        });
    }
    reader.finish()?;
    let value = Map { entries };
    validate_map(&value)?;
    Ok(value)
}

/// Serializes a `DataSpaceMap` value.
///
/// # Errors
///
/// Returns [`Error::Invalid`] when the value fails validation.
pub fn write_map(value: &Map) -> Result<Vec<u8>, Error> {
    validate_map(value)?;
    let mut output = Vec::new();
    output.extend_from_slice(&HEADER_LENGTH.to_le_bytes());
    write_count(&mut output, value.entries.len(), "DataSpaceMap.EntryCount")?;
    for entry in &value.entries {
        let start = output.len();
        output.extend_from_slice(&0u32.to_le_bytes());
        write_count(
            &mut output,
            entry.references.len(),
            "DataSpaceMapEntry.ReferenceComponentCount",
        )?;
        for reference in &entry.references {
            let kind = match reference.kind {
                ReferenceKind::Stream => 0u32,
                ReferenceKind::Storage => 1u32,
            };
            output.extend_from_slice(&kind.to_le_bytes());
            write_unicode_lpp4(&mut output, &reference.component)?;
        }
        write_unicode_lpp4(&mut output, &entry.data_space_name)?;
        let length = u32::try_from(output.len() - start)
            .map_err(|_err| invalid("DataSpaceMapEntry.Length exceeds u32"))?;
        output[start..start + 4].copy_from_slice(&length.to_le_bytes());
    }
    Ok(output)
}

/// Parses a `DataSpaceDefinition` stream.
///
/// # Errors
///
/// Returns [`Error::Invalid`] when the input is truncated, malformed, or fails validation.
pub fn parse_definition(data: &[u8]) -> Result<Definition, Error> {
    let mut reader = SliceReader::new(data)?;
    require_u32(
        reader.u32()?,
        HEADER_LENGTH,
        "DataSpaceDefinition.HeaderLength",
    )?;
    let count = bounded_count(reader.u32()?, "DataSpaceDefinition.TransformReferenceCount")?;
    if count == 0 {
        return Err(invalid(
            "DataSpaceDefinition requires at least one transform",
        ));
    }
    let mut transforms = Vec::with_capacity(count);
    for _ in 0..count {
        transforms.push(reader.unicode_lpp4()?);
    }
    reader.finish()?;
    let value = Definition { transforms };
    validate_definition(&value)?;
    Ok(value)
}

/// Serializes a `DataSpaceDefinition` value.
///
/// # Errors
///
/// Returns [`Error::Invalid`] when the value fails validation.
pub fn write_definition(value: &Definition) -> Result<Vec<u8>, Error> {
    validate_definition(value)?;
    let mut output = Vec::new();
    output.extend_from_slice(&HEADER_LENGTH.to_le_bytes());
    write_count(
        &mut output,
        value.transforms.len(),
        "DataSpaceDefinition.TransformReferenceCount",
    )?;
    for transform in &value.transforms {
        write_unicode_lpp4(&mut output, transform)?;
    }
    Ok(output)
}

/// Parses a transform header, returning the header and the byte count it consumed.
///
/// # Errors
///
/// Returns [`Error::Invalid`] when the input is truncated, malformed, or fails validation.
pub fn parse_transform_header(data: &[u8]) -> Result<(Header, usize), Error> {
    let mut reader = SliceReader::new(data)?;
    let transform_length = usize::try_from(reader.u32()?)
        .map_err(|_err| invalid("TransformLength overflows usize"))?;
    require_u32(reader.u32()?, TRANSFORM_TYPE, "TransformType")?;
    let transform_id = reader.unicode_lpp4()?;
    if reader.position() != transform_length {
        return Err(invalid("TransformLength does not end before TransformName"));
    }
    let value = Header {
        transform_id,
        transform_name: reader.unicode_lpp4()?,
        reader: reader.version()?,
        updater: reader.version()?,
        writer: reader.version()?,
    };
    validate_transform_header(&value)?;
    Ok((value, reader.position()))
}

/// Serializes a transform header.
///
/// # Errors
///
/// Returns [`Error::Invalid`] when the value fails validation.
pub fn write_transform_header(value: &Header) -> Result<Vec<u8>, Error> {
    validate_transform_header(value)?;
    let mut output = Vec::new();
    output.extend_from_slice(&0u32.to_le_bytes());
    output.extend_from_slice(&TRANSFORM_TYPE.to_le_bytes());
    write_unicode_lpp4(&mut output, &value.transform_id)?;
    let transform_length =
        u32::try_from(output.len()).map_err(|_err| invalid("TransformLength exceeds u32"))?;
    output[..4].copy_from_slice(&transform_length.to_le_bytes());
    write_unicode_lpp4(&mut output, &value.transform_name)?;
    write_version(&mut output, value.reader);
    write_version(&mut output, value.updater);
    write_version(&mut output, value.writer);
    Ok(output)
}

/// Parses an IRM transform stream.
///
/// # Errors
///
/// Returns [`Error::Invalid`] when the input is truncated, malformed, or fails validation.
pub fn parse_irm_transform(data: &[u8]) -> Result<IrmTransform, Error> {
    let (header, consumed) = parse_transform_header(data)?;
    validate_drm_header(&header)?;
    let mut reader = SliceReader::at(data, consumed)?;
    require_u32(
        reader.u32()?,
        EXTENSIBILITY_HEADER_LENGTH,
        "ExtensibilityHeader.Length",
    )?;
    let publishing_license = reader.utf8_lpp4()?;
    let license = publishing_license
        .as_deref()
        .filter(|license| !license.is_empty())
        .ok_or_else(|| invalid("IRM publishing license cannot be null or empty"))?;
    validate_inert_xml(license, "IRM publishing license")?;
    reader.finish()?;
    Ok(IrmTransform {
        header,
        publishing_license,
    })
}

/// Serializes an IRM transform.
///
/// # Errors
///
/// Returns [`Error::Invalid`] when the value fails validation.
pub fn write_irm_transform(value: &IrmTransform) -> Result<Vec<u8>, Error> {
    validate_drm_header(&value.header)?;
    let license = value
        .publishing_license
        .as_deref()
        .filter(|license| !license.is_empty())
        .ok_or_else(|| invalid("IRM publishing license cannot be null or empty"))?;
    validate_inert_xml(license, "IRM publishing license")?;
    let mut output = write_transform_header(&value.header)?;
    output.extend_from_slice(&EXTENSIBILITY_HEADER_LENGTH.to_le_bytes());
    write_utf8_lpp4(&mut output, value.publishing_license.as_deref())?;
    Ok(output)
}

/// Parses an `EncryptionTransformInfo` stream.
///
/// # Errors
///
/// Returns [`Error::Invalid`] when the input is truncated, malformed, or fails validation.
pub fn parse_encryption_transform(data: &[u8]) -> Result<EncryptionTransform, Error> {
    let (header, consumed) = parse_transform_header(data)?;
    validate_encryption_header(&header)?;
    let mut reader = SliceReader::at(data, consumed)?;
    let value = EncryptionTransform {
        header,
        encryption_name: reader.utf8_lpp4()?,
        encryption_block_size: reader.u32()?,
        cipher_mode: reader.u32()?,
    };
    require_u32(reader.u32()?, 4, "EncryptionTransformInfo.Reserved")?;
    reader.finish()?;
    validate_encryption_transform(&value)?;
    Ok(value)
}

/// Serializes an `EncryptionTransformInfo` value.
///
/// # Errors
///
/// Returns [`Error::Invalid`] when the value fails validation.
pub fn write_encryption_transform(value: &EncryptionTransform) -> Result<Vec<u8>, Error> {
    validate_encryption_transform(value)?;
    let mut output = write_transform_header(&value.header)?;
    write_utf8_lpp4(&mut output, value.encryption_name.as_deref())?;
    output.extend_from_slice(&value.encryption_block_size.to_le_bytes());
    output.extend_from_slice(&value.cipher_mode.to_le_bytes());
    output.extend_from_slice(&4u32.to_le_bytes());
    Ok(output)
}

/// Parses an end-user license stream.
///
/// # Errors
///
/// Returns [`Error::Invalid`] when the input is truncated, malformed, or fails validation.
pub fn parse_license(stream_name: &str, data: &[u8]) -> Result<License, Error> {
    validate_eul_stream_name(stream_name)?;
    let mut reader = SliceReader::new(data)?;
    let header_start = reader.position();
    let header_length = usize::try_from(reader.u32()?)
        .map_err(|_err| invalid("EndUserLicenseHeader.Length overflows usize"))?;
    if header_length < 8 || header_length > data.len() {
        return Err(invalid("invalid EndUserLicenseHeader.Length"));
    }
    let encoded_license_id = reader
        .utf8_lpp4()?
        .ok_or_else(|| invalid("EndUserLicenseHeader.ID_String cannot be null"))?;
    if reader.position() != header_start + header_length {
        return Err(invalid(
            "EndUserLicenseHeader.Length does not match ID_String",
        ));
    }
    let certificate_chain = reader.utf8_lpp4()?;
    let chain = certificate_chain
        .as_deref()
        .filter(|chain| !chain.is_empty())
        .ok_or_else(|| invalid("end-user license certificate chain cannot be null or empty"))?;
    validate_inert_xml(chain, "end-user license certificate chain")?;
    reader.finish()?;
    Ok(License {
        stream_name: stream_name.to_string(),
        encoded_license_id,
        certificate_chain,
    })
}

/// Serializes an end-user license stream.
///
/// # Errors
///
/// Returns [`Error::Invalid`] when the value fails validation.
pub fn write_license(value: &License) -> Result<Vec<u8>, Error> {
    validate_eul_stream_name(&value.stream_name)?;
    if value.encoded_license_id.is_empty() {
        return Err(invalid("EndUserLicenseHeader.ID_String cannot be empty"));
    }
    let chain = value
        .certificate_chain
        .as_deref()
        .filter(|chain| !chain.is_empty())
        .ok_or_else(|| invalid("end-user license certificate chain cannot be null or empty"))?;
    validate_inert_xml(chain, "end-user license certificate chain")?;
    let mut output = vec![0; 4];
    write_utf8_lpp4(&mut output, Some(&value.encoded_license_id))?;
    let header_length = u32::try_from(output.len())
        .map_err(|_err| invalid("EndUserLicenseHeader.Length exceeds u32"))?;
    output[..4].copy_from_slice(&header_length.to_le_bytes());
    write_utf8_lpp4(&mut output, value.certificate_chain.as_deref())?;
    Ok(output)
}

/// Inspect and cross-validate a complete `DataSpaces` graph in an OLE file.
///
/// # Errors
///
/// Returns [`Error`] when the OLE container or a `DataSpaces` stream cannot be read or validated.
pub fn inspect<R: Read + Seek>(ole: &mut OleFile<R>) -> Result<Option<Graph>, Error> {
    let custom_xml_data_store = inspect_custom_xml(ole)
        .map_err(|error| invalid(format!("MsoDataStore validation failed: {error}")))?;
    if !ole.exists(&[STORAGE]) {
        validate_custom_xml_promotion(custom_xml_data_store.as_ref(), None)?;
        return Ok(None);
    }
    let version = parse_version_info(&read_stream(ole, &[STORAGE, "Version"])?)?;
    let map = parse_map(&read_stream(ole, &[STORAGE, "DataSpaceMap"])?)?;

    let definition_entries = ole.list_directory_entries(&[STORAGE, "DataSpaceInfo"])?;
    if definition_entries.len() > MAX_ENTRIES {
        return Err(invalid("too many DataSpaceInfo entries"));
    }
    let mut definition_names = Vec::with_capacity(definition_entries.len());
    for entry in definition_entries {
        if entry.entry_type != 2 {
            return Err(invalid(format!(
                "DataSpaceInfo child '{}' is not a stream",
                entry.name
            )));
        }
        definition_names.push(entry.name.clone());
    }
    definition_names.sort();
    let mut definitions = Vec::with_capacity(definition_names.len());
    for name in definition_names {
        definitions.push(NamedDefinition {
            definition: parse_definition(&read_stream(ole, &[STORAGE, "DataSpaceInfo", &name])?)?,
            name,
        });
    }

    let transform_entries = ole.list_directory_entries(&[STORAGE, "TransformInfo"])?;
    if transform_entries.len() > MAX_ENTRIES {
        return Err(invalid("too many TransformInfo entries"));
    }
    let mut transform_names = Vec::with_capacity(transform_entries.len());
    for entry in transform_entries {
        if entry.entry_type != 1 {
            // LabelInfo is a permitted stream sibling, not a transform.
            if entry.entry_type == 2 && entry.name == "LabelInfo" {
                continue;
            }
            return Err(invalid(format!(
                "TransformInfo child '{}' is not a storage",
                entry.name
            )));
        }
        transform_names.push(entry.name.clone());
    }
    transform_names.sort();
    let mut transforms = Vec::with_capacity(transform_names.len());
    for name in transform_names {
        let child_entries = ole
            .list_directory_entries(&[STORAGE, "TransformInfo", &name])?
            .iter()
            .map(|entry| (entry.name.clone(), entry.entry_type))
            .collect::<Vec<_>>();
        if child_entries.len() > MAX_ENTRIES {
            return Err(invalid("too many transform-storage entries"));
        }
        let bytes = read_stream(ole, &[STORAGE, "TransformInfo", &name, PRIMARY])?;
        let (header, consumed) = parse_transform_header(&bytes)?;
        let irm = if header.transform_id == DRM_ID && header.transform_name == DRM_NAME {
            Some(parse_irm_transform(&bytes)?)
        } else {
            None
        };
        let encryption =
            if header.transform_id == ENCRYPTION_ID && header.transform_name == ENCRYPTION_NAME {
                Some(parse_encryption_transform(&bytes)?)
            } else {
                None
            };
        let parsed_known_transform = irm.is_some() || encryption.is_some();
        let mut end_user_licenses = Vec::new();
        for (child_name, entry_type) in child_entries {
            if child_name == PRIMARY {
                if entry_type != 2 {
                    return Err(invalid("transform Primary entry is not a stream"));
                }
                continue;
            }
            if child_name.starts_with("EUL-") {
                if entry_type != 2 {
                    return Err(invalid("end-user license entry is not a stream"));
                }
                end_user_licenses.push(parse_license(
                    &child_name,
                    &read_stream(ole, &[STORAGE, "TransformInfo", &name, &child_name])?,
                )?);
            } else if entry_type != STGTY_STREAM && entry_type != STGTY_STORAGE {
                return Err(invalid(format!(
                    "unsupported transform-storage entry '{child_name}'"
                )));
            }
        }
        if irm.is_some() && end_user_licenses.is_empty() {
            return Err(invalid(format!(
                "IRM transform '{name}' has no end-user license stream"
            )));
        }
        transforms.push(Transform {
            name,
            header,
            irm,
            encryption,
            end_user_licenses,
            opaque_tail: if parsed_known_transform {
                Vec::new()
            } else {
                bytes[consumed..].to_vec()
            },
        });
    }

    validate_graph(ole, &map, &definitions, &transforms)?;
    let irm = classify_irm(&map, &definitions, &transforms)?;
    let (label_info, labels) = if ole.exists(&[STORAGE, "TransformInfo", "LabelInfo"]) {
        let bytes = read_stream(ole, &[STORAGE, "TransformInfo", "LabelInfo"])?;
        let labels = labels::parse(&bytes)
            .map_err(|error| invalid(format!("LabelInfo validation failed: {error}")))?;
        (Some(bytes), Some(labels))
    } else {
        (None, None)
    };
    if labels.is_some()
        && !irm.as_ref().is_some_and(|profile| {
            transforms.iter().any(|transform| {
                transform.name == profile.transform_name
                    && transform
                        .irm
                        .as_ref()
                        .is_some_and(|metadata| metadata.publishing_license.is_some())
            })
        })
    {
        return Err(invalid(
            "LabelInfo requires an IRM transform with a publishing license",
        ));
    }
    let summary_information_integrity =
        inspect_integrity(ole, SUMMARY_HASH_STREAM, SUMMARY_STREAM)?;
    let document_summary_information_integrity =
        inspect_integrity(ole, DOCUMENT_SUMMARY_HASH_STREAM, DOCUMENT_SUMMARY_STREAM)?;
    if (summary_information_integrity.is_some() || document_summary_information_integrity.is_some())
        && irm.is_none()
        && !transforms
            .iter()
            .any(|transform| transform.encryption.is_some())
    {
        return Err(invalid(
            "encrypted property hash stream is present without an encryption or IRM transform",
        ));
    }
    validate_custom_xml_promotion(custom_xml_data_store.as_ref(), irm.as_ref())?;
    Ok(Some(Graph {
        version,
        map,
        definitions,
        transforms,
        irm,
        label_info,
        labels,
        summary_information_integrity,
        document_summary_information_integrity,
        custom_xml_data_store,
    }))
}

fn validate_custom_xml_promotion(store: Option<&Store>, irm: Option<&Irm>) -> Result<(), Error> {
    if store.is_some_and(|inner| inner.promotion != Promotion::Unspecified) && irm.is_none() {
        return Err(invalid(
            "MsoDataStore promotion marker requires an IRM data space",
        ));
    }
    Ok(())
}

/// Open an OLE compound file and inspect its `DataSpaces` graph.
///
/// # Errors
///
/// Returns [`Error`] when the bytes are not a valid OLE container or its `DataSpaces` graph cannot be read or validated.
pub fn inspect_bytes(bytes: &[u8]) -> Result<Option<Graph>, Error> {
    let mut ole = OleFile::open(std::io::Cursor::new(bytes))?;
    inspect(&mut ole)
}

fn inspect_integrity<R: Read + Seek>(
    ole: &mut OleFile<R>,
    info_stream: &str,
    property_stream: &str,
) -> Result<Option<Integrity>, Error> {
    if !ole.exists(&[info_stream]) {
        return Ok(None);
    }
    if !ole.exists(&[property_stream]) {
        return Err(invalid(format!(
            "{info_stream} is present without {property_stream}"
        )));
    }
    let info = integrity::parse(&read_stream(ole, &[info_stream])?)
        .map_err(|error| invalid(format!("{info_stream} is malformed: {error}")))?;
    let valid = integrity::verify(&info, &read_stream(ole, &[property_stream])?);
    Ok(Some(Integrity { info, valid }))
}

fn validate_graph<R: Read + Seek>(
    ole: &OleFile<R>,
    map: &Map,
    definitions: &[NamedDefinition],
    transforms: &[Transform],
) -> Result<(), Error> {
    let components = collect_components(ole)?;
    validate_graph_model(map, definitions, transforms)?;
    validate_component_references(map, &components)
}

fn classify_irm(
    map: &Map,
    definitions: &[NamedDefinition],
    transforms: &[Transform],
) -> Result<Option<Irm>, Error> {
    if let Some(entry) = map
        .entries
        .iter()
        .find(|entry| entry.data_space_name == "DRMEncryptedDataSpace")
    {
        if map.entries.len() != 1 {
            return Err(invalid("OOXML IRM requires exactly one DataSpaceMap entry"));
        }
        require_single_stream(entry, "EncryptedPackage")?;
        require_definition(
            definitions,
            "DRMEncryptedDataSpace",
            &["DRMEncryptedTransform"],
        )?;
        require_drm_transform(transforms, "DRMEncryptedTransform")?;
        return Ok(Some(Irm {
            document_kind: DocumentKind::Ooxml,
            protected_stream: "EncryptedPackage".to_string(),
            viewer_content_stream: None,
            transform_name: "DRMEncryptedTransform".to_string(),
        }));
    }
    if let Some(entry) = map
        .entries
        .iter()
        .find(|entry| entry.data_space_name == "0x09DRMDataSpace")
    {
        require_single_stream(entry, "0x09DRMContent")?;
        require_definition(definitions, "0x09DRMDataSpace", &["0x09DRMTransform"])?;
        require_drm_transform(transforms, "0x09DRMTransform")?;
        let viewer_content_stream = if let Some(viewer) = map
            .entries
            .iter()
            .find(|candidate| candidate.data_space_name == "0x09LZXDRMDataSpace")
        {
            require_single_stream(viewer, "0x09DRMViewerContent")?;
            require_definition(
                definitions,
                "0x09LZXDRMDataSpace",
                &["0x09DRMTransform", "0x09LZXTransform"],
            )?;
            require_named_transform(transforms, "0x09LZXTransform", LZX_ID, LZX_NAME)?;
            Some("0x09DRMViewerContent".to_string())
        } else {
            None
        };
        let expected_entry_count = if viewer_content_stream.is_some() {
            2
        } else {
            1
        };
        if map.entries.len() != expected_entry_count {
            return Err(invalid(
                "binary IRM contains an unexpected DataSpaceMap entry",
            ));
        }
        return Ok(Some(Irm {
            document_kind: DocumentKind::LegacyBinary,
            protected_stream: "0x09DRMContent".to_string(),
            viewer_content_stream,
            transform_name: "0x09DRMTransform".to_string(),
        }));
    }
    Ok(None)
}

fn require_single_stream(entry: &MapEntry, expected: &str) -> Result<(), Error> {
    if entry.references.as_slice()
        != [Reference {
            kind: ReferenceKind::Stream,
            component: expected.to_string(),
        }]
    {
        return Err(invalid(format!(
            "IRM data space '{}' does not reference exactly '{expected}'",
            entry.data_space_name
        )));
    }
    Ok(())
}

fn require_definition(
    definitions: &[NamedDefinition],
    name: &str,
    expected: &[&str],
) -> Result<(), Error> {
    let definition = definitions
        .iter()
        .find(|definition| definition.name == name)
        .ok_or_else(|| invalid(format!("missing IRM data space '{name}'")))?;
    if definition
        .definition
        .transforms
        .iter()
        .map(String::as_str)
        .ne(expected.iter().copied())
    {
        return Err(invalid(format!(
            "IRM data space '{name}' has the wrong transform chain"
        )));
    }
    Ok(())
}

fn require_drm_transform(transforms: &[Transform], name: &str) -> Result<(), Error> {
    let transform = transforms
        .iter()
        .find(|transform| transform.name == name)
        .ok_or_else(|| invalid(format!("missing IRM transform '{name}'")))?;
    validate_drm_header(&transform.header)
}

fn require_named_transform(
    transforms: &[Transform],
    name: &str,
    transform_id: &str,
    transform_name: &str,
) -> Result<(), Error> {
    let transform = transforms
        .iter()
        .find(|transform| transform.name == name)
        .ok_or_else(|| invalid(format!("missing transform '{name}'")))?;
    if transform.header.transform_id != transform_id
        || transform.header.transform_name != transform_name
        || transform.header.reader != Version::V1_0
        || transform.header.updater != Version::V1_0
        || transform.header.writer != Version::V1_0
    {
        return Err(invalid(format!("transform '{name}' has an invalid header")));
    }
    Ok(())
}

fn validate_version_info(value: &VersionInfo) -> Result<(), Error> {
    if value.feature_identifier != FEATURE
        || value.reader != Version::V1_0
        || value.updater != Version::V1_0
        || value.writer != Version::V1_0
    {
        return Err(invalid("unsupported DataSpaceVersionInfo"));
    }
    Ok(())
}

fn validate_map(value: &Map) -> Result<(), Error> {
    if value.entries.is_empty() || value.entries.len() > MAX_ENTRIES {
        return Err(invalid("DataSpaceMap entry count is out of bounds"));
    }
    let mut names = std::collections::HashSet::with_capacity(value.entries.len());
    for entry in &value.entries {
        validate_name(&entry.data_space_name, "data space name")?;
        if !names.insert(entry.data_space_name.as_str()) {
            return Err(invalid("duplicate data space name"));
        }
        if entry.references.is_empty() || entry.references.len() > MAX_ENTRIES {
            return Err(invalid("reference component count is out of bounds"));
        }
        for reference in &entry.references {
            validate_name(&reference.component, "reference component")?;
        }
    }
    Ok(())
}

fn validate_definition(value: &Definition) -> Result<(), Error> {
    if value.transforms.is_empty() || value.transforms.len() > MAX_ENTRIES {
        return Err(invalid("transform reference count is out of bounds"));
    }
    let mut names = std::collections::HashSet::with_capacity(value.transforms.len());
    for transform in &value.transforms {
        validate_name(transform, "transform reference")?;
        if !names.insert(transform.as_str()) {
            return Err(invalid("duplicate transform reference"));
        }
    }
    Ok(())
}

fn validate_transform_header(value: &Header) -> Result<(), Error> {
    validate_name(&value.transform_id, "transform identifier")?;
    validate_name(&value.transform_name, "transform name")?;
    Ok(())
}

fn validate_drm_header(value: &Header) -> Result<(), Error> {
    if value.transform_id != DRM_ID
        || value.transform_name != DRM_NAME
        || value.reader != Version::V1_0
        || value.updater != Version::V1_0
        || value.writer != Version::V1_0
    {
        return Err(invalid("invalid IRM transform header"));
    }
    Ok(())
}

fn validate_encryption_header(value: &Header) -> Result<(), Error> {
    if value.transform_id != ENCRYPTION_ID
        || value.transform_name != ENCRYPTION_NAME
        || value.reader != Version::V1_0
        || value.updater != Version::V1_0
        || value.writer != Version::V1_0
    {
        return Err(invalid("invalid encryption transform header"));
    }
    Ok(())
}

fn validate_encryption_transform(value: &EncryptionTransform) -> Result<(), Error> {
    validate_encryption_header(&value.header)?;
    if value.encryption_block_size == 0 {
        return Err(invalid("encryption transform block size cannot be zero"));
    }
    if value
        .encryption_name
        .as_deref()
        .is_some_and(|name| name.is_empty() || name.len() > MAX_STRING_BYTES)
    {
        return Err(invalid("invalid encryption transform algorithm name"));
    }
    Ok(())
}

fn validate_name(value: &str, label: &str) -> Result<(), Error> {
    if value.is_empty()
        || value.len() > MAX_STRING_BYTES
        || value.chars().any(|character| character == '\0')
    {
        return Err(invalid(format!(
            "{label} is empty, too long, or contains NUL"
        )));
    }
    Ok(())
}

fn validate_eul_stream_name(value: &str) -> Result<(), Error> {
    let Some(encoded_guid) = value.strip_prefix("EUL-") else {
        return Err(invalid("end-user license stream name lacks EUL- prefix"));
    };
    if encoded_guid.len() != 26
        || !encoded_guid
            .bytes()
            .all(|byte| byte.is_ascii_alphanumeric())
    {
        return Err(invalid(
            "end-user license stream name does not contain a 26-character base-32 GUID",
        ));
    }
    Ok(())
}

fn validate_inert_xml(value: &str, label: &str) -> Result<(), Error> {
    use quick_xml::Reader;
    use quick_xml::events::Event;

    let mut reader = Reader::from_str(value);
    reader.config_mut().trim_text(false);
    let mut depth = 0usize;
    let mut roots = 0usize;
    loop {
        match reader.read_event() {
            Ok(Event::Start(_)) => {
                if depth == 0 {
                    roots += 1;
                }
                depth = depth
                    .checked_add(1)
                    .ok_or_else(|| invalid(format!("{label} XML depth overflow")))?;
                if depth > MAX_XML_DEPTH {
                    return Err(invalid(format!("{label} XML is too deeply nested")));
                }
            },
            Ok(Event::Empty(_)) => {
                if depth == 0 {
                    roots += 1;
                }
            },
            Ok(Event::End(_)) => {
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| invalid(format!("{label} XML is unbalanced")))?;
            },
            Ok(Event::DocType(_)) => {
                return Err(invalid(format!("{label} XML contains a forbidden DTD")));
            },
            Ok(Event::Text(text)) if depth == 0 => {
                if !text
                    .decode()
                    .map_err(|error| invalid(format!("{label} XML text is invalid: {error}")))?
                    .trim()
                    .is_empty()
                {
                    return Err(invalid(format!("{label} XML has text outside its root")));
                }
            },
            Ok(Event::CData(text)) if depth == 0 => {
                if !text
                    .decode()
                    .map_err(|error| invalid(format!("{label} XML CDATA is invalid: {error}")))?
                    .trim()
                    .is_empty()
                {
                    return Err(invalid(format!("{label} XML has CDATA outside its root")));
                }
            },
            Ok(Event::Eof) => break,
            Ok(_) => {},
            Err(error) => return Err(invalid(format!("{label} XML is malformed: {error}"))),
        }
    }
    if roots != 1 || depth != 0 {
        return Err(invalid(format!(
            "{label} XML must contain exactly one complete root"
        )));
    }
    Ok(())
}

fn bounded_count(value: u32, label: &str) -> Result<usize, Error> {
    let count =
        usize::try_from(value).map_err(|_err| invalid(format!("{label} overflows usize")))?;
    if count > MAX_ENTRIES {
        return Err(invalid(format!("{label} exceeds {MAX_ENTRIES}")));
    }
    Ok(count)
}

fn write_count(output: &mut Vec<u8>, count: usize, label: &str) -> Result<(), Error> {
    if count > MAX_ENTRIES {
        return Err(invalid(format!("{label} exceeds {MAX_ENTRIES}")));
    }
    output.extend_from_slice(
        &u32::try_from(count)
            .map_err(|_err| invalid(format!("{label} exceeds u32")))?
            .to_le_bytes(),
    );
    Ok(())
}

fn require_u32(value: u32, expected: u32, label: &str) -> Result<(), Error> {
    if value != expected {
        return Err(invalid(format!(
            "{label} is {value:#010X}, expected {expected:#010X}"
        )));
    }
    Ok(())
}

fn write_version(output: &mut Vec<u8>, version: Version) {
    output.extend_from_slice(&version.major.to_le_bytes());
    output.extend_from_slice(&version.minor.to_le_bytes());
}

fn write_unicode_lpp4(output: &mut Vec<u8>, value: &str) -> Result<(), Error> {
    validate_name(value, "UNICODE-LP-P4 string")?;
    let units = value.encode_utf16().collect::<Vec<_>>();
    let byte_len = units
        .len()
        .checked_mul(2)
        .ok_or_else(|| invalid("UNICODE-LP-P4 length overflow"))?;
    output.extend_from_slice(
        &u32::try_from(byte_len)
            .map_err(|_err| invalid("UNICODE-LP-P4 length exceeds u32"))?
            .to_le_bytes(),
    );
    for unit in units {
        output.extend_from_slice(&unit.to_le_bytes());
    }
    if byte_len % 4 == 2 {
        output.extend_from_slice(&[0, 0]);
    }
    Ok(())
}

fn write_utf8_lpp4(output: &mut Vec<u8>, value: Option<&str>) -> Result<(), Error> {
    let Some(text) = value else {
        output.extend_from_slice(&0u32.to_le_bytes());
        return Ok(());
    };
    if text.len() > MAX_STRING_BYTES || text.contains('\0') {
        return Err(invalid("UTF-8-LP-P4 string is too long or contains NUL"));
    }
    output.extend_from_slice(
        &u32::try_from(text.len())
            .map_err(|_err| invalid("UTF-8-LP-P4 length exceeds u32"))?
            .to_le_bytes(),
    );
    output.extend_from_slice(text.as_bytes());
    output.resize(output.len().next_multiple_of(4), 0);
    Ok(())
}

fn read_stream<R: Read + Seek>(ole: &mut OleFile<R>, path: &[&str]) -> Result<Vec<u8>, Error> {
    let bytes = ole.open_stream(path)?;
    if bytes.len() > MAX_STREAM_BYTES {
        return Err(invalid(format!(
            "stream '{}' exceeds {MAX_STREAM_BYTES} bytes",
            path.join("/")
        )));
    }
    Ok(bytes)
}

fn invalid(message: impl Into<String>) -> Error {
    Error::Invalid(message.into())
}

#[cfg(test)]
mod tests {
    #![allow(
        clippy::unwrap_used,
        reason = "test code panics on failure; unwrap keeps assertions concise"
    )]
    use super::*;
    use litchi_cfb::OleWriter;
    use std::io::Cursor;

    fn drm_header() -> Header {
        Header {
            transform_id: DRM_ID.to_string(),
            transform_name: DRM_NAME.to_string(),
            reader: Version::V1_0,
            updater: Version::V1_0,
            writer: Version::V1_0,
        }
    }

    fn editable_package() -> Vec<u8> {
        let map = write_map(&Map {
            entries: vec![MapEntry {
                references: vec![Reference {
                    kind: ReferenceKind::Stream,
                    component: "Payload".to_string(),
                }],
                data_space_name: "DataSpace".to_string(),
            }],
        })
        .unwrap();
        let definition = write_definition(&Definition {
            transforms: vec!["FirstTransform".to_string()],
        })
        .unwrap();
        let first = Header {
            transform_id: "{11111111-1111-1111-1111-111111111111}".to_string(),
            transform_name: "FirstTransform".to_string(),
            reader: Version::V1_0,
            updater: Version::V1_0,
            writer: Version::V1_0,
        };
        let second = Header {
            transform_id: "{22222222-2222-2222-2222-222222222222}".to_string(),
            transform_name: "SecondTransform".to_string(),
            reader: Version::V1_0,
            updater: Version::V1_0,
            writer: Version::V1_0,
        };
        let mut writer = OleWriter::new();
        writer
            .create_stream(&["Payload"], b"protected-but-inert")
            .unwrap();
        writer
            .create_stream(&["Unrelated"], b"preserve-me-byte-for-byte")
            .unwrap();
        writer.create_storage(&["EmptyStorage"]).unwrap();
        writer.create_storage(&[STORAGE]).unwrap();
        writer.create_storage(&[STORAGE, "DataSpaceInfo"]).unwrap();
        writer.create_storage(&[STORAGE, "TransformInfo"]).unwrap();
        writer
            .create_storage(&[STORAGE, "TransformInfo", "FirstTransform"])
            .unwrap();
        writer
            .create_storage(&[STORAGE, "TransformInfo", "SecondTransform"])
            .unwrap();
        writer
            .create_storage(&[STORAGE, "TransformInfo", "FirstTransform", "OpaqueStorage"])
            .unwrap();
        writer
            .create_stream(
                &[
                    STORAGE,
                    "TransformInfo",
                    "FirstTransform",
                    "OpaqueStorage",
                    "OpaqueStream",
                ],
                b"opaque-transform-child",
            )
            .unwrap();
        writer
            .create_stream(
                &[STORAGE, "Version"],
                &write_version_info(&VersionInfo::default()).unwrap(),
            )
            .unwrap();
        writer
            .create_stream(&[STORAGE, "DataSpaceMap"], &map)
            .unwrap();
        writer
            .create_stream(&[STORAGE, "DataSpaceInfo", "DataSpace"], &definition)
            .unwrap();
        writer
            .create_stream(
                &[STORAGE, "TransformInfo", "FirstTransform", PRIMARY],
                &write_transform_header(&first).unwrap(),
            )
            .unwrap();
        writer
            .create_stream(
                &[STORAGE, "TransformInfo", "SecondTransform", PRIMARY],
                &write_transform_header(&second).unwrap(),
            )
            .unwrap();
        let mut bytes = Cursor::new(Vec::new());
        writer.write_to(&mut bytes).unwrap();
        bytes.into_inner()
    }

    fn encryption_package() -> Vec<u8> {
        let map = write_map(&Map {
            entries: vec![MapEntry {
                references: vec![Reference {
                    kind: ReferenceKind::Stream,
                    component: "EncryptedPackage".to_string(),
                }],
                data_space_name: "StrongEncryptionDataSpace".to_string(),
            }],
        })
        .unwrap();
        let definition = write_definition(&Definition {
            transforms: vec!["StrongEncryptionTransform".to_string()],
        })
        .unwrap();
        let encryption = EncryptionTransform {
            header: Header {
                transform_id: ENCRYPTION_ID.to_string(),
                transform_name: ENCRYPTION_NAME.to_string(),
                reader: Version::V1_0,
                updater: Version::V1_0,
                writer: Version::V1_0,
            },
            encryption_name: None,
            encryption_block_size: 16,
            cipher_mode: 0,
        };
        let mut writer = OleWriter::new();
        writer
            .create_stream(&["EncryptedPackage"], b"encrypted-payload")
            .unwrap();
        writer.create_storage(&[STORAGE]).unwrap();
        writer.create_storage(&[STORAGE, "DataSpaceInfo"]).unwrap();
        writer.create_storage(&[STORAGE, "TransformInfo"]).unwrap();
        writer
            .create_storage(&[STORAGE, "TransformInfo", "StrongEncryptionTransform"])
            .unwrap();
        writer
            .create_stream(
                &[STORAGE, "Version"],
                &write_version_info(&VersionInfo::default()).unwrap(),
            )
            .unwrap();
        writer
            .create_stream(&[STORAGE, "DataSpaceMap"], &map)
            .unwrap();
        writer
            .create_stream(
                &[STORAGE, "DataSpaceInfo", "StrongEncryptionDataSpace"],
                &definition,
            )
            .unwrap();
        writer
            .create_stream(
                &[
                    STORAGE,
                    "TransformInfo",
                    "StrongEncryptionTransform",
                    PRIMARY,
                ],
                &write_encryption_transform(&encryption).unwrap(),
            )
            .unwrap();
        let mut bytes = Cursor::new(Vec::new());
        writer.write_to(&mut bytes).unwrap();
        bytes.into_inner()
    }

    #[test]
    fn core_streams_round_trip() {
        let version = VersionInfo::default();
        assert_eq!(
            parse_version_info(&write_version_info(&version).unwrap()).unwrap(),
            version
        );
        let map = Map {
            entries: vec![MapEntry {
                references: vec![Reference {
                    kind: ReferenceKind::Stream,
                    component: "EncryptedPackage".to_string(),
                }],
                data_space_name: "DRMEncryptedDataSpace".to_string(),
            }],
        };
        assert_eq!(parse_map(&write_map(&map).unwrap()).unwrap(), map);
        let definition = Definition {
            transforms: vec!["DRMEncryptedTransform".to_string()],
        };
        assert_eq!(
            parse_definition(&write_definition(&definition).unwrap()).unwrap(),
            definition
        );
    }

    #[test]
    fn irm_transform_round_trip_preserves_license() {
        let transform = IrmTransform {
            header: drm_header(),
            publishing_license: Some("<XrML>inert</XrML>".to_string()),
        };
        assert_eq!(
            parse_irm_transform(&write_irm_transform(&transform).unwrap()).unwrap(),
            transform
        );
    }

    #[test]
    fn encryption_transform_round_trip_preserves_typed_parameters() {
        let transform = EncryptionTransform {
            header: Header {
                transform_id: ENCRYPTION_ID.to_string(),
                transform_name: ENCRYPTION_NAME.to_string(),
                reader: Version::V1_0,
                updater: Version::V1_0,
                writer: Version::V1_0,
            },
            encryption_name: None,
            encryption_block_size: 16,
            cipher_mode: 0,
        };
        assert_eq!(
            parse_encryption_transform(&write_encryption_transform(&transform).unwrap()).unwrap(),
            transform
        );
    }

    #[test]
    fn end_user_license_round_trip_preserves_inert_xml() {
        let license = License {
            stream_name: "EUL-ETRHA1143ZLUDD412YTI3M5CTZ".to_string(),
            encoded_license_id: "VwBpAG4AZABvAHcAOgB1AHMAZQByAEA".to_string(),
            certificate_chain: Some("<?xml version=\"1.0\"?><certificatechain/>".to_string()),
        };
        assert_eq!(
            parse_license(&license.stream_name, &write_license(&license).unwrap()).unwrap(),
            license
        );
    }

    #[test]
    fn rejects_malformed_lengths_padding_counts_and_drm_identity() {
        let mut version = write_version_info(&VersionInfo::default()).unwrap();
        version[0] = 3;
        assert!(parse_version_info(&version).is_err());

        let map = Map {
            entries: Vec::new(),
        };
        assert!(write_map(&map).is_err());

        let mut transform = write_irm_transform(&IrmTransform {
            header: drm_header(),
            publishing_license: Some("<XrML/>".to_string()),
        })
        .unwrap();
        transform[4] = 2;
        assert!(parse_irm_transform(&transform).is_err());
        assert!(
            write_irm_transform(&IrmTransform {
                header: drm_header(),
                publishing_license: Some("<!DOCTYPE x><x/>".to_string()),
            })
            .is_err()
        );
    }

    #[test]
    fn inspects_and_classifies_complete_ooxml_irm_graph() {
        let map = write_map(&Map {
            entries: vec![MapEntry {
                references: vec![Reference {
                    kind: ReferenceKind::Stream,
                    component: "EncryptedPackage".to_string(),
                }],
                data_space_name: "DRMEncryptedDataSpace".to_string(),
            }],
        })
        .unwrap();
        let definition = write_definition(&Definition {
            transforms: vec!["DRMEncryptedTransform".to_string()],
        })
        .unwrap();
        let primary = write_irm_transform(&IrmTransform {
            header: drm_header(),
            publishing_license: Some("<XrML/>".to_string()),
        })
        .unwrap();
        let end_user_license = License {
            stream_name: "EUL-ETRHA1143ZLUDD412YTI3M5CTZ".to_string(),
            encoded_license_id: "VwBpAG4AZABvAHcAOgB1AHMAZQByAEA".to_string(),
            certificate_chain: Some("<certificatechain/>".to_string()),
        };
        let mut writer = OleWriter::new();
        writer
            .create_stream(&["EncryptedPackage"], &[0; 16])
            .unwrap();
        writer.create_storage(&[STORAGE]).unwrap();
        writer.create_storage(&[STORAGE, "DataSpaceInfo"]).unwrap();
        writer.create_storage(&[STORAGE, "TransformInfo"]).unwrap();
        writer
            .create_storage(&[STORAGE, "TransformInfo", "DRMEncryptedTransform"])
            .unwrap();
        writer
            .create_stream(&[STORAGE, "DataSpaceMap"], &map)
            .unwrap();
        writer
            .create_stream(
                &[STORAGE, "DataSpaceInfo", "DRMEncryptedDataSpace"],
                &definition,
            )
            .unwrap();
        writer
            .create_stream(
                &[STORAGE, "TransformInfo", "DRMEncryptedTransform", PRIMARY],
                &primary,
            )
            .unwrap();
        writer
            .create_stream(
                &[
                    STORAGE,
                    "TransformInfo",
                    "DRMEncryptedTransform",
                    &end_user_license.stream_name,
                ],
                &write_license(&end_user_license).unwrap(),
            )
            .unwrap();
        writer
            .create_stream(
                &[STORAGE, "Version"],
                &write_version_info(&VersionInfo::default()).unwrap(),
            )
            .unwrap();
        let label_info = format!(
            "<?xml version=\"1.0\" encoding=\"utf-8\"?><clbl:labelList xmlns:clbl=\"{}\"><clbl:label id=\"aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee\" enabled=\"1\" method=\"Standard\" siteId=\"12345678-1234-5678-90ab-1234567890ab\" contentBits=\"8\" removed=\"0\"/></clbl:labelList>",
            labels::NAMESPACE
        );
        writer
            .create_stream(
                &[STORAGE, "TransformInfo", "LabelInfo"],
                label_info.as_bytes(),
            )
            .unwrap();
        let summary_information = b"public summary property set";
        writer
            .create_stream(&[SUMMARY_STREAM], summary_information)
            .unwrap();
        writer
            .create_stream(
                &[SUMMARY_HASH_STREAM],
                &integrity::write(integrity::crc32(summary_information), &[]),
            )
            .unwrap();
        let custom_properties = litchi_ole_common::custom_xml::Properties {
            item_id: "{11111111-2222-3333-4444-555555555555}".parse().unwrap(),
            schema_references: vec!["urn:test".to_string()],
        };
        let custom_item = litchi_ole_common::custom_xml::Item::new(
            custom_properties.item_id.storage_name(),
            br#"<test xmlns="urn:test"/>"#.to_vec(),
            custom_properties,
        )
        .unwrap();
        litchi_ole_common::custom_xml::write(
            &mut writer,
            &Store::new(Promotion::Modified, vec![custom_item]).unwrap(),
        )
        .unwrap();
        let mut bytes = Cursor::new(Vec::new());
        writer.write_to(&mut bytes).unwrap();

        let mut ole = OleFile::open(Cursor::new(bytes.into_inner())).unwrap();
        let graph = inspect(&mut ole).unwrap().unwrap();
        let irm = graph.irm.unwrap();
        assert_eq!(irm.document_kind, DocumentKind::Ooxml);
        assert_eq!(irm.protected_stream, "EncryptedPackage");
        assert_eq!(graph.transforms[0].end_user_licenses, [end_user_license]);
        assert_eq!(graph.labels.unwrap().labels.len(), 1);
        assert_eq!(
            graph.summary_information_integrity.unwrap().valid,
            Some(true)
        );
        assert!(graph.document_summary_information_integrity.is_none());
        let custom_xml = graph.custom_xml_data_store.unwrap();
        assert_eq!(custom_xml.promotion, Promotion::Modified);
        assert_eq!(custom_xml.items().len(), 1);
    }

    #[test]
    fn classifies_binary_irm_with_optional_viewer_chain() {
        let map = Map {
            entries: vec![
                MapEntry {
                    references: vec![Reference {
                        kind: ReferenceKind::Stream,
                        component: "0x09DRMContent".to_string(),
                    }],
                    data_space_name: "0x09DRMDataSpace".to_string(),
                },
                MapEntry {
                    references: vec![Reference {
                        kind: ReferenceKind::Stream,
                        component: "0x09DRMViewerContent".to_string(),
                    }],
                    data_space_name: "0x09LZXDRMDataSpace".to_string(),
                },
            ],
        };
        let definitions = vec![
            NamedDefinition {
                name: "0x09DRMDataSpace".to_string(),
                definition: Definition {
                    transforms: vec!["0x09DRMTransform".to_string()],
                },
            },
            NamedDefinition {
                name: "0x09LZXDRMDataSpace".to_string(),
                definition: Definition {
                    transforms: vec![
                        "0x09DRMTransform".to_string(),
                        "0x09LZXTransform".to_string(),
                    ],
                },
            },
        ];
        let transforms = vec![
            Transform {
                name: "0x09DRMTransform".to_string(),
                header: drm_header(),
                irm: None,
                encryption: None,
                end_user_licenses: Vec::new(),
                opaque_tail: Vec::new(),
            },
            Transform {
                name: "0x09LZXTransform".to_string(),
                header: Header {
                    transform_id: LZX_ID.to_string(),
                    transform_name: LZX_NAME.to_string(),
                    reader: Version::V1_0,
                    updater: Version::V1_0,
                    writer: Version::V1_0,
                },
                irm: None,
                encryption: None,
                end_user_licenses: Vec::new(),
                opaque_tail: Vec::new(),
            },
        ];

        let irm = classify_irm(&map, &definitions, &transforms)
            .unwrap()
            .unwrap();
        assert_eq!(irm.document_kind, DocumentKind::LegacyBinary);
        assert_eq!(
            irm.viewer_content_stream.as_deref(),
            Some("0x09DRMViewerContent")
        );
    }

    #[test]
    fn rejects_custom_xml_promotion_without_irm() {
        let store = Store::new(Promotion::Redundant, Vec::new()).unwrap();
        assert!(validate_custom_xml_promotion(Some(&store), None).is_err());

        let mut writer = OleWriter::new();
        litchi_ole_common::custom_xml::write(&mut writer, &store).unwrap();
        let mut bytes = Cursor::new(Vec::new());
        writer.write_to(&mut bytes).unwrap();
        let mut ole = OleFile::open(Cursor::new(bytes.into_inner())).unwrap();
        assert!(inspect(&mut ole).is_err());

        let unspecified = Store::default();
        assert!(validate_custom_xml_promotion(Some(&unspecified), None).is_ok());
    }

    #[test]
    fn snapshot_noop_replays_exact_dataspaces_source() {
        let mut ole = OleFile::open(Cursor::new(editable_package())).unwrap();
        let snapshot = Snapshot::from_ole(&mut ole).unwrap().unwrap();
        let commit = snapshot.edit().commit().unwrap();

        assert!(!commit.changed());
        assert!(commit.patch().is_noop());
        assert_eq!(commit.snapshot(), &snapshot);
        assert_eq!(commit.patch().apply(&snapshot).unwrap(), snapshot);
    }

    #[test]
    fn transaction_changes_definition_and_rejects_stale_patch_sources() {
        let mut ole = OleFile::open(Cursor::new(editable_package())).unwrap();
        let snapshot = Snapshot::from_ole(&mut ole).unwrap().unwrap();
        let mut transaction = snapshot.edit();
        transaction
            .set_definition(
                "DataSpace",
                Definition {
                    transforms: vec!["SecondTransform".to_string()],
                },
            )
            .unwrap();
        let commit = transaction.commit().unwrap();

        assert!(commit.changed());
        assert_eq!(commit.patch().changes().len(), 1);
        assert_eq!(
            commit.patch().changes()[0].path(),
            [
                STORAGE.to_string(),
                "DataSpaceInfo".to_string(),
                "DataSpace".to_string()
            ]
        );
        assert_eq!(commit.patch().apply(&snapshot).unwrap(), *commit.snapshot());
        assert!(commit.patch().apply(commit.snapshot()).is_err());
    }

    #[test]
    fn package_rebuild_preserves_unrelated_streams_and_empty_storages() {
        let mut ole = OleFile::open(Cursor::new(editable_package())).unwrap();
        let snapshot = Snapshot::from_ole(&mut ole).unwrap().unwrap();
        let mut transaction = snapshot.edit();
        transaction
            .set_definition(
                "DataSpace",
                Definition {
                    transforms: vec!["SecondTransform".to_string()],
                },
            )
            .unwrap();
        let commit = transaction.commit().unwrap();
        let mut rebuilt = Cursor::new(Vec::new());
        commit.patch().write_to(&mut ole, &mut rebuilt).unwrap();

        let mut output = OleFile::open(Cursor::new(rebuilt.into_inner())).unwrap();
        assert_eq!(
            output.open_stream(&["Payload"]).unwrap(),
            b"protected-but-inert"
        );
        assert_eq!(
            output.open_stream(&["Unrelated"]).unwrap(),
            b"preserve-me-byte-for-byte"
        );
        assert_eq!(
            output
                .open_stream(&[
                    STORAGE,
                    "TransformInfo",
                    "FirstTransform",
                    "OpaqueStorage",
                    "OpaqueStream",
                ])
                .unwrap(),
            b"opaque-transform-child"
        );
        assert!(output.directory_exists(&["EmptyStorage"]));
        let graph = inspect(&mut output).unwrap().unwrap();
        assert_eq!(
            graph.definitions[0].definition.transforms,
            ["SecondTransform".to_string()]
        );
    }

    #[test]
    fn encryption_metadata_edit_does_not_touch_encrypted_payload() {
        let mut ole = OleFile::open(Cursor::new(encryption_package())).unwrap();
        let snapshot = Snapshot::from_ole(&mut ole).unwrap().unwrap();
        let mut transaction = snapshot.edit();
        transaction
            .set_encryption_info("StrongEncryptionTransform", Some("AES".to_string()), 16, 0)
            .unwrap();
        let commit = transaction.commit().unwrap();
        assert_eq!(commit.patch().changes().len(), 1);
        assert_eq!(
            commit.patch().changes()[0].path(),
            [
                STORAGE.to_string(),
                "TransformInfo".to_string(),
                "StrongEncryptionTransform".to_string(),
                PRIMARY.to_string()
            ]
        );
        let mut rebuilt = Cursor::new(Vec::new());
        commit.patch().write_to(&mut ole, &mut rebuilt).unwrap();
        let mut output = OleFile::open(Cursor::new(rebuilt.into_inner())).unwrap();
        assert_eq!(
            output.open_stream(&["EncryptedPackage"]).unwrap(),
            b"encrypted-payload"
        );
        let graph = inspect(&mut output).unwrap().unwrap();
        assert_eq!(
            graph.transforms[0]
                .encryption
                .as_ref()
                .unwrap()
                .encryption_name
                .as_deref(),
            Some("AES")
        );
    }
}
