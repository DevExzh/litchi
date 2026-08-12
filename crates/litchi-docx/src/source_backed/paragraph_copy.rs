//! Exact source-backed editing for deliberately plain main-story paragraphs.
//!
//! This capability is intentionally narrower than the general document
//! transaction API. It accepts only a canonical main story whose body contains
//! direct plain paragraphs made from `w:r` and `w:t`. Paragraphs are copied or
//! removed as exact source fragments, while every other package member is
//! raw-copied by the OPC overlay publisher.

use std::io::Write;
use std::sync::Arc;

use litchi_core::{Position, SourceVersion};
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{BlobPart, Part, SourceArtifact};
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;
use sha2::{Digest as _, Sha256};
use thiserror::Error;

use crate::package::validate_document_main_content_type;
use crate::settings::DocumentSettings;

use super::Package;

const MAGIC: &[u8; 8] = b"LDXPCPY\0";
const REMOVAL_MAGIC: &[u8; 8] = b"LDXPREM\0";
const VERSION: u8 = 1;
const HEADER_BYTES: usize = 8 + 1 + 1 + (6 * 8) + 8 + 32 + 8;
const ABSOLUTE_MAX_DURABLE_BYTES: usize = 64 * 1024 * 1024;
const TRANSITIONAL_WORD: &[u8] = b"http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const STRICT_WORD: &[u8] = b"http://purl.oclc.org/ooxml/wordprocessingml/main";
const XML_NAMESPACE: &[u8] = b"http://www.w3.org/XML/1998/namespace";
const SIGNATURE_DIRECTORY: &str = "/_xmlsignatures/";

/// Result returned by exact plain-paragraph copy operations.
pub type Result<T> = std::result::Result<T, Error>;

/// Stable reason why a package or paragraph is outside this narrow closure.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum Refusal {
    /// The main document is not the ordinary `/word/document.xml` DOCX part.
    MainDocumentShape,
    /// The package contains digital-signature infrastructure.
    SignatureInfrastructure,
    /// The package is macro-enabled or contains a VBA relationship or part.
    MacroEnabled,
    /// An external relationship exists anywhere in the package.
    ExternalRelationship,
    /// A comments, notes, custom-XML, or alternative-format dependency exists.
    UnsupportedDependency,
    /// Document or write protection is enforced.
    Protection,
    /// The main story contains MCE, sections, tables, wrappers, or unknown XML.
    ComplexDocument,
    /// A direct paragraph contains anything other than plain runs and text.
    ComplexParagraph,
}

impl std::fmt::Display for Refusal {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter.write_str(match self {
            Self::MainDocumentShape => "main document is not the canonical ordinary DOCX part",
            Self::SignatureInfrastructure => "package contains digital-signature infrastructure",
            Self::MacroEnabled => "package is macro-enabled",
            Self::ExternalRelationship => "package contains an external relationship",
            Self::UnsupportedDependency => "package contains an unsupported story dependency",
            Self::Protection => "document or write protection is enforced",
            Self::ComplexDocument => "main story is not a direct plain-paragraph document",
            Self::ComplexParagraph => "paragraph is not composed only of plain runs and text",
        })
    }
}

/// A bounded exact-copy failure.
#[derive(Debug, Error)]
#[non_exhaustive]
pub enum Error {
    /// The underlying DOCX or OPC package is invalid.
    #[error(transparent)]
    Document(#[from] crate::Error),
    /// The requested source or insertion position is outside the source order.
    #[error("paragraph {kind} position {position} is out of bounds for length {len}")]
    OutOfBounds {
        /// Whether the position selects a source paragraph or insertion slot.
        kind: &'static str,
        /// Requested zero-based position.
        position: usize,
        /// Number of source paragraphs.
        len: usize,
    },
    /// The package cannot be represented by this exact one-part closure.
    #[error("plain paragraph edit refused: {0}")]
    Refused(Refusal),
    /// A configured finite resource ceiling was exceeded.
    #[error("plain paragraph edit {resource} limit exceeded: {actual} > {max}")]
    Limit {
        /// Bounded resource.
        resource: &'static str,
        /// Maximum accepted value.
        max: usize,
        /// Observed or requested value.
        actual: usize,
    },
    /// A patch was applied to bytes other than its exact captured source.
    #[error("plain paragraph edit patch source is stale")]
    StaleSource,
    /// A deterministic durable patch envelope was malformed or non-canonical.
    #[error("invalid durable plain paragraph patch")]
    InvalidDurable,
    /// A bounded allocation failed before publication began.
    #[error("plain paragraph edit allocation failed for {resource}: {source}")]
    Allocation {
        /// Allocation being attempted.
        resource: &'static str,
        /// Allocator failure.
        #[source]
        source: std::collections::TryReserveError,
    },
}

/// Finite limits for scanning, editing, output, and durable patches.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Limits {
    max_xml_bytes: usize,
    max_paragraphs: usize,
    max_events: usize,
    max_depth: usize,
    max_output_bytes: usize,
    max_durable_bytes: usize,
}

impl Limits {
    /// Construct a validated finite policy.
    pub fn new(
        max_xml_bytes: usize,
        max_paragraphs: usize,
        max_events: usize,
        max_depth: usize,
        max_output_bytes: usize,
        max_durable_bytes: usize,
    ) -> Result<Self> {
        let values = [
            ("XML bytes", max_xml_bytes),
            ("paragraphs", max_paragraphs),
            ("XML events", max_events),
            ("XML depth", max_depth),
            ("output bytes", max_output_bytes),
            ("durable patch bytes", max_durable_bytes),
        ];
        if values.iter().any(|(_, value)| *value == 0) {
            return Err(Error::InvalidDurable);
        }
        if max_durable_bytes > ABSOLUTE_MAX_DURABLE_BYTES {
            return Err(Error::InvalidDurable);
        }
        Ok(Self {
            max_xml_bytes,
            max_paragraphs,
            max_events,
            max_depth,
            max_output_bytes,
            max_durable_bytes,
        })
    }

    /// Maximum accepted main-document XML bytes.
    #[must_use]
    pub const fn max_xml_bytes(self) -> usize {
        self.max_xml_bytes
    }

    /// Maximum accepted direct paragraphs.
    #[must_use]
    pub const fn max_paragraphs(self) -> usize {
        self.max_paragraphs
    }

    /// Maximum accepted XML parser events.
    #[must_use]
    pub const fn max_events(self) -> usize {
        self.max_events
    }

    /// Maximum accepted XML nesting depth.
    #[must_use]
    pub const fn max_depth(self) -> usize {
        self.max_depth
    }

    /// Maximum accepted projected main-document XML bytes.
    #[must_use]
    pub const fn max_output_bytes(self) -> usize {
        self.max_output_bytes
    }

    /// Maximum accepted deterministic patch bytes.
    #[must_use]
    pub const fn max_durable_bytes(self) -> usize {
        self.max_durable_bytes
    }
}

impl Default for Limits {
    fn default() -> Self {
        Self {
            max_xml_bytes: 16 * 1024 * 1024,
            max_paragraphs: 65_536,
            max_events: 1_000_000,
            max_depth: 16,
            max_output_bytes: 32 * 1024 * 1024,
            max_durable_bytes: 64 * 1024 * 1024,
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
struct Range {
    start: usize,
    end: usize,
}

#[derive(Debug, Clone)]
struct Layout {
    paragraphs: Arc<[Range]>,
    body_end: usize,
}

/// Immutable plain main-story bytes bound to an exact source artifact.
#[derive(Debug, Clone)]
pub struct Snapshot {
    xml: Arc<Vec<u8>>,
    layout: Layout,
    source_version: SourceVersion,
    artifact_fingerprint: [u8; 32],
    limits: Limits,
}

impl Snapshot {
    /// Borrow the exact main-document bytes.
    #[must_use]
    pub fn xml_bytes(&self) -> &[u8] {
        self.xml.as_slice()
    }

    /// Return the number of direct source-order paragraphs.
    #[must_use]
    pub fn paragraph_count(&self) -> usize {
        self.layout.paragraphs.len()
    }

    /// Return the exact package source version captured with this snapshot.
    #[must_use]
    pub const fn source_version(&self) -> SourceVersion {
        self.source_version
    }

    /// Start an isolated exact-fragment edit.
    #[must_use]
    pub fn edit(&self) -> Edit {
        Edit {
            base: self.clone(),
            projected: self.clone(),
            operation: None,
        }
    }

    /// Start an isolated exact-fragment removal edit.
    #[must_use]
    pub fn removal_edit(&self) -> RemovalEdit {
        RemovalEdit {
            base: self.clone(),
            projected: self.clone(),
            position: None,
        }
    }

    fn with_xml(&self, xml: Vec<u8>) -> Result<Self> {
        let layout = scan_document(&xml, self.limits)?;
        Ok(Self {
            xml: Arc::new(xml),
            layout,
            source_version: self.source_version,
            artifact_fingerprint: self.artifact_fingerprint,
            limits: self.limits,
        })
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
struct CopyOperation {
    source: Position,
    before: Position,
}

/// One staged exact plain-paragraph copy.
#[derive(Debug, Clone)]
pub struct Edit {
    base: Snapshot,
    projected: Snapshot,
    operation: Option<CopyOperation>,
}

impl Edit {
    /// Borrow the unchanged source snapshot.
    #[must_use]
    pub const fn source(&self) -> &Snapshot {
        &self.base
    }

    /// Borrow the projected snapshot.
    #[must_use]
    pub const fn projected(&self) -> &Snapshot {
        &self.projected
    }

    /// Copy one source-order paragraph before one source-order insertion slot.
    ///
    /// `source` selects `0..paragraph_count`; `before` selects
    /// `0..=paragraph_count`. Both positions always refer to the immutable
    /// source order. Thus `before == paragraph_count` appends immediately
    /// before `</w:body>`. The exact selected `w:p` bytes are duplicated.
    pub fn copy_plain_paragraph(
        &mut self,
        source: Position,
        before: Position,
    ) -> Result<&mut Self> {
        if self.operation.is_some() {
            return Err(Error::Limit {
                resource: "operations",
                max: 1,
                actual: 2,
            });
        }
        let projected = copy_fragment(&self.base, source, before)?;
        self.operation = Some(CopyOperation { source, before });
        self.projected = projected;
        Ok(self)
    }

    /// Finish the edit as an exact reversible patch.
    #[must_use]
    pub fn commit(self) -> Commit {
        let patch = Patch {
            before: Arc::clone(&self.base.xml),
            after: Arc::clone(&self.projected.xml),
            expected_fingerprint: self.base.artifact_fingerprint,
            source_version: Some(self.base.source_version),
            limits: self.base.limits,
            operation: self.operation,
            direction: Direction::Forward,
        };
        Commit {
            projected: self.projected,
            patch,
        }
    }
}

/// Completed exact-copy edit.
#[derive(Debug, Clone)]
pub struct Commit {
    projected: Snapshot,
    patch: Patch,
}

impl Commit {
    /// Borrow the projected main-story snapshot.
    #[must_use]
    pub const fn projected(&self) -> &Snapshot {
        &self.projected
    }

    /// Borrow the reversible exact patch.
    #[must_use]
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Return the deterministic semantic effect.
    #[must_use]
    pub fn effect_report(&self) -> EffectReport {
        self.patch.effect_report()
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum Direction {
    Forward,
    Inverse,
}

/// Exact, deterministic, reversible main-document patch.
#[derive(Debug, Clone)]
pub struct Patch {
    before: Arc<Vec<u8>>,
    after: Arc<Vec<u8>>,
    expected_fingerprint: [u8; 32],
    source_version: Option<SourceVersion>,
    limits: Limits,
    operation: Option<CopyOperation>,
    direction: Direction,
}

impl Patch {
    /// Return whether the patch changes no source bytes.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.before.as_slice() == self.after.as_slice()
    }

    /// Apply only to the byte-exact source artifact and main story.
    pub fn apply(&self, source: &Snapshot) -> Result<Snapshot> {
        if source.artifact_fingerprint != self.expected_fingerprint
            || source.xml.as_slice() != self.before.as_slice()
            || self
                .source_version
                .is_some_and(|expected| source.source_version != expected)
        {
            return Err(Error::StaleSource);
        }
        source.with_xml(checked_clone(self.after.as_slice(), "patch application")?)
    }

    /// Return the exact patch that restores the input bytes.
    ///
    /// This form is authorized for the projected in-memory [`Snapshot`],
    /// which deliberately retains the original artifact identity. After ZIP
    /// publication, use [`Publication::inverse_patch`] to obtain an inverse
    /// authorized for the exact emitted artifact.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            before: Arc::clone(&self.after),
            after: Arc::clone(&self.before),
            expected_fingerprint: self.expected_fingerprint,
            source_version: self.source_version,
            limits: self.limits,
            operation: self.operation,
            direction: match self.direction {
                Direction::Forward => Direction::Inverse,
                Direction::Inverse => Direction::Forward,
            },
        }
    }

    /// Return the deterministic semantic effect.
    #[must_use]
    pub fn effect_report(&self) -> EffectReport {
        let copied_bytes = self.operation.map_or(0, |operation| {
            let range = self
                .source_layout()
                .ok()
                .and_then(|layout| layout.paragraphs.get(operation.source.get()).copied());
            range.map_or(0, |range| range.end.saturating_sub(range.start))
        });
        EffectReport {
            copied_paragraphs: usize::from(self.operation.is_some()),
            copied_bytes,
        }
    }

    /// Encode a canonical bounded durable patch.
    pub fn to_bytes(&self) -> Result<Vec<u8>> {
        let total = HEADER_BYTES
            .checked_add(self.before.len())
            .and_then(|value| value.checked_add(self.after.len()))
            .ok_or(Error::InvalidDurable)?;
        check_limit("durable patch bytes", self.limits.max_durable_bytes, total)?;
        let mut bytes = Vec::new();
        bytes
            .try_reserve_exact(total)
            .map_err(|source| Error::Allocation {
                resource: "durable patch",
                source,
            })?;
        bytes.extend_from_slice(MAGIC);
        bytes.push(VERSION);
        bytes.push(match (self.operation, self.direction) {
            (None, _) => 0,
            (Some(_), Direction::Forward) => 1,
            (Some(_), Direction::Inverse) => 2,
        });
        for value in [
            self.limits.max_xml_bytes,
            self.limits.max_paragraphs,
            self.limits.max_events,
            self.limits.max_depth,
            self.limits.max_output_bytes,
            self.limits.max_durable_bytes,
        ] {
            push_u64(&mut bytes, value)?;
        }
        let (source, before) = self.operation.map_or((0, 0), |operation| {
            (operation.source.get(), operation.before.get())
        });
        push_u32(&mut bytes, source)?;
        push_u32(&mut bytes, before)?;
        bytes.extend_from_slice(&self.expected_fingerprint);
        push_u32(&mut bytes, self.before.len())?;
        push_u32(&mut bytes, self.after.len())?;
        bytes.extend_from_slice(self.before.as_slice());
        bytes.extend_from_slice(self.after.as_slice());
        if bytes.len() != total {
            return Err(Error::InvalidDurable);
        }
        Ok(bytes)
    }

    /// Decode and fully validate a canonical bounded durable patch.
    pub fn from_bytes(bytes: &[u8]) -> Result<Self> {
        if bytes.len() > Limits::default().max_durable_bytes || bytes.len() < HEADER_BYTES {
            return Err(Error::InvalidDurable);
        }
        let mut cursor = 0usize;
        if take(bytes, &mut cursor, MAGIC.len())? != MAGIC {
            return Err(Error::InvalidDurable);
        }
        let version = take(bytes, &mut cursor, 1)?[0];
        if version != VERSION {
            return Err(Error::InvalidDurable);
        }
        let opcode = take(bytes, &mut cursor, 1)?[0];
        let limits = Limits::new(
            take_usize(bytes, &mut cursor)?,
            take_usize(bytes, &mut cursor)?,
            take_usize(bytes, &mut cursor)?,
            take_usize(bytes, &mut cursor)?,
            take_usize(bytes, &mut cursor)?,
            take_usize(bytes, &mut cursor)?,
        )?;
        if bytes.len() > limits.max_durable_bytes {
            return Err(Error::InvalidDurable);
        }
        let source = take_u32(bytes, &mut cursor)? as usize;
        let before_position = take_u32(bytes, &mut cursor)? as usize;
        let mut expected_fingerprint = [0u8; 32];
        expected_fingerprint.copy_from_slice(take(bytes, &mut cursor, 32)?);
        let before_len = take_u32(bytes, &mut cursor)? as usize;
        let after_len = take_u32(bytes, &mut cursor)? as usize;
        let before = Arc::new(checked_clone(
            take(bytes, &mut cursor, before_len)?,
            "durable before bytes",
        )?);
        let after = Arc::new(checked_clone(
            take(bytes, &mut cursor, after_len)?,
            "durable after bytes",
        )?);
        if cursor != bytes.len() {
            return Err(Error::InvalidDurable);
        }
        let operation = match opcode {
            0 if source == 0 && before_position == 0 && before == after => None,
            1 | 2 => Some(CopyOperation {
                source: Position::new(source),
                before: Position::new(before_position),
            }),
            _ => return Err(Error::InvalidDurable),
        };
        let direction = if opcode == 2 {
            Direction::Inverse
        } else {
            Direction::Forward
        };
        validate_durable_shape(&before, &after, operation, direction, limits)?;
        Ok(Self {
            before,
            after,
            expected_fingerprint,
            source_version: None,
            limits,
            operation,
            direction,
        })
    }

    fn source_layout(&self) -> Result<Layout> {
        match self.direction {
            Direction::Forward => scan_document(self.before.as_slice(), self.limits),
            Direction::Inverse => scan_document(self.after.as_slice(), self.limits),
        }
    }
}

/// Deterministic copy effect with no inferred package changes.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct EffectReport {
    /// Number of copied paragraphs (`0` or `1`).
    pub copied_paragraphs: usize,
    /// Exact number of duplicated main-document bytes.
    pub copied_bytes: usize,
}

impl EffectReport {
    /// Return whether no main-document bytes change.
    #[must_use]
    pub const fn is_noop(self) -> bool {
        self.copied_paragraphs == 0
    }
}

/// Exact publication evidence retained for one consuming inverse.
pub struct Publication {
    snapshot: Snapshot,
    original_snapshot: Snapshot,
    original_artifact: SourceArtifact,
    published_fingerprint: [u8; 32],
    inverse_patch: Patch,
}

impl Publication {
    /// Borrow the published semantic snapshot.
    #[must_use]
    pub const fn snapshot(&self) -> &Snapshot {
        &self.snapshot
    }

    /// Borrow the inverse patch authorized for the exact emitted artifact.
    ///
    /// This patch can be serialized, reopened, and published through
    /// [`Package::publish_plain_paragraph_copy_patch_to_stream`]. A package
    /// differing anywhere—not just in `document.xml`—is rejected.
    #[must_use]
    pub const fn inverse_patch(&self) -> &Patch {
        &self.inverse_patch
    }
}

/// One staged exact removal of a direct plain paragraph.
#[derive(Debug, Clone)]
pub struct RemovalEdit {
    base: Snapshot,
    projected: Snapshot,
    position: Option<Position>,
}

impl RemovalEdit {
    /// Borrow the unchanged source snapshot.
    #[must_use]
    pub const fn source(&self) -> &Snapshot {
        &self.base
    }

    /// Borrow the projected snapshot.
    #[must_use]
    pub const fn projected(&self) -> &Snapshot {
        &self.projected
    }

    /// Remove exactly one direct paragraph selected in immutable source order.
    ///
    /// The selected `w:p` fragment, including its original lexical spelling,
    /// is removed without rebuilding any surrounding XML. An empty body is a
    /// valid result when the selected paragraph is the only paragraph.
    pub fn remove_plain_paragraph(&mut self, position: Position) -> Result<&mut Self> {
        if self.position.is_some() {
            return Err(Error::Limit {
                resource: "operations",
                max: 1,
                actual: 2,
            });
        }
        let projected = remove_fragment(&self.base, position)?;
        self.position = Some(position);
        self.projected = projected;
        Ok(self)
    }

    /// Finish the removal as an exact reversible patch.
    #[must_use]
    pub fn commit(self) -> RemovalCommit {
        let patch = RemovalPatch {
            before: Arc::clone(&self.base.xml),
            after: Arc::clone(&self.projected.xml),
            expected_fingerprint: self.base.artifact_fingerprint,
            source_version: Some(self.base.source_version),
            limits: self.base.limits,
            position: self.position,
            direction: Direction::Forward,
        };
        RemovalCommit {
            projected: self.projected,
            patch,
        }
    }
}

/// Completed exact-removal edit.
#[derive(Debug, Clone)]
pub struct RemovalCommit {
    projected: Snapshot,
    patch: RemovalPatch,
}

impl RemovalCommit {
    /// Borrow the projected main-story snapshot.
    #[must_use]
    pub const fn projected(&self) -> &Snapshot {
        &self.projected
    }

    /// Borrow the reversible exact removal patch.
    #[must_use]
    pub const fn patch(&self) -> &RemovalPatch {
        &self.patch
    }

    /// Return the deterministic removal effect.
    #[must_use]
    pub fn effect_report(&self) -> RemovalEffectReport {
        self.patch.effect_report()
    }
}

/// Exact, deterministic, reversible main-document removal patch.
#[derive(Debug, Clone)]
pub struct RemovalPatch {
    before: Arc<Vec<u8>>,
    after: Arc<Vec<u8>>,
    expected_fingerprint: [u8; 32],
    source_version: Option<SourceVersion>,
    limits: Limits,
    position: Option<Position>,
    direction: Direction,
}

impl RemovalPatch {
    /// Return whether the patch changes no source bytes.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.before.as_slice() == self.after.as_slice()
    }

    /// Apply only to the byte-exact source artifact and main story.
    pub fn apply(&self, source: &Snapshot) -> Result<Snapshot> {
        if source.artifact_fingerprint != self.expected_fingerprint
            || source.xml.as_slice() != self.before.as_slice()
            || self
                .source_version
                .is_some_and(|expected| source.source_version != expected)
        {
            return Err(Error::StaleSource);
        }
        source.with_xml(checked_clone(
            self.after.as_slice(),
            "removal patch application",
        )?)
    }

    /// Return the exact patch that restores the input bytes.
    ///
    /// After ZIP publication, use [`RemovalPublication::inverse_patch`] for an
    /// inverse authorized against the exact emitted artifact fingerprint.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            before: Arc::clone(&self.after),
            after: Arc::clone(&self.before),
            expected_fingerprint: self.expected_fingerprint,
            source_version: self.source_version,
            limits: self.limits,
            position: self.position,
            direction: match self.direction {
                Direction::Forward => Direction::Inverse,
                Direction::Inverse => Direction::Forward,
            },
        }
    }

    /// Return the deterministic semantic effect.
    #[must_use]
    pub fn effect_report(&self) -> RemovalEffectReport {
        let removed_bytes = self.position.map_or(0, |position| {
            self.source_layout()
                .ok()
                .and_then(|layout| layout.paragraphs.get(position.get()).copied())
                .map_or(0, |range| range.end.saturating_sub(range.start))
        });
        RemovalEffectReport {
            removed_paragraphs: usize::from(self.position.is_some()),
            removed_bytes,
        }
    }

    /// Encode a canonical bounded durable removal patch.
    pub fn to_bytes(&self) -> Result<Vec<u8>> {
        let total = HEADER_BYTES
            .checked_add(self.before.len())
            .and_then(|value| value.checked_add(self.after.len()))
            .ok_or(Error::InvalidDurable)?;
        check_limit("durable patch bytes", self.limits.max_durable_bytes, total)?;
        let mut bytes = Vec::new();
        bytes
            .try_reserve_exact(total)
            .map_err(|source| Error::Allocation {
                resource: "durable removal patch",
                source,
            })?;
        bytes.extend_from_slice(REMOVAL_MAGIC);
        bytes.push(VERSION);
        bytes.push(match (self.position, self.direction) {
            (None, _) => 0,
            (Some(_), Direction::Forward) => 1,
            (Some(_), Direction::Inverse) => 2,
        });
        for value in [
            self.limits.max_xml_bytes,
            self.limits.max_paragraphs,
            self.limits.max_events,
            self.limits.max_depth,
            self.limits.max_output_bytes,
            self.limits.max_durable_bytes,
        ] {
            push_u64(&mut bytes, value)?;
        }
        push_u32(&mut bytes, self.position.map_or(0, Position::get))?;
        push_u32(&mut bytes, 0)?;
        bytes.extend_from_slice(&self.expected_fingerprint);
        push_u32(&mut bytes, self.before.len())?;
        push_u32(&mut bytes, self.after.len())?;
        bytes.extend_from_slice(self.before.as_slice());
        bytes.extend_from_slice(self.after.as_slice());
        if bytes.len() != total {
            return Err(Error::InvalidDurable);
        }
        Ok(bytes)
    }

    /// Decode and semantically validate a canonical bounded removal patch.
    pub fn from_bytes(bytes: &[u8]) -> Result<Self> {
        if bytes.len() > Limits::default().max_durable_bytes || bytes.len() < HEADER_BYTES {
            return Err(Error::InvalidDurable);
        }
        let mut cursor = 0usize;
        if take(bytes, &mut cursor, REMOVAL_MAGIC.len())? != REMOVAL_MAGIC {
            return Err(Error::InvalidDurable);
        }
        if take(bytes, &mut cursor, 1)?[0] != VERSION {
            return Err(Error::InvalidDurable);
        }
        let opcode = take(bytes, &mut cursor, 1)?[0];
        let limits = Limits::new(
            take_usize(bytes, &mut cursor)?,
            take_usize(bytes, &mut cursor)?,
            take_usize(bytes, &mut cursor)?,
            take_usize(bytes, &mut cursor)?,
            take_usize(bytes, &mut cursor)?,
            take_usize(bytes, &mut cursor)?,
        )?;
        if bytes.len() > limits.max_durable_bytes {
            return Err(Error::InvalidDurable);
        }
        let raw_position = take_u32(bytes, &mut cursor)? as usize;
        let reserved = take_u32(bytes, &mut cursor)?;
        let mut expected_fingerprint = [0u8; 32];
        expected_fingerprint.copy_from_slice(take(bytes, &mut cursor, 32)?);
        let before_len = take_u32(bytes, &mut cursor)? as usize;
        let after_len = take_u32(bytes, &mut cursor)? as usize;
        let before = Arc::new(checked_clone(
            take(bytes, &mut cursor, before_len)?,
            "durable removal before bytes",
        )?);
        let after = Arc::new(checked_clone(
            take(bytes, &mut cursor, after_len)?,
            "durable removal after bytes",
        )?);
        if cursor != bytes.len() || reserved != 0 {
            return Err(Error::InvalidDurable);
        }
        let position = match opcode {
            0 if raw_position == 0 && before == after => None,
            1 | 2 => Some(Position::new(raw_position)),
            _ => return Err(Error::InvalidDurable),
        };
        let direction = if opcode == 2 {
            Direction::Inverse
        } else {
            Direction::Forward
        };
        validate_durable_removal_shape(&before, &after, position, direction, limits)?;
        Ok(Self {
            before,
            after,
            expected_fingerprint,
            source_version: None,
            limits,
            position,
            direction,
        })
    }

    fn source_layout(&self) -> Result<Layout> {
        match self.direction {
            Direction::Forward => scan_document(self.before.as_slice(), self.limits),
            Direction::Inverse => scan_document(self.after.as_slice(), self.limits),
        }
    }
}

/// Deterministic removal effect with no inferred package changes.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct RemovalEffectReport {
    /// Number of removed paragraphs (`0` or `1`).
    pub removed_paragraphs: usize,
    /// Exact number of removed main-document bytes.
    pub removed_bytes: usize,
}

impl RemovalEffectReport {
    /// Return whether no main-document bytes change.
    #[must_use]
    pub const fn is_noop(self) -> bool {
        self.removed_paragraphs == 0
    }
}

/// Exact removal publication evidence retained for one consuming inverse.
pub struct RemovalPublication {
    snapshot: Snapshot,
    original_snapshot: Snapshot,
    original_artifact: SourceArtifact,
    published_fingerprint: [u8; 32],
    inverse_patch: RemovalPatch,
}

impl RemovalPublication {
    /// Borrow the published semantic snapshot.
    #[must_use]
    pub const fn snapshot(&self) -> &Snapshot {
        &self.snapshot
    }

    /// Borrow the inverse authorized for the exact emitted artifact.
    #[must_use]
    pub const fn inverse_patch(&self) -> &RemovalPatch {
        &self.inverse_patch
    }
}

impl Package {
    /// Capture the exact bounded closure for a plain main-story paragraph copy.
    pub fn plain_paragraph_copy_snapshot(&self) -> Result<Snapshot> {
        self.plain_paragraph_copy_snapshot_with_limits(Limits::default())
    }

    /// Capture the exact closure under caller-selected finite limits.
    pub fn plain_paragraph_copy_snapshot_with_limits(&self, limits: Limits) -> Result<Snapshot> {
        let source_version = self.package.source_version().map_err(crate::Error::from)?;
        self.validate_plain_paragraph_copy_topology()?;
        let main = self
            .package
            .main_document_part()
            .map_err(crate::Error::from)?;
        validate_document_main_content_type(main.content_type())?;
        let xml = main.data().map_err(crate::Error::from)?.into_arc();
        let layout = scan_document(xml.as_slice(), limits)?;
        let artifact_fingerprint = fingerprint_artifact(&self.package.source_artifact())?;
        let observed = self.package.source_version().map_err(crate::Error::from)?;
        if observed != source_version {
            return Err(crate::Error::Opc(litchi_opc::OpcError::SourceChanged {
                expected: source_version,
                actual: observed,
            })
            .into());
        }
        Ok(Snapshot {
            xml,
            layout,
            source_version,
            artifact_fingerprint,
            limits,
        })
    }

    /// Start one isolated source-backed exact paragraph-copy edit.
    pub fn edit_plain_paragraph_copy(&self) -> Result<Edit> {
        Ok(self.plain_paragraph_copy_snapshot()?.edit())
    }

    /// Capture the exact bounded closure for plain-paragraph removal.
    pub fn plain_paragraph_removal_snapshot(&self) -> Result<Snapshot> {
        self.plain_paragraph_removal_snapshot_with_limits(Limits::default())
    }

    /// Capture the exact removal closure under caller-selected finite limits.
    pub fn plain_paragraph_removal_snapshot_with_limits(&self, limits: Limits) -> Result<Snapshot> {
        self.plain_paragraph_copy_snapshot_with_limits(limits)
    }

    /// Start one isolated source-backed exact paragraph-removal edit.
    pub fn edit_plain_paragraph_removal(&self) -> Result<RemovalEdit> {
        Ok(self.plain_paragraph_removal_snapshot()?.removal_edit())
    }

    /// Publish an exact-source-checked copy while raw-copying every other ZIP member.
    pub fn publish_plain_paragraph_copy_to_stream<W: Write>(
        self,
        writer: W,
        commit: &Commit,
    ) -> Result<Publication> {
        self.publish_plain_paragraph_copy_patch_to_stream(writer, commit.patch())
    }

    /// Publish a live or rehydrated durable patch against its exact source.
    ///
    /// A live patch also requires its captured process-local source version.
    /// A rehydrated patch instead authorizes an identical reopened package by
    /// the complete source-artifact SHA-256 retained in its canonical wire.
    pub fn publish_plain_paragraph_copy_patch_to_stream<W: Write>(
        self,
        writer: W,
        patch: &Patch,
    ) -> Result<Publication> {
        let current = self.plain_paragraph_copy_snapshot_with_limits(patch.limits)?;
        let target = patch.apply(&current)?;
        let original_artifact = self.package.source_artifact();
        let main = self
            .package
            .main_document_part()
            .map_err(crate::Error::from)?
            .partname()
            .clone();
        let mut output = FingerprintingWriter::new(writer);
        if patch.is_noop() {
            original_artifact
                .write_to_stream(&mut output)
                .map_err(crate::Error::from)?;
        } else {
            let replacement = checked_clone(target.xml_bytes(), "publication replacement")?;
            self.package
                .write_part_overlay_to_stream(&mut output, &main, replacement)
                .map_err(crate::Error::from)?;
        }
        let published_fingerprint = output.finish();
        let mut inverse_patch = patch.inverse();
        inverse_patch.expected_fingerprint = published_fingerprint;
        inverse_patch.source_version = None;
        Ok(Publication {
            snapshot: target,
            original_snapshot: current,
            original_artifact,
            published_fingerprint,
            inverse_patch,
        })
    }

    /// Restore the exact original package from a byte-exact published artifact.
    pub fn publish_plain_paragraph_copy_inverse_to_stream<W: Write>(
        self,
        writer: W,
        publication: &Publication,
    ) -> Result<Snapshot> {
        let current = fingerprint_artifact(&self.package.source_artifact())?;
        if current != publication.published_fingerprint {
            return Err(Error::StaleSource);
        }
        publication
            .original_artifact
            .write_to_stream(writer)
            .map_err(crate::Error::from)?;
        Ok(publication.original_snapshot.clone())
    }

    /// Publish one exact-source-checked removal while raw-copying every other
    /// physical ZIP member.
    pub fn publish_plain_paragraph_removal_to_stream<W: Write>(
        self,
        writer: W,
        commit: &RemovalCommit,
    ) -> Result<RemovalPublication> {
        self.publish_plain_paragraph_removal_patch_to_stream(writer, commit.patch())
    }

    /// Publish a live or rehydrated durable removal patch against its exact
    /// complete source artifact.
    pub fn publish_plain_paragraph_removal_patch_to_stream<W: Write>(
        self,
        writer: W,
        patch: &RemovalPatch,
    ) -> Result<RemovalPublication> {
        let current = self.plain_paragraph_removal_snapshot_with_limits(patch.limits)?;
        let target = patch.apply(&current)?;
        let original_artifact = self.package.source_artifact();
        let main = self
            .package
            .main_document_part()
            .map_err(crate::Error::from)?
            .partname()
            .clone();
        let mut output = FingerprintingWriter::new(writer);
        if patch.is_noop() {
            original_artifact
                .write_to_stream(&mut output)
                .map_err(crate::Error::from)?;
        } else {
            let replacement = checked_clone(target.xml_bytes(), "removal publication replacement")?;
            self.package
                .write_part_overlay_to_stream(&mut output, &main, replacement)
                .map_err(crate::Error::from)?;
        }
        let published_fingerprint = output.finish();
        let mut inverse_patch = patch.inverse();
        inverse_patch.expected_fingerprint = published_fingerprint;
        inverse_patch.source_version = None;
        Ok(RemovalPublication {
            snapshot: target,
            original_snapshot: current,
            original_artifact,
            published_fingerprint,
            inverse_patch,
        })
    }

    /// Restore the exact original package from a byte-exact removal output.
    pub fn publish_plain_paragraph_removal_inverse_to_stream<W: Write>(
        self,
        writer: W,
        publication: &RemovalPublication,
    ) -> Result<Snapshot> {
        let current = fingerprint_artifact(&self.package.source_artifact())?;
        if current != publication.published_fingerprint {
            return Err(Error::StaleSource);
        }
        publication
            .original_artifact
            .write_to_stream(writer)
            .map_err(crate::Error::from)?;
        Ok(publication.original_snapshot.clone())
    }

    fn validate_plain_paragraph_copy_topology(&self) -> Result<()> {
        let main = self
            .package
            .main_document_part()
            .map_err(crate::Error::from)?;
        let package_strict = self.package.rels().iter().find_map(|relationship| {
            matches!(
                relationship.reltype(),
                rt::OFFICE_DOCUMENT | rt::STRICT_OFFICE_DOCUMENT
            )
            .then_some(relationship.reltype() == rt::STRICT_OFFICE_DOCUMENT)
        });
        if main.partname().as_str() != "/word/document.xml" {
            return Err(Error::Refused(Refusal::MainDocumentShape));
        }
        let main_data = main.data().map_err(crate::Error::from)?;
        let root_strict = inspect_document_root_dialect(main_data.as_bytes())?;
        if package_strict != Some(root_strict) {
            return Err(Error::Refused(Refusal::MainDocumentShape));
        }
        if main.content_type() != ct::WML_DOCUMENT_MAIN {
            return Err(Error::Refused(
                if matches!(
                    main.content_type(),
                    ct::WML_DOCUMENT_MACRO_MAIN | ct::WML_TEMPLATE_MACRO_MAIN
                ) {
                    Refusal::MacroEnabled
                } else {
                    Refusal::MainDocumentShape
                },
            ));
        }

        for relationship in self.package.rels().iter() {
            validate_relationship(relationship)?;
        }
        let mut settings = None;
        for part in self.package.iter_parts() {
            validate_part(part.partname().as_str(), part.content_type())?;
            for relationship in part.rels().iter() {
                validate_relationship(relationship)?;
                if part.partname() == main.partname()
                    && matches!(relationship.reltype(), rt::SETTINGS | STRICT_SETTINGS)
                {
                    if settings
                        .replace((
                            relationship.target_partname().map_err(crate::Error::from)?,
                            relationship.reltype() == STRICT_SETTINGS,
                        ))
                        .is_some()
                    {
                        return Err(Error::Refused(Refusal::UnsupportedDependency));
                    }
                }
            }
        }
        if let Some((settings_name, settings_strict)) = settings {
            if package_strict != Some(settings_strict) {
                return Err(Error::Refused(Refusal::UnsupportedDependency));
            }
            let settings_part = self
                .package
                .part(&settings_name)
                .map_err(crate::Error::from)?;
            if settings_part.content_type() != ct::WML_SETTINGS {
                return Err(Error::Refused(Refusal::UnsupportedDependency));
            }
            let settings_data = settings_part.data().map_err(crate::Error::from)?;
            if inspect_word_root_dialect(settings_data.as_bytes(), b"settings")? != settings_strict
            {
                return Err(Error::Refused(Refusal::UnsupportedDependency));
            }
            let mut staged = BlobPart::new(
                settings_name,
                ct::WML_SETTINGS.to_owned(),
                checked_clone(settings_data.as_bytes(), "settings inspection")?,
            );
            for relationship in settings_part.rels().iter() {
                staged
                    .rels_mut()
                    .try_add_relationship(
                        relationship.reltype().to_owned(),
                        relationship.target_ref().to_owned(),
                        relationship.r_id().to_owned(),
                        relationship.target_mode(),
                    )
                    .map_err(crate::Error::from)?;
            }
            let parsed = DocumentSettings::extract_from_part(&staged)?;
            if parsed.is_protected() || parsed.is_write_protected() || parsed.track_revisions() {
                return Err(Error::Refused(Refusal::Protection));
            }
        }
        Ok(())
    }
}

const STRICT_SETTINGS: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships/settings";
const STRICT_FOOTNOTES: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships/footnotes";
const STRICT_ENDNOTES: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships/endnotes";
const STRICT_CUSTOM_XML: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships/customXml";
const STRICT_ALTERNATIVE_FORMAT_IMPORT: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships/afChunk";

fn validate_relationship(relationship: &litchi_opc::Relationship) -> Result<()> {
    if relationship.is_external() {
        return Err(Error::Refused(Refusal::ExternalRelationship));
    }
    if is_signature_relationship(relationship.reltype()) {
        return Err(Error::Refused(Refusal::SignatureInfrastructure));
    }
    if matches!(
        relationship.reltype(),
        rt::VBA_PROJECT
            | rt::VBA_PROJECT_SIGNATURE
            | rt::VBA_PROJECT_SIGNATURE_AGILE
            | rt::WORD_VBA_DATA
    ) {
        return Err(Error::Refused(Refusal::MacroEnabled));
    }
    if matches!(
        relationship.reltype(),
        rt::COMMENTS
            | rt::STRICT_COMMENTS
            | rt::FOOTNOTES
            | rt::ENDNOTES
            | rt::CUSTOM_XML
            | rt::ALTERNATIVE_FORMAT_IMPORT
            | rt::MS_ALTERNATIVE_FORMAT_IMPORT
            | STRICT_FOOTNOTES
            | STRICT_ENDNOTES
            | STRICT_CUSTOM_XML
            | STRICT_ALTERNATIVE_FORMAT_IMPORT
    ) {
        return Err(Error::Refused(Refusal::UnsupportedDependency));
    }
    Ok(())
}

fn validate_part(path: &str, content_type: &str) -> Result<()> {
    if path
        .as_bytes()
        .get(..SIGNATURE_DIRECTORY.len())
        .is_some_and(|prefix| prefix.eq_ignore_ascii_case(SIGNATURE_DIRECTORY.as_bytes()))
        || matches!(
            content_type,
            ct::OPC_DIGITAL_SIGNATURE_ORIGIN
                | ct::OPC_DIGITAL_SIGNATURE_XMLSIGNATURE
                | ct::OPC_DIGITAL_SIGNATURE_CERTIFICATE
        )
    {
        return Err(Error::Refused(Refusal::SignatureInfrastructure));
    }
    if matches!(
        content_type,
        ct::WML_DOCUMENT_MACRO_MAIN
            | ct::WML_TEMPLATE_MACRO_MAIN
            | ct::OFC_VBA_PROJECT
            | ct::OFC_VBA_PROJECT_SIGNATURE
            | ct::OFC_VBA_PROJECT_SIGNATURE_AGILE
            | ct::WML_VBA_DATA
    ) {
        return Err(Error::Refused(Refusal::MacroEnabled));
    }
    if matches!(
        content_type,
        ct::WML_COMMENTS | ct::WML_FOOTNOTES | ct::WML_ENDNOTES | ct::OFC_CUSTOM_XML_PROPERTIES
    ) || path.starts_with("/customXml/")
    {
        return Err(Error::Refused(Refusal::UnsupportedDependency));
    }
    Ok(())
}

fn is_signature_relationship(value: &str) -> bool {
    matches!(
        value,
        rt::DIGITAL_SIGNATURE_ORIGIN
            | "http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/signature"
            | "http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/certificate"
    )
}

fn copy_fragment(source: &Snapshot, from: Position, before: Position) -> Result<Snapshot> {
    let len = source.paragraph_count();
    let source_range =
        source
            .layout
            .paragraphs
            .get(from.get())
            .copied()
            .ok_or(Error::OutOfBounds {
                kind: "source",
                position: from.get(),
                len,
            })?;
    if before.get() > len {
        return Err(Error::OutOfBounds {
            kind: "insertion",
            position: before.get(),
            len,
        });
    }
    let insertion = source
        .layout
        .paragraphs
        .get(before.get())
        .map_or(source.layout.body_end, |range| range.start);
    let fragment = source
        .xml
        .get(source_range.start..source_range.end)
        .ok_or(Error::InvalidDurable)?;
    let output_len = source
        .xml
        .len()
        .checked_add(fragment.len())
        .ok_or(Error::InvalidDurable)?;
    check_limit("output bytes", source.limits.max_output_bytes, output_len)?;
    let mut output = Vec::new();
    output
        .try_reserve_exact(output_len)
        .map_err(|source| Error::Allocation {
            resource: "main-document output",
            source,
        })?;
    output.extend_from_slice(&source.xml[..insertion]);
    output.extend_from_slice(fragment);
    output.extend_from_slice(&source.xml[insertion..]);
    if output.len() != output_len {
        return Err(Error::InvalidDurable);
    }
    let candidate = source.with_xml(output)?;
    validate_copy_readback(source, &candidate, from, before)?;
    Ok(candidate)
}

fn remove_fragment(source: &Snapshot, position: Position) -> Result<Snapshot> {
    let len = source.paragraph_count();
    let removed = source
        .layout
        .paragraphs
        .get(position.get())
        .copied()
        .ok_or(Error::OutOfBounds {
            kind: "removal",
            position: position.get(),
            len,
        })?;
    let removed_len = removed
        .end
        .checked_sub(removed.start)
        .ok_or(Error::InvalidDurable)?;
    let output_len = source
        .xml
        .len()
        .checked_sub(removed_len)
        .ok_or(Error::InvalidDurable)?;
    check_limit("output bytes", source.limits.max_output_bytes, output_len)?;
    let mut output = Vec::new();
    output
        .try_reserve_exact(output_len)
        .map_err(|source| Error::Allocation {
            resource: "main-document removal output",
            source,
        })?;
    output.extend_from_slice(&source.xml[..removed.start]);
    output.extend_from_slice(&source.xml[removed.end..]);
    if output.len() != output_len {
        return Err(Error::InvalidDurable);
    }
    let candidate = source.with_xml(output)?;
    validate_removal_readback(source, &candidate, position)?;
    Ok(candidate)
}

fn validate_removal_readback(
    source: &Snapshot,
    target: &Snapshot,
    position: Position,
) -> Result<()> {
    if source.paragraph_count() != target.paragraph_count().saturating_add(1) {
        return Err(Error::InvalidDurable);
    }
    for target_index in 0..target.paragraph_count() {
        let source_index = if target_index < position.get() {
            target_index
        } else {
            target_index.checked_add(1).ok_or(Error::InvalidDurable)?
        };
        if paragraph_bytes(target, target_index)? != paragraph_bytes(source, source_index)? {
            return Err(Error::InvalidDurable);
        }
    }
    Ok(())
}

fn validate_copy_readback(
    source: &Snapshot,
    target: &Snapshot,
    from: Position,
    before: Position,
) -> Result<()> {
    if target.paragraph_count() != source.paragraph_count().saturating_add(1) {
        return Err(Error::InvalidDurable);
    }
    let copied = paragraph_bytes(source, from.get())?;
    for target_index in 0..target.paragraph_count() {
        let expected = if target_index == before.get() {
            copied
        } else {
            let source_index = if target_index < before.get() {
                target_index
            } else {
                target_index - 1
            };
            paragraph_bytes(source, source_index)?
        };
        if paragraph_bytes(target, target_index)? != expected {
            return Err(Error::InvalidDurable);
        }
    }
    Ok(())
}

fn paragraph_bytes(snapshot: &Snapshot, index: usize) -> Result<&[u8]> {
    let range = snapshot
        .layout
        .paragraphs
        .get(index)
        .ok_or(Error::InvalidDurable)?;
    snapshot
        .xml
        .get(range.start..range.end)
        .ok_or(Error::InvalidDurable)
}

fn validate_durable_shape(
    before: &[u8],
    after: &[u8],
    operation: Option<CopyOperation>,
    direction: Direction,
    limits: Limits,
) -> Result<()> {
    let before_layout = scan_document(before, limits)?;
    let after_layout = scan_document(after, limits)?;
    let before_snapshot = Snapshot {
        xml: Arc::new(checked_clone(before, "durable validation source")?),
        layout: before_layout,
        source_version: SourceVersion::new(0, 0),
        artifact_fingerprint: [0; 32],
        limits,
    };
    let after_snapshot = Snapshot {
        xml: Arc::new(checked_clone(after, "durable validation target")?),
        layout: after_layout,
        source_version: SourceVersion::new(0, 0),
        artifact_fingerprint: [0; 32],
        limits,
    };
    let Some(operation) = operation else {
        return if before == after {
            Ok(())
        } else {
            Err(Error::InvalidDurable)
        };
    };
    let (source, expected) = match direction {
        Direction::Forward => (&before_snapshot, &after_snapshot),
        Direction::Inverse => (&after_snapshot, &before_snapshot),
    };
    let rebuilt = copy_fragment(source, operation.source, operation.before)?;
    if rebuilt.xml.as_slice() != expected.xml.as_slice() {
        return Err(Error::InvalidDurable);
    }
    Ok(())
}

fn validate_durable_removal_shape(
    before: &[u8],
    after: &[u8],
    position: Option<Position>,
    direction: Direction,
    limits: Limits,
) -> Result<()> {
    let before_layout = scan_document(before, limits)?;
    let after_layout = scan_document(after, limits)?;
    let before_snapshot = Snapshot {
        xml: Arc::new(checked_clone(before, "durable removal validation source")?),
        layout: before_layout,
        source_version: SourceVersion::new(0, 0),
        artifact_fingerprint: [0; 32],
        limits,
    };
    let after_snapshot = Snapshot {
        xml: Arc::new(checked_clone(after, "durable removal validation target")?),
        layout: after_layout,
        source_version: SourceVersion::new(0, 0),
        artifact_fingerprint: [0; 32],
        limits,
    };
    let Some(position) = position else {
        return if before == after {
            Ok(())
        } else {
            Err(Error::InvalidDurable)
        };
    };
    let (source, expected) = match direction {
        Direction::Forward => (&before_snapshot, &after_snapshot),
        Direction::Inverse => (&after_snapshot, &before_snapshot),
    };
    let rebuilt = remove_fragment(source, position)?;
    if rebuilt.xml.as_slice() != expected.xml.as_slice() {
        return Err(Error::InvalidDurable);
    }
    Ok(())
}

#[derive(Clone, Copy, PartialEq, Eq)]
enum Scope {
    Document,
    Body,
    Paragraph,
    Run,
    Text,
}

fn scan_document(xml: &[u8], limits: Limits) -> Result<Layout> {
    check_limit("XML bytes", limits.max_xml_bytes, xml.len())?;
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut stack = Vec::new();
    stack
        .try_reserve_exact(limits.max_depth.min(64))
        .map_err(|source| Error::Allocation {
            resource: "XML scope stack",
            source,
        })?;
    let mut paragraphs = Vec::new();
    paragraphs
        .try_reserve_exact(limits.max_paragraphs.min(4_096))
        .map_err(|source| Error::Allocation {
            resource: "paragraph ranges",
            source,
        })?;
    let mut paragraph_start = None;
    let mut body_end = None;
    let mut events = 0usize;
    let mut saw_document = false;
    let mut saw_body = false;
    let mut finished_document = false;
    let mut word_namespace: Option<Vec<u8>> = None;

    loop {
        let event_start =
            usize::try_from(reader.buffer_position()).map_err(|_error| Error::InvalidDurable)?;
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|_error| Error::Refused(Refusal::ComplexDocument))?;
        events = events.checked_add(1).ok_or(Error::InvalidDurable)?;
        check_limit("XML events", limits.max_events, events)?;
        match event {
            Event::Start(start) => {
                let local = checked_clone(start.local_name().as_ref(), "XML local name")?;
                let resolved = resolved_word_namespace(namespace)?;
                if let Some(expected) = word_namespace.as_deref() {
                    if resolved != expected {
                        return Err(Error::Refused(Refusal::ComplexDocument));
                    }
                } else if local.as_slice() == b"document" && stack.is_empty() {
                    word_namespace = Some(checked_clone(resolved, "Word namespace")?);
                }
                let scope = next_scope(&stack, local.as_slice(), false)?;
                validate_attributes(&start, scope, word_namespace.as_deref().unwrap_or(resolved))?;
                if scope == Scope::Document {
                    if saw_document || finished_document {
                        return Err(Error::Refused(Refusal::ComplexDocument));
                    }
                    saw_document = true;
                } else if scope == Scope::Body {
                    if saw_body {
                        return Err(Error::Refused(Refusal::ComplexDocument));
                    }
                    saw_body = true;
                } else if scope == Scope::Paragraph {
                    paragraph_start = Some(event_start);
                }
                reserve_one(&mut stack, "XML scope stack")?;
                stack.push(scope);
                check_limit("XML depth", limits.max_depth, stack.len())?;
            },
            Event::Empty(empty) => {
                let local = checked_clone(empty.local_name().as_ref(), "XML local name")?;
                let resolved = resolved_word_namespace(namespace)?;
                let Some(expected) = word_namespace.as_deref() else {
                    return Err(Error::Refused(Refusal::ComplexDocument));
                };
                if resolved != expected {
                    return Err(Error::Refused(Refusal::ComplexDocument));
                }
                let scope = next_scope(&stack, local.as_slice(), true)?;
                validate_attributes(&empty, scope, expected)?;
                if scope == Scope::Paragraph {
                    check_limit(
                        "paragraphs",
                        limits.max_paragraphs,
                        paragraphs.len().saturating_add(1),
                    )?;
                    reserve_one(&mut paragraphs, "paragraph ranges")?;
                    paragraphs.push(Range {
                        start: event_start,
                        end: usize::try_from(reader.buffer_position())
                            .map_err(|_error| Error::InvalidDurable)?,
                    });
                }
            },
            Event::End(end) => {
                let local = checked_clone(end.local_name().as_ref(), "XML local name")?;
                let resolved = resolved_word_namespace(namespace)?;
                let Some(expected) = word_namespace.as_deref() else {
                    return Err(Error::Refused(Refusal::ComplexDocument));
                };
                if resolved != expected {
                    return Err(Error::Refused(Refusal::ComplexDocument));
                }
                let scope = stack
                    .pop()
                    .ok_or(Error::Refused(Refusal::ComplexDocument))?;
                if scope_local(scope) != local.as_slice() {
                    return Err(Error::Refused(Refusal::ComplexDocument));
                }
                if scope == Scope::Paragraph {
                    check_limit(
                        "paragraphs",
                        limits.max_paragraphs,
                        paragraphs.len().saturating_add(1),
                    )?;
                    reserve_one(&mut paragraphs, "paragraph ranges")?;
                    paragraphs.push(Range {
                        start: paragraph_start.take().ok_or(Error::InvalidDurable)?,
                        end: usize::try_from(reader.buffer_position())
                            .map_err(|_error| Error::InvalidDurable)?,
                    });
                } else if scope == Scope::Body {
                    body_end = Some(event_start);
                } else if scope == Scope::Document {
                    finished_document = true;
                }
            },
            Event::Text(text) => {
                if stack.last() != Some(&Scope::Text)
                    && text.as_ref().iter().any(|byte| !byte.is_ascii_whitespace())
                {
                    return Err(Error::Refused(Refusal::ComplexDocument));
                }
            },
            Event::Decl(_) if stack.is_empty() && !saw_document => {},
            Event::Eof => break,
            Event::GeneralRef(reference)
                if stack.last() == Some(&Scope::Text)
                    && matches!(
                        reference.as_ref(),
                        b"amp" | b"lt" | b"gt" | b"apos" | b"quot"
                    ) => {},
            Event::Comment(_)
            | Event::CData(_)
            | Event::PI(_)
            | Event::DocType(_)
            | Event::GeneralRef(_) => {
                return Err(Error::Refused(Refusal::ComplexDocument));
            },
            _ => return Err(Error::Refused(Refusal::ComplexDocument)),
        }
        buffer.clear();
    }
    if !saw_document || !saw_body || !finished_document || !stack.is_empty() {
        return Err(Error::Refused(Refusal::ComplexDocument));
    }
    Ok(Layout {
        paragraphs: paragraphs.into(),
        body_end: body_end.ok_or(Error::Refused(Refusal::ComplexDocument))?,
    })
}

fn inspect_document_root_dialect(xml: &[u8]) -> Result<bool> {
    inspect_word_root_dialect(xml, b"document")
}

fn inspect_word_root_dialect(xml: &[u8], root: &[u8]) -> Result<bool> {
    let mut reader = NsReader::from_reader(xml);
    let mut buffer = Vec::new();
    loop {
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|_error| Error::Refused(Refusal::ComplexDocument))?;
        match event {
            Event::Start(start) if start.local_name().as_ref() == root => {
                return match namespace {
                    ResolveResult::Bound(Namespace(value)) if value == TRANSITIONAL_WORD => {
                        Ok(false)
                    },
                    ResolveResult::Bound(Namespace(value)) if value == STRICT_WORD => Ok(true),
                    ResolveResult::Bound(_)
                    | ResolveResult::Unknown(_)
                    | ResolveResult::Unbound => Err(Error::Refused(Refusal::ComplexDocument)),
                };
            },
            Event::Decl(_) => {},
            Event::Text(text) if text.as_ref().iter().all(u8::is_ascii_whitespace) => {},
            Event::Eof => return Err(Error::Refused(Refusal::ComplexDocument)),
            _ => return Err(Error::Refused(Refusal::ComplexDocument)),
        }
        buffer.clear();
    }
}

fn resolved_word_namespace(namespace: ResolveResult<'_>) -> Result<&[u8]> {
    match namespace {
        ResolveResult::Bound(Namespace(value))
            if value == TRANSITIONAL_WORD || value == STRICT_WORD =>
        {
            Ok(value)
        },
        ResolveResult::Bound(_) | ResolveResult::Unknown(_) | ResolveResult::Unbound => {
            Err(Error::Refused(Refusal::ComplexDocument))
        },
    }
}

fn next_scope(stack: &[Scope], local: &[u8], empty: bool) -> Result<Scope> {
    let scope = match (stack, local) {
        ([], b"document") if !empty => Scope::Document,
        ([Scope::Document], b"body") if !empty => Scope::Body,
        ([Scope::Document, Scope::Body], b"p") => Scope::Paragraph,
        ([Scope::Document, Scope::Body, Scope::Paragraph], b"r") => Scope::Run,
        ([Scope::Document, Scope::Body, Scope::Paragraph, Scope::Run], b"t") => Scope::Text,
        _ => {
            return Err(Error::Refused(
                if stack.last() == Some(&Scope::Paragraph)
                    || stack.last() == Some(&Scope::Run)
                    || stack.last() == Some(&Scope::Text)
                {
                    Refusal::ComplexParagraph
                } else {
                    Refusal::ComplexDocument
                },
            ));
        },
    };
    if empty && matches!(scope, Scope::Document | Scope::Body | Scope::Run) {
        return Err(Error::Refused(Refusal::ComplexDocument));
    }
    Ok(scope)
}

const fn scope_local(scope: Scope) -> &'static [u8] {
    match scope {
        Scope::Document => b"document",
        Scope::Body => b"body",
        Scope::Paragraph => b"p",
        Scope::Run => b"r",
        Scope::Text => b"t",
    }
}

fn validate_attributes(start: &BytesStart<'_>, scope: Scope, word_namespace: &[u8]) -> Result<()> {
    for attribute in start.attributes() {
        let attribute = attribute.map_err(|_error| Error::Refused(Refusal::ComplexDocument))?;
        let key = attribute.key.as_ref();
        let value = attribute.value.as_ref();
        let namespace_declaration = key == b"xmlns" || key.starts_with(b"xmlns:");
        if namespace_declaration {
            // Namespace declarations alone are inert and ordinary producers
            // place many of them on `w:document`. Any use by an element or a
            // non-namespace attribute is still rejected by the namespace-
            // aware scanner below this root.
            if scope != Scope::Document && value != word_namespace && value != XML_NAMESPACE {
                return Err(Error::Refused(Refusal::ComplexDocument));
            }
            continue;
        }
        if scope == Scope::Text && key == b"xml:space" && matches!(value, b"preserve" | b"default")
        {
            continue;
        }
        return Err(Error::Refused(
            if matches!(scope, Scope::Paragraph | Scope::Run | Scope::Text) {
                Refusal::ComplexParagraph
            } else {
                Refusal::ComplexDocument
            },
        ));
    }
    Ok(())
}

fn check_limit(resource: &'static str, max: usize, actual: usize) -> Result<()> {
    if actual > max {
        return Err(Error::Limit {
            resource,
            max,
            actual,
        });
    }
    Ok(())
}

fn checked_clone(source_bytes: &[u8], resource: &'static str) -> Result<Vec<u8>> {
    let mut target = Vec::new();
    target
        .try_reserve_exact(source_bytes.len())
        .map_err(|source| Error::Allocation { resource, source })?;
    target.extend_from_slice(source_bytes);
    Ok(target)
}

fn reserve_one<T>(target: &mut Vec<T>, resource: &'static str) -> Result<()> {
    if target.len() == target.capacity() {
        target
            .try_reserve_exact(1)
            .map_err(|source| Error::Allocation { resource, source })?;
    }
    Ok(())
}

struct FingerprintingWriter<W> {
    inner: W,
    hasher: Sha256,
}

impl<W> FingerprintingWriter<W> {
    fn new(inner: W) -> Self {
        Self {
            inner,
            hasher: Sha256::new(),
        }
    }

    fn finish(self) -> [u8; 32] {
        self.hasher.finalize().into()
    }
}

impl<W: Write> Write for FingerprintingWriter<W> {
    fn write(&mut self, bytes: &[u8]) -> std::io::Result<usize> {
        let written = self.inner.write(bytes)?;
        if written > bytes.len() {
            return Err(std::io::Error::new(
                std::io::ErrorKind::InvalidData,
                "plain paragraph copy sink over-reported progress",
            ));
        }
        self.hasher.update(&bytes[..written]);
        Ok(written)
    }

    fn flush(&mut self) -> std::io::Result<()> {
        self.inner.flush()
    }
}

struct HashWriter(Sha256);

impl Write for HashWriter {
    fn write(&mut self, bytes: &[u8]) -> std::io::Result<usize> {
        self.0.update(bytes);
        Ok(bytes.len())
    }

    fn flush(&mut self) -> std::io::Result<()> {
        Ok(())
    }
}

fn fingerprint_artifact(artifact: &SourceArtifact) -> Result<[u8; 32]> {
    let mut writer = HashWriter(Sha256::new());
    artifact
        .write_to_stream(&mut writer)
        .map_err(crate::Error::from)?;
    Ok(writer.0.finalize().into())
}

fn push_u32(bytes: &mut Vec<u8>, value: usize) -> Result<()> {
    let value = u32::try_from(value).map_err(|_error| Error::InvalidDurable)?;
    bytes.extend_from_slice(&value.to_le_bytes());
    Ok(())
}

fn push_u64(bytes: &mut Vec<u8>, value: usize) -> Result<()> {
    let value = u64::try_from(value).map_err(|_error| Error::InvalidDurable)?;
    bytes.extend_from_slice(&value.to_le_bytes());
    Ok(())
}

fn take<'a>(bytes: &'a [u8], cursor: &mut usize, len: usize) -> Result<&'a [u8]> {
    let end = cursor.checked_add(len).ok_or(Error::InvalidDurable)?;
    let value = bytes.get(*cursor..end).ok_or(Error::InvalidDurable)?;
    *cursor = end;
    Ok(value)
}

fn take_u32(bytes: &[u8], cursor: &mut usize) -> Result<u32> {
    let raw: [u8; 4] = take(bytes, cursor, 4)?
        .try_into()
        .map_err(|_error| Error::InvalidDurable)?;
    Ok(u32::from_le_bytes(raw))
}

fn take_usize(bytes: &[u8], cursor: &mut usize) -> Result<usize> {
    let raw: [u8; 8] = take(bytes, cursor, 8)?
        .try_into()
        .map_err(|_error| Error::InvalidDurable)?;
    usize::try_from(u64::from_le_bytes(raw)).map_err(|_error| Error::InvalidDurable)
}
