//! Source-backed text snapshots for one selected Word story.
//!
//! This module deliberately keeps the selection closure to one main, header,
//! footer, footnote, endnote, or comment story. Relationship metadata is
//! resolved without reading payloads; only the owning XML member is
//! materialized. Edits reuse the main document transaction engine through a
//! private `w:document`/`w:body` adapter. For collection parts, only the exact
//! selected entry body range is spliced back, preserving the root and all
//! sibling entries byte-for-byte.

use std::collections::TryReserveError;
use std::io::{self, Write};
use std::sync::Arc;

use litchi_core::{
    ExecutionContext, ExecutionError, Position, SourceVersion, TextObjectKind, TextOutputError,
    TextOutputOptions, TextOutputReport,
};
use litchi_ooxml_common::xml::decode_xml_reference;
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{
    PackURI, PartData, PartView, SourceArtifact, SourceArtifactFingerprint, SourceBackedPackage,
    SourceLineage,
};
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::QName;
use quick_xml::reader::Reader;
use sha2::{Digest as _, Sha256};
use thiserror::Error as ThisError;

use crate::document::{
    Edit as DocumentEdit, MAX_DOCUMENT_XML_BYTES, Snapshot as DocumentSnapshot, TransactionError,
    validate_paragraph_text,
};

use super::Package;

const TRANSITIONAL_WORD_NAMESPACE: &[u8] =
    b"http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const STRICT_WORD_NAMESPACE: &[u8] = b"http://purl.oclc.org/ooxml/wordprocessingml/main";
const MCE_NAMESPACE: &[u8] = b"http://schemas.openxmlformats.org/markup-compatibility/2006";
const XML_NAMESPACE: &[u8] = b"http://www.w3.org/XML/1998/namespace";
const DEFAULT_MAX_NAMESPACE_BINDINGS: usize = 256;
const DEFAULT_MAX_NAMESPACE_BYTES: usize = 4 * 1024 * 1024;
const MAX_STORY_REFERENCE_BYTES: usize = 64 * 1024;

/// A semantic selector for one source-backed Word story.
///
/// Header and footer indices are assigned after sorting the relationship
/// targets by canonical absolute part name. The selector never exposes an OPC
/// URI or relationship identifier.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum Selector {
    /// The package's unique main document part.
    Main,
    /// One header selected by canonical part-name order.
    Header(usize),
    /// One footer selected by canonical part-name order.
    Footer(usize),
    /// One normal footnote selected by its positive WordprocessingML ID.
    Footnote(u32),
    /// One normal endnote selected by its positive WordprocessingML ID.
    Endnote(u32),
    /// One comment selected by its WordprocessingML ID.
    Comment(u32),
}

impl Selector {
    /// Construct the main-story selector.
    #[must_use]
    pub const fn main() -> Self {
        Self::Main
    }

    /// Construct a header selector.
    #[must_use]
    pub const fn header(index: usize) -> Self {
        Self::Header(index)
    }

    /// Construct a footer selector.
    #[must_use]
    pub const fn footer(index: usize) -> Self {
        Self::Footer(index)
    }

    /// Construct a normal-footnote selector.
    #[must_use]
    pub const fn footnote(id: u32) -> Self {
        Self::Footnote(id)
    }

    /// Construct a normal-endnote selector.
    #[must_use]
    pub const fn endnote(id: u32) -> Self {
        Self::Endnote(id)
    }

    /// Construct a comment selector.
    #[must_use]
    pub const fn comment(id: u32) -> Self {
        Self::Comment(id)
    }
}

/// A source-backed story operation failure.
#[derive(Debug, ThisError)]
#[non_exhaustive]
pub enum Error {
    /// The underlying DOCX or OPC operation failed.
    #[error(transparent)]
    Document(#[from] crate::Error),
    /// The reusable main-document transaction refused the selected edit.
    #[error(transparent)]
    Transaction(#[from] TransactionError),
    /// A bounded story operation exceeded one finite resource ceiling.
    #[error("source-backed story text {resource} limit exceeded: {actual} > {maximum}")]
    Limit {
        /// Bounded resource name.
        resource: &'static str,
        /// Observed value.
        actual: usize,
        /// Maximum accepted value.
        maximum: usize,
    },
    /// A patch was applied to a source with different selected-part bytes.
    #[error("source-backed story text patch source is stale")]
    StaleSource,
    /// A patch was applied to a package lineage other than its source.
    #[error("source-backed story text patch belongs to a different source lineage")]
    ForeignSource,
    /// A patch was applied to a different selected story or package topology.
    #[error("source-backed story text patch conflicts with the selected story topology")]
    TopologyConflict,
    /// The requested secondary-story entry is absent.
    #[error("source-backed {story} entry {id} is absent")]
    MissingEntry {
        /// Secondary story kind.
        story: &'static str,
        /// Requested semantic entry ID.
        id: u32,
    },
    /// More than one secondary-story entry has the requested ID.
    #[error("source-backed {story} entry {id} is ambiguous")]
    AmbiguousEntry {
        /// Secondary story kind.
        story: &'static str,
        /// Requested semantic entry ID.
        id: u32,
    },
    /// The requested note ID or type is reserved rather than normal content.
    #[error("source-backed {story} entry {id} is reserved")]
    ReservedEntry {
        /// Secondary story kind.
        story: &'static str,
        /// Requested semantic entry ID.
        id: u32,
    },
    /// A publication inverse was applied to a foreign complete artifact.
    #[error("source-backed story text publication conflicts with the complete source artifact")]
    ArtifactConflict,
    /// A bounded allocation failed before publication began.
    #[error("source-backed story text allocation failed for {resource}: {source}")]
    Allocation {
        /// Allocation being attempted.
        resource: &'static str,
        /// Allocator failure.
        #[source]
        source: TryReserveError,
    },
}

/// Result returned by source-backed story-text operations.
pub type Result<T> = std::result::Result<T, Error>;

impl From<litchi_opc::OpcError> for Error {
    fn from(error: litchi_opc::OpcError) -> Self {
        Self::Document(crate::Error::Opc(error))
    }
}

/// Finite bounds for selected-story parsing, text output, and staged edits.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Limits {
    max_xml_bytes: usize,
    max_paragraphs: usize,
    max_events: usize,
    max_depth: usize,
    max_output_bytes: usize,
    max_replacements: usize,
    max_replacement_text_bytes: usize,
    max_entries: usize,
    max_entry_bytes: usize,
    max_namespace_bindings: usize,
    max_namespace_bytes: usize,
}

impl Limits {
    /// Construct a finite story-text policy.
    pub fn new(
        max_xml_bytes: usize,
        max_paragraphs: usize,
        max_events: usize,
        max_depth: usize,
        max_output_bytes: usize,
        max_replacements: usize,
        max_replacement_text_bytes: usize,
    ) -> Result<Self> {
        let values = [
            ("XML bytes", max_xml_bytes),
            ("paragraphs", max_paragraphs),
            ("XML events", max_events),
            ("XML depth", max_depth),
            ("output bytes", max_output_bytes),
            ("replacements", max_replacements),
            ("replacement text bytes", max_replacement_text_bytes),
        ];
        if let Some((resource, _)) = values.iter().find(|(_, value)| *value == 0) {
            return Err(Error::Limit {
                resource,
                actual: 0,
                maximum: 1,
            });
        }
        Ok(Self {
            max_xml_bytes,
            max_paragraphs,
            max_events,
            max_depth,
            max_output_bytes,
            max_replacements,
            max_replacement_text_bytes,
            max_entries: max_paragraphs,
            max_entry_bytes: max_xml_bytes,
            max_namespace_bindings: DEFAULT_MAX_NAMESPACE_BINDINGS,
            max_namespace_bytes: max_xml_bytes.min(DEFAULT_MAX_NAMESPACE_BYTES),
        })
    }

    /// Override finite bounds used while locating a secondary-story entry.
    pub fn with_secondary_entry_limits(
        mut self,
        max_entries: usize,
        max_entry_bytes: usize,
        max_namespace_bindings: usize,
    ) -> Result<Self> {
        let values = [
            ("story entries", max_entries),
            ("selected entry bytes", max_entry_bytes),
            ("namespace bindings", max_namespace_bindings),
        ];
        if let Some((resource, _)) = values.iter().find(|(_, value)| *value == 0) {
            return Err(Error::Limit {
                resource,
                actual: 0,
                maximum: 1,
            });
        }
        self.max_entries = max_entries;
        self.max_entry_bytes = max_entry_bytes;
        self.max_namespace_bindings = max_namespace_bindings;
        Ok(self)
    }

    /// Maximum selected story XML bytes.
    #[must_use]
    pub const fn max_xml_bytes(self) -> usize {
        self.max_xml_bytes
    }

    /// Maximum direct story paragraphs.
    #[must_use]
    pub const fn max_paragraphs(self) -> usize {
        self.max_paragraphs
    }

    /// Maximum XML parser events.
    #[must_use]
    pub const fn max_events(self) -> usize {
        self.max_events
    }

    /// Maximum XML nesting depth.
    #[must_use]
    pub const fn max_depth(self) -> usize {
        self.max_depth
    }

    /// Maximum generated story payload bytes.
    #[must_use]
    pub const fn max_output_bytes(self) -> usize {
        self.max_output_bytes
    }

    /// Maximum staged replacement count.
    #[must_use]
    pub const fn max_replacements(self) -> usize {
        self.max_replacements
    }

    /// Maximum aggregate authored replacement text bytes.
    #[must_use]
    pub const fn max_replacement_text_bytes(self) -> usize {
        self.max_replacement_text_bytes
    }

    /// Maximum direct entries inspected in a secondary-story part.
    #[must_use]
    pub const fn max_entries(self) -> usize {
        self.max_entries
    }

    /// Maximum bytes occupied by the selected secondary-story entry.
    #[must_use]
    pub const fn max_entry_bytes(self) -> usize {
        self.max_entry_bytes
    }

    /// Maximum in-scope XML namespace bindings.
    #[must_use]
    pub const fn max_namespace_bindings(self) -> usize {
        self.max_namespace_bindings
    }

    /// Maximum aggregate bytes retained for in-scope namespace bindings.
    #[must_use]
    pub const fn max_namespace_bytes(self) -> usize {
        self.max_namespace_bytes
    }

    /// Override the independent aggregate namespace-byte ceiling.
    pub fn with_max_namespace_bytes(mut self, max_namespace_bytes: usize) -> Result<Self> {
        if max_namespace_bytes == 0 {
            return Err(Error::Limit {
                resource: "namespace bytes",
                actual: 0,
                maximum: 1,
            });
        }
        self.max_namespace_bytes = max_namespace_bytes;
        Ok(self)
    }
}

impl Default for Limits {
    fn default() -> Self {
        Self {
            max_xml_bytes: 16 * 1024 * 1024,
            max_paragraphs: 65_536,
            max_events: 1_000_000,
            max_depth: 64,
            max_output_bytes: 32 * 1024 * 1024,
            max_replacements: 4_096,
            max_replacement_text_bytes: 8 * 1024 * 1024,
            max_entries: 65_536,
            max_entry_bytes: 16 * 1024 * 1024,
            max_namespace_bindings: DEFAULT_MAX_NAMESPACE_BINDINGS,
            max_namespace_bytes: DEFAULT_MAX_NAMESPACE_BYTES,
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
struct Range {
    start: usize,
    end: usize,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum RootKind {
    Main,
    Header,
    Footer,
    Footnotes,
    Endnotes,
    Comments,
}

impl RootKind {
    const fn entry_local_name(self) -> Option<&'static [u8]> {
        match self {
            Self::Footnotes => Some(b"footnote"),
            Self::Endnotes => Some(b"endnote"),
            Self::Comments => Some(b"comment"),
            Self::Main | Self::Header | Self::Footer => None,
        }
    }
}

#[derive(Clone)]
struct Layout {
    paragraphs: Arc<[Range]>,
    root_start_end: usize,
    root_end_start: usize,
    root_kind: RootKind,
    namespace_attributes: Arc<Vec<u8>>,
    word_namespace: Arc<[u8]>,
    has_w_namespace: bool,
}

#[derive(Debug, Clone, Copy)]
struct Envelope {
    inner_start: usize,
    inner_end: usize,
    wrapped_prefix_len: usize,
    wrapped_suffix_len: usize,
}

#[derive(Clone)]
enum Payload {
    Owned(Arc<Vec<u8>>),
    Managed(PartData),
}

impl Payload {
    fn as_bytes(&self) -> &[u8] {
        match self {
            Self::Owned(bytes) => bytes.as_slice(),
            Self::Managed(data) => data.as_bytes(),
        }
    }

    const fn is_managed(&self) -> bool {
        matches!(self, Self::Managed(_))
    }
}

/// An immutable selected-story snapshot.
pub struct Snapshot {
    selector: Selector,
    partname: PackURI,
    payload: Payload,
    layout: Layout,
    transaction: Option<DocumentSnapshot>,
    envelope: Option<Envelope>,
    source_version: SourceVersion,
    source_lineage: SourceLineage,
    part_fingerprint: [u8; 32],
    artifact_fingerprint: Option<SourceArtifactFingerprint>,
    limits: Limits,
    execution: Option<ExecutionContext>,
}

impl Clone for Snapshot {
    fn clone(&self) -> Self {
        Self {
            selector: self.selector,
            partname: self.partname.clone(),
            payload: self.payload.clone(),
            layout: self.layout.clone(),
            transaction: self.transaction.clone(),
            envelope: self.envelope,
            source_version: self.source_version,
            source_lineage: self.source_lineage.clone(),
            part_fingerprint: self.part_fingerprint,
            artifact_fingerprint: self.artifact_fingerprint,
            limits: self.limits,
            execution: self.execution.clone(),
        }
    }
}

impl Snapshot {
    fn check_execution(&self) -> Result<()> {
        check_execution_context(self.execution.as_ref())
    }

    /// Return the semantic selector captured by this snapshot.
    #[must_use]
    pub const fn selector(&self) -> Selector {
        self.selector
    }

    /// Return the exact source revision captured by this snapshot.
    #[must_use]
    pub const fn source_version(&self) -> SourceVersion {
        self.source_version
    }

    /// Return the SHA-256 digest of the selected story payload.
    #[must_use]
    pub const fn part_fingerprint(&self) -> [u8; 32] {
        self.part_fingerprint
    }

    /// Return the number of direct paragraphs in the selected story.
    pub fn paragraph_count(&self) -> Result<usize> {
        self.check_execution()?;
        Ok(self.layout.paragraphs.len())
    }

    /// Return one direct paragraph's caller-owned visible text.
    pub fn paragraph_text(&self, index: usize) -> Result<Option<String>> {
        self.check_execution()?;
        let Some(range) = self.layout.paragraphs.get(index).copied() else {
            return Ok(None);
        };
        let bytes = self
            .payload
            .as_bytes()
            .get(range.start..range.end)
            .ok_or_else(|| {
                Error::Document(crate::Error::InvalidFormat(
                    "story paragraph range is outside the selected XML".into(),
                ))
            })?;
        let mut text = String::new();
        extract_story_paragraph_text(
            bytes,
            &mut text,
            self.limits.max_output_bytes,
            "story paragraph text",
        )?;
        self.check_execution()?;
        Ok(Some(text))
    }

    /// Extract direct selected-story paragraph text into caller-owned memory.
    pub fn extract_text(&self) -> Result<String> {
        self.check_execution()?;
        let mut output = String::new();
        for (index, range) in self.layout.paragraphs.iter().copied().enumerate() {
            let bytes = self
                .payload
                .as_bytes()
                .get(range.start..range.end)
                .ok_or_else(|| {
                    Error::Document(crate::Error::InvalidFormat(
                        "story paragraph range is outside the selected XML".into(),
                    ))
                })?;
            extract_story_paragraph_text(
                bytes,
                &mut output,
                self.limits.max_output_bytes,
                "story text output",
            )?;
            if index % 64 == 0 {
                self.check_execution()?;
            }
        }
        self.check_execution()?;
        Ok(output)
    }

    /// Stream direct selected-story paragraph text through the shared bounded
    /// text-output contract.
    pub fn write_text_to<W: Write + ?Sized>(
        &self,
        output: &mut W,
        options: TextOutputOptions<'_>,
    ) -> std::result::Result<TextOutputReport, TextOutputError<Error>> {
        let story_max_output_bytes =
            u64::try_from(self.limits.max_output_bytes).unwrap_or(u64::MAX);
        let bounded_options = TextOutputOptions::new(
            options.paragraph_separator(),
            options.slide_separator(),
            options.max_output_bytes().min(story_max_output_bytes),
            options.max_objects(),
        )
        .with_empty_objects(options.include_empty_objects());
        let mut bounded_output = BoundedStoryWriter::new(output, self.limits.max_output_bytes);
        let mut writer =
            litchi_core::SequentialTextWriter::new(&mut bounded_output, bounded_options);
        if let Err(error) = self.check_execution() {
            return Err(writer.document_error(error));
        }
        for range in self.layout.paragraphs.iter().copied() {
            if let Err(error) = self.check_execution() {
                return Err(writer.document_error(error));
            }
            let bytes = match self.payload.as_bytes().get(range.start..range.end) {
                Some(bytes) => bytes,
                None => {
                    return Err(writer.document_error(Error::Document(
                        crate::Error::InvalidFormat(
                            "story paragraph range is outside the selected XML".into(),
                        ),
                    )));
                },
            };
            let paragraph_bytes = match measure_story_paragraph_text(bytes) {
                Ok(bytes) => bytes,
                Err(error) => return Err(writer.document_error(error)),
            };
            let emits_object = paragraph_bytes != 0 || bounded_options.include_empty_objects();
            let separator_bytes = if emits_object && writer.progress().objects_written() != 0 {
                bounded_options.paragraph_separator().len()
            } else {
                0
            };
            let required_output_bytes = separator_bytes
                .checked_add(paragraph_bytes)
                .and_then(|value| u64::try_from(value).ok())
                .and_then(|value| writer.progress().bytes_written().checked_add(value))
                .unwrap_or(u64::MAX);
            if emits_object && required_output_bytes > bounded_options.max_output_bytes() {
                let probe = output_limit_probe(bytes, paragraph_bytes);
                let complete_fragments = paragraph_bytes / probe.len();
                let remainder_bytes = paragraph_bytes % probe.len();
                return match writer.write_joined_object::<Error, _, _>(
                    TextObjectKind::Paragraph,
                    || {
                        std::iter::repeat_n(probe, complete_fragments)
                            .chain(std::iter::once(&probe[..remainder_bytes]))
                    },
                    "",
                ) {
                    Err(error) => Err(error),
                    Ok(()) => Err(writer.document_error(Error::Limit {
                        resource: "story text output",
                        actual: usize::try_from(required_output_bytes).unwrap_or(usize::MAX),
                        maximum: bounded_options
                            .max_output_bytes()
                            .try_into()
                            .unwrap_or(usize::MAX),
                    })),
                };
            }
            let remaining_output_bytes = bounded_options
                .max_output_bytes()
                .checked_sub(writer.progress().bytes_written())
                .and_then(|remaining| {
                    remaining.checked_sub(u64::try_from(separator_bytes).unwrap_or(u64::MAX))
                })
                .and_then(|remaining| usize::try_from(remaining).ok())
                .unwrap_or(0);
            let mut text = String::new();
            extract_story_paragraph_text(
                bytes,
                &mut text,
                remaining_output_bytes,
                "story paragraph text",
            )
            .map_err(|error| writer.document_error(error))?;
            writer.write_object::<Error>(TextObjectKind::Paragraph, &text)?;
        }
        if let Err(error) = self.check_execution() {
            return Err(writer.document_error(error));
        }
        Ok(writer.finish())
    }

    /// Start a staged direct-paragraph text edit.
    ///
    /// Managed source-backed packages support bounded read and stream access,
    /// but edits are refused rather than detaching a second owned XML buffer
    /// from the managed `PartData` reservation. This matches the established
    /// main-document transaction boundary.
    pub fn edit(&self) -> Result<Edit> {
        self.check_execution()?;
        if self.payload.is_managed() {
            return Err(Error::Document(crate::Error::UnsafeEdit {
                format: "DOCX",
                operation: "source-backed story text edit",
                reason: "managed source-backed story edits require an owned edit snapshot; use an unmanaged compatibility constructor",
            }));
        }
        let mut base = self.clone();
        let transaction = if let Some(transaction) = self.transaction.clone() {
            transaction
        } else if !matches!(self.selector, Selector::Main) {
            let effective_output_limit = self.limits.max_output_bytes.min(MAX_DOCUMENT_XML_BYTES);
            let (document_xml, envelope) =
                make_wrapped_document(self.raw_bytes(), &self.layout, effective_output_limit)?;
            base.envelope = Some(envelope);
            DocumentSnapshot::from_xml(document_xml)?
        } else {
            let document_xml = checked_vec_clone(self.raw_bytes(), "story transaction XML")?;
            DocumentSnapshot::from_xml(document_xml)?
        };
        base.transaction = Some(transaction.clone());
        Ok(Edit {
            base,
            projected: transaction.edit(),
            replacements: 0,
            replacement_text_bytes: 0,
        })
    }

    fn raw_bytes(&self) -> &[u8] {
        self.payload.as_bytes()
    }

    fn with_transaction_snapshot(&self, transaction: DocumentSnapshot) -> Result<Self> {
        let (raw, envelope) = if let Some(envelope) = self.envelope {
            let wrapped = transaction.shared_xml();
            let wrapped_end = wrapped
                .len()
                .checked_sub(envelope.wrapped_suffix_len)
                .ok_or_else(|| {
                    Error::Document(crate::Error::InvalidFormat(
                        "wrapped story suffix exceeds projected XML".into(),
                    ))
                })?;
            let inner = wrapped
                .get(envelope.wrapped_prefix_len..wrapped_end)
                .ok_or_else(|| {
                    Error::Document(crate::Error::InvalidFormat(
                        "wrapped story body is outside projected XML".into(),
                    ))
                })?;
            let before = self
                .raw_bytes()
                .get(..envelope.inner_start)
                .ok_or_else(|| {
                    Error::Document(crate::Error::InvalidFormat(
                        "story root prefix is outside source XML".into(),
                    ))
                })?;
            let after = self.raw_bytes().get(envelope.inner_end..).ok_or_else(|| {
                Error::Document(crate::Error::InvalidFormat(
                    "story root suffix is outside source XML".into(),
                ))
            })?;
            let total = before
                .len()
                .checked_add(inner.len())
                .and_then(|value| value.checked_add(after.len()))
                .ok_or(Error::Limit {
                    resource: "story output bytes",
                    actual: usize::MAX,
                    maximum: self.limits.max_output_bytes,
                })?;
            if total > self.limits.max_output_bytes {
                return Err(Error::Limit {
                    resource: "story output bytes",
                    actual: total,
                    maximum: self.limits.max_output_bytes,
                });
            }
            let mut raw = Vec::new();
            raw.try_reserve_exact(total)
                .map_err(|source| Error::Allocation {
                    resource: "story output bytes",
                    source,
                })?;
            raw.extend_from_slice(before);
            raw.extend_from_slice(inner);
            raw.extend_from_slice(after);
            (raw, Some(envelope))
        } else {
            (
                checked_story_output_clone(
                    transaction.shared_xml().as_ref(),
                    &self.limits,
                    "story transaction XML",
                )?,
                None,
            )
        };
        let payload = Payload::Owned(Arc::new(raw));
        let layout = scan_story(
            payload.as_bytes(),
            self.selector,
            Some(self.layout.word_namespace.as_ref()),
            self.limits,
            self.execution.as_ref(),
        )?;
        let envelope = envelope.map(|mut envelope| {
            envelope.inner_start = layout.root_start_end;
            envelope.inner_end = layout.root_end_start;
            envelope
        });
        let part_fingerprint = digest(payload.as_bytes());
        Ok(Self {
            selector: self.selector,
            partname: self.partname.clone(),
            payload,
            layout,
            transaction: Some(transaction),
            envelope,
            source_version: self.source_version,
            source_lineage: self.source_lineage.clone(),
            part_fingerprint,
            artifact_fingerprint: self.artifact_fingerprint,
            limits: self.limits,
            execution: self.execution.clone(),
        })
    }

    fn with_artifact_fingerprint(&self, fingerprint: SourceArtifactFingerprint) -> Self {
        let mut snapshot = self.clone();
        snapshot.artifact_fingerprint = Some(fingerprint);
        snapshot
    }
}

/// A staged selected-story direct-paragraph edit.
pub struct Edit {
    base: Snapshot,
    projected: DocumentEdit,
    replacements: usize,
    replacement_text_bytes: usize,
}

impl Edit {
    /// Borrow the immutable source snapshot.
    #[must_use]
    pub const fn source(&self) -> &Snapshot {
        &self.base
    }

    /// Borrow the current projected selected-story snapshot.
    pub fn projected(&self) -> Result<Snapshot> {
        self.base
            .with_transaction_snapshot(self.projected.projected().clone())
    }

    /// Replace all text in one direct selected-story paragraph.
    pub fn replace_paragraph_text(
        &mut self,
        position: Position,
        authored_text: impl AsRef<str>,
    ) -> Result<&mut Self> {
        self.base.check_execution()?;
        let next_replacements = self.replacements.checked_add(1).ok_or(Error::Limit {
            resource: "replacements",
            actual: usize::MAX,
            maximum: self.base.limits.max_replacements,
        })?;
        if next_replacements > self.base.limits.max_replacements {
            return Err(Error::Limit {
                resource: "replacements",
                actual: next_replacements,
                maximum: self.base.limits.max_replacements,
            });
        }
        let authored_text = authored_text.as_ref();
        let next_bytes = self
            .replacement_text_bytes
            .checked_add(authored_text.len())
            .ok_or(Error::Limit {
                resource: "replacement text bytes",
                actual: usize::MAX,
                maximum: self.base.limits.max_replacement_text_bytes,
            })?;
        if next_bytes > self.base.limits.max_replacement_text_bytes {
            return Err(Error::Limit {
                resource: "replacement text bytes",
                actual: next_bytes,
                maximum: self.base.limits.max_replacement_text_bytes,
            });
        }
        validate_paragraph_text(position, authored_text).map_err(Error::Transaction)?;
        let text = checked_string_clone(authored_text, "story replacement text")?;
        let mut candidate = self.projected.clone();
        candidate
            .replace_paragraph_text(position, text)
            .map_err(Error::Transaction)?;
        self.base.check_execution()?;
        let _validated = self
            .base
            .with_transaction_snapshot(candidate.projected().clone())?;
        self.base.check_execution()?;
        self.projected = candidate;
        self.replacements = next_replacements;
        self.replacement_text_bytes = next_bytes;
        Ok(self)
    }

    /// Finish the staged edit as a source-bound reversible patch.
    pub fn commit(self) -> Result<Commit> {
        self.base.check_execution()?;
        // The document transaction commit intentionally compacts changed
        // main-document XML. Collection stories use an envelope, however:
        // their projected snapshot already contains the exact raw paragraph
        // range replacements produced by the validated transaction. Reusing
        // that projection avoids normalizing unrelated bytes inside the
        // selected entry while preserving the outer source splice.
        let transaction_snapshot = if self.base.envelope.is_some() {
            self.projected.projected().clone()
        } else {
            self.projected.commit()?.snapshot().clone()
        };
        self.base.check_execution()?;
        let target = self.base.with_transaction_snapshot(transaction_snapshot)?;
        let patch = Patch {
            before: self.base,
            after: target.clone(),
        };
        Ok(Commit {
            snapshot: target,
            patch,
        })
    }
}

/// A completed selected-story edit.
pub struct Commit {
    snapshot: Snapshot,
    patch: Patch,
}

impl Commit {
    /// Borrow the projected selected-story snapshot.
    #[must_use]
    pub const fn snapshot(&self) -> &Snapshot {
        &self.snapshot
    }

    /// Borrow the source-bound reversible patch.
    #[must_use]
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }
}

/// A reversible selected-story patch bound to one source lineage, revision,
/// selected part, and exact selected-part bytes.
#[derive(Clone)]
pub struct Patch {
    before: Snapshot,
    after: Snapshot,
}

impl Patch {
    /// Borrow the exact source snapshot required by this patch.
    #[must_use]
    pub const fn source(&self) -> &Snapshot {
        &self.before
    }

    /// Borrow the exact projected target snapshot produced by this patch.
    #[must_use]
    pub const fn target(&self) -> &Snapshot {
        &self.after
    }

    /// Return whether the selected story bytes are unchanged.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.before.raw_bytes() == self.after.raw_bytes()
    }

    /// Return the exact inverse patch.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            before: self.after.clone(),
            after: self.before.clone(),
        }
    }

    /// Apply only to an equivalent selected-story source snapshot.
    pub fn apply(&self, source: &Snapshot) -> Result<Snapshot> {
        if source.selector != self.before.selector
            || !same_part_name(&source.partname, &self.before.partname)
        {
            return Err(Error::TopologyConflict);
        }
        if source.source_lineage != self.before.source_lineage {
            return Err(Error::ForeignSource);
        }
        if source.source_version != self.before.source_version {
            return Err(Error::StaleSource);
        }
        if source.part_fingerprint != self.before.part_fingerprint
            || source.raw_bytes() != self.before.raw_bytes()
        {
            return Err(Error::StaleSource);
        }
        Ok(if self.is_noop() {
            source.clone()
        } else {
            self.after.clone()
        })
    }
}

/// Exact publication evidence for one selected-story overlay.
pub struct Publication {
    snapshot: Snapshot,
    original_snapshot: Snapshot,
    original_artifact: SourceArtifact,
    published_fingerprint: SourceArtifactFingerprint,
    inverse_patch: Patch,
}

impl Publication {
    /// Borrow the selected-story snapshot represented by the emitted package.
    #[must_use]
    pub const fn snapshot(&self) -> &Snapshot {
        &self.snapshot
    }

    /// Borrow the inverse patch authorized for the emitted snapshot.
    #[must_use]
    pub const fn inverse_patch(&self) -> &Patch {
        &self.inverse_patch
    }

    /// Return the complete emitted-artifact fingerprint.
    #[must_use]
    pub const fn published_fingerprint(&self) -> SourceArtifactFingerprint {
        self.published_fingerprint
    }
}

impl Package {
    /// Capture one bounded source-backed Word text story.
    pub fn story_text_snapshot(&self, selector: Selector) -> Result<Snapshot> {
        self.story_text_snapshot_with_limits(selector, Limits::default())
    }

    /// Capture one selected story under explicit finite bounds.
    pub fn story_text_snapshot_with_limits(
        &self,
        selector: Selector,
        limits: Limits,
    ) -> Result<Snapshot> {
        capture(self, selector, limits)
    }

    /// Start one staged selected-story text edit.
    pub fn edit_story_text(&self, selector: Selector) -> Result<Edit> {
        self.story_text_snapshot(selector)?.edit()
    }

    /// Start one staged selected-story text edit under explicit bounds.
    pub fn edit_story_text_with_limits(&self, selector: Selector, limits: Limits) -> Result<Edit> {
        self.story_text_snapshot_with_limits(selector, limits)?
            .edit()
    }

    /// Publish one selected-story commit through the existing one-Part shared
    /// overlay path. Unchanged members are copied in their original physical
    /// form, and an exact no-op copies the complete source artifact.
    pub fn publish_story_text_commit_to_stream<W: Write>(
        self,
        writer: W,
        commit: &Commit,
    ) -> Result<Publication> {
        self.publish_story_text_patch_to_stream(writer, commit.patch())
    }

    /// Publish one source-bound selected-story patch.
    pub fn publish_story_text_patch_to_stream<W: Write>(
        self,
        writer: W,
        patch: &Patch,
    ) -> Result<Publication> {
        self.package.check_execution()?;
        let execution = self.package.execution_context();
        let current = capture(&self, patch.before.selector, patch.before.limits)?;
        let target = patch.apply(&current)?;
        let original_artifact = self.package.source_artifact();
        let mut output = FingerprintingWriter::new(writer);
        if patch.is_noop() {
            let mut cancellation_writer = CancellationWriter {
                writer: &mut output,
                execution: execution.as_ref(),
            };
            original_artifact.write_to_stream(&mut cancellation_writer)?;
            check_execution_context(execution.as_ref())?;
        } else {
            self.package.validate_topology_source_boundary()?;
            if target.raw_bytes().len() > patch.before.limits.max_output_bytes {
                return Err(Error::Limit {
                    resource: "story output bytes",
                    actual: target.raw_bytes().len(),
                    maximum: patch.before.limits.max_output_bytes,
                });
            }
            let replacement = Arc::new(checked_vec_clone(
                target.raw_bytes(),
                "story publication replacement",
            )?);
            self.package.write_part_overlay_shared_to_stream(
                &mut output,
                &current.partname,
                replacement,
            )?;
        }
        let published_fingerprint = output.finish();
        let published_snapshot = target.with_artifact_fingerprint(published_fingerprint);
        let inverse_patch = Patch {
            before: published_snapshot.clone(),
            after: current.clone(),
        };
        Ok(Publication {
            snapshot: published_snapshot,
            original_snapshot: current,
            original_artifact,
            published_fingerprint,
            inverse_patch,
        })
    }

    /// Restore the exact source artifact from a selected-story publication.
    pub fn publish_story_text_inverse_to_stream<W: Write>(
        self,
        writer: W,
        publication: &Publication,
    ) -> Result<Snapshot> {
        self.package.check_execution()?;
        let execution = self.package.execution_context();
        let current = self.package.source_artifact().fingerprint()?;
        if current != publication.published_fingerprint {
            return Err(Error::ArtifactConflict);
        }
        let mut writer = writer;
        let mut cancellation_writer = CancellationWriter {
            writer: &mut writer,
            execution: execution.as_ref(),
        };
        publication
            .original_artifact
            .write_to_stream(&mut cancellation_writer)?;
        check_execution_context(execution.as_ref())?;
        Ok(publication.original_snapshot.clone())
    }
}

fn capture(package: &Package, selector: Selector, limits: Limits) -> Result<Snapshot> {
    package.package.check_execution()?;
    let source_version = package.package.source_version()?;
    let source_lineage = package.package.source_lineage();
    let execution = package.package.execution_context();
    let (partname, part, relationship_word_namespace) = resolve_part(&package.package, selector)?;
    let declared = part.declared_uncompressed_size()?;
    if declared > limits.max_xml_bytes as u64 {
        return Err(Error::Limit {
            resource: "XML bytes",
            actual: usize::try_from(declared).unwrap_or(usize::MAX),
            maximum: limits.max_xml_bytes,
        });
    }
    let data = part.data()?;
    if data.as_bytes().len() > limits.max_xml_bytes {
        return Err(Error::Limit {
            resource: "XML bytes",
            actual: data.as_bytes().len(),
            maximum: limits.max_xml_bytes,
        });
    }
    let managed = package.package.cache_diagnostics().budget_managed;
    let payload = if managed {
        Payload::Managed(data)
    } else {
        Payload::Owned(data.into_arc()?)
    };
    let layout = scan_story(
        payload.as_bytes(),
        selector,
        relationship_word_namespace,
        limits,
        execution.as_ref(),
    )?;
    let (envelope, transaction) = (None, None);
    let part_fingerprint = digest(payload.as_bytes());
    let observed = package.package.source_version()?;
    if observed != source_version {
        return Err(Error::Document(crate::Error::Opc(
            litchi_opc::OpcError::SourceChanged {
                expected: source_version,
                actual: observed,
            },
        )));
    }
    package.package.check_execution()?;
    Ok(Snapshot {
        selector,
        partname,
        payload,
        layout,
        transaction,
        envelope,
        source_version,
        source_lineage,
        part_fingerprint,
        artifact_fingerprint: None,
        limits,
        execution,
    })
}

fn extract_story_paragraph_text(
    xml: &[u8],
    output: &mut String,
    maximum: usize,
    output_resource: &'static str,
) -> Result<()> {
    let mut paragraph_bytes = 0usize;
    crate::paragraph::for_each_word_text_chunk(xml, |chunk| {
        let next_paragraph_bytes =
            paragraph_bytes
                .checked_add(chunk.len())
                .ok_or(Error::Limit {
                    resource: "story paragraph text",
                    actual: usize::MAX,
                    maximum,
                })?;
        if next_paragraph_bytes > maximum {
            return Err(Error::Limit {
                resource: "story paragraph text",
                actual: next_paragraph_bytes,
                maximum,
            });
        }
        let next_output_bytes = output.len().checked_add(chunk.len()).ok_or(Error::Limit {
            resource: output_resource,
            actual: usize::MAX,
            maximum,
        })?;
        if next_output_bytes > maximum {
            return Err(Error::Limit {
                resource: output_resource,
                actual: next_output_bytes,
                maximum,
            });
        }
        output
            .try_reserve(chunk.len())
            .map_err(|source| Error::Allocation {
                resource: output_resource,
                source,
            })?;
        output.push_str(chunk);
        paragraph_bytes = next_paragraph_bytes;
        Ok(())
    })
}

fn measure_story_paragraph_text(xml: &[u8]) -> Result<usize> {
    let mut output_bytes = 0usize;
    crate::paragraph::for_each_word_text_chunk(xml, |chunk| {
        output_bytes = output_bytes.checked_add(chunk.len()).ok_or(Error::Limit {
            resource: "story paragraph text",
            actual: usize::MAX,
            maximum: usize::MAX,
        })?;
        Ok::<(), Error>(())
    })?;
    Ok(output_bytes)
}

fn output_limit_probe(xml: &[u8], desired_bytes: usize) -> &str {
    let mut length = xml.len().min(desired_bytes).min(4096);
    while length != 0 {
        let prefix = &xml[..length];
        if prefix.iter().all(u8::is_ascii) {
            if let Ok(probe) = std::str::from_utf8(prefix) {
                return probe;
            }
        }
        length -= 1;
    }
    "x"
}

fn validate_story_general_reference(reference: &quick_xml::events::BytesRef<'_>) -> Result<()> {
    let bytes = reference.as_ref();
    if bytes.len() > MAX_STORY_REFERENCE_BYTES {
        return Err(Error::Limit {
            resource: "XML reference bytes",
            actual: bytes.len(),
            maximum: MAX_STORY_REFERENCE_BYTES,
        });
    }
    if !is_supported_story_xml_reference(bytes) {
        return Err(unsupported_xml(
            "unsupported XML entity reference in a source-bound story",
        ));
    }
    decode_xml_reference(reference).map_err(|error| {
        Error::Document(crate::Error::Xml(format!(
            "invalid XML entity reference: {error}"
        )))
    })?;
    Ok(())
}

fn is_supported_story_xml_reference(reference: &[u8]) -> bool {
    if matches!(reference, b"lt" | b"gt" | b"amp" | b"apos" | b"quot") {
        return true;
    }
    let (digits, hexadecimal) = if reference.first() == Some(&b'#') {
        if matches!(reference.get(1), Some(b'x' | b'X')) {
            (reference.get(2..).unwrap_or_default(), true)
        } else {
            (reference.get(1..).unwrap_or_default(), false)
        }
    } else {
        return false;
    };
    !digits.is_empty()
        && digits.iter().all(|digit| {
            if hexadecimal {
                digit.is_ascii_hexdigit()
            } else {
                digit.is_ascii_digit()
            }
        })
}

fn resolve_part<'package>(
    package: &'package SourceBackedPackage,
    selector: Selector,
) -> Result<(PackURI, PartView<'package>, Option<&'static [u8]>)> {
    let execution = package.execution_context();
    check_execution_context(execution.as_ref())?;
    if let Some((story, id)) = secondary_identity(selector)
        && id == 0
        && story != "comment"
    {
        return Err(Error::ReservedEntry { story, id });
    }
    let package_dialect = package_word_dialect(package)?;
    reject_mixed_word_story_relationship_dialects(package, package_dialect, execution.as_ref())?;
    match selector {
        Selector::Main => {
            let main = package.main_document_part()?;
            if main.content_type() != ct::WML_DOCUMENT_MAIN {
                return Err(Error::Document(crate::Error::InvalidContentType {
                    expected: ct::WML_DOCUMENT_MAIN.to_owned(),
                    got: main.content_type().to_owned(),
                }));
            }
            Ok((
                main.partname().clone(),
                main,
                Some(package_dialect.word_namespace()),
            ))
        },
        Selector::Header(index) => {
            let (partname, part, relationship_namespace) =
                resolve_header_footer(package, index, true)?;
            Ok((partname, part, Some(relationship_namespace)))
        },
        Selector::Footer(index) => {
            let (partname, part, relationship_namespace) =
                resolve_header_footer(package, index, false)?;
            Ok((partname, part, Some(relationship_namespace)))
        },
        Selector::Footnote(id) => resolve_secondary_part(package, id, SecondaryProfile::Footnote),
        Selector::Endnote(id) => resolve_secondary_part(package, id, SecondaryProfile::Endnote),
        Selector::Comment(id) => resolve_secondary_part(package, id, SecondaryProfile::Comment),
    }
}

#[derive(Clone, Copy)]
enum SecondaryProfile {
    Footnote,
    Endnote,
    Comment,
}

impl SecondaryProfile {
    const fn story(self) -> &'static str {
        match self {
            Self::Footnote => "footnote",
            Self::Endnote => "endnote",
            Self::Comment => "comment",
        }
    }

    const fn relationship_types(self) -> (&'static str, &'static str) {
        match self {
            Self::Footnote => (rt::FOOTNOTES, STRICT_FOOTNOTES),
            Self::Endnote => (rt::ENDNOTES, STRICT_ENDNOTES),
            Self::Comment => (rt::COMMENTS, rt::STRICT_COMMENTS),
        }
    }

    const fn content_type(self) -> &'static str {
        match self {
            Self::Footnote => ct::WML_FOOTNOTES,
            Self::Endnote => ct::WML_ENDNOTES,
            Self::Comment => ct::WML_COMMENTS,
        }
    }
}

fn resolve_secondary_part<'package>(
    package: &'package SourceBackedPackage,
    id: u32,
    profile: SecondaryProfile,
) -> Result<(PackURI, PartView<'package>, Option<&'static [u8]>)> {
    if id == 0 && !matches!(profile, SecondaryProfile::Comment) {
        return Err(Error::ReservedEntry {
            story: profile.story(),
            id,
        });
    }
    let execution = package.execution_context();
    let main = package.main_document_part()?;
    if main.content_type() != ct::WML_DOCUMENT_MAIN {
        return Err(Error::Document(crate::Error::InvalidContentType {
            expected: ct::WML_DOCUMENT_MAIN.to_owned(),
            got: main.content_type().to_owned(),
        }));
    }
    let (transitional, strict) = profile.relationship_types();
    let mut selected: Option<(PackURI, &'static [u8])> = None;
    for relationship in main.rels().iter() {
        check_execution_context(execution.as_ref())?;
        let relationship_namespace = if relationship.reltype() == transitional {
            Some(TRANSITIONAL_WORD_NAMESPACE)
        } else if relationship.reltype() == strict {
            Some(STRICT_WORD_NAMESPACE)
        } else {
            None
        };
        let Some(relationship_namespace) = relationship_namespace else {
            continue;
        };
        if relationship.is_external() {
            return Err(Error::Document(crate::Error::InvalidRelationship(
                "secondary story relationship cannot be external".into(),
            )));
        }
        let target = relationship.target_partname()?;
        if selected.is_some() {
            return Err(Error::Document(crate::Error::InvalidRelationship(
                "secondary story has multiple owning relationships".into(),
            )));
        }
        let in_word_namespace = target
            .as_str()
            .get(..6)
            .is_some_and(|prefix| prefix.eq_ignore_ascii_case("/word/"));
        if !in_word_namespace {
            return Err(Error::Document(crate::Error::InvalidRelationship(
                "secondary story target is outside the Word part namespace".into(),
            )));
        }
        selected = Some((target, relationship_namespace));
    }
    let Some((target, relationship_namespace)) = selected else {
        return Err(Error::Document(crate::Error::InvalidRelationship(format!(
            "{} relationship is missing",
            profile.story()
        ))));
    };
    let part = package.part(&target)?;
    reject_outbound_story_relationships(&part, execution.as_ref())?;
    if part.content_type() != profile.content_type() {
        return Err(Error::Document(crate::Error::InvalidContentType {
            expected: profile.content_type().to_owned(),
            got: part.content_type().to_owned(),
        }));
    }
    if inbound_count(package, &target)? != 1 {
        return Err(Error::Document(crate::Error::InvalidRelationship(
            "secondary story target has an ambiguous inbound relationship closure".into(),
        )));
    }
    Ok((target, part, Some(relationship_namespace)))
}

#[derive(Clone, Copy, PartialEq, Eq)]
enum WordDialect {
    Transitional,
    Strict,
}

impl WordDialect {
    const fn word_namespace(self) -> &'static [u8] {
        match self {
            Self::Transitional => TRANSITIONAL_WORD_NAMESPACE,
            Self::Strict => STRICT_WORD_NAMESPACE,
        }
    }
}

fn package_word_dialect(package: &SourceBackedPackage) -> Result<WordDialect> {
    let execution = package.execution_context();
    let mut dialect = None;
    for relationship in package.rels().iter() {
        check_execution_context(execution.as_ref())?;
        let Some(candidate) = office_document_relationship_dialect(relationship.reltype()) else {
            continue;
        };
        if relationship.is_external() {
            return Err(Error::Document(crate::Error::InvalidRelationship(
                "root office document relationship cannot be external".into(),
            )));
        }
        if dialect.replace(candidate).is_some() {
            return Err(Error::Document(crate::Error::InvalidRelationship(
                "package has multiple office document relationships".into(),
            )));
        }
    }
    dialect.ok_or_else(|| {
        Error::Document(crate::Error::InvalidRelationship(
            "package office document relationship is missing".into(),
        ))
    })
}

fn office_document_relationship_dialect(reltype: &str) -> Option<WordDialect> {
    match reltype {
        rt::OFFICE_DOCUMENT => Some(WordDialect::Transitional),
        rt::STRICT_OFFICE_DOCUMENT => Some(WordDialect::Strict),
        _ => None,
    }
}

fn word_story_relationship_dialect(reltype: &str) -> Option<WordDialect> {
    match reltype {
        rt::HEADER
        | rt::FOOTER
        | rt::FOOTNOTES
        | rt::ENDNOTES
        | rt::COMMENTS
        | TRANSITIONAL_GLOSSARY => Some(WordDialect::Transitional),
        STRICT_HEADER
        | STRICT_FOOTER
        | STRICT_FOOTNOTES
        | STRICT_ENDNOTES
        | rt::STRICT_COMMENTS
        | STRICT_GLOSSARY => Some(WordDialect::Strict),
        _ => None,
    }
}

fn reject_mixed_word_story_relationship_dialects(
    package: &SourceBackedPackage,
    package_dialect: WordDialect,
    execution: Option<&ExecutionContext>,
) -> Result<()> {
    for relationship in package.rels().iter() {
        check_execution_context(execution)?;
        if word_story_relationship_dialect(relationship.reltype())
            .is_some_and(|dialect| dialect != package_dialect)
        {
            return Err(Error::Document(crate::Error::InvalidRelationship(
                "Word story relationship uses a dialect different from the package root".into(),
            )));
        }
    }
    for part in package.iter_parts() {
        check_execution_context(execution)?;
        for relationship in part.rels().iter() {
            check_execution_context(execution)?;
            if word_story_relationship_dialect(relationship.reltype())
                .is_some_and(|dialect| dialect != package_dialect)
            {
                return Err(Error::Document(crate::Error::InvalidRelationship(
                    "Word story relationship uses a dialect different from the package root".into(),
                )));
            }
        }
    }
    Ok(())
}

fn reject_outbound_story_relationships(
    part: &PartView<'_>,
    execution: Option<&ExecutionContext>,
) -> Result<()> {
    for relationship in part.rels().iter() {
        check_execution_context(execution)?;
        if word_story_relationship_dialect(relationship.reltype()).is_some() {
            return Err(Error::Document(crate::Error::InvalidRelationship(
                "selected story part has an outbound Word story relationship".into(),
            )));
        }
    }
    Ok(())
}

fn resolve_header_footer<'package>(
    package: &'package SourceBackedPackage,
    index: usize,
    header: bool,
) -> Result<(PackURI, PartView<'package>, &'static [u8])> {
    let execution = package.execution_context();
    let main = package.main_document_part()?;
    if main.content_type() != ct::WML_DOCUMENT_MAIN {
        return Err(Error::Document(crate::Error::InvalidContentType {
            expected: ct::WML_DOCUMENT_MAIN.to_owned(),
            got: main.content_type().to_owned(),
        }));
    }
    let relationship_types = if header {
        [rt::HEADER, STRICT_HEADER]
    } else {
        [rt::FOOTER, STRICT_FOOTER]
    };
    check_execution_context(execution.as_ref())?;
    let target_limit = package.iter_parts().count();
    let mut targets: Vec<(PackURI, &'static [u8])> = Vec::new();
    targets.try_reserve(1).map_err(|source| Error::Allocation {
        resource: "story relationship targets",
        source,
    })?;
    for relationship in main.rels().iter() {
        check_execution_context(execution.as_ref())?;
        if !relationship_types.contains(&relationship.reltype()) {
            continue;
        }
        let relationship_namespace = if relationship.reltype() == relationship_types[0] {
            TRANSITIONAL_WORD_NAMESPACE
        } else {
            STRICT_WORD_NAMESPACE
        };
        if relationship.is_external() {
            return Err(Error::Document(crate::Error::InvalidRelationship(
                "header/footer relationship cannot be external".into(),
            )));
        }
        let target = relationship.target_partname()?;
        if targets
            .iter()
            .any(|candidate| same_part_name(&candidate.0, &target))
        {
            return Err(Error::Document(crate::Error::InvalidRelationship(
                "multiple header/footer relationships target the same part".into(),
            )));
        }
        if !target.as_str().to_ascii_lowercase().starts_with("/word/") {
            return Err(Error::Document(crate::Error::InvalidRelationship(
                "header/footer target is outside the Word part namespace".into(),
            )));
        }
        let part = package.part(&target)?;
        let expected = if header {
            ct::WML_HEADER
        } else {
            ct::WML_FOOTER
        };
        if part.content_type() != expected {
            return Err(Error::Document(crate::Error::InvalidContentType {
                expected: expected.to_owned(),
                got: part.content_type().to_owned(),
            }));
        }
        reject_outbound_story_relationships(&part, execution.as_ref())?;
        let inbound = inbound_count(package, &target)?;
        if inbound != 1 {
            return Err(Error::Document(crate::Error::InvalidRelationship(
                "header/footer target has an ambiguous inbound relationship closure".into(),
            )));
        }

        if targets.len() >= target_limit {
            return Err(Error::Limit {
                resource: "story relationship targets",
                actual: targets.len().saturating_add(1),
                maximum: target_limit,
            });
        }
        targets.try_reserve(1).map_err(|source| Error::Allocation {
            resource: "story relationship targets",
            source,
        })?;
        targets.push((target, relationship_namespace));
    }
    targets.sort_by(|left, right| left.0.as_str().cmp(right.0.as_str()));
    let len = targets.len();
    let target = targets
        .get(index)
        .ok_or(Error::Document(crate::Error::OutOfBounds {
            object: if header { "header" } else { "footer" },
            index,
            len,
        }))?;
    let part = package.part(&target.0)?;
    Ok((target.0.clone(), part, target.1))
}

fn inbound_count(package: &SourceBackedPackage, target: &PackURI) -> Result<usize> {
    let execution = package.execution_context();
    let mut count = 0usize;
    for relationship in package.rels().iter() {
        check_execution_context(execution.as_ref())?;
        if !relationship.is_external() && same_part_name(&relationship.target_partname()?, target) {
            count = count.checked_add(1).ok_or_else(|| {
                Error::Document(crate::Error::InvalidFormat(
                    "header/footer inbound relationship count overflow".into(),
                ))
            })?;
        }
    }
    for part in package.iter_parts() {
        check_execution_context(execution.as_ref())?;
        for relationship in part.rels().iter() {
            check_execution_context(execution.as_ref())?;
            if !relationship.is_external()
                && same_part_name(&relationship.target_partname()?, target)
            {
                count = count.checked_add(1).ok_or_else(|| {
                    Error::Document(crate::Error::InvalidFormat(
                        "header/footer inbound relationship count overflow".into(),
                    ))
                })?;
            }
        }
    }
    Ok(count)
}

#[derive(Clone, Default)]
struct NamespaceContext {
    bindings: Vec<(Arc<Vec<u8>>, Arc<Vec<u8>>)>,
    raw_namespace_bytes: usize,
}

fn decode_attribute_value(value: &[u8]) -> Result<Vec<u8>> {
    let value = std::str::from_utf8(value).map_err(|error| {
        Error::Document(crate::Error::InvalidFormat(format!(
            "invalid XML attribute UTF-8: {error}"
        )))
    })?;
    let mut decoded = Vec::new();
    decoded
        .try_reserve_exact(value.len())
        .map_err(|source| Error::Allocation {
            resource: "story namespace value",
            source,
        })?;
    let mut offset = 0usize;
    while offset < value.len() {
        if value.as_bytes()[offset] != b'&' {
            let start = offset;
            while offset < value.len() && value.as_bytes()[offset] != b'&' {
                offset += 1;
            }
            decoded.extend_from_slice(&value.as_bytes()[start..offset]);
            continue;
        }
        let entity_start = offset + 1;
        let entity_end = value.as_bytes()[entity_start..]
            .iter()
            .position(|byte| *byte == b';')
            .map(|relative| entity_start + relative)
            .ok_or_else(|| {
                Error::Document(crate::Error::InvalidFormat(
                    "unterminated XML attribute escape".into(),
                ))
            })?;
        let entity = &value.as_bytes()[entity_start..entity_end];
        let replacement: &[u8];
        let mut numeric = [0u8; 4];
        let numeric_length;
        if entity == b"lt" {
            replacement = b"<";
        } else if entity == b"gt" {
            replacement = b">";
        } else if entity == b"amp" {
            replacement = b"&";
        } else if entity == b"apos" {
            replacement = b"'";
        } else if entity == b"quot" {
            replacement = b"\"";
        } else if entity.first() == Some(&b'#') {
            let (radix, digits) = if entity.get(1) == Some(&b'x') || entity.get(1) == Some(&b'X') {
                (16, &entity[2..])
            } else {
                (10, &entity[1..])
            };
            if digits.is_empty() {
                return Err(Error::Document(crate::Error::InvalidFormat(
                    "empty XML character reference".into(),
                )));
            }
            let mut codepoint = 0u32;
            for digit in digits {
                let digit_value = if digit.is_ascii_digit() {
                    u32::from(*digit - b'0')
                } else if radix == 16 && digit.is_ascii_hexdigit() {
                    u32::from(digit.to_ascii_lowercase() - b'a' + 10)
                } else {
                    return Err(Error::Document(crate::Error::InvalidFormat(
                        "invalid XML character reference".into(),
                    )));
                };
                codepoint = codepoint
                    .checked_mul(radix)
                    .and_then(|value| value.checked_add(digit_value))
                    .ok_or_else(|| {
                        Error::Document(crate::Error::InvalidFormat(
                            "XML character reference is out of range".into(),
                        ))
                    })?;
            }
            let character = char::from_u32(codepoint).ok_or_else(|| {
                Error::Document(crate::Error::InvalidFormat(
                    "XML character reference is not a Unicode scalar".into(),
                ))
            })?;
            numeric_length = character.encode_utf8(&mut numeric).len();
            replacement = &numeric[..numeric_length];
        } else {
            return Err(Error::Document(crate::Error::InvalidFormat(
                "unsupported XML attribute escape".into(),
            )));
        }
        decoded
            .try_reserve(replacement.len())
            .map_err(|source| Error::Allocation {
                resource: "story namespace value",
                source,
            })?;
        decoded.extend_from_slice(replacement);
        offset = entity_end + 1;
    }
    Ok(decoded)
}

impl NamespaceContext {
    fn resolve_prefix(&self, prefix: &[u8]) -> Option<&[u8]> {
        if prefix == b"xml" {
            return Some(XML_NAMESPACE);
        }
        self.bindings
            .iter()
            .rev()
            .find(|(candidate, _)| candidate.as_slice() == prefix)
            .and_then(|(_, namespace)| (!namespace.is_empty()).then_some(namespace.as_slice()))
    }

    fn for_element(
        parent: &Self,
        start: &BytesStart<'_>,
        max_bindings: usize,
        max_namespace_bytes: usize,
    ) -> Result<(Self, Vec<u8>)> {
        if parent.bindings.len() > max_bindings {
            return Err(Error::Limit {
                resource: "namespace bindings",
                actual: parent.bindings.len(),
                maximum: max_bindings,
            });
        }
        if parent.raw_namespace_bytes > max_namespace_bytes {
            return Err(Error::Limit {
                resource: "namespace bytes",
                actual: parent.raw_namespace_bytes,
                maximum: max_namespace_bytes,
            });
        }
        let mut context = Self {
            bindings: Vec::new(),
            raw_namespace_bytes: parent.raw_namespace_bytes,
        };
        context
            .bindings
            .try_reserve_exact(parent.bindings.len())
            .map_err(|source| Error::Allocation {
                resource: "story namespace bindings",
                source,
            })?;
        for (prefix, value) in &parent.bindings {
            context
                .bindings
                .push((Arc::clone(prefix), Arc::clone(value)));
        }

        let mut namespace_attributes = Vec::new();
        for attribute in start.attributes().with_checks(true) {
            let attribute = attribute.map_err(|error| {
                Error::Document(crate::Error::Xml(format!(
                    "invalid story attribute: {error}"
                )))
            })?;
            let key = attribute.key;
            let value = attribute.value.as_ref();
            if is_namespace_declaration(key) {
                let prefix = namespace_declared_prefix(key).unwrap_or_default();
                let raw_declaration_bytes =
                    prefix.len().checked_add(value.len()).ok_or(Error::Limit {
                        resource: "namespace bytes",
                        actual: usize::MAX,
                        maximum: max_namespace_bytes,
                    })?;
                let raw_namespace_bytes = context
                    .raw_namespace_bytes
                    .checked_add(raw_declaration_bytes)
                    .ok_or(Error::Limit {
                        resource: "namespace bytes",
                        actual: usize::MAX,
                        maximum: max_namespace_bytes,
                    })?;
                if raw_namespace_bytes > max_namespace_bytes {
                    return Err(Error::Limit {
                        resource: "namespace bytes",
                        actual: raw_namespace_bytes,
                        maximum: max_namespace_bytes,
                    });
                }
                let decoded_value = decode_attribute_value(value)?;
                if decoded_value.as_slice() == MCE_NAMESPACE {
                    return Err(unsupported_xml(
                        "markup-compatibility namespace declarations are not accepted in a source-bound story",
                    ));
                }
                if context.bindings.len() >= max_bindings {
                    return Err(Error::Limit {
                        resource: "namespace bindings",
                        actual: context.bindings.len().saturating_add(1),
                        maximum: max_bindings,
                    });
                }
                let prefix_copy = checked_vec_clone(prefix, "story namespace prefix")?;
                context
                    .bindings
                    .try_reserve(1)
                    .map_err(|source| Error::Allocation {
                        resource: "story namespace bindings",
                        source,
                    })?;
                context
                    .bindings
                    .push((Arc::new(prefix_copy), Arc::new(decoded_value)));
                context.raw_namespace_bytes = raw_namespace_bytes;
                let attribute_bytes = 1usize
                    .checked_add(key.as_ref().len())
                    .and_then(|size| size.checked_add(3))
                    .and_then(|size| size.checked_add(value.len()))
                    .ok_or(Error::Limit {
                        resource: "story namespace attributes",
                        actual: usize::MAX,
                        maximum: usize::MAX,
                    })?;
                namespace_attributes
                    .try_reserve(attribute_bytes)
                    .map_err(|source| Error::Allocation {
                        resource: "story namespace attributes",
                        source,
                    })?;
                namespace_attributes.extend_from_slice(b" ");
                namespace_attributes.extend_from_slice(key.as_ref());
                namespace_attributes.extend_from_slice(b"=\"");
                namespace_attributes.extend_from_slice(value);
                namespace_attributes.push(b'\"');
            }
        }

        for attribute in start.attributes().with_checks(true) {
            let attribute = attribute.map_err(|error| {
                Error::Document(crate::Error::Xml(format!(
                    "invalid story attribute: {error}"
                )))
            })?;
            let key = attribute.key;
            let value = attribute.value.as_ref();
            if !is_namespace_declaration(key)
                && let Some(prefix) = key.prefix()
                && context.resolve_prefix(prefix.as_ref()).is_none()
            {
                return Err(unsupported_xml("undeclared attribute namespace prefix"));
            }
            let decoded_value = if !is_namespace_declaration(key) {
                Some(decode_attribute_value(value)?)
            } else {
                None
            };
            if !is_namespace_declaration(key)
                && (is_mce_attribute(key, decoded_value.as_deref().unwrap_or_default())
                    || context
                        .resolve_attribute(key)
                        .is_some_and(|namespace| namespace == MCE_NAMESPACE))
            {
                return Err(unsupported_xml(
                    "markup-compatibility attributes are not accepted in a source-bound story",
                ));
            }
        }
        Ok((context, namespace_attributes))
    }

    fn wrapper_namespace_attributes(
        &self,
        word_namespace: &[u8],
        max_bytes: usize,
    ) -> Result<(Vec<u8>, bool)> {
        let mut seen: Vec<&[u8]> = Vec::new();
        seen.try_reserve_exact(self.bindings.len())
            .map_err(|source| Error::Allocation {
                resource: "story namespace binding projection",
                source,
            })?;
        let mut attributes = Vec::new();
        for (prefix, namespace) in self.bindings.iter().rev() {
            if seen.iter().any(|candidate| *candidate == prefix.as_slice()) {
                continue;
            }
            seen.push(prefix.as_slice());
            if prefix.as_slice() == b"xml" {
                continue;
            }
            if prefix.as_slice() == b"w" {
                if namespace.is_empty() {
                    continue;
                }
                if namespace.as_slice() != word_namespace {
                    return Err(unsupported_xml(
                        "the w prefix is rebound inside the selected story entry",
                    ));
                }
            }
            append_namespace_attribute(&mut attributes, prefix, namespace, max_bytes)?;
        }
        let has_w_namespace = self.resolve_prefix(b"w") == Some(word_namespace);
        Ok((attributes, has_w_namespace))
    }

    fn resolve_element(&self, name: QName<'_>) -> Option<&[u8]> {
        let prefix = name.prefix();
        match prefix {
            Some(prefix) => self
                .bindings
                .iter()
                .rev()
                .find(|(candidate, _)| candidate.as_slice() == prefix.as_ref())
                .and_then(|(_, value)| (!value.is_empty()).then_some(value.as_slice())),
            None => self
                .bindings
                .iter()
                .rev()
                .find(|(candidate, _)| candidate.is_empty())
                .and_then(|(_, value)| (!value.is_empty()).then_some(value.as_slice())),
        }
    }

    fn resolve_attribute(&self, name: QName<'_>) -> Option<&[u8]> {
        let prefix = name.prefix()?;
        self.bindings
            .iter()
            .rev()
            .find(|(candidate, _)| candidate.as_slice() == prefix.as_ref())
            .and_then(|(_, value)| (!value.is_empty()).then_some(value.as_slice()))
    }
}

fn append_namespace_attribute(
    output: &mut Vec<u8>,
    prefix: &[u8],
    namespace: &[u8],
    maximum: usize,
) -> Result<()> {
    let worst_case = namespace
        .len()
        .checked_mul(5)
        .and_then(|value| value.checked_add(prefix.len()))
        .and_then(|value| value.checked_add(11))
        .ok_or(Error::Limit {
            resource: "story namespace attributes",
            actual: usize::MAX,
            maximum,
        })?;
    let projected = output.len().checked_add(worst_case).ok_or(Error::Limit {
        resource: "story namespace attributes",
        actual: usize::MAX,
        maximum,
    })?;
    if projected > maximum {
        return Err(Error::Limit {
            resource: "story namespace attributes",
            actual: projected,
            maximum,
        });
    }
    output
        .try_reserve(worst_case)
        .map_err(|source| Error::Allocation {
            resource: "story namespace attributes",
            source,
        })?;
    output.extend_from_slice(b" xmlns");
    if !prefix.is_empty() {
        output.push(b':');
        output.extend_from_slice(prefix);
    }
    output.extend_from_slice(b"=\"");
    for byte in namespace {
        match byte {
            b'&' => output.extend_from_slice(b"&amp;"),
            b'<' => output.extend_from_slice(b"&lt;"),
            b'\"' => output.extend_from_slice(b"&quot;"),
            _ => output.push(*byte),
        }
    }
    output.push(b'\"');
    Ok(())
}

fn reader_position(reader: &Reader<&[u8]>) -> Result<usize> {
    let position = reader.buffer_position();
    usize::try_from(position).map_err(|_| {
        Error::Document(crate::Error::Xml(format!(
            "XML buffer position {position} exceeds addressable input",
        )))
    })
}

fn secondary_identity(selector: Selector) -> Option<(&'static str, u32)> {
    match selector {
        Selector::Footnote(id) => Some(("footnote", id)),
        Selector::Endnote(id) => Some(("endnote", id)),
        Selector::Comment(id) => Some(("comment", id)),
        Selector::Main | Selector::Header(_) | Selector::Footer(_) => None,
    }
}

fn selector_matches_entry(
    start: &BytesStart<'_>,
    context: &NamespaceContext,
    word_namespace: &[u8],
    selector: Selector,
) -> Result<bool> {
    let Some((story, selected_id)) = secondary_identity(selector) else {
        return Ok(false);
    };
    let id = word_attribute(start, context, word_namespace, b"id")
        .transpose()?
        .ok_or_else(|| unsupported_xml("secondary story entry is missing w:id"))?;
    let id = std::str::from_utf8(&id)
        .ok()
        .and_then(|value| value.parse::<i64>().ok())
        .ok_or_else(|| unsupported_xml("secondary story entry has an invalid w:id"))?;
    let matches = match selector {
        Selector::Footnote(_) | Selector::Endnote(_) => id == i64::from(selected_id),
        Selector::Comment(_) => u32::try_from(id).ok() == Some(selected_id),
        Selector::Main | Selector::Header(_) | Selector::Footer(_) => false,
    };
    if !matches {
        return Ok(false);
    }
    if matches!(selector, Selector::Footnote(_) | Selector::Endnote(_)) {
        if selected_id == 0 {
            return Err(Error::ReservedEntry {
                story,
                id: selected_id,
            });
        }
        let entry_type = word_attribute(start, context, word_namespace, b"type").transpose()?;
        if entry_type
            .as_deref()
            .is_some_and(|entry_type| entry_type != b"normal")
        {
            return Err(Error::ReservedEntry {
                story,
                id: selected_id,
            });
        }
    }
    Ok(true)
}

fn word_attribute(
    start: &BytesStart<'_>,
    context: &NamespaceContext,
    word_namespace: &[u8],
    local_name: &[u8],
) -> Option<Result<Vec<u8>>> {
    let mut selected = None;
    for attribute in start.attributes().with_checks(true) {
        let attribute = match attribute {
            Ok(attribute) => attribute,
            Err(error) => {
                return Some(Err(Error::Document(crate::Error::Xml(format!(
                    "invalid story attribute: {error}"
                )))));
            },
        };
        if attribute.key.local_name().as_ref() != local_name
            || context.resolve_attribute(attribute.key) != Some(word_namespace)
        {
            continue;
        }
        if selected.is_some() {
            return Some(Err(unsupported_xml(
                "secondary story entry repeats a WordprocessingML attribute",
            )));
        }
        selected = Some(decode_attribute_value(attribute.value.as_ref()));
    }
    selected
}

fn scan_story(
    xml: &[u8],
    selector: Selector,
    expected_word_namespace: Option<&[u8]>,
    limits: Limits,
    execution: Option<&ExecutionContext>,
) -> Result<Layout> {
    if xml.len() > limits.max_xml_bytes {
        return Err(Error::Limit {
            resource: "XML bytes",
            actual: xml.len(),
            maximum: limits.max_xml_bytes,
        });
    }
    let expected_kind = match selector {
        Selector::Main => RootKind::Main,
        Selector::Header(_) => RootKind::Header,
        Selector::Footer(_) => RootKind::Footer,
        Selector::Footnote(_) => RootKind::Footnotes,
        Selector::Endnote(_) => RootKind::Endnotes,
        Selector::Comment(_) => RootKind::Comments,
    };
    let mut reader = Reader::from_reader(xml);
    reader.config_mut().trim_text(false);
    reader.config_mut().check_end_names = true;
    let mut event_count = 0usize;
    let mut depth = 0usize;
    let mut root_seen = false;
    let mut root_closed = false;
    let mut root_start_end = 0usize;
    let mut root_end_start = 0usize;
    let mut root_kind = expected_kind;
    let mut namespace_attributes = Vec::new();
    let mut word_namespace = Vec::new();
    let mut has_w_namespace = false;
    let mut namespace_stack = Vec::new();
    let mut body_depth = None;
    let mut body_count = 0usize;
    let mut paragraph_stack = Vec::new();
    let mut paragraphs = Vec::new();
    let mut entry_count = 0usize;
    let mut selected_entry_count = 0usize;
    let mut selected_entry_depth = None;
    let mut selected_entry_outer_start = None;
    let mut selected_entry_inner_start = None;
    let mut selected_entry_inner_end = None;
    let mut selected_entry_namespace_attributes = None;
    let mut selected_entry_has_w_namespace = false;
    let empty_namespace_context = NamespaceContext::default();

    loop {
        let before = reader_position(&reader)?;
        let event = reader.read_event().map_err(|error| {
            Error::Document(crate::Error::InvalidFormat(format!(
                "invalid story XML: {error}",
            )))
        })?;
        event_count = event_count.checked_add(1).ok_or(Error::Limit {
            resource: "XML events",
            actual: usize::MAX,
            maximum: limits.max_events,
        })?;
        if event_count > limits.max_events {
            return Err(Error::Limit {
                resource: "XML events",
                actual: event_count,
                maximum: limits.max_events,
            });
        }
        if event_count & 63 == 0 {
            check_execution_context(execution)?;
        }
        let after = reader_position(&reader)?;
        match event {
            Event::Start(start) => {
                let event_start = before;
                let parent_depth = depth;
                let parent_context = namespace_stack.last().unwrap_or(&empty_namespace_context);
                let (context, element_namespace_attributes) = NamespaceContext::for_element(
                    parent_context,
                    &start,
                    limits.max_namespace_bindings,
                    limits.max_namespace_bytes,
                )?;
                if parent_depth == 0 {
                    if root_seen || root_closed {
                        return Err(unsupported_xml("multiple story roots"));
                    }
                    let info = inspect_root(
                        &start,
                        expected_kind,
                        &context,
                        element_namespace_attributes,
                    )?;
                    if expected_word_namespace
                        .is_some_and(|expected| info.word_namespace.as_slice() != expected)
                    {
                        return Err(unsupported_xml(
                            "story relationship and XML dialects differ",
                        ));
                    }
                    root_kind = info.kind;
                    namespace_attributes = info.namespace_attributes;
                    word_namespace = info.word_namespace;
                    has_w_namespace = info.has_w_namespace;
                    root_start_end = after;
                    root_seen = true;
                } else if !root_seen || root_closed {
                    return Err(unsupported_xml("story content occurs outside its root"));
                }
                reject_wrong_word_namespace(start.name(), &context, &word_namespace)?;
                if parent_depth == 1
                    && root_kind.entry_local_name().is_some_and(|local| {
                        is_word_element(start.name(), &context, &word_namespace, local)
                    })
                {
                    entry_count = entry_count.checked_add(1).ok_or(Error::Limit {
                        resource: "story entries",
                        actual: usize::MAX,
                        maximum: limits.max_entries,
                    })?;
                    if entry_count > limits.max_entries {
                        return Err(Error::Limit {
                            resource: "story entries",
                            actual: entry_count,
                            maximum: limits.max_entries,
                        });
                    }
                    if selector_matches_entry(&start, &context, &word_namespace, selector)? {
                        if selected_entry_count != 0 {
                            let (story, id) = secondary_identity(selector).ok_or_else(|| {
                                unsupported_xml("secondary story selector is incomplete")
                            })?;
                            return Err(Error::AmbiguousEntry { story, id });
                        }
                        selected_entry_count = 1;
                        selected_entry_depth =
                            Some(parent_depth.checked_add(1).ok_or(Error::Limit {
                                resource: "XML depth",
                                actual: usize::MAX,
                                maximum: limits.max_depth,
                            })?);
                        selected_entry_outer_start = Some(event_start);
                        selected_entry_inner_start = Some(after);
                        let (attributes, has_w_namespace) = context
                            .wrapper_namespace_attributes(&word_namespace, limits.max_xml_bytes)?;
                        selected_entry_namespace_attributes = Some(attributes);
                        selected_entry_has_w_namespace = has_w_namespace;
                    }
                }
                if root_kind == RootKind::Main
                    && parent_depth == 1
                    && is_word_element(start.name(), &context, &word_namespace, b"body")
                {
                    body_count = body_count.checked_add(1).ok_or(Error::Limit {
                        resource: "story body elements",
                        actual: usize::MAX,
                        maximum: 1,
                    })?;
                    if body_count > 1 {
                        return Err(unsupported_xml("main story contains multiple bodies"));
                    }
                    body_depth = Some(parent_depth.checked_add(1).ok_or(Error::Limit {
                        resource: "XML depth",
                        actual: usize::MAX,
                        maximum: limits.max_depth,
                    })?);
                }
                let direct_paragraph =
                    is_word_element(start.name(), &context, &word_namespace, b"p")
                        && match root_kind {
                            RootKind::Main => body_depth == Some(parent_depth),
                            RootKind::Header | RootKind::Footer => parent_depth == 1,
                            RootKind::Footnotes | RootKind::Endnotes | RootKind::Comments => {
                                selected_entry_depth == Some(parent_depth)
                            },
                        };
                if direct_paragraph {
                    ensure_paragraph_capacity(
                        paragraphs.len(),
                        paragraph_stack.len(),
                        limits.max_paragraphs,
                    )?;
                    reserve_one(&mut paragraphs, "story paragraph ranges")?;
                    let paragraph_depth = depth.checked_add(1).ok_or(Error::Limit {
                        resource: "XML depth",
                        actual: usize::MAX,
                        maximum: limits.max_depth,
                    })?;
                    reserve_one(&mut paragraph_stack, "story paragraph stack")?;
                    paragraph_stack.push((paragraph_depth, event_start));
                }
                depth = depth.checked_add(1).ok_or(Error::Limit {
                    resource: "XML depth",
                    actual: usize::MAX,
                    maximum: limits.max_depth,
                })?;
                if depth > limits.max_depth {
                    return Err(Error::Limit {
                        resource: "XML depth",
                        actual: depth,
                        maximum: limits.max_depth,
                    });
                }
                reserve_one(&mut namespace_stack, "story namespace stack")?;
                namespace_stack.push(context);
            },
            Event::Empty(start) => {
                let event_start = before;
                let parent_depth = depth;
                let logical_depth = parent_depth.checked_add(1).ok_or(Error::Limit {
                    resource: "XML depth",
                    actual: usize::MAX,
                    maximum: limits.max_depth,
                })?;
                if logical_depth > limits.max_depth {
                    return Err(Error::Limit {
                        resource: "XML depth",
                        actual: logical_depth,
                        maximum: limits.max_depth,
                    });
                }
                let parent_context = namespace_stack.last().unwrap_or(&empty_namespace_context);
                let (context, element_namespace_attributes) = NamespaceContext::for_element(
                    parent_context,
                    &start,
                    limits.max_namespace_bindings,
                    limits.max_namespace_bytes,
                )?;
                if parent_depth == 0 {
                    if root_seen || root_closed {
                        return Err(unsupported_xml("multiple story roots"));
                    }
                    let info = inspect_root(
                        &start,
                        expected_kind,
                        &context,
                        element_namespace_attributes,
                    )?;
                    if expected_word_namespace
                        .is_some_and(|expected| info.word_namespace.as_slice() != expected)
                    {
                        return Err(unsupported_xml(
                            "story relationship and XML dialects differ",
                        ));
                    }
                    root_kind = info.kind;
                    namespace_attributes = info.namespace_attributes;
                    word_namespace = info.word_namespace;
                    has_w_namespace = info.has_w_namespace;
                    root_start_end = after;
                    root_end_start = after;
                    root_seen = true;
                    root_closed = true;
                } else {
                    if !root_seen || root_closed {
                        return Err(unsupported_xml("story content occurs outside its root"));
                    }
                    reject_wrong_word_namespace(start.name(), &context, &word_namespace)?;
                    if parent_depth == 1
                        && root_kind.entry_local_name().is_some_and(|local| {
                            is_word_element(start.name(), &context, &word_namespace, local)
                        })
                    {
                        entry_count = entry_count.checked_add(1).ok_or(Error::Limit {
                            resource: "story entries",
                            actual: usize::MAX,
                            maximum: limits.max_entries,
                        })?;
                        if entry_count > limits.max_entries {
                            return Err(Error::Limit {
                                resource: "story entries",
                                actual: entry_count,
                                maximum: limits.max_entries,
                            });
                        }
                        if selector_matches_entry(&start, &context, &word_namespace, selector)? {
                            if selected_entry_count != 0 {
                                let (story, id) =
                                    secondary_identity(selector).ok_or_else(|| {
                                        unsupported_xml("secondary story selector is incomplete")
                                    })?;
                                return Err(Error::AmbiguousEntry { story, id });
                            }
                            let selected_bytes =
                                after.checked_sub(event_start).ok_or_else(|| {
                                    unsupported_xml("selected story entry range is inverted")
                                })?;
                            if selected_bytes > limits.max_entry_bytes {
                                return Err(Error::Limit {
                                    resource: "selected entry bytes",
                                    actual: selected_bytes,
                                    maximum: limits.max_entry_bytes,
                                });
                            }
                            selected_entry_count = 1;
                            selected_entry_inner_start = Some(after);
                            selected_entry_inner_end = Some(after);
                            let (attributes, has_w_namespace) = context
                                .wrapper_namespace_attributes(
                                    &word_namespace,
                                    limits.max_xml_bytes,
                                )?;
                            selected_entry_namespace_attributes = Some(attributes);
                            selected_entry_has_w_namespace = has_w_namespace;
                        }
                    }
                    if root_kind == RootKind::Main
                        && parent_depth == 1
                        && is_word_element(start.name(), &context, &word_namespace, b"body")
                    {
                        body_count = body_count.checked_add(1).ok_or(Error::Limit {
                            resource: "story body elements",
                            actual: usize::MAX,
                            maximum: 1,
                        })?;
                        if body_count > 1 {
                            return Err(unsupported_xml("main story contains multiple bodies"));
                        }
                    }
                    let direct_paragraph =
                        is_word_element(start.name(), &context, &word_namespace, b"p")
                            && match root_kind {
                                RootKind::Main => body_depth == Some(parent_depth),
                                RootKind::Header | RootKind::Footer => parent_depth == 1,
                                RootKind::Footnotes | RootKind::Endnotes | RootKind::Comments => {
                                    selected_entry_depth == Some(parent_depth)
                                },
                            };
                    if direct_paragraph {
                        ensure_paragraph_capacity(
                            paragraphs.len(),
                            paragraph_stack.len(),
                            limits.max_paragraphs,
                        )?;
                        reserve_one(&mut paragraphs, "story paragraph ranges")?;
                        paragraphs.push(Range {
                            start: event_start,
                            end: after,
                        });
                    }
                }
            },
            Event::End(end) => {
                let event_start = before;
                if depth == 0 {
                    return Err(unsupported_xml("story XML has an unmatched end element"));
                }
                let context = namespace_stack
                    .last()
                    .ok_or_else(|| unsupported_xml("story namespace stack is incomplete"))?;
                reject_wrong_word_namespace(end.name(), context, &word_namespace)?;
                if is_word_element(end.name(), context, &word_namespace, b"p") {
                    if let Some((paragraph_depth, start)) = paragraph_stack.pop() {
                        if paragraph_depth != depth {
                            return Err(unsupported_xml("nested or crossing story paragraphs"));
                        }
                        if paragraphs.len() >= limits.max_paragraphs {
                            return Err(Error::Limit {
                                resource: "paragraphs",
                                actual: paragraphs.len().saturating_add(1),
                                maximum: limits.max_paragraphs,
                            });
                        }
                        paragraphs.push(Range { start, end: after });
                    }
                }
                if selected_entry_depth == Some(depth) {
                    let outer_start = selected_entry_outer_start
                        .ok_or_else(|| unsupported_xml("selected story entry start is missing"))?;
                    let selected_bytes = after
                        .checked_sub(outer_start)
                        .ok_or_else(|| unsupported_xml("selected story entry range is inverted"))?;
                    if selected_bytes > limits.max_entry_bytes {
                        return Err(Error::Limit {
                            resource: "selected entry bytes",
                            actual: selected_bytes,
                            maximum: limits.max_entry_bytes,
                        });
                    }
                    selected_entry_inner_end = Some(event_start);
                    selected_entry_depth = None;
                }
                if depth == 1 {
                    if root_closed {
                        return Err(unsupported_xml("story root closes more than once"));
                    }
                    root_end_start = event_start;
                    root_closed = true;
                }
                if body_depth == Some(depth) {
                    body_depth = None;
                }
                namespace_stack.pop();
                depth -= 1;
            },
            Event::DocType(_) => {
                return Err(unsupported_xml("DOCTYPE declarations are not accepted"));
            },
            Event::PI(_) => {
                return Err(unsupported_xml("processing instructions are not accepted"));
            },
            Event::CData(_) => {
                return Err(unsupported_xml("CDATA sections are not accepted"));
            },
            Event::GeneralRef(reference) => {
                validate_story_general_reference(&reference)?;
            },
            Event::Decl(_) if root_seen => {
                return Err(unsupported_xml(
                    "XML declarations are only accepted before the story root",
                ));
            },
            Event::Text(text)
                if depth == 0 && !text.as_ref().iter().all(u8::is_ascii_whitespace) =>
            {
                return Err(unsupported_xml(
                    "non-whitespace text occurs outside story root",
                ));
            },
            Event::Eof => break,
            _ => {},
        }
    }

    if !root_seen || !root_closed || depth != 0 || !paragraph_stack.is_empty() {
        return Err(unsupported_xml("story XML root is incomplete"));
    }
    if root_kind == RootKind::Main && body_count != 1 {
        return Err(unsupported_xml("main story must contain exactly one body"));
    }
    if let Some((story, id)) = secondary_identity(selector) {
        if selected_entry_count == 0 {
            return Err(Error::MissingEntry { story, id });
        }
        root_start_end = selected_entry_inner_start
            .ok_or_else(|| unsupported_xml("selected story entry body start is missing"))?;
        root_end_start = selected_entry_inner_end
            .ok_or_else(|| unsupported_xml("selected story entry body end is missing"))?;
        namespace_attributes = selected_entry_namespace_attributes
            .ok_or_else(|| unsupported_xml("selected story namespace context is missing"))?;
        has_w_namespace = selected_entry_has_w_namespace;
    }
    check_execution_context(execution)?;
    paragraphs.sort_by_key(|range| range.start);
    Ok(Layout {
        paragraphs: paragraphs.into(),
        root_start_end,
        root_end_start,
        root_kind,
        namespace_attributes: namespace_attributes.into(),
        word_namespace: word_namespace.into(),
        has_w_namespace,
    })
}

struct RootInfo {
    kind: RootKind,
    namespace_attributes: Vec<u8>,
    word_namespace: Vec<u8>,
    has_w_namespace: bool,
}

fn inspect_root(
    start: &BytesStart<'_>,
    expected: RootKind,
    context: &NamespaceContext,
    namespace_attributes: Vec<u8>,
) -> Result<RootInfo> {
    let local = start.name().local_name();
    let expected_local = match expected {
        RootKind::Main => b"document".as_slice(),
        RootKind::Header => b"hdr".as_slice(),
        RootKind::Footer => b"ftr".as_slice(),
        RootKind::Footnotes => b"footnotes".as_slice(),
        RootKind::Endnotes => b"endnotes".as_slice(),
        RootKind::Comments => b"comments".as_slice(),
    };
    if local.as_ref() != expected_local {
        return Err(unsupported_xml("selected story root has the wrong element"));
    }
    let Some(word_namespace) = context.resolve_element(start.name()) else {
        return Err(unsupported_xml(
            "selected story root is not in a supported WordprocessingML namespace",
        ));
    };
    if word_namespace != TRANSITIONAL_WORD_NAMESPACE && word_namespace != STRICT_WORD_NAMESPACE {
        return Err(unsupported_xml(
            "selected story root is not in a supported WordprocessingML namespace",
        ));
    }
    if let Some(namespace) = context.resolve_prefix(b"w")
        && namespace != word_namespace
    {
        return Err(unsupported_xml(
            "the w prefix is bound to a namespace other than the story namespace",
        ));
    }
    let has_w_namespace = context.resolve_prefix(b"w") == Some(word_namespace);
    Ok(RootInfo {
        kind: expected,
        namespace_attributes,
        word_namespace: checked_vec_clone(word_namespace, "story namespace value")?,
        has_w_namespace,
    })
}

fn make_wrapped_document(
    xml: &[u8],
    layout: &Layout,
    maximum: usize,
) -> Result<(Vec<u8>, Envelope)> {
    if layout.root_kind == RootKind::Main {
        return Err(unsupported_xml("main stories do not require a wrapper"));
    }
    let inner = xml
        .get(layout.root_start_end..layout.root_end_start)
        .ok_or_else(|| {
            Error::Document(crate::Error::InvalidFormat(
                "story root body is outside source XML".into(),
            ))
        })?;
    let extra_namespace = if layout.has_w_namespace {
        0
    } else {
        b" xmlns:w=\""
            .len()
            .checked_add(layout.word_namespace.len())
            .and_then(|value| value.checked_add(1))
            .ok_or(Error::Limit {
                resource: "wrapped story XML bytes",
                actual: usize::MAX,
                maximum,
            })?
    };
    let prefix_len = b"<w:document"
        .len()
        .checked_add(layout.namespace_attributes.len())
        .and_then(|value| value.checked_add(extra_namespace))
        .and_then(|value| value.checked_add(b"><w:body>".len()))
        .ok_or(Error::Limit {
            resource: "wrapped story XML bytes",
            actual: usize::MAX,
            maximum,
        })?;
    let suffix = b"</w:body></w:document>";
    let total = prefix_len
        .checked_add(inner.len())
        .and_then(|value| value.checked_add(suffix.len()))
        .ok_or(Error::Limit {
            resource: "wrapped story XML bytes",
            actual: usize::MAX,
            maximum,
        })?;
    if total > maximum {
        return Err(Error::Limit {
            resource: "wrapped story XML bytes",
            actual: total,
            maximum,
        });
    }
    let mut wrapped = Vec::new();
    wrapped
        .try_reserve_exact(total)
        .map_err(|source| Error::Allocation {
            resource: "wrapped story XML",
            source,
        })?;
    wrapped.extend_from_slice(b"<w:document");
    wrapped.extend_from_slice(layout.namespace_attributes.as_slice());
    if !layout.has_w_namespace {
        wrapped.extend_from_slice(b" xmlns:w=\"");
        wrapped.extend_from_slice(layout.word_namespace.as_ref());
        wrapped.push(b'\"');
    }
    wrapped.extend_from_slice(b"><w:body>");
    let wrapped_prefix_len = wrapped.len();
    wrapped.extend_from_slice(inner);
    wrapped.extend_from_slice(suffix);
    Ok((
        wrapped,
        Envelope {
            inner_start: layout.root_start_end,
            inner_end: layout.root_end_start,
            wrapped_prefix_len,
            wrapped_suffix_len: suffix.len(),
        },
    ))
}

fn is_namespace_declaration(key: QName<'_>) -> bool {
    key.as_ref() == b"xmlns"
        || key
            .prefix()
            .as_ref()
            .is_some_and(|prefix| prefix.as_ref() == b"xmlns")
}

fn namespace_declared_prefix<'a>(key: QName<'a>) -> Option<&'a [u8]> {
    key.into_inner().strip_prefix(b"xmlns:")
}

fn is_mce_attribute(key: QName<'_>, value: &[u8]) -> bool {
    value == MCE_NAMESPACE
        || key
            .prefix()
            .as_ref()
            .is_some_and(|prefix| prefix.as_ref() == b"mc")
        || key.as_ref() == b"mc:Ignorable"
}

fn is_word_element(
    name: QName<'_>,
    context: &NamespaceContext,
    word_namespace: &[u8],
    local: &[u8],
) -> bool {
    name.local_name().as_ref() == local && context.resolve_element(name) == Some(word_namespace)
}

fn reject_wrong_word_namespace(
    name: QName<'_>,
    context: &NamespaceContext,
    word_namespace: &[u8],
) -> Result<()> {
    let local = name.local_name();
    let namespace = context.resolve_element(name);
    if name.prefix().is_some() && namespace.is_none() {
        return Err(unsupported_xml("undeclared element namespace prefix"));
    }
    if namespace == Some(MCE_NAMESPACE) {
        return Err(unsupported_xml(
            "markup-compatibility elements are not accepted in a source-bound story",
        ));
    }
    if matches!(
        local.as_ref(),
        b"document"
            | b"body"
            | b"hdr"
            | b"ftr"
            | b"footnotes"
            | b"footnote"
            | b"endnotes"
            | b"endnote"
            | b"comments"
            | b"comment"
            | b"p"
    ) && namespace != Some(word_namespace)
    {
        return Err(unsupported_xml(
            "WordprocessingML element has an incorrect namespace binding",
        ));
    }
    Ok(())
}

fn unsupported_xml(reason: &'static str) -> Error {
    Error::Document(crate::Error::UnsafeEdit {
        format: "DOCX",
        operation: "source-backed story text",
        reason,
    })
}

fn check_execution_context(context: Option<&ExecutionContext>) -> Result<()> {
    let Some(context) = context else {
        return Ok(());
    };
    context.check().map_err(|error| {
        Error::Document(crate::Error::Opc(match error {
            ExecutionError::Cancelled => litchi_opc::OpcError::Cancelled,
            other => litchi_opc::OpcError::Execution(other),
        }))
    })
}

fn reserve_one<T>(items: &mut Vec<T>, resource: &'static str) -> Result<()> {
    items
        .try_reserve(1)
        .map_err(|source| Error::Allocation { resource, source })
}

fn ensure_paragraph_capacity(committed: usize, pending: usize, maximum: usize) -> Result<()> {
    let actual = committed.checked_add(pending).ok_or(Error::Limit {
        resource: "paragraphs",
        actual: usize::MAX,
        maximum,
    })?;
    if actual >= maximum {
        return Err(Error::Limit {
            resource: "paragraphs",
            actual: actual.saturating_add(1),
            maximum,
        });
    }
    Ok(())
}

fn checked_vec_clone(bytes: &[u8], resource: &'static str) -> Result<Vec<u8>> {
    let mut copy = Vec::new();
    copy.try_reserve_exact(bytes.len())
        .map_err(|source| Error::Allocation { resource, source })?;
    copy.extend_from_slice(bytes);
    Ok(copy)
}

fn checked_string_clone(value: &str, resource: &'static str) -> Result<String> {
    let mut copy = String::new();
    copy.try_reserve_exact(value.len())
        .map_err(|source| Error::Allocation { resource, source })?;
    copy.push_str(value);
    Ok(copy)
}

fn checked_story_output_clone(
    bytes: &[u8],
    limits: &Limits,
    resource: &'static str,
) -> Result<Vec<u8>> {
    if bytes.len() > limits.max_output_bytes {
        return Err(Error::Limit {
            resource,
            actual: bytes.len(),
            maximum: limits.max_output_bytes,
        });
    }
    checked_vec_clone(bytes, resource)
}

fn digest(bytes: &[u8]) -> [u8; 32] {
    Sha256::digest(bytes).into()
}

fn same_part_name(left: &PackURI, right: &PackURI) -> bool {
    left.as_str().eq_ignore_ascii_case(right.as_str())
}

fn error_to_io(error: Error) -> io::Error {
    io::Error::new(io::ErrorKind::Interrupted, error.to_string())
}

struct CancellationWriter<'a, W> {
    writer: &'a mut W,
    execution: Option<&'a ExecutionContext>,
}

impl<'a, W: Write> Write for CancellationWriter<'a, W> {
    fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
        check_execution_context(self.execution).map_err(error_to_io)?;
        let accepted = self.writer.write(bytes)?;
        if accepted > bytes.len() {
            return Err(io::Error::new(
                io::ErrorKind::InvalidData,
                "writer accepted more bytes than supplied",
            ));
        }
        check_execution_context(self.execution).map_err(error_to_io)?;
        Ok(accepted)
    }

    fn flush(&mut self) -> io::Result<()> {
        check_execution_context(self.execution).map_err(error_to_io)?;
        self.writer.flush()
    }
}

#[derive(Debug)]
struct StoryOutputLimit {
    attempted: usize,
    maximum: usize,
}

impl std::fmt::Display for StoryOutputLimit {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        write!(
            formatter,
            "story text output limit exceeded (attempted {}, maximum {})",
            self.attempted, self.maximum
        )
    }
}

impl std::error::Error for StoryOutputLimit {}

struct BoundedStoryWriter<'a, W: Write + ?Sized> {
    inner: &'a mut W,
    accepted: usize,
    maximum: usize,
}

impl<'a, W: Write + ?Sized> BoundedStoryWriter<'a, W> {
    fn new(inner: &'a mut W, maximum: usize) -> Self {
        Self {
            inner,
            accepted: 0,
            maximum,
        }
    }
}

impl<W: Write + ?Sized> Write for BoundedStoryWriter<'_, W> {
    fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
        let remaining = self.maximum.checked_sub(self.accepted).ok_or_else(|| {
            io::Error::new(
                io::ErrorKind::InvalidData,
                "story output byte count overflow",
            )
        })?;
        if bytes.len() > remaining {
            return Err(io::Error::new(
                io::ErrorKind::WriteZero,
                StoryOutputLimit {
                    attempted: self.accepted.saturating_add(bytes.len()),
                    maximum: self.maximum,
                },
            ));
        }
        let accepted = self.inner.write(bytes)?;
        if accepted > bytes.len() {
            return Err(io::Error::new(
                io::ErrorKind::InvalidData,
                "writer accepted more bytes than supplied",
            ));
        }
        self.accepted = self.accepted.checked_add(accepted).ok_or_else(|| {
            io::Error::new(
                io::ErrorKind::InvalidData,
                "story output byte count overflow",
            )
        })?;
        Ok(accepted)
    }

    fn flush(&mut self) -> io::Result<()> {
        self.inner.flush()
    }
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

    fn finish(self) -> SourceArtifactFingerprint {
        SourceArtifactFingerprint::from_sha256(self.hasher.finalize().into())
    }
}

impl<W: Write> Write for FingerprintingWriter<W> {
    fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
        let accepted = self.inner.write(bytes)?;
        if accepted > bytes.len() {
            return Err(io::Error::new(
                io::ErrorKind::InvalidData,
                "story publication sink reported more bytes than supplied",
            ));
        }
        self.hasher.update(&bytes[..accepted]);
        Ok(accepted)
    }

    fn flush(&mut self) -> io::Result<()> {
        self.inner.flush()
    }
}

const STRICT_HEADER: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships/header";
const STRICT_FOOTER: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships/footer";
const STRICT_FOOTNOTES: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships/footnotes";
const STRICT_ENDNOTES: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships/endnotes";
const TRANSITIONAL_GLOSSARY: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/glossaryDocument";
const STRICT_GLOSSARY: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships/glossaryDocument";
