//! Source-backed text snapshots for one selected Word story or one bounded
//! glossary-entry batch.
//!
//! This module deliberately keeps the selection closure to one main, header,
//! footer, footnote, endnote, comment, or glossary-entry story. Relationship metadata is
//! resolved without reading payloads; only the owning XML member is
//! materialized. Edits reuse the main document transaction engine through a
//! private `w:document`/`w:body` adapter. For collection parts, only the exact
//! selected entry body range is spliced back, preserving the root and all
//! sibling entries byte-for-byte.

use std::cell::RefCell;
use std::collections::TryReserveError;
use std::io::{self, Write};
use std::sync::Arc;

use litchi_core::xml::ReaderOrigin;
use litchi_core::{
    ExecutionContext, ExecutionError, Position, SourceVersion, TextObjectKind, TextOutputError,
    TextOutputOptions, TextOutputReport,
};
use litchi_ooxml_common::xml::decode_xml_reference;
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{
    PackURI, PartData, PartView, SourceArtifactFingerprint, SourceBackedPackage, SourceLineage,
};
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::QName;
use quick_xml::reader::Reader;
use sha2::{Digest as _, Sha256};
use thiserror::Error as ThisError;

use crate::document::raw::Tree as DocumentEdit;
use crate::document::{
    Error as TransactionError, MAX_DOCUMENT_XML_BYTES, Snapshot as DocumentSnapshot,
    validate_paragraph_text,
};

use super::Package;
use litchi_ooxml_common::xml::attributes::BytesStartExt as _;

const TRANSITIONAL_WORD_NAMESPACE: &[u8] =
    b"http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const STRICT_WORD_NAMESPACE: &[u8] = b"http://purl.oclc.org/ooxml/wordprocessingml/main";
const TRANSITIONAL_RELATIONSHIP_NAMESPACE: &[u8] =
    b"http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const STRICT_RELATIONSHIP_NAMESPACE: &[u8] =
    b"http://purl.oclc.org/ooxml/officeDocument/relationships";
const VML_NAMESPACE: &[u8] = b"urn:schemas-microsoft-com:vml";
const MCE_NAMESPACE: &[u8] = b"http://schemas.openxmlformats.org/markup-compatibility/2006";
const XML_NAMESPACE: &[u8] = b"http://www.w3.org/XML/1998/namespace";
const DEFAULT_MAX_NAMESPACE_BINDINGS: usize = 256;
const DEFAULT_MAX_NAMESPACE_BYTES: usize = 4 * 1024 * 1024;
const MAX_STORY_REFERENCE_BYTES: usize = 64 * 1024;
const MAX_GLOSSARY_SELECTOR_NAME_BYTES: usize = 1024 * 1024;
const MAX_GLOSSARY_SELECTOR_ID_BYTES: usize = 256;

/// An owned semantic selector for one existing glossary document part.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub enum GlossarySelector {
    /// Select by the Unicode-case-folded `w:name` value.
    ByName(String),
    /// Select by the semantic `w:guid` value.
    ById(crate::glossary::Id),
    /// Select the entry whose `w:name` and `w:guid` both match.
    ByNameAndId {
        /// Requested building-block name.
        name: String,
        /// Requested building-block GUID.
        id: crate::glossary::Id,
    },
    /// Select by source-order index as a snapshot-local fallback.
    ByIndex(usize),
}

/// The compact diagnostic kind of a glossary selector.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum GlossarySelectorKind {
    /// A Unicode-case-folded name selector.
    ByName,
    /// A semantic GUID selector.
    ById,
    /// A selector requiring both name and GUID equality.
    ByNameAndId,
    /// A source-order index selector.
    ByIndex,
}

impl GlossarySelector {
    /// Return the compact selector kind used by typed lookup diagnostics.
    #[must_use]
    pub const fn kind(&self) -> GlossarySelectorKind {
        match self {
            Self::ByName(_) => GlossarySelectorKind::ByName,
            Self::ById(_) => GlossarySelectorKind::ById,
            Self::ByNameAndId { .. } => GlossarySelectorKind::ByNameAndId,
            Self::ByIndex(_) => GlossarySelectorKind::ByIndex,
        }
    }

    /// Construct a name-based glossary selector.
    #[must_use]
    pub fn by_name(name: impl Into<String>) -> Self {
        Self::ByName(name.into())
    }

    /// Construct a GUID-based glossary selector.
    #[must_use]
    pub const fn by_id(id: crate::glossary::Id) -> Self {
        Self::ById(id)
    }

    /// Construct a selector requiring both name and GUID equality.
    #[must_use]
    pub fn by_name_and_id(name: impl Into<String>, id: crate::glossary::Id) -> Self {
        Self::ByNameAndId {
            name: name.into(),
            id,
        }
    }

    /// Construct a checked source-order selector.
    #[must_use]
    pub const fn by_index(index: usize) -> Self {
        Self::ByIndex(index)
    }
}

/// An owned, non-empty selection of existing glossary entries.
///
/// The selectors are resolved against one glossary Part. During capture they
/// are canonicalized to the source order of the selected entries; duplicate,
/// missing, and ambiguous semantic matches are refused before editing. The
/// selector count is checked against [`Limits::max_entries`] when captured.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct GlossaryBatchSelector {
    selectors: Vec<GlossarySelector>,
}

impl GlossaryBatchSelector {
    /// Construct an exactly-two glossary-entry convenience selector.
    #[must_use]
    pub fn new(first: GlossarySelector, second: GlossarySelector) -> Self {
        Self {
            selectors: vec![first, second],
        }
    }

    /// Construct a checked non-empty glossary-entry selector set.
    pub fn try_new(selectors: Vec<GlossarySelector>) -> Result<Self> {
        if selectors.is_empty() {
            return Err(Error::Limit {
                kind: LimitKind::GlossaryBatchSelectors,
                actual: 0,
                maximum: 1,
            });
        }
        Ok(Self { selectors })
    }

    /// Borrow the non-empty semantic selectors in their requested order.
    ///
    /// A captured [`GlossaryBatchSnapshot`] exposes its canonical source
    /// order through [`GlossaryBatchSnapshot::selector`] and
    /// [`GlossaryBatchSnapshot::entry_selector`].
    #[must_use]
    pub fn selectors(&self) -> &[GlossarySelector] {
        &self.selectors
    }
}

/// A semantic selector for one source-backed Word story.
///
/// Header and footer indices are assigned after sorting the relationship
/// targets by canonical absolute part name. The selector never exposes an OPC
/// URI or relationship identifier.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
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
    /// One existing glossary document part selected by semantic identity.
    Glossary(GlossarySelector),
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

    /// Construct a glossary-entry selector.
    #[must_use]
    pub fn glossary(selector: GlossarySelector) -> Self {
        Self::Glossary(selector)
    }

    /// Construct a source-order glossary-entry selector.
    #[must_use]
    pub const fn glossary_index(index: usize) -> Self {
        Self::Glossary(GlossarySelector::ByIndex(index))
    }

    /// Construct a name-based glossary-entry selector.
    #[must_use]
    pub fn glossary_by_name(name: impl Into<String>) -> Self {
        Self::Glossary(GlossarySelector::by_name(name))
    }

    /// Construct a GUID-based glossary-entry selector.
    #[must_use]
    pub const fn glossary_by_id(id: crate::glossary::Id) -> Self {
        Self::Glossary(GlossarySelector::ById(id))
    }

    /// Construct a glossary-entry selector requiring a name and GUID.
    #[must_use]
    pub fn glossary_by_name_and_id(name: impl Into<String>, id: crate::glossary::Id) -> Self {
        Self::Glossary(GlossarySelector::by_name_and_id(name, id))
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
    DocumentEdit(#[from] TransactionError),
    /// A bounded story operation exceeded one finite resource ceiling.
    #[error("source-backed story text {kind} limit exceeded: {actual} > {maximum}")]
    Limit {
        /// The exceeded budget (record 0355a).
        kind: LimitKind,
        /// Observed value.
        actual: usize,
        /// Maximum accepted value.
        maximum: usize,
    },
    /// A patch was applied to a source with different selected-part bytes.
    #[error("source-backed story text patch source is stale")]
    PatchConflict,
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
    /// The requested glossary entry is absent.
    #[error("source-backed glossary selector {selector:?} is absent")]
    MissingGlossaryEntry {
        /// Compact selector kind that matched no source entry.
        selector: GlossarySelectorKind,
    },
    /// More than one glossary entry matches the semantic selector.
    #[error("source-backed glossary selector {selector:?} is ambiguous")]
    AmbiguousGlossaryEntry {
        /// Compact selector kind that matched multiple source entries.
        selector: GlossarySelectorKind,
    },
    /// A batch edit attempted to replace one direct paragraph more than once.
    #[error(
        "source-backed glossary batch paragraph {paragraph_index} in entry {entry_index} is repeated"
    )]
    RepeatedParagraphEdit {
        /// Canonical source-order entry index.
        entry_index: usize,
        /// Direct paragraph index within that entry.
        paragraph_index: usize,
    },
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
            (LimitKind::XmlBytes, max_xml_bytes),
            (LimitKind::Paragraphs, max_paragraphs),
            (LimitKind::XmlEvents, max_events),
            (LimitKind::XmlDepth, max_depth),
            (LimitKind::OutputBytes, max_output_bytes),
            (LimitKind::Replacements, max_replacements),
            (LimitKind::ReplacementTextBytes, max_replacement_text_bytes),
        ];
        if let Some((resource, _)) = values.iter().find(|(_, value)| *value == 0) {
            return Err(Error::Limit {
                kind: *resource,
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
            (LimitKind::Entries, max_entries),
            (LimitKind::SelectedEntryBytes, max_entry_bytes),
            (LimitKind::NamespaceBindings, max_namespace_bindings),
        ];
        if let Some((resource, _)) = values.iter().find(|(_, value)| *value == 0) {
            return Err(Error::Limit {
                kind: *resource,
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
                kind: LimitKind::NamespaceBytes,
                actual: 0,
                maximum: 1,
            });
        }
        self.max_namespace_bytes = max_namespace_bytes;
        Ok(self)
    }

    fn intersect(self, other: Self) -> Self {
        Self {
            max_xml_bytes: self.max_xml_bytes.min(other.max_xml_bytes),
            max_paragraphs: self.max_paragraphs.min(other.max_paragraphs),
            max_events: self.max_events.min(other.max_events),
            max_depth: self.max_depth.min(other.max_depth),
            max_output_bytes: self.max_output_bytes.min(other.max_output_bytes),
            max_replacements: self.max_replacements.min(other.max_replacements),
            max_replacement_text_bytes: self
                .max_replacement_text_bytes
                .min(other.max_replacement_text_bytes),
            max_entries: self.max_entries.min(other.max_entries),
            max_entry_bytes: self.max_entry_bytes.min(other.max_entry_bytes),
            max_namespace_bindings: self
                .max_namespace_bindings
                .min(other.max_namespace_bindings),
            max_namespace_bytes: self.max_namespace_bytes.min(other.max_namespace_bytes),
        }
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
    Glossary,
}

impl RootKind {
    const fn entry_local_name(self) -> Option<&'static [u8]> {
        match self {
            Self::Footnotes => Some(b"footnote"),
            Self::Endnotes => Some(b"endnote"),
            Self::Comments => Some(b"comment"),
            Self::Main | Self::Header | Self::Footer | Self::Glossary => None,
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
    /// A validated view into the transaction XML.  Main-story projections
    /// already own the complete document buffer; retaining its range avoids
    /// copying that buffer merely to construct the raw story snapshot.
    OwnedRange {
        bytes: Arc<Vec<u8>>,
        range: Range,
    },
    Managed(PartData),
}

impl Payload {
    fn as_bytes(&self) -> &[u8] {
        match self {
            Self::Owned(bytes) => bytes.as_slice(),
            Self::OwnedRange { bytes, range } => &bytes[range.start..range.end],
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
            selector: self.selector.clone(),
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
    pub fn selector(&self) -> Selector {
        self.selector.clone()
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

    /// Return the policy under which this snapshot was admitted.
    #[must_use]
    pub const fn limits_value(&self) -> Limits {
        self.limits
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
            LimitKind::ParagraphText,
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
                LimitKind::TextOutput,
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
                        kind: LimitKind::TextOutput,
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
                LimitKind::ParagraphText,
            )
            .map_err(|error| writer.document_error(error))?;
            writer.write_object::<Error>(TextObjectKind::Paragraph, &text)?;
        }
        if let Err(error) = self.check_execution() {
            return Err(writer.document_error(error));
        }
        Ok(writer.finish())
    }

    /// Open the selected story as a bounded mutable raw codec model.
    ///
    /// The supplied policy is intersected with the policy retained by this
    /// admitted snapshot. Raw models publish [`Patch`] values rather than
    /// package owners; source-backed package edits use the private consuming
    /// transaction engine instead.
    pub fn raw_tree(&self, limits: Limits) -> Result<raw::Tree> {
        let base = self.with_raw_limits(limits)?;
        Ok(raw::Tree::from_engine(base.edit_engine()?))
    }

    pub(crate) fn edit_engine(&self) -> Result<Edit> {
        self.check_execution()?;
        if self.payload.is_managed() {
            return Err(Error::Document(crate::Error::refused(
                litchi_core::edit::RefusalKind::UnsupportedOperation,
                "source-backed story text edit",
                "managed source-backed story edits require an owned edit snapshot; use an unmanaged compatibility constructor",
            )));
        }
        let mut base = self.clone();
        let transaction = if let Some(transaction) = self.transaction.clone() {
            transaction
        } else if !matches!(&self.selector, Selector::Main) {
            let effective_output_limit = self.limits.max_output_bytes.min(MAX_DOCUMENT_XML_BYTES);
            let (document_xml, envelope) =
                make_wrapped_document(self.raw_bytes(), &self.layout, effective_output_limit)?;
            base.envelope = Some(envelope);
            DocumentSnapshot::from_xml(document_xml)?
        } else {
            let source_bytes = self.raw_bytes();
            if source_bytes.len() > self.limits.max_output_bytes {
                return Err(Error::Limit {
                    kind: LimitKind::TransactionXml,
                    actual: source_bytes.len(),
                    maximum: self.limits.max_output_bytes,
                });
            }
            let document_xml = checked_vec_clone(source_bytes, "story transaction XML")?;
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

    fn with_raw_limits(&self, requested: Limits) -> Result<Self> {
        let limits = self.limits.intersect(requested);
        if limits == self.limits {
            return Ok(self.clone());
        }
        let layout = scan_selected_story(
            self.raw_bytes(),
            &self.selector,
            Some(self.layout.word_namespace.as_ref()),
            limits,
            self.execution.as_ref(),
        )?;
        let mut base = self.clone();
        base.layout = layout;
        base.limits = limits;
        base.transaction = None;
        base.envelope = None;
        Ok(base)
    }

    fn with_transaction_snapshot(&self, transaction: DocumentSnapshot) -> Result<Self> {
        let (payload, envelope) = if let Some(envelope) = self.envelope {
            let wrapped = transaction.shared_xml().map_err(Error::DocumentEdit)?;
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
                    kind: LimitKind::OutputBytes,
                    actual: usize::MAX,
                    maximum: self.limits.max_output_bytes,
                })?;
            if total > self.limits.max_output_bytes {
                return Err(Error::Limit {
                    kind: LimitKind::OutputBytes,
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
            (Payload::Owned(Arc::new(raw)), Some(envelope))
        } else {
            let bytes = transaction.shared_xml().map_err(Error::DocumentEdit)?;
            if bytes.len() > self.limits.max_output_bytes {
                return Err(Error::Limit {
                    kind: LimitKind::TransactionXml,
                    actual: bytes.len(),
                    maximum: self.limits.max_output_bytes,
                });
            }
            let range = Range {
                start: 0,
                end: bytes.len(),
            };
            (Payload::OwnedRange { bytes, range }, None)
        };
        let layout = scan_selected_story(
            payload.as_bytes(),
            &self.selector,
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
            selector: self.selector.clone(),
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
}

/// Private transaction engine for the raw codec model and source-backed owner
/// family. The public raw model wraps this engine with a bounded projection
/// cache; the owner family consumes it directly.
pub(crate) struct Edit {
    base: Snapshot,
    projected: DocumentEdit,
    replacements: usize,
    replacement_text_bytes: usize,
}

impl Edit {
    /// Borrow the immutable source snapshot.
    #[must_use]
    pub(crate) const fn source(&self) -> &Snapshot {
        &self.base
    }

    /// Borrow the current projected selected-story snapshot.
    pub(crate) fn projected(&self) -> Result<Snapshot> {
        self.base
            .with_transaction_snapshot(self.projected.projected().clone())
    }

    /// Replace all text in one direct selected-story paragraph.
    pub(crate) fn replace_paragraph_text(
        &mut self,
        position: Position,
        authored_text: impl AsRef<str>,
    ) -> Result<&mut Self> {
        self.base.check_execution()?;
        let next_replacements = self.replacements.checked_add(1).ok_or(Error::Limit {
            kind: LimitKind::Replacements,
            actual: usize::MAX,
            maximum: self.base.limits.max_replacements,
        })?;
        if next_replacements > self.base.limits.max_replacements {
            return Err(Error::Limit {
                kind: LimitKind::Replacements,
                actual: next_replacements,
                maximum: self.base.limits.max_replacements,
            });
        }
        let authored_text = authored_text.as_ref();
        let next_bytes = self
            .replacement_text_bytes
            .checked_add(authored_text.len())
            .ok_or(Error::Limit {
                kind: LimitKind::ReplacementTextBytes,
                actual: usize::MAX,
                maximum: self.base.limits.max_replacement_text_bytes,
            })?;
        if next_bytes > self.base.limits.max_replacement_text_bytes {
            return Err(Error::Limit {
                kind: LimitKind::ReplacementTextBytes,
                actual: next_bytes,
                maximum: self.base.limits.max_replacement_text_bytes,
            });
        }
        validate_paragraph_text(position, authored_text).map_err(Error::DocumentEdit)?;
        let text = checked_string_clone(authored_text, "story replacement text")?;
        let mut candidate = self.projected.try_clone().map_err(Error::DocumentEdit)?;
        candidate
            .replace_paragraph_text(position, text)
            .map_err(Error::DocumentEdit)?;
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

    fn target(&self) -> Result<Snapshot> {
        self.base.check_execution()?;
        // The staged transaction has already validated every replacement and
        // `with_transaction_snapshot` below performs the selected-story
        // admission again.  Reusing its projected snapshot keeps raw output
        // source-preserving and avoids cloning the complete transaction just
        // to call its consuming owner commit.  The owner family still uses
        // `finish`, which consumes the transaction at publication time.
        let transaction_snapshot = self.projected.projected().clone();
        self.base.check_execution()?;
        let target = self.base.with_transaction_snapshot(transaction_snapshot)?;
        Ok(target)
    }
}

impl Edit {
    pub(crate) fn finish(self) -> Result<Patch> {
        self.base.check_execution()?;
        let transaction_snapshot = if self.base.envelope.is_some() {
            self.projected.projected().clone()
        } else {
            self.projected.commit()?.snapshot().clone()
        };
        self.base.check_execution()?;
        let target = self.base.with_transaction_snapshot(transaction_snapshot)?;
        Ok(Patch {
            before: self.base,
            after: target,
        })
    }
}

/// A reversible selected-story patch carrying the selected-part identity and
/// exact before and after bytes. Owner publication adds its complete-artifact
/// and source-freshness proof; raw replay matches the selected content.
#[derive(Clone)]
pub struct Patch {
    before: Snapshot,
    after: Snapshot,
}

impl Patch {
    /// Borrow the before snapshot represented by this patch.
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

    /// Apply to a selected-story snapshot with the patch's exact selector,
    /// part and before bytes, from any source (record 0355a: a package
    /// replay reads the snapshot it applies to).
    pub(crate) fn apply_content(&self, source: &Snapshot) -> Result<Snapshot> {
        self.check_content(source)?;
        Ok(if self.is_noop() {
            source.clone()
        } else {
            self.after.clone()
        })
    }

    fn check_content(&self, source: &Snapshot) -> Result<()> {
        if source.selector != self.before.selector
            || !same_part_name(&source.partname, &self.before.partname)
        {
            return Err(Error::TopologyConflict);
        }
        if !same_content(
            source.part_fingerprint,
            source.raw_bytes(),
            self.before.raw_bytes(),
            self.before.part_fingerprint,
        ) {
            return Err(Error::PatchConflict);
        }
        Ok(())
    }

    /// Replay this raw patch on an independently opened matching snapshot.
    ///
    /// Raw codec replay is scoped to the selected XML member. The owning
    /// package family performs its separate complete-artifact and source
    /// freshness checks before publication. The returned snapshot is
    /// admitted under the recipient policy and carries the recipient source
    /// identity, so a subsequent raw edit cannot inherit the author's fence.
    pub fn apply(&self, source: &Snapshot) -> Result<Snapshot> {
        source.check_execution()?;
        self.check_content(source)?;
        self.rebased_target(source)
    }

    fn rebased_target(&self, source: &Snapshot) -> Result<Snapshot> {
        if self.is_noop() {
            return Ok(source.clone());
        }
        let bytes = self.after.raw_bytes();
        if bytes.len() > source.limits.max_output_bytes {
            return Err(Error::Limit {
                kind: LimitKind::OutputBytes,
                actual: bytes.len(),
                maximum: source.limits.max_output_bytes,
            });
        }
        let layout = scan_selected_story(
            bytes,
            &source.selector,
            Some(source.layout.word_namespace.as_ref()),
            source.limits,
            source.execution.as_ref(),
        )?;
        let mut target = self.after.clone();
        target.selector = source.selector.clone();
        target.partname = source.partname.clone();
        target.layout = layout;
        target.transaction = None;
        target.envelope = None;
        target.source_version = source.source_version;
        target.source_lineage = source.source_lineage.clone();
        target.part_fingerprint = digest(bytes);
        target.artifact_fingerprint = source.artifact_fingerprint;
        target.limits = source.limits;
        target.execution = source.execution.clone();
        Ok(target)
    }

    pub(crate) fn selector(&self) -> Selector {
        self.before.selector.clone()
    }
}

/// Raw mutable codec models for selected-story and glossary text records.
pub mod raw {
    use super::{
        Edit, GlossaryBatchEdit, GlossaryBatchPatch, GlossaryBatchSnapshot, Patch, Position,
        RefCell, Result, Snapshot,
    };

    /// A mutable raw selected-story text model.
    ///
    /// The model keeps only the private transaction engine and one cached
    /// projection. Repeated output calls borrow the cached snapshot; they do
    /// not clone the complete transaction. The consuming engine remains
    /// private to the source-backed owner family.
    pub struct Tree {
        inner: Edit,
        projected: RefCell<Option<Snapshot>>,
    }

    impl Tree {
        pub(super) fn from_engine(inner: Edit) -> Self {
            Self {
                inner,
                projected: RefCell::new(None),
            }
        }

        /// Borrow the immutable source snapshot.
        #[must_use]
        pub const fn source(&self) -> &Snapshot {
            self.inner.source()
        }

        /// Return the current bounded projection.
        pub fn projected(&self) -> Result<Snapshot> {
            if let Some(projected) = self.projected.borrow().as_ref() {
                return Ok(projected.clone());
            }
            let projected = self.inner.target()?;
            *self.projected.borrow_mut() = Some(projected.clone());
            Ok(projected)
        }

        /// Replace all text in one direct paragraph.
        pub fn replace_paragraph_text(
            &mut self,
            position: Position,
            authored_text: impl AsRef<str>,
        ) -> Result<&mut Self> {
            self.inner.replace_paragraph_text(position, authored_text)?;
            self.projected.borrow_mut().take();
            Ok(self)
        }

        /// Produce the source-bound raw patch without publishing a package.
        pub fn patch(&self) -> Result<Patch> {
            let target = self.projected()?;
            Ok(Patch {
                before: self.inner.source().clone(),
                after: target,
            })
        }

        /// Replay a selected-story patch against this model's source.
        ///
        /// The returned snapshot is a raw projection; no package is
        /// published and no source lineage is required.
        pub fn apply(&self, patch: &Patch) -> Result<Snapshot> {
            patch.apply(self.source())
        }
    }

    /// A mutable raw glossary-entry batch model.
    pub struct GlossaryTree {
        inner: GlossaryBatchEdit,
        projected: RefCell<Option<GlossaryBatchSnapshot>>,
    }

    impl GlossaryTree {
        pub(super) fn from_engine(inner: GlossaryBatchEdit) -> Self {
            Self {
                inner,
                projected: RefCell::new(None),
            }
        }

        /// Borrow the immutable source snapshot.
        #[must_use]
        pub const fn source(&self) -> &GlossaryBatchSnapshot {
            self.inner.source()
        }

        /// Return the current bounded projection.
        pub fn projected(&self) -> Result<GlossaryBatchSnapshot> {
            if let Some(projected) = self.projected.borrow().as_ref() {
                return Ok(projected.clone());
            }
            let projected = self.inner.projected()?;
            *self.projected.borrow_mut() = Some(projected.clone());
            Ok(projected)
        }

        /// Replace all text in one direct paragraph of one selected entry.
        pub fn replace_paragraph_text(
            &mut self,
            entry_index: usize,
            position: Position,
            authored_text: impl AsRef<str>,
        ) -> Result<&mut Self> {
            self.inner
                .replace_paragraph_text(entry_index, position, authored_text)?;
            self.projected.borrow_mut().take();
            Ok(self)
        }

        /// Produce the source-bound raw patch without publishing a package.
        pub fn patch(&self) -> Result<GlossaryBatchPatch> {
            let target = self.projected()?;
            Ok(GlossaryBatchPatch {
                before: self.inner.source().clone(),
                after: target,
            })
        }

        /// Replay a glossary batch patch against this model's source.
        ///
        /// The returned snapshot is a raw projection; no package is
        /// published and no source lineage is required.
        pub fn apply(&self, patch: &GlossaryBatchPatch) -> Result<GlossaryBatchSnapshot> {
            patch.apply(self.source())
        }
    }
}

impl Package {
    /// Capture one bounded source-backed Word text story.
    pub fn story_text_snapshot(&self, selector: Selector) -> Result<Snapshot> {
        self.story_text_snapshot_bounded(selector, self.semantic_limits.story_text_value())
    }

    /// Capture one selected story under explicit finite bounds.
    pub(crate) fn story_text_snapshot_bounded(
        &self,
        selector: Selector,
        limits: Limits,
    ) -> Result<Snapshot> {
        capture(self, selector, limits)
    }

    /// Check a selected-story patch against this package and return the part
    /// and its rewritten bytes (record 0355a).
    pub(crate) fn story_text_target(&self, patch: &Patch) -> Result<(PackURI, Arc<Vec<u8>>)> {
        self.trim_clean_paragraph_index_for_operation();
        self.physical_package()?.check_execution()?;
        let recipient_limits = self.semantic_limits.story_text_value();
        let current = capture(self, patch.selector(), recipient_limits)?;
        let target = patch.apply_content(&current)?;
        story_replacement(
            self,
            patch.is_noop(),
            &current.partname,
            target.raw_bytes(),
            recipient_limits,
            LimitKind::OutputBytes,
            "story publication replacement",
        )
    }
}

/// The replacement bytes of one changed story part, after the output and
/// topology checks of a publication.
fn story_replacement(
    package: &Package,
    noop: bool,
    partname: &PackURI,
    bytes: &[u8],
    limits: Limits,
    kind: LimitKind,
    resource: &'static str,
) -> Result<(PackURI, Arc<Vec<u8>>)> {
    if noop {
        return Ok((partname.clone(), Arc::new(Vec::new())));
    }
    package
        .physical_package()?
        .validate_topology_source_boundary()?;
    if bytes.len() > limits.max_output_bytes {
        return Err(Error::Limit {
            kind,
            actual: bytes.len(),
            maximum: limits.max_output_bytes,
        });
    }
    Ok((
        partname.clone(),
        Arc::new(checked_vec_clone(bytes, resource)?),
    ))
}

fn capture(package: &Package, selector: Selector, limits: Limits) -> Result<Snapshot> {
    package.trim_clean_paragraph_index_for_operation();
    package.physical_package()?.check_execution()?;
    if let Selector::Glossary(glossary) = &selector {
        validate_glossary_selector(glossary)?;
    }
    let source_version = package.physical_package()?.source_version()?;
    let source_lineage = package.physical_package()?.source_lineage();
    let execution = package.physical_package()?.execution_context();
    let (partname, part, relationship_word_namespace) =
        resolve_part(package.physical_package()?, &selector)?;
    let declared = part.declared_uncompressed_size()?;
    if declared > limits.max_xml_bytes as u64 {
        return Err(Error::Limit {
            kind: LimitKind::XmlBytes,
            actual: usize::try_from(declared).unwrap_or(usize::MAX),
            maximum: limits.max_xml_bytes,
        });
    }
    let data = part.data()?;
    if data.as_bytes().len() > limits.max_xml_bytes {
        return Err(Error::Limit {
            kind: LimitKind::XmlBytes,
            actual: data.as_bytes().len(),
            maximum: limits.max_xml_bytes,
        });
    }
    let managed = package
        .physical_package()?
        .cache_diagnostics()
        .budget_managed;
    let payload = if managed {
        Payload::Managed(data)
    } else {
        Payload::Owned(data.into_arc()?)
    };
    let layout = scan_selected_story(
        payload.as_bytes(),
        &selector,
        relationship_word_namespace,
        limits,
        execution.as_ref(),
    )?;
    let (envelope, transaction) = (None, None);
    let part_fingerprint = digest(payload.as_bytes());
    let observed = package.physical_package()?.source_version()?;
    if observed != source_version {
        return Err(Error::Document(crate::Error::Opc(
            litchi_opc::OpcError::SourceChanged {
                expected: source_version,
                actual: observed,
            },
        )));
    }
    package.physical_package()?.check_execution()?;
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
    output_resource: LimitKind,
) -> Result<()> {
    let mut paragraph_bytes = 0usize;
    crate::paragraph::for_each_word_text_chunk(xml, |chunk| {
        let next_paragraph_bytes =
            paragraph_bytes
                .checked_add(chunk.len())
                .ok_or(Error::Limit {
                    kind: LimitKind::ParagraphText,
                    actual: usize::MAX,
                    maximum,
                })?;
        if next_paragraph_bytes > maximum {
            return Err(Error::Limit {
                kind: LimitKind::ParagraphText,
                actual: next_paragraph_bytes,
                maximum,
            });
        }
        let next_output_bytes = output.len().checked_add(chunk.len()).ok_or(Error::Limit {
            kind: output_resource,
            actual: usize::MAX,
            maximum,
        })?;
        if next_output_bytes > maximum {
            return Err(Error::Limit {
                kind: output_resource,
                actual: next_output_bytes,
                maximum,
            });
        }
        output
            .try_reserve(chunk.len())
            .map_err(|source| Error::Allocation {
                resource: output_resource.as_str(),
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
            kind: LimitKind::ParagraphText,
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
            kind: LimitKind::XmlReferenceBytes,
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
    selector: &Selector,
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
            crate::package::validate_document_main_content_type(main.content_type())
                .map_err(Error::Document)?;
            Ok((
                main.partname().clone(),
                main,
                Some(package_dialect.word_namespace()),
            ))
        },
        Selector::Header(index) => {
            let (partname, part, relationship_namespace) =
                resolve_header_footer(package, *index, true)?;
            Ok((partname, part, Some(relationship_namespace)))
        },
        Selector::Footer(index) => {
            let (partname, part, relationship_namespace) =
                resolve_header_footer(package, *index, false)?;
            Ok((partname, part, Some(relationship_namespace)))
        },
        Selector::Footnote(id) => resolve_secondary_part(package, *id, SecondaryProfile::Footnote),
        Selector::Endnote(id) => resolve_secondary_part(package, *id, SecondaryProfile::Endnote),
        Selector::Comment(id) => resolve_secondary_part(package, *id, SecondaryProfile::Comment),
        Selector::Glossary(_) => resolve_glossary_part(package, package_dialect),
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
    crate::package::validate_document_main_content_type(main.content_type())
        .map_err(Error::Document)?;
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

fn resolve_glossary_part<'package>(
    package: &'package SourceBackedPackage,
    package_dialect: WordDialect,
) -> Result<(PackURI, PartView<'package>, Option<&'static [u8]>)> {
    let execution = package.execution_context();
    let main = package.main_document_part()?;
    crate::package::validate_document_main_content_type(main.content_type())
        .map_err(Error::Document)?;
    if package.rels().iter().any(|relationship| {
        matches!(
            relationship.reltype(),
            TRANSITIONAL_GLOSSARY | STRICT_GLOSSARY
        )
    }) {
        return Err(Error::Document(crate::Error::InvalidRelationship(
            "package root cannot source a glossary-document relationship".into(),
        )));
    }
    let expected_relationship = match package_dialect {
        WordDialect::Transitional => TRANSITIONAL_GLOSSARY,
        WordDialect::Strict => STRICT_GLOSSARY,
    };
    let mut selected = None;
    for part in package.iter_parts() {
        check_execution_context(execution.as_ref())?;
        let is_main = part.partname() == main.partname();
        for relationship in part.rels().iter() {
            check_execution_context(execution.as_ref())?;
            if !matches!(
                relationship.reltype(),
                TRANSITIONAL_GLOSSARY | STRICT_GLOSSARY
            ) {
                continue;
            }
            if !is_main {
                return Err(Error::Document(crate::Error::InvalidRelationship(
                    "glossary-document relationship has an invalid source".into(),
                )));
            }
            if relationship.reltype() != expected_relationship {
                return Err(Error::Document(crate::Error::InvalidRelationship(
                    "glossary relationship does not match package conformance".into(),
                )));
            }
            if relationship.is_external() {
                return Err(Error::Document(crate::Error::InvalidRelationship(
                    "glossary relationship cannot be external".into(),
                )));
            }
            if selected.is_some() {
                return Err(Error::Document(crate::Error::InvalidRelationship(
                    "main document has multiple glossary relationships".into(),
                )));
            }
            selected = Some(relationship.target_partname()?);
        }
    }
    let target = selected.ok_or_else(|| {
        Error::Document(crate::Error::InvalidRelationship(
            "main document glossary relationship is missing".into(),
        ))
    })?;
    let part = package.part(&target)?;
    if part.content_type() != ct::WML_DOCUMENT_GLOSSARY {
        return Err(Error::Document(crate::Error::InvalidContentType {
            expected: ct::WML_DOCUMENT_GLOSSARY.to_owned(),
            got: part.content_type().to_owned(),
        }));
    }
    if inbound_count(package, &target)? != 1 {
        return Err(Error::Document(crate::Error::InvalidRelationship(
            "glossary target has an ambiguous inbound relationship closure".into(),
        )));
    }
    for candidate in package.iter_parts() {
        check_execution_context(execution.as_ref())?;
        if candidate.content_type() == ct::WML_DOCUMENT_GLOSSARY
            && candidate.partname() != part.partname()
        {
            return Err(Error::Document(crate::Error::InvalidRelationship(
                "orphan glossary content-type part exists beside the owned root".into(),
            )));
        }
    }
    validate_glossary_root_relationships(package, &part, package_dialect, execution.as_ref())?;
    Ok((
        part.partname().clone(),
        part,
        Some(package_dialect.word_namespace()),
    ))
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

#[derive(Clone, Copy, PartialEq, Eq)]
enum OwnedStoryKind {
    Header,
    Footer,
    Footnotes,
    Endnotes,
    Comments,
    Glossary,
}

impl OwnedStoryKind {
    const fn index(self) -> Option<usize> {
        match self {
            Self::Header => Some(0),
            Self::Footer => Some(1),
            Self::Footnotes => Some(2),
            Self::Endnotes => Some(3),
            Self::Comments => Some(4),
            Self::Glossary => None,
        }
    }

    const fn content_type(self) -> Option<&'static str> {
        match self {
            Self::Header => Some(ct::WML_HEADER),
            Self::Footer => Some(ct::WML_FOOTER),
            Self::Footnotes => Some(ct::WML_FOOTNOTES),
            Self::Endnotes => Some(ct::WML_ENDNOTES),
            Self::Comments => Some(ct::WML_COMMENTS),
            Self::Glossary => None,
        }
    }
}

fn validate_glossary_root_relationships(
    package: &SourceBackedPackage,
    root: &PartView<'_>,
    package_dialect: WordDialect,
    execution: Option<&ExecutionContext>,
) -> Result<()> {
    let mut seen = [false; 5];
    for relationship in root.rels().iter() {
        check_execution_context(execution)?;
        if word_relationship_uri_dialect(relationship.reltype())
            .is_some_and(|dialect| dialect != package_dialect)
        {
            return Err(Error::Document(crate::Error::InvalidRelationship(
                "glossary auxiliary relationship uses a dialect different from the package root"
                    .into(),
            )));
        }
        let Some((dialect, kind)) = word_story_relationship_kind(relationship.reltype()) else {
            if !relationship.is_external() {
                let target = relationship.target_partname()?;
                let target_part = package.part(&target)?;
                if inbound_count(package, target_part.partname())? != 1 {
                    return Err(Error::Document(crate::Error::InvalidRelationship(
                        "glossary auxiliary target has an ambiguous inbound relationship closure"
                            .into(),
                    )));
                }
            }
            continue;
        };
        if dialect != package_dialect {
            return Err(Error::Document(crate::Error::InvalidRelationship(
                "Word story relationship uses a dialect different from the package root".into(),
            )));
        }
        if kind == OwnedStoryKind::Glossary {
            return Err(Error::Document(crate::Error::InvalidRelationship(
                "a glossary story cannot own another glossary story".into(),
            )));
        }
        if relationship.is_external() {
            return Err(Error::Document(crate::Error::InvalidRelationship(
                "glossary-owned Word story relationship cannot be external".into(),
            )));
        }
        if let Some(index) = kind.index()
            && std::mem::replace(&mut seen[index], true)
        {
            return Err(Error::Document(crate::Error::InvalidRelationship(
                "glossary story has multiple relationships of one story kind".into(),
            )));
        }
        let target = relationship.target_partname()?;
        let target_part = package.part(&target)?;
        let expected = kind.content_type().ok_or(unsupported_xml(
            "glossary story relationship has no content type",
        ))?;
        if target_part.content_type() != expected {
            return Err(Error::Document(crate::Error::InvalidContentType {
                expected: expected.to_owned(),
                got: target_part.content_type().to_owned(),
            }));
        }
        if inbound_count(package, target_part.partname())? != 1 {
            return Err(Error::Document(crate::Error::InvalidRelationship(
                "glossary auxiliary story has an ambiguous inbound relationship closure".into(),
            )));
        }
        reject_outbound_story_relationships(&target_part, execution)?;
    }
    Ok(())
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

fn word_story_relationship_kind(reltype: &str) -> Option<(WordDialect, OwnedStoryKind)> {
    match reltype {
        rt::HEADER => Some((WordDialect::Transitional, OwnedStoryKind::Header)),
        rt::FOOTER => Some((WordDialect::Transitional, OwnedStoryKind::Footer)),
        rt::FOOTNOTES => Some((WordDialect::Transitional, OwnedStoryKind::Footnotes)),
        rt::ENDNOTES => Some((WordDialect::Transitional, OwnedStoryKind::Endnotes)),
        rt::COMMENTS => Some((WordDialect::Transitional, OwnedStoryKind::Comments)),
        TRANSITIONAL_GLOSSARY => Some((WordDialect::Transitional, OwnedStoryKind::Glossary)),
        STRICT_HEADER => Some((WordDialect::Strict, OwnedStoryKind::Header)),
        STRICT_FOOTER => Some((WordDialect::Strict, OwnedStoryKind::Footer)),
        STRICT_FOOTNOTES => Some((WordDialect::Strict, OwnedStoryKind::Footnotes)),
        STRICT_ENDNOTES => Some((WordDialect::Strict, OwnedStoryKind::Endnotes)),
        rt::STRICT_COMMENTS => Some((WordDialect::Strict, OwnedStoryKind::Comments)),
        STRICT_GLOSSARY => Some((WordDialect::Strict, OwnedStoryKind::Glossary)),
        _ => None,
    }
}

fn word_story_relationship_dialect(reltype: &str) -> Option<WordDialect> {
    word_story_relationship_kind(reltype).map(|(dialect, _)| dialect)
}

fn word_relationship_uri_dialect(reltype: &str) -> Option<WordDialect> {
    if reltype.starts_with("http://schemas.openxmlformats.org/officeDocument/2006/relationships/") {
        Some(WordDialect::Transitional)
    } else if reltype.starts_with("http://purl.oclc.org/ooxml/officeDocument/relationships/") {
        Some(WordDialect::Strict)
    } else {
        None
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
    crate::package::validate_document_main_content_type(main.content_type())
        .map_err(Error::Document)?;
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
                kind: LimitKind::RelationshipTargets,
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

fn preflight_attribute_value(value: &[u8], maximum: usize, resource: LimitKind) -> Result<()> {
    let raw_maximum = maximum.checked_mul(4).unwrap_or(usize::MAX);
    if value.len() > raw_maximum {
        return Err(Error::Limit {
            kind: resource,
            actual: value.len(),
            maximum: raw_maximum,
        });
    }
    let _ = checked_decoded_attribute_value_length(value, maximum, resource)?;
    Ok(())
}

fn checked_decoded_attribute_value_length(
    value: &[u8],
    maximum: usize,
    resource: LimitKind,
) -> Result<usize> {
    let raw_maximum = maximum.checked_mul(4).unwrap_or(usize::MAX);
    if value.len() > raw_maximum {
        return Err(Error::Limit {
            kind: resource,
            actual: value.len(),
            maximum: raw_maximum,
        });
    }
    let mut decoded_length = 0usize;
    let mut offset = 0usize;
    while offset < value.len() {
        if value[offset] != b'&' {
            decoded_length = decoded_length.checked_add(1).ok_or(Error::Limit {
                kind: resource,
                actual: usize::MAX,
                maximum,
            })?;
            offset += 1;
            continue;
        }
        let entity_start = offset + 1;
        let entity_end = value[entity_start..]
            .iter()
            .position(|byte| *byte == b';')
            .map(|relative| entity_start + relative)
            .ok_or_else(|| {
                Error::Document(crate::Error::InvalidFormat(
                    "unterminated XML attribute escape".into(),
                ))
            })?;
        let entity = &value[entity_start..entity_end];
        let replacement_length = if matches!(entity, b"lt" | b"gt" | b"amp" | b"apos" | b"quot") {
            1
        } else if entity.first() == Some(&b'#') {
            let (radix, digits) = if entity.get(1) == Some(&b'x') || entity.get(1) == Some(&b'X') {
                (16_u32, &entity[2..])
            } else {
                (10_u32, &entity[1..])
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
            char::from_u32(codepoint)
                .ok_or_else(|| {
                    Error::Document(crate::Error::InvalidFormat(
                        "XML character reference is not a Unicode scalar".into(),
                    ))
                })?
                .len_utf8()
        } else {
            return Err(Error::Document(crate::Error::InvalidFormat(
                "unsupported XML attribute escape".into(),
            )));
        };
        decoded_length = decoded_length
            .checked_add(replacement_length)
            .ok_or(Error::Limit {
                kind: resource,
                actual: usize::MAX,
                maximum,
            })?;
        if decoded_length > maximum {
            return Err(Error::Limit {
                kind: resource,
                actual: decoded_length,
                maximum,
            });
        }
        offset = entity_end + 1;
    }
    Ok(decoded_length)
}

fn decode_attribute_value(value: &[u8]) -> Result<Vec<u8>> {
    let decoded_length =
        checked_decoded_attribute_value_length(value, usize::MAX, LimitKind::AttributeValue)?;
    let value = std::str::from_utf8(value).map_err(|error| {
        Error::Document(crate::Error::InvalidFormat(format!(
            "invalid XML attribute UTF-8: {error}"
        )))
    })?;
    let mut decoded = Vec::new();
    decoded
        .try_reserve_exact(decoded_length)
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
        let mut numeric = [0u8; 4];
        let replacement: &[u8] = if entity == b"lt" {
            b"<"
        } else if entity == b"gt" {
            b">"
        } else if entity == b"amp" {
            b"&"
        } else if entity == b"apos" {
            b"'"
        } else if entity == b"quot" {
            b"\""
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
            let numeric_length = character.encode_utf8(&mut numeric).len();
            &numeric[..numeric_length]
        } else {
            return Err(Error::Document(crate::Error::InvalidFormat(
                "unsupported XML attribute escape".into(),
            )));
        };
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
        expected_word_namespace: Option<&[u8]>,
    ) -> Result<(Self, Vec<u8>)> {
        if parent.bindings.len() > max_bindings {
            return Err(Error::Limit {
                kind: LimitKind::NamespaceBindings,
                actual: parent.bindings.len(),
                maximum: max_bindings,
            });
        }
        if parent.raw_namespace_bytes > max_namespace_bytes {
            return Err(Error::Limit {
                kind: LimitKind::NamespaceBytes,
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
        for attribute in start.checked_attributes() {
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
                        kind: LimitKind::NamespaceBytes,
                        actual: usize::MAX,
                        maximum: max_namespace_bytes,
                    })?;
                let raw_namespace_bytes = context
                    .raw_namespace_bytes
                    .checked_add(raw_declaration_bytes)
                    .ok_or(Error::Limit {
                        kind: LimitKind::NamespaceBytes,
                        actual: usize::MAX,
                        maximum: max_namespace_bytes,
                    })?;
                if raw_namespace_bytes > max_namespace_bytes {
                    return Err(Error::Limit {
                        kind: LimitKind::NamespaceBytes,
                        actual: raw_namespace_bytes,
                        maximum: max_namespace_bytes,
                    });
                }
                preflight_attribute_value(
                    value,
                    max_namespace_bytes,
                    LimitKind::NamespaceAttributeValue,
                )?;
                let decoded_value = decode_attribute_value(value)?;
                if context.bindings.len() >= max_bindings {
                    return Err(Error::Limit {
                        kind: LimitKind::NamespaceBindings,
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
                        kind: LimitKind::NamespaceAttributes,
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

        for attribute in start.checked_attributes() {
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
            if !is_namespace_declaration(key) {
                let mce_attribute = context
                    .resolve_attribute(key)
                    .is_some_and(|namespace| namespace == MCE_NAMESPACE);
                if mce_attribute {
                    preflight_attribute_value(
                        value,
                        max_namespace_bytes,
                        LimitKind::AttributeValue,
                    )?;
                    if key.local_name().as_ref() != b"Ignorable" {
                        return Err(unsupported_xml(
                            "markup-compatibility branch directives are not accepted in a source-bound story",
                        ));
                    }
                }
            }
        }
        let word_namespace = expected_word_namespace
            .or_else(|| context.resolve_element(start.name()))
            .unwrap_or_default();
        reject_word_dialect_attributes(start, &context, word_namespace)?;
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
    let escaped_namespace_bytes = namespace.iter().try_fold(0usize, |total, byte| {
        let byte_len = match byte {
            b'&' => 5,
            b'<' => 4,
            b'"' => 6,
            _ => 1,
        };
        total.checked_add(byte_len).ok_or(Error::Limit {
            kind: LimitKind::NamespaceAttributes,
            actual: usize::MAX,
            maximum,
        })
    })?;
    let prefix_bytes = if prefix.is_empty() {
        0
    } else {
        prefix.len().checked_add(1).ok_or(Error::Limit {
            kind: LimitKind::NamespaceAttributes,
            actual: usize::MAX,
            maximum,
        })?
    };
    let declaration_bytes = b" xmlns"
        .len()
        .checked_add(prefix_bytes)
        .and_then(|value| value.checked_add(b"=\"".len()))
        .and_then(|value| value.checked_add(escaped_namespace_bytes))
        .and_then(|value| value.checked_add(1))
        .ok_or(Error::Limit {
            kind: LimitKind::NamespaceAttributes,
            actual: usize::MAX,
            maximum,
        })?;
    let projected = output
        .len()
        .checked_add(declaration_bytes)
        .ok_or(Error::Limit {
            kind: LimitKind::NamespaceAttributes,
            actual: usize::MAX,
            maximum,
        })?;
    if projected > maximum {
        return Err(Error::Limit {
            kind: LimitKind::NamespaceAttributes,
            actual: projected,
            maximum,
        });
    }
    output
        .try_reserve_exact(declaration_bytes)
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

/// The byte offset in the reader's input of its current position; `origin`
/// is that input's [`ReaderOrigin`].
fn reader_position(reader: &Reader<&[u8]>, origin: ReaderOrigin) -> Result<usize> {
    let position = reader.buffer_position();
    origin.offset(position).ok_or_else(|| {
        Error::Document(crate::Error::Xml(format!(
            "XML buffer position {position} exceeds addressable input",
        )))
    })
}

fn secondary_identity(selector: &Selector) -> Option<(&'static str, u32)> {
    match selector {
        Selector::Footnote(id) => Some(("footnote", *id)),
        Selector::Endnote(id) => Some(("endnote", *id)),
        Selector::Comment(id) => Some(("comment", *id)),
        Selector::Main | Selector::Header(_) | Selector::Footer(_) | Selector::Glossary(_) => None,
    }
}

fn selector_matches_entry(
    start: &BytesStart<'_>,
    context: &NamespaceContext,
    word_namespace: &[u8],
    selector: &Selector,
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
        Selector::Main | Selector::Header(_) | Selector::Footer(_) | Selector::Glossary(_) => false,
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
    for attribute in start.checked_attributes() {
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
        if let Err(error) = preflight_attribute_value(
            attribute.value.as_ref(),
            MAX_GLOSSARY_SELECTOR_NAME_BYTES,
            LimitKind::AttributeValue,
        ) {
            return Some(Err(error));
        }
        selected = Some(decode_attribute_value(attribute.value.as_ref()));
    }
    selected
}

fn validate_glossary_selector(selector: &GlossarySelector) -> Result<()> {
    match selector {
        GlossarySelector::ByName(name) => {
            let _ = glossary_name_key(name)?;
        },
        GlossarySelector::ById(id) => validate_glossary_selector_id(id)?,
        GlossarySelector::ByNameAndId { name, id } => {
            let _ = glossary_name_key(name)?;
            validate_glossary_selector_id(id)?;
        },
        GlossarySelector::ByIndex(_) => {},
    }
    Ok(())
}

fn validate_glossary_selector_id(id: &crate::glossary::Id) -> Result<()> {
    if id.as_str().len() > MAX_GLOSSARY_SELECTOR_ID_BYTES {
        return Err(Error::Limit {
            kind: LimitKind::GlossarySelectorIdBytes,
            actual: id.as_str().len(),
            maximum: MAX_GLOSSARY_SELECTOR_ID_BYTES,
        });
    }
    Ok(())
}

fn glossary_name_key(value: &str) -> Result<String> {
    if value.len() > MAX_GLOSSARY_SELECTOR_NAME_BYTES {
        return Err(Error::Limit {
            kind: LimitKind::GlossarySelectorNameBytes,
            actual: value.len(),
            maximum: MAX_GLOSSARY_SELECTOR_NAME_BYTES,
        });
    }
    use caseless::Caseless;
    use unicode_normalization::UnicodeNormalization;

    let mut key = String::new();
    key.try_reserve(value.len())
        .map_err(|source| Error::Allocation {
            resource: "glossary selector name key",
            source,
        })?;
    for character in value.chars().nfd().default_case_fold().nfd() {
        let next = key
            .len()
            .checked_add(character.len_utf8())
            .ok_or(Error::Limit {
                kind: LimitKind::GlossarySelectorNameKey,
                actual: usize::MAX,
                maximum: MAX_GLOSSARY_SELECTOR_NAME_BYTES,
            })?;
        if next > MAX_GLOSSARY_SELECTOR_NAME_BYTES {
            return Err(Error::Limit {
                kind: LimitKind::GlossarySelectorNameKey,
                actual: next,
                maximum: MAX_GLOSSARY_SELECTOR_NAME_BYTES,
            });
        }
        if key.capacity() < next {
            key.try_reserve(character.len_utf8())
                .map_err(|source| Error::Allocation {
                    resource: "glossary selector name key",
                    source,
                })?;
        }
        key.push(character);
    }
    Ok(key)
}

fn glossary_metadata_text(
    start: &BytesStart<'_>,
    context: &NamespaceContext,
    word_namespace: &[u8],
    field: &'static str,
) -> Result<Option<String>> {
    let Some(value) = word_attribute(start, context, word_namespace, b"val").transpose()? else {
        return Ok(None);
    };
    if value.len() > MAX_GLOSSARY_SELECTOR_NAME_BYTES {
        return Err(Error::Limit {
            kind: LimitKind::GlossarySelectorMetadataBytes,
            actual: value.len(),
            maximum: MAX_GLOSSARY_SELECTOR_NAME_BYTES,
        });
    }
    let value = String::from_utf8(value).map_err(|_| {
        unsupported_xml(match field {
            "name" => "glossary entry name is not valid UTF-8",
            _ => "glossary entry GUID is not valid UTF-8",
        })
    })?;
    validate_xml_1_0_value(
        &value,
        match field {
            "name" => "glossary entry name contains a character forbidden by XML 1.0",
            _ => "glossary entry GUID contains a character forbidden by XML 1.0",
        },
    )?;
    Ok(Some(value))
}

fn validate_xml_1_0_value(value: &str, reason: &'static str) -> Result<()> {
    if value.chars().all(|character| {
        matches!(
            character,
            '\u{9}' | '\u{a}' | '\u{d}' | '\u{20}'..='\u{d7ff}'
                | '\u{e000}'..='\u{fffd}'
                | '\u{10000}'..='\u{10ffff}'
        )
    }) {
        Ok(())
    } else {
        Err(unsupported_xml(reason))
    }
}

fn glossary_metadata_id(
    start: &BytesStart<'_>,
    context: &NamespaceContext,
    word_namespace: &[u8],
) -> Result<Option<crate::glossary::Id>> {
    let Some(value) = glossary_metadata_text(start, context, word_namespace, "guid")? else {
        return Ok(None);
    };
    let id = crate::glossary::Id::new(value)
        .map_err(|_| unsupported_xml("glossary entry GUID is not a semantic building-block ID"))?;
    Ok(Some(id))
}

fn glossary_selector_matches(
    selector: &GlossarySelector,
    requested_name_key: Option<&str>,
    index: usize,
    name_key: Option<&str>,
    id: Option<&str>,
) -> bool {
    match selector {
        GlossarySelector::ByName(_) => requested_name_key == name_key,
        GlossarySelector::ById(requested) => id == Some(requested.as_str()),
        GlossarySelector::ByNameAndId {
            id: requested_id, ..
        } => requested_name_key == name_key && id == Some(requested_id.as_str()),
        GlossarySelector::ByIndex(requested) => *requested == index,
    }
}

fn scan_selected_story(
    xml: &[u8],
    selector: &Selector,
    expected_word_namespace: Option<&[u8]>,
    limits: Limits,
    execution: Option<&ExecutionContext>,
) -> Result<Layout> {
    match selector {
        Selector::Glossary(glossary) => {
            scan_glossary(xml, glossary, expected_word_namespace, limits, execution)
        },
        _ => scan_story(xml, selector, expected_word_namespace, limits, execution),
    }
}

fn scan_glossary(
    xml: &[u8],
    selector: &GlossarySelector,
    expected_word_namespace: Option<&[u8]>,
    limits: Limits,
    execution: Option<&ExecutionContext>,
) -> Result<Layout> {
    let mut layouts = scan_glossary_many(
        xml,
        std::slice::from_ref(selector),
        expected_word_namespace,
        limits,
        execution,
    )?;
    layouts.pop().ok_or_else(|| {
        Error::Document(crate::Error::InvalidFormat(
            "glossary selector inventory returned no layout".into(),
        ))
    })
}

fn scan_glossary_many(
    xml: &[u8],
    selectors: &[GlossarySelector],
    expected_word_namespace: Option<&[u8]>,
    limits: Limits,
    execution: Option<&ExecutionContext>,
) -> Result<Vec<Layout>> {
    if xml.len() > limits.max_xml_bytes {
        return Err(Error::Limit {
            kind: LimitKind::XmlBytes,
            actual: xml.len(),
            maximum: limits.max_xml_bytes,
        });
    }
    validate_glossary_batch_selector_limits(selectors, limits)?;
    let mut requested_name_keys = Vec::new();
    requested_name_keys
        .try_reserve_exact(selectors.len())
        .map_err(|source| Error::Allocation {
            resource: "glossary selector keys",
            source,
        })?;
    for selector in selectors {
        requested_name_keys.push(match selector {
            GlossarySelector::ByName(name) | GlossarySelector::ByNameAndId { name, .. } => {
                Some(glossary_name_key(name)?)
            },
            GlossarySelector::ById(_) | GlossarySelector::ByIndex(_) => None,
        });
    }
    let mut matched = Vec::new();
    matched
        .try_reserve_exact(selectors.len())
        .map_err(|source| Error::Allocation {
            resource: "glossary selector matches",
            source,
        })?;
    matched.resize_with(selectors.len(), || false);
    let mut reader = Reader::from_reader(xml);
    let origin = ReaderOrigin::of(xml);
    reader.config_mut().trim_text(false);
    reader.config_mut().check_end_names = true;
    let mut event_count = 0usize;
    let mut depth = 0usize;
    let mut root_seen = false;
    let mut root_closed = false;
    let mut word_namespace = Vec::new();
    let mut word_namespace_shared: Option<Arc<[u8]>> = None;
    let mut namespace_stack = Vec::new();
    let empty_namespace_context = NamespaceContext::default();
    let mut root_has_background = false;
    let mut root_has_doc_parts = false;
    let mut doc_parts_depth = None;
    let mut doc_part_depth = None;
    let mut doc_part_pr_depth = None;
    let mut body_depth = None;
    let mut entry_count = 0usize;
    let mut current_entry_index = 0usize;
    let mut current_entry_start = None;
    let mut current_entry_selected = false;
    let mut current_entry_pr_seen = false;
    let mut current_entry_name_key = None;
    let mut current_entry_id = None;
    let mut current_body_count = 0usize;
    let mut current_body_inner_start = None;
    let mut current_body_inner_end = None;
    let mut current_body_namespace_attributes = None;
    let mut current_body_has_w_namespace = false;
    let mut paragraph_stack = Vec::new();
    let mut current_paragraphs = Vec::new();
    let mut selected_layouts = Vec::new();
    selected_layouts
        .try_reserve_exact(selectors.len())
        .map_err(|source| Error::Allocation {
            resource: "glossary selected layouts",
            source,
        })?;

    loop {
        let before = reader_position(&reader, origin)?;
        let event = reader.read_event().map_err(|error| {
            Error::Document(crate::Error::InvalidFormat(format!(
                "invalid glossary XML: {error}",
            )))
        })?;
        event_count = event_count.checked_add(1).ok_or(Error::Limit {
            kind: LimitKind::XmlEvents,
            actual: usize::MAX,
            maximum: limits.max_events,
        })?;
        if event_count > limits.max_events {
            return Err(Error::Limit {
                kind: LimitKind::XmlEvents,
                actual: event_count,
                maximum: limits.max_events,
            });
        }
        if event_count & 63 == 0 {
            check_execution_context(execution)?;
        }
        let after = reader_position(&reader, origin)?;
        match event {
            Event::Start(start) => {
                let event_start = before;
                let parent_depth = depth;
                let logical_depth = parent_depth.checked_add(1).ok_or(Error::Limit {
                    kind: LimitKind::XmlDepth,
                    actual: usize::MAX,
                    maximum: limits.max_depth,
                })?;
                if logical_depth > limits.max_depth {
                    return Err(Error::Limit {
                        kind: LimitKind::XmlDepth,
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
                    expected_word_namespace,
                )?;
                if parent_depth == 0 {
                    if root_seen || root_closed {
                        return Err(unsupported_xml("multiple glossary roots"));
                    }
                    let info = inspect_root(
                        &start,
                        RootKind::Glossary,
                        &context,
                        element_namespace_attributes,
                    )?;
                    if expected_word_namespace
                        .is_some_and(|expected| info.word_namespace.as_slice() != expected)
                    {
                        return Err(unsupported_xml(
                            "glossary relationship and XML dialects differ",
                        ));
                    }
                    word_namespace = info.word_namespace;
                    word_namespace_shared = Some(word_namespace.clone().into());
                    root_seen = true;
                } else if !root_seen || root_closed {
                    return Err(unsupported_xml("glossary content occurs outside its root"));
                }
                reject_wrong_word_namespace(start.name(), &context, &word_namespace)?;
                if parent_depth == 1 {
                    if is_word_element(start.name(), &context, &word_namespace, b"background") {
                        if root_has_background || root_has_doc_parts {
                            return Err(unsupported_xml(
                                "glossary root has an invalid background order",
                            ));
                        }
                        root_has_background = true;
                    } else if is_word_element(start.name(), &context, &word_namespace, b"docParts")
                    {
                        if root_has_doc_parts {
                            return Err(unsupported_xml(
                                "glossary root contains multiple docParts",
                            ));
                        }
                        root_has_doc_parts = true;
                        doc_parts_depth = Some(logical_depth);
                    } else if is_glossary_root_extension(start.name(), &context, &word_namespace) {
                        // Preserve qualified extension metadata outside the Word schema.
                    } else {
                        return Err(unsupported_xml(
                            "glossary root has an unexpected direct child",
                        ));
                    }
                } else if doc_parts_depth == Some(parent_depth) {
                    if !is_word_element(start.name(), &context, &word_namespace, b"docPart") {
                        return Err(unsupported_xml(
                            "glossary docParts contains a non-docPart child",
                        ));
                    }
                    if doc_part_depth.is_some() {
                        return Err(unsupported_xml("glossary docPart nesting is invalid"));
                    }
                    entry_count = entry_count.checked_add(1).ok_or(Error::Limit {
                        kind: LimitKind::Entries,
                        actual: usize::MAX,
                        maximum: limits.max_entries,
                    })?;
                    if entry_count > limits.max_entries {
                        return Err(Error::Limit {
                            kind: LimitKind::Entries,
                            actual: entry_count,
                            maximum: limits.max_entries,
                        });
                    }
                    current_entry_index = entry_count - 1;
                    current_entry_start = Some(event_start);
                    current_entry_selected = false;
                    current_entry_pr_seen = false;
                    current_entry_name_key = None;
                    current_entry_id = None;
                    current_body_count = 0;
                    current_body_inner_start = None;
                    current_body_inner_end = None;
                    current_body_namespace_attributes = None;
                    current_body_has_w_namespace = false;
                    current_paragraphs.clear();
                    doc_part_depth = Some(logical_depth);
                } else if doc_part_depth == Some(parent_depth) {
                    if is_word_element(start.name(), &context, &word_namespace, b"docPartPr") {
                        if current_entry_pr_seen || current_body_count != 0 {
                            return Err(unsupported_xml(
                                "glossary docPartPr must precede its docPartBody and occur once",
                            ));
                        }
                        current_entry_pr_seen = true;
                        doc_part_pr_depth = Some(logical_depth);
                    } else if is_word_element(
                        start.name(),
                        &context,
                        &word_namespace,
                        b"docPartBody",
                    ) {
                        if current_body_count != 0 {
                            return Err(unsupported_xml(
                                "glossary docPart has multiple docPartBody elements",
                            ));
                        }
                        current_body_count =
                            current_body_count.checked_add(1).ok_or(Error::Limit {
                                kind: LimitKind::GlossaryEntryBodies,
                                actual: usize::MAX,
                                maximum: 1,
                            })?;
                        if body_depth.is_some() {
                            return Err(unsupported_xml("glossary docPartBody nesting is invalid"));
                        }
                        current_entry_selected =
                            selectors
                                .iter()
                                .enumerate()
                                .any(|(selector_index, requested)| {
                                    glossary_selector_matches(
                                        requested,
                                        requested_name_keys[selector_index].as_deref(),
                                        current_entry_index,
                                        current_entry_name_key.as_deref(),
                                        current_entry_id.as_ref().map(crate::glossary::Id::as_str),
                                    )
                                });
                        body_depth = Some(logical_depth);
                        if current_entry_selected {
                            current_body_inner_start = Some(after);
                            let (attributes, has_w_namespace) = context
                                .wrapper_namespace_attributes(
                                    &word_namespace,
                                    limits.max_xml_bytes,
                                )?;
                            current_body_namespace_attributes = Some(attributes);
                            current_body_has_w_namespace = has_w_namespace;
                        }
                    } else if context.resolve_element(start.name())
                        == Some(word_namespace.as_slice())
                    {
                        return Err(unsupported_xml(
                            "glossary docPart has an unexpected direct Word child",
                        ));
                    }
                } else if doc_part_pr_depth == Some(parent_depth)
                    && is_word_element(start.name(), &context, &word_namespace, b"name")
                {
                    let value = glossary_metadata_text(&start, &context, &word_namespace, "name")?;
                    if let Some(value) = value {
                        current_entry_name_key = Some(glossary_name_key(&value)?);
                    }
                } else if doc_part_pr_depth == Some(parent_depth)
                    && is_word_element(start.name(), &context, &word_namespace, b"guid")
                {
                    current_entry_id = glossary_metadata_id(&start, &context, &word_namespace)?;
                } else if body_depth == Some(parent_depth)
                    && current_entry_selected
                    && is_word_element(start.name(), &context, &word_namespace, b"p")
                {
                    ensure_paragraph_capacity(
                        current_paragraphs.len(),
                        paragraph_stack.len(),
                        limits.max_paragraphs,
                    )?;
                    reserve_one(&mut paragraph_stack, "glossary paragraph stack")?;
                    paragraph_stack.push((logical_depth, event_start));
                    reserve_one(&mut current_paragraphs, "glossary paragraph ranges")?;
                }
                depth = logical_depth;
                reserve_one(&mut namespace_stack, "glossary namespace stack")?;
                namespace_stack.push(context);
            },
            Event::Empty(start) => {
                let event_start = before;
                let parent_depth = depth;
                let logical_depth = parent_depth.checked_add(1).ok_or(Error::Limit {
                    kind: LimitKind::XmlDepth,
                    actual: usize::MAX,
                    maximum: limits.max_depth,
                })?;
                if logical_depth > limits.max_depth {
                    return Err(Error::Limit {
                        kind: LimitKind::XmlDepth,
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
                    expected_word_namespace,
                )?;
                if parent_depth == 0 {
                    if root_seen || root_closed {
                        return Err(unsupported_xml("multiple glossary roots"));
                    }
                    let info = inspect_root(
                        &start,
                        RootKind::Glossary,
                        &context,
                        element_namespace_attributes,
                    )?;
                    if expected_word_namespace
                        .is_some_and(|expected| info.word_namespace.as_slice() != expected)
                    {
                        return Err(unsupported_xml(
                            "glossary relationship and XML dialects differ",
                        ));
                    }
                    word_namespace = info.word_namespace;
                    word_namespace_shared = Some(word_namespace.clone().into());
                    root_seen = true;
                    root_closed = true;
                } else if !root_seen || root_closed {
                    return Err(unsupported_xml("glossary content occurs outside its root"));
                } else {
                    reject_wrong_word_namespace(start.name(), &context, &word_namespace)?;
                    if parent_depth == 1 {
                        if is_word_element(start.name(), &context, &word_namespace, b"background") {
                            if root_has_background || root_has_doc_parts {
                                return Err(unsupported_xml(
                                    "glossary root has an invalid background order",
                                ));
                            }
                            root_has_background = true;
                        } else if is_word_element(
                            start.name(),
                            &context,
                            &word_namespace,
                            b"docParts",
                        ) {
                            if root_has_doc_parts {
                                return Err(unsupported_xml(
                                    "glossary root contains multiple docParts",
                                ));
                            }
                            root_has_doc_parts = true;
                        } else if is_glossary_root_extension(
                            start.name(),
                            &context,
                            &word_namespace,
                        ) {
                            // Preserve qualified extension metadata outside the Word schema.
                        } else {
                            return Err(unsupported_xml(
                                "glossary root has an unexpected direct child",
                            ));
                        }
                    } else if doc_parts_depth == Some(parent_depth) {
                        if !is_word_element(start.name(), &context, &word_namespace, b"docPart") {
                            return Err(unsupported_xml(
                                "glossary docParts contains a non-docPart child",
                            ));
                        }
                        entry_count = entry_count.checked_add(1).ok_or(Error::Limit {
                            kind: LimitKind::Entries,
                            actual: usize::MAX,
                            maximum: limits.max_entries,
                        })?;
                        if entry_count > limits.max_entries {
                            return Err(Error::Limit {
                                kind: LimitKind::Entries,
                                actual: entry_count,
                                maximum: limits.max_entries,
                            });
                        }
                        return Err(unsupported_xml(
                            "glossary docPart must contain exactly one docPartBody",
                        ));
                    } else if doc_part_depth == Some(parent_depth) {
                        if is_word_element(start.name(), &context, &word_namespace, b"docPartPr") {
                            if current_entry_pr_seen || current_body_count != 0 {
                                return Err(unsupported_xml(
                                    "glossary docPartPr must precede its docPartBody and occur once",
                                ));
                            }
                            current_entry_pr_seen = true;
                        } else if is_word_element(
                            start.name(),
                            &context,
                            &word_namespace,
                            b"docPartBody",
                        ) {
                            if current_body_count != 0 {
                                return Err(unsupported_xml(
                                    "glossary docPart has multiple docPartBody elements",
                                ));
                            }
                            current_body_count =
                                current_body_count.checked_add(1).ok_or(Error::Limit {
                                    kind: LimitKind::GlossaryEntryBodies,
                                    actual: usize::MAX,
                                    maximum: 1,
                                })?;
                            current_entry_selected =
                                selectors
                                    .iter()
                                    .enumerate()
                                    .any(|(selector_index, requested)| {
                                        glossary_selector_matches(
                                            requested,
                                            requested_name_keys[selector_index].as_deref(),
                                            current_entry_index,
                                            current_entry_name_key.as_deref(),
                                            current_entry_id
                                                .as_ref()
                                                .map(crate::glossary::Id::as_str),
                                        )
                                    });
                            if current_entry_selected {
                                current_body_inner_start = Some(after);
                                current_body_inner_end = Some(after);
                                let (attributes, has_w_namespace) = context
                                    .wrapper_namespace_attributes(
                                        &word_namespace,
                                        limits.max_xml_bytes,
                                    )?;
                                current_body_namespace_attributes = Some(attributes);
                                current_body_has_w_namespace = has_w_namespace;
                            }
                        } else if context.resolve_element(start.name())
                            == Some(word_namespace.as_slice())
                        {
                            return Err(unsupported_xml(
                                "glossary docPart has an unexpected direct Word child",
                            ));
                        }
                    } else if doc_part_pr_depth == Some(parent_depth)
                        && is_word_element(start.name(), &context, &word_namespace, b"name")
                    {
                        let value =
                            glossary_metadata_text(&start, &context, &word_namespace, "name")?;
                        if let Some(value) = value {
                            current_entry_name_key = Some(glossary_name_key(&value)?);
                        }
                    } else if doc_part_pr_depth == Some(parent_depth)
                        && is_word_element(start.name(), &context, &word_namespace, b"guid")
                    {
                        current_entry_id = glossary_metadata_id(&start, &context, &word_namespace)?;
                    } else if body_depth == Some(parent_depth)
                        && current_entry_selected
                        && is_word_element(start.name(), &context, &word_namespace, b"p")
                    {
                        ensure_paragraph_capacity(
                            current_paragraphs.len(),
                            paragraph_stack.len(),
                            limits.max_paragraphs,
                        )?;
                        reserve_one(&mut current_paragraphs, "glossary paragraph ranges")?;
                        current_paragraphs.push(Range {
                            start: event_start,
                            end: after,
                        });
                    }
                }
            },
            Event::End(end) => {
                let event_start = before;
                if depth == 0 {
                    return Err(unsupported_xml("glossary XML has an unmatched end element"));
                }
                let context = namespace_stack
                    .last()
                    .ok_or(unsupported_xml("glossary namespace stack is incomplete"))?;
                reject_wrong_word_namespace(end.name(), context, &word_namespace)?;
                if is_word_element(end.name(), context, &word_namespace, b"p")
                    && let Some((paragraph_depth, start)) = paragraph_stack.pop()
                {
                    if paragraph_depth != depth {
                        return Err(unsupported_xml("nested or crossing glossary paragraphs"));
                    }
                    current_paragraphs.push(Range { start, end: after });
                }
                if body_depth == Some(depth) {
                    if !paragraph_stack.is_empty() {
                        return Err(unsupported_xml(
                            "glossary body closes with an open paragraph",
                        ));
                    }
                    if current_entry_selected {
                        current_body_inner_end = Some(event_start);
                    }
                    body_depth = None;
                }
                if doc_part_pr_depth == Some(depth) {
                    doc_part_pr_depth = None;
                }
                if doc_part_depth == Some(depth) {
                    if current_body_count != 1 {
                        return Err(unsupported_xml(
                            "every glossary docPart must contain exactly one docPartBody",
                        ));
                    }
                    if current_entry_selected {
                        let entry_start = current_entry_start
                            .ok_or(unsupported_xml("glossary docPart start is missing"))?;
                        let entry_bytes = after
                            .checked_sub(entry_start)
                            .ok_or(unsupported_xml("selected glossary entry range is inverted"))?;
                        let mut matching_selectors = Vec::new();
                        for (selector_index, requested) in selectors.iter().enumerate() {
                            if glossary_selector_matches(
                                requested,
                                requested_name_keys[selector_index].as_deref(),
                                current_entry_index,
                                current_entry_name_key.as_deref(),
                                current_entry_id.as_ref().map(crate::glossary::Id::as_str),
                            ) {
                                reserve_one(&mut matching_selectors, "glossary selector matches")?;
                                matching_selectors.push(selector_index);
                            }
                        }
                        if !matching_selectors.is_empty() {
                            if entry_bytes > limits.max_entry_bytes {
                                return Err(Error::Limit {
                                    kind: LimitKind::SelectedEntryBytes,
                                    actual: entry_bytes,
                                    maximum: limits.max_entry_bytes,
                                });
                            }
                            let inner_start = current_body_inner_start.ok_or(unsupported_xml(
                                "selected glossary entry body start is missing",
                            ))?;
                            let inner_end = current_body_inner_end.ok_or(unsupported_xml(
                                "selected glossary entry body end is missing",
                            ))?;
                            let namespace_attributes = current_body_namespace_attributes
                                .take()
                                .ok_or(unsupported_xml(
                                    "selected glossary namespace context is missing",
                                ))?;
                            let paragraphs: Arc<[Range]> =
                                std::mem::take(&mut current_paragraphs).into();
                            let namespace_attributes: Arc<Vec<u8>> = namespace_attributes.into();
                            let word_namespace = word_namespace_shared
                                .as_ref()
                                .ok_or(unsupported_xml("glossary word namespace is missing"))?;
                            for selector_index in matching_selectors {
                                if matched[selector_index] {
                                    return Err(Error::AmbiguousGlossaryEntry {
                                        selector: selectors[selector_index].kind(),
                                    });
                                }
                                matched[selector_index] = true;
                                selected_layouts.push((
                                    selector_index,
                                    Layout {
                                        paragraphs: paragraphs.clone(),
                                        root_start_end: inner_start,
                                        root_end_start: inner_end,
                                        root_kind: RootKind::Glossary,
                                        namespace_attributes: namespace_attributes.clone(),
                                        word_namespace: word_namespace.clone(),
                                        has_w_namespace: current_body_has_w_namespace,
                                    },
                                ));
                            }
                        } else {
                            current_paragraphs.clear();
                        }
                    } else {
                        current_paragraphs.clear();
                    }
                    doc_part_depth = None;
                    current_entry_start = None;
                    current_entry_selected = false;
                    current_entry_pr_seen = false;
                    current_entry_name_key = None;
                    current_entry_id = None;
                    current_body_count = 0;
                    current_body_inner_start = None;
                    current_body_inner_end = None;
                    current_body_has_w_namespace = false;
                }
                if doc_parts_depth == Some(depth) {
                    doc_parts_depth = None;
                }
                if depth == 1 {
                    if root_closed {
                        return Err(unsupported_xml("glossary root closes more than once"));
                    }
                    root_closed = true;
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
                    "XML declarations are only accepted before the glossary root",
                ));
            },
            Event::Text(text)
                if depth == 0 && !text.as_ref().iter().all(u8::is_ascii_whitespace) =>
            {
                return Err(unsupported_xml(
                    "non-whitespace text occurs outside glossary root",
                ));
            },
            Event::Eof => break,
            _ => {},
        }
    }

    if !root_seen || !root_closed || depth != 0 || !namespace_stack.is_empty() {
        return Err(unsupported_xml("glossary XML root is incomplete"));
    }
    if !root_has_doc_parts {
        return Err(unsupported_xml("glossary root must contain docParts"));
    }
    for (selector_index, selector) in selectors.iter().enumerate() {
        if !matched[selector_index] {
            return Err(Error::MissingGlossaryEntry {
                selector: selector.kind(),
            });
        }
    }
    selected_layouts.sort_unstable_by_key(|(selector_index, _)| *selector_index);
    if selected_layouts.len() != selectors.len() {
        return Err(Error::TopologyConflict);
    }
    let mut layouts = Vec::new();
    layouts
        .try_reserve_exact(selected_layouts.len())
        .map_err(|source| Error::Allocation {
            resource: "glossary selected layouts",
            source,
        })?;
    for (selector_index, layout) in selected_layouts {
        if selector_index >= selectors.len() {
            return Err(Error::TopologyConflict);
        }
        layouts.push(layout);
    }
    check_execution_context(execution)?;
    admit_story_compatibility(xml, limits, execution)?;
    Ok(layouts)
}

fn validate_glossary_batch_selector_limits(
    selectors: &[GlossarySelector],
    limits: Limits,
) -> Result<()> {
    if selectors.is_empty() {
        return Err(Error::Limit {
            kind: LimitKind::GlossaryBatchSelectors,
            actual: 0,
            maximum: 1,
        });
    }
    if selectors.len() > limits.max_entries {
        return Err(Error::Limit {
            kind: LimitKind::GlossaryBatchSelectors,
            actual: selectors.len(),
            maximum: limits.max_entries,
        });
    }
    let mut total = 0usize;
    for selector in selectors {
        let bytes = match selector {
            GlossarySelector::ByName(name) => {
                if name.len() > MAX_GLOSSARY_SELECTOR_NAME_BYTES {
                    return Err(Error::Limit {
                        kind: LimitKind::GlossarySelectorNameBytes,
                        actual: name.len(),
                        maximum: MAX_GLOSSARY_SELECTOR_NAME_BYTES,
                    });
                }
                name.len()
            },
            GlossarySelector::ById(id) => {
                validate_glossary_selector_id(id)?;
                id.as_str().len()
            },
            GlossarySelector::ByNameAndId { name, id } => {
                if name.len() > MAX_GLOSSARY_SELECTOR_NAME_BYTES {
                    return Err(Error::Limit {
                        kind: LimitKind::GlossarySelectorNameBytes,
                        actual: name.len(),
                        maximum: MAX_GLOSSARY_SELECTOR_NAME_BYTES,
                    });
                }
                validate_glossary_selector_id(id)?;
                name.len()
                    .checked_add(id.as_str().len())
                    .ok_or(Error::Limit {
                        kind: LimitKind::GlossarySelectorBytes,
                        actual: usize::MAX,
                        maximum: limits.max_xml_bytes,
                    })?
            },
            GlossarySelector::ByIndex(_) => 0,
        };
        total = total.checked_add(bytes).ok_or(Error::Limit {
            kind: LimitKind::GlossarySelectorBytes,
            actual: usize::MAX,
            maximum: limits.max_xml_bytes,
        })?;
        if total > limits.max_xml_bytes {
            return Err(Error::Limit {
                kind: LimitKind::GlossarySelectorBytes,
                actual: total,
                maximum: limits.max_xml_bytes,
            });
        }
    }
    Ok(())
}

fn scan_story(
    xml: &[u8],
    selector: &Selector,
    expected_word_namespace: Option<&[u8]>,
    limits: Limits,
    execution: Option<&ExecutionContext>,
) -> Result<Layout> {
    if xml.len() > limits.max_xml_bytes {
        return Err(Error::Limit {
            kind: LimitKind::XmlBytes,
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
        Selector::Glossary(_) => {
            return Err(unsupported_xml(
                "glossary selectors require a glossary scan",
            ));
        },
    };
    let mut reader = Reader::from_reader(xml);
    let origin = ReaderOrigin::of(xml);
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
        let before = reader_position(&reader, origin)?;
        let event = reader.read_event().map_err(|error| {
            Error::Document(crate::Error::InvalidFormat(format!(
                "invalid story XML: {error}",
            )))
        })?;
        event_count = event_count.checked_add(1).ok_or(Error::Limit {
            kind: LimitKind::XmlEvents,
            actual: usize::MAX,
            maximum: limits.max_events,
        })?;
        if event_count > limits.max_events {
            return Err(Error::Limit {
                kind: LimitKind::XmlEvents,
                actual: event_count,
                maximum: limits.max_events,
            });
        }
        if event_count & 63 == 0 {
            check_execution_context(execution)?;
        }
        let after = reader_position(&reader, origin)?;
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
                    expected_word_namespace,
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
                        kind: LimitKind::Entries,
                        actual: usize::MAX,
                        maximum: limits.max_entries,
                    })?;
                    if entry_count > limits.max_entries {
                        return Err(Error::Limit {
                            kind: LimitKind::Entries,
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
                                kind: LimitKind::XmlDepth,
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
                        kind: LimitKind::BodyElements,
                        actual: usize::MAX,
                        maximum: 1,
                    })?;
                    if body_count > 1 {
                        return Err(unsupported_xml("main story contains multiple bodies"));
                    }
                    body_depth = Some(parent_depth.checked_add(1).ok_or(Error::Limit {
                        kind: LimitKind::XmlDepth,
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
                            RootKind::Glossary => false,
                        };
                if direct_paragraph {
                    ensure_paragraph_capacity(
                        paragraphs.len(),
                        paragraph_stack.len(),
                        limits.max_paragraphs,
                    )?;
                    reserve_one(&mut paragraphs, "story paragraph ranges")?;
                    let paragraph_depth = depth.checked_add(1).ok_or(Error::Limit {
                        kind: LimitKind::XmlDepth,
                        actual: usize::MAX,
                        maximum: limits.max_depth,
                    })?;
                    reserve_one(&mut paragraph_stack, "story paragraph stack")?;
                    paragraph_stack.push((paragraph_depth, event_start));
                }
                depth = depth.checked_add(1).ok_or(Error::Limit {
                    kind: LimitKind::XmlDepth,
                    actual: usize::MAX,
                    maximum: limits.max_depth,
                })?;
                if depth > limits.max_depth {
                    return Err(Error::Limit {
                        kind: LimitKind::XmlDepth,
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
                    kind: LimitKind::XmlDepth,
                    actual: usize::MAX,
                    maximum: limits.max_depth,
                })?;
                if logical_depth > limits.max_depth {
                    return Err(Error::Limit {
                        kind: LimitKind::XmlDepth,
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
                    expected_word_namespace,
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
                            kind: LimitKind::Entries,
                            actual: usize::MAX,
                            maximum: limits.max_entries,
                        })?;
                        if entry_count > limits.max_entries {
                            return Err(Error::Limit {
                                kind: LimitKind::Entries,
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
                                    kind: LimitKind::SelectedEntryBytes,
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
                            kind: LimitKind::BodyElements,
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
                                RootKind::Glossary => false,
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
                                kind: LimitKind::Paragraphs,
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
                            kind: LimitKind::SelectedEntryBytes,
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
    admit_story_compatibility(xml, limits, execution)?;
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
        RootKind::Glossary => b"glossaryDocument".as_slice(),
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
                kind: LimitKind::WrappedStoryXmlBytes,
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
            kind: LimitKind::WrappedStoryXmlBytes,
            actual: usize::MAX,
            maximum,
        })?;
    let suffix = b"</w:body></w:document>";
    let total = prefix_len
        .checked_add(inner.len())
        .and_then(|value| value.checked_add(suffix.len()))
        .ok_or(Error::Limit {
            kind: LimitKind::WrappedStoryXmlBytes,
            actual: usize::MAX,
            maximum,
        })?;
    if total > maximum {
        return Err(Error::Limit {
            kind: LimitKind::WrappedStoryXmlBytes,
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

/// Source edits keep raw bytes. Prove that their Word content agrees with
/// the reader's MCE projection instead of removing compatibility markup.
/// Run after the lexical scan has enforced the caller's namespace, XML,
/// event and entry bounds. The proof builds no projected document.
fn admit_story_compatibility(
    xml: &[u8],
    limits: Limits,
    execution: Option<&ExecutionContext>,
) -> Result<()> {
    use litchi_ooxml_common::mce;
    if !mce::mentions_namespace(xml) {
        return Ok(());
    }
    let proof_limits = mce::Limits {
        max_input_bytes: limits.max_xml_bytes,
        // The proof checks a conservative serialization bound, but emits
        // no bytes; this is not the story's caller-owned text-output budget.
        max_output_bytes: limits.max_xml_bytes.saturating_mul(8),
        max_depth: limits.max_depth,
        max_namespace_bindings: limits.max_namespace_bindings,
        ..mce::Limits::default()
    };
    let verdict = mce::admission::admit_with(
        xml,
        &crate::parts::document_part::visible_document_capabilities(),
        &proof_limits,
        |_| check_execution_context(execution),
    )?;
    verdict.map(|_| ()).map_err(|_| {
        unsupported_xml(
            "markup-compatibility processing would change the selected story's understood content",
        )
    })
}

fn is_word_element(
    name: QName<'_>,
    context: &NamespaceContext,
    word_namespace: &[u8],
    local: &[u8],
) -> bool {
    name.local_name().as_ref() == local && context.resolve_element(name) == Some(word_namespace)
}

fn is_glossary_root_extension(
    name: QName<'_>,
    context: &NamespaceContext,
    word_namespace: &[u8],
) -> bool {
    context
        .resolve_element(name)
        .is_some_and(|namespace| !namespace.is_empty() && namespace != word_namespace)
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
    if (word_namespace == STRICT_WORD_NAMESPACE
        && namespace == Some(TRANSITIONAL_RELATIONSHIP_NAMESPACE))
        || (word_namespace == TRANSITIONAL_WORD_NAMESPACE
            && namespace == Some(STRICT_RELATIONSHIP_NAMESPACE))
    {
        return Err(unsupported_xml(
            "glossary content mixes Strict and Transitional relationship namespaces",
        ));
    }
    if word_namespace == STRICT_WORD_NAMESPACE && namespace == Some(VML_NAMESPACE) {
        return Err(unsupported_xml(
            "Strict glossary content cannot contain VML",
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
            | b"glossaryDocument"
            | b"background"
            | b"docParts"
            | b"docPart"
            | b"docPartPr"
            | b"docPartBody"
            | b"name"
            | b"guid"
            | b"p"
    ) && namespace != Some(word_namespace)
    {
        return Err(unsupported_xml(
            "WordprocessingML element has an incorrect namespace binding",
        ));
    }
    Ok(())
}

fn reject_word_dialect_attributes(
    start: &BytesStart<'_>,
    context: &NamespaceContext,
    word_namespace: &[u8],
) -> Result<()> {
    for attribute in start.checked_attributes() {
        let attribute = attribute.map_err(|error| {
            Error::Document(crate::Error::Xml(format!(
                "invalid story attribute: {error}"
            )))
        })?;
        if is_namespace_declaration(attribute.key) {
            continue;
        }
        let Some(prefix) = attribute.key.prefix() else {
            continue;
        };
        let Some(namespace) = context.resolve_prefix(prefix.as_ref()) else {
            continue;
        };
        if (word_namespace == STRICT_WORD_NAMESPACE
            && namespace == TRANSITIONAL_RELATIONSHIP_NAMESPACE)
            || (word_namespace == TRANSITIONAL_WORD_NAMESPACE
                && namespace == STRICT_RELATIONSHIP_NAMESPACE)
        {
            return Err(unsupported_xml(
                "glossary content mixes Strict and Transitional relationship attributes",
            ));
        }
        if word_namespace == STRICT_WORD_NAMESPACE && namespace == VML_NAMESPACE {
            return Err(unsupported_xml(
                "Strict glossary content cannot contain VML attributes",
            ));
        }
    }
    Ok(())
}

fn unsupported_xml(reason: &'static str) -> Error {
    Error::Document(crate::Error::refused(
        litchi_core::edit::RefusalKind::UnsupportedContent,
        "source-backed story text",
        reason,
    ))
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
        kind: LimitKind::Paragraphs,
        actual: usize::MAX,
        maximum,
    })?;
    if actual >= maximum {
        return Err(Error::Limit {
            kind: LimitKind::Paragraphs,
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

fn digest(bytes: &[u8]) -> [u8; 32] {
    Sha256::digest(bytes).into()
}

fn same_part_name(left: &PackURI, right: &PackURI) -> bool {
    left.as_str().eq_ignore_ascii_case(right.as_str())
}

fn same_content(
    actual_fingerprint: [u8; 32],
    actual: &[u8],
    expected: &[u8],
    expected_fingerprint: [u8; 32],
) -> bool {
    // Snapshots are immutable and their fingerprint was computed when the
    // source was admitted. Reusing that authenticated value avoids hashing
    // the same payload again on every owner publication.
    actual.len() == expected.len()
        && actual_fingerprint == expected_fingerprint
        && actual == expected
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

const STRICT_HEADER: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships/header";
const STRICT_FOOTER: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships/footer";
const STRICT_FOOTNOTES: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships/footnotes";
const STRICT_ENDNOTES: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships/endnotes";
const TRANSITIONAL_GLOSSARY: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/glossaryDocument";
const STRICT_GLOSSARY: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships/glossaryDocument";

#[derive(Clone)]
struct GlossaryBatchEntry {
    selector_index: usize,
    layout: Layout,
}

/// An immutable source-backed batch of existing glossary-entry text stories.
///
/// Entries are resolved atomically in one glossary Part and retained in
/// canonical source order. The payload is shared once by the whole batch;
/// entry records contain only bounded ranges, selectors, and namespace
/// metadata rather than duplicate XML buffers.
#[derive(Clone)]
pub struct GlossaryBatchSnapshot {
    selector: GlossaryBatchSelector,
    entries: Vec<GlossaryBatchEntry>,
    payload: Payload,
    partname: PackURI,
    source_version: SourceVersion,
    source_lineage: SourceLineage,
    part_fingerprint: [u8; 32],
    artifact_fingerprint: Option<SourceArtifactFingerprint>,
    limits: Limits,
    execution: Option<ExecutionContext>,
}

impl GlossaryBatchSnapshot {
    /// Borrow the canonical source-order selector set.
    #[must_use]
    pub const fn selector(&self) -> &GlossaryBatchSelector {
        &self.selector
    }

    /// Return the number of selected glossary entries.
    pub fn entry_count(&self) -> Result<usize> {
        self.check_execution()?;
        Ok(self.entries.len())
    }

    /// Borrow one canonical source-order selector.
    pub fn entry_selector(&self, index: usize) -> Option<&GlossarySelector> {
        let entry = self.entries.get(index)?;
        self.selector.selectors.get(entry.selector_index)
    }

    /// Return the source revision captured by this batch.
    #[must_use]
    pub const fn source_version(&self) -> SourceVersion {
        self.source_version
    }

    /// Return the SHA-256 digest of the complete selected glossary Part.
    #[must_use]
    pub const fn part_fingerprint(&self) -> [u8; 32] {
        self.part_fingerprint
    }

    /// Return the policy under which this batch was admitted.
    #[must_use]
    pub const fn limits_value(&self) -> Limits {
        self.limits
    }

    /// Return the direct paragraph count for one selected entry.
    pub fn paragraph_count(&self, entry_index: usize) -> Result<Option<usize>> {
        self.check_execution()?;
        let Some(entry) = self.entries.get(entry_index) else {
            return Ok(None);
        };
        Ok(Some(entry.layout.paragraphs.len()))
    }

    /// Return one direct paragraph from one selected entry.
    pub fn paragraph_text(
        &self,
        entry_index: usize,
        paragraph_index: usize,
    ) -> Result<Option<String>> {
        self.check_execution()?;
        let Some(entry) = self.entries.get(entry_index) else {
            return Ok(None);
        };
        let Some(range) = entry.layout.paragraphs.get(paragraph_index).copied() else {
            return Ok(None);
        };
        let bytes = self
            .raw_bytes()
            .get(range.start..range.end)
            .ok_or_else(|| {
                Error::Document(crate::Error::InvalidFormat(
                    "glossary batch paragraph range is outside the selected XML".into(),
                ))
            })?;
        let mut text = String::new();
        extract_story_paragraph_text(
            bytes,
            &mut text,
            self.limits.max_output_bytes,
            LimitKind::GlossaryBatchParagraphText,
        )?;
        self.check_execution()?;
        Ok(Some(text))
    }

    /// Extract direct paragraph text from one selected entry.
    pub fn extract_text(&self, entry_index: usize) -> Result<Option<String>> {
        self.check_execution()?;
        let Some(entry) = self.entries.get(entry_index) else {
            return Ok(None);
        };
        let mut output = String::new();
        for (index, range) in entry.layout.paragraphs.iter().copied().enumerate() {
            let bytes = self
                .raw_bytes()
                .get(range.start..range.end)
                .ok_or_else(|| {
                    Error::Document(crate::Error::InvalidFormat(
                        "glossary batch paragraph range is outside the selected XML".into(),
                    ))
                })?;
            extract_story_paragraph_text(
                bytes,
                &mut output,
                self.limits.max_output_bytes,
                LimitKind::GlossaryBatchTextOutput,
            )?;
            if index % 64 == 0 {
                self.check_execution()?;
            }
        }
        self.check_execution()?;
        Ok(Some(output))
    }

    /// Open the selected entries as a bounded mutable raw codec model.
    ///
    /// The supplied policy is intersected with the policy retained by this
    /// admitted snapshot. The model publishes a raw
    /// [`GlossaryBatchPatch`] and never a detached package owner.
    pub fn raw_tree(&self, limits: Limits) -> Result<raw::GlossaryTree> {
        let base = self.with_raw_limits(limits)?;
        Ok(raw::GlossaryTree::from_engine(base.edit_engine()?))
    }

    pub(crate) fn edit_engine(&self) -> Result<GlossaryBatchEdit> {
        self.check_execution()?;
        if self.payload.is_managed() {
            return Err(Error::Document(crate::Error::refused(
                litchi_core::edit::RefusalKind::UnsupportedOperation,
                "source-backed glossary batch text edit",
                "managed source-backed story edits require an owned edit snapshot; use an unmanaged compatibility constructor",
            )));
        }
        Ok(GlossaryBatchEdit {
            base: self.clone(),
            intents: Vec::new(),
            replacements: 0,
            replacement_text_bytes: 0,
        })
    }

    fn check_execution(&self) -> Result<()> {
        check_execution_context(self.execution.as_ref())
    }

    fn raw_bytes(&self) -> &[u8] {
        self.payload.as_bytes()
    }

    fn with_raw_limits(&self, requested: Limits) -> Result<Self> {
        let limits = self.limits.intersect(requested);
        if limits == self.limits {
            return Ok(self.clone());
        }
        let expected_word_namespace = self
            .entries
            .first()
            .map(|entry| entry.layout.word_namespace.as_ref());
        let layouts = scan_glossary_many(
            self.raw_bytes(),
            &self.selector.selectors,
            expected_word_namespace,
            limits,
            self.execution.as_ref(),
        )?;
        if layouts.len() != self.entries.len() {
            return Err(Error::TopologyConflict);
        }
        let mut entries = Vec::new();
        entries
            .try_reserve_exact(layouts.len())
            .map_err(|source| Error::Allocation {
                resource: "glossary raw tree entries",
                source,
            })?;
        for (selector_index, layout) in layouts.into_iter().enumerate() {
            entries.push(GlossaryBatchEntry {
                selector_index,
                layout,
            });
        }
        entries.sort_unstable_by_key(|entry| entry.layout.root_start_end);
        validate_batch_ranges(&entries, self.raw_bytes().len())?;
        for (canonical_index, entry) in entries.iter_mut().enumerate() {
            entry.selector_index = canonical_index;
        }
        let mut base = self.clone();
        base.entries = entries;
        base.limits = limits;
        Ok(base)
    }
}

#[derive(Clone)]
struct GlossaryBatchEditIntent {
    entry_index: usize,
    position: Position,
    text: String,
}

/// A staged atomic edit of a bounded glossary-entry batch.
///
/// Only compact replacement intents are retained. During projection one
/// selected entry is wrapped and edited at a time, so the batch does not hold
/// one full document wrapper or XML payload per selected entry.
pub(crate) struct GlossaryBatchEdit {
    base: GlossaryBatchSnapshot,
    intents: Vec<GlossaryBatchEditIntent>,
    replacements: usize,
    replacement_text_bytes: usize,
}

impl GlossaryBatchEdit {
    /// Borrow the immutable batch source snapshot.
    #[must_use]
    pub(crate) const fn source(&self) -> &GlossaryBatchSnapshot {
        &self.base
    }

    /// Return the current projected batch after atomically validating intents.
    pub(crate) fn projected(&self) -> Result<GlossaryBatchSnapshot> {
        self.base.check_execution()?;
        let target = assemble_batch_snapshot(&self.base, &self.intents)?;
        self.base.check_execution()?;
        Ok(target)
    }

    /// Replace all text in one direct paragraph of one selected entry.
    ///
    /// A paragraph may be addressed only once in one staged batch. This
    /// removes overlapping intent ambiguity and leaves every failed operation
    /// free to return without changing counters or stored intents.
    pub(crate) fn replace_paragraph_text(
        &mut self,
        entry_index: usize,
        position: Position,
        authored_text: impl AsRef<str>,
    ) -> Result<&mut Self> {
        self.base.check_execution()?;
        let Some(entry) = self.base.entries.get(entry_index) else {
            return Err(batch_entry_index_error(
                entry_index,
                self.base.entries.len(),
            ));
        };
        let paragraph_index = position.get();
        if paragraph_index >= entry.layout.paragraphs.len() {
            return Err(batch_paragraph_index_error(
                entry_index,
                paragraph_index,
                entry.layout.paragraphs.len(),
            ));
        }
        let insert_at = match self.intents.binary_search_by(|intent| {
            intent
                .entry_index
                .cmp(&entry_index)
                .then_with(|| intent.position.cmp(&position))
        }) {
            Ok(_) => {
                return Err(Error::RepeatedParagraphEdit {
                    entry_index,
                    paragraph_index,
                });
            },
            Err(index) => index,
        };
        let next_replacements = self.replacements.checked_add(1).ok_or(Error::Limit {
            kind: LimitKind::Replacements,
            actual: usize::MAX,
            maximum: self.base.limits.max_replacements,
        })?;
        if next_replacements > self.base.limits.max_replacements {
            return Err(Error::Limit {
                kind: LimitKind::Replacements,
                actual: next_replacements,
                maximum: self.base.limits.max_replacements,
            });
        }
        let authored_text = authored_text.as_ref();
        let next_bytes = self
            .replacement_text_bytes
            .checked_add(authored_text.len())
            .ok_or(Error::Limit {
                kind: LimitKind::ReplacementTextBytes,
                actual: usize::MAX,
                maximum: self.base.limits.max_replacement_text_bytes,
            })?;
        if next_bytes > self.base.limits.max_replacement_text_bytes {
            return Err(Error::Limit {
                kind: LimitKind::ReplacementTextBytes,
                actual: next_bytes,
                maximum: self.base.limits.max_replacement_text_bytes,
            });
        }
        validate_paragraph_text(position, authored_text).map_err(Error::DocumentEdit)?;
        self.intents
            .try_reserve(1)
            .map_err(|source| Error::Allocation {
                resource: "glossary batch edit intents",
                source,
            })?;
        let text = checked_string_clone(authored_text, "glossary batch replacement text")?;
        self.intents.insert(
            insert_at,
            GlossaryBatchEditIntent {
                entry_index,
                position,
                text,
            },
        );
        if let Err(error) = self.base.check_execution() {
            self.intents.remove(insert_at);
            return Err(error);
        }
        self.replacements = next_replacements;
        self.replacement_text_bytes = next_bytes;
        Ok(self)
    }

    pub(crate) fn finish(self) -> Result<GlossaryBatchPatch> {
        self.base.check_execution()?;
        let target = self.projected()?;
        self.base.check_execution()?;
        Ok(GlossaryBatchPatch {
            before: self.base,
            after: target,
        })
    }
}

/// A reversible patch for a glossary Part containing a bounded set of selected
/// entries. Owner publication adds its complete-artifact and source-freshness
/// proof; raw replay matches the selected content.
#[derive(Clone)]
pub struct GlossaryBatchPatch {
    before: GlossaryBatchSnapshot,
    after: GlossaryBatchSnapshot,
}

impl GlossaryBatchPatch {
    /// Borrow the before batch snapshot represented by this patch.
    #[must_use]
    pub const fn source(&self) -> &GlossaryBatchSnapshot {
        &self.before
    }

    /// Borrow the projected target represented by this patch.
    #[must_use]
    pub const fn target(&self) -> &GlossaryBatchSnapshot {
        &self.after
    }

    /// Return whether the complete glossary Part is unchanged.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.before.raw_bytes() == self.after.raw_bytes()
    }

    /// Return the exact inverse batch patch.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            before: self.after.clone(),
            after: self.before.clone(),
        }
    }

    /// Apply only to an equivalent source-backed batch snapshot.
    /// [`Patch::apply_content`] for a glossary batch.
    pub(crate) fn apply_content(
        &self,
        source: &GlossaryBatchSnapshot,
    ) -> Result<GlossaryBatchSnapshot> {
        self.check_content(source)?;
        Ok(if self.is_noop() {
            source.clone()
        } else {
            self.after.clone()
        })
    }

    fn check_content(&self, source: &GlossaryBatchSnapshot) -> Result<()> {
        if self.before.selector != source.selector
            || !same_part_name(&self.before.partname, &source.partname)
        {
            return Err(Error::TopologyConflict);
        }
        if !same_content(
            source.part_fingerprint,
            source.raw_bytes(),
            self.before.raw_bytes(),
            self.before.part_fingerprint,
        ) {
            return Err(Error::PatchConflict);
        }
        Ok(())
    }

    /// Replay this raw glossary patch on an independently opened matching
    /// batch snapshot. Replay is scoped to the selected glossary XML part;
    /// package publication performs complete-artifact and source-freshness
    /// checks separately.
    pub fn apply(&self, source: &GlossaryBatchSnapshot) -> Result<GlossaryBatchSnapshot> {
        source.check_execution()?;
        self.check_content(source)?;
        self.rebased_target(source)
    }

    fn rebased_target(&self, source: &GlossaryBatchSnapshot) -> Result<GlossaryBatchSnapshot> {
        if self.is_noop() {
            return Ok(source.clone());
        }
        let bytes = self.after.raw_bytes();
        if bytes.len() > source.limits.max_output_bytes {
            return Err(Error::Limit {
                kind: LimitKind::GlossaryBatchOutputBytes,
                actual: bytes.len(),
                maximum: source.limits.max_output_bytes,
            });
        }
        let expected_word_namespace = source
            .entries
            .first()
            .map(|entry| entry.layout.word_namespace.as_ref());
        let layouts = scan_glossary_many(
            bytes,
            &source.selector.selectors,
            expected_word_namespace,
            source.limits,
            source.execution.as_ref(),
        )?;
        if layouts.len() != source.entries.len() {
            return Err(Error::TopologyConflict);
        }
        let mut entries = Vec::new();
        entries
            .try_reserve_exact(layouts.len())
            .map_err(|source| Error::Allocation {
                resource: "glossary raw replay entries",
                source,
            })?;
        for (selector_index, layout) in layouts.into_iter().enumerate() {
            entries.push(GlossaryBatchEntry {
                selector_index,
                layout,
            });
        }
        entries.sort_unstable_by_key(|entry| entry.layout.root_start_end);
        validate_batch_ranges(&entries, bytes.len())?;
        for (canonical_index, entry) in entries.iter_mut().enumerate() {
            entry.selector_index = canonical_index;
        }
        let mut target = self.after.clone();
        target.selector = source.selector.clone();
        target.entries = entries;
        target.partname = source.partname.clone();
        target.source_version = source.source_version;
        target.source_lineage = source.source_lineage.clone();
        target.part_fingerprint = digest(bytes);
        target.artifact_fingerprint = source.artifact_fingerprint;
        target.limits = source.limits;
        target.execution = source.execution.clone();
        Ok(target)
    }
}

impl Package {
    /// Capture a bounded batch of existing glossary-entry text stories.
    pub fn glossary_text_batch_snapshot(
        &self,
        selector: GlossaryBatchSelector,
    ) -> Result<GlossaryBatchSnapshot> {
        self.glossary_text_batch_snapshot_bounded(selector, self.semantic_limits.story_text_value())
    }

    /// Capture a glossary-entry batch under explicit finite bounds.
    pub(crate) fn glossary_text_batch_snapshot_bounded(
        &self,
        selector: GlossaryBatchSelector,
        limits: Limits,
    ) -> Result<GlossaryBatchSnapshot> {
        capture_glossary_batch(self, selector, limits)
    }

    /// Check a glossary batch patch against this package and return the part
    /// and its rewritten bytes (record 0355a).
    pub(crate) fn glossary_text_batch_target(
        &self,
        patch: &GlossaryBatchPatch,
    ) -> Result<(PackURI, Arc<Vec<u8>>)> {
        self.trim_clean_paragraph_index_for_operation();
        self.physical_package()?.check_execution()?;
        let recipient_limits = self.semantic_limits.story_text_value();
        let current =
            capture_glossary_batch(self, patch.before.selector.clone(), recipient_limits)?;
        let target = patch.apply_content(&current)?;
        story_replacement(
            self,
            patch.is_noop(),
            &current.partname,
            target.raw_bytes(),
            recipient_limits,
            LimitKind::GlossaryBatchOutputBytes,
            "glossary batch publication replacement",
        )
    }
}

fn capture_glossary_batch(
    package: &Package,
    selector: GlossaryBatchSelector,
    limits: Limits,
) -> Result<GlossaryBatchSnapshot> {
    package.trim_clean_paragraph_index_for_operation();
    package.physical_package()?.check_execution()?;
    validate_glossary_batch_selector_limits(&selector.selectors, limits)?;

    let source_version = package.physical_package()?.source_version()?;
    let source_lineage = package.physical_package()?.source_lineage();
    let execution = package.physical_package()?.execution_context();
    let first_selector = Selector::Glossary(
        selector
            .selectors
            .first()
            .ok_or(Error::TopologyConflict)?
            .clone(),
    );
    // Resolve the package topology and materialize the glossary Part once.
    // One semantic inventory scan resolves every selector against those bytes.
    let (partname, part, relationship_word_namespace) =
        resolve_part(package.physical_package()?, &first_selector)?;
    let declared = part.declared_uncompressed_size()?;
    if declared > limits.max_xml_bytes as u64 {
        return Err(Error::Limit {
            kind: LimitKind::XmlBytes,
            actual: usize::try_from(declared).unwrap_or(usize::MAX),
            maximum: limits.max_xml_bytes,
        });
    }
    let data = part.data()?;
    if data.as_bytes().len() > limits.max_xml_bytes {
        return Err(Error::Limit {
            kind: LimitKind::XmlBytes,
            actual: data.as_bytes().len(),
            maximum: limits.max_xml_bytes,
        });
    }
    let managed = package
        .physical_package()?
        .cache_diagnostics()
        .budget_managed;
    let payload = if managed {
        Payload::Managed(data)
    } else {
        Payload::Owned(data.into_arc()?)
    };
    let part_fingerprint = digest(payload.as_bytes());
    let layouts = scan_glossary_many(
        payload.as_bytes(),
        &selector.selectors,
        relationship_word_namespace,
        limits,
        execution.as_ref(),
    )?;
    if layouts.len() != selector.selectors.len() {
        return Err(Error::TopologyConflict);
    }
    let mut entries = Vec::new();
    entries
        .try_reserve_exact(selector.selectors.len())
        .map_err(|source| Error::Allocation {
            resource: "glossary batch entries",
            source,
        })?;
    let mut selected_bytes = 0usize;
    for (selector_index, layout) in layouts.into_iter().enumerate() {
        let entry_bytes = layout
            .root_end_start
            .checked_sub(layout.root_start_end)
            .ok_or(Error::TopologyConflict)?;
        selected_bytes = selected_bytes
            .checked_add(entry_bytes)
            .ok_or(Error::Limit {
                kind: LimitKind::GlossaryBatchEntryBytes,
                actual: usize::MAX,
                maximum: limits.max_output_bytes,
            })?;
        if selected_bytes > limits.max_output_bytes {
            return Err(Error::Limit {
                kind: LimitKind::GlossaryBatchEntryBytes,
                actual: selected_bytes,
                maximum: limits.max_output_bytes,
            });
        }
        entries.push(GlossaryBatchEntry {
            selector_index,
            layout,
        });
    }
    let observed = package.physical_package()?.source_version()?;
    if observed != source_version {
        return Err(Error::Document(crate::Error::Opc(
            litchi_opc::OpcError::SourceChanged {
                expected: source_version,
                actual: observed,
            },
        )));
    }
    package.physical_package()?.check_execution()?;

    entries.sort_unstable_by_key(|entry| entry.layout.root_start_end);
    validate_batch_ranges(&entries, payload.as_bytes().len())?;
    let mut canonical = Vec::new();
    canonical
        .try_reserve_exact(entries.len())
        .map_err(|source| Error::Allocation {
            resource: "glossary batch selectors",
            source,
        })?;
    for (canonical_index, entry) in entries.iter_mut().enumerate() {
        let requested = selector
            .selectors
            .get(entry.selector_index)
            .ok_or(Error::TopologyConflict)?;
        canonical.push(requested.clone());
        entry.selector_index = canonical_index;
    }
    Ok(GlossaryBatchSnapshot {
        selector: GlossaryBatchSelector {
            selectors: canonical,
        },
        entries,
        payload,
        partname,
        source_version,
        source_lineage,
        part_fingerprint,
        artifact_fingerprint: None,
        limits,
        execution,
    })
}

struct GlossaryBatchMeasurement {
    entry_index: usize,
    intent_start: usize,
    intent_end: usize,
    range: Range,
    length: usize,
}

struct GlossaryBatchProjection {
    range: Range,
    length: usize,
    bytes: Option<Vec<u8>>,
}

struct GlossaryBatchReplacement {
    range: Range,
    bytes: Vec<u8>,
}

fn assemble_batch_snapshot(
    base: &GlossaryBatchSnapshot,
    intents: &[GlossaryBatchEditIntent],
) -> Result<GlossaryBatchSnapshot> {
    base.check_execution()?;
    if intents.len() > base.limits.max_replacements {
        return Err(Error::Limit {
            kind: LimitKind::Replacements,
            actual: intents.len(),
            maximum: base.limits.max_replacements,
        });
    }
    if base.raw_bytes().len() > base.limits.max_output_bytes {
        return Err(Error::Limit {
            kind: LimitKind::GlossaryBatchOutputBytes,
            actual: base.raw_bytes().len(),
            maximum: base.limits.max_output_bytes,
        });
    }
    let mut authored_bytes = 0usize;
    for intent in intents {
        authored_bytes = authored_bytes
            .checked_add(intent.text.len())
            .ok_or(Error::Limit {
                kind: LimitKind::GlossaryBatchTransientBytes,
                actual: usize::MAX,
                maximum: base.limits.max_output_bytes,
            })?;
    }
    let transient_bound = authored_bytes
        .checked_mul(6)
        .and_then(|value| base.raw_bytes().len().checked_add(value))
        .ok_or(Error::Limit {
            kind: LimitKind::GlossaryBatchTransientBytes,
            actual: usize::MAX,
            maximum: base.limits.max_output_bytes,
        })?;
    if transient_bound > base.limits.max_output_bytes {
        return Err(Error::Limit {
            kind: LimitKind::GlossaryBatchTransientBytes,
            actual: transient_bound,
            maximum: base.limits.max_output_bytes,
        });
    }
    let mut measurements = Vec::new();
    measurements
        .try_reserve_exact(intents.len())
        .map_err(|source| Error::Allocation {
            resource: "glossary batch replacement measurements",
            source,
        })?;
    let mut aggregate_replacement_bytes = 0usize;
    let mut intent_start = 0usize;
    for entry_index in 0..base.entries.len() {
        let group_start = intent_start;
        while intents
            .get(intent_start)
            .is_some_and(|intent| intent.entry_index == entry_index)
        {
            intent_start += 1;
        }
        if group_start == intent_start {
            continue;
        }
        base.check_execution()?;
        let projection = project_batch_entry(
            base,
            entry_index,
            intents
                .get(group_start..intent_start)
                .ok_or(Error::TopologyConflict)?,
            false,
        )?;
        aggregate_replacement_bytes = aggregate_replacement_bytes
            .checked_add(projection.length)
            .ok_or(Error::Limit {
                kind: LimitKind::GlossaryBatchReplacementBytes,
                actual: usize::MAX,
                maximum: base.limits.max_output_bytes,
            })?;
        if aggregate_replacement_bytes > base.limits.max_output_bytes {
            return Err(Error::Limit {
                kind: LimitKind::GlossaryBatchReplacementBytes,
                actual: aggregate_replacement_bytes,
                maximum: base.limits.max_output_bytes,
            });
        }
        if projection.bytes.is_some() {
            return Err(Error::TopologyConflict);
        }
        measurements.push(GlossaryBatchMeasurement {
            entry_index,
            intent_start: group_start,
            intent_end: intent_start,
            range: projection.range,
            length: projection.length,
        });
    }
    if intent_start != intents.len() {
        return Err(Error::TopologyConflict);
    }

    let mut final_len = base.raw_bytes().len();
    for measurement in &measurements {
        if measurement.range.start > measurement.range.end
            || measurement.range.end > base.raw_bytes().len()
        {
            return Err(Error::TopologyConflict);
        }
        let old_len = measurement.range.end - measurement.range.start;
        final_len = final_len
            .checked_sub(old_len)
            .and_then(|value| value.checked_add(measurement.length))
            .ok_or(Error::Limit {
                kind: LimitKind::GlossaryBatchOutputBytes,
                actual: usize::MAX,
                maximum: base.limits.max_output_bytes,
            })?;
    }
    if final_len > base.limits.max_output_bytes {
        return Err(Error::Limit {
            kind: LimitKind::GlossaryBatchOutputBytes,
            actual: final_len,
            maximum: base.limits.max_output_bytes,
        });
    }

    let mut replacements = Vec::new();
    replacements
        .try_reserve_exact(measurements.len())
        .map_err(|source| Error::Allocation {
            resource: "glossary batch replacements",
            source,
        })?;
    for measurement in measurements {
        base.check_execution()?;
        let intent_group = intents
            .get(measurement.intent_start..measurement.intent_end)
            .ok_or(Error::TopologyConflict)?;
        let projection = project_batch_entry(base, measurement.entry_index, intent_group, true)?;
        let Some(bytes) = projection.bytes else {
            return Err(Error::TopologyConflict);
        };
        if bytes.len() != measurement.length || projection.range != measurement.range {
            return Err(Error::TopologyConflict);
        }
        replacements.push(GlossaryBatchReplacement {
            range: projection.range,
            bytes,
        });
    }

    let payload = if replacements.is_empty() {
        base.payload.clone()
    } else {
        let mut raw = Vec::new();
        raw.try_reserve_exact(final_len)
            .map_err(|source| Error::Allocation {
                resource: "glossary batch output bytes",
                source,
            })?;
        let mut cursor = 0usize;
        for replacement in &replacements {
            if replacement.range.start < cursor
                || replacement.range.start > replacement.range.end
                || replacement.range.end > base.raw_bytes().len()
            {
                return Err(Error::TopologyConflict);
            }
            let untouched = base
                .raw_bytes()
                .get(cursor..replacement.range.start)
                .ok_or(Error::TopologyConflict)?;
            raw.extend_from_slice(untouched);
            raw.extend_from_slice(&replacement.bytes);
            cursor = replacement.range.end;
        }
        let suffix = base
            .raw_bytes()
            .get(cursor..)
            .ok_or(Error::TopologyConflict)?;
        raw.extend_from_slice(suffix);
        if raw.len() != final_len {
            return Err(Error::TopologyConflict);
        }
        Payload::Owned(Arc::new(raw))
    };

    let first_entry = base.entries.first().ok_or(Error::TopologyConflict)?;
    let expected_word_namespace = first_entry.layout.word_namespace.as_ref();
    let final_layouts = scan_glossary_many(
        payload.as_bytes(),
        &base.selector.selectors,
        Some(expected_word_namespace),
        base.limits,
        base.execution.as_ref(),
    )?;
    if final_layouts.len() != base.entries.len() {
        return Err(Error::TopologyConflict);
    }
    let mut target_entries = Vec::new();
    target_entries
        .try_reserve_exact(base.entries.len())
        .map_err(|source| Error::Allocation {
            resource: "glossary batch target entries",
            source,
        })?;
    for entry in &base.entries {
        base.check_execution()?;
        let layout = final_layouts
            .get(entry.selector_index)
            .ok_or(Error::TopologyConflict)?
            .clone();
        target_entries.push(GlossaryBatchEntry {
            selector_index: entry.selector_index,
            layout,
        });
    }
    target_entries.sort_unstable_by_key(|entry| entry.layout.root_start_end);
    validate_batch_ranges(&target_entries, payload.as_bytes().len())?;
    let part_fingerprint = digest(payload.as_bytes());
    base.check_execution()?;
    Ok(GlossaryBatchSnapshot {
        selector: base.selector.clone(),
        entries: target_entries,
        payload,
        partname: base.partname.clone(),
        source_version: base.source_version,
        source_lineage: base.source_lineage.clone(),
        part_fingerprint,
        artifact_fingerprint: base.artifact_fingerprint,
        limits: base.limits,
        execution: base.execution.clone(),
    })
}

fn project_batch_entry(
    base: &GlossaryBatchSnapshot,
    entry_index: usize,
    intents: &[GlossaryBatchEditIntent],
    materialize: bool,
) -> Result<GlossaryBatchProjection> {
    let entry = base
        .entries
        .get(entry_index)
        .ok_or_else(|| batch_entry_index_error(entry_index, base.entries.len()))?;
    let range = Range {
        start: entry.layout.root_start_end,
        end: entry.layout.root_end_start,
    };
    let source_len = range
        .end
        .checked_sub(range.start)
        .ok_or(Error::TopologyConflict)?;
    let mut authored_bytes = 0usize;
    for intent in intents {
        if intent.entry_index != entry_index {
            return Err(Error::TopologyConflict);
        }
        authored_bytes = authored_bytes
            .checked_add(intent.text.len())
            .ok_or(Error::Limit {
                kind: LimitKind::GlossaryBatchTransientBytes,
                actual: usize::MAX,
                maximum: base.limits.max_output_bytes,
            })?;
    }
    let transient_estimate = source_len.checked_add(authored_bytes).ok_or(Error::Limit {
        kind: LimitKind::GlossaryBatchTransientBytes,
        actual: usize::MAX,
        maximum: base.limits.max_output_bytes,
    })?;
    if transient_estimate > base.limits.max_output_bytes {
        return Err(Error::Limit {
            kind: LimitKind::GlossaryBatchTransientBytes,
            actual: transient_estimate,
            maximum: base.limits.max_output_bytes,
        });
    }
    let effective_output_limit = base.limits.max_output_bytes.min(MAX_DOCUMENT_XML_BYTES);
    let (document_xml, envelope) =
        make_wrapped_document(base.raw_bytes(), &entry.layout, effective_output_limit)?;
    let document = DocumentSnapshot::from_xml(document_xml)?;
    let mut projected = document.edit();
    for intent in intents {
        if intent.entry_index != entry_index {
            return Err(Error::TopologyConflict);
        }
        base.check_execution()?;
        validate_paragraph_text(intent.position, &intent.text).map_err(Error::DocumentEdit)?;
        let text = checked_string_clone(&intent.text, "glossary batch replacement text")?;
        projected
            .replace_paragraph_text(intent.position, text)
            .map_err(Error::DocumentEdit)?;
    }
    base.check_execution()?;
    let wrapped = projected
        .projected()
        .shared_xml()
        .map_err(Error::DocumentEdit)?;
    let wrapped_end = wrapped
        .len()
        .checked_sub(envelope.wrapped_suffix_len)
        .ok_or_else(|| {
            Error::Document(crate::Error::InvalidFormat(
                "glossary batch wrapper suffix exceeds projected XML".into(),
            ))
        })?;
    let inner = wrapped
        .get(envelope.wrapped_prefix_len..wrapped_end)
        .ok_or_else(|| {
            Error::Document(crate::Error::InvalidFormat(
                "glossary batch wrapper body is outside projected XML".into(),
            ))
        })?;
    if inner.len() > base.limits.max_entry_bytes {
        return Err(Error::Limit {
            kind: LimitKind::SelectedEntryBytes,
            actual: inner.len(),
            maximum: base.limits.max_entry_bytes,
        });
    }
    let bytes = if materialize {
        Some(checked_vec_clone(inner, "glossary batch replacement")?)
    } else {
        None
    };
    Ok(GlossaryBatchProjection {
        range,
        length: inner.len(),
        bytes,
    })
}

fn validate_batch_ranges(entries: &[GlossaryBatchEntry], xml_len: usize) -> Result<()> {
    for entry in entries {
        if entry.layout.root_start_end > entry.layout.root_end_start
            || entry.layout.root_end_start > xml_len
        {
            return Err(Error::TopologyConflict);
        }
    }
    for pair in entries.windows(2) {
        if pair[0].layout.root_start_end == pair[1].layout.root_start_end
            || pair[0].layout.root_end_start == pair[1].layout.root_end_start
            || pair[0].layout.root_end_start > pair[1].layout.root_start_end
        {
            return Err(Error::TopologyConflict);
        }
    }
    Ok(())
}

fn batch_entry_index_error(index: usize, len: usize) -> Error {
    Error::Limit {
        kind: LimitKind::GlossaryBatchEntryIndex,
        actual: index.saturating_add(1),
        maximum: len,
    }
}

fn batch_paragraph_index_error(_entry_index: usize, index: usize, len: usize) -> Error {
    Error::Limit {
        kind: LimitKind::GlossaryBatchParagraphIndex,
        actual: index.saturating_add(1),
        maximum: len,
    }
}

/// The bounded resources of source-backed story-text edits, carried by [`Error::Limit`] (record 0355a).
///
/// Each variant names one bounded resource; [`std::fmt::Display`] gives its
/// human-readable name, the one the error messages print.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum LimitKind {
    /// Story attribute value.
    AttributeValue,
    /// Story body elements.
    BodyElements,
    /// Story entries.
    Entries,
    /// Glossary batch entry bytes.
    GlossaryBatchEntryBytes,
    /// Glossary batch entry index.
    GlossaryBatchEntryIndex,
    /// Glossary batch output bytes.
    GlossaryBatchOutputBytes,
    /// Glossary batch paragraph index.
    GlossaryBatchParagraphIndex,
    /// Glossary batch paragraph text.
    GlossaryBatchParagraphText,
    /// Glossary batch replacement bytes.
    GlossaryBatchReplacementBytes,
    /// Glossary batch selectors.
    GlossaryBatchSelectors,
    /// Glossary batch text output.
    GlossaryBatchTextOutput,
    /// Glossary batch transient bytes.
    GlossaryBatchTransientBytes,
    /// Glossary entry bodies.
    GlossaryEntryBodies,
    /// Glossary selector bytes.
    GlossarySelectorBytes,
    /// Glossary selector ID bytes.
    GlossarySelectorIdBytes,
    /// Glossary selector metadata bytes.
    GlossarySelectorMetadataBytes,
    /// Glossary selector name bytes.
    GlossarySelectorNameBytes,
    /// Glossary selector name key.
    GlossarySelectorNameKey,
    /// Namespace attribute value.
    NamespaceAttributeValue,
    /// Story namespace attributes.
    NamespaceAttributes,
    /// Namespace bindings.
    NamespaceBindings,
    /// Namespace bytes.
    NamespaceBytes,
    /// Story output bytes.
    OutputBytes,
    /// Story paragraph text.
    ParagraphText,
    /// Paragraphs.
    Paragraphs,
    /// Story relationship targets.
    RelationshipTargets,
    /// Replacement text bytes.
    ReplacementTextBytes,
    /// Replacements.
    Replacements,
    /// Selected entry bytes.
    SelectedEntryBytes,
    /// Story text output.
    TextOutput,
    /// Story transaction XML.
    TransactionXml,
    /// Wrapped story XML bytes.
    WrappedStoryXmlBytes,
    /// XML bytes.
    XmlBytes,
    /// XML depth.
    XmlDepth,
    /// XML events.
    XmlEvents,
    /// XML reference bytes.
    XmlReferenceBytes,
}

impl LimitKind {
    /// The resource's human-readable name.
    #[must_use]
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::AttributeValue => "story attribute value",
            Self::BodyElements => "story body elements",
            Self::Entries => "story entries",
            Self::GlossaryBatchEntryBytes => "glossary batch entry bytes",
            Self::GlossaryBatchEntryIndex => "glossary batch entry index",
            Self::GlossaryBatchOutputBytes => "glossary batch output bytes",
            Self::GlossaryBatchParagraphIndex => "glossary batch paragraph index",
            Self::GlossaryBatchParagraphText => "glossary batch paragraph text",
            Self::GlossaryBatchReplacementBytes => "glossary batch replacement bytes",
            Self::GlossaryBatchSelectors => "glossary batch selectors",
            Self::GlossaryBatchTextOutput => "glossary batch text output",
            Self::GlossaryBatchTransientBytes => "glossary batch transient bytes",
            Self::GlossaryEntryBodies => "glossary entry bodies",
            Self::GlossarySelectorBytes => "glossary selector bytes",
            Self::GlossarySelectorIdBytes => "glossary selector ID bytes",
            Self::GlossarySelectorMetadataBytes => "glossary selector metadata bytes",
            Self::GlossarySelectorNameBytes => "glossary selector name bytes",
            Self::GlossarySelectorNameKey => "glossary selector name key",
            Self::NamespaceAttributeValue => "namespace attribute value",
            Self::NamespaceAttributes => "story namespace attributes",
            Self::NamespaceBindings => "namespace bindings",
            Self::NamespaceBytes => "namespace bytes",
            Self::OutputBytes => "story output bytes",
            Self::ParagraphText => "story paragraph text",
            Self::Paragraphs => "paragraphs",
            Self::RelationshipTargets => "story relationship targets",
            Self::ReplacementTextBytes => "replacement text bytes",
            Self::Replacements => "replacements",
            Self::SelectedEntryBytes => "selected entry bytes",
            Self::TextOutput => "story text output",
            Self::TransactionXml => "story transaction XML",
            Self::WrappedStoryXmlBytes => "wrapped story XML bytes",
            Self::XmlBytes => "XML bytes",
            Self::XmlDepth => "XML depth",
            Self::XmlEvents => "XML events",
            Self::XmlReferenceBytes => "XML reference bytes",
        }
    }
}

impl std::fmt::Display for LimitKind {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter.write_str(self.as_str())
    }
}
