//! Source-preserving main-document snapshots, edits, and reversible patches.

mod durable;
#[cfg(test)]
mod publication_proof_tests;
#[cfg(test)]
mod scan_differential_tests;

use std::mem::size_of;
use std::sync::Arc;

use litchi_core::{ExecutionContext, Position, Reservation, Resource, SourceVersion};
use litchi_ooxml_common::private::{BindingTracker, split_qualified_name};
use litchi_ooxml_common::properties::time::DateTime;
use litchi_opc::{PackURI, PartData, SourceLineage, SourceXmlPart};
use quick_xml::events::Event;
use quick_xml::name::{Namespace, PrefixDeclaration, ResolveResult};
use quick_xml::reader::NsReader;
use quick_xml::{Reader, XmlVersion};
use thiserror::Error;

use crate::namespace::{
    NamespaceBindings, NamespaceCapture, STRICT_WORDPROCESSINGML_NAMESPACE,
    WORDPROCESSINGML_NAMESPACE, is_wordprocessing_namespace,
};
use crate::paragraph::Paragraph;

pub(crate) use durable::durable_transfer_operations;
pub use durable::{Composition, History, JoinError, PreparedEdit, ThreeWayError, ThreeWayPlan};
pub use litchi_core::patch::{
    CompositionLimits, HistoryLimits, MergeChoice, SubEditConflict, SubEditJoinFailure,
    ThreeWayMergeFailure,
};

pub(crate) const MAX_DOCUMENT_XML_BYTES: usize = 32 * 1024 * 1024;
const UTF8_BYTE_ORDER_MARK: &[u8] = b"\xEF\xBB\xBF";
const UTF8_BYTE_ORDER_MARK_TEXT: &str = "\u{feff}";
const MAX_DOCUMENT_DEPTH: usize = 256;
const MAX_DOCUMENT_NODES: usize = 1_000_000;
const MAX_OPERATIONS: usize = 4_096;
const MAX_REPLACEMENT_TEXT_BYTES: usize = 16 * 1024 * 1024;
// The managed layout scanner owns each quick-xml event and grows bounded
// layout indexes. `scan_text_owner` additionally retains one `TextSlot` for
// every editable text element and copies each element's prefix/local-name
// bytes. An empty `<t/>` is four bytes, so the slot count is conservatively
// bounded by `xml_len / 4`; each slot's vector capacity may double and its two
// backing allocations carry allocator metadata. The input multiplier covers
// quick-xml's owned event, decoded text, namespace state, and the retained
// `String` capacity. These bounds are charged before the first event so a
// cancellation or budget failure cannot leave a partially admitted scan.
const MANAGED_SCAN_MIN_TEXT_SLOT_BYTES: usize = 4;
const MANAGED_SCAN_SLOT_ALLOCATOR_OVERHEAD: usize = 64;
const MANAGED_SCAN_INPUT_MEMORY_MULTIPLIER: usize = 8;
const MANAGED_SCAN_FIXED_MEMORY: usize = 4096;
const MANAGED_SCAN_OBJECT_MULTIPLIER: usize = 2;
const MANAGED_SCAN_WORK_MULTIPLIER: usize = 2;
// The inherited resolver scan temporarily holds the source reader's
// namespace buffers while copying the effective bindings into the resolver
// lease. The lease is then cloned by the namespace-aware codec, so account for
// allocator geometric growth across those overlapping resolver states before
// parsing starts.
const MANAGED_NAMESPACE_SCAN_MEMORY_MULTIPLIER: usize = 32;
const MANAGED_NAMESPACE_SCAN_FIXED_MEMORY: usize = 4096;
const MAX_REVISION_ATTRIBUTES: usize = 256;
const MAX_REVISION_METADATA_VALUE_BYTES: usize = 64 * 1024;
const MAX_REVISION_TAG_ATTRIBUTE_BYTES: usize = 8 * 1024 * 1024;
const MARKUP_COMPATIBILITY_NAMESPACE: &[u8] =
    b"http://schemas.openxmlformats.org/markup-compatibility/2006";
const XML_NAMESPACE: &[u8] = b"http://www.w3.org/XML/1998/namespace";
const WORD_2023_DATE_UTC_NAMESPACE: &[u8] =
    b"http://schemas.microsoft.com/office/word/2023/wordml/word16du";

/// Result returned by main-document transaction operations.
pub type TransactionResult<T> = Result<T, TransactionError>;

pub(crate) fn validate_paragraph_text(position: Position, text: &str) -> TransactionResult<()> {
    validate_authored_text(text).map_err(|reason| TransactionError::Refused {
        position: position.get(),
        reason,
    })
}

/// A typed reason why a paragraph operation cannot be represented safely.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum Refusal {
    /// A composite selector is empty, duplicated, or not in canonical
    /// ownership order, so its scope cannot be represented unambiguously.
    AmbiguousCompositeSelector,
    /// The selected owner contains unsupported structural content.
    ComplexContent,
    /// The selected owner has no editable text element.
    ComplexRun,
    /// The selected hyperlink does not exist.
    HyperlinkNotFound,
    /// The selected direct paragraph run does not exist.
    RunNotFound,
    /// The selected simple field does not exist.
    FieldNotFound,
    /// The selected begin/separate/end complex field does not exist.
    ComplexFieldNotFound,
    /// The selected tracked insertion or deletion does not exist.
    RevisionNotFound,
    /// The selected tracked insertion or deletion contains a structure whose
    /// ownership or dependency semantics this bounded operation cannot prove.
    RevisionDependency,
    /// The selected direct inline content control does not exist.
    ContentControlNotFound,
    /// The selected table, row, or cell does not exist.
    CellNotFound,
    /// The requested text needs structural run elements such as `w:tab` or
    /// `w:br`, which this focused text operation does not synthesize.
    StructuralText,
}

impl std::fmt::Display for Refusal {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter.write_str(match self {
            Self::AmbiguousCompositeSelector => {
                "composite selector is empty, duplicated, or not canonically ordered"
            },
            Self::ComplexContent => "selected owner contains unsupported structural content",
            Self::ComplexRun => "selected content has no editable Word text",
            Self::HyperlinkNotFound => "direct paragraph hyperlink was not found",
            Self::RunNotFound => "direct paragraph run was not found",
            Self::FieldNotFound => "direct simple field was not found",
            Self::ComplexFieldNotFound => "complex field sequence was not found",
            Self::RevisionNotFound => "direct tracked revision was not found",
            Self::RevisionDependency => {
                "tracked revision contains unsupported move, range, property, table, or nested dependency content"
            },
            Self::ContentControlNotFound => "direct inline content control was not found",
            Self::CellNotFound => "table cell was not found",
            Self::StructuralText => "text requires structural WordprocessingML elements",
        })
    }
}

/// Direct tracked-revision wrapper selected for inert text replacement.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum RevisionKind {
    /// `w:ins` tracked insertion content.
    Insertion,
    /// `w:del` tracked deletion content.
    Deletion,
}

/// The Word redline disposition to apply to one direct inline revision.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum RevisionAction {
    /// Keep inserted content or remove deleted content.
    Accept,
    /// Remove inserted content or restore deleted content as ordinary text.
    Reject,
}

impl RevisionAction {
    const fn removes_content(self, kind: RevisionKind) -> bool {
        matches!(
            (self, kind),
            (Self::Reject, RevisionKind::Insertion) | (Self::Accept, RevisionKind::Deletion)
        )
    }
}

/// A checked direct-body paragraph/revision selector.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct RevisionSelector {
    paragraph: Position,
    kind: RevisionKind,
    revision: Position,
}

impl RevisionSelector {
    /// Construct a direct inline revision selector.
    #[must_use]
    pub const fn new(paragraph: Position, kind: RevisionKind, revision: Position) -> Self {
        Self {
            paragraph,
            kind,
            revision,
        }
    }

    /// Return the direct-body paragraph position.
    #[must_use]
    pub const fn paragraph(self) -> Position {
        self.paragraph
    }

    /// Return the selected tracked revision family.
    #[must_use]
    pub const fn kind(self) -> RevisionKind {
        self.kind
    }

    /// Return the direct wrapper position among revisions of [`Self::kind`].
    #[must_use]
    pub const fn revision(self) -> Position {
        self.revision
    }
}

/// One table/row/cell step in a bounded nested-table selector path.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct TableCellAddress {
    /// Direct table position in the body for the first step, or in the
    /// preceding cell for every nested step.
    pub table: Position,
    /// Direct row position in the selected table.
    pub row: Position,
    /// Direct cell position in the selected row.
    pub cell: Position,
}

impl TableCellAddress {
    /// Construct one checked-by-use nested table-cell address step.
    #[must_use]
    pub const fn new(table: Position, row: Position, cell: Position) -> Self {
        Self { table, row, cell }
    }
}

/// One direct hyperlink address within an owner-scoped paragraph collection.
///
/// The paragraph position is relative to the body, deepest block content
/// control, or deepest nested-table cell selected by the batch operation.
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub struct ParagraphHyperlinkAddress {
    /// Direct paragraph position within the selected owner.
    pub paragraph: Position,
    /// Direct hyperlink position within the selected paragraph.
    pub hyperlink: Position,
}

impl ParagraphHyperlinkAddress {
    /// Construct one checked-by-use paragraph/hyperlink address.
    #[must_use]
    pub const fn new(paragraph: Position, hyperlink: Position) -> Self {
        Self {
            paragraph,
            hyperlink,
        }
    }
}

/// Authored text paired with one owner-scoped direct hyperlink address.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct HyperlinkTextReplacement {
    address: ParagraphHyperlinkAddress,
    text: String,
}

/// Authored text paired with one direct-body paragraph position.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ParagraphTextReplacement {
    position: Position,
    text: String,
}

impl ParagraphTextReplacement {
    /// Construct one direct-body paragraph text replacement.
    #[must_use]
    pub fn new(position: Position, text: impl Into<String>) -> Self {
        Self {
            position,
            text: text.into(),
        }
    }

    /// Return the selected direct-body paragraph position.
    #[must_use]
    pub const fn position(&self) -> Position {
        self.position
    }

    /// Borrow the authored replacement text.
    #[must_use]
    pub fn text(&self) -> &str {
        &self.text
    }
}

impl HyperlinkTextReplacement {
    /// Construct one hyperlink text replacement.
    #[must_use]
    pub fn new(address: ParagraphHyperlinkAddress, text: impl Into<String>) -> Self {
        Self {
            address,
            text: text.into(),
        }
    }

    /// Return the selected direct paragraph/hyperlink address.
    #[must_use]
    pub const fn address(&self) -> ParagraphHyperlinkAddress {
        self.address
    }

    /// Borrow the authored replacement text.
    #[must_use]
    pub fn text(&self) -> &str {
        &self.text
    }
}

/// Dependency-checked paragraph payload prepared for one exact receiving
/// document. Package planning rewrites relationship references to dependencies
/// already proven present in the receiver before constructing this value.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ParagraphTransfer {
    target: Arc<Vec<u8>>,
    fragment: Arc<Vec<u8>>,
    dependency_digest: Arc<str>,
    inverse_dependency_digest: Arc<str>,
    graph: Arc<TransferGraph>,
}

impl ParagraphTransfer {
    pub(crate) fn new(
        target: Arc<Vec<u8>>,
        fragment: Vec<u8>,
        dependency_digest: String,
        inverse_dependency_digest: String,
        graph: TransferGraph,
    ) -> Self {
        Self {
            target,
            fragment: Arc::new(fragment),
            dependency_digest: dependency_digest.into(),
            inverse_dependency_digest: inverse_dependency_digest.into(),
            graph: Arc::new(graph),
        }
    }

    /// Exact compact paragraph XML retained by the plan.
    #[must_use]
    pub fn xml_bytes(&self) -> &[u8] {
        self.fragment.as_slice()
    }
}

/// Opaque dependency subgraph carried by a paragraph-transfer operation.
///
/// Package planning and publication own its bounded contents; callers retain
/// it only when inspecting or replaying a native operation.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct TransferGraph {
    pub(crate) main_relationships: Arc<[TransferRelationship]>,
    pub(crate) parts: Arc<[TransferPart]>,
}

impl TransferGraph {
    pub(crate) fn empty() -> Self {
        Self {
            main_relationships: Arc::new([]),
            parts: Arc::new([]),
        }
    }

    pub(crate) fn is_empty(&self) -> bool {
        self.main_relationships.is_empty() && self.parts.is_empty()
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) struct TransferPart {
    pub(crate) name: String,
    pub(crate) content_type: String,
    pub(crate) blob: Arc<Vec<u8>>,
    pub(crate) relationships: Arc<[TransferRelationship]>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) struct TransferRelationship {
    pub(crate) id: String,
    pub(crate) relationship_type: String,
    pub(crate) target: String,
    pub(crate) external: bool,
}

impl RevisionKind {
    const fn local_name(self) -> &'static [u8] {
        match self {
            Self::Insertion => b"ins",
            Self::Deletion => b"del",
        }
    }
}

/// Typed reason why a cross-package paragraph dependency closure cannot be
/// represented without copying or guessing package resources.
#[derive(Debug, Clone, PartialEq, Eq)]
#[non_exhaustive]
pub enum TransferRefusal {
    /// A relationship reference in the donor paragraph is dangling.
    MissingDonorRelationship(String),
    /// The receiving package has no semantically equivalent dependency edge.
    MissingEquivalentDependency {
        /// OPC relationship type URI.
        relationship_type: String,
        /// Exact external or relative internal target reference.
        target: String,
    },
    /// Multiple receiver edges are equivalent, so choosing an ID would guess.
    AmbiguousEquivalentDependency {
        /// OPC relationship type URI.
        relationship_type: String,
        /// Exact external or relative internal target reference.
        target: String,
    },
    /// The selected donor paragraph could not be represented as compact XML.
    InvalidParagraphXml,
}

impl std::fmt::Display for TransferRefusal {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::MissingDonorRelationship(identifier) => {
                write!(formatter, "donor relationship {identifier} is missing")
            },
            Self::MissingEquivalentDependency {
                relationship_type,
                target,
            } => write!(
                formatter,
                "receiver lacks dependency {relationship_type} -> {target}"
            ),
            Self::AmbiguousEquivalentDependency {
                relationship_type,
                target,
            } => write!(
                formatter,
                "receiver dependency {relationship_type} -> {target} is ambiguous"
            ),
            Self::InvalidParagraphXml => {
                formatter.write_str("donor paragraph XML is not transferable")
            },
        }
    }
}

/// A main-document transaction failure.
#[derive(Debug, Error)]
#[non_exhaustive]
pub enum TransactionError {
    /// The underlying DOCX document or package is invalid.
    #[error(transparent)]
    Document(#[from] crate::Error),
    /// A checked paragraph position is outside the projected document.
    #[error("paragraph position {position} is out of bounds for length {len}")]
    OutOfBounds {
        /// Requested zero-based paragraph position.
        position: usize,
        /// Projected direct-body paragraph count.
        len: usize,
    },
    /// The selected paragraph cannot be changed without guessing how to
    /// rewrite dependent or structured content.
    #[error("paragraph {position} edit refused: {reason}")]
    Refused {
        /// Selected zero-based paragraph position.
        position: usize,
        /// Stable refusal category.
        reason: Refusal,
    },
    /// A configured transaction resource ceiling was exceeded.
    #[error("document transaction {resource} limit exceeded: {actual} > {max}")]
    Limit {
        /// Bounded resource.
        resource: &'static str,
        /// Maximum accepted value.
        max: usize,
        /// Observed or requested value.
        actual: usize,
    },
    /// The patch target no longer has the exact source bytes captured by the
    /// edit.
    #[error("document patch source is stale")]
    StaleSource,
    /// A semantic durable operation's expected value did not match source.
    #[error("document patch semantic precondition does not match")]
    SemanticPrecondition,
    /// A common durable patch could not be constructed or decoded.
    #[error(transparent)]
    Durable(#[from] litchi_core::patch::PatchError),
    /// A common disjoint-composition bound or identifier was invalid.
    #[error(transparent)]
    Composition(#[from] litchi_core::patch::CompositionError),
    /// A durable patch used an unsupported or malformed DOCX vocabulary.
    #[error("invalid DOCX durable patch: {0}")]
    InvalidDurable(String),
    /// A cross-package transfer dependency could not be closed safely.
    #[error("paragraph transfer refused: {0}")]
    Transfer(TransferRefusal),
}

#[derive(Debug, Clone, PartialEq, Eq)]
struct SourceIdentity {
    lineage: SourceLineage,
    version: SourceVersion,
    partname: PackURI,
}

/// Reservations and the source execution policy retained by one managed
/// snapshot.  The token is shared by immutable snapshot and view clones; a
/// candidate receives a fresh token so its indexes remain charged while the
/// source snapshot is still live.
#[derive(Debug)]
pub(crate) struct ManagedAdmission {
    context: ExecutionContext,
    _memory: Reservation,
    _objects: Reservation,
}

impl ManagedAdmission {
    fn new(
        context: ExecutionContext,
        xml_len: usize,
        partname_len: usize,
    ) -> TransactionResult<Arc<Self>> {
        let index_bytes = xml_len
            .checked_mul(3)
            .and_then(|count| count.checked_mul(size_of::<Range>()))
            .ok_or(TransactionError::Limit {
                resource: "managed snapshot index bytes",
                max: usize::MAX,
                actual: usize::MAX,
            })?;
        let metadata_bytes = size_of::<Snapshot>()
            .checked_add(size_of::<ManagedXml>())
            .and_then(|bytes| bytes.checked_add(size_of::<Arc<ManagedXml>>()))
            .and_then(|bytes| bytes.checked_add(size_of::<PartData>()))
            .and_then(|bytes| bytes.checked_add(size_of::<Arc<PartData>>()))
            .and_then(|bytes| bytes.checked_add(size_of::<SourceXmlPart>()))
            .and_then(|bytes| bytes.checked_add(size_of::<Arc<SourceXmlPart>>()))
            .and_then(|bytes| bytes.checked_add(size_of::<SourceIdentity>()))
            .and_then(|bytes| bytes.checked_add(size_of::<Arc<SourceIdentity>>()))
            .and_then(|bytes| bytes.checked_add(size_of::<SourceLineage>()))
            .and_then(|bytes| bytes.checked_add(size_of::<ManagedAdmission>()))
            .and_then(|bytes| bytes.checked_add(size_of::<Arc<ManagedAdmission>>()))
            .and_then(|bytes| bytes.checked_add(partname_len.saturating_mul(2)))
            .and_then(|bytes| bytes.checked_add(4 * size_of::<Arc<[Range]>>()))
            .ok_or(TransactionError::Limit {
                resource: "managed snapshot metadata bytes",
                max: usize::MAX,
                actual: usize::MAX,
            })?;
        let memory = metadata_bytes
            .checked_add(index_bytes)
            .ok_or(TransactionError::Limit {
                resource: "managed snapshot memory",
                max: usize::MAX,
                actual: usize::MAX,
            })?;
        let objects =
            xml_len
                .min(MAX_DOCUMENT_NODES)
                .checked_add(4)
                .ok_or(TransactionError::Limit {
                    resource: "managed snapshot objects",
                    max: usize::MAX,
                    actual: usize::MAX,
                })?;
        let memory = reserve_managed(&context, Resource::Memory, memory)?;
        let objects = reserve_managed(&context, Resource::Objects, objects)?;
        Ok(Arc::new(Self {
            context,
            _memory: memory,
            _objects: objects,
        }))
    }

    fn scan(&self, xml_len: usize) -> TransactionResult<ScanAdmission> {
        let scan_memory = reserve_managed(
            &self.context,
            Resource::Memory,
            managed_scan_memory(xml_len)?,
        )?;
        let scan_objects = reserve_managed(
            &self.context,
            Resource::Objects,
            managed_scan_slot_count(xml_len)?
                .checked_mul(MANAGED_SCAN_OBJECT_MULTIPLIER)
                .and_then(|objects| objects.checked_add(8))
                .ok_or(TransactionError::Limit {
                    resource: "managed scan objects",
                    max: usize::MAX,
                    actual: usize::MAX,
                })?,
        )?;
        let depth = reserve_managed(&self.context, Resource::Depth, MAX_DOCUMENT_DEPTH)?;
        consume_managed(
            &self.context,
            Resource::Work,
            xml_len
                .checked_mul(MANAGED_SCAN_WORK_MULTIPLIER)
                .ok_or(TransactionError::Limit {
                    resource: "managed scan work",
                    max: usize::MAX,
                    actual: usize::MAX,
                })?,
        )?;
        Ok(ScanAdmission {
            _memory: scan_memory,
            _objects: scan_objects,
            _depth: Some(depth),
        })
    }

    /// Admit a paragraph/run owner scan. Unlike the document layout scanner,
    /// the owner parser resolves every qualified child name while retaining
    /// text-slot prefixes. Until that parser is instrumented per event, its
    /// finite worst case is bounded by the square of the input span: both the
    /// number of events and the active namespace bindings are input-sized.
    /// Charge this before the parser grows any owner state; layout scans keep
    /// their observed per-event charge instead of paying this quadratic bound.
    fn owner_scan(&self, xml_len: usize) -> TransactionResult<ScanAdmission> {
        let scan = self.scan(xml_len)?;
        let units = xml_len.checked_add(1).ok_or(TransactionError::Limit {
            resource: "managed namespace lookup work",
            max: usize::MAX,
            actual: usize::MAX,
        })?;
        let namespace_work = units.checked_mul(units).ok_or(TransactionError::Limit {
            resource: "managed namespace lookup work",
            max: usize::MAX,
            actual: usize::MAX,
        })?;
        consume_managed(&self.context, Resource::Work, namespace_work)?;
        Ok(scan)
    }

    /// Admit the transient parser state used by a retained paragraph or run
    /// view. The returned token is deliberately scoped to one fallible read;
    /// caller-owned result values retain their own source owner separately.
    pub(crate) fn parser_admission(
        &self,
        xml_len: usize,
    ) -> crate::error::Result<ManagedParserAdmission> {
        self.owner_scan(xml_len)
            .map(|scan| ManagedParserAdmission { _scan: scan })
            .map_err(managed_parser_error)
    }

    pub(crate) fn namespace_scan_admission(
        &self,
        fragment_len: usize,
        owner_span_len: usize,
    ) -> crate::error::Result<ManagedNamespaceAdmission> {
        let memory = owner_span_len
            .checked_mul(MANAGED_NAMESPACE_SCAN_MEMORY_MULTIPLIER)
            .and_then(|bytes| bytes.checked_add(MANAGED_NAMESPACE_SCAN_FIXED_MEMORY))
            .and_then(|bytes| bytes.checked_add(size_of::<ManagedNamespaceAdmission>()))
            .and_then(|bytes| bytes.checked_add(size_of::<NsReader<&[u8]>>()))
            .ok_or_else(|| {
                crate::Error::InvalidFormat(
                    "managed inherited namespace scan memory overflows usize".into(),
                )
            })?;
        let objects = owner_span_len.checked_add(8).ok_or_else(|| {
            crate::Error::InvalidFormat(
                "managed inherited namespace scan objects overflow usize".into(),
            )
        })?;
        let memory = reserve_managed(&self.context, Resource::Memory, memory)
            .map_err(managed_parser_error)?;
        let objects = reserve_managed(&self.context, Resource::Objects, objects)
            .map_err(managed_parser_error)?;
        let depth = reserve_managed(&self.context, Resource::Depth, MAX_DOCUMENT_DEPTH)
            .map_err(managed_parser_error)?;
        // Namespace lookup can compare each retained fragment event with all
        // bindings accumulated in the owner prefix.  Charge that finite
        // fragment-by-owner cross-term before the resolver or codec starts;
        // the scanner still charges its observed event-level lookups below.
        let fragment_units = fragment_len.checked_add(1).ok_or_else(|| {
            crate::Error::InvalidFormat(
                "managed inherited namespace fragment length overflows usize".into(),
            )
        })?;
        let owner_units = owner_span_len.checked_add(1).ok_or_else(|| {
            crate::Error::InvalidFormat(
                "managed inherited namespace owner length overflows usize".into(),
            )
        })?;
        let cross_term = fragment_units.checked_mul(owner_units).ok_or_else(|| {
            crate::Error::InvalidFormat(
                "managed inherited namespace lookup work overflows usize".into(),
            )
        })?;
        let work = owner_span_len
            .checked_mul(MANAGED_SCAN_WORK_MULTIPLIER)
            .and_then(|work| work.checked_add(cross_term))
            .ok_or_else(|| {
                crate::Error::InvalidFormat(
                    "managed inherited namespace scan work overflows usize".into(),
                )
            })?;
        consume_managed(&self.context, Resource::Work, work).map_err(managed_parser_error)?;
        Ok(ManagedNamespaceAdmission {
            context: self.context.clone(),
            _memory: memory,
            _objects: objects,
            _depth: depth,
        })
    }
}

#[derive(Debug)]
struct ScanAdmission {
    _memory: Reservation,
    _objects: Reservation,
    _depth: Option<Reservation>,
}

impl ScanAdmission {
    /// Release parser nesting capacity once the owned text slots are built.
    /// Memory/object charges remain with the caller while those slots live.
    fn release_depth(&mut self) {
        drop(self._depth.take());
    }
}

#[derive(Debug)]
pub(crate) struct ManagedParserAdmission {
    _scan: ScanAdmission,
}

#[derive(Debug)]
pub(crate) struct ManagedNamespaceAdmission {
    context: ExecutionContext,
    _memory: Reservation,
    _objects: Reservation,
    _depth: Reservation,
}

impl ManagedNamespaceAdmission {
    pub(crate) fn check(&self) -> crate::error::Result<()> {
        self.context.check().map_err(managed_error)
    }

    pub(crate) fn consume_lookup_work(
        &self,
        event_bytes: usize,
        active_bindings: usize,
    ) -> crate::error::Result<()> {
        let work = event_bytes
            .checked_mul(active_bindings.saturating_add(1))
            .ok_or_else(|| {
                crate::Error::InvalidFormat(
                    "managed inherited namespace lookup work overflows usize".into(),
                )
            })?;
        consume_managed(&self.context, Resource::Work, work).map_err(managed_parser_error)
    }
}

#[derive(Debug)]
struct OperationAdmission {
    _memory: Reservation,
    _objects: Reservation,
}

#[derive(Debug)]
struct StringAdmission {
    _memory: Reservation,
}

fn reserve_managed(
    context: &ExecutionContext,
    resource: Resource,
    amount: usize,
) -> TransactionResult<Reservation> {
    let amount = u64::try_from(amount).map_err(|_error| TransactionError::Limit {
        resource: "managed reservation amount",
        max: usize::MAX,
        actual: usize::MAX,
    })?;
    context.reserve(resource, amount).map_err(managed_execution)
}

fn consume_managed(
    context: &ExecutionContext,
    resource: Resource,
    amount: usize,
) -> TransactionResult<()> {
    let amount = u64::try_from(amount).map_err(|_error| TransactionError::Limit {
        resource: "managed work amount",
        max: usize::MAX,
        actual: usize::MAX,
    })?;
    context.consume(resource, amount).map_err(managed_execution)
}

fn managed_error(error: litchi_core::ExecutionError) -> crate::Error {
    let error = match error {
        litchi_core::ExecutionError::Cancelled => litchi_opc::OpcError::Cancelled,
        error => litchi_opc::OpcError::Execution(error),
    };
    crate::Error::Opc(error)
}

fn managed_execution(error: litchi_core::ExecutionError) -> TransactionError {
    TransactionError::Document(managed_error(error))
}

fn managed_parser_error(error: TransactionError) -> crate::Error {
    match error {
        TransactionError::Document(error) => error,
        error => crate::Error::Other(format!(
            "managed paragraph parser admission failed: {error}"
        )),
    }
}

fn managed_scan_memory(xml_len: usize) -> TransactionResult<usize> {
    let slot_count = managed_scan_slot_count(xml_len)?;
    let slot_bytes = size_of::<TextSlot>()
        .checked_mul(2)
        .and_then(|bytes| bytes.checked_add(MANAGED_SCAN_SLOT_ALLOCATOR_OVERHEAD))
        .and_then(|bytes| bytes.checked_mul(slot_count))
        .ok_or(TransactionError::Limit {
            resource: "managed text-slot memory",
            max: usize::MAX,
            actual: usize::MAX,
        })?;
    MANAGED_SCAN_FIXED_MEMORY
        .checked_add(
            xml_len
                .checked_mul(MANAGED_SCAN_INPUT_MEMORY_MULTIPLIER)
                .ok_or(TransactionError::Limit {
                    resource: "managed scan parser memory",
                    max: usize::MAX,
                    actual: usize::MAX,
                })?,
        )
        .and_then(|bytes| bytes.checked_add(slot_bytes))
        .ok_or(TransactionError::Limit {
            resource: "managed scan memory",
            max: usize::MAX,
            actual: usize::MAX,
        })
}

fn managed_scan_slot_count(xml_len: usize) -> TransactionResult<usize> {
    xml_len
        .checked_add(MANAGED_SCAN_MIN_TEXT_SLOT_BYTES - 1)
        .and_then(|bytes| bytes.checked_div(MANAGED_SCAN_MIN_TEXT_SLOT_BYTES))
        .ok_or(TransactionError::Limit {
            resource: "managed text-slot count",
            max: usize::MAX,
            actual: usize::MAX,
        })
}

impl SourceIdentity {
    fn matches(&self, other: &Self) -> bool {
        self.lineage == other.lineage
            && self.version == other.version
            && self.partname.is_equivalent_to(&other.partname)
    }

    fn matches_parts(
        &self,
        lineage: &SourceLineage,
        version: SourceVersion,
        partname: &PackURI,
    ) -> bool {
        self.lineage.eq(lineage)
            && self.version == version
            && self.partname.is_equivalent_to(partname)
    }
}

#[derive(Debug, Clone)]
struct ManagedXml {
    data: Arc<PartData>,
    identity: Arc<SourceIdentity>,
}

#[derive(Debug, Clone)]
enum XmlStorage {
    Owned(Arc<Vec<u8>>),
    /// Unmanaged bytes whose package source identity is retained solely for
    /// exact same-source checks.  This is used when an OPC package cannot
    /// retain a managed `PartData` owner, while still refusing equal-byte
    /// patches from a foreign package lineage.
    OwnedWithIdentity {
        xml: Arc<Vec<u8>>,
        identity: Arc<SourceIdentity>,
    },
    Managed(Arc<ManagedXml>),
    Source {
        xml: Arc<SourceXmlPart>,
        identity: Arc<SourceIdentity>,
    },
}

impl XmlStorage {
    fn bytes(&self) -> &[u8] {
        match self {
            Self::Owned(xml) => xml.as_slice(),
            Self::OwnedWithIdentity { xml, .. } => xml.as_slice(),
            Self::Managed(xml) => xml.data.as_bytes(),
            Self::Source { xml, .. } => xml.bytes(),
        }
    }

    fn identity(&self) -> Option<&SourceIdentity> {
        match self {
            Self::Owned(_) => None,
            Self::OwnedWithIdentity { identity, .. } => Some(identity.as_ref()),
            Self::Managed(xml) => Some(xml.identity.as_ref()),
            Self::Source { identity, .. } => Some(identity.as_ref()),
        }
    }

    fn source_xml(&self) -> Option<SourceXmlPart> {
        match self {
            Self::Source { xml, .. } => Some(xml.as_ref().clone()),
            Self::Owned(_) | Self::OwnedWithIdentity { .. } | Self::Managed(_) => None,
        }
    }
}

/// An immutable, cheaply clonable snapshot of the main document XML.
#[derive(Debug, Clone)]
pub struct Snapshot {
    xml: XmlStorage,
    paragraphs: Arc<[Range]>,
    paragraph_namespaces: Arc<[NamespaceBindings]>,
    tables: Arc<[Range]>,
    block_controls: Arc<[Range]>,
    content_end: u32,
    conformance: Conformance,
    admission: Option<Arc<ManagedAdmission>>,
    /// Proof that these exact bytes passed the package writer's source
    /// publication audit, set only by [`Edit::commit`] on the preserving route
    /// (change 0754). Every other constructor leaves it empty, and a derived
    /// snapshot never inherits it, because a proof covers one allocation.
    publication: Option<xml_minifier::audit::VerifiedSource>,
}

impl Snapshot {
    /// Parse and retain one bounded `WordprocessingML` main document.
    ///
    /// # Errors
    ///
    /// Returns a typed document or resource-limit error when the XML is
    /// malformed, unsupported, or exceeds the transaction bounds.
    pub fn from_xml(source_xml: impl Into<Vec<u8>>) -> TransactionResult<Self> {
        let xml = source_xml.into();
        if xml.len() > MAX_DOCUMENT_XML_BYTES {
            return Err(TransactionError::Limit {
                resource: "XML bytes",
                max: MAX_DOCUMENT_XML_BYTES,
                actual: xml.len(),
            });
        }
        let layout = scan_document(&xml)?;
        Ok(Self {
            xml: XmlStorage::Owned(Arc::new(xml)),
            paragraphs: layout.paragraphs.into(),
            paragraph_namespaces: layout.paragraph_namespaces.into(),
            tables: layout.tables.into(),
            block_controls: layout.block_controls.into(),
            content_end: layout.content_end,
            conformance: layout.conformance,
            admission: None,
            publication: None,
        })
    }

    pub(crate) fn from_shared_xml(xml: Arc<Vec<u8>>) -> TransactionResult<Self> {
        if xml.len() > MAX_DOCUMENT_XML_BYTES {
            return Err(TransactionError::Limit {
                resource: "XML bytes",
                max: MAX_DOCUMENT_XML_BYTES,
                actual: xml.len(),
            });
        }
        let layout = scan_document(&xml)?;
        Ok(Self {
            xml: XmlStorage::Owned(xml),
            paragraphs: layout.paragraphs.into(),
            paragraph_namespaces: layout.paragraph_namespaces.into(),
            tables: layout.tables.into(),
            block_controls: layout.block_controls.into(),
            content_end: layout.content_end,
            conformance: layout.conformance,
            admission: None,
            publication: None,
        })
    }

    /// Parse shared XML while retaining the OPC source identity that produced
    /// it.  The bytes remain unmanaged, so this constructor does not claim a
    /// managed cache reservation or a source-publication token; the identity
    /// only participates in exact same-source checks for patches and commits.
    pub(crate) fn from_shared_xml_with_source_identity(
        xml: Arc<Vec<u8>>,
        lineage: SourceLineage,
        version: SourceVersion,
        partname: &PackURI,
    ) -> TransactionResult<Self> {
        Self::from_owned_xml_with_identity(
            xml,
            Arc::new(SourceIdentity {
                lineage,
                version,
                partname: partname.clone(),
            }),
        )
    }

    fn from_owned_xml_with_identity(
        xml: Arc<Vec<u8>>,
        identity: Arc<SourceIdentity>,
    ) -> TransactionResult<Self> {
        let mut snapshot = Self::from_shared_xml(Arc::clone(&xml))?;
        snapshot.xml = XmlStorage::OwnedWithIdentity { xml, identity };
        Ok(snapshot)
    }

    fn with_rewritten_xml(&self, xml: Vec<u8>) -> TransactionResult<Self> {
        if let Some(identity) = self.source_identity() {
            Self::from_owned_xml_with_identity(Arc::new(xml), identity)
        } else {
            Self::from_xml(xml)
        }
    }

    /// Rewrite the retained XML after an exact splice of whole direct-body
    /// paragraphs, deriving the new layout instead of rescanning the part.
    ///
    /// [`Self::with_rewritten_xml`] rebuilds the layout with a full
    /// [`scan_document`] pass over the complete main document, which is the
    /// dominant cost of an ordinary one-paragraph edit. Every splice accepted
    /// here replaces one whole direct-body paragraph with a fragment of the
    /// same structural shape, so the new layout is the old one with those
    /// paragraphs resized and everything after them shifted, and every
    /// whole-document verdict the scanner reaches is unchanged (see
    /// [`preserves_body_child_shape`]). Anything the proof does not establish
    /// falls back to the full rescan, so the incremental path is never the only
    /// thing between unproven bytes and a snapshot.
    fn with_spliced_paragraphs(
        &self,
        xml: Vec<u8>,
        splices: &[ParagraphSplice<'_>],
    ) -> TransactionResult<Self> {
        match self.spliced_layout(xml.len(), splices)? {
            Some(layout) => Ok(self.with_spliced_layout(xml, layout)),
            None => self.with_rewritten_xml(xml),
        }
    }

    /// Derive the layout that [`scan_document`] would produce for the spliced
    /// XML, or `None` when the splices are not provably layout-preserving.
    ///
    /// `None` is never a refusal: it routes the caller back to the full rescan,
    /// which keeps every typed error, limit and refusal exactly where it was.
    fn spliced_layout(
        &self,
        xml_len: usize,
        splices: &[ParagraphSplice<'_>],
    ) -> TransactionResult<Option<SplicedLayout>> {
        if splices.is_empty() || xml_len > MAX_DOCUMENT_XML_BYTES {
            return Ok(None);
        }
        let source = self.xml.bytes();
        let mut bounds = Vec::new();
        bounds
            .try_reserve_exact(splices.len())
            .map_err(|source| crate::Error::Allocation {
                resource: "document paragraph splice bounds",
                source,
            })?;
        let mut previous_end = 0usize;
        for splice in splices {
            if splice.start < previous_end || splice.start > splice.end {
                return Ok(None);
            }
            let Some(before) = source.get(splice.start..splice.end) else {
                return Ok(None);
            };
            if !preserves_body_child_shape(before, splice.replacement) {
                return Ok(None);
            }
            // Each paragraph's inherited-namespace snapshot is carried over
            // unchanged, which is exact only while neither side of a splice
            // declares a namespace; any declaration takes the full rescan.
            if contains_namespace_declaration_marker(before)
                || contains_namespace_declaration_marker(splice.replacement)
            {
                return Ok(None);
            }
            let (Ok(start), Ok(end), Ok(length)) = (
                u32::try_from(splice.start),
                u32::try_from(splice.end),
                u32::try_from(splice.replacement.len()),
            ) else {
                return Ok(None);
            };
            bounds.push(SpliceBound { start, end, length });
            previous_end = splice.end;
        }

        let mut paragraphs = Vec::new();
        paragraphs
            .try_reserve_exact(self.paragraphs.len())
            .map_err(|source| crate::Error::Allocation {
                resource: "spliced document paragraph ranges",
                source,
            })?;
        let mut index = 0usize;
        let mut delta = 0i64;
        for range in self.paragraphs.iter() {
            let Some(start) = shifted_offset(range.start, delta) else {
                return Ok(None);
            };
            match bounds.get(index) {
                Some(bound) if bound.start == range.start => {
                    if range.start.checked_add(range.length) != Some(bound.end) {
                        return Ok(None);
                    }
                    let Some(next) = bound.accumulate(delta) else {
                        return Ok(None);
                    };
                    paragraphs.push(Range {
                        start,
                        length: bound.length,
                    });
                    delta = next;
                    index = index.saturating_add(1);
                },
                _ => paragraphs.push(Range {
                    start,
                    length: range.length,
                }),
            }
        }
        if index != bounds.len() {
            return Ok(None);
        }

        let (Some(tables), Some(block_controls)) = (
            shift_sibling_ranges(&self.tables, &bounds)?,
            shift_sibling_ranges(&self.block_controls, &bounds)?,
        ) else {
            return Ok(None);
        };
        // Every direct-body child ends at or before the body-final section
        // properties, and at or before `</w:body>` when there are none, so the
        // insertion offset always sits after the last rewritten paragraph and
        // moves by the total length change.
        let Some(last) = bounds.last() else {
            return Ok(None);
        };
        if self.content_end < last.end {
            return Ok(None);
        }
        let Some(content_end) = shifted_offset(self.content_end, delta) else {
            return Ok(None);
        };
        if usize::try_from(content_end).is_ok_and(|offset| offset > xml_len) {
            return Ok(None);
        }
        Ok(Some(SplicedLayout {
            paragraphs: paragraphs.into(),
            paragraph_namespaces: Arc::clone(&self.paragraph_namespaces),
            tables,
            block_controls,
            content_end,
        }))
    }

    /// Retain spliced XML under a layout that was derived rather than scanned.
    ///
    /// The storage variant, the conformance and the absent managed admission
    /// are exactly what [`Self::with_rewritten_xml`] would have produced for
    /// the same bytes; only the layout arrives by a different route. The
    /// document-size limit is checked by [`Self::spliced_layout`], which sends
    /// oversized input to the rescan so the typed limit error keeps its
    /// identity.
    fn with_spliced_layout(&self, xml: Vec<u8>, layout: SplicedLayout) -> Self {
        let xml = Arc::new(xml);
        Self {
            xml: match self.source_identity() {
                Some(identity) => XmlStorage::OwnedWithIdentity { xml, identity },
                None => XmlStorage::Owned(xml),
            },
            paragraphs: layout.paragraphs,
            paragraph_namespaces: layout.paragraph_namespaces,
            tables: layout.tables,
            block_controls: layout.block_controls,
            content_end: layout.content_end,
            conformance: self.conformance,
            admission: None,
            publication: None,
        }
    }

    pub(crate) fn from_managed_part(
        data: PartData,
        lineage: SourceLineage,
        version: SourceVersion,
        partname: &PackURI,
        context: ExecutionContext,
    ) -> TransactionResult<Self> {
        if data.as_bytes().len() > MAX_DOCUMENT_XML_BYTES {
            return Err(TransactionError::Limit {
                resource: "XML bytes",
                max: MAX_DOCUMENT_XML_BYTES,
                actual: data.as_bytes().len(),
            });
        }
        let admission =
            ManagedAdmission::new(context, data.as_bytes().len(), partname.as_str().len())?;
        let scan_admission = admission.scan(data.as_bytes().len())?;
        admission.context.check().map_err(managed_execution)?;
        let layout = scan_document_with_context(data.as_bytes(), Some(&admission.context))?;
        admission.context.check().map_err(managed_execution)?;
        drop(scan_admission);
        Ok(Self {
            xml: XmlStorage::Managed(Arc::new(ManagedXml {
                data: Arc::new(data),
                identity: Arc::new(SourceIdentity {
                    lineage,
                    version,
                    partname: partname.clone(),
                }),
            })),
            paragraphs: layout.paragraphs.into(),
            paragraph_namespaces: layout.paragraph_namespaces.into(),
            tables: layout.tables.into(),
            block_controls: layout.block_controls.into(),
            content_end: layout.content_end,
            conformance: layout.conformance,
            admission: Some(admission),
            publication: None,
        })
    }

    pub(crate) fn from_source_xml(
        source_xml: SourceXmlPart,
        lineage: SourceLineage,
        version: SourceVersion,
        partname: &PackURI,
        context: ExecutionContext,
    ) -> TransactionResult<Self> {
        if source_xml.bytes().len() > MAX_DOCUMENT_XML_BYTES {
            return Err(TransactionError::Limit {
                resource: "XML bytes",
                max: MAX_DOCUMENT_XML_BYTES,
                actual: source_xml.bytes().len(),
            });
        }
        let admission =
            ManagedAdmission::new(context, source_xml.bytes().len(), partname.as_str().len())?;
        let scan_admission = admission.scan(source_xml.bytes().len())?;
        admission.context.check().map_err(managed_execution)?;
        let layout = scan_document_with_context(source_xml.bytes(), Some(&admission.context))?;
        admission.context.check().map_err(managed_execution)?;
        drop(scan_admission);
        Ok(Self {
            xml: XmlStorage::Source {
                xml: Arc::new(source_xml),
                identity: Arc::new(SourceIdentity {
                    lineage,
                    version,
                    partname: partname.clone(),
                }),
            },
            paragraphs: layout.paragraphs.into(),
            paragraph_namespaces: layout.paragraph_namespaces.into(),
            tables: layout.tables.into(),
            block_controls: layout.block_controls.into(),
            content_end: layout.content_end,
            conformance: layout.conformance,
            admission: Some(admission),
            publication: None,
        })
    }

    fn from_source_xml_with_identity(
        source_xml: SourceXmlPart,
        identity: Arc<SourceIdentity>,
        context: ExecutionContext,
    ) -> TransactionResult<Self> {
        if source_xml.bytes().len() > MAX_DOCUMENT_XML_BYTES {
            return Err(TransactionError::Limit {
                resource: "XML bytes",
                max: MAX_DOCUMENT_XML_BYTES,
                actual: source_xml.bytes().len(),
            });
        }
        let admission = ManagedAdmission::new(
            context,
            source_xml.bytes().len(),
            identity.partname.as_str().len(),
        )?;
        let scan_admission = admission.scan(source_xml.bytes().len())?;
        admission.context.check().map_err(managed_execution)?;
        let layout = scan_document_with_context(source_xml.bytes(), Some(&admission.context))?;
        admission.context.check().map_err(managed_execution)?;
        drop(scan_admission);
        Ok(Self {
            xml: XmlStorage::Source {
                xml: Arc::new(source_xml),
                identity,
            },
            paragraphs: layout.paragraphs.into(),
            paragraph_namespaces: layout.paragraph_namespaces.into(),
            tables: layout.tables.into(),
            block_controls: layout.block_controls.into(),
            content_end: layout.content_end,
            conformance: layout.conformance,
            admission: Some(admission),
            publication: None,
        })
    }

    /// The proof that these exact bytes passed the package writer's source
    /// publication audit, when the commit that produced this snapshot
    /// established one.
    pub(crate) fn publication_proof(&self) -> Option<&xml_minifier::audit::VerifiedSource> {
        self.publication
            .as_ref()
            .filter(|proof| proof.covers(self.xml_bytes(), xml_minifier::audit::Limits::default()))
    }

    /// Attach a publication audit proof. A proof that does not cover exactly
    /// these bytes is dropped, so a snapshot never carries a proof for other
    /// bytes.
    fn with_publication_proof(mut self, proof: xml_minifier::audit::VerifiedSource) -> Self {
        if proof.covers(self.xml_bytes(), xml_minifier::audit::Limits::default()) {
            self.publication = Some(proof);
        }
        self
    }

    pub(crate) fn shared_xml(&self) -> TransactionResult<Arc<Vec<u8>>> {
        match &self.xml {
            XmlStorage::Owned(xml) => Ok(Arc::clone(xml)),
            XmlStorage::OwnedWithIdentity { xml, .. } => Ok(Arc::clone(xml)),
            XmlStorage::Managed(_) | XmlStorage::Source { .. } => {
                Err(TransactionError::Document(crate::Error::UnsafeEdit {
                    format: "DOCX",
                    operation: "snapshot.shared_xml",
                    reason: "managed document XML cannot escape its retained source owner",
                }))
            },
        }
    }

    /// Borrow the exact main-document XML bytes.
    #[must_use]
    pub fn xml_bytes(&self) -> &[u8] {
        self.xml.bytes()
    }

    /// Return the number of direct main-body paragraphs.
    #[must_use]
    pub fn paragraph_count(&self) -> usize {
        self.paragraphs.len()
    }

    fn range(&self, position: Position) -> TransactionResult<Range> {
        self.paragraphs
            .get(position.get())
            .copied()
            .ok_or(TransactionError::OutOfBounds {
                position: position.get(),
                len: self.paragraph_count(),
            })
    }

    /// Borrow one direct main-body paragraph through a checked position.
    #[must_use]
    pub fn paragraph(&self, position: Position) -> Option<Paragraph> {
        self.paragraphs
            .get(position.get())
            .and_then(|range| match &self.xml {
                XmlStorage::Owned(xml) | XmlStorage::OwnedWithIdentity { xml, .. } => {
                    // Every scan records one inherited-namespace snapshot per
                    // paragraph range, so this lookup cannot miss.
                    self.paragraph_namespaces
                        .get(position.get())
                        .map(|namespaces| {
                            Paragraph::from_arc_range_with_context(
                                Arc::clone(xml),
                                range.start,
                                range.length,
                                Arc::clone(namespaces),
                            )
                        })
                },
                XmlStorage::Managed(xml) => self.admission.as_ref().map(|admission| {
                    Paragraph::from_managed_range(
                        Arc::clone(&xml.data),
                        Arc::clone(admission),
                        range.start,
                        range.length,
                    )
                }),
                XmlStorage::Source { xml, .. } => self.admission.as_ref().map(|admission| {
                    Paragraph::from_source_range(
                        Arc::clone(xml),
                        Arc::clone(admission),
                        range.start,
                        range.length,
                    )
                }),
            })
    }

    /// Return all direct main-body paragraphs without copying their XML.
    #[must_use]
    pub fn paragraphs(&self) -> Vec<Paragraph> {
        self.paragraphs
            .iter()
            .zip(self.paragraph_namespaces.iter())
            .filter_map(|(range, namespaces)| match &self.xml {
                XmlStorage::Owned(xml) | XmlStorage::OwnedWithIdentity { xml, .. } => {
                    Some(Paragraph::from_arc_range_with_context(
                        Arc::clone(xml),
                        range.start,
                        range.length,
                        Arc::clone(namespaces),
                    ))
                },
                XmlStorage::Managed(xml) => self.admission.as_ref().map(|admission| {
                    Paragraph::from_managed_range(
                        Arc::clone(&xml.data),
                        Arc::clone(admission),
                        range.start,
                        range.length,
                    )
                }),
                XmlStorage::Source { xml, .. } => self.admission.as_ref().map(|admission| {
                    Paragraph::from_source_range(
                        Arc::clone(xml),
                        Arc::clone(admission),
                        range.start,
                        range.length,
                    )
                }),
            })
            .collect()
    }

    /// Return the number of direct main-body tables.
    #[must_use]
    pub fn table_count(&self) -> usize {
        self.tables.len()
    }

    /// Return the number of direct main-body block content controls.
    #[must_use]
    pub fn block_content_control_count(&self) -> usize {
        self.block_controls.len()
    }

    /// Start an isolated edit whose selectors resolve against its projected
    /// state.
    #[must_use]
    pub fn edit(&self) -> Edit {
        Edit {
            base: self.clone(),
            projected: self.clone(),
            operations: Vec::new(),
            replacement_text_bytes: 0,
            operation_admission: None,
            operation_string_admission: None,
            compaction: CompactionPolicy::default(),
        }
    }

    fn managed_context(&self) -> Option<ExecutionContext> {
        self.admission
            .as_ref()
            .map(|admission| admission.context.clone())
    }

    fn same_source(&self, other: &Self) -> bool {
        // The retained bytes are immutable, so the same address and length
        // is an exact equality proof; distinct allocations still receive the
        // full byte comparison. An edit's projection shares its base's
        // allocation until it changes, so an exact no-op settles here.
        let (left, right) = (self.xml.bytes(), other.xml.bytes());
        if !std::ptr::eq(left, right) && left != right {
            return false;
        }
        match (self.xml.identity(), other.xml.identity()) {
            (Some(left), Some(right)) => left.matches(right),
            (None, None) => true,
            _ => false,
        }
    }

    pub(crate) fn is_source_backed(&self) -> bool {
        matches!(&self.xml, XmlStorage::Source { .. })
    }

    pub(crate) fn is_managed(&self) -> bool {
        matches!(
            &self.xml,
            XmlStorage::Managed(_) | XmlStorage::Source { .. }
        )
    }

    pub(crate) fn source_xml(&self) -> Option<SourceXmlPart> {
        self.xml.source_xml()
    }

    pub(crate) fn source_identity_matches(
        &self,
        lineage: &SourceLineage,
        version: SourceVersion,
        partname: &PackURI,
    ) -> bool {
        self.xml
            .identity()
            .is_some_and(|identity| identity.matches_parts(lineage, version, partname))
    }

    pub(crate) fn reuse_if_source_xml_matches(
        &self,
        source_xml: &SourceXmlPart,
        lineage: &SourceLineage,
        version: SourceVersion,
        partname: &PackURI,
    ) -> Option<Self> {
        if !self.source_identity_matches(lineage, version, partname) {
            return None;
        }
        let before_bytes = self.xml.bytes();
        let current_bytes = source_xml.bytes();
        // The OPC proof retains immutable bytes. Pointer identity is therefore
        // an exact equality proof and avoids a second full comparison on a
        // cache hit; distinct allocations still receive the byte check.
        if before_bytes.len() != current_bytes.len()
            || (before_bytes.as_ptr() != current_bytes.as_ptr() && before_bytes != current_bytes)
        {
            return None;
        }
        Some(self.clone())
    }

    fn source_identity(&self) -> Option<Arc<SourceIdentity>> {
        match &self.xml {
            XmlStorage::Owned(_) => None,
            XmlStorage::OwnedWithIdentity { identity, .. } => Some(Arc::clone(identity)),
            XmlStorage::Managed(xml) => Some(Arc::clone(&xml.identity)),
            XmlStorage::Source { identity, .. } => Some(Arc::clone(identity)),
        }
    }

    /// Whether this snapshot retains exactly `bytes` and carries no source
    /// identity.
    ///
    /// [`Package::document_snapshot`] builds an owned, identity-free snapshot
    /// over a copy of the opened main part, so `same_source` against such a
    /// snapshot reduces to exactly this test: equal bytes, and `None` on both
    /// identity sides. The retained bytes are immutable, so pointer identity
    /// is an exact equality proof; distinct allocations still receive the full
    /// byte comparison.
    ///
    /// [`Package::document_snapshot`]: crate::Package::document_snapshot
    fn retains_exact_unmanaged_xml(&self, bytes: &[u8]) -> bool {
        if self.xml.identity().is_some() {
            return false;
        }
        let retained = self.xml.bytes();
        retained.len() == bytes.len() && (retained.as_ptr() == bytes.as_ptr() || retained == bytes)
    }
}

/// A semantic main-document operation recorded in a reversible patch.
#[derive(Debug, Clone, PartialEq, Eq)]
#[non_exhaustive]
pub enum Operation {
    /// Replace the complete text of one direct-body paragraph.
    ReplaceParagraphText {
        /// Projected paragraph position at the time of the operation.
        position: Position,
        /// Text required before applying the operation.
        before: String,
        /// Text produced by the operation.
        after: String,
    },
    /// Replace text in one direct hyperlink while retaining its relationships.
    ReplaceHyperlinkText {
        /// Direct-body paragraph position.
        paragraph: Position,
        /// Direct hyperlink position within that paragraph.
        hyperlink: Position,
        /// Text required before applying the operation.
        before: String,
        /// Text produced by the operation.
        after: String,
    },
    /// Replace text in one direct run inside an otherwise rich paragraph.
    ReplaceRunText {
        /// Direct-body paragraph position.
        paragraph: Position,
        /// Direct run position within the paragraph.
        run: Position,
        /// Text required before applying the operation.
        before: String,
        /// Text produced by the operation.
        after: String,
    },
    /// Replace the displayed result of one direct `w:fldSimple` field.
    ReplaceSimpleFieldText {
        /// Direct-body paragraph position.
        paragraph: Position,
        /// Direct simple-field position within the paragraph.
        field: Position,
        /// Text required before applying the operation.
        before: String,
        /// Text produced by the operation.
        after: String,
    },
    /// Replace the displayed result of one run-delimited complex field.
    ReplaceComplexFieldText {
        /// Direct-body paragraph position.
        paragraph: Position,
        /// Position among complex field begin markers.
        field: Position,
        /// Text required before applying the operation.
        before: String,
        /// Text produced by the operation.
        after: String,
    },
    /// Replace inert text inside one direct tracked revision wrapper.
    ReplaceRevisionText {
        /// Direct-body paragraph position.
        paragraph: Position,
        /// Tracked wrapper family.
        kind: RevisionKind,
        /// Position among direct wrappers of the selected family.
        revision: Position,
        /// Text required before applying the operation.
        before: String,
        /// Text produced by the operation.
        after: String,
    },
    /// Apply or reject one direct inline tracked revision.
    ///
    /// The complete owning paragraph before and after bytes are retained so
    /// replay and inversion remain exact even though the selected wrapper is
    /// removed or unwrapped by the action.
    ApplyRevision {
        /// Direct paragraph/revision selector used to construct the action.
        selector: RevisionSelector,
        /// Word redline disposition requested by the caller.
        action: RevisionAction,
        /// Exact owning paragraph bytes required before replay.
        before: Arc<Vec<u8>>,
        /// Exact owning paragraph bytes produced by the action.
        after: Arc<Vec<u8>>,
    },
    /// Replace text in one direct inline content control.
    ReplaceContentControlText {
        /// Direct-body paragraph position.
        paragraph: Position,
        /// Direct inline content-control position within the paragraph.
        control: Position,
        /// Text required before applying the operation.
        before: String,
        /// Text produced by the operation.
        after: String,
    },
    /// Replace text in a nested inline content control selected by direct-child path.
    ReplaceNestedContentControlText {
        /// Direct-body paragraph position.
        paragraph: Position,
        /// Non-empty direct `w:sdt` path, first in the paragraph and then in
        /// each preceding `w:sdtContent`.
        controls: Arc<[Position]>,
        /// Text required before applying the operation.
        before: String,
        /// Text produced by the operation.
        after: String,
    },
    /// Replace a direct hyperlink inside a nested inline content control.
    ReplaceNestedContentControlHyperlinkText {
        /// Direct-body paragraph position.
        paragraph: Position,
        /// Non-empty direct content-control path.
        controls: Arc<[Position]>,
        /// Direct hyperlink position in the deepest `w:sdtContent`.
        hyperlink: Position,
        /// Text required before applying the operation.
        before: String,
        /// Text produced by the operation.
        after: String,
    },
    /// Replace one direct paragraph inside a possibly nested block control.
    ReplaceBlockContentControlParagraphText {
        /// Non-empty path starting at a direct-body `w:sdt`.
        controls: Arc<[Position]>,
        /// Direct paragraph position in the deepest `w:sdtContent`.
        paragraph: Position,
        /// Text required before applying the operation.
        before: String,
        /// Text produced by the operation.
        after: String,
    },
    /// Replace a direct hyperlink in a path-addressed block-control paragraph.
    ReplaceBlockContentControlParagraphHyperlinkText {
        /// Non-empty path starting at a direct-body `w:sdt`.
        controls: Arc<[Position]>,
        /// Direct paragraph position in the deepest `w:sdtContent`.
        paragraph: Position,
        /// Direct hyperlink position in the selected paragraph.
        hyperlink: Position,
        /// Text required before applying the operation.
        before: String,
        /// Text produced by the operation.
        after: String,
    },
    /// Replace text in one basic direct-body table cell.
    ReplaceCellText {
        /// Direct-body table position.
        table: Position,
        /// Direct row position in the table.
        row: Position,
        /// Direct cell position in the row.
        cell: Position,
        /// Text required before applying the operation.
        before: String,
        /// Text produced by the operation.
        after: String,
    },
    /// Replace one direct paragraph inside a rich or multi-paragraph cell.
    ReplaceCellParagraphText {
        /// Direct-body table position.
        table: Position,
        /// Direct row position in the table.
        row: Position,
        /// Direct cell position in the row.
        cell: Position,
        /// Direct paragraph position in the cell.
        paragraph: Position,
        /// Text required before applying the operation.
        before: String,
        /// Text produced by the operation.
        after: String,
    },
    /// Replace a direct paragraph in a cell reached through nested tables.
    ReplaceNestedCellParagraphText {
        /// Non-empty body-to-nested-cell path.
        path: Arc<[TableCellAddress]>,
        /// Direct paragraph position in the deepest cell.
        paragraph: Position,
        /// Text required before applying the operation.
        before: String,
        /// Text produced by the operation.
        after: String,
    },
    /// Replace a direct hyperlink in a path-addressed nested-table paragraph.
    ReplaceNestedCellParagraphHyperlinkText {
        /// Non-empty body-to-nested-cell path.
        path: Arc<[TableCellAddress]>,
        /// Direct paragraph position in the deepest cell.
        paragraph: Position,
        /// Direct hyperlink position in the selected paragraph.
        hyperlink: Position,
        /// Text required before applying the operation.
        before: String,
        /// Text produced by the operation.
        after: String,
    },
    /// Insert one compact plain-text paragraph.
    InsertParagraph {
        /// Projected insertion position at the time of the operation.
        position: Position,
        /// Inserted inert text.
        text: String,
    },
    /// Remove a paragraph previously inserted by the inverse patch.
    RemoveParagraph {
        /// Projected paragraph position at the time of the operation.
        position: Position,
        /// Removed inert text.
        text: String,
    },
    /// Insert one dependency-checked compact paragraph fragment.
    InsertTransferredParagraph {
        /// Projected insertion position.
        position: Position,
        /// Compact paragraph XML with receiver-local relationship references.
        xml: Arc<Vec<u8>>,
        /// Exact receiver relationship/resource inventory required at publish.
        dependency_digest: Arc<str>,
        /// Exact graph digest required by the inverse removal.
        inverse_dependency_digest: Arc<str>,
        /// Complete receiver-local dependency closure added by publication.
        graph: Arc<TransferGraph>,
    },
    /// Remove the exact transferred paragraph fragment.
    RemoveTransferredParagraph {
        /// Projected paragraph position.
        position: Position,
        /// Exact compact paragraph XML expected at the position.
        xml: Arc<Vec<u8>>,
        /// Exact receiver relationship/resource inventory required at publish.
        dependency_digest: Arc<str>,
        /// Exact graph digest required by the inverse insertion.
        inverse_dependency_digest: Arc<str>,
        /// Complete receiver-local dependency closure removed by publication.
        graph: Arc<TransferGraph>,
    },
}

impl Operation {
    pub(crate) const fn supports_source_backed_main_document_overlay(&self) -> bool {
        match self {
            Self::ReplaceParagraphText { .. }
            | Self::ReplaceHyperlinkText { .. }
            | Self::ReplaceRunText { .. }
            | Self::ReplaceSimpleFieldText { .. }
            | Self::ReplaceComplexFieldText { .. }
            | Self::ReplaceRevisionText { .. }
            | Self::ApplyRevision { .. }
            | Self::ReplaceContentControlText { .. }
            | Self::ReplaceNestedContentControlText { .. }
            | Self::ReplaceNestedContentControlHyperlinkText { .. }
            | Self::ReplaceBlockContentControlParagraphText { .. }
            | Self::ReplaceBlockContentControlParagraphHyperlinkText { .. }
            | Self::ReplaceCellText { .. }
            | Self::ReplaceCellParagraphText { .. }
            | Self::ReplaceNestedCellParagraphText { .. }
            | Self::ReplaceNestedCellParagraphHyperlinkText { .. }
            | Self::InsertParagraph { .. }
            | Self::RemoveParagraph { .. } => true,
            Self::InsertTransferredParagraph { .. } | Self::RemoveTransferredParagraph { .. } => {
                false
            },
        }
    }

    fn inverse(&self) -> Self {
        match self {
            Self::ReplaceParagraphText {
                position,
                before,
                after,
            } => Self::ReplaceParagraphText {
                position: *position,
                before: after.clone(),
                after: before.clone(),
            },
            Self::ReplaceHyperlinkText {
                paragraph,
                hyperlink,
                before,
                after,
            } => Self::ReplaceHyperlinkText {
                paragraph: *paragraph,
                hyperlink: *hyperlink,
                before: after.clone(),
                after: before.clone(),
            },
            Self::ReplaceRunText {
                paragraph,
                run,
                before,
                after,
            } => Self::ReplaceRunText {
                paragraph: *paragraph,
                run: *run,
                before: after.clone(),
                after: before.clone(),
            },
            Self::ReplaceSimpleFieldText {
                paragraph,
                field,
                before,
                after,
            } => Self::ReplaceSimpleFieldText {
                paragraph: *paragraph,
                field: *field,
                before: after.clone(),
                after: before.clone(),
            },
            Self::ReplaceComplexFieldText {
                paragraph,
                field,
                before,
                after,
            } => Self::ReplaceComplexFieldText {
                paragraph: *paragraph,
                field: *field,
                before: after.clone(),
                after: before.clone(),
            },
            Self::ReplaceRevisionText {
                paragraph,
                kind,
                revision,
                before,
                after,
            } => Self::ReplaceRevisionText {
                paragraph: *paragraph,
                kind: *kind,
                revision: *revision,
                before: after.clone(),
                after: before.clone(),
            },
            Self::ApplyRevision {
                selector,
                action,
                before,
                after,
            } => Self::ApplyRevision {
                selector: *selector,
                action: *action,
                before: Arc::clone(after),
                after: Arc::clone(before),
            },
            Self::ReplaceContentControlText {
                paragraph,
                control,
                before,
                after,
            } => Self::ReplaceContentControlText {
                paragraph: *paragraph,
                control: *control,
                before: after.clone(),
                after: before.clone(),
            },
            Self::ReplaceNestedContentControlText {
                paragraph,
                controls,
                before,
                after,
            } => Self::ReplaceNestedContentControlText {
                paragraph: *paragraph,
                controls: Arc::clone(controls),
                before: after.clone(),
                after: before.clone(),
            },
            Self::ReplaceNestedContentControlHyperlinkText {
                paragraph,
                controls,
                hyperlink,
                before,
                after,
            } => Self::ReplaceNestedContentControlHyperlinkText {
                paragraph: *paragraph,
                controls: Arc::clone(controls),
                hyperlink: *hyperlink,
                before: after.clone(),
                after: before.clone(),
            },
            Self::ReplaceBlockContentControlParagraphText {
                controls,
                paragraph,
                before,
                after,
            } => Self::ReplaceBlockContentControlParagraphText {
                controls: Arc::clone(controls),
                paragraph: *paragraph,
                before: after.clone(),
                after: before.clone(),
            },
            Self::ReplaceBlockContentControlParagraphHyperlinkText {
                controls,
                paragraph,
                hyperlink,
                before,
                after,
            } => Self::ReplaceBlockContentControlParagraphHyperlinkText {
                controls: Arc::clone(controls),
                paragraph: *paragraph,
                hyperlink: *hyperlink,
                before: after.clone(),
                after: before.clone(),
            },
            Self::ReplaceCellText {
                table,
                row,
                cell,
                before,
                after,
            } => Self::ReplaceCellText {
                table: *table,
                row: *row,
                cell: *cell,
                before: after.clone(),
                after: before.clone(),
            },
            Self::ReplaceCellParagraphText {
                table,
                row,
                cell,
                paragraph,
                before,
                after,
            } => Self::ReplaceCellParagraphText {
                table: *table,
                row: *row,
                cell: *cell,
                paragraph: *paragraph,
                before: after.clone(),
                after: before.clone(),
            },
            Self::ReplaceNestedCellParagraphText {
                path,
                paragraph,
                before,
                after,
            } => Self::ReplaceNestedCellParagraphText {
                path: Arc::clone(path),
                paragraph: *paragraph,
                before: after.clone(),
                after: before.clone(),
            },
            Self::ReplaceNestedCellParagraphHyperlinkText {
                path,
                paragraph,
                hyperlink,
                before,
                after,
            } => Self::ReplaceNestedCellParagraphHyperlinkText {
                path: Arc::clone(path),
                paragraph: *paragraph,
                hyperlink: *hyperlink,
                before: after.clone(),
                after: before.clone(),
            },
            Self::InsertParagraph { position, text } => Self::RemoveParagraph {
                position: *position,
                text: text.clone(),
            },
            Self::RemoveParagraph { position, text } => Self::InsertParagraph {
                position: *position,
                text: text.clone(),
            },
            Self::InsertTransferredParagraph {
                position,
                xml,
                dependency_digest,
                inverse_dependency_digest,
                graph,
            } => Self::RemoveTransferredParagraph {
                position: *position,
                xml: Arc::clone(xml),
                dependency_digest: Arc::clone(inverse_dependency_digest),
                inverse_dependency_digest: Arc::clone(dependency_digest),
                graph: Arc::clone(graph),
            },
            Self::RemoveTransferredParagraph {
                position,
                xml,
                dependency_digest,
                inverse_dependency_digest,
                graph,
            } => Self::InsertTransferredParagraph {
                position: *position,
                xml: Arc::clone(xml),
                dependency_digest: Arc::clone(inverse_dependency_digest),
                inverse_dependency_digest: Arc::clone(dependency_digest),
                graph: Arc::clone(graph),
            },
        }
    }
}

/// How a committed main-document edit compacts the `WordprocessingML` it
/// publishes.
///
/// The main document is the only part a document transaction rewrites, and an
/// edit changes only the body children it was asked to change. This policy
/// decides what happens to the *rest* of that part when the edit is committed.
///
/// The default is [`Self::PreserveUnmodified`]: preservation by default, the
/// rule the rest of the library already follows for parts an edit never
/// names. [`Self::WholeDocument`] is the opt-in that re-serializes the entire
/// main document, which is what every committed edit did before this policy
/// existed.
///
/// ```rust,no_run
/// use litchi_core::Position;
/// use litchi_docx::Package;
/// use litchi_docx::document::CompactionPolicy;
///
/// let mut package = Package::open("input.docx")?;
/// let mut edit = package
///     .edit_document()?
///     .with_compaction_policy(CompactionPolicy::WholeDocument);
/// edit.replace_paragraph_text(Position::new(0), "hello")?;
/// package.publish_document_edit(edit)?;
/// # Ok::<(), Box<dyn std::error::Error + Send + Sync>>(())
/// ```
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
#[non_exhaustive]
pub enum CompactionPolicy {
    /// Compact only the body paragraphs the edit changed, and publish every
    /// other byte of the main document exactly as it was read.
    ///
    /// A paragraph whose bytes differ from the source paragraph at the same
    /// position is compacted on its own, as a single closed root; a paragraph
    /// whose bytes are unchanged is never re-serialized, so its whitespace,
    /// attribute spelling and comments survive the edit. Bytes outside the
    /// direct-body paragraphs — the document prologue, tables, block content
    /// controls, the body-final section properties and the epilogue — are
    /// never re-serialized under this policy.
    ///
    /// Because unmodified markup is not re-serialized, the refusals that
    /// belong to re-serialization are not raised for it: a main document whose
    /// untouched markup carries, for example, an `xml:space` value other than
    /// `default` or `preserve` publishes under this policy and is refused
    /// under [`Self::WholeDocument`]. Every refusal that belongs to the
    /// snapshot itself — the document byte, node and depth limits, document
    /// type declarations, unbalanced nesting — is unchanged, because the
    /// snapshot scan is unchanged.
    #[default]
    PreserveUnmodified,
    /// Re-serialize the whole main document in compact form on every changed
    /// commit.
    ///
    /// This strips compactable inter-element whitespace from every paragraph,
    /// including paragraphs the edit never touched, and normalizes attribute
    /// quoting and escaping across the part. It is the behaviour every
    /// committed edit had before this policy existed.
    WholeDocument,
}

/// A staged main-document edit.
#[derive(Debug)]
pub struct Edit {
    base: Snapshot,
    projected: Snapshot,
    operations: Vec<Operation>,
    replacement_text_bytes: usize,
    operation_admission: Option<Arc<OperationAdmission>>,
    operation_string_admission: Option<Arc<StringAdmission>>,
    compaction: CompactionPolicy,
}

impl Edit {
    /// Borrow the immutable source snapshot.
    #[must_use]
    pub const fn source(&self) -> &Snapshot {
        &self.base
    }

    /// Borrow the current projected snapshot.
    #[must_use]
    pub const fn projected(&self) -> &Snapshot {
        &self.projected
    }

    /// Choose how this edit compacts the main document it publishes.
    ///
    /// The default is [`CompactionPolicy::PreserveUnmodified`], which leaves
    /// every paragraph the edit did not change byte for byte as it was read.
    #[must_use]
    pub const fn with_compaction_policy(mut self, policy: CompactionPolicy) -> Self {
        self.compaction = policy;
        self
    }

    /// Return the compaction policy this edit will publish under.
    #[must_use]
    pub const fn compaction_policy(&self) -> CompactionPolicy {
        self.compaction
    }

    /// Clone a staged edit after admitting the new operation ledger.
    ///
    /// `Edit` intentionally does not implement infallible [`Clone`]: cloning
    /// its operation vector can allocate, and managed clones must charge that
    /// allocation against the retained execution context before growing the
    /// vector. Managed clones also admit the copied operation string payloads
    /// before `String::clone` grows them.
    pub fn try_clone(&self) -> TransactionResult<Self> {
        let mut operations = Vec::new();
        let operation_admission = if let Some(admission) = self.projected.admission.as_ref() {
            if self.operations.is_empty() {
                None
            } else {
                let operation_bytes = managed_operation_memory_bytes(self.operations.len())?;
                let memory =
                    reserve_managed(&admission.context, Resource::Memory, operation_bytes)?;
                let objects =
                    reserve_managed(&admission.context, Resource::Objects, self.operations.len())?;
                Some(Arc::new(OperationAdmission {
                    _memory: memory,
                    _objects: objects,
                }))
            }
        } else {
            None
        };
        let operation_string_admission = if let Some(admission) = self.projected.admission.as_ref()
        {
            let string_bytes = self
                .operations
                .iter()
                .try_fold(0usize, |total, operation| {
                    total.checked_add(operation_string_bytes(operation)).ok_or(
                        TransactionError::Limit {
                            resource: "managed cloned operation string bytes",
                            max: usize::MAX,
                            actual: usize::MAX,
                        },
                    )
                })?;
            if string_bytes == 0 {
                None
            } else {
                let memory = string_bytes
                    .checked_add(size_of::<StringAdmission>())
                    .and_then(|bytes| bytes.checked_add(size_of::<Arc<StringAdmission>>()))
                    .ok_or(TransactionError::Limit {
                        resource: "managed cloned operation string metadata bytes",
                        max: usize::MAX,
                        actual: usize::MAX,
                    })?;
                Some(Arc::new(StringAdmission {
                    _memory: reserve_managed(&admission.context, Resource::Memory, memory)?,
                }))
            }
        } else {
            None
        };
        operations
            .try_reserve_exact(self.operations.len())
            .map_err(|source| crate::Error::Allocation {
                resource: "document operation clone",
                source,
            })?;
        operations.extend(self.operations.iter().cloned());
        Ok(Self {
            base: self.base.clone(),
            projected: self.projected.clone(),
            operations,
            replacement_text_bytes: self.replacement_text_bytes,
            operation_admission,
            operation_string_admission,
            compaction: self.compaction,
        })
    }

    /// Replace all text in a direct-body paragraph while retaining run
    /// boundaries, formatting, drawings, and unknown run XML.
    ///
    /// Replacement characters are assigned to the existing text slots in
    /// order: each slot keeps up to its original character count and the final
    /// slot receives any remainder. Direct hyperlinks and other paragraph
    /// wrappers use their focused operations and are refused here.
    /// The input must expose its borrowed `str` view as well as a `String`
    /// conversion so managed callers can be admitted before conversion.
    ///
    /// # Errors
    ///
    /// Returns a typed refusal, checked-position error, resource-limit error,
    /// or malformed-document error without changing the projected snapshot.
    pub fn replace_paragraph_text(
        &mut self,
        position: Position,
        authored_text: impl Into<String> + AsRef<str>,
    ) -> TransactionResult<&mut Self> {
        if self.projected.is_managed() {
            let input_bytes = authored_text.as_ref().len();
            let input_admission = self.reserve_input_text(input_bytes)?;
            let text = authored_text.into();
            let input_capacity_admission = if text.capacity() > input_bytes {
                self.reserve_input_text(text.capacity() - input_bytes)?
            } else {
                None
            };
            return self.replace_managed_paragraph_text(
                position,
                text,
                input_admission,
                input_capacity_admission,
            );
        }
        self.reserve_operation()?;
        let text = authored_text.into();
        validate_authored_text(&text).map_err(|reason| TransactionError::Refused {
            position: position.get(),
            reason,
        })?;
        let range = self.range(position)?;
        let paragraph_start = usize::try_from(range.start).map_err(|_conversion_error| {
            crate::Error::InvalidFormat("paragraph offset does not fit usize".into())
        })?;
        let paragraph_end = paragraph_start
            .checked_add(usize::try_from(range.length).map_err(|_conversion_error| {
                crate::Error::InvalidFormat("paragraph length does not fit usize".into())
            })?)
            .ok_or_else(|| crate::Error::InvalidFormat("paragraph range overflow".into()))?;
        let paragraph = self
            .projected
            .xml_bytes()
            .get(paragraph_start..paragraph_end)
            .ok_or_else(|| crate::Error::InvalidFormat("paragraph range is outside XML".into()))?;
        let owner =
            scan_text_owner(paragraph, b"p").map_err(|reason| TransactionError::Refused {
                position: position.get(),
                reason,
            })?;
        if owner.text == text {
            return Ok(self);
        }
        let replacement_text_bytes = self.checked_text_total(text.len())?;
        let replacement = rewrite_text_owner(paragraph, &owner, &text)?;
        let xml = replace_range(
            self.projected.xml_bytes(),
            paragraph_start,
            paragraph_end,
            &replacement,
        )?;
        let candidate = self.projected.with_spliced_paragraphs(
            xml,
            &[ParagraphSplice {
                start: paragraph_start,
                end: paragraph_end,
                replacement: &replacement,
            }],
        )?;
        let readback = candidate
            .paragraph(position)
            .ok_or(TransactionError::OutOfBounds {
                position: position.get(),
                len: candidate.paragraph_count(),
            })?
            .text()?;
        if readback != text {
            return Err(crate::Error::InvalidFormat(
                "document text edit failed semantic readback".into(),
            )
            .into());
        }
        self.operations.push(Operation::ReplaceParagraphText {
            position,
            before: owner.text,
            after: text,
        });
        self.replacement_text_bytes = replacement_text_bytes;
        self.projected = candidate;
        Ok(self)
    }

    fn replace_managed_paragraph_text(
        &mut self,
        position: Position,
        authored_text: String,
        input_admission: Option<Arc<StringAdmission>>,
        input_capacity_admission: Option<Arc<StringAdmission>>,
    ) -> TransactionResult<&mut Self> {
        let context = self.base.managed_context().ok_or({
            TransactionError::Document(crate::Error::UnsafeEdit {
                format: "DOCX",
                operation: "replace_paragraph_text",
                reason: "managed candidate is missing its execution admission",
            })
        })?;
        context.check().map_err(managed_execution)?;
        validate_authored_text(&authored_text).map_err(|reason| TransactionError::Refused {
            position: position.get(),
            reason,
        })?;
        let range = self.range(position)?;
        let paragraph_start = usize::try_from(range.start).map_err(|_conversion_error| {
            crate::Error::InvalidFormat("paragraph offset does not fit usize".into())
        })?;
        let paragraph_end = paragraph_start
            .checked_add(usize::try_from(range.length).map_err(|_conversion_error| {
                crate::Error::InvalidFormat("paragraph length does not fit usize".into())
            })?)
            .ok_or_else(|| crate::Error::InvalidFormat("paragraph range overflow".into()))?;
        let paragraph = self
            .projected
            .xml_bytes()
            .get(paragraph_start..paragraph_end)
            .ok_or_else(|| crate::Error::InvalidFormat("paragraph range is outside XML".into()))?;
        let mut text_scan_admission = self
            .projected
            .admission
            .as_ref()
            .map(|admission| admission.owner_scan(paragraph.len()))
            .transpose()?;
        let owner =
            scan_text_owner(paragraph, b"p").map_err(|reason| TransactionError::Refused {
                position: position.get(),
                reason,
            })?;
        if let Some(admission) = text_scan_admission.as_mut() {
            admission.release_depth();
        }
        if owner.text == authored_text {
            drop((
                input_admission,
                input_capacity_admission,
                text_scan_admission,
            ));
            return Ok(self);
        }

        // Every managed replacement is planned against the immutable base
        // source. A derived SourceXmlPart is deliberately unable to issue a
        // second proof, so rebuilding from `self.base` keeps all proofs tied
        // to one original payload even when callers edit disjoint paragraphs
        // in arbitrary order.
        let base_range = self.base.range(position)?;
        let base_paragraph_start =
            usize::try_from(base_range.start).map_err(|_conversion_error| {
                crate::Error::InvalidFormat("base paragraph offset does not fit usize".into())
            })?;
        let base_paragraph_end = base_paragraph_start
            .checked_add(
                usize::try_from(base_range.length).map_err(|_conversion_error| {
                    crate::Error::InvalidFormat("base paragraph length does not fit usize".into())
                })?,
            )
            .ok_or_else(|| crate::Error::InvalidFormat("base paragraph range overflow".into()))?;
        let base_paragraph = self
            .base
            .xml_bytes()
            .get(base_paragraph_start..base_paragraph_end)
            .ok_or_else(|| {
                crate::Error::InvalidFormat("base paragraph range is outside XML".into())
            })?;
        let mut base_scan_admission = self
            .base
            .admission
            .as_ref()
            .map(|admission| admission.owner_scan(base_paragraph.len()))
            .transpose()?;
        let base_owner =
            scan_text_owner(base_paragraph, b"p").map_err(|reason| TransactionError::Refused {
                position: position.get(),
                reason,
            })?;
        if let Some(admission) = base_scan_admission.as_mut() {
            admission.release_depth();
        }
        let TextOwner {
            text: base_text, ..
        } = base_owner;

        let mut existing_index = None;
        for (index, operation) in self.operations.iter().enumerate() {
            let Operation::ReplaceParagraphText {
                position: operation_position,
                before,
                ..
            } = operation
            else {
                return Err(TransactionError::Document(crate::Error::UnsafeEdit {
                    format: "DOCX",
                    operation: "replace_paragraph_text",
                    reason: "managed transactions only support direct paragraph text operations",
                }));
            };
            if *operation_position == position {
                if existing_index.replace(index).is_some() {
                    return Err(crate::Error::InvalidFormat(
                        "managed paragraph replacement plan contains a duplicate position".into(),
                    )
                    .into());
                }
                if before != &base_text {
                    return Err(TransactionError::StaleSource);
                }
            }
        }
        if existing_index.is_none() && owner.text != base_text {
            return Err(crate::Error::InvalidFormat(
                "managed paragraph replacement plan lost its base operation".into(),
            )
            .into());
        }

        let remove_existing = existing_index.is_some() && authored_text == base_text;
        let add_new = existing_index.is_none() && authored_text != base_text;
        let final_count = self
            .operations
            .len()
            .checked_sub(usize::from(remove_existing))
            .and_then(|count| count.checked_add(usize::from(add_new)))
            .ok_or(TransactionError::Limit {
                resource: "managed paragraph operation count",
                max: MAX_OPERATIONS,
                actual: usize::MAX,
            })?;
        if final_count > MAX_OPERATIONS {
            return Err(TransactionError::Limit {
                resource: "managed paragraph operation count",
                max: MAX_OPERATIONS,
                actual: final_count,
            });
        }

        let mut planned_string_bytes = 0usize;
        for (index, operation) in self.operations.iter().enumerate() {
            if Some(index) == existing_index {
                if remove_existing {
                    continue;
                }
                planned_string_bytes = planned_string_bytes
                    .checked_add(base_text.capacity())
                    .and_then(|bytes| bytes.checked_add(authored_text.capacity()))
                    .ok_or(TransactionError::Limit {
                        resource: "managed operation string bytes",
                        max: MAX_REPLACEMENT_TEXT_BYTES,
                        actual: usize::MAX,
                    })?;
            } else {
                planned_string_bytes = planned_string_bytes
                    .checked_add(operation_string_bytes(operation))
                    .ok_or(TransactionError::Limit {
                        resource: "managed operation string bytes",
                        max: MAX_REPLACEMENT_TEXT_BYTES,
                        actual: usize::MAX,
                    })?;
            }
        }
        if add_new {
            planned_string_bytes = planned_string_bytes
                .checked_add(base_text.capacity())
                .and_then(|bytes| bytes.checked_add(authored_text.capacity()))
                .ok_or(TransactionError::Limit {
                    resource: "managed operation string bytes",
                    max: MAX_REPLACEMENT_TEXT_BYTES,
                    actual: usize::MAX,
                })?;
        }

        if final_count == 0 {
            // Reverting the last staged operation restores the exact base
            // SourceXmlPart, including all untouched source bytes and its
            // original source proof. Drop the operation ledger as well so a
            // subsequent edit starts with a fresh, correctly sized budget.
            self.operations = Vec::new();
            self.operation_admission = None;
            self.operation_string_admission = None;
            self.replacement_text_bytes = 0;
            self.projected = self.base.clone();
            drop((
                input_admission,
                input_capacity_admission,
                text_scan_admission,
                base_scan_admission,
            ));
            return Ok(self);
        }

        let source = self.base.source_xml().ok_or(crate::Error::UnsafeEdit {
            format: "DOCX",
            operation: "replace_paragraph_text",
            reason: "managed paragraph source authorization is unavailable",
        })?;
        let identity = self
            .base
            .source_identity()
            .ok_or(crate::Error::UnsafeEdit {
                format: "DOCX",
                operation: "replace_paragraph_text",
                reason: "managed paragraph source identity is unavailable",
            })?;

        // Admit the replacement ledger and all retained operation strings
        // before cloning or growing the prospective operation vector. The
        // previous ledger remains live until the candidate succeeds, which
        // keeps this method atomic if any later proof or parser check fails.
        let operation_admission = self.reserve_managed_operation_admission(final_count)?;
        let operation_string_admission = if planned_string_bytes == 0 {
            None
        } else {
            Some(self.reserve_string_admission(planned_string_bytes)?)
        };
        let mut planned_operations = Vec::new();
        planned_operations
            .try_reserve_exact(final_count)
            .map_err(|source| crate::Error::Allocation {
                resource: "managed document operation metadata",
                source,
            })?;
        let mut authored_text = Some(authored_text);
        for (index, operation) in self.operations.iter().enumerate() {
            let Operation::ReplaceParagraphText {
                position: operation_position,
                before,
                after,
            } = operation
            else {
                return Err(TransactionError::Document(crate::Error::UnsafeEdit {
                    format: "DOCX",
                    operation: "replace_paragraph_text",
                    reason: "managed transactions only support direct paragraph text operations",
                }));
            };
            if Some(index) == existing_index {
                if remove_existing {
                    continue;
                }
                let after = authored_text.take().ok_or_else(|| {
                    crate::Error::InvalidFormat(
                        "managed replacement text was consumed before its operation".into(),
                    )
                })?;
                planned_operations.push(Operation::ReplaceParagraphText {
                    position: *operation_position,
                    before: before.clone(),
                    after,
                });
            } else {
                planned_operations.push(Operation::ReplaceParagraphText {
                    position: *operation_position,
                    before: before.clone(),
                    after: after.clone(),
                });
            }
        }
        if add_new {
            let after = authored_text.take().ok_or_else(|| {
                crate::Error::InvalidFormat(
                    "managed replacement text was consumed before insertion".into(),
                )
            })?;
            planned_operations.push(Operation::ReplaceParagraphText {
                position,
                before: base_text,
                after,
            });
        }
        drop((authored_text, base_scan_admission));
        debug_assert_eq!(planned_operations.len(), final_count);
        self.reconstruct_managed_paragraph_operations(
            planned_operations,
            operation_admission,
            operation_string_admission,
            "replace_paragraph_text",
            (source, identity),
            (
                input_admission,
                input_capacity_admission,
                text_scan_admission,
            ),
        )
    }

    fn reconstruct_managed_paragraph_operations(
        &mut self,
        planned_operations: Vec<Operation>,
        operation_admission: Arc<OperationAdmission>,
        operation_string_admission: Option<Arc<StringAdmission>>,
        operation: &'static str,
        (source, identity): (SourceXmlPart, Arc<SourceIdentity>),
        _temporary_admissions: impl Sized,
    ) -> TransactionResult<&mut Self> {
        let context = self.base.managed_context().ok_or({
            TransactionError::Document(crate::Error::UnsafeEdit {
                format: "DOCX",
                operation,
                reason: "managed candidate is missing its execution admission",
            })
        })?;
        context.check().map_err(managed_execution)?;
        debug_assert!(!planned_operations.is_empty());
        let replacement_text_bytes = managed_replacement_text_bytes(&planned_operations)?;

        let plan_memory = planned_operations
            .len()
            .checked_mul(size_of::<ManagedParagraphPlan>())
            .and_then(|bytes| bytes.checked_add(size_of::<Vec<ManagedParagraphPlan>>()));
        let plan_memory = plan_memory.ok_or(TransactionError::Limit {
            resource: "managed paragraph reconstruction plan memory",
            max: usize::MAX,
            actual: usize::MAX,
        })?;
        let _plan_memory = reserve_managed(&context, Resource::Memory, plan_memory)?;
        let _plan_objects = reserve_managed(&context, Resource::Objects, planned_operations.len())?;
        let mut plans = Vec::new();
        plans
            .try_reserve_exact(planned_operations.len())
            .map_err(|source| crate::Error::Allocation {
                resource: "managed paragraph reconstruction plan",
                source,
            })?;
        let mut fragment_bytes = 0usize;
        let mut fragment_count = 0usize;
        for (operation_index, planned_operation) in planned_operations.iter().enumerate() {
            let Operation::ReplaceParagraphText {
                position: operation_position,
                before,
                after,
            } = planned_operation
            else {
                return Err(TransactionError::Document(crate::Error::UnsafeEdit {
                    format: "DOCX",
                    operation,
                    reason: "managed transactions only support direct paragraph text operations",
                }));
            };
            let range = self.base.range(*operation_position)?;
            let paragraph_start = checked_start(range, "base paragraph")?;
            let paragraph_end = checked_end(range, "base paragraph")?;
            let paragraph = checked_slice(
                self.base.xml_bytes(),
                paragraph_start,
                paragraph_end,
                "base paragraph",
            )?;
            let base_admission = self.base.admission.as_ref().ok_or({
                TransactionError::Document(crate::Error::UnsafeEdit {
                    format: "DOCX",
                    operation,
                    reason: "managed base snapshot is missing its execution admission",
                })
            })?;
            let mut scan_admission = base_admission.owner_scan(paragraph.len())?;
            let owner =
                scan_text_owner(paragraph, b"p").map_err(|reason| TransactionError::Refused {
                    position: operation_position.get(),
                    reason,
                })?;
            scan_admission.release_depth();
            if owner.text != *before {
                return Err(TransactionError::StaleSource);
            }
            let (_, owner_fragment_bytes) = preflight_text_owner_rewrite(paragraph, &owner, after)?;
            fragment_bytes = fragment_bytes.checked_add(owner_fragment_bytes).ok_or(
                TransactionError::Limit {
                    resource: "managed paragraph fragment bytes",
                    max: MAX_DOCUMENT_XML_BYTES,
                    actual: usize::MAX,
                },
            )?;
            fragment_count =
                fragment_count
                    .checked_add(owner.slots.len())
                    .ok_or(TransactionError::Limit {
                        resource: "managed paragraph fragment count",
                        max: MAX_OPERATIONS,
                        actual: usize::MAX,
                    })?;
            plans.push(ManagedParagraphPlan {
                range,
                operation_index,
                owner,
                _scan: scan_admission,
            });
        }
        plans.sort_unstable_by_key(|plan| (plan.range.start, plan.range.length));
        for pair in plans.windows(2) {
            let previous_end = usize::try_from(pair[0].range.start)
                .ok()
                .and_then(|start| {
                    usize::try_from(pair[0].range.length)
                        .ok()
                        .and_then(|len| start.checked_add(len))
                })
                .ok_or_else(|| {
                    crate::Error::InvalidFormat("managed paragraph range overflows".into())
                })?;
            if previous_end
                > usize::try_from(pair[1].range.start).map_err(|_conversion_error| {
                    crate::Error::InvalidFormat(
                        "managed paragraph offset does not fit usize".into(),
                    )
                })?
            {
                return Err(crate::Error::InvalidFormat(
                    "managed paragraph replacement ranges overlap".into(),
                )
                .into());
            }
        }

        let fragment_memory = fragment_bytes
            .checked_add(
                fragment_count
                    .checked_mul(
                        size_of::<Vec<u8>>() + size_of::<litchi_opc::AuthoredXmlFragment>(),
                    )
                    .ok_or(TransactionError::Limit {
                        resource: "managed paragraph fragment metadata",
                        max: usize::MAX,
                        actual: usize::MAX,
                    })?,
            )
            .ok_or(TransactionError::Limit {
                resource: "managed paragraph fragment memory",
                max: usize::MAX,
                actual: usize::MAX,
            })?;
        let _fragment_memory = reserve_managed(&context, Resource::Memory, fragment_memory)?;
        let reconstruction_work = self
            .base
            .xml_bytes()
            .len()
            .checked_add(fragment_bytes)
            .and_then(|bytes| {
                planned_operations
                    .len()
                    .checked_mul(size_of::<Operation>())
                    .and_then(|operation_bytes| bytes.checked_add(operation_bytes))
            })
            .ok_or(TransactionError::Limit {
                resource: "managed paragraph reconstruction work",
                max: usize::MAX,
                actual: usize::MAX,
            })?;
        consume_managed(&context, Resource::Work, reconstruction_work)?;

        let source_for_proofs = source.clone();
        let mut publication = source.into_publication().map_err(crate::Error::from)?;
        let base_xml = self.base.xml_bytes();
        for plan in &plans {
            let Operation::ReplaceParagraphText { after, .. } =
                &planned_operations[plan.operation_index]
            else {
                return Err(crate::Error::InvalidFormat(
                    "managed paragraph reconstruction plan contains a different operation".into(),
                )
                .into());
            };
            let paragraph_start = checked_start(plan.range, "base paragraph")?;
            let total_characters = after.chars().count();
            let mut characters = after.char_indices();
            let mut character_cursor = 0usize;
            let mut byte_cursor = 0usize;
            for (index, slot) in plan.owner.slots.iter().enumerate() {
                let remaining =
                    total_characters
                        .checked_sub(character_cursor)
                        .ok_or_else(|| {
                            crate::Error::InvalidFormat(
                                "paragraph text slot cursor exceeded authored text".into(),
                            )
                        })?;
                let count = if index + 1 == plan.owner.slots.len() {
                    remaining
                } else {
                    slot.characters.min(remaining)
                };
                let value_start = byte_cursor;
                for _ in 0..count {
                    let (_, character) = characters.next().ok_or_else(|| {
                        crate::Error::InvalidFormat(
                            "paragraph text slot is outside authored text".into(),
                        )
                    })?;
                    byte_cursor = byte_cursor.checked_add(character.len_utf8()).ok_or(
                        TransactionError::Limit {
                            resource: "replacement text bytes",
                            max: MAX_REPLACEMENT_TEXT_BYTES,
                            actual: usize::MAX,
                        },
                    )?;
                }
                character_cursor =
                    character_cursor
                        .checked_add(count)
                        .ok_or(TransactionError::Limit {
                            resource: "replacement text characters",
                            max: MAX_REPLACEMENT_TEXT_BYTES,
                            actual: usize::MAX,
                        })?;
                let value = &after[value_start..byte_cursor];
                let fragment = try_run_content_fragment(&slot.prefix, &slot.local_name, value)?;
                let start = paragraph_start.checked_add(slot.start).ok_or_else(|| {
                    crate::Error::InvalidFormat("paragraph text source range overflows".into())
                })?;
                let end = paragraph_start.checked_add(slot.end).ok_or_else(|| {
                    crate::Error::InvalidFormat("paragraph text source range overflows".into())
                })?;
                let expected = base_xml.get(start..end).ok_or_else(|| {
                    crate::Error::InvalidFormat("text source range is outside XML".into())
                })?;
                let proof = source_for_proofs
                    .checked_range(start..end, expected)
                    .map_err(crate::Error::from)?;
                let fragment = litchi_opc::AuthoredXmlFragment::markup_with_execution_context(
                    fragment, &context,
                )
                .map_err(crate::Error::from)?;
                publication
                    .replace(proof, fragment)
                    .map_err(crate::Error::from)?;
            }
        }
        let candidate_source = publication.finish().map_err(crate::Error::from)?;
        let candidate =
            Snapshot::from_source_xml_with_identity(candidate_source, identity, context)?;
        for plan in &plans {
            let Operation::ReplaceParagraphText {
                position: operation_position,
                after,
                ..
            } = &planned_operations[plan.operation_index]
            else {
                return Err(crate::Error::InvalidFormat(
                    "managed paragraph reconstruction plan contains a different operation".into(),
                )
                .into());
            };
            let readback = candidate
                .paragraph(*operation_position)
                .ok_or(TransactionError::OutOfBounds {
                    position: operation_position.get(),
                    len: candidate.paragraph_count(),
                })?
                .text()?;
            if readback != *after {
                return Err(crate::Error::InvalidFormat(
                    "managed document text edit failed semantic readback".into(),
                )
                .into());
            }
        }
        self.operations = planned_operations;
        self.operation_admission = Some(operation_admission);
        self.operation_string_admission = operation_string_admission;
        self.replacement_text_bytes = replacement_text_bytes;
        self.projected = candidate;
        Ok(self)
    }

    /// Atomically replace complete text across direct-body paragraphs.
    ///
    /// Positions must be non-empty, unique, and strictly increasing. Every
    /// paragraph retains its run boundaries, formatting, drawings, and unknown
    /// run XML under the same refusal rules as [`Self::replace_paragraph_text`].
    ///
    /// # Errors
    ///
    /// Returns a typed ambiguity refusal for an empty, duplicate, or
    /// non-canonical selector list, or a checked leaf/resource failure without
    /// changing the edit.
    pub fn replace_body_paragraph_texts(
        &mut self,
        replacements: &[ParagraphTextReplacement],
    ) -> TransactionResult<&mut Self> {
        if self.projected.is_managed() {
            validate_paragraph_replacements(replacements)?;
            return self.replace_managed_body_paragraph_texts(replacements);
        }
        self.ensure_unmanaged("replace_body_paragraph_texts")?;
        validate_paragraph_replacements(replacements)?;
        let mut candidate = self.try_clone()?;
        let first_operation = candidate.operations.len();
        let mut ranges = Vec::new();
        ranges
            .try_reserve_exact(replacements.len())
            .map_err(|allocation_error| crate::Error::Allocation {
                resource: "document paragraph replacement plan",
                source: allocation_error,
            })?;
        for replacement in replacements {
            candidate.reserve_operation()?;
            validate_authored_text(replacement.text.as_str()).map_err(|reason| {
                TransactionError::Refused {
                    position: replacement.position.get(),
                    reason,
                }
            })?;
            let replacement_text_bytes = candidate.checked_text_total(replacement.text.len())?;
            let range = candidate.range(replacement.position)?;
            let paragraph_start = checked_start(range, "paragraph")?;
            let paragraph_end = checked_end(range, "paragraph")?;
            let paragraph = checked_slice(
                candidate.projected.xml_bytes(),
                paragraph_start,
                paragraph_end,
                "paragraph",
            )?;
            let owner =
                scan_text_owner(paragraph, b"p").map_err(|reason| TransactionError::Refused {
                    position: replacement.position.get(),
                    reason,
                })?;
            if owner.text == replacement.text {
                continue;
            }
            ranges.push((
                paragraph_start,
                paragraph_end,
                rewrite_text_owner(paragraph, &owner, replacement.text.as_str())?,
            ));
            candidate.operations.push(Operation::ReplaceParagraphText {
                position: replacement.position,
                before: owner.text,
                after: replacement.text.clone(),
            });
            candidate.replacement_text_bytes = replacement_text_bytes;
        }
        if ranges.is_empty() {
            *self = candidate;
            return Ok(self);
        }
        let mut splices = Vec::new();
        splices
            .try_reserve_exact(ranges.len())
            .map_err(|allocation_error| crate::Error::Allocation {
                resource: "document paragraph splice plan",
                source: allocation_error,
            })?;
        splices.extend(
            ranges
                .iter()
                .map(|(start, end, replacement)| ParagraphSplice {
                    start: *start,
                    end: *end,
                    replacement: replacement.as_slice(),
                }),
        );
        let xml = replace_ranges(candidate.projected.xml_bytes(), &ranges)?;
        let projected = candidate.projected.with_spliced_paragraphs(xml, &splices)?;
        for operation in &candidate.operations[first_operation..] {
            let Operation::ReplaceParagraphText {
                position, after, ..
            } = operation
            else {
                return Err(crate::Error::InvalidFormat(
                    "paragraph replacement plan contains a different operation".into(),
                )
                .into());
            };
            let readback = projected
                .paragraph(*position)
                .ok_or(TransactionError::OutOfBounds {
                    position: position.get(),
                    len: projected.paragraph_count(),
                })?
                .text()?;
            if readback != *after {
                return Err(crate::Error::InvalidFormat(
                    "document paragraph batch failed semantic readback".into(),
                )
                .into());
            }
        }
        candidate.projected = projected;
        *self = candidate;
        Ok(self)
    }

    fn replace_managed_body_paragraph_texts(
        &mut self,
        replacements: &[ParagraphTextReplacement],
    ) -> TransactionResult<&mut Self> {
        let context = self.base.managed_context().ok_or({
            TransactionError::Document(crate::Error::UnsafeEdit {
                format: "DOCX",
                operation: "replace_body_paragraph_texts",
                reason: "managed candidate is missing its execution admission",
            })
        })?;
        context.check().map_err(managed_execution)?;

        // The caller owns the replacement strings, but the managed operation
        // must still admit every borrowed length and excess capacity before
        // any decision or operation vector can grow. Keep each token live
        // until the shared reconstruction either publishes or fails.
        let input_token_capacity =
            replacements
                .len()
                .checked_mul(2)
                .ok_or(TransactionError::Limit {
                    resource: "managed batch input admission capacity",
                    max: usize::MAX,
                    actual: usize::MAX,
                })?;
        let input_vector_memory = input_token_capacity
            .checked_mul(size_of::<Arc<StringAdmission>>())
            .and_then(|bytes| bytes.checked_add(size_of::<Vec<Arc<StringAdmission>>>()))
            .ok_or(TransactionError::Limit {
                resource: "managed batch input admission metadata",
                max: usize::MAX,
                actual: usize::MAX,
            })?;
        let _input_vector_memory =
            reserve_managed(&context, Resource::Memory, input_vector_memory)?;
        let mut input_admissions = Vec::new();
        input_admissions
            .try_reserve_exact(input_token_capacity)
            .map_err(|source| crate::Error::Allocation {
                resource: "managed batch input admissions",
                source,
            })?;
        for replacement in replacements {
            input_admissions.push(self.reserve_string_admission(replacement.text.len())?);
            if replacement.text.capacity() > replacement.text.len() {
                input_admissions.push(self.reserve_string_admission(
                    replacement.text.capacity() - replacement.text.len(),
                )?);
            }
        }
        for replacement in replacements {
            validate_authored_text(replacement.text.as_str()).map_err(|reason| {
                TransactionError::Refused {
                    position: replacement.position.get(),
                    reason,
                }
            })?;
        }

        // Decisions retain the immutable-base text and its scan admission
        // until the final ledger is built. Admit that bounded vector before
        // retaining either the scan result or its owned text.
        let decision_memory = replacements
            .len()
            .checked_mul(size_of::<ManagedBatchDecision<'_>>())
            .and_then(|bytes| bytes.checked_add(size_of::<Vec<ManagedBatchDecision<'_>>>()))
            .ok_or(TransactionError::Limit {
                resource: "managed paragraph batch decision metadata",
                max: usize::MAX,
                actual: usize::MAX,
            })?;
        let _decision_memory = reserve_managed(&context, Resource::Memory, decision_memory)?;
        let _decision_objects = reserve_managed(&context, Resource::Objects, replacements.len())?;
        let mut decisions = Vec::new();
        decisions
            .try_reserve_exact(replacements.len())
            .map_err(|source| crate::Error::Allocation {
                resource: "managed paragraph batch decisions",
                source,
            })?;

        for replacement in replacements {
            context.check().map_err(managed_execution)?;
            let range = self.range(replacement.position)?;
            let paragraph_start = checked_start(range, "paragraph")?;
            let paragraph_end = checked_end(range, "paragraph")?;
            let paragraph = checked_slice(
                self.projected.xml_bytes(),
                paragraph_start,
                paragraph_end,
                "paragraph",
            )?;
            let (owner, _projected_scan) = {
                let admission = self.projected.admission.as_ref().ok_or({
                    TransactionError::Document(crate::Error::UnsafeEdit {
                        format: "DOCX",
                        operation: "replace_body_paragraph_texts",
                        reason: "managed projected snapshot is missing its execution admission",
                    })
                })?;
                let mut scan = admission.owner_scan(paragraph.len())?;
                let owner = scan_text_owner(paragraph, b"p").map_err(|reason| {
                    TransactionError::Refused {
                        position: replacement.position.get(),
                        reason,
                    }
                })?;
                scan.release_depth();
                (owner, scan)
            };
            if owner.text == replacement.text {
                continue;
            }

            let base_range = self.base.range(replacement.position)?;
            let base_start = checked_start(base_range, "base paragraph")?;
            let base_end = checked_end(base_range, "base paragraph")?;
            let base_paragraph = checked_slice(
                self.base.xml_bytes(),
                base_start,
                base_end,
                "base paragraph",
            )?;
            let base_admission = self.base.admission.as_ref().ok_or({
                TransactionError::Document(crate::Error::UnsafeEdit {
                    format: "DOCX",
                    operation: "replace_body_paragraph_texts",
                    reason: "managed base snapshot is missing its execution admission",
                })
            })?;
            let mut base_scan = base_admission.owner_scan(base_paragraph.len())?;
            let base_owner = scan_text_owner(base_paragraph, b"p").map_err(|reason| {
                TransactionError::Refused {
                    position: replacement.position.get(),
                    reason,
                }
            })?;
            base_scan.release_depth();
            let base_text = base_owner.text;

            let mut existing_index = None;
            for (index, operation) in self.operations.iter().enumerate() {
                let Operation::ReplaceParagraphText {
                    position: operation_position,
                    before,
                    ..
                } = operation
                else {
                    return Err(TransactionError::Document(crate::Error::UnsafeEdit {
                        format: "DOCX",
                        operation: "replace_body_paragraph_texts",
                        reason: "managed transactions only support direct paragraph text operations",
                    }));
                };
                if *operation_position == replacement.position {
                    if existing_index.replace(index).is_some() {
                        return Err(crate::Error::InvalidFormat(
                            "managed paragraph replacement plan contains a duplicate position"
                                .into(),
                        )
                        .into());
                    }
                    if before != &base_text {
                        return Err(TransactionError::StaleSource);
                    }
                }
            }
            if existing_index.is_none() && owner.text != base_text {
                return Err(crate::Error::InvalidFormat(
                    "managed paragraph replacement plan lost its base operation".into(),
                )
                .into());
            }

            let remove_existing = existing_index.is_some() && replacement.text == base_text;
            let add_new = existing_index.is_none() && replacement.text != base_text;
            decisions.push(ManagedBatchDecision {
                replacement,
                base_text,
                existing_index,
                remove_existing,
                add_new,
                _scan: base_scan,
            });
        }

        if decisions.is_empty() {
            drop((decisions, input_admissions));
            return Ok(self);
        }

        let remove_count = decisions
            .iter()
            .filter(|decision| decision.remove_existing)
            .count();
        let add_count = decisions.iter().filter(|decision| decision.add_new).count();
        let final_count = self
            .operations
            .len()
            .checked_sub(remove_count)
            .and_then(|count| count.checked_add(add_count))
            .ok_or(TransactionError::Limit {
                resource: "managed paragraph operation count",
                max: MAX_OPERATIONS,
                actual: usize::MAX,
            })?;
        if final_count > MAX_OPERATIONS {
            return Err(TransactionError::Limit {
                resource: "managed paragraph operation count",
                max: MAX_OPERATIONS,
                actual: final_count,
            });
        }

        if final_count == 0 {
            drop((decisions, input_admissions));
            self.operations = Vec::new();
            self.operation_admission = None;
            self.operation_string_admission = None;
            self.replacement_text_bytes = 0;
            self.projected = self.base.clone();
            return Ok(self);
        }

        let source = self.base.source_xml().ok_or(crate::Error::UnsafeEdit {
            format: "DOCX",
            operation: "replace_body_paragraph_texts",
            reason: "managed paragraph source authorization is unavailable",
        })?;
        let identity = self
            .base
            .source_identity()
            .ok_or(crate::Error::UnsafeEdit {
                format: "DOCX",
                operation: "replace_body_paragraph_texts",
                reason: "managed paragraph source identity is unavailable",
            })?;

        // Existing edits may have been staged in arbitrary scalar order. Map
        // each retained operation to its batch decision once so final-ledger
        // planning stays linear in the prior ledger and selected positions.
        let decision_map_memory = self
            .operations
            .len()
            .checked_mul(size_of::<Option<usize>>())
            .and_then(|bytes| bytes.checked_add(size_of::<Vec<Option<usize>>>()))
            .ok_or(TransactionError::Limit {
                resource: "managed paragraph batch decision map metadata",
                max: usize::MAX,
                actual: usize::MAX,
            })?;
        let _decision_map_memory =
            reserve_managed(&context, Resource::Memory, decision_map_memory)?;
        let mut decision_by_operation = Vec::new();
        decision_by_operation
            .try_reserve_exact(self.operations.len())
            .map_err(|source| crate::Error::Allocation {
                resource: "managed paragraph batch decision map",
                source,
            })?;
        decision_by_operation.resize(self.operations.len(), None);
        for (decision_index, decision) in decisions.iter().enumerate() {
            if let Some(operation_index) = decision.existing_index {
                if decision_by_operation[operation_index]
                    .replace(decision_index)
                    .is_some()
                {
                    return Err(crate::Error::InvalidFormat(
                        "managed paragraph replacement plan contains a duplicate operation".into(),
                    )
                    .into());
                }
            }
        }

        let mut planned_string_bytes = 0usize;
        for (index, operation) in self.operations.iter().enumerate() {
            let decision = decision_by_operation[index]
                .and_then(|decision_index| decisions.get(decision_index));
            if let Some(decision) = decision {
                if decision.remove_existing {
                    continue;
                }
                let Operation::ReplaceParagraphText { before, .. } = operation else {
                    return Err(TransactionError::Document(crate::Error::UnsafeEdit {
                        format: "DOCX",
                        operation: "replace_body_paragraph_texts",
                        reason: "managed transactions only support direct paragraph text operations",
                    }));
                };
                planned_string_bytes = planned_string_bytes
                    .checked_add(before.capacity())
                    .and_then(|bytes| bytes.checked_add(decision.replacement.text.capacity()))
                    .ok_or(TransactionError::Limit {
                        resource: "managed operation string bytes",
                        max: MAX_REPLACEMENT_TEXT_BYTES,
                        actual: usize::MAX,
                    })?;
            } else {
                planned_string_bytes = planned_string_bytes
                    .checked_add(operation_string_bytes(operation))
                    .ok_or(TransactionError::Limit {
                        resource: "managed operation string bytes",
                        max: MAX_REPLACEMENT_TEXT_BYTES,
                        actual: usize::MAX,
                    })?;
            }
        }
        for decision in decisions.iter().filter(|decision| decision.add_new) {
            planned_string_bytes = planned_string_bytes
                .checked_add(decision.base_text.capacity())
                .and_then(|bytes| bytes.checked_add(decision.replacement.text.capacity()))
                .ok_or(TransactionError::Limit {
                    resource: "managed operation string bytes",
                    max: MAX_REPLACEMENT_TEXT_BYTES,
                    actual: usize::MAX,
                })?;
        }

        let operation_admission = self.reserve_managed_operation_admission(final_count)?;
        let operation_string_admission = if planned_string_bytes == 0 {
            None
        } else {
            Some(self.reserve_string_admission(planned_string_bytes)?)
        };
        let mut planned_operations = Vec::new();
        planned_operations
            .try_reserve_exact(final_count)
            .map_err(|source| crate::Error::Allocation {
                resource: "managed document operation metadata",
                source,
            })?;
        for (index, operation) in self.operations.iter().enumerate() {
            let decision = decision_by_operation[index]
                .and_then(|decision_index| decisions.get(decision_index));
            if let Some(decision) = decision {
                if decision.remove_existing {
                    continue;
                }
                let Operation::ReplaceParagraphText {
                    position, before, ..
                } = operation
                else {
                    return Err(TransactionError::Document(crate::Error::UnsafeEdit {
                        format: "DOCX",
                        operation: "replace_body_paragraph_texts",
                        reason: "managed transactions only support direct paragraph text operations",
                    }));
                };
                planned_operations.push(Operation::ReplaceParagraphText {
                    position: *position,
                    before: before.clone(),
                    after: decision.replacement.text.clone(),
                });
            } else {
                planned_operations.push(operation.clone());
            }
        }
        for decision in decisions.iter_mut().filter(|decision| decision.add_new) {
            let before = std::mem::take(&mut decision.base_text);
            planned_operations.push(Operation::ReplaceParagraphText {
                position: decision.replacement.position,
                before,
                after: decision.replacement.text.clone(),
            });
        }
        drop(decisions);
        drop(decision_by_operation);
        debug_assert_eq!(planned_operations.len(), final_count);

        self.reconstruct_managed_paragraph_operations(
            planned_operations,
            operation_admission,
            operation_string_admission,
            "replace_body_paragraph_texts",
            (source, identity),
            input_admissions,
        )
    }

    /// Replace all text in one direct paragraph hyperlink while leaving its
    /// anchor, tooltip, relationship, target frame, and unknown XML untouched.
    ///
    /// # Errors
    ///
    /// Returns a checked selector/refusal, resource-limit, or malformed XML
    /// error without changing the projected snapshot.
    pub fn replace_hyperlink_text(
        &mut self,
        paragraph: Position,
        hyperlink: Position,
        authored_text: impl Into<String>,
    ) -> TransactionResult<&mut Self> {
        self.ensure_unmanaged("replace_hyperlink_text")?;
        self.reserve_operation()?;
        let text = authored_text.into();
        validate_authored_text(&text).map_err(|reason| TransactionError::Refused {
            position: paragraph.get(),
            reason,
        })?;
        let replacement_text_bytes = self.checked_text_total(text.len())?;
        let paragraph_range = self.range(paragraph)?;
        let paragraph_start = checked_start(paragraph_range, "paragraph")?;
        let paragraph_end = checked_end(paragraph_range, "paragraph")?;
        let paragraph_xml = checked_slice(
            self.projected.xml_bytes(),
            paragraph_start,
            paragraph_end,
            "paragraph",
        )?;
        let hyperlink_range = select_direct_child(
            paragraph_xml,
            b"p",
            b"hyperlink",
            hyperlink,
            Refusal::HyperlinkNotFound,
        )
        .map_err(|reason| TransactionError::Refused {
            position: paragraph.get(),
            reason,
        })?;
        let hyperlink_start = checked_relative_start(paragraph_start, hyperlink_range)?;
        let hyperlink_end = checked_relative_end(paragraph_start, hyperlink_range)?;
        let hyperlink_xml = checked_slice(
            self.projected.xml_bytes(),
            hyperlink_start,
            hyperlink_end,
            "hyperlink",
        )?;
        let owner = scan_text_owner(hyperlink_xml, b"hyperlink").map_err(|reason| {
            TransactionError::Refused {
                position: paragraph.get(),
                reason,
            }
        })?;
        if owner.text == text {
            return Ok(self);
        }
        let replacement = rewrite_text_owner(hyperlink_xml, &owner, &text)?;
        let xml = replace_range(
            self.projected.xml_bytes(),
            hyperlink_start,
            hyperlink_end,
            &replacement,
        )?;
        let candidate = self.projected.with_rewritten_xml(xml)?;
        let actual = selected_hyperlink_text(&candidate, paragraph, hyperlink)?;
        if actual != text {
            return Err(crate::Error::InvalidFormat(
                "document hyperlink edit failed semantic readback".into(),
            )
            .into());
        }
        self.operations.push(Operation::ReplaceHyperlinkText {
            paragraph,
            hyperlink,
            before: owner.text,
            after: text,
        });
        self.replacement_text_bytes = replacement_text_bytes;
        self.projected = candidate;
        Ok(self)
    }

    /// Atomically replace direct hyperlinks across one or more direct-body
    /// paragraphs.
    ///
    /// Addresses must be non-empty, unique, and strictly increasing by
    /// paragraph then hyperlink position. Each leaf remains inside one direct
    /// paragraph, so this bounded composite selector cannot cross ownership or
    /// relationship scopes.
    ///
    /// # Errors
    ///
    /// Returns a typed ambiguity refusal for an empty, duplicate, or
    /// non-canonical selector list, or a resource-limit error for an oversized
    /// batch. Any checked leaf failure also leaves the entire edit unchanged.
    pub fn replace_body_hyperlink_texts(
        &mut self,
        replacements: &[HyperlinkTextReplacement],
    ) -> TransactionResult<&mut Self> {
        self.ensure_unmanaged("replace_body_hyperlink_texts")?;
        validate_hyperlink_replacements(replacements)?;
        let mut candidate = self.try_clone()?;
        for replacement in replacements {
            candidate.replace_hyperlink_text(
                replacement.address.paragraph,
                replacement.address.hyperlink,
                replacement.text.as_str(),
            )?;
        }
        *self = candidate;
        Ok(self)
    }

    /// Replace text and structural characters in one direct run while
    /// retaining its `w:rPr`, drawings, and opaque run children. Tabs, line
    /// breaks, carriage returns, non-breaking hyphens, and soft hyphens map to
    /// their native `WordprocessingML` run elements.
    ///
    /// # Errors
    ///
    /// Returns a checked selector/refusal, resource-limit, or malformed XML
    /// error without changing the projected snapshot.
    pub fn replace_run_text(
        &mut self,
        paragraph: Position,
        run: Position,
        authored_text: impl Into<String>,
    ) -> TransactionResult<&mut Self> {
        self.ensure_unmanaged("replace_run_text")?;
        self.replace_direct_paragraph_owner_text(
            paragraph,
            run,
            b"r",
            Refusal::RunNotFound,
            authored_text.into(),
            |before, after| Operation::ReplaceRunText {
                paragraph,
                run,
                before,
                after,
            },
        )
    }

    /// Replace the displayed text of one direct simple field while preserving
    /// its instruction, dirty/lock state, formatting, and opaque XML.
    ///
    /// Field instructions remain inert and are never evaluated or refreshed.
    ///
    /// # Errors
    ///
    /// Returns a checked selector/refusal, resource-limit, or malformed XML
    /// error without changing the projected snapshot.
    pub fn replace_simple_field_text(
        &mut self,
        paragraph: Position,
        field: Position,
        authored_text: impl Into<String>,
    ) -> TransactionResult<&mut Self> {
        self.ensure_unmanaged("replace_simple_field_text")?;
        self.replace_direct_paragraph_owner_text(
            paragraph,
            field,
            b"fldSimple",
            Refusal::FieldNotFound,
            authored_text.into(),
            |before, after| Operation::ReplaceSimpleFieldText {
                paragraph,
                field,
                before,
                after,
            },
        )
    }

    /// Replace the result text between one complex field's direct-run
    /// `begin`/`separate`/`end` markers. Instruction and marker runs remain
    /// exact; nested complex fields inside the selected result are refused.
    ///
    /// # Errors
    ///
    /// Returns a checked selector/refusal, resource-limit, or malformed XML
    /// error without changing the projected snapshot.
    pub fn replace_complex_field_result_text(
        &mut self,
        paragraph: Position,
        field: Position,
        authored_text: impl Into<String>,
    ) -> TransactionResult<&mut Self> {
        self.ensure_unmanaged("replace_complex_field_result_text")?;
        self.reserve_operation()?;
        let text = authored_text.into();
        validate_authored_text(&text).map_err(|reason| TransactionError::Refused {
            position: paragraph.get(),
            reason,
        })?;
        let replacement_text_bytes = self.checked_text_total(text.len())?;
        let paragraph_range = self.range(paragraph)?;
        let paragraph_start = checked_start(paragraph_range, "paragraph")?;
        let paragraph_end = checked_end(paragraph_range, "paragraph")?;
        let paragraph_xml = checked_slice(
            self.projected.xml_bytes(),
            paragraph_start,
            paragraph_end,
            "paragraph",
        )?;
        let result_range = select_complex_field_result(paragraph_xml, field).map_err(|reason| {
            TransactionError::Refused {
                position: paragraph.get(),
                reason,
            }
        })?;
        let start = checked_relative_start(paragraph_start, result_range)?;
        let end = checked_relative_end(paragraph_start, result_range)?;
        let result_xml = checked_slice(
            self.projected.xml_bytes(),
            start,
            end,
            "complex field result",
        )?;
        let (before, replacement) =
            rewrite_run_region(result_xml, &text).map_err(|reason| TransactionError::Refused {
                position: paragraph.get(),
                reason,
            })?;
        if before == text {
            return Ok(self);
        }
        let candidate = self.projected.with_rewritten_xml(replace_range(
            self.projected.xml_bytes(),
            start,
            end,
            &replacement,
        )?)?;
        let actual = selected_complex_field_result_text(&candidate, paragraph, field)?;
        if actual != text {
            return Err(crate::Error::InvalidFormat(
                "complex field result edit failed semantic readback".into(),
            )
            .into());
        }
        self.operations.push(Operation::ReplaceComplexFieldText {
            paragraph,
            field,
            before,
            after: text,
        });
        self.replacement_text_bytes = replacement_text_bytes;
        self.projected = candidate;
        Ok(self)
    }

    /// Replace inert text inside one direct tracked insertion or deletion.
    /// Revision metadata and the wrapper itself remain byte-preserved.
    ///
    /// # Errors
    ///
    /// Returns a checked selector/refusal, resource-limit, or malformed XML
    /// error without changing the projected snapshot.
    pub fn replace_revision_text(
        &mut self,
        paragraph: Position,
        kind: RevisionKind,
        revision: Position,
        authored_text: impl Into<String>,
    ) -> TransactionResult<&mut Self> {
        self.ensure_unmanaged("replace_revision_text")?;
        self.replace_direct_paragraph_owner_text(
            paragraph,
            revision,
            kind.local_name(),
            Refusal::RevisionNotFound,
            authored_text.into(),
            |before, after| Operation::ReplaceRevisionText {
                paragraph,
                kind,
                revision,
                before,
                after,
            },
        )
    }

    /// Apply one direct inline tracked revision using Word redline semantics.
    ///
    /// Insertions are accepted by unwrapping their runs and rejected by
    /// removing them. Deletions are accepted by removing them and rejected by
    /// unwrapping their runs while converting `w:delText` to ordinary
    /// `w:t`. The wrapper's revision metadata is intentionally consumed by
    /// either disposition; bytes belonging to retained children remain exact,
    /// including opaque extension markup and namespace declarations hoisted
    /// from the removed wrapper when needed.
    ///
    /// Only a direct paragraph child is selectable. Moves, range markers,
    /// property/table revisions, nested revision owners, fields, controls,
    /// hyperlinks, and other structures whose ownership or dependency closure
    /// cannot be proven are refused before the edit changes state.
    ///
    /// Like every operation other than direct paragraph text replacement, a
    /// disposition is refused on a budget-managed transaction.
    ///
    /// # Errors
    ///
    /// Returns a checked selector/refusal, resource-limit, or malformed XML
    /// error without changing the projected snapshot.
    pub fn apply_revision(
        &mut self,
        selector: RevisionSelector,
        action: RevisionAction,
    ) -> TransactionResult<&mut Self> {
        self.apply_revision_as("apply_revision", selector, action)
    }

    /// Accept one direct inline tracked insertion or deletion.
    ///
    /// # Errors
    ///
    /// Returns the same failures as [`Self::apply_revision`].
    pub fn accept_revision(
        &mut self,
        paragraph: Position,
        kind: RevisionKind,
        revision: Position,
    ) -> TransactionResult<&mut Self> {
        self.apply_revision_as(
            "accept_revision",
            RevisionSelector::new(paragraph, kind, revision),
            RevisionAction::Accept,
        )
    }

    /// Reject one direct inline tracked insertion or deletion.
    ///
    /// # Errors
    ///
    /// Returns the same failures as [`Self::apply_revision`].
    pub fn reject_revision(
        &mut self,
        paragraph: Position,
        kind: RevisionKind,
        revision: Position,
    ) -> TransactionResult<&mut Self> {
        self.apply_revision_as(
            "reject_revision",
            RevisionSelector::new(paragraph, kind, revision),
            RevisionAction::Reject,
        )
    }

    /// Apply one revision disposition for the named public operation.
    ///
    /// The managed refusal comes first. A disposition rebuilds the projection
    /// with [`Snapshot::from_xml`], which would drop a managed snapshot's
    /// source identity and budget admission and publish uncharged bytes.
    fn apply_revision_as(
        &mut self,
        operation: &'static str,
        selector: RevisionSelector,
        action: RevisionAction,
    ) -> TransactionResult<&mut Self> {
        self.ensure_unmanaged(operation)?;
        self.reserve_operation()?;
        let paragraph_range = self.range(selector.paragraph)?;
        let paragraph_start = checked_start(paragraph_range, "paragraph")?;
        let paragraph_end = checked_end(paragraph_range, "paragraph")?;
        let paragraph_xml = checked_slice(
            self.projected.xml_bytes(),
            paragraph_start,
            paragraph_end,
            "paragraph",
        )?;
        let inherited_namespaces = self
            .projected
            .paragraph_namespaces
            .get(selector.paragraph.get())
            .ok_or(TransactionError::OutOfBounds {
                position: selector.paragraph.get(),
                len: self.projected.paragraph_count(),
            })?;
        let paragraph_after =
            rewrite_revision_paragraph(paragraph_xml, selector, action, inherited_namespaces)?;
        if paragraph_after == paragraph_xml {
            return Ok(self);
        }
        let candidate = Snapshot::from_xml(replace_range(
            self.projected.xml_bytes(),
            paragraph_start,
            paragraph_end,
            &paragraph_after,
        )?)?;
        if candidate.paragraph_count() != self.projected.paragraph_count() {
            return Err(crate::Error::InvalidFormat(
                "revision action changed the paragraph count".into(),
            )
            .into());
        }
        let before = Arc::new(paragraph_xml.to_vec());
        let after = Arc::new(paragraph_after);
        self.operations.push(Operation::ApplyRevision {
            selector,
            action,
            before,
            after,
        });
        self.projected = candidate;
        Ok(self)
    }

    /// Replace text inside one direct inline content control while retaining
    /// `w:sdtPr`, data binding, lock state, wrapper metadata, and opaque XML.
    ///
    /// This operation supports an inline `w:sdtContent` whose direct children
    /// are runs; block controls and nested control structures are refused.
    ///
    /// # Errors
    ///
    /// Returns a checked selector/refusal, resource-limit, or malformed XML
    /// error without changing the projected snapshot.
    pub fn replace_content_control_text(
        &mut self,
        paragraph: Position,
        control: Position,
        authored_text: impl Into<String>,
    ) -> TransactionResult<&mut Self> {
        self.ensure_unmanaged("replace_content_control_text")?;
        self.reserve_operation()?;
        let text = authored_text.into();
        validate_authored_text(&text).map_err(|reason| TransactionError::Refused {
            position: paragraph.get(),
            reason,
        })?;
        let replacement_text_bytes = self.checked_text_total(text.len())?;
        let content = select_content_control_content(&self.projected, paragraph, control)?;
        let content_xml = checked_slice(
            self.projected.xml_bytes(),
            content.0,
            content.1,
            "content control content",
        )?;
        let owner = scan_text_owner(content_xml, b"sdtContent").map_err(|reason| {
            TransactionError::Refused {
                position: paragraph.get(),
                reason,
            }
        })?;
        if owner.text == text {
            return Ok(self);
        }
        let replacement = rewrite_text_owner(content_xml, &owner, &text)?;
        let candidate = self.projected.with_rewritten_xml(replace_range(
            self.projected.xml_bytes(),
            content.0,
            content.1,
            &replacement,
        )?)?;
        let actual = selected_content_control_text(&candidate, paragraph, control)?;
        if actual != text {
            return Err(crate::Error::InvalidFormat(
                "document content-control edit failed semantic readback".into(),
            )
            .into());
        }
        self.operations.push(Operation::ReplaceContentControlText {
            paragraph,
            control,
            before: owner.text,
            after: text,
        });
        self.replacement_text_bytes = replacement_text_bytes;
        self.projected = candidate;
        Ok(self)
    }

    /// Replace text in the deepest inline content control reached by a
    /// non-empty direct-child path. The first position selects a direct
    /// paragraph `w:sdt`; later positions select a direct `w:sdt` in the
    /// preceding `w:sdtContent`.
    ///
    /// # Errors
    ///
    /// Returns a checked path/refusal, resource-limit, or malformed XML error
    /// without changing the projected snapshot.
    pub fn replace_nested_content_control_text(
        &mut self,
        paragraph: Position,
        controls: &[Position],
        authored_text: impl Into<String>,
    ) -> TransactionResult<&mut Self> {
        self.ensure_unmanaged("replace_nested_content_control_text")?;
        let path: Arc<[Position]> = controls.into();
        let (start, end) = select_nested_inline_control_content(&self.projected, paragraph, &path)?;
        let readback_path = Arc::clone(&path);
        self.replace_selected_owner_text(
            (start, end),
            b"sdtContent",
            paragraph.get(),
            authored_text.into(),
            move |before, after| Operation::ReplaceNestedContentControlText {
                paragraph,
                controls: path,
                before,
                after,
            },
            move |candidate| {
                selected_nested_inline_control_text(candidate, paragraph, &readback_path)
            },
        )
    }

    /// Replace one direct hyperlink inside the deepest nested inline control.
    /// Every selector step is a direct child, so the path cannot cross owner
    /// boundaries or relationship scopes.
    ///
    /// # Errors
    ///
    /// Returns a checked path/refusal, resource-limit, or malformed XML error
    /// without changing the projected snapshot.
    pub fn replace_nested_content_control_hyperlink_text(
        &mut self,
        paragraph: Position,
        controls: &[Position],
        hyperlink: Position,
        authored_text: impl Into<String>,
    ) -> TransactionResult<&mut Self> {
        self.ensure_unmanaged("replace_nested_content_control_hyperlink_text")?;
        let path: Arc<[Position]> = controls.into();
        let content = select_nested_inline_control_content(&self.projected, paragraph, &path)?;
        let range = select_hyperlink_owner(
            &self.projected,
            content,
            b"sdtContent",
            hyperlink,
            paragraph.get(),
        )?;
        let readback_path = Arc::clone(&path);
        self.replace_selected_owner_text(
            range,
            b"hyperlink",
            paragraph.get(),
            authored_text.into(),
            move |before, after| Operation::ReplaceNestedContentControlHyperlinkText {
                paragraph,
                controls: path,
                hyperlink,
                before,
                after,
            },
            move |candidate| {
                selected_nested_inline_control_hyperlink_text(
                    candidate,
                    paragraph,
                    &readback_path,
                    hyperlink,
                )
            },
        )
    }

    /// Replace one direct paragraph inside the deepest block content control
    /// reached by a non-empty path. The first position selects a direct-body
    /// `w:sdt`; later positions select a direct nested `w:sdt`.
    ///
    /// # Errors
    ///
    /// Returns a checked path/refusal, resource-limit, or malformed XML error
    /// without changing the projected snapshot.
    pub fn replace_block_content_control_paragraph_text(
        &mut self,
        controls: &[Position],
        paragraph: Position,
        authored_text: impl Into<String>,
    ) -> TransactionResult<&mut Self> {
        self.ensure_unmanaged("replace_block_content_control_paragraph_text")?;
        let path: Arc<[Position]> = controls.into();
        let (start, end) = select_block_control_paragraph(&self.projected, &path, paragraph)?;
        let readback_path = Arc::clone(&path);
        self.replace_selected_owner_text(
            (start, end),
            b"p",
            path.first().map_or(0, |position| position.get()),
            authored_text.into(),
            move |before, after| Operation::ReplaceBlockContentControlParagraphText {
                controls: path,
                paragraph,
                before,
                after,
            },
            move |candidate| {
                selected_block_control_paragraph_text(candidate, &readback_path, paragraph)
            },
        )
    }

    /// Replace one direct hyperlink inside a path-addressed block-control
    /// paragraph without changing its relationship identifier.
    ///
    /// # Errors
    ///
    /// Returns a checked path/refusal, resource-limit, or malformed XML error
    /// without changing the projected snapshot.
    pub fn replace_block_content_control_paragraph_hyperlink_text(
        &mut self,
        controls: &[Position],
        paragraph: Position,
        hyperlink: Position,
        authored_text: impl Into<String>,
    ) -> TransactionResult<&mut Self> {
        self.ensure_unmanaged("replace_block_content_control_paragraph_hyperlink_text")?;
        let path: Arc<[Position]> = controls.into();
        let owner = select_block_control_paragraph(&self.projected, &path, paragraph)?;
        let error_position = path.first().map_or(0, |position| position.get());
        let range =
            select_hyperlink_owner(&self.projected, owner, b"p", hyperlink, error_position)?;
        let readback_path = Arc::clone(&path);
        self.replace_selected_owner_text(
            range,
            b"hyperlink",
            error_position,
            authored_text.into(),
            move |before, after| Operation::ReplaceBlockContentControlParagraphHyperlinkText {
                controls: path,
                paragraph,
                hyperlink,
                before,
                after,
            },
            move |candidate| {
                selected_block_control_paragraph_hyperlink_text(
                    candidate,
                    &readback_path,
                    paragraph,
                    hyperlink,
                )
            },
        )
    }

    /// Atomically replace direct hyperlinks across one or more direct
    /// paragraphs in the deepest block content control reached by `controls`.
    ///
    /// Addresses must be non-empty, unique, and strictly increasing by
    /// paragraph then hyperlink position. Direct-child selection keeps every
    /// leaf within the same path-addressed control owner.
    ///
    /// # Errors
    ///
    /// Returns a typed ambiguity refusal for an empty, duplicate, or
    /// non-canonical selector list, or a resource-limit error for an oversized
    /// batch. Any checked path or leaf failure also leaves the entire edit
    /// unchanged.
    pub fn replace_block_content_control_paragraph_hyperlink_texts(
        &mut self,
        controls: &[Position],
        replacements: &[HyperlinkTextReplacement],
    ) -> TransactionResult<&mut Self> {
        self.ensure_unmanaged("replace_block_content_control_paragraph_hyperlink_texts")?;
        validate_hyperlink_replacements(replacements)?;
        let mut candidate = self.try_clone()?;
        for replacement in replacements {
            candidate.replace_block_content_control_paragraph_hyperlink_text(
                controls,
                replacement.address.paragraph,
                replacement.address.hyperlink,
                replacement.text.as_str(),
            )?;
        }
        *self = candidate;
        Ok(self)
    }

    /// Replace text in a basic direct-body table cell.
    ///
    /// The supported cell contains one direct paragraph. Its existing runs,
    /// formatting, cell properties, drawings, and unknown run XML remain in
    /// place; nested tables, controls, and multiple cell paragraphs are
    /// refused.
    ///
    /// # Errors
    ///
    /// Returns a checked selector/refusal, resource-limit, or malformed XML
    /// error without changing the projected snapshot.
    pub fn replace_table_cell_text(
        &mut self,
        table: Position,
        row: Position,
        cell: Position,
        authored_text: impl Into<String>,
    ) -> TransactionResult<&mut Self> {
        self.ensure_unmanaged("replace_table_cell_text")?;
        self.reserve_operation()?;
        let text = authored_text.into();
        validate_authored_text(&text).map_err(|reason| TransactionError::Refused {
            position: table.get(),
            reason,
        })?;
        let replacement_text_bytes = self.checked_text_total(text.len())?;
        let cell_selection = select_cell(&self.projected, table, row, cell)?;
        let paragraph_range = single_cell_paragraph(cell_selection.xml)?;
        let paragraph_start = checked_relative_start(cell_selection.start, paragraph_range)?;
        let paragraph_end = checked_relative_end(cell_selection.start, paragraph_range)?;
        let paragraph_xml = checked_slice(
            self.projected.xml_bytes(),
            paragraph_start,
            paragraph_end,
            "table cell paragraph",
        )?;
        let owner =
            scan_text_owner(paragraph_xml, b"p").map_err(|reason| TransactionError::Refused {
                position: table.get(),
                reason,
            })?;
        if owner.text == text {
            return Ok(self);
        }
        let replacement = rewrite_text_owner(paragraph_xml, &owner, &text)?;
        let xml = replace_range(
            self.projected.xml_bytes(),
            paragraph_start,
            paragraph_end,
            &replacement,
        )?;
        let candidate = self.projected.with_rewritten_xml(xml)?;
        let actual = selected_cell_text(&candidate, table, row, cell)?;
        if actual != text {
            return Err(crate::Error::InvalidFormat(
                "document table-cell edit failed semantic readback".into(),
            )
            .into());
        }
        self.operations.push(Operation::ReplaceCellText {
            table,
            row,
            cell,
            before: owner.text,
            after: text,
        });
        self.replacement_text_bytes = replacement_text_bytes;
        self.projected = candidate;
        Ok(self)
    }

    /// Replace one direct paragraph in a rich or multi-paragraph table cell.
    /// Other paragraphs, nested tables, cell properties, and opaque cell XML
    /// remain untouched.
    ///
    /// # Errors
    ///
    /// Returns a checked selector/refusal, resource-limit, or malformed XML
    /// error without changing the projected snapshot.
    pub fn replace_table_cell_paragraph_text(
        &mut self,
        table: Position,
        row: Position,
        cell: Position,
        paragraph: Position,
        authored_text: impl Into<String>,
    ) -> TransactionResult<&mut Self> {
        self.ensure_unmanaged("replace_table_cell_paragraph_text")?;
        self.reserve_operation()?;
        let text = authored_text.into();
        validate_authored_text(&text).map_err(|reason| TransactionError::Refused {
            position: table.get(),
            reason,
        })?;
        let replacement_text_bytes = self.checked_text_total(text.len())?;
        let cell_selection = select_cell(&self.projected, table, row, cell)?;
        let paragraph_range = select_direct_child(
            cell_selection.xml,
            b"tc",
            b"p",
            paragraph,
            Refusal::CellNotFound,
        )
        .map_err(|reason| TransactionError::Refused {
            position: table.get(),
            reason,
        })?;
        let paragraph_start = checked_relative_start(cell_selection.start, paragraph_range)?;
        let paragraph_end = checked_relative_end(cell_selection.start, paragraph_range)?;
        let paragraph_xml = checked_slice(
            self.projected.xml_bytes(),
            paragraph_start,
            paragraph_end,
            "table cell paragraph",
        )?;
        let owner =
            scan_text_owner(paragraph_xml, b"p").map_err(|reason| TransactionError::Refused {
                position: table.get(),
                reason,
            })?;
        if owner.text == text {
            return Ok(self);
        }
        let replacement = rewrite_text_owner(paragraph_xml, &owner, &text)?;
        let xml = replace_range(
            self.projected.xml_bytes(),
            paragraph_start,
            paragraph_end,
            &replacement,
        )?;
        let candidate = self.projected.with_rewritten_xml(xml)?;
        let actual = selected_cell_paragraph_text(&candidate, table, row, cell, paragraph)?;
        if actual != text {
            return Err(crate::Error::InvalidFormat(
                "document table-cell paragraph edit failed semantic readback".into(),
            )
            .into());
        }
        self.operations.push(Operation::ReplaceCellParagraphText {
            table,
            row,
            cell,
            paragraph,
            before: owner.text,
            after: text,
        });
        self.replacement_text_bytes = replacement_text_bytes;
        self.projected = candidate;
        Ok(self)
    }

    /// Replace one direct paragraph in a cell reached through a non-empty
    /// body-to-nested-table path. Each later table must be a direct child of
    /// the preceding selected cell.
    ///
    /// # Errors
    ///
    /// Returns a checked path/refusal, resource-limit, or malformed XML error
    /// without changing the projected snapshot.
    pub fn replace_nested_table_cell_paragraph_text(
        &mut self,
        path: &[TableCellAddress],
        paragraph: Position,
        authored_text: impl Into<String>,
    ) -> TransactionResult<&mut Self> {
        self.ensure_unmanaged("replace_nested_table_cell_paragraph_text")?;
        let path_arc: Arc<[TableCellAddress]> = path.into();
        let (start, end) = select_nested_cell_paragraph(&self.projected, &path_arc, paragraph)?;
        let readback_path = Arc::clone(&path_arc);
        self.replace_selected_owner_text(
            (start, end),
            b"p",
            path_arc.first().map_or(0, |address| address.table.get()),
            authored_text.into(),
            move |before, after| Operation::ReplaceNestedCellParagraphText {
                path: path_arc,
                paragraph,
                before,
                after,
            },
            move |candidate| {
                selected_nested_cell_paragraph_text(candidate, &readback_path, paragraph)
            },
        )
    }

    /// Replace one direct hyperlink inside a path-addressed nested-table cell
    /// paragraph without changing its relationship identifier.
    ///
    /// # Errors
    ///
    /// Returns a checked path/refusal, resource-limit, or malformed XML error
    /// without changing the projected snapshot.
    pub fn replace_nested_table_cell_paragraph_hyperlink_text(
        &mut self,
        path: &[TableCellAddress],
        paragraph: Position,
        hyperlink: Position,
        authored_text: impl Into<String>,
    ) -> TransactionResult<&mut Self> {
        self.ensure_unmanaged("replace_nested_table_cell_paragraph_hyperlink_text")?;
        let path_arc: Arc<[TableCellAddress]> = path.into();
        let owner = select_nested_cell_paragraph(&self.projected, &path_arc, paragraph)?;
        let error_position = path_arc.first().map_or(0, |address| address.table.get());
        let range =
            select_hyperlink_owner(&self.projected, owner, b"p", hyperlink, error_position)?;
        let readback_path = Arc::clone(&path_arc);
        self.replace_selected_owner_text(
            range,
            b"hyperlink",
            error_position,
            authored_text.into(),
            move |before, after| Operation::ReplaceNestedCellParagraphHyperlinkText {
                path: path_arc,
                paragraph,
                hyperlink,
                before,
                after,
            },
            move |candidate| {
                selected_nested_cell_paragraph_hyperlink_text(
                    candidate,
                    &readback_path,
                    paragraph,
                    hyperlink,
                )
            },
        )
    }

    /// Atomically replace direct hyperlinks across one or more direct
    /// paragraphs in the deepest cell reached by `path`.
    ///
    /// Addresses must be non-empty, unique, and strictly increasing by
    /// paragraph then hyperlink position. Direct-child selection keeps every
    /// leaf within the same path-addressed cell owner.
    ///
    /// # Errors
    ///
    /// Returns a typed ambiguity refusal for an empty, duplicate, or
    /// non-canonical selector list, or a resource-limit error for an oversized
    /// batch. Any checked path or leaf failure also leaves the entire edit
    /// unchanged.
    pub fn replace_nested_table_cell_paragraph_hyperlink_texts(
        &mut self,
        path: &[TableCellAddress],
        replacements: &[HyperlinkTextReplacement],
    ) -> TransactionResult<&mut Self> {
        self.ensure_unmanaged("replace_nested_table_cell_paragraph_hyperlink_texts")?;
        validate_hyperlink_replacements(replacements)?;
        let mut candidate = self.try_clone()?;
        for replacement in replacements {
            candidate.replace_nested_table_cell_paragraph_hyperlink_text(
                path,
                replacement.address.paragraph,
                replacement.address.hyperlink,
                replacement.text.as_str(),
            )?;
        }
        *self = candidate;
        Ok(self)
    }

    /// Insert a compact plain-text paragraph at a projected zero-based
    /// position. `position == paragraph_count()` appends before the body-final
    /// section properties.
    ///
    /// # Errors
    ///
    /// Returns a typed refusal, checked-position error, resource-limit error,
    /// or malformed-document error without changing the projected snapshot.
    pub fn insert_paragraph(
        &mut self,
        position: Position,
        authored_text: impl Into<String>,
    ) -> TransactionResult<&mut Self> {
        self.ensure_unmanaged("insert_paragraph")?;
        self.reserve_operation()?;
        let text = authored_text.into();
        validate_authored_text(&text).map_err(|reason| TransactionError::Refused {
            position: position.get(),
            reason,
        })?;
        let replacement_text_bytes = self.checked_text_total(text.len())?;
        let count = self.projected.paragraph_count();
        if position.get() > count {
            return Err(TransactionError::OutOfBounds {
                position: position.get(),
                len: count,
            });
        }
        let offset = if position.get() == count {
            usize::try_from(self.projected.content_end).map_err(|_conversion_error| {
                crate::Error::InvalidFormat("document insertion offset does not fit usize".into())
            })?
        } else {
            usize::try_from(self.range(position)?.start).map_err(|_conversion_error| {
                crate::Error::InvalidFormat("paragraph offset does not fit usize".into())
            })?
        };
        let paragraph = try_plain_paragraph(self.projected.conformance, &text)?;
        let xml = replace_range(self.projected.xml_bytes(), offset, offset, &paragraph)?;
        let candidate = self.projected.with_rewritten_xml(xml)?;
        let readback = candidate
            .paragraph(position)
            .ok_or(TransactionError::OutOfBounds {
                position: position.get(),
                len: candidate.paragraph_count(),
            })?
            .text()?;
        let expected_count = count.checked_add(1).ok_or(TransactionError::Limit {
            resource: "paragraphs",
            max: usize::MAX,
            actual: usize::MAX,
        })?;
        if readback != text || candidate.paragraph_count() != expected_count {
            return Err(crate::Error::InvalidFormat(
                "document paragraph insertion failed semantic readback".into(),
            )
            .into());
        }
        self.operations
            .push(Operation::InsertParagraph { position, text });
        self.replacement_text_bytes = replacement_text_bytes;
        self.projected = candidate;
        Ok(self)
    }

    /// Insert one non-mutating dependency-checked paragraph transfer plan.
    ///
    /// The plan is receiver-specific: every relationship reference was mapped
    /// to an equivalent relationship already owned by the exact target
    /// package. This method never copies or guesses package dependencies.
    ///
    /// # Errors
    ///
    /// Returns a stale-plan, position, operation-bound, or XML validation
    /// error without changing the projected snapshot.
    pub fn insert_paragraph_transfer(
        &mut self,
        position: Position,
        plan: &ParagraphTransfer,
    ) -> TransactionResult<&mut Self> {
        self.ensure_unmanaged("insert_paragraph_transfer")?;
        if plan.target.as_slice() != self.base.xml_bytes() {
            return Err(TransactionError::StaleSource);
        }
        self.insert_transferred_paragraph(
            position,
            Arc::clone(&plan.fragment),
            Arc::clone(&plan.dependency_digest),
            Arc::clone(&plan.inverse_dependency_digest),
            Arc::clone(&plan.graph),
        )
    }

    fn apply_operation(&mut self, operation: &Operation) -> TransactionResult<&mut Self> {
        match operation {
            Operation::ReplaceParagraphText {
                position,
                before,
                after,
            } => {
                let actual = self
                    .projected
                    .paragraph(*position)
                    .ok_or(TransactionError::OutOfBounds {
                        position: position.get(),
                        len: self.projected.paragraph_count(),
                    })?
                    .text()?;
                if &actual != before {
                    return Err(TransactionError::SemanticPrecondition);
                }
                self.replace_paragraph_text(*position, after.clone())
            },
            Operation::ReplaceHyperlinkText {
                paragraph,
                hyperlink,
                before,
                after,
            } => {
                if selected_hyperlink_text(&self.projected, *paragraph, *hyperlink)? != *before {
                    return Err(TransactionError::SemanticPrecondition);
                }
                self.replace_hyperlink_text(*paragraph, *hyperlink, after.clone())
            },
            Operation::ReplaceRunText {
                paragraph,
                run,
                before,
                after,
            } => {
                if selected_direct_paragraph_owner_text(
                    &self.projected,
                    *paragraph,
                    *run,
                    b"r",
                    Refusal::RunNotFound,
                )? != *before
                {
                    return Err(TransactionError::SemanticPrecondition);
                }
                self.replace_run_text(*paragraph, *run, after.clone())
            },
            Operation::ReplaceSimpleFieldText {
                paragraph,
                field,
                before,
                after,
            } => {
                if selected_direct_paragraph_owner_text(
                    &self.projected,
                    *paragraph,
                    *field,
                    b"fldSimple",
                    Refusal::FieldNotFound,
                )? != *before
                {
                    return Err(TransactionError::SemanticPrecondition);
                }
                self.replace_simple_field_text(*paragraph, *field, after.clone())
            },
            Operation::ReplaceComplexFieldText {
                paragraph,
                field,
                before,
                after,
            } => {
                if selected_complex_field_result_text(&self.projected, *paragraph, *field)?
                    != *before
                {
                    return Err(TransactionError::SemanticPrecondition);
                }
                self.replace_complex_field_result_text(*paragraph, *field, after.clone())
            },
            Operation::ReplaceRevisionText {
                paragraph,
                kind,
                revision,
                before,
                after,
            } => {
                if selected_direct_paragraph_owner_text(
                    &self.projected,
                    *paragraph,
                    *revision,
                    kind.local_name(),
                    Refusal::RevisionNotFound,
                )? != *before
                {
                    return Err(TransactionError::SemanticPrecondition);
                }
                self.replace_revision_text(*paragraph, *kind, *revision, after.clone())
            },
            Operation::ApplyRevision {
                selector,
                action,
                before,
                after,
            } => self.apply_raw_revision_operation(*selector, *action, before, after),
            Operation::ReplaceContentControlText {
                paragraph,
                control,
                before,
                after,
            } => {
                if selected_content_control_text(&self.projected, *paragraph, *control)? != *before
                {
                    return Err(TransactionError::SemanticPrecondition);
                }
                self.replace_content_control_text(*paragraph, *control, after.clone())
            },
            Operation::ReplaceNestedContentControlText {
                paragraph,
                controls,
                before,
                after,
            } => {
                if selected_nested_inline_control_text(&self.projected, *paragraph, controls)?
                    != *before
                {
                    return Err(TransactionError::SemanticPrecondition);
                }
                self.replace_nested_content_control_text(*paragraph, controls, after.clone())
            },
            Operation::ReplaceNestedContentControlHyperlinkText {
                paragraph,
                controls,
                hyperlink,
                before,
                after,
            } => {
                if selected_nested_inline_control_hyperlink_text(
                    &self.projected,
                    *paragraph,
                    controls,
                    *hyperlink,
                )? != *before
                {
                    return Err(TransactionError::SemanticPrecondition);
                }
                self.replace_nested_content_control_hyperlink_text(
                    *paragraph,
                    controls,
                    *hyperlink,
                    after.clone(),
                )
            },
            Operation::ReplaceBlockContentControlParagraphText {
                controls,
                paragraph,
                before,
                after,
            } => {
                if selected_block_control_paragraph_text(&self.projected, controls, *paragraph)?
                    != *before
                {
                    return Err(TransactionError::SemanticPrecondition);
                }
                self.replace_block_content_control_paragraph_text(
                    controls,
                    *paragraph,
                    after.clone(),
                )
            },
            Operation::ReplaceBlockContentControlParagraphHyperlinkText {
                controls,
                paragraph,
                hyperlink,
                before,
                after,
            } => {
                if selected_block_control_paragraph_hyperlink_text(
                    &self.projected,
                    controls,
                    *paragraph,
                    *hyperlink,
                )? != *before
                {
                    return Err(TransactionError::SemanticPrecondition);
                }
                self.replace_block_content_control_paragraph_hyperlink_text(
                    controls,
                    *paragraph,
                    *hyperlink,
                    after.clone(),
                )
            },
            Operation::ReplaceCellText {
                table,
                row,
                cell,
                before,
                after,
            } => {
                if selected_cell_text(&self.projected, *table, *row, *cell)? != *before {
                    return Err(TransactionError::SemanticPrecondition);
                }
                self.replace_table_cell_text(*table, *row, *cell, after.clone())
            },
            Operation::ReplaceCellParagraphText {
                table,
                row,
                cell,
                paragraph,
                before,
                after,
            } => {
                if selected_cell_paragraph_text(&self.projected, *table, *row, *cell, *paragraph)?
                    != *before
                {
                    return Err(TransactionError::SemanticPrecondition);
                }
                self.replace_table_cell_paragraph_text(
                    *table,
                    *row,
                    *cell,
                    *paragraph,
                    after.clone(),
                )
            },
            Operation::ReplaceNestedCellParagraphText {
                path,
                paragraph,
                before,
                after,
            } => {
                if selected_nested_cell_paragraph_text(&self.projected, path, *paragraph)?
                    != *before
                {
                    return Err(TransactionError::SemanticPrecondition);
                }
                self.replace_nested_table_cell_paragraph_text(path, *paragraph, after.clone())
            },
            Operation::ReplaceNestedCellParagraphHyperlinkText {
                path,
                paragraph,
                hyperlink,
                before,
                after,
            } => {
                if selected_nested_cell_paragraph_hyperlink_text(
                    &self.projected,
                    path,
                    *paragraph,
                    *hyperlink,
                )? != *before
                {
                    return Err(TransactionError::SemanticPrecondition);
                }
                self.replace_nested_table_cell_paragraph_hyperlink_text(
                    path,
                    *paragraph,
                    *hyperlink,
                    after.clone(),
                )
            },
            Operation::InsertParagraph { position, text } => {
                self.insert_paragraph(*position, text.clone())
            },
            Operation::RemoveParagraph { position, text } => {
                self.remove_plain_paragraph(*position, text)
            },
            Operation::InsertTransferredParagraph {
                position,
                xml,
                dependency_digest,
                inverse_dependency_digest,
                graph,
            } => self.insert_transferred_paragraph(
                *position,
                Arc::clone(xml),
                Arc::clone(dependency_digest),
                Arc::clone(inverse_dependency_digest),
                Arc::clone(graph),
            ),
            Operation::RemoveTransferredParagraph {
                position,
                xml,
                dependency_digest,
                inverse_dependency_digest,
                graph,
            } => self.remove_transferred_paragraph(
                *position,
                xml,
                dependency_digest,
                inverse_dependency_digest,
                graph,
            ),
        }
    }

    fn apply_raw_revision_operation(
        &mut self,
        selector: RevisionSelector,
        action: RevisionAction,
        before: &Arc<Vec<u8>>,
        after: &Arc<Vec<u8>>,
    ) -> TransactionResult<&mut Self> {
        // Replay rebuilds the projection the same way as `apply_revision_as`,
        // so it carries the same managed refusal.
        self.ensure_unmanaged("apply_revision")?;
        self.reserve_operation()?;
        let range = self.range(selector.paragraph)?;
        let start = checked_start(range, "paragraph")?;
        let end = checked_end(range, "paragraph")?;
        let source = checked_slice(self.projected.xml_bytes(), start, end, "paragraph")?;
        if source != before.as_slice() {
            return Err(TransactionError::SemanticPrecondition);
        }
        let inherited_namespaces = self
            .projected
            .paragraph_namespaces
            .get(selector.paragraph.get())
            .ok_or(TransactionError::OutOfBounds {
                position: selector.paragraph.get(),
                len: self.projected.paragraph_count(),
            })?;
        let forward_match =
            rewrite_revision_paragraph(source, selector, action, inherited_namespaces)
                .is_ok_and(|candidate| candidate.as_slice() == after.as_slice());
        let reverse_match = !forward_match
            && rewrite_revision_paragraph(after, selector, action, inherited_namespaces)
                .is_ok_and(|candidate| candidate.as_slice() == source);
        if !forward_match && !reverse_match {
            return Err(TransactionError::SemanticPrecondition);
        }
        let candidate = Snapshot::from_xml(replace_range(
            self.projected.xml_bytes(),
            start,
            end,
            after,
        )?)?;
        if candidate.paragraph_count() != self.projected.paragraph_count() {
            return Err(crate::Error::InvalidFormat(
                "revision replay changed the paragraph count".into(),
            )
            .into());
        }
        self.operations.push(Operation::ApplyRevision {
            selector,
            action,
            before: Arc::clone(before),
            after: Arc::clone(after),
        });
        self.projected = candidate;
        Ok(self)
    }

    fn remove_plain_paragraph(
        &mut self,
        position: Position,
        expected_text: &str,
    ) -> TransactionResult<&mut Self> {
        self.reserve_operation()?;
        let range = self.range(position)?;
        let start = checked_start(range, "paragraph")?;
        let end = checked_end(range, "paragraph")?;
        let source = checked_slice(self.projected.xml_bytes(), start, end, "paragraph")?;
        let expected = try_plain_paragraph(self.projected.conformance, expected_text)?;
        if source != expected.as_slice() {
            return Err(TransactionError::SemanticPrecondition);
        }
        let xml = replace_range(self.projected.xml_bytes(), start, end, &[])?;
        let candidate = self.projected.with_rewritten_xml(xml)?;
        if candidate.paragraph_count().checked_add(1) != Some(self.projected.paragraph_count()) {
            return Err(crate::Error::InvalidFormat(
                "document paragraph removal failed semantic readback".into(),
            )
            .into());
        }
        self.operations.push(Operation::RemoveParagraph {
            position,
            text: expected_text.to_owned(),
        });
        self.projected = candidate;
        Ok(self)
    }

    fn insert_transferred_paragraph(
        &mut self,
        position: Position,
        xml: Arc<Vec<u8>>,
        dependency_digest: Arc<str>,
        inverse_dependency_digest: Arc<str>,
        graph: Arc<TransferGraph>,
    ) -> TransactionResult<&mut Self> {
        self.reserve_operation()?;
        let count = self.projected.paragraph_count();
        if position.get() > count {
            return Err(TransactionError::OutOfBounds {
                position: position.get(),
                len: count,
            });
        }
        let offset = if position.get() == count {
            usize::try_from(self.projected.content_end).map_err(|_error| {
                crate::Error::InvalidFormat("document insertion offset does not fit usize".into())
            })?
        } else {
            checked_start(self.range(position)?, "paragraph")?
        };
        let candidate = self.projected.with_rewritten_xml(replace_range(
            self.projected.xml_bytes(),
            offset,
            offset,
            xml.as_slice(),
        )?)?;
        let inserted = candidate
            .paragraph(position)
            .ok_or(TransactionError::OutOfBounds {
                position: position.get(),
                len: candidate.paragraph_count(),
            })?;
        if inserted.xml_bytes() != xml.as_slice() {
            return Err(crate::Error::InvalidFormat(
                "transferred paragraph failed exact readback".into(),
            )
            .into());
        }
        self.operations.push(Operation::InsertTransferredParagraph {
            position,
            xml,
            dependency_digest,
            inverse_dependency_digest,
            graph,
        });
        self.projected = candidate;
        Ok(self)
    }

    fn remove_transferred_paragraph(
        &mut self,
        position: Position,
        xml: &Arc<Vec<u8>>,
        dependency_digest: &Arc<str>,
        inverse_dependency_digest: &Arc<str>,
        graph: &Arc<TransferGraph>,
    ) -> TransactionResult<&mut Self> {
        self.reserve_operation()?;
        let range = self.range(position)?;
        let start = checked_start(range, "paragraph")?;
        let end = checked_end(range, "paragraph")?;
        let source = checked_slice(self.projected.xml_bytes(), start, end, "paragraph")?;
        if source != xml.as_slice() {
            return Err(TransactionError::SemanticPrecondition);
        }
        let candidate = self.projected.with_rewritten_xml(replace_range(
            self.projected.xml_bytes(),
            start,
            end,
            &[],
        )?)?;
        self.operations.push(Operation::RemoveTransferredParagraph {
            position,
            xml: Arc::clone(xml),
            dependency_digest: Arc::clone(dependency_digest),
            inverse_dependency_digest: Arc::clone(inverse_dependency_digest),
            graph: Arc::clone(graph),
        });
        self.projected = candidate;
        Ok(self)
    }

    fn replace_direct_paragraph_owner_text(
        &mut self,
        paragraph: Position,
        owner: Position,
        child_name: &[u8],
        missing: Refusal,
        text: String,
        operation: impl FnOnce(String, String) -> Operation,
    ) -> TransactionResult<&mut Self> {
        self.reserve_operation()?;
        let validation = if child_name == b"r" {
            validate_authored_run_content(&text)
        } else {
            validate_authored_text(&text)
        };
        validation.map_err(|reason| TransactionError::Refused {
            position: paragraph.get(),
            reason,
        })?;
        let replacement_text_bytes = self.checked_text_total(text.len())?;
        let paragraph_range = self.range(paragraph)?;
        let paragraph_start = checked_start(paragraph_range, "paragraph")?;
        let paragraph_end = checked_end(paragraph_range, "paragraph")?;
        let paragraph_xml = checked_slice(
            self.projected.xml_bytes(),
            paragraph_start,
            paragraph_end,
            "paragraph",
        )?;
        let owner_range = select_direct_child(paragraph_xml, b"p", child_name, owner, missing)
            .map_err(|reason| TransactionError::Refused {
                position: paragraph.get(),
                reason,
            })?;
        let owner_start = checked_relative_start(paragraph_start, owner_range)?;
        let owner_end = checked_relative_end(paragraph_start, owner_range)?;
        let owner_xml = checked_slice(
            self.projected.xml_bytes(),
            owner_start,
            owner_end,
            "paragraph text owner",
        )?;
        let scanned =
            scan_text_owner(owner_xml, child_name).map_err(|reason| TransactionError::Refused {
                position: paragraph.get(),
                reason,
            })?;
        if scanned.text == text {
            return Ok(self);
        }
        let replacement = rewrite_text_owner(owner_xml, &scanned, &text)?;
        let xml = replace_range(
            self.projected.xml_bytes(),
            owner_start,
            owner_end,
            &replacement,
        )?;
        let candidate = self.projected.with_rewritten_xml(xml)?;
        let actual = selected_direct_paragraph_owner_text(
            &candidate, paragraph, owner, child_name, missing,
        )?;
        if actual != text {
            return Err(crate::Error::InvalidFormat(
                "document owner text edit failed semantic readback".into(),
            )
            .into());
        }
        self.operations.push(operation(scanned.text, text));
        self.replacement_text_bytes = replacement_text_bytes;
        self.projected = candidate;
        Ok(self)
    }

    fn replace_selected_owner_text(
        &mut self,
        range: (usize, usize),
        root_name: &[u8],
        error_position: usize,
        text: String,
        operation: impl FnOnce(String, String) -> Operation,
        readback: impl FnOnce(&Snapshot) -> TransactionResult<String>,
    ) -> TransactionResult<&mut Self> {
        self.reserve_operation()?;
        let (start, end) = range;
        validate_authored_text(&text).map_err(|reason| TransactionError::Refused {
            position: error_position,
            reason,
        })?;
        let replacement_text_bytes = self.checked_text_total(text.len())?;
        let owner_xml = checked_slice(self.projected.xml_bytes(), start, end, "selected owner")?;
        let owner =
            scan_text_owner(owner_xml, root_name).map_err(|reason| TransactionError::Refused {
                position: error_position,
                reason,
            })?;
        if owner.text == text {
            return Ok(self);
        }
        let replacement = rewrite_text_owner(owner_xml, &owner, &text)?;
        let candidate = self.projected.with_rewritten_xml(replace_range(
            self.projected.xml_bytes(),
            start,
            end,
            &replacement,
        )?)?;
        if readback(&candidate)? != text {
            return Err(crate::Error::InvalidFormat(
                "selected owner edit failed semantic readback".into(),
            )
            .into());
        }
        self.operations.push(operation(owner.text, text));
        self.replacement_text_bytes = replacement_text_bytes;
        self.projected = candidate;
        Ok(self)
    }

    /// Validate and publish the projected snapshot without changing the
    /// source snapshot.
    ///
    /// # Errors
    ///
    /// Reserved for commit-time document validation failures.
    pub fn commit(self) -> TransactionResult<Commit> {
        let managed_context = self.projected.managed_context();
        if let Some(context) = managed_context.as_ref() {
            context.check().map_err(managed_execution)?;
        }
        // Revision dispositions preserve retained opaque payload bytes, including
        // whitespace and attribute spelling. Global compaction would rewrite them.
        let preserve_revision_source = self
            .operations
            .iter()
            .any(|operation| matches!(operation, Operation::ApplyRevision { .. }));
        let projected = if self.base.same_source(&self.projected) {
            self.projected
        } else if self.base.is_source_backed() && self.projected.is_source_backed() {
            // Source-authorized edits are already validated and retain the
            // checked splice/output reservation. Running them through the
            // authored compacting writer would detach that owner and lose
            // the exact source proof required by publication.
            self.projected
        } else if preserve_revision_source {
            // The preserved bytes skip compaction but not its UTF-8 check: a
            // changed document that is not UTF-8 is refused either way.
            std::str::from_utf8(self.projected.xml_bytes()).map_err(|error| {
                crate::Error::InvalidFormat(format!(
                    "changed main-document XML is not UTF-8: {error}"
                ))
            })?;
            self.projected
        } else {
            match self.compaction {
                CompactionPolicy::PreserveUnmodified => {
                    commit_preserving_unmodified(&self.base, &self.projected)?
                },
                CompactionPolicy::WholeDocument => compact_whole_document(&self.projected)?,
            }
        };
        if let Some(context) = managed_context.as_ref() {
            context.check().map_err(managed_execution)?;
        }
        let (inverse_operations, inverse_admission, inverse_string_admission) =
            build_inverse_operations(&self.operations, managed_context.as_ref())?;
        let diagnostics = Diagnostics {
            operations: self.operations.len(),
            changed: !self.base.same_source(&projected),
        };
        let patch = Patch {
            before: self.base,
            after: projected.clone(),
            operations: OperationList::from_vec(self.operations),
            operation_admission: self.operation_admission,
            operation_string_admission: self.operation_string_admission,
            inverse_operations,
            inverse_admission,
            inverse_string_admission,
        };
        Ok(Commit {
            snapshot: projected,
            patch,
            diagnostics,
        })
    }

    fn range(&self, position: Position) -> TransactionResult<Range> {
        self.projected.range(position)
    }

    fn reserve_operation(&mut self) -> TransactionResult<()> {
        if self.operations.len() >= MAX_OPERATIONS {
            return Err(TransactionError::Limit {
                resource: "operations",
                max: MAX_OPERATIONS,
                actual: self.operations.len().saturating_add(1),
            });
        }
        if self.operation_admission.is_none() {
            if self.projected.admission.is_some() {
                let operation_admission = self.reserve_managed_operation_admission(1)?;
                self.operations.try_reserve_exact(1).map_err(|source| {
                    crate::Error::Allocation {
                        resource: "managed document operation metadata",
                        source,
                    }
                })?;
                self.operation_admission = Some(operation_admission);
            } else {
                self.operations.try_reserve_exact(1).map_err(|source| {
                    crate::Error::Allocation {
                        resource: "document operation metadata",
                        source,
                    }
                })?;
            }
        } else if self.projected.is_managed() && self.operations.is_empty() {
            // A managed operation ledger must always cover every retained
            // operation. Multi-paragraph source reconstruction installs a
            // freshly sized ledger atomically before replacing this vector.
            return Err(TransactionError::Document(crate::Error::UnsafeEdit {
                format: "DOCX",
                operation: "document.edit",
                reason: "managed operation metadata admission is inconsistent",
            }));
        }
        Ok(())
    }

    fn reserve_managed_operation_admission(
        &self,
        operation_count: usize,
    ) -> TransactionResult<Arc<OperationAdmission>> {
        let admission = self.projected.admission.as_ref().ok_or({
            TransactionError::Document(crate::Error::UnsafeEdit {
                format: "DOCX",
                operation: "document.edit",
                reason: "managed operation admission is unavailable",
            })
        })?;
        let operation_bytes = managed_operation_memory_bytes(operation_count)?;
        let memory = reserve_managed(&admission.context, Resource::Memory, operation_bytes)?;
        let objects = reserve_managed(&admission.context, Resource::Objects, operation_count)?;
        Ok(Arc::new(OperationAdmission {
            _memory: memory,
            _objects: objects,
        }))
    }

    fn reserve_input_text(&self, bytes: usize) -> TransactionResult<Option<Arc<StringAdmission>>> {
        if !self.projected.is_managed() {
            return Ok(None);
        }
        Ok(Some(self.reserve_string_admission(bytes)?))
    }

    fn reserve_string_admission(&self, bytes: usize) -> TransactionResult<Arc<StringAdmission>> {
        let admission = self.projected.admission.as_ref().ok_or({
            TransactionError::Document(crate::Error::UnsafeEdit {
                format: "DOCX",
                operation: "document.edit",
                reason: "managed string admission is unavailable",
            })
        })?;
        let memory = bytes
            .checked_add(size_of::<StringAdmission>())
            .and_then(|bytes| bytes.checked_add(size_of::<Arc<StringAdmission>>()))
            .ok_or(TransactionError::Limit {
                resource: "operation string metadata bytes",
                max: usize::MAX,
                actual: usize::MAX,
            })?;
        Ok(Arc::new(StringAdmission {
            _memory: reserve_managed(&admission.context, Resource::Memory, memory)?,
        }))
    }

    fn ensure_unmanaged(&self, operation: &'static str) -> TransactionResult<()> {
        if self.projected.is_managed() {
            return Err(TransactionError::Document(crate::Error::UnsafeEdit {
                format: "DOCX",
                operation,
                reason: "this managed transaction operation has no retained source-preserving implementation",
            }));
        }
        Ok(())
    }

    fn checked_text_total(&self, bytes: usize) -> TransactionResult<usize> {
        let actual =
            self.replacement_text_bytes
                .checked_add(bytes)
                .ok_or(TransactionError::Limit {
                    resource: "replacement text bytes",
                    max: MAX_REPLACEMENT_TEXT_BYTES,
                    actual: usize::MAX,
                })?;
        if actual > MAX_REPLACEMENT_TEXT_BYTES {
            return Err(TransactionError::Limit {
                resource: "replacement text bytes",
                max: MAX_REPLACEMENT_TEXT_BYTES,
                actual,
            });
        }
        Ok(actual)
    }
}

/// Diagnostics for one successful main-document commit.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Diagnostics {
    operations: usize,
    changed: bool,
}

impl Diagnostics {
    /// Number of semantic operations in the commit.
    #[must_use]
    pub const fn operations(self) -> usize {
        self.operations
    }

    /// Whether the commit changed the exact main-document bytes.
    #[must_use]
    pub const fn changed(self) -> bool {
        self.changed
    }
}

/// A successful main-document publication.
#[derive(Debug, Clone)]
pub struct Commit {
    snapshot: Snapshot,
    patch: Patch,
    diagnostics: Diagnostics,
}

impl Commit {
    /// Borrow the published snapshot.
    #[must_use]
    pub const fn snapshot(&self) -> &Snapshot {
        &self.snapshot
    }

    /// Borrow the reversible patch.
    #[must_use]
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Return content-free commit diagnostics.
    #[must_use]
    pub const fn diagnostics(&self) -> Diagnostics {
        self.diagnostics
    }

    /// Move the snapshot and patch out of the commit.
    #[must_use]
    pub fn into_parts(self) -> (Snapshot, Patch) {
        (self.snapshot, self.patch)
    }
}

/// A reversible, exact-source-checked main-document patch.
#[derive(Debug, Clone)]
pub struct Patch {
    before: Snapshot,
    after: Snapshot,
    operations: OperationList,
    operation_admission: Option<Arc<OperationAdmission>>,
    operation_string_admission: Option<Arc<StringAdmission>>,
    inverse_operations: OperationList,
    inverse_admission: Option<Arc<OperationAdmission>>,
    inverse_string_admission: Option<Arc<StringAdmission>>,
}

impl Patch {
    /// Exact immutable source required by this patch.
    #[must_use]
    pub const fn source(&self) -> &Snapshot {
        &self.before
    }

    /// Exact immutable target produced by this patch.
    #[must_use]
    pub const fn target(&self) -> &Snapshot {
        &self.after
    }

    /// Borrow the semantic operations in staging order.
    #[must_use]
    pub fn operations(&self) -> &[Operation] {
        self.operations.as_slice()
    }

    /// Whether this patch changes the exact main-document bytes.
    #[must_use]
    pub fn changed(&self) -> bool {
        !self.before.same_source(&self.after)
    }

    /// Return the exact inverse patch.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            before: self.after.clone(),
            after: self.before.clone(),
            operations: self.inverse_operations.clone(),
            operation_admission: self.inverse_admission.clone(),
            operation_string_admission: self.inverse_string_admission.clone(),
            inverse_operations: self.operations.clone(),
            inverse_admission: self.operation_admission.clone(),
            inverse_string_admission: self.operation_string_admission.clone(),
        }
    }

    /// Apply only when the target has the exact source document bytes.
    ///
    /// # Errors
    ///
    /// Returns [`TransactionError::StaleSource`] when `source` does not match
    /// the exact bytes against which this patch was produced.
    pub fn apply(&self, source: &Snapshot) -> TransactionResult<Snapshot> {
        if !source.same_source(&self.before) {
            return Err(TransactionError::StaleSource);
        }
        Ok(if self.changed() {
            self.after.clone()
        } else {
            source.clone()
        })
    }

    /// Reuse the retained target when `bytes` are already the exact,
    /// identity-free source this patch was produced against.
    ///
    /// [`Self::apply`] reads nothing from its `source` argument except that
    /// exact-source comparison and, for an unchanged patch, the source value
    /// it hands back. An opened package can therefore answer it from the main
    /// part blob it already retains, instead of copying and rescanning the
    /// whole part into a snapshot that is discarded immediately afterwards.
    /// The returned snapshot is the value `apply` would have returned: for an
    /// unchanged patch, `before` is byte-equal and identity-equal to the
    /// snapshot `apply` would have cloned, so the two are the same value.
    ///
    /// `None` means the cheap proof did not hold. It is not a refusal: the
    /// caller falls back to the full snapshot route, which keeps every
    /// refusal, error identity and ordering unchanged.
    pub(crate) fn target_for_exact_unmanaged_source(&self, bytes: &[u8]) -> Option<Snapshot> {
        if !self.before.retains_exact_unmanaged_xml(bytes) {
            return None;
        }
        Some(if self.changed() {
            self.after.clone()
        } else {
            self.before.clone()
        })
    }
}

fn build_inverse_operations(
    operations: &[Operation],
    context: Option<&ExecutionContext>,
) -> TransactionResult<(
    OperationList,
    Option<Arc<OperationAdmission>>,
    Option<Arc<StringAdmission>>,
)> {
    if operations.is_empty() {
        return Ok((OperationList::Empty, None, None));
    }
    let admission = if let Some(context) = context {
        let string_bytes = operations.iter().try_fold(0usize, |total, operation| {
            total
                .checked_add(operation_string_bytes(operation))
                .ok_or(TransactionError::Limit {
                    resource: "inverse operation string bytes",
                    max: usize::MAX,
                    actual: usize::MAX,
                })
        })?;
        let memory = managed_operation_memory_bytes(operations.len())?;
        let memory = reserve_managed(context, Resource::Memory, memory)?;
        let objects = reserve_managed(context, Resource::Objects, operations.len())?;
        let operation_admission = Arc::new(OperationAdmission {
            _memory: memory,
            _objects: objects,
        });
        let string_admission = if string_bytes == 0 {
            None
        } else {
            let memory = string_bytes
                .checked_add(size_of::<StringAdmission>())
                .and_then(|bytes| bytes.checked_add(size_of::<Arc<StringAdmission>>()))
                .ok_or(TransactionError::Limit {
                    resource: "inverse operation string metadata bytes",
                    max: usize::MAX,
                    actual: usize::MAX,
                })?;
            Some(Arc::new(StringAdmission {
                _memory: reserve_managed(context, Resource::Memory, memory)?,
            }))
        };
        (Some(operation_admission), string_admission)
    } else {
        (None, None)
    };
    let mut inverse = Vec::new();
    inverse
        .try_reserve_exact(operations.len())
        .map_err(|source| crate::Error::Allocation {
            resource: "inverse document operation metadata",
            source,
        })?;
    for operation in operations.iter().rev() {
        inverse.push(operation.inverse());
    }
    Ok((OperationList::from_vec(inverse), admission.0, admission.1))
}

fn operation_string_bytes(operation: &Operation) -> usize {
    match operation {
        Operation::ReplaceParagraphText { before, after, .. }
        | Operation::ReplaceHyperlinkText { before, after, .. }
        | Operation::ReplaceRunText { before, after, .. }
        | Operation::ReplaceSimpleFieldText { before, after, .. }
        | Operation::ReplaceComplexFieldText { before, after, .. }
        | Operation::ReplaceRevisionText { before, after, .. }
        | Operation::ReplaceContentControlText { before, after, .. }
        | Operation::ReplaceNestedContentControlText { before, after, .. }
        | Operation::ReplaceNestedContentControlHyperlinkText { before, after, .. }
        | Operation::ReplaceBlockContentControlParagraphText { before, after, .. }
        | Operation::ReplaceBlockContentControlParagraphHyperlinkText { before, after, .. }
        | Operation::ReplaceCellText { before, after, .. }
        | Operation::ReplaceCellParagraphText { before, after, .. }
        | Operation::ReplaceNestedCellParagraphText { before, after, .. }
        | Operation::ReplaceNestedCellParagraphHyperlinkText { before, after, .. } => {
            before.capacity().saturating_add(after.capacity())
        },
        Operation::InsertParagraph { text, .. } | Operation::RemoveParagraph { text, .. } => {
            text.capacity()
        },
        // The retained paragraph images are charged by their buffer capacity,
        // like the string fields above.
        Operation::ApplyRevision { before, after, .. } => {
            before.capacity().saturating_add(after.capacity())
        },
        Operation::InsertTransferredParagraph { .. }
        | Operation::RemoveTransferredParagraph { .. } => 0,
    }
}

fn managed_replacement_text_bytes(operations: &[Operation]) -> TransactionResult<usize> {
    operations.iter().try_fold(0usize, |total, operation| {
        let Operation::ReplaceParagraphText { after, .. } = operation else {
            return Err(TransactionError::Document(crate::Error::UnsafeEdit {
                format: "DOCX",
                operation: "document.edit",
                reason: "managed operation ledger contains an unsupported operation",
            }));
        };
        let total = total
            .checked_add(after.len())
            .ok_or(TransactionError::Limit {
                resource: "replacement text bytes",
                max: MAX_REPLACEMENT_TEXT_BYTES,
                actual: usize::MAX,
            })?;
        if total > MAX_REPLACEMENT_TEXT_BYTES {
            return Err(TransactionError::Limit {
                resource: "replacement text bytes",
                max: MAX_REPLACEMENT_TEXT_BYTES,
                actual: total,
            });
        }
        Ok(total)
    })
}

#[derive(Debug, Clone)]
enum OperationList {
    Empty,
    Shared(Arc<[Operation]>),
}

impl OperationList {
    fn from_vec(operations: Vec<Operation>) -> Self {
        if operations.is_empty() {
            Self::Empty
        } else {
            Self::Shared(operations.into())
        }
    }

    fn as_slice(&self) -> &[Operation] {
        match self {
            Self::Empty => &[],
            Self::Shared(operations) => operations.as_ref(),
        }
    }

    fn is_empty(&self) -> bool {
        self.as_slice().is_empty()
    }

    fn iter(&self) -> std::slice::Iter<'_, Operation> {
        self.as_slice().iter()
    }
}

fn managed_operation_memory_bytes(operation_count: usize) -> TransactionResult<usize> {
    // A managed edit retains its operation Vec until commit converts it into
    // the patch's Arc slice. The inverse builder has the same Vec-to-Arc
    // overlap, so admit both live representations before either allocation.
    size_of::<Operation>()
        .checked_mul(operation_count)
        .and_then(|bytes| bytes.checked_mul(2))
        .and_then(|bytes| bytes.checked_add(size_of::<Vec<Operation>>()))
        .and_then(|bytes| bytes.checked_add(size_of::<Arc<[Operation]>>()))
        .and_then(|bytes| bytes.checked_add(size_of::<OperationAdmission>()))
        .and_then(|bytes| bytes.checked_add(size_of::<Arc<OperationAdmission>>()))
        .ok_or(TransactionError::Limit {
            resource: "managed operation metadata bytes",
            max: usize::MAX,
            actual: usize::MAX,
        })
}

#[derive(Debug, Clone, Copy)]
struct Range {
    start: u32,
    length: u32,
}

/// One accepted rewrite of a whole direct-body paragraph.
///
/// `start` and `end` are the exact byte range the rewrite replaced in the
/// previous XML; the incremental layout only accepts a range that is exactly
/// one entry of the retained paragraph layout.
struct ParagraphSplice<'a> {
    start: usize,
    end: usize,
    replacement: &'a [u8],
}

/// One accepted splice reduced to the offsets the layout works in.
#[derive(Clone, Copy)]
struct SpliceBound {
    start: u32,
    end: u32,
    length: u32,
}

impl SpliceBound {
    /// Add this splice's length change to a running offset delta.
    fn accumulate(self, delta: i64) -> Option<i64> {
        delta
            .checked_add(i64::from(self.length))?
            .checked_add(i64::from(self.start))?
            .checked_sub(i64::from(self.end))
    }
}

/// A layout derived from an existing one by resizing rewritten paragraphs and
/// shifting every later direct-body range.
struct SplicedLayout {
    paragraphs: Arc<[Range]>,
    /// The source snapshot's per-paragraph inherited declarations, still exact
    /// because no splice adds or removes a namespace declaration.
    paragraph_namespaces: Arc<[NamespaceBindings]>,
    tables: Arc<[Range]>,
    block_controls: Arc<[Range]>,
    content_end: u32,
}

/// The structural shape of one direct-body child element.
#[derive(Clone, Copy)]
struct BodyChildShape {
    nodes: usize,
    max_depth: usize,
    root_tag_len: usize,
}

/// Shift direct-body sibling ranges across the accepted paragraph splices.
///
/// Direct body children never overlap, so a table or block content control is
/// either entirely before a rewritten paragraph or entirely after it. `None`
/// reports an overlap that only a full rescan can resolve.
fn shift_sibling_ranges(
    ranges: &Arc<[Range]>,
    bounds: &[SpliceBound],
) -> TransactionResult<Option<Arc<[Range]>>> {
    let (Some(first), Some(last)) = (bounds.first(), ranges.last()) else {
        return Ok(Some(Arc::clone(ranges)));
    };
    let Some(last_end) = last.start.checked_add(last.length) else {
        return Ok(None);
    };
    if last_end <= first.start {
        // Every sibling ends before the first rewritten paragraph, so the
        // retained range list is still exact and can stay shared.
        return Ok(Some(Arc::clone(ranges)));
    }
    let mut shifted = Vec::new();
    shifted
        .try_reserve_exact(ranges.len())
        .map_err(|source| crate::Error::Allocation {
            resource: "spliced document sibling ranges",
            source,
        })?;
    let mut index = 0usize;
    let mut delta = 0i64;
    for range in ranges.iter() {
        let Some(end) = range.start.checked_add(range.length) else {
            return Ok(None);
        };
        while let Some(bound) = bounds.get(index) {
            if bound.end > range.start {
                break;
            }
            let Some(next) = bound.accumulate(delta) else {
                return Ok(None);
            };
            delta = next;
            index = index.saturating_add(1);
        }
        if bounds.get(index).is_some_and(|bound| bound.start < end) {
            return Ok(None);
        }
        let Some(start) = shifted_offset(range.start, delta) else {
            return Ok(None);
        };
        shifted.push(Range {
            start,
            length: range.length,
        });
    }
    Ok(Some(shifted.into()))
}

/// Move one layout offset by a signed length change.
fn shifted_offset(offset: u32, delta: i64) -> Option<u32> {
    u32::try_from(i64::from(offset).checked_add(delta)?).ok()
}

/// Report whether the package writer would accept preserved main-document
/// bytes without re-serializing them.
///
/// The OPC writer audits every XML member it publishes for encoding,
/// well-formedness, exactly one document element, character data only inside
/// it, the absence of a DTD or DOCTYPE and its finite budgets
/// (`PackageWriter::audit_published_xml`). It no longer asserts the
/// repository's compact output contract against any member: change 0665 made
/// that movement on the eager route as an implementation interpretation of
/// change 0652's decision 2 (*"Original byte contract: loose the audit, accept
/// non-compact XMLs"*), after changes 0654 and 0657 established the same
/// source profile on the source-backed route. Change 0660 named this function
/// as the one place that had to change when it landed, and this is that change.
///
/// So a producer that indents its markup, or writes a line break after the XML
/// declaration — the witness change 0602 named, and the shape 53 of this
/// repository's 55 openable DOCX fixtures have — now publishes preserved. What
/// still sends a document down the whole-document route is a *structural*
/// refusal the writer would raise on the preserved bytes: the byte-order-marked
/// main document change 0650 froze is the one this repository's corpus carries.
///
/// The gate runs the auditor the writer runs, under the writer's limits, so it
/// cannot reach a different verdict than the writer. It is asked of the
/// *source* snapshot rather than the candidate because the compactor's own
/// output always satisfies the contract, so a source the writer accepts stays
/// acceptable once compacted paragraphs are spliced into it, and a source the
/// writer would refuse is refused whichever paragraphs the edit touched.
fn publication_accepts_preserved_xml(xml: &[u8]) -> bool {
    xml_minifier::audit::verify_source(xml, xml_minifier::audit::Limits::default()).is_ok()
}

/// Commit under [`CompactionPolicy::PreserveUnmodified`].
///
/// The route is the source gate's, exactly as before: when the package writer
/// accepts the *source* snapshot's bytes preserved
/// ([`publication_accepts_preserved_xml`]), the changed paragraphs are
/// compacted in place ([`compact_changed_paragraphs`]); otherwise the whole
/// document is compacted ([`compact_whole_document`]).
///
/// Change 0754 answers the gate with the pair audit
/// [`xml_minifier::audit::VerifiedSource::verify_replacement`], whose verdict
/// on the source is `verify_source`'s own, so the route cannot differ. When
/// the candidate passes too, the same pass has audited the candidate's exact
/// allocation, mostly inside the one element the edit replaced, and the
/// resulting proof travels with the snapshot to the package writer, which
/// then does not audit those bytes again. The package writer's audit stays the
/// last line of defence for everything else: bytes without a proof, and any
/// bytes other than the proof's own allocation, are audited there.
///
/// The candidate is compacted before the gate's verdict is known, because the
/// pair audit needs it. Compaction only reads the two snapshots, so when the
/// source fails the gate its candidate — or its error — is dropped as if it
/// had never been computed, and when the source passes a compaction error is
/// returned exactly where it was. A candidate whose bytes cannot be shared
/// (a managed projection) takes the historical gate and carries no proof.
///
/// Only [`crate::Package::apply_document_patch`] consumes the proof, and it
/// publishes only patches whose source snapshot carries no source identity
/// (`Package::document_snapshot`'s). A snapshot with a source identity is
/// published by the source-backed writer, which audits its own pair, so for
/// such a source the historical gate runs alone and no proof is built.
fn commit_preserving_unmodified(
    base: &Snapshot,
    projected: &Snapshot,
) -> TransactionResult<Snapshot> {
    if base.xml.identity().is_some() {
        return if publication_accepts_preserved_xml(base.xml_bytes()) {
            compact_changed_paragraphs(base, projected)
        } else {
            compact_whole_document(projected)
        };
    }
    let compacted = compact_changed_paragraphs(base, projected);
    let replacement = compacted
        .as_ref()
        .ok()
        .and_then(|candidate| candidate.shared_xml().ok());
    let (source_publishes, proof) = match replacement {
        Some(replacement) => match xml_minifier::audit::VerifiedSource::verify_replacement(
            base.xml_bytes(),
            replacement,
            xml_minifier::audit::Limits::default(),
        ) {
            Ok((proof, _how)) => (true, Some(proof)),
            // The source's own `verify_source` error: the gate refuses.
            Err(xml_minifier::audit::ReplacementError::Original(_)) => (false, None),
            // The source passes and the candidate does not. Publish it
            // unproven, so the writer's own audit refuses it where it always
            // has.
            Err(_) => (true, None),
        },
        None => (publication_accepts_preserved_xml(base.xml_bytes()), None),
    };
    if !source_publishes {
        return compact_whole_document(projected);
    }
    let candidate = compacted?;
    Ok(match proof {
        Some(proof) => candidate.with_publication_proof(proof),
        None => candidate,
    })
}

/// Compact the whole projected main document: the
/// [`CompactionPolicy::WholeDocument`] opt-in, and the preserving policy's
/// fallback when the source cannot publish preserved — byte for byte what
/// every committed edit published before the policy existed.
fn compact_whole_document(projected: &Snapshot) -> TransactionResult<Snapshot> {
    let source = std::str::from_utf8(projected.xml_bytes()).map_err(|error| {
        crate::Error::InvalidFormat(format!("changed main-document XML is not UTF-8: {error}"))
    })?;
    let compact = compact_changed_document_preserving_bom(source)?;
    projected.with_rewritten_xml(compact.into_bytes())
}

/// Compact a whole-document edit while carrying a leading UTF-8 byte order
/// mark through the quick-xml based compactor.
///
/// `quick-xml` consumes the mark before its first event. The transaction keeps
/// the original XML bytes as its source of truth, so the fallback publication
/// route must add the mark back after compaction just as the mutable writer
/// does. Paragraph-fragment compaction does not call this helper.
fn compact_changed_document_preserving_bom(source: &str) -> TransactionResult<String> {
    let Some(rest) = source.strip_prefix(UTF8_BYTE_ORDER_MARK_TEXT) else {
        return Ok(crate::writer::doc::compact_changed_document_xml(source)?);
    };
    let compact = crate::writer::doc::compact_changed_document_xml(rest)?;
    let mut output = String::with_capacity(UTF8_BYTE_ORDER_MARK_TEXT.len() + compact.len());
    output.push_str(UTF8_BYTE_ORDER_MARK_TEXT);
    output.push_str(&compact);
    Ok(output)
}

/// The `xml:space` attribute name, whose inherited value is the only state
/// [`crate::writer::doc::compact_changed_document_xml`] carries into a body
/// child from its ancestors.
const XML_SPACE_ATTRIBUTE: &[u8] = b"xml:space";

/// Compact only the direct-body paragraphs a committed edit changed, leaving
/// every other byte of the main document exactly as it was read.
///
/// This is [`CompactionPolicy::PreserveUnmodified`]. A paragraph is *changed*
/// when its bytes differ from the bytes of the source snapshot's paragraph at
/// the same position; a changed paragraph is compacted on its own, as a single
/// closed root, and spliced back through the derived-layout route so no
/// whole-document rescan is needed.
///
/// **Why a paragraph fragment compacts to the same bytes as its range would
/// under the whole-document pass.** The compactor is a plain `Reader`: it never
/// resolves namespace prefixes, so a fragment whose prefixes are declared on an
/// ancestor compacts exactly as it would in place. Its pending-whitespace and
/// text-run state is reset at every element boundary, and a direct-body
/// paragraph begins at one. Its only inherited state is the `xml:space` stack,
/// whose value at a direct-body paragraph comes from `w:document` and `w:body`
/// alone; when either of those start tags carries an `xml:space` attribute this
/// function declines and publishes the projection unchanged, rather than
/// compact a fragment under an inherited value it cannot see.
///
/// Declining is never a refusal: it publishes the bytes the edit produced,
/// which is what this policy promises for everything it does not compact.
fn compact_changed_paragraphs(
    base: &Snapshot,
    projected: &Snapshot,
) -> TransactionResult<Snapshot> {
    if base.paragraphs.len() != projected.paragraphs.len() || inherits_xml_space(projected) {
        return Ok(projected.clone());
    }
    let source = projected.xml_bytes();
    let original = base.xml_bytes();
    let mut replacements: Vec<(usize, usize, Vec<u8>)> = Vec::new();
    for (range, base_range) in projected.paragraphs.iter().zip(base.paragraphs.iter()) {
        let (Some(candidate), Some(previous)) = (
            paragraph_slice(source, *range),
            paragraph_slice(original, *base_range),
        ) else {
            return Ok(projected.clone());
        };
        if candidate == previous {
            continue;
        }
        let text = std::str::from_utf8(candidate).map_err(|error| {
            crate::Error::InvalidFormat(format!("changed paragraph XML is not UTF-8: {error}"))
        })?;
        let compact = crate::writer::doc::compact_changed_document_xml(text)?;
        if compact.as_bytes() == candidate {
            continue;
        }
        let (Ok(start), Ok(length)) = (usize::try_from(range.start), usize::try_from(range.length))
        else {
            return Ok(projected.clone());
        };
        let Some(end) = start.checked_add(length) else {
            return Ok(projected.clone());
        };
        replacements
            .try_reserve(1)
            .map_err(|source| crate::Error::Allocation {
                resource: "changed paragraph compaction plan",
                source,
            })?;
        replacements.push((start, end, compact.into_bytes()));
    }
    if replacements.is_empty() {
        return Ok(projected.clone());
    }
    let xml = replace_ranges(source, &replacements)?;
    let mut splices = Vec::new();
    splices
        .try_reserve_exact(replacements.len())
        .map_err(|source| crate::Error::Allocation {
            resource: "changed paragraph compaction splices",
            source,
        })?;
    splices.extend(
        replacements
            .iter()
            .map(|(start, end, replacement)| ParagraphSplice {
                start: *start,
                end: *end,
                replacement,
            }),
    );
    projected.with_spliced_paragraphs(xml, &splices)
}

/// Borrow one body-child range out of the bytes that carry it.
fn paragraph_slice(xml: &[u8], range: Range) -> Option<&[u8]> {
    let start = usize::try_from(range.start).ok()?;
    let end = start.checked_add(usize::try_from(range.length).ok()?)?;
    xml.get(start..end)
}

/// Report whether any ancestor of the direct-body children declares
/// `xml:space`.
///
/// Every direct-body child is a child of `w:body`, whose only ancestor is the
/// document root, so both start tags lie before the first direct-body child.
/// Searching those bytes for the attribute name is conservative: a match in an
/// attribute *value* also declines, which costs a compaction and preserves
/// bytes.
fn inherits_xml_space(snapshot: &Snapshot) -> bool {
    let prologue_end = [
        snapshot.paragraphs.first().map(|range| range.start),
        snapshot.tables.first().map(|range| range.start),
        snapshot.block_controls.first().map(|range| range.start),
        Some(snapshot.content_end),
    ]
    .into_iter()
    .flatten()
    .min()
    .unwrap_or(0);
    let Ok(prologue_end) = usize::try_from(prologue_end) else {
        return true;
    };
    let Some(prologue) = snapshot.xml_bytes().get(..prologue_end) else {
        return true;
    };
    prologue
        .windows(XML_SPACE_ATTRIBUTE.len())
        .any(|window| window == XML_SPACE_ATTRIBUTE)
}

/// Whether splicing `after` in place of `before` leaves every verdict of
/// [`scan_document`] unchanged.
///
/// The scanner reaches only four kinds of conclusion about the inside of a
/// direct body child: how the child itself is classified (from the resolved
/// namespace and local name of its root element), the running element count
/// against `MAX_DOCUMENT_NODES`, the nesting depth against
/// `MAX_DOCUMENT_DEPTH`, and the refusals for DTDs, processing instructions
/// and unbalanced nesting. Nothing else inside a body child reaches the
/// layout.
///
/// So when both fragments are exactly one balanced element that starts at
/// their first byte and ends at their last, their root start tags are
/// byte-identical, and the replacement introduces no additional element and no
/// additional nesting, the spliced document classifies that child identically,
/// cannot exceed a total it was already under, and gains no refusal. The bytes
/// around the splice are untouched, so the namespace bindings in scope at the
/// root are the same on both sides and the identical start tag therefore
/// resolves identically.
fn preserves_body_child_shape(before: &[u8], after: &[u8]) -> bool {
    let (Some(previous), Some(next)) = (body_child_shape(before), body_child_shape(after)) else {
        return false;
    };
    next.nodes <= previous.nodes
        && next.max_depth <= previous.max_depth
        && next.root_tag_len == previous.root_tag_len
        && before.get(..previous.root_tag_len) == after.get(..next.root_tag_len)
}

/// Measure one direct-body child fragment, or decline to measure it.
///
/// `None` means the fragment is not exactly one balanced element spanning all
/// of its bytes, or that it carries markup [`scan_document`] forbids. Callers
/// treat that as "fall back to the full rescan", never as a document refusal.
fn body_child_shape(fragment: &[u8]) -> Option<BodyChildShape> {
    let mut reader = Reader::from_reader(fragment);
    let mut shape = BodyChildShape {
        nodes: 0,
        max_depth: 0,
        root_tag_len: 0,
    };
    let mut depth = 0usize;
    let mut root_seen = false;
    let mut root_closed = false;
    loop {
        let event_start = usize::try_from(reader.buffer_position()).ok()?;
        let event = reader.read_event().ok()?;
        let event_end = usize::try_from(reader.buffer_position()).ok()?;
        match event {
            Event::Start(_) => {
                if root_closed {
                    return None;
                }
                depth = depth.checked_add(1)?;
                shape.nodes = shape.nodes.checked_add(1)?;
                shape.max_depth = shape.max_depth.max(depth);
                if !root_seen {
                    if event_start != 0 {
                        return None;
                    }
                    shape.root_tag_len = event_end;
                    root_seen = true;
                }
            },
            Event::Empty(_) => {
                if root_closed {
                    return None;
                }
                shape.nodes = shape.nodes.checked_add(1)?;
                shape.max_depth = shape.max_depth.max(depth.checked_add(1)?);
                if !root_seen {
                    if event_start != 0 {
                        return None;
                    }
                    shape.root_tag_len = event_end;
                    root_seen = true;
                    root_closed = true;
                }
            },
            Event::End(_) => {
                depth = depth.checked_sub(1)?;
                if depth == 0 {
                    root_closed = true;
                }
            },
            Event::Text(text) => {
                if (!root_seen || root_closed) && !text.is_empty() {
                    return None;
                }
            },
            Event::CData(_) | Event::Comment(_) | Event::GeneralRef(_) => {
                if !root_seen || root_closed {
                    return None;
                }
            },
            Event::Decl(_) | Event::DocType(_) | Event::PI(_) => return None,
            Event::Eof => break,
        }
    }
    (root_seen && root_closed && depth == 0).then_some(shape)
}

struct Layout {
    paragraphs: Vec<Range>,
    paragraph_namespaces: Vec<NamespaceBindings>,
    tables: Vec<Range>,
    block_controls: Vec<Range>,
    content_end: u32,
    conformance: Conformance,
}

#[derive(Debug, Clone, Copy)]
enum Conformance {
    Transitional,
    Strict,
}

impl Conformance {
    const fn namespace(self) -> &'static str {
        match self {
            Self::Transitional => "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
            Self::Strict => "http://purl.oclc.org/ooxml/wordprocessingml/main",
        }
    }
}

struct TextOwner {
    slots: Vec<TextSlot>,
    text: String,
}

struct ManagedParagraphPlan {
    range: Range,
    operation_index: usize,
    owner: TextOwner,
    _scan: ScanAdmission,
}

struct ManagedBatchDecision<'a> {
    replacement: &'a ParagraphTextReplacement,
    base_text: String,
    existing_index: Option<usize>,
    remove_existing: bool,
    add_new: bool,
    _scan: ScanAdmission,
}

struct TextSlot {
    start: usize,
    end: usize,
    prefix: Vec<u8>,
    local_name: Vec<u8>,
    characters: usize,
}

#[derive(Clone, Copy)]
enum ComplexFieldMarker {
    Begin,
    Separate,
    End,
}

enum FragmentPrefix {
    Unseen,
    Unprefixed,
    Prefixed(Vec<u8>),
}

impl FragmentPrefix {
    fn from_name(name: quick_xml::name::QName<'_>) -> Self {
        name.prefix().map_or(Self::Unprefixed, |prefix| {
            Self::Prefixed(prefix.into_inner().to_vec())
        })
    }
}

struct CellSelection<'a> {
    xml: &'a [u8],
    start: usize,
}

fn validate_hyperlink_replacements(
    replacements: &[HyperlinkTextReplacement],
) -> TransactionResult<()> {
    if replacements.len() > MAX_OPERATIONS {
        return Err(TransactionError::Limit {
            resource: "composite hyperlink replacements",
            max: MAX_OPERATIONS,
            actual: replacements.len(),
        });
    }
    let ambiguous_position = replacements
        .first()
        .map_or(0, |replacement| replacement.address.paragraph.get());
    if replacements.is_empty()
        || replacements
            .windows(2)
            .any(|pair| pair[0].address >= pair[1].address)
    {
        return Err(TransactionError::Refused {
            position: ambiguous_position,
            reason: Refusal::AmbiguousCompositeSelector,
        });
    }
    Ok(())
}

fn validate_paragraph_replacements(
    replacements: &[ParagraphTextReplacement],
) -> TransactionResult<()> {
    if replacements.len() > MAX_OPERATIONS {
        return Err(TransactionError::Limit {
            resource: "composite paragraph replacements",
            max: MAX_OPERATIONS,
            actual: replacements.len(),
        });
    }
    let ambiguous_position = replacements
        .first()
        .map_or(0, |replacement| replacement.position.get());
    if replacements.is_empty()
        || replacements
            .windows(2)
            .any(|pair| pair[0].position >= pair[1].position)
    {
        return Err(TransactionError::Refused {
            position: ambiguous_position,
            reason: Refusal::AmbiguousCompositeSelector,
        });
    }
    Ok(())
}

fn scan_document(xml: &[u8]) -> TransactionResult<Layout> {
    scan_document_with_context(xml, None)
}

/// Count the namespace-declaring attributes of one start or empty tag, for
/// the managed scan's namespace work charge.
///
/// Every counted key is `xmlns` or begins with `xmlns:`, and every key is a
/// slice of the tag's raw attribute bytes, so a tag whose attribute bytes do
/// not contain `xmlns` counts zero without iterating its attributes. The
/// count is otherwise the historical one, including its `filter_map` over
/// malformed attributes.
fn event_namespace_binding_count(event: &Event<'_>) -> usize {
    match event {
        Event::Start(element) | Event::Empty(element) => {
            if !contains_namespace_declaration_marker(element.attributes_raw()) {
                return 0;
            }
            element
                .attributes()
                .with_checks(false)
                .filter_map(Result::ok)
                .filter(|attribute| attribute.key.as_namespace_binding().is_some())
                .count()
        },
        _ => 0,
    }
}

/// Whether `bytes` contain `xmlns`, the prefix every namespace-declaring
/// attribute key starts with.
fn contains_namespace_declaration_marker(bytes: &[u8]) -> bool {
    bytes
        .windows(b"xmlns".len())
        .any(|window| window == b"xmlns")
}

/// Whether a direct-body child with this local name is classified by
/// [`scan_document_with_context`], which is only when it is in a
/// `WordprocessingML` namespace.
fn is_classified_body_child(local: &[u8]) -> bool {
    matches!(local, b"p" | b"tbl" | b"sdt" | b"sectPr")
}

fn scan_document_with_context(
    xml: &[u8],
    context: Option<&ExecutionContext>,
) -> TransactionResult<Layout> {
    // A plain `Reader` and the shared binding tracker replace `NsReader`
    // (change 0229's pattern for the paragraph text scanner, applied here by
    // change 0754). `NsReader::from_reader` is `Reader::from_reader` with the
    // default configuration, so the tokenizer and every tokenizer error are
    // unchanged, and `BindingTracker` reproduces the resolver's push, deferred
    // pop and namespace errors byte for byte. Events borrow the input instead
    // of being copied, and a name is resolved only where a verdict reads its
    // namespace: the document root, every `body`, and the four classified
    // direct-body children. `scan_document_with_context_nsreader_oracle`
    // keeps the previous implementation for the differential tests.
    let mut reader = Reader::from_reader(xml);
    let mut tracker = BindingTracker::new();
    let mut pending_pop = false;
    let mut paragraphs = Vec::new();
    // Declarations each direct-body paragraph inherits, for the parsers that
    // resolve a retained span in context (spec-gap branch). Consecutive
    // paragraphs share one snapshot while their scope is unchanged.
    let mut paragraph_namespaces = Vec::new();
    let mut namespace_capture = NamespaceCapture::default();
    let mut tables = Vec::new();
    let mut block_controls = Vec::new();
    let mut body_depth = None;
    let mut body_end = None;
    let mut final_section_start = None;
    let mut pending = None::<(bool, bool, bool, bool, usize, Option<NamespaceBindings>)>;
    let mut conformance = None;
    let mut saw_document = false;
    let mut depth = 0usize;
    let mut nodes = 0usize;
    // Namespace accounting is only needed for the managed work charge. Keep
    // this state out of ordinary/unmanaged layout scans so they retain their
    // prior allocation and parsing behavior.
    let mut namespace_bindings = context.map(|_| 2usize);
    let mut namespace_scopes = context.map(|_| Vec::<usize>::new());
    // quick-xml consumes a leading UTF-8 BOM before its first event and does
    // not include those bytes in `buffer_position()`. Layout ranges address
    // the retained source buffer, so carry the prefix into every event span
    // used below. This keeps managed and unmanaged scans on the same source
    // coordinates and preserves exact paragraph/table slices.
    let bom_offset =
        usize::from(xml.starts_with(UTF8_BYTE_ORDER_MARK)) * UTF8_BYTE_ORDER_MARK.len();

    loop {
        if let Some(context) = context {
            context.check().map_err(managed_execution)?;
        }
        let event_start = usize::try_from(reader.buffer_position())
            .map_err(|_conversion_error| {
                crate::Error::InvalidFormat("document offset does not fit usize".into())
            })?
            .checked_add(bom_offset)
            .ok_or_else(|| crate::Error::InvalidFormat("document offset overflow".into()))?;
        // `NsReader` applies the deferred pop of the previous `End` or `Empty`
        // scope when it is asked for the next event.
        if pending_pop {
            tracker.pop();
            pending_pop = false;
        }
        let event = reader
            .read_event()
            .map_err(|error| crate::Error::Xml(error.to_string()))?;
        // `NsReader` pushes a `Start`/`Empty` scope before it returns the
        // event, so a namespace error preempts the event exactly where its
        // `read_event` returned `Err`. The tracker's error text is the
        // `NamespaceError` display that `quick_xml::Error::Namespace`
        // forwards, so the `Error::Xml` message is unchanged.
        match &event {
            Event::Start(element) => tracker
                .push(element)
                .map_err(|error| crate::Error::Xml(error.to_string()))?,
            Event::Empty(element) => {
                tracker
                    .push(element)
                    .map_err(|error| crate::Error::Xml(error.to_string()))?;
                pending_pop = true;
            },
            Event::End(_) => pending_pop = true,
            _ => {},
        }
        let event_end = usize::try_from(reader.buffer_position())
            .map_err(|_conversion_error| {
                crate::Error::InvalidFormat("document offset does not fit usize".into())
            })?
            .checked_add(bom_offset)
            .ok_or_else(|| crate::Error::InvalidFormat("document offset overflow".into()))?;
        let event_bytes = event_end.saturating_sub(event_start).max(1);
        let (event_namespace_bindings, active_namespace_bindings) = if namespace_bindings.is_some()
        {
            let event_namespace_bindings = event_namespace_binding_count(&event);
            let current_bindings = namespace_bindings.ok_or_else(|| {
                crate::Error::InvalidFormat("managed namespace accounting state disappeared".into())
            })?;
            let active_namespace_bindings = current_bindings
                .checked_add(event_namespace_bindings)
                .ok_or(TransactionError::Limit {
                    resource: "document namespace binding count",
                    max: usize::MAX,
                    actual: usize::MAX,
                })?;
            (event_namespace_bindings, active_namespace_bindings)
        } else {
            (0, 0)
        };
        if let Some(context) = context {
            // NamespaceResolver resolves a qualified name by scanning the
            // in-scope binding list. Charge the observed event span against
            // that list before any later layout/index allocation; this keeps
            // namespace-heavy input bounded without blind XML-size-squared
            // prepayment for ordinary documents.
            let lookup_work = event_bytes
                .checked_mul(active_namespace_bindings.saturating_add(1))
                .ok_or(TransactionError::Limit {
                    resource: "document namespace lookup work",
                    max: usize::MAX,
                    actual: usize::MAX,
                })?;
            consume_managed(context, Resource::Work, lookup_work)?;
        }
        if matches!(&event, Event::Start(_))
            && let (Some(bindings), Some(scopes)) =
                (namespace_bindings.as_mut(), namespace_scopes.as_mut())
        {
            *bindings = active_namespace_bindings;
            scopes.push(event_namespace_bindings);
        }

        if matches!(event, Event::Start(_) | Event::Empty(_)) {
            nodes = nodes.checked_add(1).ok_or_else(|| {
                crate::Error::InvalidFormat("document element counter overflow".into())
            })?;
            if nodes > MAX_DOCUMENT_NODES {
                return Err(TransactionError::Limit {
                    resource: "XML elements",
                    max: MAX_DOCUMENT_NODES,
                    actual: nodes,
                });
            }
        }

        match event {
            Event::Start(element) => {
                depth = depth.checked_add(1).ok_or_else(|| {
                    crate::Error::InvalidFormat("document XML nesting is too deep".into())
                })?;
                if depth > MAX_DOCUMENT_DEPTH {
                    return Err(TransactionError::Limit {
                        resource: "XML depth",
                        max: MAX_DOCUMENT_DEPTH,
                        actual: depth,
                    });
                }
                // `QName::local_name` and `QName::prefix` split at the same
                // first colon as this.
                let (prefix, local) = split_qualified_name(element.name().into_inner());
                let direct_body_child = body_depth.is_some_and(|body| depth == body + 1);
                // Every verdict below that reads the namespace also requires
                // one of these local names, so any other element's name is
                // never resolved. Resolution has no effect but its result.
                let namespace = ((depth == 1 && local == b"document")
                    || local == b"body"
                    || (direct_body_child && is_classified_body_child(local)))
                .then(|| tracker.resolve_prefix_bytes(prefix));
                let is_word = namespace.as_ref().is_some_and(is_wordprocessing_namespace);
                if depth == 1 && is_word && local == b"document" {
                    saw_document = true;
                }
                if is_word && local == b"body" {
                    if depth != 2 || !saw_document {
                        return Err(crate::Error::InvalidFormat(
                            "WordprocessingML body is not a direct child of the document root"
                                .into(),
                        )
                        .into());
                    }
                    if body_depth.is_some() || body_end.is_some() {
                        return Err(crate::Error::InvalidFormat(
                            "main document contains multiple bodies".into(),
                        )
                        .into());
                    }
                    body_depth = Some(depth);
                    conformance = namespace.as_ref().and_then(conformance_from_namespace);
                } else if direct_body_child {
                    let is_paragraph = is_word && local == b"p";
                    let is_table = is_word && local == b"tbl";
                    let is_control = is_word && local == b"sdt";
                    let is_section = is_word && local == b"sectPr";
                    if final_section_start.is_some() {
                        return Err(crate::Error::InvalidFormat(
                            "body-final section properties are not the final body child".into(),
                        )
                        .into());
                    }
                    let namespaces = if is_paragraph {
                        Some(namespace_capture.capture_tracker(&tracker)?)
                    } else {
                        None
                    };
                    pending = Some((
                        is_paragraph,
                        is_table,
                        is_control,
                        is_section,
                        event_start,
                        namespaces,
                    ));
                }
            },
            Event::Empty(element) => {
                let child_depth = depth.checked_add(1).ok_or_else(|| {
                    crate::Error::InvalidFormat("document XML nesting is too deep".into())
                })?;
                if body_depth.is_some_and(|body| child_depth == body + 1) {
                    let (prefix, local) = split_qualified_name(element.name().into_inner());
                    let is_word = is_classified_body_child(local)
                        && is_wordprocessing_namespace(&tracker.resolve_prefix_bytes(prefix));
                    if final_section_start.is_some() {
                        return Err(crate::Error::InvalidFormat(
                            "body-final section properties are not the final body child".into(),
                        )
                        .into());
                    }
                    if is_word && local == b"p" {
                        paragraphs.push(checked_range(event_start, event_end)?);
                        paragraph_namespaces.push(namespace_capture.capture_tracker(&tracker)?);
                    }
                    if is_word && local == b"tbl" {
                        tables.push(checked_range(event_start, event_end)?);
                    }
                    if is_word && local == b"sdt" {
                        block_controls.push(checked_range(event_start, event_end)?);
                    }
                    if is_word && local == b"sectPr" {
                        final_section_start = Some(event_start);
                    }
                }
            },
            Event::End(element) => {
                if let Some((is_paragraph, is_table, is_control, is_section, start, namespaces)) =
                    pending.take_if(|_| body_depth.is_some_and(|body| depth == body + 1))
                {
                    if is_paragraph {
                        paragraphs.push(checked_range(start, event_end)?);
                        paragraph_namespaces.push(namespaces.ok_or_else(|| {
                            crate::Error::InvalidFormat(
                                "paragraph namespace capture disappeared".into(),
                            )
                        })?);
                    }
                    if is_table {
                        tables.push(checked_range(start, event_end)?);
                    }
                    if is_control {
                        block_controls.push(checked_range(start, event_end)?);
                    }
                    if is_section {
                        final_section_start = Some(start);
                    }
                    pending = None;
                }
                // The name resolves in the scope its own start tag opened:
                // the pop is deferred to the next read, as in `NsReader`.
                if body_depth == Some(depth)
                    && let (prefix, b"body") = split_qualified_name(element.name().into_inner())
                    && is_wordprocessing_namespace(&tracker.resolve_prefix_bytes(prefix))
                {
                    body_end = Some(event_start);
                    body_depth = None;
                }
                if let (Some(bindings), Some(scopes)) =
                    (namespace_bindings.as_mut(), namespace_scopes.as_mut())
                {
                    let scope_bindings = scopes.pop().ok_or_else(|| {
                        crate::Error::InvalidFormat("document namespace scope underflow".into())
                    })?;
                    *bindings = bindings.checked_sub(scope_bindings).ok_or_else(|| {
                        crate::Error::InvalidFormat("document namespace binding underflow".into())
                    })?;
                }
                depth = depth.checked_sub(1).ok_or_else(|| {
                    crate::Error::InvalidFormat("invalid document XML nesting".into())
                })?;
            },
            Event::DocType(_) => {
                return Err(crate::Error::InvalidFormat(
                    "DTD declarations are forbidden in a Word main document".into(),
                )
                .into());
            },
            Event::PI(_) => {
                return Err(crate::Error::InvalidFormat(
                    "processing instructions are forbidden in a Word main document".into(),
                )
                .into());
            },
            Event::Eof if depth != 0 || pending.is_some() => {
                return Err(crate::Error::InvalidFormat(
                    "unterminated Word main document XML".into(),
                )
                .into());
            },
            Event::Eof => break,
            Event::Text(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::GeneralRef(_) => {},
        }
    }
    let body_end_offset = body_end.ok_or_else(|| {
        crate::Error::InvalidFormat("main document has no WordprocessingML body".into())
    })?;
    let document_conformance = conformance.ok_or_else(|| {
        crate::Error::InvalidFormat("main document body has no supported namespace".into())
    })?;
    if !saw_document {
        return Err(crate::Error::InvalidFormat(
            "main document has no WordprocessingML document root".into(),
        )
        .into());
    }
    let content_end = u32::try_from(final_section_start.unwrap_or(body_end_offset)).map_err(
        |_conversion_error| {
            crate::Error::InvalidFormat("document insertion offset exceeds u32".into())
        },
    )?;
    Ok(Layout {
        paragraphs,
        paragraph_namespaces,
        tables,
        block_controls,
        content_end,
        conformance: document_conformance,
    })
}

#[cfg(test)]
fn oracle_event_namespace_binding_count(event: &Event<'_>) -> usize {
    match event {
        Event::Start(element) | Event::Empty(element) => element
            .attributes()
            .with_checks(false)
            .filter_map(Result::ok)
            .filter(|attribute| attribute.key.as_namespace_binding().is_some())
            .count(),
        _ => 0,
    }
}

/// Change-0754 differential oracle: the `NsReader` layout scanner that
/// [`scan_document_with_context`] replaced, retained test-only so the
/// tracker-driven scan is pinned to its layouts and errors (the change-0229
/// pattern of `extract_word_text_nsreader_oracle`).
#[cfg(test)]
fn scan_document_with_context_nsreader_oracle(
    xml: &[u8],
    context: Option<&ExecutionContext>,
) -> TransactionResult<Layout> {
    let mut reader = NsReader::from_reader(xml);
    let mut paragraphs = Vec::new();
    let mut paragraph_namespaces = Vec::new();
    let mut tables = Vec::new();
    let mut block_controls = Vec::new();
    let mut body_depth = None;
    let mut body_end = None;
    let mut final_section_start = None;
    let mut pending = None::<(bool, bool, bool, bool, usize, NamespaceBindings)>;
    let mut conformance = None;
    let mut saw_document = false;
    let mut namespace_capture = NamespaceCapture::default();
    let mut depth = 0usize;
    let mut nodes = 0usize;
    // Namespace accounting is only needed for the managed work charge. Keep
    // this state out of ordinary/unmanaged layout scans so they retain their
    // prior allocation and parsing behavior.
    let mut namespace_bindings = context.map(|_| 2usize);
    let mut namespace_scopes = context.map(|_| Vec::<usize>::new());
    // quick-xml consumes a leading UTF-8 BOM before its first event and does
    // not include those bytes in `buffer_position()`. Layout ranges address
    // the retained source buffer, so carry the prefix into every event span
    // used below. This keeps managed and unmanaged scans on the same source
    // coordinates and preserves exact paragraph/table slices.
    let bom_offset =
        usize::from(xml.starts_with(UTF8_BYTE_ORDER_MARK)) * UTF8_BYTE_ORDER_MARK.len();

    loop {
        if let Some(context) = context {
            context.check().map_err(managed_execution)?;
        }
        let event_start = usize::try_from(reader.buffer_position())
            .map_err(|_conversion_error| {
                crate::Error::InvalidFormat("document offset does not fit usize".into())
            })?
            .checked_add(bom_offset)
            .ok_or_else(|| crate::Error::InvalidFormat("document offset overflow".into()))?;
        let raw_event = reader
            .read_event()
            .map_err(|error| crate::Error::Xml(error.to_string()))?;
        let event_end = usize::try_from(reader.buffer_position())
            .map_err(|_conversion_error| {
                crate::Error::InvalidFormat("document offset does not fit usize".into())
            })?
            .checked_add(bom_offset)
            .ok_or_else(|| crate::Error::InvalidFormat("document offset overflow".into()))?;
        let event_bytes = event_end.saturating_sub(event_start).max(1);
        let (event_namespace_bindings, active_namespace_bindings) = if namespace_bindings.is_some()
        {
            let event_namespace_bindings = oracle_event_namespace_binding_count(&raw_event);
            let current_bindings = namespace_bindings.ok_or_else(|| {
                crate::Error::InvalidFormat("managed namespace accounting state disappeared".into())
            })?;
            let active_namespace_bindings = current_bindings
                .checked_add(event_namespace_bindings)
                .ok_or(TransactionError::Limit {
                    resource: "document namespace binding count",
                    max: usize::MAX,
                    actual: usize::MAX,
                })?;
            (event_namespace_bindings, active_namespace_bindings)
        } else {
            (0, 0)
        };
        if let Some(context) = context {
            // NamespaceResolver resolves a qualified name by scanning the
            // in-scope binding list. Charge the observed event span against
            // that list before any later layout/index allocation; this keeps
            // namespace-heavy input bounded without blind XML-size-squared
            // prepayment for ordinary documents.
            let lookup_work = event_bytes
                .checked_mul(active_namespace_bindings.saturating_add(1))
                .ok_or(TransactionError::Limit {
                    resource: "document namespace lookup work",
                    max: usize::MAX,
                    actual: usize::MAX,
                })?;
            consume_managed(context, Resource::Work, lookup_work)?;
        }
        // The event borrows the input slice rather than the reader, so
        // namespace resolution can borrow the reader's resolver.  Cloning the
        // resolver here would copy all in-scope namespace bindings once per
        // event and turn a linear scan into a quadratic allocation path for
        // adversarial namespace-heavy input.
        if matches!(&raw_event, Event::Start(_))
            && let (Some(bindings), Some(scopes)) =
                (namespace_bindings.as_mut(), namespace_scopes.as_mut())
        {
            *bindings = active_namespace_bindings;
            scopes.push(event_namespace_bindings);
        }
        let resolver = reader.resolver();
        let (namespace, event) = resolver.resolve_event(raw_event);

        if matches!(event, Event::Start(_) | Event::Empty(_)) {
            nodes = nodes.checked_add(1).ok_or_else(|| {
                crate::Error::InvalidFormat("document element counter overflow".into())
            })?;
            if nodes > MAX_DOCUMENT_NODES {
                return Err(TransactionError::Limit {
                    resource: "XML elements",
                    max: MAX_DOCUMENT_NODES,
                    actual: nodes,
                });
            }
        }

        match event {
            Event::Start(element) => {
                depth = depth.checked_add(1).ok_or_else(|| {
                    crate::Error::InvalidFormat("document XML nesting is too deep".into())
                })?;
                if depth > MAX_DOCUMENT_DEPTH {
                    return Err(TransactionError::Limit {
                        resource: "XML depth",
                        max: MAX_DOCUMENT_DEPTH,
                        actual: depth,
                    });
                }
                let is_word = is_wordprocessing_namespace(&namespace);
                let local = element.local_name();
                if depth == 1 && is_word && local.as_ref() == b"document" {
                    saw_document = true;
                }
                if is_word && local.as_ref() == b"body" {
                    if depth != 2 || !saw_document {
                        return Err(crate::Error::InvalidFormat(
                            "WordprocessingML body is not a direct child of the document root"
                                .into(),
                        )
                        .into());
                    }
                    if body_depth.is_some() || body_end.is_some() {
                        return Err(crate::Error::InvalidFormat(
                            "main document contains multiple bodies".into(),
                        )
                        .into());
                    }
                    body_depth = Some(depth);
                    conformance = conformance_from_namespace(&namespace);
                } else if body_depth.is_some_and(|body| depth == body + 1) {
                    let is_paragraph = is_word && local.as_ref() == b"p";
                    let is_table = is_word && local.as_ref() == b"tbl";
                    let is_control = is_word && local.as_ref() == b"sdt";
                    let is_section = is_word && local.as_ref() == b"sectPr";
                    if final_section_start.is_some() {
                        return Err(crate::Error::InvalidFormat(
                            "body-final section properties are not the final body child".into(),
                        )
                        .into());
                    }
                    pending = Some((
                        is_paragraph,
                        is_table,
                        is_control,
                        is_section,
                        event_start,
                        namespace_capture.capture(reader.resolver())?,
                    ));
                }
            },
            Event::Empty(element) => {
                let child_depth = depth.checked_add(1).ok_or_else(|| {
                    crate::Error::InvalidFormat("document XML nesting is too deep".into())
                })?;
                if body_depth.is_some_and(|body| child_depth == body + 1) {
                    let is_word = is_wordprocessing_namespace(&namespace);
                    let local = element.local_name();
                    if final_section_start.is_some() {
                        return Err(crate::Error::InvalidFormat(
                            "body-final section properties are not the final body child".into(),
                        )
                        .into());
                    }
                    if is_word && local.as_ref() == b"p" {
                        paragraphs.push(checked_range(event_start, event_end)?);
                        paragraph_namespaces.push(namespace_capture.capture(reader.resolver())?);
                    }
                    if is_word && local.as_ref() == b"tbl" {
                        tables.push(checked_range(event_start, event_end)?);
                    }
                    if is_word && local.as_ref() == b"sdt" {
                        block_controls.push(checked_range(event_start, event_end)?);
                    }
                    if is_word && local.as_ref() == b"sectPr" {
                        final_section_start = Some(event_start);
                    }
                }
            },
            Event::End(element) => {
                if body_depth.is_some_and(|body| depth == body + 1)
                    && let Some((is_paragraph, is_table, is_control, is_section, start, namespaces)) =
                        pending.take()
                {
                    if is_paragraph {
                        paragraphs.push(checked_range(start, event_end)?);
                        paragraph_namespaces.push(namespaces);
                    }
                    if is_table {
                        tables.push(checked_range(start, event_end)?);
                    }
                    if is_control {
                        block_controls.push(checked_range(start, event_end)?);
                    }
                    if is_section {
                        final_section_start = Some(start);
                    }
                    pending = None;
                }
                if body_depth == Some(depth)
                    && is_wordprocessing_namespace(&namespace)
                    && element.local_name().as_ref() == b"body"
                {
                    body_end = Some(event_start);
                    body_depth = None;
                }
                if let (Some(bindings), Some(scopes)) =
                    (namespace_bindings.as_mut(), namespace_scopes.as_mut())
                {
                    let scope_bindings = scopes.pop().ok_or_else(|| {
                        crate::Error::InvalidFormat("document namespace scope underflow".into())
                    })?;
                    *bindings = bindings.checked_sub(scope_bindings).ok_or_else(|| {
                        crate::Error::InvalidFormat("document namespace binding underflow".into())
                    })?;
                }
                depth = depth.checked_sub(1).ok_or_else(|| {
                    crate::Error::InvalidFormat("invalid document XML nesting".into())
                })?;
            },
            Event::DocType(_) => {
                return Err(crate::Error::InvalidFormat(
                    "DTD declarations are forbidden in a Word main document".into(),
                )
                .into());
            },
            Event::PI(_) => {
                return Err(crate::Error::InvalidFormat(
                    "processing instructions are forbidden in a Word main document".into(),
                )
                .into());
            },
            Event::Eof if depth != 0 || pending.is_some() => {
                return Err(crate::Error::InvalidFormat(
                    "unterminated Word main document XML".into(),
                )
                .into());
            },
            Event::Eof => break,
            Event::Text(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::GeneralRef(_) => {},
        }
    }
    let body_end_offset = body_end.ok_or_else(|| {
        crate::Error::InvalidFormat("main document has no WordprocessingML body".into())
    })?;
    let document_conformance = conformance.ok_or_else(|| {
        crate::Error::InvalidFormat("main document body has no supported namespace".into())
    })?;
    if !saw_document {
        return Err(crate::Error::InvalidFormat(
            "main document has no WordprocessingML document root".into(),
        )
        .into());
    }
    let content_end = u32::try_from(final_section_start.unwrap_or(body_end_offset)).map_err(
        |_conversion_error| {
            crate::Error::InvalidFormat("document insertion offset exceeds u32".into())
        },
    )?;
    Ok(Layout {
        paragraphs,
        paragraph_namespaces,
        tables,
        block_controls,
        content_end,
        conformance: document_conformance,
    })
}

fn conformance_from_namespace(namespace: &ResolveResult<'_>) -> Option<Conformance> {
    match namespace {
        ResolveResult::Bound(Namespace(uri)) if *uri == WORDPROCESSINGML_NAMESPACE => {
            Some(Conformance::Transitional)
        },
        ResolveResult::Bound(Namespace(uri)) if *uri == STRICT_WORDPROCESSINGML_NAMESPACE => {
            Some(Conformance::Strict)
        },
        ResolveResult::Bound(_) | ResolveResult::Unbound | ResolveResult::Unknown(_) => None,
    }
}

fn checked_range(start: usize, end: usize) -> TransactionResult<Range> {
    Ok(Range {
        start: u32::try_from(start).map_err(|_conversion_error| {
            crate::Error::InvalidFormat("paragraph offset exceeds u32".into())
        })?,
        length: u32::try_from(
            end.checked_sub(start)
                .ok_or_else(|| crate::Error::InvalidFormat("paragraph range underflow".into()))?,
        )
        .map_err(|_conversion_error| {
            crate::Error::InvalidFormat("paragraph length exceeds u32".into())
        })?,
    })
}

struct RevisionFragmentInfo {
    root_open_end: usize,
    root_close_start: usize,
    child_ranges: Vec<(usize, usize)>,
    namespace_declarations: Vec<Vec<u8>>,
    scope_attributes: Vec<Vec<u8>>,
    deleted_text_names: Vec<(usize, usize, Vec<u8>)>,
}

fn rewrite_revision_paragraph(
    paragraph_xml: &[u8],
    selector: RevisionSelector,
    action: RevisionAction,
    inherited_namespaces: &[(Option<Vec<u8>>, Vec<u8>)],
) -> TransactionResult<Vec<u8>> {
    if paragraph_has_revision_range_marker(paragraph_xml, inherited_namespaces).map_err(
        |reason| TransactionError::Refused {
            position: selector.paragraph.get(),
            reason,
        },
    )? {
        return Err(TransactionError::Refused {
            position: selector.paragraph.get(),
            reason: Refusal::RevisionDependency,
        });
    }
    let revision_range = select_direct_child_with_context(
        paragraph_xml,
        b"p",
        selector.kind.local_name(),
        selector.revision,
        Refusal::RevisionNotFound,
        inherited_namespaces,
    )
    .map_err(|reason| TransactionError::Refused {
        position: selector.paragraph.get(),
        reason,
    })?;
    let revision_start = usize::try_from(revision_range.start).map_err(|_error| {
        TransactionError::Document(crate::Error::InvalidFormat(
            "revision offset does not fit usize".into(),
        ))
    })?;
    let revision_end = revision_start
        .checked_add(usize::try_from(revision_range.length).map_err(|_error| {
            TransactionError::Document(crate::Error::InvalidFormat(
                "revision length does not fit usize".into(),
            ))
        })?)
        .ok_or_else(|| {
            TransactionError::Document(crate::Error::InvalidFormat(
                "revision range overflows usize".into(),
            ))
        })?;
    let revision_xml = checked_slice(
        paragraph_xml,
        revision_start,
        revision_end,
        "tracked revision",
    )?;
    let replacement =
        rewrite_revision_fragment(revision_xml, selector.kind, action, inherited_namespaces)
            .map_err(|reason| TransactionError::Refused {
                position: selector.paragraph.get(),
                reason,
            })?;
    replace_range(paragraph_xml, revision_start, revision_end, &replacement)
}

fn paragraph_has_revision_range_marker(
    paragraph_xml: &[u8],
    inherited_namespaces: &[(Option<Vec<u8>>, Vec<u8>)],
) -> Result<bool, Refusal> {
    if paragraph_xml.len() > MAX_DOCUMENT_XML_BYTES {
        return Err(Refusal::RevisionDependency);
    }
    let mut reader = NsReader::from_reader(paragraph_xml);
    for (prefix, namespace) in inherited_namespaces {
        let prefix = prefix
            .as_deref()
            .map_or(PrefixDeclaration::Default, PrefixDeclaration::Named);
        reader
            .resolver_mut()
            .add(prefix, Namespace(namespace))
            .map_err(|_error| Refusal::RevisionDependency)?;
    }
    let mut depth = 0usize;
    let mut nodes = 0usize;
    loop {
        let (namespace, event) = reader
            .read_resolved_event()
            .map_err(|_error| Refusal::RevisionDependency)?;
        match event {
            Event::Start(element) => {
                depth = depth.checked_add(1).ok_or(Refusal::RevisionDependency)?;
                if depth > MAX_DOCUMENT_DEPTH {
                    return Err(Refusal::RevisionDependency);
                }
                nodes = nodes.checked_add(1).ok_or(Refusal::RevisionDependency)?;
                if nodes > MAX_DOCUMENT_NODES {
                    return Err(Refusal::RevisionDependency);
                }
                if depth == 2
                    && is_wordprocessing_namespace(&namespace)
                    && is_revision_range_marker(element.local_name().as_ref())
                {
                    return Ok(true);
                }
            },
            Event::Empty(element) => {
                let child_depth = depth.checked_add(1).ok_or(Refusal::RevisionDependency)?;
                if child_depth > MAX_DOCUMENT_DEPTH {
                    return Err(Refusal::RevisionDependency);
                }
                nodes = nodes.checked_add(1).ok_or(Refusal::RevisionDependency)?;
                if nodes > MAX_DOCUMENT_NODES {
                    return Err(Refusal::RevisionDependency);
                }
                if child_depth == 2
                    && is_wordprocessing_namespace(&namespace)
                    && is_revision_range_marker(element.local_name().as_ref())
                {
                    return Ok(true);
                }
            },
            Event::End(_) => {
                depth = depth.checked_sub(1).ok_or(Refusal::RevisionDependency)?;
            },
            Event::Eof => return Ok(false),
            Event::Decl(_) | Event::DocType(_) => return Err(Refusal::RevisionDependency),
            Event::Text(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::PI(_)
            | Event::GeneralRef(_) => {},
        }
    }
}

fn is_revision_range_marker(local: &[u8]) -> bool {
    matches!(
        local,
        b"moveFromRangeStart"
            | b"moveFromRangeEnd"
            | b"moveToRangeStart"
            | b"moveToRangeEnd"
            | b"commentRangeStart"
            | b"commentRangeEnd"
            | b"permStart"
            | b"permEnd"
            | b"bookmarkStart"
            | b"bookmarkEnd"
            | b"customXmlInsRangeStart"
            | b"customXmlInsRangeEnd"
            | b"customXmlDelRangeStart"
            | b"customXmlDelRangeEnd"
            | b"customXmlMoveFromRangeStart"
            | b"customXmlMoveFromRangeEnd"
            | b"customXmlMoveToRangeStart"
            | b"customXmlMoveToRangeEnd"
    )
}

fn rewrite_revision_fragment(
    xml: &[u8],
    kind: RevisionKind,
    action: RevisionAction,
    inherited_namespaces: &[(Option<Vec<u8>>, Vec<u8>)],
) -> Result<Vec<u8>, Refusal> {
    let info = scan_revision_fragment(xml, kind, inherited_namespaces)?;
    if action.removes_content(kind) {
        return Ok(Vec::new());
    }
    if info.root_close_start < info.root_open_end {
        return Err(Refusal::RevisionDependency);
    }
    let inner_range = info.root_open_end..info.root_close_start;
    let inner_source = xml
        .get(inner_range.clone())
        .ok_or(Refusal::RevisionDependency)?;
    let declaration_bytes = info
        .namespace_declarations
        .iter()
        .map(Vec::len)
        .chain(info.scope_attributes.iter().map(Vec::len))
        .try_fold(0usize, |size, length| {
            length
                .checked_add(1)
                .and_then(|length| size.checked_add(length))
        })
        .ok_or(Refusal::RevisionDependency)?;
    let expansion = info
        .child_ranges
        .len()
        .checked_mul(declaration_bytes)
        .and_then(|size| inner_source.len().checked_add(size))
        .ok_or(Refusal::RevisionDependency)?;
    if expansion > MAX_DOCUMENT_XML_BYTES {
        return Err(Refusal::RevisionDependency);
    }
    let mut inner = Vec::new();
    inner
        .try_reserve_exact(inner_source.len())
        .map_err(|_error| Refusal::RevisionDependency)?;
    inner.extend_from_slice(inner_source);
    let namespace_declarations = raw_attribute_fragments(&info.namespace_declarations)?;
    let root_namespace_bindings = raw_namespace_bindings(&namespace_declarations)?;
    let scope_attributes = raw_attribute_fragments(&info.scope_attributes)?;
    let mut replacements = Vec::new();
    let replace_deleted_text =
        matches!(kind, RevisionKind::Deletion) && matches!(action, RevisionAction::Reject);
    let replacement_count = info
        .deleted_text_names
        .len()
        .checked_add(info.child_ranges.len())
        .ok_or(Refusal::RevisionDependency)?;
    replacements
        .try_reserve_exact(replacement_count)
        .map_err(|_error| Refusal::RevisionDependency)?;
    if replace_deleted_text {
        for (start, end, replacement) in info.deleted_text_names {
            let start = start
                .checked_sub(info.root_open_end)
                .ok_or(Refusal::RevisionDependency)?;
            let end = end
                .checked_sub(info.root_open_end)
                .ok_or(Refusal::RevisionDependency)?;
            replacements.push((start, end, replacement));
        }
    }
    for (child_start, child_end) in info.child_ranges {
        let child_start = child_start
            .checked_sub(info.root_open_end)
            .ok_or(Refusal::RevisionDependency)?;
        let tag_end = child_end
            .checked_sub(info.root_open_end)
            .ok_or(Refusal::RevisionDependency)?;
        let opening = inner
            .get(child_start..tag_end)
            .ok_or(Refusal::RevisionDependency)?;
        let existing = raw_tag_attributes(opening)?;
        let mut child_namespace_declarations = Vec::new();
        child_namespace_declarations
            .try_reserve_exact(existing.len())
            .map_err(|_error| Refusal::RevisionDependency)?;
        for attribute in &existing {
            if is_namespace_declaration_name(&attribute.name) {
                child_namespace_declarations.push(attribute);
            }
        }
        let name_end = raw_tag_name_end(opening)?
            .checked_add(child_start)
            .ok_or(Refusal::RevisionDependency)?;
        let mut insertion = Vec::new();
        insertion
            .try_reserve_exact(declaration_bytes)
            .map_err(|_error| Refusal::RevisionDependency)?;
        for declaration in &namespace_declarations {
            if let Some(current) = existing.iter().find(|item| item.name == declaration.name) {
                if current.value != declaration.value {
                    return Err(Refusal::RevisionDependency);
                }
                continue;
            }
            insertion.push(b' ');
            insertion.extend_from_slice(&declaration.raw);
        }
        for scope_attribute in &scope_attributes {
            let expected_namespace = raw_attribute_namespace_kind(
                &scope_attribute.name,
                &[],
                &root_namespace_bindings,
                inherited_namespaces,
            )?;
            let child_namespace = raw_attribute_namespace_kind(
                &scope_attribute.name,
                &child_namespace_declarations,
                &root_namespace_bindings,
                inherited_namespaces,
            )?;
            if child_namespace != expected_namespace {
                return Err(Refusal::RevisionDependency);
            }
            let mut matched = false;
            for current in &existing {
                if is_namespace_declaration_name(&current.name) {
                    continue;
                }
                let current_namespace = raw_attribute_namespace_kind(
                    &current.name,
                    &child_namespace_declarations,
                    &root_namespace_bindings,
                    inherited_namespaces,
                )?;
                let current_local = raw_attribute_local_name(&current.name)?;
                if current_namespace == expected_namespace
                    && current_local == raw_attribute_local_name(&scope_attribute.name)?
                {
                    if matched || current.value != scope_attribute.value {
                        return Err(Refusal::RevisionDependency);
                    }
                    matched = true;
                }
            }
            if matched {
                continue;
            }
            insertion.push(b' ');
            insertion.extend_from_slice(&scope_attribute.raw);
        }
        if !insertion.is_empty() {
            replacements.push((name_end, name_end, insertion));
        }
    }
    replacements.sort_by_key(|(start, _, _)| *start);
    replace_ranges(&inner, &replacements).map_err(|_error| Refusal::RevisionDependency)
}

fn scan_revision_fragment(
    xml: &[u8],
    kind: RevisionKind,
    inherited_namespaces: &[(Option<Vec<u8>>, Vec<u8>)],
) -> Result<RevisionFragmentInfo, Refusal> {
    let mut reader = NsReader::from_reader(xml);
    for (prefix, namespace) in inherited_namespaces {
        let prefix = prefix
            .as_deref()
            .map_or(PrefixDeclaration::Default, PrefixDeclaration::Named);
        reader
            .resolver_mut()
            .add(prefix, Namespace(namespace))
            .map_err(|_error| Refusal::RevisionDependency)?;
    }
    let mut root_started = false;
    let mut root_closed = false;
    let mut root_open_end = None;
    let mut root_close_start = None;
    let mut child_ranges = Vec::new();
    let mut namespace_declarations = Vec::new();
    let mut scope_attributes = Vec::new();
    let mut deleted_text_names = Vec::new();
    let mut run_depth = None;
    let mut text_depth = None;
    let mut opaque_depth = None;
    let mut depth = 0usize;
    let mut nodes = 0usize;

    loop {
        let event_start = usize::try_from(reader.buffer_position())
            .map_err(|_error| Refusal::RevisionDependency)?;
        let raw_event = reader
            .read_event()
            .map_err(|_error| Refusal::RevisionDependency)?;
        let (namespace, event) = reader.resolver().resolve_event(raw_event);
        let event_end = usize::try_from(reader.buffer_position())
            .map_err(|_error| Refusal::RevisionDependency)?;
        if matches!(event, Event::Start(_) | Event::Empty(_)) {
            nodes = nodes.checked_add(1).ok_or(Refusal::RevisionDependency)?;
            if nodes > MAX_DOCUMENT_NODES {
                return Err(Refusal::RevisionDependency);
            }
            if let Event::Start(element) | Event::Empty(element) = &event {
                validate_revision_element_attributes(element, reader.resolver(), reader.decoder())?;
            }
        }

        match event {
            Event::Start(element) => {
                depth = depth.checked_add(1).ok_or(Refusal::RevisionDependency)?;
                if depth > MAX_DOCUMENT_DEPTH {
                    return Err(Refusal::RevisionDependency);
                }
                if !root_started {
                    if depth != 1
                        || !is_wordprocessing_namespace(&namespace)
                        || element.local_name().as_ref() != kind.local_name()
                    {
                        return Err(Refusal::RevisionNotFound);
                    }
                    root_started = true;
                    root_open_end = Some(event_end);
                    namespace_declarations = raw_namespace_declarations(
                        xml.get(event_start..event_end)
                            .ok_or(Refusal::RevisionDependency)?,
                    )?;
                    scope_attributes = raw_revision_scope_attributes(
                        xml.get(event_start..event_end)
                            .ok_or(Refusal::RevisionDependency)?,
                        &element,
                        reader.resolver(),
                        reader.decoder(),
                    )?;
                    continue;
                }
                if root_closed || depth == 1 {
                    return Err(Refusal::RevisionDependency);
                }
                if depth == 2 {
                    if !is_word_local(&namespace, element.name(), b"r") {
                        return Err(Refusal::RevisionDependency);
                    }
                    child_ranges
                        .try_reserve(1)
                        .map_err(|_error| Refusal::RevisionDependency)?;
                    child_ranges.push((event_start, event_end));
                    run_depth = Some(depth);
                } else if text_depth.is_some() {
                    return Err(Refusal::RevisionDependency);
                } else if run_depth == Some(depth - 1) {
                    let mut deleted_name = None;
                    inspect_revision_run_child(
                        xml,
                        &namespace,
                        element.name(),
                        kind,
                        false,
                        &mut text_depth,
                        depth,
                        Some((&mut deleted_name, event_start)),
                    )?;
                    if let Some((start, end, replacement)) = deleted_name {
                        deleted_text_names
                            .try_reserve(1)
                            .map_err(|_error| Refusal::RevisionDependency)?;
                        deleted_text_names.push((start, end, replacement));
                    }
                    if !is_wordprocessing_namespace(&namespace) {
                        opaque_depth = Some(depth);
                    }
                } else {
                    inspect_revision_descendant(&namespace, element.name(), depth)?;
                }
            },
            Event::Empty(element) => {
                let child_depth = depth.checked_add(1).ok_or(Refusal::RevisionDependency)?;
                if !root_started {
                    if depth != 0
                        || !is_wordprocessing_namespace(&namespace)
                        || element.local_name().as_ref() != kind.local_name()
                    {
                        return Err(Refusal::RevisionNotFound);
                    }
                    root_started = true;
                    root_closed = true;
                    root_open_end = Some(event_end);
                    root_close_start = Some(event_end);
                    namespace_declarations = raw_namespace_declarations(
                        xml.get(event_start..event_end)
                            .ok_or(Refusal::RevisionDependency)?,
                    )?;
                    scope_attributes = raw_revision_scope_attributes(
                        xml.get(event_start..event_end)
                            .ok_or(Refusal::RevisionDependency)?,
                        &element,
                        reader.resolver(),
                        reader.decoder(),
                    )?;
                } else if root_closed {
                    return Err(Refusal::RevisionDependency);
                } else if child_depth == 2 {
                    if !is_word_local(&namespace, element.name(), b"r") {
                        return Err(Refusal::RevisionDependency);
                    }
                    child_ranges
                        .try_reserve(1)
                        .map_err(|_error| Refusal::RevisionDependency)?;
                    child_ranges.push((event_start, event_end));
                } else if text_depth.is_some() {
                    return Err(Refusal::RevisionDependency);
                } else if run_depth == Some(child_depth - 1) {
                    let mut deleted_name = None;
                    inspect_revision_run_child(
                        xml,
                        &namespace,
                        element.name(),
                        kind,
                        true,
                        &mut text_depth,
                        child_depth,
                        Some((&mut deleted_name, event_start)),
                    )?;
                    if let Some((start, end, replacement)) = deleted_name {
                        deleted_text_names
                            .try_reserve(1)
                            .map_err(|_error| Refusal::RevisionDependency)?;
                        deleted_text_names.push((start, end, replacement));
                    }
                } else {
                    inspect_revision_descendant(&namespace, element.name(), child_depth)?;
                }
            },
            Event::End(element) => {
                if !root_started || root_closed || depth == 0 {
                    return Err(Refusal::RevisionDependency);
                }
                if depth == 1 {
                    if !is_word_local(&namespace, element.name(), kind.local_name()) {
                        return Err(Refusal::RevisionDependency);
                    }
                    root_close_start = Some(event_start);
                    root_closed = true;
                    depth = 0;
                } else {
                    if let Some(text_depth_value) = text_depth {
                        if text_depth_value == depth {
                            if matches!(kind, RevisionKind::Deletion)
                                && is_word_local(&namespace, element.name(), b"delText")
                            {
                                let (start, end, replacement) =
                                    revision_tag_name_replacement(xml, event_start, true)?;
                                deleted_text_names
                                    .try_reserve(1)
                                    .map_err(|_error| Refusal::RevisionDependency)?;
                                deleted_text_names.push((start, end, replacement));
                            }
                            text_depth = None;
                        } else if text_depth_value < depth {
                            return Err(Refusal::RevisionDependency);
                        }
                    }
                    if opaque_depth == Some(depth) {
                        opaque_depth = None;
                    }
                    if run_depth == Some(depth) {
                        run_depth = None;
                    }
                    depth = depth.checked_sub(1).ok_or(Refusal::RevisionDependency)?;
                }
            },
            Event::Text(event_text) => {
                validate_revision_text_event(Event::Text(event_text.borrow()))?;
                if text_depth.is_none()
                    && opaque_depth.is_none()
                    && !event_text.as_ref().iter().all(u8::is_ascii_whitespace)
                {
                    return Err(Refusal::RevisionDependency);
                }
            },
            Event::CData(event_text) => {
                validate_revision_text_event(Event::CData(event_text.borrow()))?;
                if text_depth.is_none() && opaque_depth.is_none() {
                    return Err(Refusal::RevisionDependency);
                }
            },
            Event::GeneralRef(reference) => {
                validate_revision_text_event(Event::GeneralRef(reference.borrow()))?;
                if text_depth.is_none() && opaque_depth.is_none() {
                    return Err(Refusal::RevisionDependency);
                }
            },
            Event::Eof => break,
            Event::Decl(_) | Event::DocType(_) => return Err(Refusal::RevisionDependency),
            Event::Comment(_) | Event::PI(_) => {},
        }
    }
    if !root_started
        || !root_closed
        || depth != 0
        || text_depth.is_some()
        || opaque_depth.is_some()
        || run_depth.is_some()
    {
        return Err(Refusal::RevisionDependency);
    }
    let root_open_end = root_open_end.ok_or(Refusal::RevisionDependency)?;
    Ok(RevisionFragmentInfo {
        root_open_end,
        root_close_start: root_close_start.unwrap_or(root_open_end),
        child_ranges,
        namespace_declarations,
        scope_attributes,
        deleted_text_names,
    })
}

fn inspect_revision_run_child(
    xml: &[u8],
    namespace: &ResolveResult<'_>,
    name: quick_xml::name::QName<'_>,
    kind: RevisionKind,
    empty: bool,
    text_depth: &mut Option<usize>,
    depth: usize,
    deleted_name: Option<(&mut Option<(usize, usize, Vec<u8>)>, usize)>,
) -> Result<(), Refusal> {
    if matches!(namespace, ResolveResult::Unknown(_)) {
        return Err(Refusal::RevisionDependency);
    }
    if !is_wordprocessing_namespace(namespace) {
        // Extension children are retained byte-for-byte. Any known Word
        // dependency nested below them is still rejected by the descendant
        // scanner, so this does not silently reinterpret Word semantics.
        return Ok(());
    }
    let local = name.local_name();
    let local = local.as_ref();
    if local == b"rPr" {
        return Ok(());
    }
    let expected_text = match kind {
        RevisionKind::Insertion => b"t".as_slice(),
        RevisionKind::Deletion => b"delText".as_slice(),
    };
    if local == expected_text {
        if !empty {
            if text_depth.replace(depth).is_some() {
                return Err(Refusal::RevisionDependency);
            }
        }
        if matches!(kind, RevisionKind::Deletion) {
            if let Some((slot, event_start)) = deleted_name {
                let (start, end, replacement) =
                    revision_tag_name_replacement(xml, event_start, false)?;
                *slot = Some((start, end, replacement));
            }
        }
        return Ok(());
    }
    if is_word_local(namespace, name, b"t") || is_word_local(namespace, name, b"delText") {
        return Err(Refusal::RevisionDependency);
    }
    if matches!(
        local,
        b"tab" | b"br" | b"cr" | b"noBreakHyphen" | b"softHyphen"
    ) {
        if !empty {
            return Err(Refusal::RevisionDependency);
        }
        return Ok(());
    }
    if unsupported_revision_word_local(local) {
        return Err(Refusal::RevisionDependency);
    }
    Err(Refusal::RevisionDependency)
}

fn validate_revision_text_event(event: Event<'_>) -> Result<(), Refusal> {
    let text = match event {
        Event::Text(text) => {
            let encoded = text
                .xml_content(XmlVersion::Explicit1_0)
                .map_err(|_error| Refusal::RevisionDependency)?;
            quick_xml::escape::unescape(&encoded)
                .map_err(|_error| Refusal::RevisionDependency)?
                .into_owned()
        },
        Event::CData(text) => text
            .xml_content(XmlVersion::Explicit1_0)
            .map_err(|_error| Refusal::RevisionDependency)?
            .into_owned(),
        Event::GeneralRef(reference) => litchi_ooxml_common::xml::decode_xml_reference(&reference)
            .map_err(|_error| Refusal::RevisionDependency)?,
        _ => return Ok(()),
    };
    if text.chars().all(is_legal_revision_xml_character) {
        Ok(())
    } else {
        Err(Refusal::RevisionDependency)
    }
}

fn validate_revision_element_attributes(
    element: &quick_xml::events::BytesStart<'_>,
    resolver: &quick_xml::name::NamespaceResolver,
    decoder: quick_xml::encoding::Decoder,
) -> Result<(), Refusal> {
    let mut attributes = 0usize;
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|_error| Refusal::RevisionDependency)?;
        if attribute.value.len() > MAX_REVISION_METADATA_VALUE_BYTES
            || attribute.value.contains(&b'<')
        {
            return Err(Refusal::RevisionDependency);
        }
        let decoded = attribute
            .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
            .map_err(|_error| Refusal::RevisionDependency)?;
        if !decoded.chars().all(is_legal_revision_xml_character) {
            return Err(Refusal::RevisionDependency);
        }
        attributes = attributes
            .checked_add(1)
            .ok_or(Refusal::RevisionDependency)?;
        if attributes > MAX_REVISION_ATTRIBUTES {
            return Err(Refusal::RevisionDependency);
        }
        if attribute.key.as_ref() == b"xmlns" || attribute.key.as_ref().starts_with(b"xmlns:") {
            continue;
        }
        if matches!(
            resolver.resolve_attribute(attribute.key).0,
            ResolveResult::Unknown(_)
        ) {
            return Err(Refusal::RevisionDependency);
        }
    }
    Ok(())
}

fn is_legal_revision_xml_character(character: char) -> bool {
    matches!(
        character as u32,
        0x9 | 0xA | 0xD | 0x20..=0xD7FF | 0xE000..=0xFFFD | 0x10000..=0x10FFFF
    )
}

fn inspect_revision_descendant(
    namespace: &ResolveResult<'_>,
    name: quick_xml::name::QName<'_>,
    depth: usize,
) -> Result<(), Refusal> {
    if matches!(namespace, ResolveResult::Unknown(_)) {
        return Err(Refusal::RevisionDependency);
    }
    if !is_wordprocessing_namespace(namespace) {
        return Ok(());
    }
    let local = name.local_name();
    let local = local.as_ref();
    if matches!(local, b"t" | b"delText" | b"r" | b"ins" | b"del")
        || unsupported_revision_word_local(local)
    {
        return Err(Refusal::RevisionDependency);
    }
    if depth < 3 {
        return Err(Refusal::RevisionDependency);
    }
    Ok(())
}

fn unsupported_revision_word_local(local: &[u8]) -> bool {
    matches!(
        local,
        b"moveFrom"
            | b"moveTo"
            | b"moveFromRangeStart"
            | b"moveFromRangeEnd"
            | b"moveToRangeStart"
            | b"moveToRangeEnd"
            | b"bookmarkStart"
            | b"bookmarkEnd"
            | b"commentRangeStart"
            | b"commentRangeEnd"
            | b"permStart"
            | b"permEnd"
            | b"customXml"
            | b"customXmlInsRangeStart"
            | b"customXmlInsRangeEnd"
            | b"customXmlDelRangeStart"
            | b"customXmlDelRangeEnd"
            | b"customXmlMoveFromRangeStart"
            | b"customXmlMoveFromRangeEnd"
            | b"customXmlMoveToRangeStart"
            | b"customXmlMoveToRangeEnd"
            | b"rPrChange"
            | b"pPrChange"
            | b"sectPrChange"
            | b"tblPrChange"
            | b"tblPrExChange"
            | b"tblGridChange"
            | b"trPrChange"
            | b"tcPrChange"
            | b"tbl"
            | b"tr"
            | b"tc"
            | b"tblPr"
            | b"trPr"
            | b"tcPr"
            | b"hyperlink"
            | b"fldSimple"
            | b"fldChar"
            | b"instrText"
            | b"delInstrText"
            | b"sdt"
            | b"sdtContent"
            | b"commentReference"
            | b"footnoteReference"
            | b"endnoteReference"
            | b"drawing"
            | b"pict"
            | b"object"
    )
}

fn is_word_local(
    namespace: &ResolveResult<'_>,
    name: quick_xml::name::QName<'_>,
    local: &[u8],
) -> bool {
    is_wordprocessing_namespace(namespace) && name.local_name().as_ref() == local
}

fn raw_namespace_declarations(open_tag: &[u8]) -> Result<Vec<Vec<u8>>, Refusal> {
    let attributes = raw_tag_attributes(open_tag)?;
    let mut declarations = Vec::new();
    for attribute in attributes.into_iter().filter(|attribute| {
        attribute.name.as_slice() == b"xmlns" || attribute.name.starts_with(b"xmlns:")
    }) {
        declarations
            .try_reserve(1)
            .map_err(|_error| Refusal::RevisionDependency)?;
        declarations.push(attribute.raw);
    }
    Ok(declarations)
}

fn raw_revision_scope_attributes(
    open_tag: &[u8],
    element: &quick_xml::events::BytesStart<'_>,
    resolver: &quick_xml::name::NamespaceResolver,
    decoder: quick_xml::encoding::Decoder,
) -> Result<Vec<Vec<u8>>, Refusal> {
    let raw_attributes = raw_tag_attributes(open_tag)?;
    let mut raw_attributes = raw_attributes.into_iter();
    let mut scope = Vec::new();
    scope
        .try_reserve_exact(element.attributes().size_hint().0)
        .map_err(|_error| Refusal::RevisionDependency)?;
    let mut seen_id = false;
    let mut seen_author = false;
    let mut seen_date = false;
    let mut seen_date_utc = false;
    let mut seen_user_id = false;
    let mut seen_scope_attributes: Vec<(RevisionNamespaceKind, Vec<u8>)> = Vec::new();
    let mut attribute_count = 0usize;
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|_error| Refusal::RevisionDependency)?;
        attribute_count = attribute_count
            .checked_add(1)
            .ok_or(Refusal::RevisionDependency)?;
        if attribute_count > MAX_REVISION_ATTRIBUTES {
            return Err(Refusal::RevisionDependency);
        }
        let raw = raw_attributes.next().ok_or(Refusal::RevisionDependency)?;
        if raw.name.as_slice() != attribute.key.as_ref() {
            return Err(Refusal::RevisionDependency);
        }
        if raw.name.as_slice() == b"xmlns" || raw.name.starts_with(b"xmlns:") {
            continue;
        }
        let (namespace, local_name) = resolver.resolve_attribute(attribute.key);
        let local_name = local_name.as_ref();
        let classification = match &namespace {
            ResolveResult::Bound(Namespace(uri)) if is_wordprocessing_namespace(&namespace) => {
                match local_name {
                    b"id" => RevisionAttributeClass::Metadata(RevisionMetadataAttribute::Id),
                    b"author" => {
                        RevisionAttributeClass::Metadata(RevisionMetadataAttribute::Author)
                    },
                    b"date" => RevisionAttributeClass::Metadata(RevisionMetadataAttribute::Date),
                    b"userId" => {
                        RevisionAttributeClass::Metadata(RevisionMetadataAttribute::UserId)
                    },
                    _ => return Err(Refusal::RevisionDependency),
                }
            },
            ResolveResult::Bound(Namespace(uri))
                if *uri == WORD_2023_DATE_UTC_NAMESPACE && local_name == b"dateUtc" =>
            {
                RevisionAttributeClass::Metadata(RevisionMetadataAttribute::DateUtc)
            },
            ResolveResult::Bound(Namespace(uri))
                if *uri == MARKUP_COMPATIBILITY_NAMESPACE
                    && is_markup_compatibility_scope_attribute(local_name) =>
            {
                RevisionAttributeClass::Scope(RevisionNamespaceKind::MarkupCompatibility)
            },
            ResolveResult::Bound(Namespace(uri))
                if *uri == XML_NAMESPACE && is_xml_scope_attribute(local_name) =>
            {
                RevisionAttributeClass::Scope(RevisionNamespaceKind::Xml)
            },
            ResolveResult::Bound(_) | ResolveResult::Unbound | ResolveResult::Unknown(_) => {
                return Err(Refusal::RevisionDependency);
            },
        };
        if let RevisionAttributeClass::Metadata(metadata) = classification {
            let value = attribute
                .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
                .map_err(|_error| Refusal::RevisionDependency)?;
            validate_revision_metadata(metadata, &value)?;
            match metadata {
                RevisionMetadataAttribute::Id => {
                    if seen_id {
                        return Err(Refusal::RevisionDependency);
                    }
                    seen_id = true;
                },
                RevisionMetadataAttribute::Author => {
                    if seen_author {
                        return Err(Refusal::RevisionDependency);
                    }
                    seen_author = true;
                },
                RevisionMetadataAttribute::Date => {
                    if seen_date {
                        return Err(Refusal::RevisionDependency);
                    }
                    seen_date = true;
                },
                RevisionMetadataAttribute::DateUtc => {
                    if seen_date_utc {
                        return Err(Refusal::RevisionDependency);
                    }
                    seen_date_utc = true;
                },
                RevisionMetadataAttribute::UserId => {
                    if seen_user_id {
                        return Err(Refusal::RevisionDependency);
                    }
                    seen_user_id = true;
                },
            }
        } else {
            let RevisionAttributeClass::Scope(namespace_kind) = classification else {
                return Err(Refusal::RevisionDependency);
            };
            if seen_scope_attributes.iter().any(|(current, local)| {
                *current == namespace_kind && local.as_slice() == local_name
            }) {
                return Err(Refusal::RevisionDependency);
            }
            seen_scope_attributes
                .try_reserve(1)
                .map_err(|_error| Refusal::RevisionDependency)?;
            seen_scope_attributes.push((namespace_kind, local_name.to_vec()));
            scope.push(raw.raw);
        }
    }
    if raw_attributes.next().is_some() || !seen_id || !seen_author {
        return Err(Refusal::RevisionDependency);
    }
    Ok(scope)
}

#[derive(Clone, Copy)]
enum RevisionMetadataAttribute {
    Id,
    Author,
    Date,
    DateUtc,
    UserId,
}

enum RevisionAttributeClass {
    Metadata(RevisionMetadataAttribute),
    Scope(RevisionNamespaceKind),
}

fn validate_revision_metadata(
    attribute: RevisionMetadataAttribute,
    value: &str,
) -> Result<(), Refusal> {
    if value.len() > MAX_REVISION_METADATA_VALUE_BYTES
        || !value.chars().all(is_legal_revision_xml_character)
    {
        return Err(Refusal::RevisionDependency);
    }
    match attribute {
        RevisionMetadataAttribute::Id => {
            let normalized = collapse_revision_metadata_whitespace(value);
            let digits = normalized
                .strip_prefix('+')
                .or_else(|| normalized.strip_prefix('-'))
                .unwrap_or(&normalized);
            if digits.is_empty() || !digits.bytes().all(|byte| byte.is_ascii_digit()) {
                return Err(Refusal::RevisionDependency);
            }
        },
        RevisionMetadataAttribute::Author => {},
        RevisionMetadataAttribute::Date => {
            DateTime::new(value.to_owned()).map_err(|_error| Refusal::RevisionDependency)?;
        },
        RevisionMetadataAttribute::DateUtc => {
            if !value.is_ascii() || value.len() > 128 {
                return Err(Refusal::RevisionDependency);
            }
            DateTime::new(value.to_owned()).map_err(|_error| Refusal::RevisionDependency)?;
            let normalized = collapse_revision_metadata_whitespace(value);
            if !normalized.ends_with('Z')
                && !normalized.ends_with("+00:00")
                && !normalized.ends_with("-00:00")
            {
                return Err(Refusal::RevisionDependency);
            }
        },
        RevisionMetadataAttribute::UserId => {
            if value.trim().is_empty() {
                return Err(Refusal::RevisionDependency);
            }
        },
    }
    Ok(())
}

fn collapse_revision_metadata_whitespace(value: &str) -> String {
    let mut output = String::with_capacity(value.len());
    let mut pending_space = false;
    for character in value.chars() {
        if matches!(character, '\u{9}' | '\u{A}' | '\u{D}' | ' ') {
            pending_space = true;
        } else {
            if pending_space && !output.is_empty() {
                output.push(' ');
            }
            output.push(character);
            pending_space = false;
        }
    }
    output
}

fn is_markup_compatibility_scope_attribute(local_name: &[u8]) -> bool {
    matches!(
        local_name,
        b"Ignorable" | b"ProcessContent" | b"MustUnderstand"
    )
}

fn is_xml_scope_attribute(local_name: &[u8]) -> bool {
    matches!(local_name, b"space" | b"lang")
}

struct RawTagAttribute {
    name: Vec<u8>,
    value: Vec<u8>,
    raw: Vec<u8>,
}

#[derive(Clone, Copy, PartialEq, Eq)]
enum RevisionNamespaceKind {
    Word,
    MarkupCompatibility,
    Xml,
    DateUtc,
    Other,
}

fn raw_attribute_fragments(fragments: &[Vec<u8>]) -> Result<Vec<RawTagAttribute>, Refusal> {
    let mut attributes = Vec::new();
    attributes
        .try_reserve_exact(fragments.len())
        .map_err(|_error| Refusal::RevisionDependency)?;
    for fragment in fragments {
        let mut tag = Vec::new();
        let capacity = b"<x "
            .len()
            .checked_add(fragment.len())
            .and_then(|size| size.checked_add(1))
            .ok_or(Refusal::RevisionDependency)?;
        tag.try_reserve_exact(capacity)
            .map_err(|_error| Refusal::RevisionDependency)?;
        tag.extend_from_slice(b"<x ");
        tag.extend_from_slice(fragment);
        tag.push(b'>');
        let mut parsed = raw_tag_attributes(&tag)?;
        if parsed.len() != 1 {
            return Err(Refusal::RevisionDependency);
        }
        attributes.push(parsed.pop().ok_or(Refusal::RevisionDependency)?);
    }
    Ok(attributes)
}

fn raw_namespace_bindings(
    declarations: &[RawTagAttribute],
) -> Result<Vec<(Option<Vec<u8>>, Vec<u8>)>, Refusal> {
    let mut bindings: Vec<(Option<Vec<u8>>, Vec<u8>)> = Vec::new();
    bindings
        .try_reserve_exact(declarations.len())
        .map_err(|_error| Refusal::RevisionDependency)?;
    for declaration in declarations {
        let prefix = raw_namespace_declaration_prefix(&declaration.name)?;
        if bindings
            .iter()
            .any(|(current, _)| current.as_deref() == prefix)
        {
            return Err(Refusal::RevisionDependency);
        }
        bindings.push((prefix.map(ToOwned::to_owned), declaration.value.clone()));
    }
    Ok(bindings)
}

fn is_namespace_declaration_name(name: &[u8]) -> bool {
    name == b"xmlns" || name.starts_with(b"xmlns:")
}

fn raw_namespace_declaration_prefix(name: &[u8]) -> Result<Option<&[u8]>, Refusal> {
    if name == b"xmlns" {
        return Ok(None);
    }
    let prefix = name
        .strip_prefix(b"xmlns:")
        .filter(|prefix| !prefix.is_empty() && !prefix.contains(&b':'))
        .ok_or(Refusal::RevisionDependency)?;
    Ok(Some(prefix))
}

fn raw_attribute_local_name(name: &[u8]) -> Result<&[u8], Refusal> {
    let Some(index) = name.iter().position(|byte| *byte == b':') else {
        if name.is_empty() {
            return Err(Refusal::RevisionDependency);
        }
        return Ok(name);
    };
    let local = name
        .get(index + 1..)
        .filter(|local| !local.is_empty() && !local.contains(&b':'))
        .ok_or(Refusal::RevisionDependency)?;
    Ok(local)
}

fn raw_attribute_prefix(name: &[u8]) -> Result<Option<&[u8]>, Refusal> {
    let Some(index) = name.iter().position(|byte| *byte == b':') else {
        return Ok(None);
    };
    let prefix = name
        .get(..index)
        .filter(|prefix| !prefix.is_empty())
        .ok_or(Refusal::RevisionDependency)?;
    Ok(Some(prefix))
}

fn raw_attribute_namespace_kind(
    name: &[u8],
    child_declarations: &[&RawTagAttribute],
    root_bindings: &[(Option<Vec<u8>>, Vec<u8>)],
    inherited_namespaces: &[(Option<Vec<u8>>, Vec<u8>)],
) -> Result<RevisionNamespaceKind, Refusal> {
    let Some(prefix) = raw_attribute_prefix(name)? else {
        raw_attribute_local_name(name)?;
        return Ok(RevisionNamespaceKind::Other);
    };
    if prefix == b"xml" {
        return Ok(RevisionNamespaceKind::Xml);
    }
    if prefix == b"xmlns" {
        return Err(Refusal::RevisionDependency);
    }
    let namespace = child_declarations
        .iter()
        .find_map(|attribute| {
            raw_namespace_declaration_prefix(&attribute.name)
                .ok()
                .flatten()
                .filter(|current| *current == prefix)
                .map(|_| attribute.value.as_slice())
        })
        .or_else(|| {
            root_bindings.iter().find_map(|(current, namespace)| {
                current
                    .as_deref()
                    .filter(|current| *current == prefix)
                    .map(|_| namespace.as_slice())
            })
        })
        .or_else(|| {
            inherited_namespaces
                .iter()
                .find_map(|(current, namespace)| {
                    current
                        .as_deref()
                        .filter(|current| *current == prefix)
                        .map(|_| namespace.as_slice())
                })
        })
        .ok_or(Refusal::RevisionDependency)?;
    if namespace.is_empty() {
        return Err(Refusal::RevisionDependency);
    }
    Ok(match namespace {
        value
            if value == WORDPROCESSINGML_NAMESPACE
                || value == STRICT_WORDPROCESSINGML_NAMESPACE =>
        {
            RevisionNamespaceKind::Word
        },
        value if value == MARKUP_COMPATIBILITY_NAMESPACE => {
            RevisionNamespaceKind::MarkupCompatibility
        },
        value if value == XML_NAMESPACE => RevisionNamespaceKind::Xml,
        value if value == WORD_2023_DATE_UTC_NAMESPACE => RevisionNamespaceKind::DateUtc,
        _ => RevisionNamespaceKind::Other,
    })
}

fn raw_tag_name_end(tag: &[u8]) -> Result<usize, Refusal> {
    if tag.first() != Some(&b'<') {
        return Err(Refusal::RevisionDependency);
    }
    let mut end = 1usize;
    if tag.get(end) == Some(&b'/') {
        end += 1;
    }
    let start = end;
    while end < tag.len() && !matches!(tag[end], b'>' | b'/' | b' ' | b'\t' | b'\r' | b'\n') {
        end += 1;
    }
    if end == start {
        return Err(Refusal::RevisionDependency);
    }
    Ok(end)
}

fn raw_tag_attributes(tag: &[u8]) -> Result<Vec<RawTagAttribute>, Refusal> {
    if tag.len() > MAX_REVISION_TAG_ATTRIBUTE_BYTES {
        return Err(Refusal::RevisionDependency);
    }
    let mut name_end = 1usize;
    while name_end < tag.len()
        && !matches!(
            tag[name_end],
            b':' | b'>' | b'/' | b'=' | b' ' | b'\t' | b'\r' | b'\n'
        )
    {
        name_end += 1;
    }
    if name_end < tag.len() && tag[name_end] == b':' {
        name_end += 1;
        while name_end < tag.len()
            && !matches!(
                tag[name_end],
                b'>' | b'/' | b'=' | b' ' | b'\t' | b'\r' | b'\n'
            )
        {
            name_end += 1;
        }
    }
    let mut cursor = name_end;
    let mut attributes = Vec::new();
    let mut attribute_count = 0usize;
    let mut copied_bytes = 0usize;
    while cursor < tag.len() {
        while cursor < tag.len() && tag[cursor].is_ascii_whitespace() {
            cursor += 1;
        }
        if cursor >= tag.len() || matches!(tag[cursor], b'>' | b'/') {
            break;
        }
        let start = cursor;
        while cursor < tag.len()
            && !matches!(
                tag[cursor],
                b'=' | b'>' | b'/' | b' ' | b'\t' | b'\r' | b'\n'
            )
        {
            cursor += 1;
        }
        let name = tag.get(start..cursor).ok_or(Refusal::RevisionDependency)?;
        if name.is_empty() || name.len() > MAX_REVISION_METADATA_VALUE_BYTES {
            return Err(Refusal::RevisionDependency);
        }
        while cursor < tag.len() && tag[cursor].is_ascii_whitespace() {
            cursor += 1;
        }
        if tag.get(cursor) != Some(&b'=') {
            return Err(Refusal::RevisionDependency);
        }
        cursor += 1;
        while cursor < tag.len() && tag[cursor].is_ascii_whitespace() {
            cursor += 1;
        }
        let quote = *tag.get(cursor).ok_or(Refusal::RevisionDependency)?;
        if !matches!(quote, b'"' | b'\'') {
            return Err(Refusal::RevisionDependency);
        }
        cursor += 1;
        let value_start = cursor;
        while cursor < tag.len() && tag[cursor] != quote {
            cursor += 1;
        }
        let value = tag
            .get(value_start..cursor)
            .ok_or(Refusal::RevisionDependency)?;
        if value.len() > MAX_REVISION_METADATA_VALUE_BYTES {
            return Err(Refusal::RevisionDependency);
        }
        if cursor >= tag.len() {
            return Err(Refusal::RevisionDependency);
        }
        cursor += 1;
        attribute_count = attribute_count
            .checked_add(1)
            .ok_or(Refusal::RevisionDependency)?;
        if attribute_count > MAX_REVISION_ATTRIBUTES {
            return Err(Refusal::RevisionDependency);
        }
        let raw = tag.get(start..cursor).ok_or(Refusal::RevisionDependency)?;
        copied_bytes = copied_bytes
            .checked_add(name.len())
            .and_then(|size| size.checked_add(value.len()))
            .and_then(|size| size.checked_add(raw.len()))
            .ok_or(Refusal::RevisionDependency)?;
        if copied_bytes > MAX_REVISION_TAG_ATTRIBUTE_BYTES {
            return Err(Refusal::RevisionDependency);
        }
        let mut name_owned = Vec::new();
        name_owned
            .try_reserve_exact(name.len())
            .map_err(|_error| Refusal::RevisionDependency)?;
        name_owned.extend_from_slice(name);
        let mut value_owned = Vec::new();
        value_owned
            .try_reserve_exact(value.len())
            .map_err(|_error| Refusal::RevisionDependency)?;
        value_owned.extend_from_slice(value);
        let mut raw_owned = Vec::new();
        raw_owned
            .try_reserve_exact(raw.len())
            .map_err(|_error| Refusal::RevisionDependency)?;
        raw_owned.extend_from_slice(raw);
        attributes
            .try_reserve_exact(1)
            .map_err(|_error| Refusal::RevisionDependency)?;
        attributes.push(RawTagAttribute {
            name: name_owned,
            value: value_owned,
            raw: raw_owned,
        });
    }
    Ok(attributes)
}

fn revision_tag_name_replacement(
    xml: &[u8],
    event_start: usize,
    end_tag: bool,
) -> Result<(usize, usize, Vec<u8>), Refusal> {
    let name_start = event_start
        .checked_add(if end_tag { 2 } else { 1 })
        .ok_or(Refusal::RevisionDependency)?;
    revision_tag_name_replacement_at(xml, name_start)
}

fn revision_tag_name_replacement_at(
    xml: &[u8],
    name_start: usize,
) -> Result<(usize, usize, Vec<u8>), Refusal> {
    let name_end = xml
        .get(name_start..)
        .and_then(|suffix| {
            suffix
                .iter()
                .position(|byte| matches!(*byte, b'>' | b'/' | b' ' | b'\t' | b'\r' | b'\n'))
        })
        .and_then(|offset| offset.checked_add(name_start))
        .ok_or(Refusal::RevisionDependency)?;
    let local_start = xml
        .get(name_start..name_end)
        .and_then(|name| name.iter().rposition(|byte| *byte == b':'))
        .map_or(name_start, |offset| name_start + offset + 1);
    Ok((local_start, name_end, b"t".to_vec()))
}

fn scan_text_owner(xml: &[u8], root_name: &[u8]) -> Result<TextOwner, Refusal> {
    let mut reader = NsReader::from_reader(xml);
    let mut fragment_prefix = FragmentPrefix::Unseen;
    let mut root_depth = None;
    let mut run_depth = None;
    let mut slots = Vec::new();
    let mut open_text = None::<(usize, Vec<u8>, Vec<u8>)>;
    let mut text = String::new();
    let mut depth = 0usize;
    let mut saw_root = false;

    loop {
        let event_start = usize::try_from(reader.buffer_position())
            .map_err(|_conversion_error| Refusal::ComplexContent)?;
        let raw_event = reader
            .read_event()
            .map_err(|_xml_error| Refusal::ComplexContent)?
            .into_owned();
        // `raw_event` owns its bytes; borrow the resolver instead of copying
        // the complete namespace environment for every event.
        let resolver = reader.resolver();
        let (namespace, event) = resolver.resolve_event(raw_event);
        let event_end = usize::try_from(reader.buffer_position())
            .map_err(|_conversion_error| Refusal::ComplexContent)?;

        match event {
            Event::Start(element) => {
                depth = depth.checked_add(1).ok_or(Refusal::ComplexContent)?;
                if root_depth.is_none() {
                    fragment_prefix = FragmentPrefix::from_name(element.name());
                    if saw_root
                        || !is_transaction_fragment_word_name(
                            &namespace,
                            element.name(),
                            root_name,
                            &fragment_prefix,
                        )
                    {
                        return Err(Refusal::ComplexContent);
                    }
                    root_depth = Some(depth);
                    if root_name == b"r" {
                        run_depth = Some(depth);
                    }
                    saw_root = true;
                } else if run_depth.is_some_and(|run| depth == run + 1) {
                    if is_owner_text_element(
                        root_name,
                        &namespace,
                        element.name(),
                        &fragment_prefix,
                    ) {
                        let prefix = element
                            .name()
                            .prefix()
                            .map_or_else(Vec::new, |value| value.into_inner().to_vec());
                        open_text =
                            Some((event_start, prefix, element.local_name().as_ref().to_vec()));
                    } else if is_structural_run_text(&namespace, element.name(), &fragment_prefix) {
                        return Err(Refusal::ComplexRun);
                    }
                } else if root_depth.is_some_and(|root| depth == root + 1) {
                    if is_transaction_fragment_word_name(
                        &namespace,
                        element.name(),
                        b"r",
                        &fragment_prefix,
                    ) {
                        run_depth = Some(depth);
                    } else if is_fragment_word_element(&namespace, &fragment_prefix)
                        && !(root_name == b"p"
                            && is_transaction_fragment_word_name(
                                &namespace,
                                element.name(),
                                b"pPr",
                                &fragment_prefix,
                            ))
                    {
                        return Err(Refusal::ComplexContent);
                    }
                } else if open_text.is_some() {
                    return Err(Refusal::ComplexRun);
                }
            },
            Event::Empty(element) => {
                let child_depth = depth.checked_add(1).ok_or(Refusal::ComplexContent)?;
                if run_depth.is_some_and(|run| child_depth == run + 1) {
                    if is_owner_text_element(
                        root_name,
                        &namespace,
                        element.name(),
                        &fragment_prefix,
                    ) {
                        let prefix = element
                            .name()
                            .prefix()
                            .map_or_else(Vec::new, |value| value.into_inner().to_vec());
                        slots.push(TextSlot {
                            start: event_start,
                            end: event_end,
                            prefix,
                            local_name: element.local_name().as_ref().to_vec(),
                            characters: 0,
                        });
                    } else if let Some(character) =
                        structural_run_character(&namespace, element.name(), &fragment_prefix)
                    {
                        let prefix = element
                            .name()
                            .prefix()
                            .map_or_else(Vec::new, |value| value.into_inner().to_vec());
                        text.push(character);
                        slots.push(TextSlot {
                            start: event_start,
                            end: event_end,
                            prefix,
                            local_name: element.local_name().as_ref().to_vec(),
                            characters: 1,
                        });
                    } else if is_structural_run_text(&namespace, element.name(), &fragment_prefix) {
                        return Err(Refusal::ComplexRun);
                    }
                } else if root_depth.is_some_and(|root| child_depth == root + 1) {
                    if is_transaction_fragment_word_name(
                        &namespace,
                        element.name(),
                        b"r",
                        &fragment_prefix,
                    ) {
                        continue;
                    }
                    if is_fragment_word_element(&namespace, &fragment_prefix)
                        && !(root_name == b"p"
                            && is_transaction_fragment_word_name(
                                &namespace,
                                element.name(),
                                b"pPr",
                                &fragment_prefix,
                            ))
                    {
                        return Err(Refusal::ComplexContent);
                    }
                }
            },
            Event::End(element) => {
                if open_text.is_some()
                    && is_owner_text_element(
                        root_name,
                        &namespace,
                        element.name(),
                        &fragment_prefix,
                    )
                {
                    let (start, prefix, local_name) =
                        open_text.take().ok_or(Refusal::ComplexRun)?;
                    let value = decode_text_fragment(
                        xml.get(start..event_end).ok_or(Refusal::ComplexRun)?,
                    )?;
                    let characters = value.chars().count();
                    text.push_str(&value);
                    slots.push(TextSlot {
                        start,
                        end: event_end,
                        prefix,
                        local_name,
                        characters,
                    });
                }
                if run_depth == Some(depth)
                    && is_transaction_fragment_word_name(
                        &namespace,
                        element.name(),
                        b"r",
                        &fragment_prefix,
                    )
                {
                    run_depth = None;
                }
                depth = depth.checked_sub(1).ok_or(Refusal::ComplexContent)?;
            },
            Event::Text(event_text) if open_text.is_none() => {
                if !event_text.as_ref().iter().all(u8::is_ascii_whitespace) {
                    return Err(if run_depth.is_some() {
                        Refusal::ComplexRun
                    } else {
                        Refusal::ComplexContent
                    });
                }
            },
            Event::CData(_) | Event::GeneralRef(_) if open_text.is_none() => {
                return Err(if run_depth.is_some() {
                    Refusal::ComplexRun
                } else {
                    Refusal::ComplexContent
                });
            },
            Event::Eof => break,
            Event::Text(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::PI(_)
            | Event::DocType(_)
            | Event::GeneralRef(_) => {},
        }
    }
    if !saw_root || depth != 0 || open_text.is_some() {
        return Err(Refusal::ComplexContent);
    }
    if slots.is_empty() {
        return Err(Refusal::ComplexRun);
    }
    Ok(TextOwner { slots, text })
}

fn is_owner_text_element(
    root_name: &[u8],
    namespace: &ResolveResult<'_>,
    name: quick_xml::name::QName<'_>,
    fragment_prefix: &FragmentPrefix,
) -> bool {
    is_transaction_fragment_word_name(
        namespace,
        name,
        if root_name == b"del" {
            b"delText"
        } else {
            b"t"
        },
        fragment_prefix,
    )
}

fn decode_text_fragment(xml: &[u8]) -> Result<String, Refusal> {
    let mut reader = Reader::from_reader(xml);
    let mut value = String::new();
    let mut depth = 0usize;
    loop {
        match reader.read_event().map_err(|_error| Refusal::ComplexRun)? {
            Event::Start(_) => {
                depth = depth.checked_add(1).ok_or(Refusal::ComplexRun)?;
                if depth > 1 {
                    return Err(Refusal::ComplexRun);
                }
            },
            Event::Empty(_) => return Err(Refusal::ComplexRun),
            Event::Text(text) if depth == 1 => {
                let decoded = text
                    .xml_content(XmlVersion::Explicit1_0)
                    .map_err(|_error| Refusal::ComplexRun)?;
                let unescaped =
                    quick_xml::escape::unescape(&decoded).map_err(|_error| Refusal::ComplexRun)?;
                value.push_str(&unescaped);
            },
            Event::CData(text) if depth == 1 => {
                let decoded = text
                    .xml_content(XmlVersion::Explicit1_0)
                    .map_err(|_error| Refusal::ComplexRun)?;
                value.push_str(&decoded);
            },
            Event::GeneralRef(reference) if depth == 1 => {
                value.push_str(
                    &litchi_ooxml_common::xml::decode_xml_reference(&reference)
                        .map_err(|_error| Refusal::ComplexRun)?,
                );
            },
            Event::End(_) => {
                depth = depth.checked_sub(1).ok_or(Refusal::ComplexRun)?;
            },
            Event::Eof => {
                if depth == 0 {
                    break;
                }
                return Err(Refusal::ComplexRun);
            },
            Event::Text(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::PI(_)
            | Event::DocType(_)
            | Event::GeneralRef(_) => {},
        }
    }
    Ok(value)
}

fn rewrite_text_owner(xml: &[u8], owner: &TextOwner, text: &str) -> TransactionResult<Vec<u8>> {
    let _ = preflight_text_owner_rewrite(xml, owner, text)?;
    let total_characters = text.chars().count();
    let mut characters = text.char_indices();
    let mut character_cursor = 0usize;
    let mut byte_cursor = 0usize;
    let mut replacements = Vec::new();
    replacements
        .try_reserve_exact(owner.slots.len())
        .map_err(|source| {
            TransactionError::Document(crate::Error::Allocation {
                resource: "paragraph replacements",
                source,
            })
        })?;
    for (index, slot) in owner.slots.iter().enumerate() {
        let remaining = total_characters
            .checked_sub(character_cursor)
            .ok_or_else(|| {
                TransactionError::Document(crate::Error::InvalidFormat(
                    "paragraph text slot cursor exceeded authored text".into(),
                ))
            })?;
        let count = if index + 1 == owner.slots.len() {
            remaining
        } else {
            slot.characters.min(remaining)
        };
        let value_start = byte_cursor;
        for _ in 0..count {
            let (_, character) = characters.next().ok_or_else(|| {
                TransactionError::Document(crate::Error::InvalidFormat(
                    "paragraph text slot is outside authored text".into(),
                ))
            })?;
            byte_cursor =
                byte_cursor
                    .checked_add(character.len_utf8())
                    .ok_or(TransactionError::Limit {
                        resource: "replacement text bytes",
                        max: MAX_REPLACEMENT_TEXT_BYTES,
                        actual: usize::MAX,
                    })?;
        }
        character_cursor = character_cursor
            .checked_add(count)
            .ok_or(TransactionError::Limit {
                resource: "replacement text characters",
                max: MAX_REPLACEMENT_TEXT_BYTES,
                actual: usize::MAX,
            })?;
        let value = &text[value_start..byte_cursor];
        let fragment = try_run_content_fragment(&slot.prefix, &slot.local_name, value)?;
        replacements.push((slot.start, slot.end, fragment));
    }
    replace_ranges(xml, &replacements)
}

fn preflight_text_owner_rewrite(
    xml: &[u8],
    owner: &TextOwner,
    text: &str,
) -> TransactionResult<(usize, usize)> {
    let total_characters = text.chars().count();
    let mut characters = text.char_indices();
    let mut character_cursor = 0usize;
    let mut byte_cursor = 0usize;
    let mut output_len = xml.len();
    let mut fragment_bytes = 0usize;
    for (index, slot) in owner.slots.iter().enumerate() {
        let remaining = total_characters
            .checked_sub(character_cursor)
            .ok_or_else(|| {
                TransactionError::Document(crate::Error::InvalidFormat(
                    "paragraph text slot cursor exceeded authored text".into(),
                ))
            })?;
        let count = if index + 1 == owner.slots.len() {
            remaining
        } else {
            slot.characters.min(remaining)
        };
        let value_start = byte_cursor;
        for _ in 0..count {
            let (_, character) = characters.next().ok_or_else(|| {
                TransactionError::Document(crate::Error::InvalidFormat(
                    "paragraph text slot is outside authored text".into(),
                ))
            })?;
            byte_cursor =
                byte_cursor
                    .checked_add(character.len_utf8())
                    .ok_or(TransactionError::Limit {
                        resource: "replacement text bytes",
                        max: MAX_REPLACEMENT_TEXT_BYTES,
                        actual: usize::MAX,
                    })?;
        }
        character_cursor = character_cursor
            .checked_add(count)
            .ok_or(TransactionError::Limit {
                resource: "replacement text characters",
                max: MAX_REPLACEMENT_TEXT_BYTES,
                actual: usize::MAX,
            })?;
        let value = &text[value_start..byte_cursor];
        let replacement_len = run_content_fragment_len(&slot.prefix, &slot.local_name, value)?;
        fragment_bytes =
            fragment_bytes
                .checked_add(replacement_len)
                .ok_or(TransactionError::Limit {
                    resource: "paragraph replacement XML bytes",
                    max: MAX_DOCUMENT_XML_BYTES,
                    actual: usize::MAX,
                })?;
        let removed_len = slot.end.checked_sub(slot.start).ok_or_else(|| {
            TransactionError::Document(crate::Error::InvalidFormat(
                "paragraph text slot range is inverted".into(),
            ))
        })?;
        output_len = output_len.checked_sub(removed_len).ok_or_else(|| {
            TransactionError::Document(crate::Error::InvalidFormat(
                "paragraph text slot range exceeds document XML".into(),
            ))
        })?;
        output_len = output_len
            .checked_add(replacement_len)
            .ok_or(TransactionError::Limit {
                resource: "XML bytes",
                max: MAX_DOCUMENT_XML_BYTES,
                actual: usize::MAX,
            })?;
    }
    if output_len > MAX_DOCUMENT_XML_BYTES {
        return Err(TransactionError::Limit {
            resource: "XML bytes",
            max: MAX_DOCUMENT_XML_BYTES,
            actual: output_len,
        });
    }
    Ok((output_len, fragment_bytes))
}

fn is_transaction_fragment_word_name(
    namespace: &ResolveResult<'_>,
    name: quick_xml::name::QName<'_>,
    local_name: &[u8],
    fragment_prefix: &FragmentPrefix,
) -> bool {
    name.local_name().as_ref() == local_name && is_fragment_word_element(namespace, fragment_prefix)
}

fn is_fragment_word_element(
    namespace: &ResolveResult<'_>,
    fragment_prefix: &FragmentPrefix,
) -> bool {
    if is_wordprocessing_namespace(namespace) {
        return true;
    }
    match namespace {
        ResolveResult::Unknown(prefix) => matches!(
            fragment_prefix,
            FragmentPrefix::Prefixed(candidate) if candidate.as_slice() == prefix.as_slice()
        ),
        ResolveResult::Unbound => matches!(fragment_prefix, FragmentPrefix::Unprefixed),
        ResolveResult::Bound(_) => false,
    }
}

fn is_structural_run_text(
    namespace: &ResolveResult<'_>,
    name: quick_xml::name::QName<'_>,
    fragment_prefix: &FragmentPrefix,
) -> bool {
    [
        b"tab".as_slice(),
        b"br".as_slice(),
        b"cr".as_slice(),
        b"noBreakHyphen".as_slice(),
        b"softHyphen".as_slice(),
        b"instrText".as_slice(),
        b"delText".as_slice(),
        b"fldChar".as_slice(),
    ]
    .into_iter()
    .any(|local| is_transaction_fragment_word_name(namespace, name, local, fragment_prefix))
}

fn structural_run_character(
    namespace: &ResolveResult<'_>,
    name: quick_xml::name::QName<'_>,
    fragment_prefix: &FragmentPrefix,
) -> Option<char> {
    [
        (b"tab".as_slice(), '\t'),
        (b"br".as_slice(), '\n'),
        (b"cr".as_slice(), '\r'),
        (b"noBreakHyphen".as_slice(), '\u{2011}'),
        (b"softHyphen".as_slice(), '\u{00AD}'),
    ]
    .into_iter()
    .find_map(|(local, character)| {
        is_transaction_fragment_word_name(namespace, name, local, fragment_prefix)
            .then_some(character)
    })
}

fn select_complex_field_result(xml: &[u8], position: Position) -> Result<Range, Refusal> {
    let mut active = Vec::<(usize, Option<usize>)>::new();
    let mut field_index = 0usize;
    for run_index in 0..MAX_OPERATIONS {
        let Ok(run) = select_direct_child(
            xml,
            b"p",
            b"r",
            Position::new(run_index),
            Refusal::RunNotFound,
        ) else {
            break;
        };
        let start = usize::try_from(run.start).map_err(|_error| Refusal::ComplexContent)?;
        let end = start
            .checked_add(usize::try_from(run.length).map_err(|_error| Refusal::ComplexContent)?)
            .ok_or(Refusal::ComplexContent)?;
        let marker = complex_field_marker(xml.get(start..end).ok_or(Refusal::ComplexContent)?)?;
        match marker {
            Some(ComplexFieldMarker::Begin) => {
                active.push((field_index, None));
                field_index = field_index.checked_add(1).ok_or(Refusal::ComplexContent)?;
            },
            Some(ComplexFieldMarker::Separate) => {
                let Some((_index, result_start)) = active.last_mut() else {
                    return Err(Refusal::ComplexContent);
                };
                if result_start.replace(end).is_some() {
                    return Err(Refusal::ComplexContent);
                }
            },
            Some(ComplexFieldMarker::End) => {
                let Some((index, result_start)) = active.pop() else {
                    return Err(Refusal::ComplexContent);
                };
                if index == position.get() {
                    let resolved_start = result_start.ok_or(Refusal::ComplexContent)?;
                    return checked_range(resolved_start, start)
                        .map_err(|_error| Refusal::ComplexContent);
                }
            },
            None => {},
        }
    }
    Err(Refusal::ComplexFieldNotFound)
}

fn complex_field_marker(xml: &[u8]) -> Result<Option<ComplexFieldMarker>, Refusal> {
    let mut reader = NsReader::from_reader(xml);
    let mut fragment_prefix = FragmentPrefix::Unseen;
    loop {
        let raw = reader
            .read_event()
            .map_err(|_error| Refusal::ComplexContent)?
            .into_owned();
        let resolver = reader.resolver().clone();
        let (namespace, event) = resolver.resolve_event(raw);
        match event {
            Event::Start(element) | Event::Empty(element) => {
                if matches!(fragment_prefix, FragmentPrefix::Unseen) {
                    fragment_prefix = FragmentPrefix::from_name(element.name());
                }
                if is_transaction_fragment_word_name(
                    &namespace,
                    element.name(),
                    b"fldChar",
                    &fragment_prefix,
                ) {
                    let value = element
                        .attributes()
                        .filter_map(Result::ok)
                        .find(|attribute| attribute.key.local_name().as_ref() == b"fldCharType")
                        .ok_or(Refusal::ComplexContent)?
                        .decoded_and_normalized_value(XmlVersion::Explicit1_0, reader.decoder())
                        .map_err(|_error| Refusal::ComplexContent)?;
                    return match value.as_ref() {
                        "begin" => Ok(Some(ComplexFieldMarker::Begin)),
                        "separate" => Ok(Some(ComplexFieldMarker::Separate)),
                        "end" => Ok(Some(ComplexFieldMarker::End)),
                        _ => Err(Refusal::ComplexContent),
                    };
                }
            },
            Event::Eof => return Ok(None),
            Event::DocType(_) | Event::PI(_) => return Err(Refusal::ComplexContent),
            Event::End(_)
            | Event::Text(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::GeneralRef(_) => {},
        }
    }
}

fn rewrite_run_region(xml: &[u8], text: &str) -> Result<(String, Vec<u8>), Refusal> {
    let (wrapper, prefix_length, suffix_length, owner) = scan_run_region(xml)?;
    let replacement =
        rewrite_text_owner(&wrapper, &owner, text).map_err(|_error| Refusal::ComplexContent)?;
    let end = replacement
        .len()
        .checked_sub(suffix_length)
        .ok_or(Refusal::ComplexContent)?;
    Ok((
        owner.text,
        replacement
            .get(prefix_length..end)
            .ok_or(Refusal::ComplexContent)?
            .to_vec(),
    ))
}

fn scan_run_region(xml: &[u8]) -> Result<(Vec<u8>, usize, usize, TextOwner), Refusal> {
    let first = xml
        .iter()
        .position(|byte| *byte == b'<')
        .ok_or(Refusal::ComplexRun)?;
    let name_end = xml
        .get(first + 1..)
        .ok_or(Refusal::ComplexRun)?
        .iter()
        .position(|byte| matches!(*byte, b':' | b' ' | b'>' | b'/'))
        .ok_or(Refusal::ComplexRun)?
        .checked_add(first + 1)
        .ok_or(Refusal::ComplexRun)?;
    let prefix = if xml.get(name_end) == Some(&b':') {
        xml.get(first + 1..name_end).ok_or(Refusal::ComplexRun)?
    } else {
        &[]
    };
    let prefix_text = std::str::from_utf8(prefix).map_err(|_error| Refusal::ComplexRun)?;
    let (open, close) = if prefix_text.is_empty() {
        ("<region>".to_owned(), "</region>".to_owned())
    } else {
        (
            format!("<{prefix_text}:region>"),
            format!("</{prefix_text}:region>"),
        )
    };
    let mut wrapper = Vec::with_capacity(open.len() + xml.len() + close.len());
    wrapper.extend_from_slice(open.as_bytes());
    wrapper.extend_from_slice(xml);
    wrapper.extend_from_slice(close.as_bytes());
    let owner = scan_text_owner(&wrapper, b"region")?;
    Ok((wrapper, open.len(), close.len(), owner))
}

fn select_direct_child(
    xml: &[u8],
    root_name: &[u8],
    child_name: &[u8],
    position: Position,
    missing: Refusal,
) -> Result<Range, Refusal> {
    select_direct_child_with_context(xml, root_name, child_name, position, missing, &[])
}

fn select_direct_child_with_context(
    xml: &[u8],
    root_name: &[u8],
    child_name: &[u8],
    position: Position,
    missing: Refusal,
    inherited_namespaces: &[(Option<Vec<u8>>, Vec<u8>)],
) -> Result<Range, Refusal> {
    let mut reader = NsReader::from_reader(xml);
    for (prefix, namespace) in inherited_namespaces {
        let prefix = prefix
            .as_deref()
            .map_or(PrefixDeclaration::Default, PrefixDeclaration::Named);
        reader
            .resolver_mut()
            .add(prefix, Namespace(namespace))
            .map_err(|_error| missing)?;
    }
    let mut fragment_prefix = FragmentPrefix::Unseen;
    let mut root_depth = None;
    let mut capture = None::<(usize, usize)>;
    let mut depth = 0usize;
    let mut index = 0usize;
    let mut saw_root = false;
    loop {
        let start = usize::try_from(reader.buffer_position()).map_err(|_error| missing)?;
        let raw_event = reader.read_event().map_err(|_error| missing)?;
        let (namespace, event) = reader.resolver().resolve_event(raw_event);
        let end = usize::try_from(reader.buffer_position()).map_err(|_error| missing)?;
        match event {
            Event::Start(element) => {
                depth = depth.checked_add(1).ok_or(missing)?;
                if root_depth.is_none() {
                    fragment_prefix = FragmentPrefix::from_name(element.name());
                    if saw_root
                        || !is_transaction_fragment_word_name(
                            &namespace,
                            element.name(),
                            root_name,
                            &fragment_prefix,
                        )
                    {
                        return Err(missing);
                    }
                    saw_root = true;
                    root_depth = Some(depth);
                } else if root_depth.is_some_and(|root| depth == root + 1)
                    && is_transaction_fragment_word_name(
                        &namespace,
                        element.name(),
                        child_name,
                        &fragment_prefix,
                    )
                {
                    if index == position.get() {
                        capture = Some((start, depth));
                    }
                    index = index.checked_add(1).ok_or(missing)?;
                }
            },
            Event::Empty(element) => {
                let child_depth = depth.checked_add(1).ok_or(missing)?;
                if root_depth.is_some_and(|root| child_depth == root + 1)
                    && is_transaction_fragment_word_name(
                        &namespace,
                        element.name(),
                        child_name,
                        &fragment_prefix,
                    )
                {
                    if index == position.get() {
                        return checked_range(start, end).map_err(|_error| missing);
                    }
                    index = index.checked_add(1).ok_or(missing)?;
                }
            },
            Event::End(_) => {
                if let Some((capture_start, capture_depth)) = capture
                    && depth == capture_depth
                {
                    return checked_range(capture_start, end).map_err(|_error| missing);
                }
                depth = depth.checked_sub(1).ok_or(missing)?;
            },
            Event::Eof if depth != 0 || capture.is_some() => return Err(missing),
            Event::Eof => break,
            Event::Text(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::PI(_)
            | Event::DocType(_)
            | Event::GeneralRef(_) => {},
        }
    }
    Err(missing)
}

fn single_cell_paragraph(xml: &[u8]) -> TransactionResult<Range> {
    validate_basic_cell_children(xml).map_err(|reason| TransactionError::Refused {
        position: 0,
        reason,
    })?;
    let first = select_direct_child(xml, b"tc", b"p", Position::new(0), Refusal::ComplexContent)
        .map_err(|reason| TransactionError::Refused {
            position: 0,
            reason,
        })?;
    if select_direct_child(xml, b"tc", b"p", Position::new(1), Refusal::CellNotFound).is_ok() {
        return Err(TransactionError::Refused {
            position: 0,
            reason: Refusal::ComplexContent,
        });
    }
    Ok(first)
}

fn validate_basic_cell_children(xml: &[u8]) -> Result<(), Refusal> {
    let mut reader = NsReader::from_reader(xml);
    let mut fragment_prefix = FragmentPrefix::Unseen;
    let mut root_depth = None;
    let mut depth = 0usize;
    let mut saw_root = false;
    loop {
        let raw_event = reader
            .read_event()
            .map_err(|_error| Refusal::ComplexContent)?
            .into_owned();
        let resolver = reader.resolver().clone();
        let (namespace, event) = resolver.resolve_event(raw_event);
        match event {
            Event::Start(element) => {
                depth = depth.checked_add(1).ok_or(Refusal::ComplexContent)?;
                if root_depth.is_none() {
                    fragment_prefix = FragmentPrefix::from_name(element.name());
                    if saw_root
                        || !is_transaction_fragment_word_name(
                            &namespace,
                            element.name(),
                            b"tc",
                            &fragment_prefix,
                        )
                    {
                        return Err(Refusal::ComplexContent);
                    }
                    saw_root = true;
                    root_depth = Some(depth);
                } else if root_depth.is_some_and(|root| depth == root + 1)
                    && is_fragment_word_element(&namespace, &fragment_prefix)
                    && !is_transaction_fragment_word_name(
                        &namespace,
                        element.name(),
                        b"tcPr",
                        &fragment_prefix,
                    )
                    && !is_transaction_fragment_word_name(
                        &namespace,
                        element.name(),
                        b"p",
                        &fragment_prefix,
                    )
                {
                    return Err(Refusal::ComplexContent);
                }
            },
            Event::Empty(element) => {
                let child_depth = depth.checked_add(1).ok_or(Refusal::ComplexContent)?;
                if root_depth.is_some_and(|root| child_depth == root + 1)
                    && is_fragment_word_element(&namespace, &fragment_prefix)
                    && !is_transaction_fragment_word_name(
                        &namespace,
                        element.name(),
                        b"tcPr",
                        &fragment_prefix,
                    )
                    && !is_transaction_fragment_word_name(
                        &namespace,
                        element.name(),
                        b"p",
                        &fragment_prefix,
                    )
                {
                    return Err(Refusal::ComplexContent);
                }
            },
            Event::End(_) => {
                depth = depth.checked_sub(1).ok_or(Refusal::ComplexContent)?;
            },
            Event::Text(text)
                if root_depth.is_some_and(|root| depth == root)
                    && !text.as_ref().iter().all(u8::is_ascii_whitespace) =>
            {
                return Err(Refusal::ComplexContent);
            },
            Event::CData(_) | Event::GeneralRef(_)
                if root_depth.is_some_and(|root| depth == root) =>
            {
                return Err(Refusal::ComplexContent);
            },
            Event::Eof => break,
            Event::Text(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::PI(_)
            | Event::DocType(_)
            | Event::GeneralRef(_) => {},
        }
    }
    if saw_root && depth == 0 {
        Ok(())
    } else {
        Err(Refusal::ComplexContent)
    }
}

fn select_cell(
    snapshot: &Snapshot,
    table: Position,
    row: Position,
    cell: Position,
) -> TransactionResult<CellSelection<'_>> {
    let table_range =
        snapshot
            .tables
            .get(table.get())
            .copied()
            .ok_or(TransactionError::Refused {
                position: table.get(),
                reason: Refusal::CellNotFound,
            })?;
    let table_start = checked_start(table_range, "table")?;
    let table_end = checked_end(table_range, "table")?;
    let table_xml = checked_slice(snapshot.xml_bytes(), table_start, table_end, "table")?;
    select_cell_in_table(snapshot, table_start, table_xml, row, cell, table.get())
}

fn select_cell_in_table<'a>(
    snapshot: &'a Snapshot,
    table_start: usize,
    table_xml: &[u8],
    row: Position,
    cell: Position,
    error_position: usize,
) -> TransactionResult<CellSelection<'a>> {
    let row_range = select_direct_child(table_xml, b"tbl", b"tr", row, Refusal::CellNotFound)
        .map_err(|reason| TransactionError::Refused {
            position: error_position,
            reason,
        })?;
    let row_start = checked_relative_start(table_start, row_range)?;
    let row_end = checked_relative_end(table_start, row_range)?;
    let row_xml = checked_slice(snapshot.xml_bytes(), row_start, row_end, "table row")?;
    let cell_range = select_direct_child(row_xml, b"tr", b"tc", cell, Refusal::CellNotFound)
        .map_err(|reason| TransactionError::Refused {
            position: error_position,
            reason,
        })?;
    let cell_start = checked_relative_start(row_start, cell_range)?;
    let cell_end = checked_relative_end(row_start, cell_range)?;
    Ok(CellSelection {
        xml: checked_slice(snapshot.xml_bytes(), cell_start, cell_end, "table cell")?,
        start: cell_start,
    })
}

fn validate_path_length(length: usize, resource: &'static str) -> TransactionResult<()> {
    if length == 0 {
        return Err(TransactionError::Refused {
            position: 0,
            reason: Refusal::ComplexContent,
        });
    }
    if length > MAX_OPERATIONS {
        return Err(TransactionError::Limit {
            resource,
            max: MAX_OPERATIONS,
            actual: length,
        });
    }
    Ok(())
}

fn select_single_control_content(
    snapshot: &Snapshot,
    control_start: usize,
    control_end: usize,
    error_position: usize,
) -> TransactionResult<(usize, usize)> {
    let control_xml = checked_slice(
        snapshot.xml_bytes(),
        control_start,
        control_end,
        "content control",
    )?;
    let content = select_direct_child(
        control_xml,
        b"sdt",
        b"sdtContent",
        Position::new(0),
        Refusal::ComplexContent,
    )
    .map_err(|reason| TransactionError::Refused {
        position: error_position,
        reason,
    })?;
    if select_direct_child(
        control_xml,
        b"sdt",
        b"sdtContent",
        Position::new(1),
        Refusal::ComplexContent,
    )
    .is_ok()
    {
        return Err(TransactionError::Refused {
            position: error_position,
            reason: Refusal::ComplexContent,
        });
    }
    Ok((
        checked_relative_start(control_start, content)?,
        checked_relative_end(control_start, content)?,
    ))
}

fn select_direct_control_content(
    snapshot: &Snapshot,
    root_start: usize,
    root_end: usize,
    root_name: &[u8],
    control: Position,
    error_position: usize,
) -> TransactionResult<(usize, usize)> {
    let root_xml = checked_slice(snapshot.xml_bytes(), root_start, root_end, "control owner")?;
    let control_range = select_direct_child(
        root_xml,
        root_name,
        b"sdt",
        control,
        Refusal::ContentControlNotFound,
    )
    .map_err(|reason| TransactionError::Refused {
        position: error_position,
        reason,
    })?;
    let control_start = checked_relative_start(root_start, control_range)?;
    let control_end = checked_relative_end(root_start, control_range)?;
    select_single_control_content(snapshot, control_start, control_end, error_position)
}

fn select_hyperlink_owner(
    snapshot: &Snapshot,
    owner: (usize, usize),
    root_name: &[u8],
    hyperlink: Position,
    error_position: usize,
) -> TransactionResult<(usize, usize)> {
    let owner_xml = checked_slice(snapshot.xml_bytes(), owner.0, owner.1, "composite owner")?;
    let hyperlink_range = select_direct_child(
        owner_xml,
        root_name,
        b"hyperlink",
        hyperlink,
        Refusal::HyperlinkNotFound,
    )
    .map_err(|reason| TransactionError::Refused {
        position: error_position,
        reason,
    })?;
    Ok((
        checked_relative_start(owner.0, hyperlink_range)?,
        checked_relative_end(owner.0, hyperlink_range)?,
    ))
}

fn selected_composite_hyperlink_text(
    snapshot: &Snapshot,
    owner: (usize, usize),
    root_name: &[u8],
    hyperlink: Position,
    error_position: usize,
) -> TransactionResult<String> {
    let (start, end) =
        select_hyperlink_owner(snapshot, owner, root_name, hyperlink, error_position)?;
    scan_text_owner(
        checked_slice(snapshot.xml_bytes(), start, end, "composite hyperlink")?,
        b"hyperlink",
    )
    .map(|scanned| scanned.text)
    .map_err(|reason| TransactionError::Refused {
        position: error_position,
        reason,
    })
}

fn select_nested_inline_control_content(
    snapshot: &Snapshot,
    paragraph: Position,
    controls: &[Position],
) -> TransactionResult<(usize, usize)> {
    validate_path_length(controls.len(), "content control path")?;
    let paragraph_range =
        snapshot
            .paragraphs
            .get(paragraph.get())
            .copied()
            .ok_or(TransactionError::OutOfBounds {
                position: paragraph.get(),
                len: snapshot.paragraph_count(),
            })?;
    let mut start = checked_start(paragraph_range, "paragraph")?;
    let mut end = checked_end(paragraph_range, "paragraph")?;
    let mut root_name = b"p".as_slice();
    for control in controls {
        (start, end) = select_direct_control_content(
            snapshot,
            start,
            end,
            root_name,
            *control,
            paragraph.get(),
        )?;
        root_name = b"sdtContent";
    }
    Ok((start, end))
}

fn selected_nested_inline_control_text(
    snapshot: &Snapshot,
    paragraph: Position,
    controls: &[Position],
) -> TransactionResult<String> {
    let (start, end) = select_nested_inline_control_content(snapshot, paragraph, controls)?;
    scan_text_owner(
        checked_slice(snapshot.xml_bytes(), start, end, "content control content")?,
        b"sdtContent",
    )
    .map(|owner| owner.text)
    .map_err(|reason| TransactionError::Refused {
        position: paragraph.get(),
        reason,
    })
}

fn selected_nested_inline_control_hyperlink_text(
    snapshot: &Snapshot,
    paragraph: Position,
    controls: &[Position],
    hyperlink: Position,
) -> TransactionResult<String> {
    let owner = select_nested_inline_control_content(snapshot, paragraph, controls)?;
    selected_composite_hyperlink_text(snapshot, owner, b"sdtContent", hyperlink, paragraph.get())
}

fn select_block_control_content(
    snapshot: &Snapshot,
    controls: &[Position],
) -> TransactionResult<(usize, usize)> {
    validate_path_length(controls.len(), "block content control path")?;
    let first = controls[0];
    let range =
        snapshot
            .block_controls
            .get(first.get())
            .copied()
            .ok_or(TransactionError::Refused {
                position: first.get(),
                reason: Refusal::ContentControlNotFound,
            })?;
    let mut start = checked_start(range, "block content control")?;
    let mut end = checked_end(range, "block content control")?;
    (start, end) = select_single_control_content(snapshot, start, end, first.get())?;
    for control in &controls[1..] {
        (start, end) = select_direct_control_content(
            snapshot,
            start,
            end,
            b"sdtContent",
            *control,
            first.get(),
        )?;
    }
    Ok((start, end))
}

fn select_block_control_paragraph(
    snapshot: &Snapshot,
    controls: &[Position],
    paragraph: Position,
) -> TransactionResult<(usize, usize)> {
    let (content_start, content_end) = select_block_control_content(snapshot, controls)?;
    let content_xml = checked_slice(
        snapshot.xml_bytes(),
        content_start,
        content_end,
        "block content control content",
    )?;
    let paragraph_range = select_direct_child(
        content_xml,
        b"sdtContent",
        b"p",
        paragraph,
        Refusal::ContentControlNotFound,
    )
    .map_err(|reason| TransactionError::Refused {
        position: controls.first().map_or(0, |position| position.get()),
        reason,
    })?;
    Ok((
        checked_relative_start(content_start, paragraph_range)?,
        checked_relative_end(content_start, paragraph_range)?,
    ))
}

fn selected_block_control_paragraph_text(
    snapshot: &Snapshot,
    controls: &[Position],
    paragraph: Position,
) -> TransactionResult<String> {
    let (start, end) = select_block_control_paragraph(snapshot, controls, paragraph)?;
    scan_text_owner(
        checked_slice(snapshot.xml_bytes(), start, end, "block control paragraph")?,
        b"p",
    )
    .map(|owner| owner.text)
    .map_err(|reason| TransactionError::Refused {
        position: controls.first().map_or(0, |position| position.get()),
        reason,
    })
}

fn selected_block_control_paragraph_hyperlink_text(
    snapshot: &Snapshot,
    controls: &[Position],
    paragraph: Position,
    hyperlink: Position,
) -> TransactionResult<String> {
    let owner = select_block_control_paragraph(snapshot, controls, paragraph)?;
    selected_composite_hyperlink_text(
        snapshot,
        owner,
        b"p",
        hyperlink,
        controls.first().map_or(0, |position| position.get()),
    )
}

fn select_nested_cell<'a>(
    snapshot: &'a Snapshot,
    path: &[TableCellAddress],
) -> TransactionResult<CellSelection<'a>> {
    validate_path_length(path.len(), "nested table path")?;
    let first = path[0];
    let table_range =
        snapshot
            .tables
            .get(first.table.get())
            .copied()
            .ok_or(TransactionError::Refused {
                position: first.table.get(),
                reason: Refusal::CellNotFound,
            })?;
    let mut table_start = checked_start(table_range, "table")?;
    let table_end = checked_end(table_range, "table")?;
    let mut table_xml = checked_slice(snapshot.xml_bytes(), table_start, table_end, "table")?;
    let mut selection = select_cell_in_table(
        snapshot,
        table_start,
        table_xml,
        first.row,
        first.cell,
        first.table.get(),
    )?;
    for address in &path[1..] {
        let nested_range = select_direct_child(
            selection.xml,
            b"tc",
            b"tbl",
            address.table,
            Refusal::CellNotFound,
        )
        .map_err(|reason| TransactionError::Refused {
            position: first.table.get(),
            reason,
        })?;
        table_start = checked_relative_start(selection.start, nested_range)?;
        let nested_end = checked_relative_end(selection.start, nested_range)?;
        table_xml = checked_slice(
            snapshot.xml_bytes(),
            table_start,
            nested_end,
            "nested table",
        )?;
        selection = select_cell_in_table(
            snapshot,
            table_start,
            table_xml,
            address.row,
            address.cell,
            first.table.get(),
        )?;
    }
    Ok(selection)
}

fn select_nested_cell_paragraph(
    snapshot: &Snapshot,
    path: &[TableCellAddress],
    paragraph: Position,
) -> TransactionResult<(usize, usize)> {
    let selection = select_nested_cell(snapshot, path)?;
    let paragraph_range =
        select_direct_child(selection.xml, b"tc", b"p", paragraph, Refusal::CellNotFound).map_err(
            |reason| TransactionError::Refused {
                position: path.first().map_or(0, |address| address.table.get()),
                reason,
            },
        )?;
    Ok((
        checked_relative_start(selection.start, paragraph_range)?,
        checked_relative_end(selection.start, paragraph_range)?,
    ))
}

fn selected_nested_cell_paragraph_text(
    snapshot: &Snapshot,
    path: &[TableCellAddress],
    paragraph: Position,
) -> TransactionResult<String> {
    let (start, end) = select_nested_cell_paragraph(snapshot, path, paragraph)?;
    scan_text_owner(
        checked_slice(snapshot.xml_bytes(), start, end, "nested table paragraph")?,
        b"p",
    )
    .map(|owner| owner.text)
    .map_err(|reason| TransactionError::Refused {
        position: path.first().map_or(0, |address| address.table.get()),
        reason,
    })
}

fn selected_nested_cell_paragraph_hyperlink_text(
    snapshot: &Snapshot,
    path: &[TableCellAddress],
    paragraph: Position,
    hyperlink: Position,
) -> TransactionResult<String> {
    let owner = select_nested_cell_paragraph(snapshot, path, paragraph)?;
    selected_composite_hyperlink_text(
        snapshot,
        owner,
        b"p",
        hyperlink,
        path.first().map_or(0, |address| address.table.get()),
    )
}

fn selected_hyperlink_text(
    snapshot: &Snapshot,
    paragraph: Position,
    hyperlink: Position,
) -> TransactionResult<String> {
    let paragraph_range =
        snapshot
            .paragraphs
            .get(paragraph.get())
            .copied()
            .ok_or(TransactionError::OutOfBounds {
                position: paragraph.get(),
                len: snapshot.paragraph_count(),
            })?;
    let start = checked_start(paragraph_range, "paragraph")?;
    let end = checked_end(paragraph_range, "paragraph")?;
    let paragraph_xml = checked_slice(snapshot.xml_bytes(), start, end, "paragraph")?;
    let range = select_direct_child(
        paragraph_xml,
        b"p",
        b"hyperlink",
        hyperlink,
        Refusal::HyperlinkNotFound,
    )
    .map_err(|reason| TransactionError::Refused {
        position: paragraph.get(),
        reason,
    })?;
    let child_start = checked_relative_start(start, range)?;
    let child_end = checked_relative_end(start, range)?;
    scan_text_owner(
        checked_slice(snapshot.xml_bytes(), child_start, child_end, "hyperlink")?,
        b"hyperlink",
    )
    .map(|scanned| scanned.text)
    .map_err(|reason| TransactionError::Refused {
        position: paragraph.get(),
        reason,
    })
}

fn selected_direct_paragraph_owner_text(
    snapshot: &Snapshot,
    paragraph: Position,
    owner: Position,
    child_name: &[u8],
    missing: Refusal,
) -> TransactionResult<String> {
    let paragraph_range =
        snapshot
            .paragraphs
            .get(paragraph.get())
            .copied()
            .ok_or(TransactionError::OutOfBounds {
                position: paragraph.get(),
                len: snapshot.paragraph_count(),
            })?;
    let paragraph_start = checked_start(paragraph_range, "paragraph")?;
    let paragraph_end = checked_end(paragraph_range, "paragraph")?;
    let paragraph_xml = checked_slice(
        snapshot.xml_bytes(),
        paragraph_start,
        paragraph_end,
        "paragraph",
    )?;
    let owner_range = select_direct_child(paragraph_xml, b"p", child_name, owner, missing)
        .map_err(|reason| TransactionError::Refused {
            position: paragraph.get(),
            reason,
        })?;
    let start = checked_relative_start(paragraph_start, owner_range)?;
    let end = checked_relative_end(paragraph_start, owner_range)?;
    scan_text_owner(
        checked_slice(snapshot.xml_bytes(), start, end, "paragraph text owner")?,
        child_name,
    )
    .map(|scanned| scanned.text)
    .map_err(|reason| TransactionError::Refused {
        position: paragraph.get(),
        reason,
    })
}

fn selected_complex_field_result_text(
    snapshot: &Snapshot,
    paragraph: Position,
    field: Position,
) -> TransactionResult<String> {
    let paragraph_range =
        snapshot
            .paragraphs
            .get(paragraph.get())
            .copied()
            .ok_or(TransactionError::OutOfBounds {
                position: paragraph.get(),
                len: snapshot.paragraph_count(),
            })?;
    let paragraph_start = checked_start(paragraph_range, "paragraph")?;
    let paragraph_end = checked_end(paragraph_range, "paragraph")?;
    let paragraph_xml = checked_slice(
        snapshot.xml_bytes(),
        paragraph_start,
        paragraph_end,
        "paragraph",
    )?;
    let result_range = select_complex_field_result(paragraph_xml, field).map_err(|reason| {
        TransactionError::Refused {
            position: paragraph.get(),
            reason,
        }
    })?;
    let start = checked_relative_start(paragraph_start, result_range)?;
    let end = checked_relative_end(paragraph_start, result_range)?;
    scan_run_region(checked_slice(
        snapshot.xml_bytes(),
        start,
        end,
        "complex field result",
    )?)
    .map(|(_wrapper, _prefix_length, _suffix_length, owner)| owner.text)
    .map_err(|reason| TransactionError::Refused {
        position: paragraph.get(),
        reason,
    })
}

fn select_content_control_content(
    snapshot: &Snapshot,
    paragraph: Position,
    control: Position,
) -> TransactionResult<(usize, usize)> {
    let paragraph_range =
        snapshot
            .paragraphs
            .get(paragraph.get())
            .copied()
            .ok_or(TransactionError::OutOfBounds {
                position: paragraph.get(),
                len: snapshot.paragraph_count(),
            })?;
    let paragraph_start = checked_start(paragraph_range, "paragraph")?;
    let paragraph_end = checked_end(paragraph_range, "paragraph")?;
    let paragraph_xml = checked_slice(
        snapshot.xml_bytes(),
        paragraph_start,
        paragraph_end,
        "paragraph",
    )?;
    let control_range = select_direct_child(
        paragraph_xml,
        b"p",
        b"sdt",
        control,
        Refusal::ContentControlNotFound,
    )
    .map_err(|reason| TransactionError::Refused {
        position: paragraph.get(),
        reason,
    })?;
    let control_start = checked_relative_start(paragraph_start, control_range)?;
    let control_end = checked_relative_end(paragraph_start, control_range)?;
    let control_xml = checked_slice(
        snapshot.xml_bytes(),
        control_start,
        control_end,
        "content control",
    )?;
    let content_range = select_direct_child(
        control_xml,
        b"sdt",
        b"sdtContent",
        Position::new(0),
        Refusal::ComplexContent,
    )
    .map_err(|reason| TransactionError::Refused {
        position: paragraph.get(),
        reason,
    })?;
    if select_direct_child(
        control_xml,
        b"sdt",
        b"sdtContent",
        Position::new(1),
        Refusal::ComplexContent,
    )
    .is_ok()
    {
        return Err(TransactionError::Refused {
            position: paragraph.get(),
            reason: Refusal::ComplexContent,
        });
    }
    Ok((
        checked_relative_start(control_start, content_range)?,
        checked_relative_end(control_start, content_range)?,
    ))
}

fn selected_content_control_text(
    snapshot: &Snapshot,
    paragraph: Position,
    control: Position,
) -> TransactionResult<String> {
    let (start, end) = select_content_control_content(snapshot, paragraph, control)?;
    scan_text_owner(
        checked_slice(snapshot.xml_bytes(), start, end, "content control content")?,
        b"sdtContent",
    )
    .map(|owner| owner.text)
    .map_err(|reason| TransactionError::Refused {
        position: paragraph.get(),
        reason,
    })
}

fn selected_cell_text(
    snapshot: &Snapshot,
    table: Position,
    row: Position,
    cell: Position,
) -> TransactionResult<String> {
    let selection = select_cell(snapshot, table, row, cell)?;
    let paragraph = single_cell_paragraph(selection.xml)?;
    let start = checked_relative_start(selection.start, paragraph)?;
    let end = checked_relative_end(selection.start, paragraph)?;
    scan_text_owner(
        checked_slice(snapshot.xml_bytes(), start, end, "table cell paragraph")?,
        b"p",
    )
    .map(|owner| owner.text)
    .map_err(|reason| TransactionError::Refused {
        position: table.get(),
        reason,
    })
}

fn selected_cell_paragraph_text(
    snapshot: &Snapshot,
    table: Position,
    row: Position,
    cell: Position,
    paragraph: Position,
) -> TransactionResult<String> {
    let selection = select_cell(snapshot, table, row, cell)?;
    let paragraph_range =
        select_direct_child(selection.xml, b"tc", b"p", paragraph, Refusal::CellNotFound).map_err(
            |reason| TransactionError::Refused {
                position: table.get(),
                reason,
            },
        )?;
    let start = checked_relative_start(selection.start, paragraph_range)?;
    let end = checked_relative_end(selection.start, paragraph_range)?;
    scan_text_owner(
        checked_slice(snapshot.xml_bytes(), start, end, "table cell paragraph")?,
        b"p",
    )
    .map(|owner| owner.text)
    .map_err(|reason| TransactionError::Refused {
        position: table.get(),
        reason,
    })
}

fn checked_start(range: Range, resource: &'static str) -> TransactionResult<usize> {
    usize::try_from(range.start).map_err(|_error| {
        crate::Error::InvalidFormat(format!("{resource} offset does not fit usize")).into()
    })
}

fn checked_end(range: Range, resource: &'static str) -> TransactionResult<usize> {
    checked_start(range, resource)?
        .checked_add(usize::try_from(range.length).map_err(|_error| {
            crate::Error::InvalidFormat(format!("{resource} length does not fit usize"))
        })?)
        .ok_or_else(|| {
            crate::Error::InvalidFormat(format!("{resource} range overflows usize")).into()
        })
}

fn checked_relative_start(base: usize, range: Range) -> TransactionResult<usize> {
    base.checked_add(usize::try_from(range.start).map_err(|_error| {
        crate::Error::InvalidFormat("relative XML offset does not fit usize".into())
    })?)
    .ok_or_else(|| crate::Error::InvalidFormat("relative XML offset overflows".into()).into())
}

fn checked_relative_end(base: usize, range: Range) -> TransactionResult<usize> {
    checked_relative_start(base, range)?
        .checked_add(usize::try_from(range.length).map_err(|_error| {
            crate::Error::InvalidFormat("relative XML length does not fit usize".into())
        })?)
        .ok_or_else(|| crate::Error::InvalidFormat("relative XML range overflows".into()).into())
}

fn checked_slice<'a>(
    source: &'a [u8],
    start: usize,
    end: usize,
    resource: &'static str,
) -> TransactionResult<&'a [u8]> {
    source.get(start..end).ok_or_else(|| {
        crate::Error::InvalidFormat(format!("{resource} range is outside document XML")).into()
    })
}

fn replace_ranges(
    source: &[u8],
    replacements: &[(usize, usize, Vec<u8>)],
) -> TransactionResult<Vec<u8>> {
    let mut capacity = source.len();
    let mut previous_end = 0usize;
    for (start, end, replacement) in replacements {
        if *start < previous_end || start > end || *end > source.len() {
            return Err(crate::Error::InvalidFormat(
                "document rewrite ranges are invalid or overlap".into(),
            )
            .into());
        }
        capacity = capacity
            .checked_sub(end - start)
            .and_then(|size| size.checked_add(replacement.len()))
            .ok_or(TransactionError::Limit {
                resource: "projected XML bytes",
                max: MAX_DOCUMENT_XML_BYTES,
                actual: usize::MAX,
            })?;
        previous_end = *end;
    }
    if capacity > MAX_DOCUMENT_XML_BYTES {
        return Err(TransactionError::Limit {
            resource: "projected XML bytes",
            max: MAX_DOCUMENT_XML_BYTES,
            actual: capacity,
        });
    }
    let mut output = Vec::new();
    output
        .try_reserve_exact(capacity)
        .map_err(|allocation_error| crate::Error::Allocation {
            resource: "document transaction XML",
            source: allocation_error,
        })?;
    let mut cursor = 0usize;
    for (start, end, replacement) in replacements {
        output.extend_from_slice(&source[cursor..*start]);
        output.extend_from_slice(replacement);
        cursor = *end;
    }
    output.extend_from_slice(&source[cursor..]);
    Ok(output)
}

fn validate_authored_text(text: &str) -> Result<(), Refusal> {
    if text.contains(['\t', '\n', '\r'])
        || text.chars().any(|character| {
            !matches!(character, '\u{20}'..='\u{D7FF}' | '\u{E000}'..='\u{FFFD}' | '\u{10000}'..='\u{10FFFF}')
        })
    {
        return Err(Refusal::StructuralText);
    }
    Ok(())
}

fn validate_authored_run_content(text: &str) -> Result<(), Refusal> {
    if text.chars().any(|character| {
        !matches!(character, '\u{9}' | '\u{A}' | '\u{D}' | '\u{20}'..='\u{D7FF}' | '\u{E000}'..='\u{FFFD}' | '\u{10000}'..='\u{10FFFF}')
    }) {
        return Err(Refusal::StructuralText);
    }
    Ok(())
}

fn run_content_fragment_len(
    prefix: &[u8],
    original_local_name: &[u8],
    text: &str,
) -> TransactionResult<usize> {
    let text_local_name = if original_local_name == b"delText" {
        b"delText".as_slice()
    } else {
        b"t".as_slice()
    };
    if text.is_empty() {
        return text_element_named_len(prefix, text_local_name, "");
    }
    let mut output_len = 0usize;
    let mut plain_start = 0usize;
    for (index, character) in text.char_indices() {
        let structural = match character {
            '\t' => Some(b"tab".as_slice()),
            '\n' => Some(b"br".as_slice()),
            '\r' => Some(b"cr".as_slice()),
            '\u{2011}' => Some(b"noBreakHyphen".as_slice()),
            '\u{00AD}' => Some(b"softHyphen".as_slice()),
            _ => None,
        };
        if let Some(local_name) = structural {
            let plain = &text[plain_start..index];
            if !plain.is_empty() {
                output_len = replacement_len_add(
                    output_len,
                    text_element_named_len(prefix, text_local_name, plain)?,
                )?;
            }
            let name_len = qualified_name_len(prefix, local_name)?;
            output_len = replacement_len_add(output_len, replacement_len_add(3, name_len)?)?;
            plain_start =
                index
                    .checked_add(character.len_utf8())
                    .ok_or(TransactionError::Limit {
                        resource: "replacement text bytes",
                        max: MAX_REPLACEMENT_TEXT_BYTES,
                        actual: usize::MAX,
                    })?;
        }
    }
    let plain = &text[plain_start..];
    if !plain.is_empty() {
        output_len = replacement_len_add(
            output_len,
            text_element_named_len(prefix, text_local_name, plain)?,
        )?;
    }
    Ok(output_len)
}

fn try_run_content_fragment(
    prefix: &[u8],
    original_local_name: &[u8],
    text: &str,
) -> TransactionResult<Vec<u8>> {
    let capacity = run_content_fragment_len(prefix, original_local_name, text)?;
    let text_local_name = if original_local_name == b"delText" {
        b"delText".as_slice()
    } else {
        b"t".as_slice()
    };
    let mut output = Vec::new();
    output.try_reserve_exact(capacity).map_err(|source| {
        TransactionError::Document(crate::Error::Allocation {
            resource: "paragraph replacement XML",
            source,
        })
    })?;
    if text.is_empty() {
        append_text_element_named(&mut output, prefix, text_local_name, text);
        return Ok(output);
    }
    let mut plain_start = 0usize;
    for (index, character) in text.char_indices() {
        let structural = match character {
            '\t' => Some(b"tab".as_slice()),
            '\n' => Some(b"br".as_slice()),
            '\r' => Some(b"cr".as_slice()),
            '\u{2011}' => Some(b"noBreakHyphen".as_slice()),
            '\u{00AD}' => Some(b"softHyphen".as_slice()),
            _ => None,
        };
        if let Some(local_name) = structural {
            let plain = &text[plain_start..index];
            if !plain.is_empty() {
                append_text_element_named(&mut output, prefix, text_local_name, plain);
            }
            output.push(b'<');
            append_qualified_name(&mut output, prefix, local_name);
            output.extend_from_slice(b"/>");
            plain_start =
                index
                    .checked_add(character.len_utf8())
                    .ok_or(TransactionError::Limit {
                        resource: "replacement text bytes",
                        max: MAX_REPLACEMENT_TEXT_BYTES,
                        actual: usize::MAX,
                    })?;
        }
    }
    let plain = &text[plain_start..];
    if !plain.is_empty() {
        append_text_element_named(&mut output, prefix, text_local_name, plain);
    }
    debug_assert_eq!(output.len(), capacity);
    Ok(output)
}

fn replacement_len_add(left: usize, right: usize) -> TransactionResult<usize> {
    left.checked_add(right).ok_or(TransactionError::Limit {
        resource: "replacement XML bytes",
        max: MAX_DOCUMENT_XML_BYTES,
        actual: usize::MAX,
    })
}

fn qualified_name_len(prefix: &[u8], local_name: &[u8]) -> TransactionResult<usize> {
    replacement_len_add(
        replacement_len_add(prefix.len(), if prefix.is_empty() { 0 } else { 1 })?,
        local_name.len(),
    )
}

fn text_element_named_len(
    prefix: &[u8],
    local_name: &[u8],
    text: &str,
) -> TransactionResult<usize> {
    let name_len = qualified_name_len(prefix, local_name)?;
    if text.is_empty() {
        return replacement_len_add(3, name_len);
    }
    let preserve = text.chars().next().is_some_and(char::is_whitespace)
        || text.chars().next_back().is_some_and(char::is_whitespace);
    let attribute_len = if preserve {
        b" xml:space=\"preserve\"".len()
    } else {
        0
    };
    let mut length = replacement_len_add(5, name_len)?;
    length = replacement_len_add(length, name_len)?;
    length = replacement_len_add(length, attribute_len)?;
    replacement_len_add(length, escaped_xml_len(text)?)
}

fn escaped_xml_len(text: &str) -> TransactionResult<usize> {
    let mut length = 0usize;
    for byte in text.bytes() {
        let encoded_len = match byte {
            b'&' => 5,
            b'<' | b'>' => 4,
            b'"' | b'\'' => 6,
            _ => 1,
        };
        length = replacement_len_add(length, encoded_len)?;
    }
    Ok(length)
}

fn append_qualified_name(output: &mut Vec<u8>, prefix: &[u8], local_name: &[u8]) {
    if !prefix.is_empty() {
        output.extend_from_slice(prefix);
        output.push(b':');
    }
    output.extend_from_slice(local_name);
}

fn append_text_element_named(output: &mut Vec<u8>, prefix: &[u8], local_name: &[u8], text: &str) {
    output.push(b'<');
    append_qualified_name(output, prefix, local_name);
    if text.is_empty() {
        output.extend_from_slice(b"/>");
        return;
    }
    let preserve = text.chars().next().is_some_and(char::is_whitespace)
        || text.chars().next_back().is_some_and(char::is_whitespace);
    if preserve {
        output.extend_from_slice(b" xml:space=\"preserve\"");
    }
    output.push(b'>');
    append_escaped_xml(output, text);
    output.extend_from_slice(b"</");
    append_qualified_name(output, prefix, local_name);
    output.push(b'>');
}

fn append_escaped_xml(output: &mut Vec<u8>, text: &str) {
    let mut cursor = 0usize;
    for (index, byte) in text.bytes().enumerate() {
        let replacement = match byte {
            b'&' => Some(b"&amp;".as_slice()),
            b'<' => Some(b"&lt;".as_slice()),
            b'>' => Some(b"&gt;".as_slice()),
            b'"' => Some(b"&quot;".as_slice()),
            b'\'' => Some(b"&apos;".as_slice()),
            _ => None,
        };
        if let Some(replacement) = replacement {
            output.extend_from_slice(&text.as_bytes()[cursor..index]);
            output.extend_from_slice(replacement);
            cursor = index + 1;
        }
    }
    output.extend_from_slice(&text.as_bytes()[cursor..]);
}

fn try_plain_paragraph(conformance: Conformance, text: &str) -> TransactionResult<Vec<u8>> {
    let namespace = conformance.namespace().as_bytes();
    let prefix = b"<w:p xmlns:w=\"";
    if text.is_empty() {
        let capacity = replacement_len_add(
            replacement_len_add(prefix.len(), namespace.len())?,
            b"\"/>".len(),
        )?;
        if capacity > MAX_DOCUMENT_XML_BYTES {
            return Err(TransactionError::Limit {
                resource: "XML bytes",
                max: MAX_DOCUMENT_XML_BYTES,
                actual: capacity,
            });
        }
        let mut paragraph = Vec::new();
        paragraph.try_reserve_exact(capacity).map_err(|source| {
            TransactionError::Document(crate::Error::Allocation {
                resource: "plain paragraph XML",
                source,
            })
        })?;
        paragraph.extend_from_slice(prefix);
        paragraph.extend_from_slice(namespace);
        paragraph.extend_from_slice(b"\"/>");
        return Ok(paragraph);
    }
    let text_len = text_element_named_len(b"w", b"t", text)?;
    let capacity = replacement_len_add(
        replacement_len_add(
            replacement_len_add(prefix.len(), namespace.len())?,
            b"\"><w:r>".len(),
        )?,
        replacement_len_add(text_len, b"</w:r></w:p>".len())?,
    )?;
    if capacity > MAX_DOCUMENT_XML_BYTES {
        return Err(TransactionError::Limit {
            resource: "XML bytes",
            max: MAX_DOCUMENT_XML_BYTES,
            actual: capacity,
        });
    }
    let mut paragraph = Vec::new();
    paragraph.try_reserve_exact(capacity).map_err(|source| {
        TransactionError::Document(crate::Error::Allocation {
            resource: "plain paragraph XML",
            source,
        })
    })?;
    paragraph.extend_from_slice(prefix);
    paragraph.extend_from_slice(namespace);
    paragraph.extend_from_slice(b"\"><w:r>");
    append_text_element_named(&mut paragraph, b"w", b"t", text);
    paragraph.extend_from_slice(b"</w:r></w:p>");
    debug_assert_eq!(paragraph.len(), capacity);
    Ok(paragraph)
}

fn replace_range(
    source: &[u8],
    start: usize,
    end: usize,
    replacement: &[u8],
) -> TransactionResult<Vec<u8>> {
    if start > end || end > source.len() {
        return Err(crate::Error::InvalidFormat("invalid document rewrite range".into()).into());
    }
    let capacity = source
        .len()
        .checked_sub(end - start)
        .and_then(|size| size.checked_add(replacement.len()))
        .ok_or(TransactionError::Limit {
            resource: "projected XML bytes",
            max: MAX_DOCUMENT_XML_BYTES,
            actual: usize::MAX,
        })?;
    if capacity > MAX_DOCUMENT_XML_BYTES {
        return Err(TransactionError::Limit {
            resource: "projected XML bytes",
            max: MAX_DOCUMENT_XML_BYTES,
            actual: capacity,
        });
    }
    let mut output = Vec::new();
    output
        .try_reserve_exact(capacity)
        .map_err(|allocation_error| crate::Error::Allocation {
            resource: "document transaction XML",
            source: allocation_error,
        })?;
    output.extend_from_slice(&source[..start]);
    output.extend_from_slice(replacement);
    output.extend_from_slice(&source[end..]);
    Ok(output)
}

#[cfg(test)]
mod tests {
    use super::*;

    const WORD: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    /// Compare a snapshot's retained layout with a full rescan of its own
    /// bytes.
    ///
    /// The incremental splice path must be indistinguishable from
    /// [`Snapshot::from_xml`], so this is the differential the corpus tests
    /// assert on every rewrite they accept.
    fn layout_matches_rescan(snapshot: &Snapshot) -> bool {
        fn ranges_equal(retained: &[Range], scanned: &[Range]) -> bool {
            retained.len() == scanned.len()
                && retained
                    .iter()
                    .zip(scanned)
                    .all(|(left, right)| left.start == right.start && left.length == right.length)
        }

        let Ok(layout) = scan_document(snapshot.xml_bytes()) else {
            return false;
        };
        ranges_equal(&snapshot.paragraphs, &layout.paragraphs)
            && ranges_equal(&snapshot.tables, &layout.tables)
            && ranges_equal(&snapshot.block_controls, &layout.block_controls)
            && snapshot.content_end == layout.content_end
            && snapshot.conformance.namespace() == layout.conformance.namespace()
    }

    fn paragraph_bytes(snapshot: &Snapshot, index: usize) -> (usize, usize) {
        let range = snapshot.paragraphs[index];
        let start = usize::try_from(range.start).unwrap();
        (start, start + usize::try_from(range.length).unwrap())
    }

    /// Assert the differential for one rewrite, and report whether the
    /// incremental layout was the route that produced it.
    fn rewrite_matches_rescan(base: &Snapshot, index: usize, text: &str) -> Option<bool> {
        let mut edit = base.edit();
        edit.replace_paragraph_text(Position::new(index), text)
            .ok()?;
        let projected = edit.projected();
        assert!(
            layout_matches_rescan(projected),
            "incremental layout differs from a rescan after rewriting paragraph {index}"
        );
        let (start, end) = paragraph_bytes(base, index);
        let (replacement_start, replacement_end) = paragraph_bytes(projected, index);
        let replacement = &projected.xml_bytes()[replacement_start..replacement_end];
        Some(
            base.spliced_layout(
                projected.xml_bytes().len(),
                &[ParagraphSplice {
                    start,
                    end,
                    replacement,
                }],
            )
            .unwrap()
            .is_some(),
        )
    }

    fn sampled_positions(count: usize, wanted: usize) -> Vec<usize> {
        if count == 0 || wanted == 0 {
            return Vec::new();
        }
        let step = count.div_ceil(wanted).max(1);
        (0..count).step_by(step).collect()
    }

    fn docx_fixture_paths() -> Vec<std::path::PathBuf> {
        let mut paths = Vec::new();
        let mut stack = vec![
            std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
                .join("..")
                .join("..")
                .join("test-data"),
        ];
        while let Some(directory) = stack.pop() {
            let Ok(entries) = std::fs::read_dir(&directory) else {
                continue;
            };
            for entry in entries.flatten() {
                let path = entry.path();
                if path.is_dir() {
                    stack.push(path);
                } else if path
                    .extension()
                    .is_some_and(|extension| extension.eq_ignore_ascii_case("docx"))
                {
                    paths.push(path);
                }
            }
        }
        paths.sort();
        paths
    }

    fn generated_corpus_snapshot(paragraphs: usize) -> Snapshot {
        let mut package = crate::Package::new().unwrap();
        {
            let document = package.document_mut().unwrap();
            for index in 0..paragraphs {
                document.add_paragraph_with_text(&format!(
                    "litchi-perf-baseline-docx-semantic-v1-source-{index:05}"
                ));
            }
        }
        let mut output = std::io::Cursor::new(Vec::new());
        package.to_stream(&mut output).unwrap();
        crate::Package::from_reader(std::io::Cursor::new(output.into_inner()))
            .unwrap()
            .document_snapshot()
            .unwrap()
    }

    /// A compact main document whose untouched paragraph carries an attribute
    /// value the compactor re-escapes.
    ///
    /// The document satisfies the publication audit exactly as written — one
    /// ASCII space between attributes, no whitespace-only text run, no
    /// whitespace before a tag close — so the preserving policy publishes it
    /// as it stands, while whole-document compaction rewrites `'` to `&apos;`
    /// in a paragraph no edit named.
    fn apostrophe_corpus() -> Vec<u8> {
        document(concat!(
            "<w:p><w:pPr><w:pStyle w:val=\"it's\"/></w:pPr><w:r><w:t>first</w:t></w:r></w:p>",
            "<w:p><w:r><w:t>second</w:t></w:r></w:p>",
        ))
    }

    #[test]
    fn the_default_policy_leaves_an_untouched_paragraph_byte_for_byte() {
        let base = Snapshot::from_xml(apostrophe_corpus()).unwrap();
        assert_eq!(
            base.edit().compaction_policy(),
            CompactionPolicy::PreserveUnmodified
        );
        let untouched = paragraph_bytes(&base, 0);
        let untouched = base.xml_bytes()[untouched.0..untouched.1].to_vec();
        assert!(untouched.windows(6).any(|window| window == b"it's\"/"));

        let mut edit = base.edit();
        edit.replace_paragraph_text(Position::new(1), "second edited")
            .unwrap();
        let commit = edit.commit().unwrap();
        let published = commit.snapshot();
        let range = paragraph_bytes(published, 0);
        assert_eq!(
            &published.xml_bytes()[range.0..range.1],
            untouched.as_slice(),
            "the default policy must republish an untouched paragraph byte for byte"
        );
        assert_eq!(
            published
                .paragraph(Position::new(1))
                .unwrap()
                .text()
                .unwrap(),
            "second edited"
        );
    }

    #[test]
    fn the_opt_in_policy_compacts_every_paragraph_including_untouched_ones() {
        let base = Snapshot::from_xml(apostrophe_corpus()).unwrap();
        let mut edit = base
            .edit()
            .with_compaction_policy(CompactionPolicy::WholeDocument);
        assert_eq!(edit.compaction_policy(), CompactionPolicy::WholeDocument);
        edit.replace_paragraph_text(Position::new(1), "second edited")
            .unwrap();
        let commit = edit.commit().unwrap();
        let published = commit.snapshot();
        let range = paragraph_bytes(published, 0);
        let untouched = &published.xml_bytes()[range.0..range.1];
        assert!(
            untouched.windows(10).any(|window| window == b"it&apos;s\""),
            "whole-document compaction re-escapes an untouched paragraph's attributes"
        );
        assert_eq!(
            published
                .paragraph(Position::new(1))
                .unwrap()
                .text()
                .unwrap(),
            "second edited"
        );
        assert_eq!(published.paragraph_count(), base.paragraph_count());
    }

    #[test]
    fn a_source_the_publication_audit_refuses_takes_the_whole_document_route() {
        // Change 0665 leaves every structural and finite-budget refusal in
        // place, so the preserving policy still has a fallback. One token
        // larger than the auditor's `TokenBytes` budget is such a refusal and
        // the snapshot admits it, because the snapshot's own budgets are
        // counted differently.
        let oversized = "a".repeat(4 * 1024 * 1024 + 1);
        let xml = document(&format!(
            "<w:p><w:pPr><w:pStyle w:val=\"{oversized}\"/></w:pPr><w:r><w:t>first</w:t></w:r></w:p>\
             <w:p><w:r><w:t>second</w:t></w:r></w:p>"
        ));
        assert!(!publication_accepts_preserved_xml(&xml));
        let base = Snapshot::from_xml(xml).unwrap();

        let mut preserving = base.edit();
        preserving
            .replace_paragraph_text(Position::new(1), "second edited")
            .unwrap();
        let preserved = preserving.commit().unwrap();

        let mut compacting = base
            .edit()
            .with_compaction_policy(CompactionPolicy::WholeDocument);
        compacting
            .replace_paragraph_text(Position::new(1), "second edited")
            .unwrap();
        let compacted = compacting.commit().unwrap();

        assert_eq!(
            preserved.snapshot().xml_bytes(),
            compacted.snapshot().xml_bytes(),
            "the fallback must publish exactly what the whole-document route publishes"
        );
    }

    #[test]
    fn an_exact_no_op_is_untouched_by_either_policy() {
        let base = Snapshot::from_xml(apostrophe_corpus()).unwrap();
        for policy in [
            CompactionPolicy::PreserveUnmodified,
            CompactionPolicy::WholeDocument,
        ] {
            let commit = base.edit().with_compaction_policy(policy).commit().unwrap();
            assert!(!commit.patch().changed());
            assert_eq!(commit.snapshot().xml_bytes(), base.xml_bytes());
        }
    }

    #[test]
    fn the_preserving_policy_survives_a_clone_and_an_inherited_ancestor_space() {
        let base = Snapshot::from_xml(apostrophe_corpus()).unwrap();
        let edit = base
            .edit()
            .with_compaction_policy(CompactionPolicy::WholeDocument);
        assert_eq!(
            edit.try_clone().unwrap().compaction_policy(),
            CompactionPolicy::WholeDocument
        );

        // `w:body` carrying `xml:space` is state the paragraph fragment cannot
        // see, so the preserving policy declines to compact fragments at all
        // and publishes the projection unchanged.
        let xml = format!(
            "<w:document xmlns:w=\"{WORD}\"><w:body xml:space=\"preserve\">\
             <w:p><w:r><w:t>one</w:t></w:r></w:p><w:sectPr/></w:body></w:document>"
        )
        .into_bytes();
        let base = Snapshot::from_xml(xml).unwrap();
        assert!(inherits_xml_space(&base));
        let mut edit = base.edit();
        edit.replace_paragraph_text(Position::new(0), "two")
            .unwrap();
        let projected = edit.projected().xml_bytes().to_vec();
        let commit = edit.commit().unwrap();
        assert_eq!(commit.snapshot().xml_bytes(), projected.as_slice());
    }

    #[test]
    fn the_gate_accepts_every_noncompact_spelling_and_still_refuses_structure() {
        let compact = apostrophe_corpus();
        assert!(publication_accepts_preserved_xml(&compact));

        // Change 0665: every one of these is a compactness verdict the writer
        // no longer raises, so the preserving policy must publish them. Each
        // declares its prefix, as change 0750's namespace check requires.
        let accepted: [&[u8]; 4] = [
            b"<?xml version=\"1.0\"?>\n<w:p xmlns:w=\"urn:w\"/>",
            b" <w:p xmlns:w=\"urn:w\"/>",
            b"<w:p xmlns:w=\"urn:w\"/>\n",
            b"<w:p xmlns:w=\"urn:w\" a=\"1\"\n     b=\"2\"/>",
        ];
        for case in accepted {
            assert!(
                publication_accepts_preserved_xml(case),
                "the gate should accept {:?}",
                String::from_utf8_lossy(case)
            );
        }

        // Structural refusals stay, and the gate still keeps them off the
        // preserving route.
        let refused: [&[u8]; 4] = [
            // Character data outside the document element.
            b"x<w:p/>",
            // An unterminated processing instruction.
            b"<?xml version=\"1.0\"",
            // Nothing at all.
            b"",
            // A document type declaration.
            b"<!DOCTYPE w:document><w:p/>",
        ];
        for case in refused {
            assert!(
                !publication_accepts_preserved_xml(case),
                "the gate should refuse {:?}",
                String::from_utf8_lossy(case)
            );
        }

        // The gate is the auditor the writer runs, over the whole corpus.
        let mut audited = 0usize;
        let mut accepted_fixtures = 0usize;
        for path in docx_fixture_paths() {
            let Ok(archive) = std::fs::read(&path) else {
                continue;
            };
            let Ok(package) = crate::Package::from_reader(std::io::Cursor::new(archive)) else {
                continue;
            };
            let Ok(base) = package.document_snapshot() else {
                continue;
            };
            audited += 1;
            let xml = base.xml_bytes();
            assert_eq!(
                publication_accepts_preserved_xml(xml),
                xml_minifier::audit::verify_source(xml, xml_minifier::audit::Limits::default())
                    .is_ok(),
                "the gate contradicted the auditor on {}",
                path.display()
            );
            if publication_accepts_preserved_xml(xml) {
                accepted_fixtures += 1;
            }
        }
        assert!(
            audited >= 50,
            "the corpus should reach this test: {audited}"
        );
        assert!(
            accepted_fixtures + 1 >= audited,
            "at most one corpus fixture may still take the whole-document route: \
             {accepted_fixtures} of {audited} accepted"
        );
    }

    #[test]
    fn both_policies_agree_across_the_docx_fixture_corpus() {
        let mut fixtures = 0usize;
        let mut edited = 0usize;
        let mut preserving = 0usize;
        let mut identical = 0usize;
        for path in docx_fixture_paths() {
            let Ok(archive) = std::fs::read(&path) else {
                continue;
            };
            let Ok(package) = crate::Package::from_reader(std::io::Cursor::new(archive)) else {
                continue;
            };
            let Ok(base) = package.document_snapshot() else {
                continue;
            };
            fixtures += 1;
            let gate = publication_accepts_preserved_xml(base.xml_bytes());
            preserving += usize::from(gate);
            for index in sampled_positions(base.paragraph_count(), 16) {
                let mut default_edit = base.edit();
                let Ok(_staged) =
                    default_edit.replace_paragraph_text(Position::new(index), "litchi 0660")
                else {
                    continue;
                };
                let mut optin_edit = base
                    .edit()
                    .with_compaction_policy(CompactionPolicy::WholeDocument);
                optin_edit
                    .replace_paragraph_text(Position::new(index), "litchi 0660")
                    .expect("the same rewrite is accepted under either policy");
                let default_commit = default_edit.commit().unwrap();
                let optin_commit = optin_edit.commit().unwrap();
                edited += 1;

                // Either policy keeps the document readable and the edit
                // exactly where it was made.
                for commit in [&default_commit, &optin_commit] {
                    assert_eq!(
                        commit.snapshot().paragraph_count(),
                        base.paragraph_count(),
                        "{} paragraph {index}",
                        path.display()
                    );
                    assert_eq!(
                        commit.snapshot().table_count(),
                        base.table_count(),
                        "{} paragraph {index}",
                        path.display()
                    );
                    assert_eq!(
                        commit
                            .snapshot()
                            .paragraph(Position::new(index))
                            .unwrap()
                            .text()
                            .unwrap(),
                        "litchi 0660",
                        "{} paragraph {index}",
                        path.display()
                    );
                    assert!(layout_matches_rescan(commit.snapshot()));
                }

                if gate {
                    // Preservation: every paragraph the edit did not name is
                    // republished byte for byte.
                    for other in 0..base.paragraph_count() {
                        if other == index {
                            continue;
                        }
                        let before = paragraph_bytes(&base, other);
                        let after = paragraph_bytes(default_commit.snapshot(), other);
                        assert_eq!(
                            &base.xml_bytes()[before.0..before.1],
                            &default_commit.snapshot().xml_bytes()[after.0..after.1],
                            "{} paragraph {other} moved under the preserving policy",
                            path.display()
                        );
                    }
                } else {
                    // The fallback publishes exactly what the whole-document
                    // route publishes.
                    assert_eq!(
                        default_commit.snapshot().xml_bytes(),
                        optin_commit.snapshot().xml_bytes(),
                        "{} paragraph {index} diverged on the fallback",
                        path.display()
                    );
                    identical += 1;
                }
            }
        }
        assert!(
            fixtures >= 50 && edited >= 80,
            "the corpus should reach this test: {fixtures} fixtures, {edited} edits"
        );
        // Change 0665 removed the compactness refusal from the eager
        // publication audit, so the gate no longer sends this corpus down the
        // whole-document fallback: every fixture preserves.
        assert_eq!(
            identical, 0,
            "no fixture should still need the audit fallback: {identical}/{edited}"
        );
        assert_eq!(
            preserving, fixtures,
            "every fixture should take the preserving route: {preserving}/{fixtures}"
        );
    }

    #[test]
    fn incremental_paragraph_layout_equals_a_rescan_across_the_docx_fixture_corpus() {
        let mut fixtures = 0usize;
        let mut rewrites = 0usize;
        let mut incremental = 0usize;
        let mut batches = 0usize;
        for path in docx_fixture_paths() {
            let Ok(archive) = std::fs::read(&path) else {
                continue;
            };
            let Ok(package) = crate::Package::from_reader(std::io::Cursor::new(archive)) else {
                continue;
            };
            let Ok(base) = package.document_snapshot() else {
                continue;
            };
            fixtures += 1;
            let positions = sampled_positions(base.paragraph_count(), 16);
            for index in &positions {
                for text in ["s", "", "a much longer replacement paragraph body"] {
                    if let Some(spliced) = rewrite_matches_rescan(&base, *index, text) {
                        rewrites += 1;
                        incremental += usize::from(spliced);
                    }
                }
            }
            if positions.len() > 1 {
                let replacements = positions
                    .iter()
                    .map(|index| {
                        ParagraphTextReplacement::new(
                            Position::new(*index),
                            format!("batch {index} replacement"),
                        )
                    })
                    .collect::<Vec<_>>();
                let mut edit = base.edit();
                if edit.replace_body_paragraph_texts(&replacements).is_ok() {
                    batches += 1;
                    assert!(
                        layout_matches_rescan(edit.projected()),
                        "the incremental batch layout differs from a rescan for {}",
                        path.display()
                    );
                }
            }
        }
        assert!(
            fixtures >= 50,
            "the DOCX fixture corpus should reach this test: {fixtures} snapshot-able fixtures"
        );
        assert!(
            rewrites >= 200 && batches >= 2,
            "the differential should accept real rewrites: {rewrites} single, {batches} batched"
        );
        assert!(
            incremental * 10 >= rewrites * 9,
            "the incremental layout should carry the fixture rewrites: {incremental}/{rewrites}"
        );
    }

    #[test]
    fn incremental_paragraph_layout_equals_a_rescan_on_the_generated_corpora() {
        for paragraphs in [24usize, 200, 10_000] {
            let base = generated_corpus_snapshot(paragraphs);
            assert_eq!(base.paragraph_count(), paragraphs);
            for index in [0, paragraphs / 2, paragraphs - 1] {
                for text in [
                    "",
                    "short",
                    "a replacement that is longer than the source text",
                ] {
                    let spliced = rewrite_matches_rescan(&base, index, text)
                        .expect("a generated paragraph rewrite is always accepted");
                    assert!(
                        spliced,
                        "the incremental layout should carry the {paragraphs}-paragraph corpus"
                    );
                }
            }
            let updates = paragraphs.div_ceil(100);
            let replacements = (0..updates)
                .map(|update| {
                    let index = update * paragraphs / updates;
                    ParagraphTextReplacement::new(
                        Position::new(index),
                        format!("litchi-perf-baseline-docx-semantic-v1-updated-{index:05}"),
                    )
                })
                .collect::<Vec<_>>();
            let mut edit = base.edit();
            edit.replace_body_paragraph_texts(&replacements).unwrap();
            assert!(
                layout_matches_rescan(edit.projected()),
                "the incremental batch layout differs from a rescan on {paragraphs} paragraphs"
            );
            let commit = edit.commit().unwrap();
            assert!(layout_matches_rescan(commit.snapshot()));
        }
    }

    #[test]
    fn incremental_paragraph_layout_tracks_tables_controls_and_the_final_section() {
        let xml = document(concat!(
            "<w:tbl><w:tr><w:tc><w:p><w:r><w:t>cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>",
            "<w:p><w:r><w:t>one</w:t></w:r></w:p>",
            "<w:sdt><w:sdtContent><w:p><w:r><w:t>control</w:t></w:r></w:p></w:sdtContent></w:sdt>",
            "<w:p><w:r><w:t>two</w:t></w:r></w:p>",
            "<w:tbl><w:tr><w:tc><w:p><w:r><w:t>tail</w:t></w:r></w:p></w:tc></w:tr></w:tbl>",
        ));
        let base = Snapshot::from_xml(xml).unwrap();
        assert_eq!(base.paragraph_count(), 2);
        assert_eq!(base.table_count(), 2);
        assert_eq!(base.block_content_control_count(), 1);
        for index in [0usize, 1] {
            for text in ["", "x", "a considerably longer paragraph body than before"] {
                let spliced = rewrite_matches_rescan(&base, index, text)
                    .expect("a direct-body paragraph rewrite is accepted here");
                assert!(
                    spliced,
                    "paragraph {index} should take the incremental route"
                );
            }
        }
    }

    #[test]
    fn a_shape_changing_replacement_declines_the_incremental_layout() {
        let xml = document("<w:p><w:r><w:t>a</w:t></w:r></w:p>");
        let base = Snapshot::from_xml(xml).unwrap();
        let (start, end) = paragraph_bytes(&base, 0);
        let source = &base.xml_bytes()[start..end];
        for replacement in [
            // One more element than the fragment it replaces.
            &b"<w:p><w:r><w:t>a</w:t><w:t>b</w:t></w:r></w:p>"[..],
            // One level deeper than the fragment it replaces.
            &b"<w:p><w:r><w:smartTag><w:t>a</w:t></w:smartTag></w:r></w:p>"[..],
            // A different root element.
            &b"<w:tbl><w:tr><w:tc><w:p><w:r><w:t>a</w:t></w:r></w:p></w:tc></w:tr></w:tbl>"[..],
            // Two roots rather than one.
            &b"<w:p><w:r><w:t>a</w:t></w:r></w:p><w:p/>"[..],
            // Markup the document scanner refuses.
            &b"<w:p><?target instruction?></w:p>"[..],
            // Not balanced.
            &b"<w:p><w:r><w:t>a</w:t></w:r>"[..],
            // Padded rather than exactly the element.
            &b" <w:p><w:r><w:t>a</w:t></w:r></w:p>"[..],
            &b"<w:p><w:r><w:t>a</w:t></w:r></w:p> "[..],
        ] {
            assert!(
                !preserves_body_child_shape(source, replacement),
                "a shape-changing replacement must decline the incremental layout"
            );
            let spliced = base
                .spliced_layout(
                    base.xml_bytes().len() + replacement.len() - (end - start),
                    &[ParagraphSplice {
                        start,
                        end,
                        replacement,
                    }],
                )
                .unwrap();
            assert!(spliced.is_none());
        }
    }

    #[test]
    fn declining_the_incremental_layout_still_rescans_the_spliced_document() {
        let xml = document("<w:p><w:r><w:t>a</w:t></w:r></w:p><w:p><w:r><w:t>b</w:t></w:r></w:p>");
        let base = Snapshot::from_xml(xml).unwrap();
        let (start, end) = paragraph_bytes(&base, 0);
        let replacement = &b"<w:p><w:r><w:t>a</w:t></w:r><w:r><w:t>grown</w:t></w:r></w:p>"[..];
        let spliced = replace_range(base.xml_bytes(), start, end, replacement).unwrap();
        let expected = Snapshot::from_xml(spliced.clone()).unwrap();
        let candidate = base
            .with_spliced_paragraphs(
                spliced,
                &[ParagraphSplice {
                    start,
                    end,
                    replacement,
                }],
            )
            .unwrap();
        assert!(layout_matches_rescan(&candidate));
        assert_eq!(candidate.paragraph_count(), expected.paragraph_count());
        assert_eq!(candidate.xml_bytes(), expected.xml_bytes());
    }

    #[test]
    fn a_splice_that_is_not_a_whole_paragraph_declines_the_incremental_layout() {
        let xml = document("<w:p><w:r><w:t>a</w:t></w:r></w:p>");
        let base = Snapshot::from_xml(xml).unwrap();
        let (start, end) = paragraph_bytes(&base, 0);
        let replacement = &b"<w:p><w:r><w:t>a</w:t></w:r></w:p>"[..];
        for (start, end) in [(start + 1, end), (start, end - 1), (start, end + 1)] {
            assert!(
                base.spliced_layout(
                    base.xml_bytes().len(),
                    &[ParagraphSplice {
                        start,
                        end,
                        replacement,
                    }],
                )
                .unwrap()
                .is_none()
            );
        }
    }

    #[test]
    fn body_child_shape_measures_one_balanced_element_and_refuses_anything_else() {
        let shape = body_child_shape(b"<w:p><w:r><w:t>a</w:t></w:r></w:p>").unwrap();
        assert_eq!(shape.nodes, 3);
        assert_eq!(shape.max_depth, 3);
        assert_eq!(shape.root_tag_len, "<w:p>".len());
        let empty = body_child_shape(b"<w:p/>").unwrap();
        assert_eq!(empty.nodes, 1);
        assert_eq!(empty.max_depth, 1);
        assert_eq!(empty.root_tag_len, "<w:p/>".len());
        assert!(body_child_shape(b"<!DOCTYPE w:p><w:p/>").is_none());
        assert!(body_child_shape(b"<?xml version=\"1.0\"?><w:p/>").is_none());
        assert!(body_child_shape(b"<w:p></w:r>").is_none());
        assert!(body_child_shape(b"").is_none());
        assert!(body_child_shape(b"text").is_none());
    }

    fn document(body: &str) -> Vec<u8> {
        format!("<w:document xmlns:w=\"{WORD}\"><w:body>{body}<w:sectPr/></w:body></w:document>")
            .into_bytes()
    }

    fn durable_limits() -> litchi_core::patch::PatchLimits {
        litchi_core::patch::PatchLimits::new(
            litchi_core::patch::BlobLimits::new(1, MAX_DOCUMENT_XML_BYTES, MAX_DOCUMENT_XML_BYTES),
            1024 * 1024,
            32,
            8,
            256 * 1024,
            512 * 1024,
        )
    }

    #[test]
    fn source_backed_overlay_classification_refuses_transfer_operations() {
        let ordinary = Operation::InsertParagraph {
            position: Position::new(0),
            text: "plain".into(),
        };
        let transferred = Operation::InsertTransferredParagraph {
            position: Position::new(0),
            xml: Arc::new(b"<w:p/>".to_vec()),
            dependency_digest: Arc::from("before"),
            inverse_dependency_digest: Arc::from("after"),
            graph: Arc::new(TransferGraph::empty()),
        };

        assert!(ordinary.supports_source_backed_main_document_overlay());
        assert!(!transferred.supports_source_backed_main_document_overlay());
    }

    #[test]
    fn source_backed_snapshot_reuses_the_exact_raw_xml_allocation() {
        let xml = Arc::new(document("<w:p><w:r><w:t>source backed</w:t></w:r></w:p>"));
        let allocation = xml.as_ptr();
        let snapshot = Snapshot::from_shared_xml(Arc::clone(&xml)).unwrap();

        assert_eq!(snapshot.xml_bytes().as_ptr(), allocation);
        assert!(matches!(&snapshot.xml, XmlStorage::Owned(value) if Arc::ptr_eq(value, &xml)));
    }

    #[test]
    fn byte_order_marked_snapshot_keeps_layout_ranges_and_edit_output() {
        let mut marked = UTF8_BYTE_ORDER_MARK.to_vec();
        marked.extend_from_slice(&document(
            "<w:p><w:r><w:t>before</w:t></w:r></w:p><w:p><w:r><w:t>kept</w:t></w:r></w:p>",
        ));
        let source = Snapshot::from_xml(marked).unwrap();
        assert_eq!(source.paragraph_count(), 2);
        assert_eq!(
            source.paragraph(Position::new(0)).unwrap().text().unwrap(),
            "before"
        );
        assert_eq!(
            source.paragraph(Position::new(1)).unwrap().text().unwrap(),
            "kept"
        );
        assert!(layout_matches_rescan(&source));

        let mut edit = source.edit();
        edit.replace_paragraph_text(Position::new(0), "after")
            .unwrap();
        let committed = edit.commit().unwrap();
        assert_eq!(
            committed
                .snapshot()
                .paragraph(Position::new(0))
                .unwrap()
                .text()
                .unwrap(),
            "after"
        );
        assert_eq!(
            committed
                .snapshot()
                .paragraph(Position::new(1))
                .unwrap()
                .text()
                .unwrap(),
            "kept"
        );
        assert!(
            committed
                .snapshot()
                .xml_bytes()
                .starts_with(UTF8_BYTE_ORDER_MARK)
        );
        assert!(layout_matches_rescan(committed.snapshot()));
    }

    #[test]
    fn length_changing_text_edit_preserves_formatting_and_is_reversible() {
        let source = Snapshot::from_xml(document(
            "<w:p w:rsidR=\"1\"><w:pPr><w:keepNext/></w:pPr><w:r><w:rPr><w:b/></w:rPr><w:t>old</w:t></w:r></w:p>",
        ))
        .unwrap();
        let mut edit = source.edit();
        edit.replace_paragraph_text(Position::new(0), " longer & text ")
            .unwrap();
        let commit = edit.commit().unwrap();

        assert_eq!(
            commit
                .snapshot()
                .paragraph(Position::new(0))
                .unwrap()
                .text()
                .unwrap(),
            " longer & text "
        );
        assert!(std::str::from_utf8(commit.snapshot().xml_bytes()).unwrap().contains("<w:pPr><w:keepNext/></w:pPr><w:r><w:rPr><w:b/></w:rPr><w:t xml:space=\"preserve\"> longer &amp; text </w:t></w:r>"));
        assert_eq!(commit.diagnostics().operations(), 1);
        assert!(commit.diagnostics().changed());

        let restored = commit.patch().inverse().apply(commit.snapshot()).unwrap();
        assert_eq!(restored.xml_bytes(), source.xml_bytes());
        assert!(commit.patch().apply(&restored).is_ok());
    }

    #[test]
    fn paragraph_text_batch_is_canonical_atomic_and_reversible() {
        let source = Snapshot::from_xml(document(
            "<w:p w:rsidR=\"1\"><w:r><w:rPr><w:b/></w:rPr><w:t>first</w:t></w:r></w:p><w:p><w:r><w:t>retained</w:t></w:r></w:p><w:p><w:r><w:rPr><w:i/></w:rPr><w:t>third</w:t></w:r></w:p>",
        ))
        .unwrap();
        let replacements = [
            ParagraphTextReplacement::new(Position::new(0), "first changed"),
            ParagraphTextReplacement::new(Position::new(2), "third changed"),
        ];
        let mut edit = source.edit();
        edit.replace_body_paragraph_texts(&replacements).unwrap();
        let commit = edit.commit().unwrap();

        let mut scalar = source.edit();
        for replacement in &replacements {
            scalar
                .replace_paragraph_text(replacement.position(), replacement.text())
                .unwrap();
        }
        let scalar = scalar.commit().unwrap();
        assert_eq!(commit.snapshot().xml_bytes(), scalar.snapshot().xml_bytes());

        assert_eq!(commit.diagnostics().operations(), 2);
        assert_eq!(
            commit
                .snapshot()
                .paragraph(Position::new(0))
                .unwrap()
                .text()
                .unwrap(),
            "first changed"
        );
        assert_eq!(
            commit
                .snapshot()
                .paragraph(Position::new(1))
                .unwrap()
                .text()
                .unwrap(),
            "retained"
        );
        let xml = std::str::from_utf8(commit.snapshot().xml_bytes()).unwrap();
        assert!(xml.contains("<w:rPr><w:b/></w:rPr>"));
        assert!(xml.contains("<w:rPr><w:i/></w:rPr>"));
        assert_eq!(
            commit
                .patch()
                .inverse()
                .apply(commit.snapshot())
                .unwrap()
                .xml_bytes(),
            source.xml_bytes()
        );
        let durable = commit.patch().to_durable(durable_limits()).unwrap();
        let wire = durable.to_deterministic_json().unwrap();
        let decoded =
            litchi_core::patch::Patch::<litchi_core::patch::Reversible>::from_deterministic_json(
                &wire,
                durable_limits(),
            )
            .unwrap();
        let applied = source.apply_durable(&decoded).unwrap();
        assert_eq!(applied.xml_bytes(), commit.snapshot().xml_bytes());
        assert_eq!(
            applied
                .apply_durable(&decoded.inverse())
                .unwrap()
                .xml_bytes(),
            source.xml_bytes()
        );

        let noops = [
            ParagraphTextReplacement::new(Position::new(0), "first"),
            ParagraphTextReplacement::new(Position::new(2), "third"),
        ];
        let mut noop = source.edit();
        noop.replace_body_paragraph_texts(&noops).unwrap();
        let noop = noop.commit().unwrap();
        assert!(!noop.diagnostics().changed());
        assert_eq!(noop.diagnostics().operations(), 0);
        assert_eq!(
            noop.snapshot().xml_bytes().as_ptr(),
            source.xml_bytes().as_ptr()
        );

        let duplicate = [
            ParagraphTextReplacement::new(Position::new(0), "one"),
            ParagraphTextReplacement::new(Position::new(0), "duplicate"),
        ];
        let out_of_order = [
            ParagraphTextReplacement::new(Position::new(2), "third first"),
            ParagraphTextReplacement::new(Position::new(0), "first second"),
        ];
        let late_failure = [
            ParagraphTextReplacement::new(Position::new(0), "would change"),
            ParagraphTextReplacement::new(Position::new(2), "line\nbreak"),
        ];
        let mut refused = source.edit();
        assert!(matches!(
            refused.replace_body_paragraph_texts(&duplicate),
            Err(TransactionError::Refused {
                reason: Refusal::AmbiguousCompositeSelector,
                ..
            })
        ));
        assert!(matches!(
            refused.replace_body_paragraph_texts(&[]),
            Err(TransactionError::Refused {
                reason: Refusal::AmbiguousCompositeSelector,
                ..
            })
        ));
        assert!(matches!(
            refused.replace_body_paragraph_texts(&out_of_order),
            Err(TransactionError::Refused {
                reason: Refusal::AmbiguousCompositeSelector,
                ..
            })
        ));
        assert!(refused.replace_body_paragraph_texts(&late_failure).is_err());
        assert_eq!(refused.projected().xml_bytes(), source.xml_bytes());
    }

    #[test]
    fn insertion_uses_projected_checked_positions_and_strict_namespace() {
        let strict = "http://purl.oclc.org/ooxml/wordprocessingml/main";
        let xml = format!(
            "<s:document xmlns:s=\"{strict}\"><s:body><s:p><s:r><s:t>A</s:t></s:r></s:p><s:sectPr/></s:body></s:document>"
        );
        let source = Snapshot::from_xml(xml.into_bytes()).unwrap();
        let mut edit = source.edit();
        edit.insert_paragraph(Position::new(0), "B")
            .unwrap()
            .insert_paragraph(Position::new(2), " C ")
            .unwrap();
        let commit = edit.commit().unwrap();

        let text = commit
            .snapshot()
            .paragraphs()
            .into_iter()
            .map(|paragraph| paragraph.text().unwrap())
            .collect::<Vec<_>>();
        assert_eq!(text, ["B", "A", " C "]);
        let xml = std::str::from_utf8(commit.snapshot().xml_bytes()).unwrap();
        assert!(xml.contains(&format!(
            "<w:p xmlns:w=\"{strict}\"><w:r><w:t>B</w:t></w:r></w:p>"
        )));
        assert!(!xml.contains('\n'));
    }

    #[test]
    fn refuses_complex_content_and_stale_patch_sources() {
        let source = Snapshot::from_xml(document(
            "<w:p><w:hyperlink w:anchor=\"a\"><w:r><w:t>linked</w:t></w:r></w:hyperlink></w:p><w:p><w:r><w:t>plain</w:t></w:r></w:p>",
        ))
        .unwrap();
        let mut refused = source.edit();
        assert!(matches!(
            refused.replace_paragraph_text(Position::new(0), "no"),
            Err(TransactionError::Refused {
                reason: Refusal::ComplexContent,
                ..
            })
        ));

        let mut edit = source.edit();
        edit.replace_paragraph_text(Position::new(1), "changed")
            .unwrap();
        let commit = edit.commit().unwrap();
        let stale = Snapshot::from_xml(document("<w:p><w:r><w:t>other</w:t></w:r></w:p>")).unwrap();
        assert!(matches!(
            commit.patch().apply(&stale),
            Err(TransactionError::StaleSource)
        ));
    }

    #[test]
    fn exact_noop_shares_snapshot_bytes_and_records_no_operation() {
        let source = Snapshot::from_xml(document("<w:p><w:r><w:t>same</w:t></w:r></w:p>")).unwrap();
        let mut edit = source.edit();
        edit.replace_paragraph_text(Position::new(0), "same")
            .unwrap();
        let commit = edit.commit().unwrap();

        assert!(!commit.patch().changed());
        assert!(commit.patch().operations().is_empty());
        assert!(matches!(
            (&source.xml, &commit.snapshot().xml),
            (XmlStorage::Owned(source), XmlStorage::Owned(snapshot))
                if Arc::ptr_eq(source, snapshot)
        ));
    }

    #[test]
    fn multi_run_hyperlink_and_cell_edits_preserve_formatting_and_unknown_xml() {
        let source = Snapshot::from_xml(document(
            "<w:p w:rsidR=\"01\"><w:pPr><w:keepNext/></w:pPr><w:r><w:rPr><w:b/></w:rPr><w:t>Bold</w:t></w:r><w:r><w:rPr><w:i/></w:rPr><w:drawing><x:opaque xmlns:x=\"urn:test\"/></w:drawing><w:t>tail</w:t></w:r></w:p><w:p><w:hyperlink r:id=\"rId9\" w:tooltip=\"tip\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><w:r><w:rPr><w:u/></w:rPr><w:t>link</w:t></w:r><w:r><w:t> text</w:t></w:r></w:hyperlink></w:p><w:tbl><w:tblPr><w:tblStyle w:val=\"Grid\"/></w:tblPr><w:tr><w:trPr><w:cantSplit/></w:trPr><w:tc><w:tcPr><w:shd w:fill=\"FFFF00\"/></w:tcPr><w:p><w:r><w:rPr><w:b/></w:rPr><w:t>cell</w:t></w:r><w:r><z:keep xmlns:z=\"urn:test\"/><w:t> tail</w:t></w:r></w:p></w:tc></w:tr></w:tbl>",
        ))
        .unwrap();
        let mut edit = source.edit();
        edit.replace_paragraph_text(Position::new(0), "Reformatted")
            .unwrap()
            .replace_hyperlink_text(Position::new(1), Position::new(0), "new target")
            .unwrap()
            .replace_table_cell_text(
                Position::new(0),
                Position::new(0),
                Position::new(0),
                "updated cell",
            )
            .unwrap();
        let commit = edit.commit().unwrap();
        let xml = std::str::from_utf8(commit.snapshot().xml_bytes()).unwrap();

        assert_eq!(
            commit
                .snapshot()
                .paragraph(Position::new(0))
                .unwrap()
                .text()
                .unwrap(),
            "Reformatted"
        );
        assert_eq!(
            selected_hyperlink_text(commit.snapshot(), Position::new(1), Position::new(0)).unwrap(),
            "new target"
        );
        assert_eq!(
            selected_cell_text(
                commit.snapshot(),
                Position::new(0),
                Position::new(0),
                Position::new(0)
            )
            .unwrap(),
            "updated cell"
        );
        for retained in [
            "<w:pPr><w:keepNext/></w:pPr>",
            "<w:rPr><w:b/></w:rPr>",
            "<w:rPr><w:i/></w:rPr>",
            "<x:opaque xmlns:x=\"urn:test\"/>",
            "r:id=\"rId9\"",
            "w:tooltip=\"tip\"",
            "<w:shd w:fill=\"FFFF00\"/>",
            "<z:keep xmlns:z=\"urn:test\"/>",
        ] {
            assert!(xml.contains(retained), "missing retained XML: {retained}");
        }
        assert_eq!(commit.patch().operations().len(), 3);
        assert_eq!(
            commit
                .patch()
                .inverse()
                .apply(commit.snapshot())
                .unwrap()
                .xml_bytes(),
            source.xml_bytes()
        );
    }

    #[test]
    fn direct_revision_actions_preserve_runs_are_durable_and_exactly_reversible() {
        let source = Snapshot::from_xml(document(
            "<w:p><w:r><w:t>before </w:t></w:r><w:ins w:id=\"7\" w:author=\"A\" xmlns:x=\"urn:extension\"><w:r><w:rPr><w:b/></w:rPr><w:t>added</w:t><x:opaque x:flag=\"1\">opaque text</x:opaque></w:r><w:r><w:t>second</w:t><x:second>second opaque</x:second></w:r></w:ins><w:del w:id=\"8\" w:author=\"A\"><w:r><w:rPr><w:i/></w:rPr><w:delText>gone &amp; old</w:delText></w:r></w:del><w:r><w:t>tail</w:t></w:r></w:p>",
        ))
        .unwrap();

        let mut edit = source.edit();
        edit.accept_revision(Position::new(0), RevisionKind::Insertion, Position::new(0))
            .unwrap()
            .reject_revision(Position::new(0), RevisionKind::Deletion, Position::new(0))
            .unwrap();
        let commit = edit.commit().unwrap();
        let xml = std::str::from_utf8(commit.snapshot().xml_bytes()).unwrap();
        assert!(!xml.contains("<w:ins"));
        assert!(!xml.contains("<w:del"));
        assert!(xml.contains("<w:r xmlns:x=\"urn:extension\"><w:rPr><w:b/></w:rPr><w:t>added</w:t><x:opaque x:flag=\"1\">opaque text</x:opaque></w:r>"));
        assert!(xml.contains("<w:r xmlns:x=\"urn:extension\"><w:t>second</w:t><x:second>second opaque</x:second></w:r>"));
        assert!(xml.contains("<w:t>gone &amp; old</w:t>"));
        assert!(xml.contains("<w:rPr><w:i/></w:rPr>"));
        assert_eq!(commit.patch().operations().len(), 2);
        assert_eq!(
            commit
                .patch()
                .inverse()
                .apply(commit.snapshot())
                .unwrap()
                .xml_bytes(),
            source.xml_bytes()
        );

        let durable = commit.patch().to_durable(durable_limits()).unwrap();
        let wire = durable.to_deterministic_json().unwrap();
        let decoded =
            litchi_core::patch::Patch::<litchi_core::patch::Reversible>::from_deterministic_json(
                &wire,
                durable_limits(),
            )
            .unwrap();
        assert_eq!(
            source.apply_durable(&decoded).unwrap().xml_bytes(),
            commit.snapshot().xml_bytes()
        );
        assert_eq!(
            source
                .apply_durable(&decoded)
                .unwrap()
                .apply_durable(&decoded.inverse())
                .unwrap()
                .xml_bytes(),
            source.xml_bytes()
        );

        let mut reject_insert = source.edit();
        reject_insert
            .reject_revision(Position::new(0), RevisionKind::Insertion, Position::new(0))
            .unwrap();
        let rejected = reject_insert.commit().unwrap();
        let rejected_xml = std::str::from_utf8(rejected.snapshot().xml_bytes()).unwrap();
        assert!(!rejected_xml.contains("added"));
        assert!(rejected_xml.contains("<w:del"));

        let mut accept_delete = source.edit();
        accept_delete
            .accept_revision(Position::new(0), RevisionKind::Deletion, Position::new(0))
            .unwrap();
        let accepted = accept_delete.commit().unwrap();
        let accepted_xml = std::str::from_utf8(accepted.snapshot().xml_bytes()).unwrap();
        assert!(!accepted_xml.contains("gone"));
        assert!(accepted_xml.contains("<w:ins"));
    }

    #[test]
    fn direct_revision_actions_are_atomic_and_empty_actions_are_exact_noops() {
        let unsupported = Snapshot::from_xml(document(
            "<w:p><w:ins w:id=\"1\" w:author=\"A\"><w:r><w:moveFrom w:id=\"2\" w:author=\"B\"><w:r><w:t>moved</w:t></w:r></w:moveFrom></w:r></w:ins></w:p>",
        ))
        .unwrap();
        let mut refused = unsupported.edit();
        assert!(matches!(
            refused.accept_revision(Position::new(0), RevisionKind::Insertion, Position::new(0),),
            Err(TransactionError::Refused {
                reason: Refusal::RevisionDependency,
                ..
            })
        ));
        assert_eq!(refused.projected().xml_bytes(), unsupported.xml_bytes());
        assert!(refused.operations.is_empty());

        let malformed_text = Snapshot::from_xml(document(
            "<w:p><w:ins w:id=\"1\" w:author=\"A\"><w:r><w:t>&unknown;</w:t></w:r></w:ins></w:p>",
        ))
        .unwrap();
        let mut malformed_edit = malformed_text.edit();
        assert!(matches!(
            malformed_edit.reject_revision(
                Position::new(0),
                RevisionKind::Insertion,
                Position::new(0),
            ),
            Err(TransactionError::Refused {
                reason: Refusal::RevisionDependency,
                ..
            })
        ));
        assert_eq!(
            malformed_edit.projected().xml_bytes(),
            malformed_text.xml_bytes()
        );
        assert!(malformed_edit.operations.is_empty());

        let ranged = Snapshot::from_xml(document(
            "<w:p><w:bookmarkStart w:id=\"1\" w:name=\"range\"/><w:ins w:id=\"2\" w:author=\"A\"><w:r><w:t>added</w:t></w:r></w:ins><w:bookmarkEnd w:id=\"1\"/></w:p>",
        ))
        .unwrap();
        let mut ranged_edit = ranged.edit();
        assert!(matches!(
            ranged_edit.accept_revision(
                Position::new(0),
                RevisionKind::Insertion,
                Position::new(0),
            ),
            Err(TransactionError::Refused {
                reason: Refusal::RevisionDependency,
                ..
            })
        ));
        assert_eq!(ranged_edit.projected().xml_bytes(), ranged.xml_bytes());
        assert!(ranged_edit.operations.is_empty());

        let empty =
            Snapshot::from_xml(document("<w:p><w:ins w:id=\"1\" w:author=\"A\"/></w:p>")).unwrap();
        let mut noop = empty.edit();
        assert!(matches!(
            noop.accept_revision(Position::new(0), RevisionKind::Insertion, Position::new(1),),
            Err(TransactionError::Refused {
                reason: Refusal::RevisionNotFound,
                ..
            })
        ));
        assert_eq!(noop.projected().xml_bytes(), empty.xml_bytes());
        assert!(noop.operations.is_empty());

        let stale =
            Snapshot::from_xml(document("<w:p><w:r><w:t>different</w:t></w:r></w:p>")).unwrap();
        assert!(matches!(
            noop.commit().unwrap().patch().apply(&stale),
            Err(TransactionError::StaleSource)
        ));
    }

    #[test]
    fn revision_replay_refuses_tampered_after_atomically() {
        let source = Snapshot::from_xml(document(
            "<w:p><w:ins w:id=\"1\" w:author=\"A\"><w:r><w:t>added</w:t></w:r></w:ins></w:p>",
        ))
        .unwrap();
        let mut edit = source.edit();
        edit.accept_revision(Position::new(0), RevisionKind::Insertion, Position::new(0))
            .unwrap();
        let commit = edit.commit().unwrap();
        let Operation::ApplyRevision {
            selector,
            action,
            before,
            ..
        } = commit.patch().operations()[0].clone()
        else {
            panic!("revision action must record ApplyRevision");
        };
        let tampered = Operation::ApplyRevision {
            selector,
            action,
            before,
            after: Arc::new(b"<w:p><w:r><w:t>tampered</w:t></w:r></w:p>".to_vec()),
        };
        let mut replay = source.edit();
        assert!(matches!(
            replay.apply_operation(&tampered),
            Err(TransactionError::SemanticPrecondition)
        ));
        assert_eq!(replay.projected().xml_bytes(), source.xml_bytes());
        assert!(replay.operations.is_empty());
    }

    #[test]
    fn managed_revision_replay_is_refused_before_any_state_change() {
        use litchi_core::{Budget, CancellationSource, ExecutionLimits, Limits, OwnedSource};
        use litchi_opc::constants::{content_type as ct, relationship_type as rt};
        use litchi_opc::{BlobPart, OpcPackage, PackageWriter};

        let xml = document(
            "<w:p><w:ins w:id=\"1\" w:author=\"A\"><w:r><w:t>added</w:t></w:r></w:ins></w:p>",
        );
        // An ordinary unmanaged edit of the same bytes records the operation
        // that the managed edit is then asked to replay.
        let mut unmanaged = Snapshot::from_xml(xml.clone()).unwrap().edit();
        unmanaged
            .accept_revision(Position::new(0), RevisionKind::Insertion, Position::new(0))
            .unwrap();
        let recorded = unmanaged.operations[0].clone();
        assert!(matches!(recorded, Operation::ApplyRevision { .. }));

        let mut opc = OpcPackage::new();
        opc.try_add_part(Box::new(BlobPart::new(
            PackURI::new("/word/document.xml").unwrap(),
            ct::WML_DOCUMENT_MAIN.to_owned(),
            xml,
        )))
        .unwrap();
        opc.relate_to("word/document.xml", rt::OFFICE_DOCUMENT);
        let memory = 1 << 20;
        let budget = Budget::root(
            "docx-managed-revision-replay",
            Limits::new(memory, u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX),
        );
        let (_cancellation_source, cancellation) = CancellationSource::pair();
        let limits = ExecutionLimits::new(
            std::num::NonZeroUsize::MIN,
            std::num::NonZeroUsize::MIN,
            std::num::NonZeroU64::new(memory).unwrap(),
            0,
        )
        .unwrap();
        let package = crate::source_backed::Package::from_read_at_with_execution_context(
            Arc::new(OwnedSource::new(PackageWriter::to_bytes(&opc).unwrap())),
            crate::ReadLimits::default(),
            ExecutionContext::new(budget.clone(), cancellation, limits),
        )
        .unwrap();
        let mut managed = package.edit_document().unwrap();
        assert!(managed.projected.is_managed());
        let projected = managed.projected().xml_bytes().to_vec();
        let reserved = budget.used(Resource::Memory);

        assert!(matches!(
            managed.apply_operation(&recorded),
            Err(TransactionError::Document(crate::Error::UnsafeEdit {
                format: "DOCX",
                operation: "apply_revision",
                ..
            }))
        ));
        assert!(managed.projected.is_managed());
        assert_eq!(managed.projected().xml_bytes(), projected.as_slice());
        assert!(managed.operations.is_empty());
        assert_eq!(budget.used(Resource::Memory), reserved);
    }

    #[test]
    fn revision_actions_resolve_metadata_and_scope_qnames() {
        let foreign_date = Snapshot::from_xml(document(
            "<w:p><w:ins w:id=\"1\" w:author=\"A\" xmlns:x=\"urn:foreign\" x:dateUtc=\"2026-01-01T00:00:00Z\"><w:r><w:t>added</w:t></w:r></w:ins></w:p>",
        ))
        .unwrap();
        let mut foreign_edit = foreign_date.edit();
        assert!(matches!(
            foreign_edit.reject_revision(
                Position::new(0),
                RevisionKind::Insertion,
                Position::new(0),
            ),
            Err(TransactionError::Refused {
                reason: Refusal::RevisionDependency,
                ..
            })
        ));
        assert_eq!(
            foreign_edit.projected().xml_bytes(),
            foreign_date.xml_bytes()
        );
        assert!(foreign_edit.operations.is_empty());

        let aliased_scope = Snapshot::from_xml(document(
            "<w:p><w:ins w:id=\"1\" w:author=\"A\" xmlns:w16du=\"http://schemas.microsoft.com/office/word/2023/wordml/word16du\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:m=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" mc:Ignorable=\"w16du\"><w:r m:Ignorable=\"w16du\"><w:t>added</w:t></w:r></w:ins></w:p>",
        ))
        .unwrap();
        let mut aliased_edit = aliased_scope.edit();
        aliased_edit
            .accept_revision(Position::new(0), RevisionKind::Insertion, Position::new(0))
            .unwrap();
        let aliased_commit = aliased_edit.commit().unwrap();
        let aliased_xml = std::str::from_utf8(aliased_commit.snapshot().xml_bytes()).unwrap();
        assert_eq!(aliased_xml.matches("Ignorable=\"w16du\"").count(), 1);
        assert!(
            aliased_xml.contains(
                "xmlns:m=\"http://schemas.openxmlformats.org/markup-compatibility/2006\""
            )
        );
        assert!(aliased_xml.contains("m:Ignorable=\"w16du\""));
    }

    #[test]
    fn revision_actions_resolve_inherited_word_aliases() {
        let source = Snapshot::from_xml(
            format!(
                "<w:document xmlns:w=\"{WORD}\" xmlns:q=\"{WORD}\"><w:body><w:p><q:ins q:id=\"1\" q:author=\"A\"><q:r><q:t>added</q:t></q:r></q:ins></w:p><w:sectPr/></w:body></w:document>"
            )
            .into_bytes(),
        )
        .unwrap();
        let mut edit = source.edit();
        edit.accept_revision(Position::new(0), RevisionKind::Insertion, Position::new(0))
            .unwrap();
        let commit = edit.commit().unwrap();
        let xml = std::str::from_utf8(commit.snapshot().xml_bytes()).unwrap();
        assert!(!xml.contains("<q:ins"));
        assert!(xml.contains("<q:r><q:t>added</q:t></q:r>"));
    }

    #[test]
    fn revision_actions_preserve_opaque_payload_quoted_attributes_and_inherited_scope() {
        let source = Snapshot::from_xml(document(
            r#"<w:p><w:ins w:id="1" w:author="A" xmlns:x="urn:opaque" xmlns:c="http://schemas.openxmlformats.org/markup-compatibility/2006" c:Ignorable="x" xml:space="preserve" xml:lang="en"><w:r x:vendor="a>b"><x:opaque>  <x:child/>  payload<![CDATA[ & ]]></x:opaque><w:t> x </w:t></w:r></w:ins></w:p>"#,
        )).unwrap();
        let mut edit = source.edit();
        edit.accept_revision(Position::new(0), RevisionKind::Insertion, Position::new(0))
            .unwrap();
        let commit = edit.commit().unwrap();
        let xml = std::str::from_utf8(commit.snapshot().xml_bytes()).unwrap();
        for expected in [
            r#"x:vendor="a>b""#,
            r#"c:Ignorable="x""#,
            r#"xml:space="preserve""#,
            r#"xml:lang="en""#,
            "<x:opaque>  <x:child/>  payload<![CDATA[ & ]]></x:opaque>",
        ] {
            assert!(xml.contains(expected), "missing {expected} in {xml}");
        }
        let durable = commit.patch().to_durable(durable_limits()).unwrap();
        let replayed = source.apply_durable(&durable).unwrap();
        assert_eq!(replayed.xml_bytes(), commit.snapshot().xml_bytes());
        assert_eq!(
            replayed
                .apply_durable(&durable.inverse())
                .unwrap()
                .xml_bytes(),
            source.xml_bytes()
        );
        let visible = litchi_ooxml_common::mce::process_markup_compatibility(
            commit.snapshot().xml_bytes(),
            &litchi_ooxml_common::mce::Capabilities::default(),
            &litchi_ooxml_common::mce::Limits::default(),
        )
        .unwrap();
        assert!(
            !std::str::from_utf8(&visible.xml)
                .unwrap()
                .contains("payload")
        );
        assert_eq!(
            commit
                .patch()
                .inverse()
                .apply(commit.snapshot())
                .unwrap()
                .xml_bytes(),
            source.xml_bytes()
        );
    }

    #[test]
    fn revision_commit_rejects_invalid_utf8_outside_the_selected_revision() {
        for opaque in [false, true] {
            let sibling = if opaque {
                r#"<w:p><w:r><x:opaque xmlns:x="urn:opaque">invalid</x:opaque></w:r></w:p>"#
            } else {
                "<w:p><w:r><w:t>invalid</w:t></w:r></w:p>"
            };
            let mut xml = document(&format!(
                r#"<w:p><w:ins w:id="1" w:author="A"><w:r><w:t>added</w:t></w:r></w:ins></w:p>{sibling}"#
            ));
            let index = xml
                .windows(7)
                .position(|bytes| bytes == b"invalid")
                .unwrap();
            xml[index] = 0xff;
            let source = Snapshot::from_xml(xml).unwrap();
            let mut edit = source.edit();
            edit.accept_revision(Position::new(0), RevisionKind::Insertion, Position::new(0))
                .unwrap();
            assert!(matches!(
                edit.commit(),
                Err(TransactionError::Document(crate::Error::InvalidFormat(_)))
            ));
        }
    }

    #[test]
    fn revision_actions_refuse_invalid_descendant_attributes_and_declarations() {
        for child in [
            r#"<w:r x:bad="&unknown;"/>"#,
            r#"<w:r x:bad="&#0;"/>"#,
            r#"<w:r><x:opaque x:bad="&#xFFFF;"/></w:r>"#,
            r#"<w:r x:bad="raw<value"/>"#,
            r#"<?xml version="1.0"?><w:r/>"#,
        ] {
            let source = Snapshot::from_xml(document(&format!(
                r#"<w:p><w:ins w:id="1" w:author="A" xmlns:x="urn:opaque">{child}</w:ins></w:p>"#
            )))
            .unwrap();
            for action in [RevisionAction::Accept, RevisionAction::Reject] {
                let mut edit = source.edit();
                assert!(
                    edit.apply_revision(
                        RevisionSelector::new(
                            Position::new(0),
                            RevisionKind::Insertion,
                            Position::new(0)
                        ),
                        action
                    )
                    .is_err(),
                    "{child}"
                );
                assert_eq!(edit.projected().xml_bytes(), source.xml_bytes());
                assert!(edit.operations.is_empty());
            }
        }
    }

    #[test]
    fn revision_namespace_expansion_and_deep_untrusted_replay_are_refused_atomically() {
        let body = format!(
            r#"<w:p><w:ins w:id="1" w:author="A" xmlns:x="urn:{}">{}</w:ins></w:p>"#,
            "x".repeat(8192),
            "<w:r/>".repeat(5000)
        );
        let source = Snapshot::from_xml(document(&body)).unwrap();
        let mut edit = source.edit();
        assert!(
            edit.accept_revision(Position::new(0), RevisionKind::Insertion, Position::new(0))
                .is_err()
        );
        assert_eq!(edit.projected().xml_bytes(), source.xml_bytes());
        assert!(edit.operations.is_empty());

        let source = Snapshot::from_xml(document(
            r#"<w:p><w:ins w:id="1" w:author="A"><w:r><w:t>x</w:t></w:r></w:ins></w:p>"#,
        ))
        .unwrap();
        let mut edit = source.edit();
        edit.accept_revision(Position::new(0), RevisionKind::Insertion, Position::new(0))
            .unwrap();
        let commit = edit.commit().unwrap();
        let Operation::ApplyRevision {
            selector,
            action,
            before,
            ..
        } = commit.patch().operations()[0].clone()
        else {
            panic!("revision operation");
        };
        let after = format!(
            "<w:p>{}{}</w:p>",
            "<w:r>".repeat(66_000),
            "</w:r>".repeat(66_000)
        );
        let malformed = Operation::ApplyRevision {
            selector,
            action,
            before,
            after: Arc::new(after.into_bytes()),
        };
        let mut replay = source.edit();
        assert!(matches!(
            replay.apply_operation(&malformed),
            Err(TransactionError::SemanticPrecondition)
        ));
        assert_eq!(replay.projected().xml_bytes(), source.xml_bytes());
        assert!(replay.operations.is_empty());
    }

    #[test]
    fn rich_owner_edits_preserve_the_source_spelling_and_stay_durable_and_reversible() {
        let source = Snapshot::from_xml(
            format!(
                "<w:document xmlns:w=\"{WORD}\">\n  <w:body>\n    <w:p><w:r><w:rPr><w:b/></w:rPr><w:t>direct</w:t><x:keep xmlns:x=\"urn:test\"/></w:r><w:fldSimple w:instr=\" AUTHOR \"><w:r><w:rPr><w:i/></w:rPr><w:t>field</w:t></w:r></w:fldSimple><w:r><w:fldChar w:fldCharType=\"begin\"/></w:r><w:r><w:instrText> DATE </w:instrText></w:r><w:r><w:fldChar w:fldCharType=\"separate\"/></w:r><w:r><w:rPr><w:color w:val=\"FF0000\"/></w:rPr><w:t>complex result</w:t></w:r><w:r><w:fldChar w:fldCharType=\"end\"/></w:r><w:ins w:id=\"7\" w:author=\"A\"><w:r><w:t>added</w:t></w:r></w:ins><w:del w:id=\"8\" w:author=\"A\"><w:r><w:delText>gone &amp; old</w:delText></w:r></w:del><w:sdt><w:sdtPr><w:tag w:val=\"kept\"/></w:sdtPr><w:sdtContent><w:r><w:rPr><w:smallCaps/></w:rPr><w:t>control</w:t></w:r></w:sdtContent></w:sdt></w:p>\n    <w:tbl><w:tr><w:tc><w:tcPr><w:shd w:fill=\"00FF00\"/></w:tcPr><w:p><w:r><w:t>first</w:t></w:r></w:p><w:tbl><w:tr><w:tc><w:p><w:r><w:t>nested kept</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:p><w:r><w:rPr><w:u/></w:rPr><w:t>second</w:t></w:r></w:p></w:tc></w:tr></w:tbl>\n    <w:sectPr/>\n  </w:body>\n</w:document>"
            )
            .into_bytes(),
        )
        .unwrap();
        let mut edit = source.edit();
        edit.replace_run_text(Position::new(0), Position::new(0), " run & entity ")
            .unwrap()
            .replace_simple_field_text(Position::new(0), Position::new(0), "new field")
            .unwrap()
            .replace_complex_field_result_text(
                Position::new(0),
                Position::new(0),
                "new complex result",
            )
            .unwrap()
            .replace_revision_text(
                Position::new(0),
                RevisionKind::Insertion,
                Position::new(0),
                "new insertion",
            )
            .unwrap()
            .replace_content_control_text(Position::new(0), Position::new(0), "new control")
            .unwrap()
            .replace_revision_text(
                Position::new(0),
                RevisionKind::Deletion,
                Position::new(0),
                "new deletion",
            )
            .unwrap()
            .replace_table_cell_paragraph_text(
                Position::new(0),
                Position::new(0),
                Position::new(0),
                Position::new(1),
                "rich cell paragraph",
            )
            .unwrap();
        let commit = edit.commit().unwrap();
        let xml = std::str::from_utf8(commit.snapshot().xml_bytes()).unwrap();
        // Change 0665: publication no longer refuses a non-compact member, so
        // the default policy preserves this source's indentation instead of
        // re-serializing the whole document to remove it.
        assert!(xml.contains("\n  "));
        for retained in [
            "<w:rPr><w:b/></w:rPr>",
            "<x:keep xmlns:x=\"urn:test\"/>",
            "w:instr=\" AUTHOR \"",
            "<w:fldChar w:fldCharType=\"begin\"/>",
            "<w:instrText> DATE </w:instrText>",
            "<w:fldChar w:fldCharType=\"separate\"/>",
            "<w:rPr><w:color w:val=\"FF0000\"/></w:rPr>",
            "<w:t>new complex result</w:t>",
            "<w:fldChar w:fldCharType=\"end\"/>",
            "<w:ins w:id=\"7\" w:author=\"A\">",
            "<w:del w:id=\"8\" w:author=\"A\">",
            "<w:delText>new deletion</w:delText>",
            "<w:tag w:val=\"kept\"/>",
            "<w:rPr><w:smallCaps/></w:rPr>",
            "<w:t>new control</w:t>",
            "<w:t>nested kept</w:t>",
            "<w:rPr><w:u/></w:rPr>",
            "<w:t xml:space=\"preserve\"> run &amp; entity </w:t>",
        ] {
            assert!(xml.contains(retained), "missing retained XML: {retained}");
        }

        let durable = commit.patch().to_durable(durable_limits()).unwrap();
        let wire = durable.to_deterministic_json().unwrap();
        let decoded =
            litchi_core::patch::Patch::<litchi_core::patch::Reversible>::from_deterministic_json(
                &wire,
                durable_limits(),
            )
            .unwrap();
        let applied = source.apply_durable(&decoded).unwrap();
        assert_eq!(applied.xml_bytes(), commit.snapshot().xml_bytes());
        assert_eq!(
            applied
                .apply_durable(&decoded.inverse())
                .unwrap()
                .xml_bytes(),
            source.xml_bytes()
        );
    }

    #[test]
    fn three_way_planning_keeps_disjoint_rich_owners_and_resolves_overlap() {
        let source = Snapshot::from_xml(document(
            "<w:p><w:r><w:t>run</w:t></w:r><w:fldSimple w:instr=\" AUTHOR \"><w:r><w:t>field</w:t></w:r></w:fldSimple></w:p>",
        ))
        .unwrap();
        let limits = CompositionLimits::new(8, 8, 32, 8);

        let mut run = source.edit();
        run.replace_run_text(Position::new(0), Position::new(0), "changed run")
            .unwrap();
        let mut left = source.compose(limits);
        left.join(run.prepare(limits, "run").unwrap()).unwrap();
        let mut field = source.edit();
        field
            .replace_simple_field_text(Position::new(0), Position::new(0), "changed field")
            .unwrap();
        let mut right = source.compose(limits);
        right.join(field.prepare(limits, "field").unwrap()).unwrap();
        let plan = source.plan_three_way(left, right).unwrap();
        assert!(plan.is_clean());
        assert_eq!(
            plan.finish()
                .unwrap()
                .commit()
                .unwrap()
                .patch()
                .operations()
                .len(),
            2
        );

        let branch = |identifier: &str, value: &str| {
            let mut edit = source.edit();
            edit.replace_simple_field_text(Position::new(0), Position::new(0), value)
                .unwrap();
            let mut composition = source.compose(limits);
            composition
                .join(edit.prepare(limits, identifier).unwrap())
                .unwrap();
            composition
        };
        let mut conflict = source
            .plan_three_way(branch("left", "left"), branch("right", "right"))
            .unwrap();
        assert!(!conflict.is_clean());
        assert!(!conflict.conflicts().is_empty());
        conflict.resolve(MergeChoice::Left);
        let merged = conflict.finish().unwrap().commit().unwrap();
        assert!(
            std::str::from_utf8(merged.snapshot().xml_bytes())
                .unwrap()
                .contains(">left</w:t>")
        );
    }

    #[test]
    fn durable_patch_is_deterministic_stale_checked_and_reversible() {
        let source = Snapshot::from_xml(document(
            "<w:p><w:r><w:t>one</w:t></w:r><w:r><w:t> two</w:t></w:r></w:p><w:tbl><w:tr><w:tc><w:p><w:r><w:t>cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>",
        ))
        .unwrap();
        let mut edit = source.edit();
        edit.replace_paragraph_text(Position::new(0), "durable text")
            .unwrap()
            .replace_table_cell_text(
                Position::new(0),
                Position::new(0),
                Position::new(0),
                "durable cell",
            )
            .unwrap();
        let commit = edit.commit().unwrap();
        let durable = commit.patch().to_durable(durable_limits()).unwrap();
        let first = durable.to_deterministic_json().unwrap();
        let second = durable.to_deterministic_json().unwrap();
        assert_eq!(first, second);
        let decoded =
            litchi_core::patch::Patch::<litchi_core::patch::Reversible>::from_deterministic_json(
                &first,
                durable_limits(),
            )
            .unwrap();

        let applied = source.apply_durable(&decoded).unwrap();
        assert_eq!(applied.xml_bytes(), commit.snapshot().xml_bytes());
        let restored = applied.apply_durable(&decoded.inverse()).unwrap();
        assert_eq!(restored.xml_bytes(), source.xml_bytes());
        let stale = Snapshot::from_xml(document("<w:p><w:r><w:t>other</w:t></w:r></w:p>")).unwrap();
        assert!(matches!(
            stale.apply_durable(&decoded),
            Err(TransactionError::StaleSource)
        ));
    }

    #[test]
    fn disjoint_composition_and_bounded_history_are_deterministic() {
        let source = Snapshot::from_xml(document(
            "<w:p><w:r><w:t>body</w:t></w:r></w:p><w:tbl><w:tr><w:tc><w:p><w:r><w:t>cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>",
        ))
        .unwrap();
        let limits = CompositionLimits::new(8, 8, 32, 8);
        let mut paragraph = source.edit();
        paragraph
            .replace_paragraph_text(Position::new(0), "changed body")
            .unwrap();
        let paragraph = paragraph.prepare(limits, "b-paragraph").unwrap();
        let mut cell = source.edit();
        cell.replace_table_cell_text(
            Position::new(0),
            Position::new(0),
            Position::new(0),
            "changed cell",
        )
        .unwrap();
        let cell = cell.prepare(limits, "a-cell").unwrap();
        let mut composition = source.compose(limits);
        composition.join(paragraph).unwrap().join(cell).unwrap();
        let commit = composition.commit().unwrap();
        assert_eq!(commit.patch().operations().len(), 2);
        assert!(matches!(
            commit.patch().operations()[0],
            Operation::ReplaceCellText { .. }
        ));

        let mut left = source.edit();
        left.replace_paragraph_text(Position::new(0), "left")
            .unwrap();
        let left = left.prepare(limits, "left").unwrap();
        let mut right = source.edit();
        right
            .replace_paragraph_text(Position::new(0), "right")
            .unwrap();
        let right = right.prepare(limits, "right").unwrap();
        let mut overlap = source.compose(limits);
        overlap.join(left).unwrap();
        assert!(matches!(
            overlap.join(right).unwrap_err().failure(),
            SubEditJoinFailure::Overlap(_)
        ));

        let mut indexed = source.edit();
        indexed
            .replace_run_text(Position::new(0), Position::new(0), "indexed")
            .unwrap();
        let indexed = indexed.prepare(limits, "indexed-owner").unwrap();
        let mut appended = source.edit();
        appended
            .insert_paragraph(Position::new(source.paragraph_count()), "appended")
            .unwrap();
        let appended = appended.prepare(limits, "append-only").unwrap();
        let mut append_composition = source.compose(limits);
        append_composition
            .join(indexed)
            .unwrap()
            .join(appended)
            .unwrap();
        append_composition.commit().unwrap();

        let mut indexed = source.edit();
        indexed
            .replace_run_text(Position::new(0), Position::new(0), "indexed")
            .unwrap();
        let indexed = indexed.prepare(limits, "indexed-owner").unwrap();
        let mut prefixed = source.edit();
        prefixed
            .insert_paragraph(Position::new(0), "prefixed")
            .unwrap();
        let prefixed = prefixed.prepare(limits, "prefix-insert").unwrap();
        let mut prefix_overlap = source.compose(limits);
        prefix_overlap.join(indexed).unwrap();
        assert!(matches!(
            prefix_overlap.join(prefixed).unwrap_err().failure(),
            SubEditJoinFailure::Overlap(_)
        ));

        let budget = u64::try_from(commit.snapshot().xml_bytes().len()).unwrap();
        let mut history = source.history(HistoryLimits::new(1, budget));
        history.record(commit).unwrap();
        assert!(history.can_undo());
        assert!(history.undo());
        assert_eq!(history.current().xml_bytes(), source.xml_bytes());
        assert!(history.redo());
    }

    #[test]
    fn composition_owner_effects_refuse_ancestor_overlaps_and_cap_atomically() {
        let source = Snapshot::from_xml(document("<w:p/>")).unwrap();
        let limits = CompositionLimits::new(8, 16, 32, 16);

        let mut direct_control = source.edit();
        direct_control
            .operations
            .push(Operation::ReplaceContentControlText {
                paragraph: Position::new(0),
                control: Position::new(0),
                before: "before".into(),
                after: "direct".into(),
            });
        let direct_control = direct_control.prepare(limits, "direct-control").unwrap();

        let mut nested_control = source.edit();
        nested_control
            .operations
            .push(Operation::ReplaceNestedContentControlText {
                paragraph: Position::new(0),
                controls: Arc::from([Position::new(0), Position::new(0)]),
                before: "before".into(),
                after: "nested".into(),
            });
        let nested_control = nested_control.prepare(limits, "nested-control").unwrap();

        let mut controls = source.compose(limits);
        controls.join(direct_control).unwrap();
        assert!(matches!(
            controls.join(nested_control).unwrap_err().failure(),
            SubEditJoinFailure::Overlap(_)
        ));
        assert_eq!(controls.len(), 1);

        let cell_source = Snapshot::from_xml(document(
            "<w:p/><w:tbl><w:tr><w:tc><w:p><w:r><w:t>cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>",
        ))
        .unwrap();
        let mut broad_cell = cell_source.edit();
        broad_cell
            .replace_table_cell_text(
                Position::new(0),
                Position::new(0),
                Position::new(0),
                "direct cell",
            )
            .unwrap();
        let broad_cell = broad_cell.prepare(limits, "broad-cell").unwrap();

        let mut nested_cell = cell_source.edit();
        nested_cell
            .replace_nested_table_cell_paragraph_text(
                &[TableCellAddress::new(
                    Position::new(0),
                    Position::new(0),
                    Position::new(0),
                )],
                Position::new(0),
                "nested cell",
            )
            .unwrap();
        let nested_cell = nested_cell.prepare(limits, "nested-cell").unwrap();

        let mut cells = cell_source.compose(limits);
        cells.join(broad_cell).unwrap();
        assert!(matches!(
            cells.join(nested_cell).unwrap_err().failure(),
            SubEditJoinFailure::Overlap(_)
        ));
        assert_eq!(cells.len(), 1);

        let capped = CompositionLimits::new(1, 16, 32, 16);
        let mut first = source.edit();
        first.operations.push(Operation::ReplaceParagraphText {
            position: Position::new(0),
            before: "before".into(),
            after: "first".into(),
        });
        let first = first.prepare(capped, "first").unwrap();
        let mut second = source.edit();
        second.operations.push(Operation::ReplaceRunText {
            paragraph: Position::new(0),
            run: Position::new(0),
            before: "before".into(),
            after: "second".into(),
        });
        let second = second.prepare(capped, "second").unwrap();
        let mut bounded = source.compose(capped);
        bounded.join(first).unwrap();
        let error = bounded.join(second).unwrap_err();
        assert!(matches!(error.failure(), SubEditJoinFailure::Limit(_)));
        assert_eq!(bounded.len(), 1);
        assert_eq!(error.into_rejected().identifier(), "second");
    }

    #[test]
    fn structural_run_text_is_native_and_complex_cells_remain_atomic_refusals() {
        let source = Snapshot::from_xml(document(
            "<w:p><w:r><w:t>safe</w:t><w:br/></w:r></w:p><w:tbl><w:tr><w:tc><w:p><w:r><w:t>one</w:t></w:r></w:p><w:p><w:r><w:t>two</w:t></w:r></w:p></w:tc></w:tr></w:tbl>",
        ))
        .unwrap();
        let mut edit = source.edit();
        edit.replace_run_text(
            Position::new(0),
            Position::new(0),
            "safe\tline\nsoft\u{2011}hyphen\u{00ad}",
        )
        .unwrap();
        assert!(
            edit.replace_table_cell_text(
                Position::new(0),
                Position::new(0),
                Position::new(0),
                "unsafe flatten",
            )
            .is_err()
        );
        let xml = std::str::from_utf8(edit.projected().xml_bytes()).unwrap();
        for structural in [
            "<w:tab/>",
            "<w:br/>",
            "<w:noBreakHyphen/>",
            "<w:softHyphen/>",
        ] {
            assert!(xml.contains(structural));
        }
        assert!(xml.contains("<w:r>"));
        assert_ne!(edit.projected().xml_bytes(), source.xml_bytes());
    }

    #[test]
    fn nested_controls_and_tables_are_durable_exact_and_path_checked() {
        let source = Snapshot::from_xml(
            format!(
                "<w:document xmlns:w=\"{WORD}\">\n<w:body><w:p><w:sdt><w:sdtPr><w:tag w:val=\"outer-inline\"/></w:sdtPr><w:sdtContent><w:sdt><w:sdtPr><w:tag w:val=\"inner-inline\"/></w:sdtPr><w:sdtContent><w:r><w:rPr><w:b/></w:rPr><w:t>inline old</w:t></w:r></w:sdtContent></w:sdt></w:sdtContent></w:sdt></w:p><w:sdt><w:sdtPr><w:alias w:val=\"outer-block\"/></w:sdtPr><w:sdtContent><w:sdt><w:sdtPr><w:alias w:val=\"inner-block\"/></w:sdtPr><w:sdtContent><w:p><w:pPr><w:keepNext/></w:pPr><w:r><w:t>block old</w:t></w:r></w:p></w:sdtContent></w:sdt></w:sdtContent></w:sdt><w:tbl><w:tblPr><w:tblStyle w:val=\"Outer\"/></w:tblPr><w:tr><w:tc><w:p><w:r><w:t>outer kept</w:t></w:r></w:p><w:tbl><w:tblPr><w:tblStyle w:val=\"Inner\"/></w:tblPr><w:tr><w:tc><w:tcPr><w:shd w:fill=\"00FF00\"/></w:tcPr><w:p><w:r><w:rPr><w:i/></w:rPr><w:t>nested old</w:t></w:r></w:p></w:tc></w:tr></w:tbl></w:tc></w:tr></w:tbl><w:sectPr/></w:body></w:document>"
            )
            .into_bytes(),
        )
        .unwrap();
        assert_eq!(source.block_content_control_count(), 1);
        let controls = [Position::new(0), Position::new(0)];
        let table_path = [
            TableCellAddress::new(Position::new(0), Position::new(0), Position::new(0)),
            TableCellAddress::new(Position::new(0), Position::new(0), Position::new(0)),
        ];
        let mut edit = source.edit();
        assert!(
            edit.replace_nested_content_control_text(Position::new(0), &[], "refused")
                .is_err()
        );
        assert_eq!(edit.projected().xml_bytes(), source.xml_bytes());
        edit.replace_nested_content_control_text(Position::new(0), &controls, "inline new")
            .unwrap()
            .replace_block_content_control_paragraph_text(&controls, Position::new(0), "block new")
            .unwrap()
            .replace_nested_table_cell_paragraph_text(&table_path, Position::new(0), "nested new")
            .unwrap();
        let commit = edit.commit().unwrap();
        let xml = std::str::from_utf8(commit.snapshot().xml_bytes()).unwrap();
        for retained in [
            "w:val=\"outer-inline\"",
            "w:val=\"inner-inline\"",
            "<w:rPr><w:b/></w:rPr>",
            "w:val=\"outer-block\"",
            "w:val=\"inner-block\"",
            "<w:pPr><w:keepNext/></w:pPr>",
            "<w:tblStyle w:val=\"Outer\"/>",
            "<w:tblStyle w:val=\"Inner\"/>",
            "<w:shd w:fill=\"00FF00\"/>",
            "<w:rPr><w:i/></w:rPr>",
            "<w:t>outer kept</w:t>",
            "<w:t>inline new</w:t>",
            "<w:t>block new</w:t>",
            "<w:t>nested new</w:t>",
        ] {
            assert!(xml.contains(retained), "missing retained XML: {retained}");
        }
        // Change 0665: the default policy preserves the source's line breaks
        // rather than re-serializing the whole document to remove them.
        assert!(xml.contains("\n"));

        let durable = commit.patch().to_durable(durable_limits()).unwrap();
        let applied = source.apply_durable(&durable).unwrap();
        assert_eq!(applied.xml_bytes(), commit.snapshot().xml_bytes());
        assert_eq!(
            applied
                .apply_durable(&durable.inverse())
                .unwrap()
                .xml_bytes(),
            source.xml_bytes()
        );
        assert_eq!(
            commit
                .patch()
                .inverse()
                .apply(commit.snapshot())
                .unwrap()
                .xml_bytes(),
            source.xml_bytes()
        );

        let composition_limits = CompositionLimits::new(8, 8, 32, 8);
        let mut inline = source.edit();
        inline
            .replace_nested_content_control_text(Position::new(0), &controls, "inline new")
            .unwrap();
        let mut block = source.edit();
        block
            .replace_block_content_control_paragraph_text(&controls, Position::new(0), "block new")
            .unwrap();
        let mut table = source.edit();
        table
            .replace_nested_table_cell_paragraph_text(&table_path, Position::new(0), "nested new")
            .unwrap();
        let mut composition = source.compose(composition_limits);
        composition
            .join(block.prepare(composition_limits, "block").unwrap())
            .unwrap()
            .join(inline.prepare(composition_limits, "inline").unwrap())
            .unwrap()
            .join(table.prepare(composition_limits, "table").unwrap())
            .unwrap();
        assert_eq!(
            composition.commit().unwrap().snapshot().xml_bytes(),
            commit.snapshot().xml_bytes()
        );

        let history_budget = u64::try_from(commit.snapshot().xml_bytes().len()).unwrap();
        let mut history = source.history(HistoryLimits::new(2, history_budget));
        history.record(commit).unwrap();
        assert!(history.undo());
        assert_eq!(history.current().xml_bytes(), source.xml_bytes());
        assert!(history.redo());
    }

    #[test]
    fn composite_nested_hyperlinks_preserve_relationship_ownership() {
        let source = Snapshot::from_xml(
            format!(
                "<w:document xmlns:w=\"{WORD}\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><w:body><w:p><w:sdt><w:sdtPr><w:tag w:val=\"outer-inline-link\"/></w:sdtPr><w:sdtContent><w:sdt><w:sdtPr><w:tag w:val=\"inner-inline-link\"/></w:sdtPr><w:sdtContent><w:hyperlink r:id=\"inlineRel\" w:tooltip=\"inline tip\"><w:r><w:rPr><w:u/></w:rPr><w:t>inline link</w:t></w:r></w:hyperlink></w:sdtContent></w:sdt></w:sdtContent></w:sdt></w:p><w:sdt><w:sdtPr><w:alias w:val=\"outer-block-link\"/></w:sdtPr><w:sdtContent><w:sdt><w:sdtPr><w:alias w:val=\"inner-block-link\"/></w:sdtPr><w:sdtContent><w:p><w:hyperlink r:id=\"blockRel\" w:tooltip=\"block tip\"><w:r><w:t>block link</w:t></w:r></w:hyperlink></w:p></w:sdtContent></w:sdt></w:sdtContent></w:sdt><w:tbl><w:tr><w:tc><w:p><w:r><w:t>outer</w:t></w:r></w:p><w:tbl><w:tr><w:tc><w:p><w:hyperlink r:id=\"tableRel\" w:tooltip=\"table tip\"><w:r><w:rPr><w:i/></w:rPr><w:t>table link</w:t></w:r></w:hyperlink></w:p></w:tc></w:tr></w:tbl></w:tc></w:tr></w:tbl><w:sectPr/></w:body></w:document>"
            )
            .into_bytes(),
        )
        .unwrap();
        let controls = [Position::new(0), Position::new(0)];
        let table_path = [
            TableCellAddress::new(Position::new(0), Position::new(0), Position::new(0)),
            TableCellAddress::new(Position::new(0), Position::new(0), Position::new(0)),
        ];
        let mut edit = source.edit();
        edit.replace_nested_content_control_hyperlink_text(
            Position::new(0),
            &controls,
            Position::new(0),
            "inline changed",
        )
        .unwrap()
        .replace_block_content_control_paragraph_hyperlink_text(
            &controls,
            Position::new(0),
            Position::new(0),
            "block changed",
        )
        .unwrap()
        .replace_nested_table_cell_paragraph_hyperlink_text(
            &table_path,
            Position::new(0),
            Position::new(0),
            "table changed",
        )
        .unwrap();
        let commit = edit.commit().unwrap();
        let xml = std::str::from_utf8(commit.snapshot().xml_bytes()).unwrap();
        for retained in [
            "r:id=\"inlineRel\"",
            "w:tooltip=\"inline tip\"",
            "<w:rPr><w:u/></w:rPr>",
            "r:id=\"blockRel\"",
            "w:tooltip=\"block tip\"",
            "r:id=\"tableRel\"",
            "w:tooltip=\"table tip\"",
            "<w:rPr><w:i/></w:rPr>",
            "<w:t>inline changed</w:t>",
            "<w:t>block changed</w:t>",
            "<w:t>table changed</w:t>",
        ] {
            assert!(xml.contains(retained), "missing retained XML: {retained}");
        }

        let durable = commit.patch().to_durable(durable_limits()).unwrap();
        let applied = source.apply_durable(&durable).unwrap();
        assert_eq!(applied.xml_bytes(), commit.snapshot().xml_bytes());
        assert_eq!(
            applied
                .apply_durable(&durable.inverse())
                .unwrap()
                .xml_bytes(),
            source.xml_bytes()
        );

        let limits = CompositionLimits::new(8, 8, 32, 8);
        let mut wide = source.edit();
        assert!(
            wide.replace_nested_content_control_text(
                Position::new(0),
                &controls,
                "cannot flatten hyperlink",
            )
            .is_err()
        );
        assert_eq!(wide.projected().xml_bytes(), source.xml_bytes());
        let mut left_edit = source.edit();
        left_edit
            .replace_nested_content_control_hyperlink_text(
                Position::new(0),
                &controls,
                Position::new(0),
                "left",
            )
            .unwrap();
        let mut right_edit = source.edit();
        right_edit
            .replace_nested_table_cell_paragraph_hyperlink_text(
                &table_path,
                Position::new(0),
                Position::new(0),
                "right",
            )
            .unwrap();
        let mut composition = source.compose(limits);
        composition
            .join(left_edit.prepare(limits, "left").unwrap())
            .unwrap()
            .join(right_edit.prepare(limits, "right").unwrap())
            .unwrap();
        assert_eq!(composition.commit().unwrap().patch().operations().len(), 2);
    }

    #[test]
    fn cross_paragraph_hyperlink_batches_are_atomic_durable_and_non_crossing() {
        let source = Snapshot::from_xml(
            format!(
                "<w:document xmlns:w=\"{WORD}\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><w:body><w:p><w:hyperlink r:id=\"body0\" w:tooltip=\"body zero\"><w:r><w:rPr><w:b/></w:rPr><w:t>body zero old</w:t></w:r></w:hyperlink></w:p><w:p><w:hyperlink r:id=\"body1\" w:tooltip=\"body one\"><w:r><w:t>body one old</w:t></w:r></w:hyperlink></w:p><w:sdt><w:sdtPr><w:alias w:val=\"outer-block\"/></w:sdtPr><w:sdtContent><w:sdt><w:sdtPr><w:alias w:val=\"inner-block\"/></w:sdtPr><w:sdtContent><w:p><w:hyperlink r:id=\"block0\" w:tooltip=\"block zero\"><w:r><w:t>block zero old</w:t></w:r></w:hyperlink></w:p><w:p><w:hyperlink r:id=\"block1\" w:tooltip=\"block one\"><w:r><w:rPr><w:i/></w:rPr><w:t>block one old</w:t></w:r></w:hyperlink></w:p></w:sdtContent></w:sdt></w:sdtContent></w:sdt><w:tbl><w:tr><w:tc><w:p><w:r><w:t>outer retained</w:t></w:r></w:p><w:tbl><w:tr><w:tc><w:tcPr><w:shd w:fill=\"00FF00\"/></w:tcPr><w:p><w:hyperlink r:id=\"table0\" w:tooltip=\"table zero\"><w:r><w:t>table zero old</w:t></w:r></w:hyperlink></w:p><w:p><w:hyperlink r:id=\"table1\" w:tooltip=\"table one\"><w:r><w:rPr><w:u/></w:rPr><w:t>table one old</w:t></w:r></w:hyperlink></w:p></w:tc></w:tr></w:tbl></w:tc></w:tr></w:tbl><w:sectPr/></w:body></w:document>"
            )
            .into_bytes(),
        )
        .unwrap();
        let controls = [Position::new(0), Position::new(0)];
        let table_path = [
            TableCellAddress::new(Position::new(0), Position::new(0), Position::new(0)),
            TableCellAddress::new(Position::new(0), Position::new(0), Position::new(0)),
        ];
        let replacements = |prefix: &str| {
            [
                HyperlinkTextReplacement::new(
                    ParagraphHyperlinkAddress::new(Position::new(0), Position::new(0)),
                    format!("{prefix} zero changed"),
                ),
                HyperlinkTextReplacement::new(
                    ParagraphHyperlinkAddress::new(Position::new(1), Position::new(0)),
                    format!("{prefix} one changed"),
                ),
            ]
        };

        let mut edit = source.edit();
        edit.replace_body_hyperlink_texts(&replacements("body"))
            .unwrap()
            .replace_block_content_control_paragraph_hyperlink_texts(
                &controls,
                &replacements("block"),
            )
            .unwrap()
            .replace_nested_table_cell_paragraph_hyperlink_texts(
                &table_path,
                &replacements("table"),
            )
            .unwrap();
        let commit = edit.commit().unwrap();
        assert_eq!(commit.patch().operations().len(), 6);
        let xml = std::str::from_utf8(commit.snapshot().xml_bytes()).unwrap();
        for retained in [
            "r:id=\"body0\"",
            "w:tooltip=\"body one\"",
            "<w:rPr><w:b/></w:rPr>",
            "r:id=\"block0\"",
            "w:tooltip=\"block one\"",
            "<w:rPr><w:i/></w:rPr>",
            "r:id=\"table0\"",
            "w:tooltip=\"table one\"",
            "<w:rPr><w:u/></w:rPr>",
            "<w:shd w:fill=\"00FF00\"/>",
            "<w:t>outer retained</w:t>",
            "<w:t>body zero changed</w:t>",
            "<w:t>body one changed</w:t>",
            "<w:t>block zero changed</w:t>",
            "<w:t>block one changed</w:t>",
            "<w:t>table zero changed</w:t>",
            "<w:t>table one changed</w:t>",
        ] {
            assert!(xml.contains(retained), "missing retained XML: {retained}");
        }

        let durable = commit.patch().to_durable(durable_limits()).unwrap();
        let wire = durable.to_deterministic_json().unwrap();
        let decoded =
            litchi_core::patch::Patch::<litchi_core::patch::Reversible>::from_deterministic_json(
                &wire,
                durable_limits(),
            )
            .unwrap();
        let applied = source.apply_durable(&decoded).unwrap();
        assert_eq!(applied.xml_bytes(), commit.snapshot().xml_bytes());
        assert_eq!(
            applied
                .apply_durable(&decoded.inverse())
                .unwrap()
                .xml_bytes(),
            source.xml_bytes()
        );

        let limits = CompositionLimits::new(8, 8, 32, 8);
        let mut body = source.edit();
        body.replace_body_hyperlink_texts(&replacements("body"))
            .unwrap();
        let mut block = source.edit();
        block
            .replace_block_content_control_paragraph_hyperlink_texts(
                &controls,
                &replacements("block"),
            )
            .unwrap();
        let mut table = source.edit();
        table
            .replace_nested_table_cell_paragraph_hyperlink_texts(
                &table_path,
                &replacements("table"),
            )
            .unwrap();
        let mut composition = source.compose(limits);
        composition
            .join(table.prepare(limits, "c-table").unwrap())
            .unwrap()
            .join(block.prepare(limits, "b-block").unwrap())
            .unwrap()
            .join(body.prepare(limits, "a-body").unwrap())
            .unwrap();
        assert_eq!(
            composition.commit().unwrap().snapshot().xml_bytes(),
            commit.snapshot().xml_bytes()
        );

        let history_budget = u64::try_from(commit.snapshot().xml_bytes().len()).unwrap();
        let mut history = source.history(HistoryLimits::new(2, history_budget));
        history.record(commit).unwrap();
        assert!(history.undo());
        assert_eq!(history.current().xml_bytes(), source.xml_bytes());
        assert!(history.redo());

        let duplicate = [
            HyperlinkTextReplacement::new(
                ParagraphHyperlinkAddress::new(Position::new(0), Position::new(0)),
                "first",
            ),
            HyperlinkTextReplacement::new(
                ParagraphHyperlinkAddress::new(Position::new(0), Position::new(0)),
                "duplicate",
            ),
        ];
        let mut refused = source.edit();
        assert!(matches!(
            refused.replace_body_hyperlink_texts(&[]),
            Err(TransactionError::Refused {
                reason: Refusal::AmbiguousCompositeSelector,
                ..
            })
        ));
        assert!(matches!(
            refused.replace_body_hyperlink_texts(&duplicate),
            Err(TransactionError::Refused {
                reason: Refusal::AmbiguousCompositeSelector,
                ..
            })
        ));
        assert_eq!(refused.projected().xml_bytes(), source.xml_bytes());

        let late_failure = [
            HyperlinkTextReplacement::new(
                ParagraphHyperlinkAddress::new(Position::new(0), Position::new(0)),
                "would otherwise change",
            ),
            HyperlinkTextReplacement::new(
                ParagraphHyperlinkAddress::new(Position::new(1), Position::new(1)),
                "missing",
            ),
        ];
        assert!(matches!(
            refused.replace_body_hyperlink_texts(&late_failure),
            Err(TransactionError::Refused {
                reason: Refusal::HyperlinkNotFound,
                ..
            })
        ));
        assert_eq!(refused.projected().xml_bytes(), source.xml_bytes());
    }
}
