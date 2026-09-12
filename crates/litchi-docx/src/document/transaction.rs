//! Source-preserving main-document snapshots, edits, and reversible patches.

mod durable;

use std::mem::size_of;
use std::sync::Arc;

use litchi_core::{ExecutionContext, Position, Reservation, Resource, SourceVersion};
use litchi_opc::{PackURI, PartData, SourceLineage, SourceXmlPart};
use quick_xml::events::Event;
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;
use quick_xml::{Reader, XmlVersion};
use thiserror::Error;

use crate::namespace::{
    STRICT_WORDPROCESSINGML_NAMESPACE, WORDPROCESSINGML_NAMESPACE, is_wordprocessing_namespace,
};
use crate::paragraph::Paragraph;

pub(crate) use durable::durable_transfer_operations;
pub use durable::{Composition, History, JoinError, PreparedEdit, ThreeWayError, ThreeWayPlan};
pub use litchi_core::patch::{
    CompositionLimits, HistoryLimits, MergeChoice, SubEditConflict, SubEditJoinFailure,
    ThreeWayMergeFailure,
};

pub(crate) const MAX_DOCUMENT_XML_BYTES: usize = 32 * 1024 * 1024;
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
            Self::ContentControlNotFound => "direct inline content control was not found",
            Self::CellNotFound => "table cell was not found",
            Self::StructuralText => "text requires structural WordprocessingML elements",
        })
    }
}

/// Direct tracked-revision wrapper selected for inert text replacement.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum RevisionKind {
    /// `w:ins` tracked insertion content.
    Insertion,
    /// `w:del` tracked deletion content.
    Deletion,
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
    tables: Arc<[Range]>,
    block_controls: Arc<[Range]>,
    content_end: u32,
    conformance: Conformance,
    admission: Option<Arc<ManagedAdmission>>,
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
            tables: layout.tables.into(),
            block_controls: layout.block_controls.into(),
            content_end: layout.content_end,
            conformance: layout.conformance,
            admission: None,
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
            tables: layout.tables.into(),
            block_controls: layout.block_controls.into(),
            content_end: layout.content_end,
            conformance: layout.conformance,
            admission: None,
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
            tables: layout.tables.into(),
            block_controls: layout.block_controls.into(),
            content_end: layout.content_end,
            conformance: layout.conformance,
            admission: Some(admission),
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
            tables: layout.tables.into(),
            block_controls: layout.block_controls.into(),
            content_end: layout.content_end,
            conformance: layout.conformance,
            admission: Some(admission),
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
            tables: layout.tables.into(),
            block_controls: layout.block_controls.into(),
            content_end: layout.content_end,
            conformance: layout.conformance,
            admission: Some(admission),
        })
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
                XmlStorage::Owned(xml) | XmlStorage::OwnedWithIdentity { xml, .. } => Some(
                    Paragraph::from_arc_range(Arc::clone(xml), range.start, range.length),
                ),
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
            .filter_map(|range| match &self.xml {
                XmlStorage::Owned(xml) | XmlStorage::OwnedWithIdentity { xml, .. } => Some(
                    Paragraph::from_arc_range(Arc::clone(xml), range.start, range.length),
                ),
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
        }
    }

    fn managed_context(&self) -> Option<ExecutionContext> {
        self.admission
            .as_ref()
            .map(|admission| admission.context.clone())
    }

    fn same_source(&self, other: &Self) -> bool {
        if self.xml.bytes() != other.xml.bytes() {
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

/// A staged main-document edit.
#[derive(Debug)]
pub struct Edit {
    base: Snapshot,
    projected: Snapshot,
    operations: Vec<Operation>,
    replacement_text_bytes: usize,
    operation_admission: Option<Arc<OperationAdmission>>,
    operation_string_admission: Option<Arc<StringAdmission>>,
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
        let candidate = self.projected.with_rewritten_xml(xml)?;
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
        let projected = candidate
            .projected
            .with_rewritten_xml(replace_ranges(candidate.projected.xml_bytes(), &ranges)?)?;
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
        let projected = if self.base.same_source(&self.projected) {
            self.projected
        } else if self.base.is_source_backed() && self.projected.is_source_backed() {
            // Source-authorized edits are already validated and retain the
            // checked splice/output reservation. Running them through the
            // authored compacting writer would detach that owner and lose
            // the exact source proof required by publication.
            self.projected
        } else {
            let source = std::str::from_utf8(self.projected.xml_bytes()).map_err(|error| {
                crate::Error::InvalidFormat(format!(
                    "changed main-document XML is not UTF-8: {error}"
                ))
            })?;
            let compact = crate::writer::doc::compact_changed_document_xml(source)?;
            self.projected.with_rewritten_xml(compact.into_bytes())?
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

struct Layout {
    paragraphs: Vec<Range>,
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

fn event_namespace_binding_count(event: &Event<'_>) -> usize {
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

fn scan_document_with_context(
    xml: &[u8],
    context: Option<&ExecutionContext>,
) -> TransactionResult<Layout> {
    let mut reader = NsReader::from_reader(xml);
    let mut paragraphs = Vec::new();
    let mut tables = Vec::new();
    let mut block_controls = Vec::new();
    let mut body_depth = None;
    let mut body_end = None;
    let mut final_section_start = None;
    let mut pending = None::<(bool, bool, bool, bool, usize)>;
    let mut conformance = None;
    let mut saw_document = false;
    let mut depth = 0usize;
    let mut nodes = 0usize;
    // Namespace accounting is only needed for the managed work charge. Keep
    // this state out of ordinary/unmanaged layout scans so they retain their
    // prior allocation and parsing behavior.
    let mut namespace_bindings = context.map(|_| 2usize);
    let mut namespace_scopes = context.map(|_| Vec::<usize>::new());

    loop {
        if let Some(context) = context {
            context.check().map_err(managed_execution)?;
        }
        let event_start =
            usize::try_from(reader.buffer_position()).map_err(|_conversion_error| {
                crate::Error::InvalidFormat("document offset does not fit usize".into())
            })?;
        let raw_event = reader
            .read_event()
            .map_err(|error| crate::Error::Xml(error.to_string()))?
            .into_owned();
        let event_end = usize::try_from(reader.buffer_position()).map_err(|_conversion_error| {
            crate::Error::InvalidFormat("document offset does not fit usize".into())
        })?;
        let event_bytes = event_end.saturating_sub(event_start).max(1);
        let (event_namespace_bindings, active_namespace_bindings) = if namespace_bindings.is_some()
        {
            let event_namespace_bindings = event_namespace_binding_count(&raw_event);
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
        // The event is owned, so namespace resolution can borrow the reader's
        // resolver.  Cloning the resolver here would copy all in-scope
        // namespace bindings once per event and turn a linear scan into a
        // quadratic allocation path for adversarial namespace-heavy input.
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
                    pending = Some((is_paragraph, is_table, is_control, is_section, event_start));
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
                if let Some((is_paragraph, is_table, is_control, is_section, start)) = pending
                    && body_depth.is_some_and(|body| depth == body + 1)
                {
                    if is_paragraph {
                        paragraphs.push(checked_range(start, event_end)?);
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
    let mut reader = NsReader::from_reader(xml);
    let mut fragment_prefix = FragmentPrefix::Unseen;
    let mut root_depth = None;
    let mut capture = None::<(usize, usize)>;
    let mut depth = 0usize;
    let mut index = 0usize;
    let mut saw_root = false;
    loop {
        let start = usize::try_from(reader.buffer_position()).map_err(|_error| missing)?;
        let raw_event = reader.read_event().map_err(|_error| missing)?.into_owned();
        let resolver = reader.resolver().clone();
        let (namespace, event) = resolver.resolve_event(raw_event);
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
    fn rich_owner_edits_are_compact_durable_and_exactly_reversible() {
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
        assert!(!xml.contains("\n  "));
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
        assert!(!xml.contains("\n"));

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
