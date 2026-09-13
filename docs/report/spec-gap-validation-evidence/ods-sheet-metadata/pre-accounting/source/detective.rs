//! Bounded, source-qualified codec for ODF formula-auditing metadata.
//!
//! This module deliberately stops at the detective owner.  It does not find
//! an effective branch in markup-compatibility XML, rebuild a worksheet, or
//! evaluate a formula.  A scan retains exact source ranges and enough
//! namespace/ancestry information for a caller to make a narrowly scoped
//! replacement decision.  Unsupported or malformed owners remain opaque so a
//! no-op can retain their bytes; every mutating plan refuses them before it
//! allocates a candidate.

use std::{
    fmt,
    ops::{Deref, Range},
    sync::Arc,
};

use litchi_core::{Error, ExecutionContext, ExecutionError, Reservation, Resource, Result};
use quick_xml::{
    XmlVersion,
    events::{BytesCData, BytesEnd, BytesStart, Event},
    name::{Namespace, NamespaceResolver, PrefixDeclaration, ResolveResult},
    reader::NsReader,
};

use crate::model::detective::{
    Detective, Direction, HighlightedRange, Operation, OperationKind, write_detective,
};

/// ODF office namespace.
pub const OFFICE_NAMESPACE: &str = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
/// ODF table namespace.
pub const TABLE_NAMESPACE: &str = "urn:oasis:names:tc:opendocument:xmlns:table:1.0";
/// ODF text namespace.
pub const TEXT_NAMESPACE: &str = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";
/// XML namespace.
pub const XML_NAMESPACE: &str = "http://www.w3.org/XML/1998/namespace";
/// XMLNS namespace.
pub const XMLNS_NAMESPACE: &str = "http://www.w3.org/2000/xmlns/";
/// Markup Compatibility namespace.  It is intentionally never interpreted.
pub const MC_NAMESPACE: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";

/// Hard content XML ceiling from the sheet-metadata design.
pub const MAX_INPUT_BYTES: usize = 256 * 1024 * 1024;
/// Hard nesting ceiling from the sheet-metadata design.
pub const MAX_DEPTH: usize = 1_024;
/// Hard owner item ceiling from the sheet-metadata design.
pub const MAX_ITEMS: usize = 1_048_576;
/// Hard text/scalar ceiling from the sheet-metadata design.
pub const MAX_TEXT_BYTES: usize = 16 * 1024 * 1024;
/// Hard format-owned work ceiling from the sheet-metadata design.
pub const MAX_WORK_UNITS: u64 = 16_000_000_000;
/// Hard XML event ceiling from the sheet-metadata design.
pub const MAX_EVENTS: usize = 4_194_304;
/// Hard namespace-binding ceiling from the sheet-metadata design.
pub const MAX_NAMESPACE_BINDINGS: usize = 4_096;
const DEFAULT_MAX_EVENTS: usize = MAX_EVENTS;
const DEFAULT_MAX_NAMESPACE_BINDINGS: usize = MAX_NAMESPACE_BINDINGS;
const DEFAULT_MAX_OUTPUT_BYTES: usize = MAX_INPUT_BYTES;
const DEFAULT_MAX_SCRATCH_BYTES: usize = 64 * 1024 * 1024;
const ARC_ALLOCATION_WORDS: usize = 2;

fn arc_allocation_bytes<T>() -> Result<u64> {
    u64_len(
        std::mem::size_of::<T>()
            .checked_add(
                ARC_ALLOCATION_WORDS
                    .checked_mul(std::mem::size_of::<usize>())
                    .ok_or_else(|| invalid("ODS detective Arc allocation size overflows"))?,
            )
            .ok_or_else(|| invalid("ODS detective Arc allocation size overflows"))?,
    )
}

fn arc_str_allocation_bytes(length: usize) -> Result<u64> {
    u64_len(
        length
            .checked_add(
                ARC_ALLOCATION_WORDS
                    .checked_mul(std::mem::size_of::<usize>())
                    .ok_or_else(|| invalid("ODS detective Arc string size overflows"))?,
            )
            .ok_or_else(|| invalid("ODS detective Arc string size overflows"))?,
    )
}

/// Finite limits for one detective scan or splice plan.
///
/// Every field is bounded by a hard ceiling.  The caller may lower a bound,
/// but cannot use an unbounded sentinel.  The shared [`ExecutionContext`]
/// remains authoritative for hierarchical resource accounting and
/// cancellation; these fields are format-owned ceilings used for deterministic
/// diagnostics and preallocation.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct Limits {
    max_input_bytes: usize,
    max_output_bytes: usize,
    max_scratch_bytes: usize,
    max_depth: usize,
    max_events: usize,
    max_owners: usize,
    max_namespace_bindings: usize,
    max_highlighted_ranges: usize,
    max_operations: usize,
    max_text_bytes: usize,
    max_work_units: u64,
}

impl Default for Limits {
    fn default() -> Self {
        Self {
            max_input_bytes: MAX_INPUT_BYTES,
            max_output_bytes: DEFAULT_MAX_OUTPUT_BYTES,
            max_scratch_bytes: DEFAULT_MAX_SCRATCH_BYTES,
            max_depth: MAX_DEPTH,
            max_events: DEFAULT_MAX_EVENTS,
            max_owners: MAX_ITEMS,
            max_namespace_bindings: DEFAULT_MAX_NAMESPACE_BINDINGS,
            max_highlighted_ranges: MAX_ITEMS,
            max_operations: MAX_ITEMS,
            max_text_bytes: MAX_TEXT_BYTES,
            max_work_units: MAX_WORK_UNITS,
        }
    }
}

impl Limits {
    /// Creates a finite limit profile.
    ///
    /// The remaining fields use the hard/default values.  Use the checked
    /// setters to lower individual collection or output bounds.
    pub fn new(
        max_input_bytes: usize,
        max_output_bytes: usize,
        max_depth: usize,
        max_events: usize,
        max_work_units: u64,
    ) -> Result<Self> {
        let limits = Self {
            max_input_bytes,
            max_output_bytes,
            max_scratch_bytes: DEFAULT_MAX_SCRATCH_BYTES.min(max_output_bytes),
            max_depth,
            max_events,
            max_work_units,
            ..Self::default()
        };
        limits.validate()?;
        Ok(limits)
    }

    /// The conservative finite server profile.
    #[must_use]
    pub const fn server() -> Self {
        Self {
            max_input_bytes: 256 * 1024 * 1024,
            max_output_bytes: 256 * 1024 * 1024,
            max_scratch_bytes: 16 * 1024 * 1024,
            max_depth: 256,
            max_events: 1_048_576,
            max_owners: 65_536,
            max_namespace_bindings: 512,
            max_highlighted_ranges: 65_536,
            max_operations: 65_536,
            max_text_bytes: 4 * 1024 * 1024,
            max_work_units: 1_000_000_000,
        }
    }

    /// Sets the owner and child collection ceilings.
    pub fn with_item_limits(
        mut self,
        max_owners: usize,
        max_highlighted_ranges: usize,
        max_operations: usize,
    ) -> Result<Self> {
        self.max_owners = max_owners;
        self.max_highlighted_ranges = max_highlighted_ranges;
        self.max_operations = max_operations;
        self.validate()?;
        Ok(self)
    }

    /// Sets the text and namespace-context ceilings.
    pub fn with_scalar_limits(
        mut self,
        max_text_bytes: usize,
        max_namespace_bindings: usize,
    ) -> Result<Self> {
        self.max_text_bytes = max_text_bytes;
        self.max_namespace_bindings = max_namespace_bindings;
        self.validate()?;
        Ok(self)
    }

    /// Sets the scratch ceiling used by candidate construction.
    pub fn with_scratch_bytes(mut self, max_scratch_bytes: usize) -> Result<Self> {
        self.max_scratch_bytes = max_scratch_bytes;
        self.validate()?;
        Ok(self)
    }

    /// Input byte ceiling.
    #[must_use]
    pub const fn max_input_bytes(self) -> usize {
        self.max_input_bytes
    }

    /// Candidate output byte ceiling.
    #[must_use]
    pub const fn max_output_bytes(self) -> usize {
        self.max_output_bytes
    }

    /// Candidate scratch byte ceiling.
    #[must_use]
    pub const fn max_scratch_bytes(self) -> usize {
        self.max_scratch_bytes
    }

    /// Maximum XML nesting depth.
    #[must_use]
    pub const fn max_depth(self) -> usize {
        self.max_depth
    }

    /// Maximum XML events visited by one scan.
    #[must_use]
    pub const fn max_events(self) -> usize {
        self.max_events
    }

    /// Maximum direct detective owners retained by one snapshot.
    #[must_use]
    pub const fn max_owners(self) -> usize {
        self.max_owners
    }

    /// Maximum active namespace bindings retained in one context.
    #[must_use]
    pub const fn max_namespace_bindings(self) -> usize {
        self.max_namespace_bindings
    }

    /// Maximum highlighted ranges in one owner.
    #[must_use]
    pub const fn max_highlighted_ranges(self) -> usize {
        self.max_highlighted_ranges
    }

    /// Maximum operations in one owner.
    #[must_use]
    pub const fn max_operations(self) -> usize {
        self.max_operations
    }

    /// Maximum decoded scalar text bytes.
    #[must_use]
    pub const fn max_text_bytes(self) -> usize {
        self.max_text_bytes
    }

    /// Maximum format-owned work units.
    #[must_use]
    pub const fn max_work_units(self) -> u64 {
        self.max_work_units
    }

    fn validate(self) -> Result<()> {
        if self.max_input_bytes > MAX_INPUT_BYTES
            || self.max_output_bytes > MAX_INPUT_BYTES
            || self.max_scratch_bytes > MAX_INPUT_BYTES
            || self.max_depth > MAX_DEPTH
            || self.max_events > MAX_EVENTS
            || self.max_owners > MAX_ITEMS
            || self.max_namespace_bindings > MAX_NAMESPACE_BINDINGS
            || self.max_highlighted_ranges > MAX_ITEMS
            || self.max_operations > MAX_ITEMS
            || self.max_text_bytes > MAX_TEXT_BYTES
            || self.max_work_units > MAX_WORK_UNITS
        {
            return Err(invalid("ODS detective limits exceed a hard ceiling"));
        }
        if self.max_input_bytes == 0
            || self.max_output_bytes == 0
            || self.max_depth == 0
            || self.max_events == 0
            || self.max_owners == 0
            || self.max_namespace_bindings == 0
            || self.max_highlighted_ranges == 0
            || self.max_operations == 0
            || self.max_text_bytes == 0
            || self.max_work_units == 0
        {
            return Err(invalid("ODS detective limits must be finite and non-zero"));
        }
        if self.max_scratch_bytes == 0 {
            return Err(invalid("ODS detective scratch limit must be non-zero"));
        }
        Ok(())
    }
}

/// A borrowed source range.  The source is never normalized or copied.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct RawRange<'source> {
    source: &'source str,
    start: usize,
    end: usize,
}

impl<'source> RawRange<'source> {
    fn new(source: &'source str, range: Range<usize>) -> Result<Self> {
        if range.start > range.end || source.get(range.clone()).is_none() {
            return Err(invalid("ODS detective source range is invalid"));
        }
        Ok(Self {
            source,
            start: range.start,
            end: range.end,
        })
    }

    /// Source byte range, half-open.
    #[must_use]
    pub fn range(self) -> Range<usize> {
        self.start..self.end
    }

    /// Exact source bytes represented by this range.
    #[must_use]
    pub fn as_str(self) -> &'source str {
        // `new` proves this range, and all ranges are constructed from reader
        // offsets over this exact source.  Keep the fallback for defensive
        // callers if a future constructor is added.
        self.source.get(self.start..self.end).unwrap_or("")
    }

    /// Number of source bytes represented by this range.
    #[must_use]
    pub fn len(self) -> usize {
        self.end.saturating_sub(self.start)
    }

    /// Whether this range contains no source bytes.
    #[must_use]
    pub const fn is_empty(self) -> bool {
        self.start == self.end
    }
}

/// One inherited namespace binding captured at an owner boundary.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct NamespaceBinding {
    prefix: String,
    uri: String,
}

impl NamespaceBinding {
    /// Namespace prefix, or the empty string for the default binding.
    #[must_use]
    pub fn prefix(&self) -> &str {
        &self.prefix
    }

    /// Expanded namespace URI.
    #[must_use]
    pub fn uri(&self) -> &str {
        &self.uri
    }
}

/// Inherited namespace context required to parse or author one owner.
#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct NamespaceContext {
    bindings: Vec<NamespaceBinding>,
}

impl NamespaceContext {
    /// Active bindings in expanded-namespace order as reported by quick-xml.
    #[must_use]
    pub fn bindings(&self) -> &[NamespaceBinding] {
        &self.bindings
    }

    /// Returns the URI bound to a prefix.
    #[must_use]
    pub fn uri_for_prefix(&self, prefix: &str) -> Option<&str> {
        self.bindings
            .iter()
            .find(|binding| binding.prefix == prefix)
            .map(|binding| binding.uri.as_str())
    }

    /// Returns one prefix currently bound to the ODF table namespace.
    #[must_use]
    pub fn table_prefix(&self) -> Option<&str> {
        self.bindings
            .iter()
            .find(|binding| binding.uri == TABLE_NAMESPACE)
            .map(|binding| binding.prefix.as_str())
    }

    fn has_table_binding(&self, prefix: &str) -> bool {
        self.uri_for_prefix(prefix) == Some(TABLE_NAMESPACE)
    }
}

fn namespace_context_bytes(context: &NamespaceContext) -> Result<usize> {
    context.bindings.iter().try_fold(0usize, |total, binding| {
        total
            .checked_add(binding.prefix.len())
            .and_then(|value| value.checked_add(binding.uri.len()))
            .and_then(|value| value.checked_add(std::mem::size_of::<NamespaceBinding>()))
            .ok_or_else(|| invalid("ODS namespace context size overflows"))
    })
}

/// The actual ODF cell parent that owns a direct detective child.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum CellKind {
    /// An ordinary physical `table:table-cell`.
    TableCell,
    /// An explicit physical `table:covered-table-cell`.
    CoveredTableCell,
}

/// Expanded ancestry captured for source qualification.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Ancestor {
    namespace: Option<String>,
    local_name: String,
}

impl Ancestor {
    /// Expanded namespace URI, or `None` for an unbound name.
    #[must_use]
    pub fn namespace(&self) -> Option<&str> {
        self.namespace.as_deref()
    }

    /// Local element name.
    #[must_use]
    pub fn local_name(&self) -> &str {
        &self.local_name
    }
}

/// Provenance needed before a direct owner can be edited.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct SourceQualification {
    parent: CellKind,
    ancestors: Vec<Ancestor>,
    namespace_context: NamespaceContext,
    foreign_ancestor: bool,
    markup_compatibility_ancestor: bool,
}

impl SourceQualification {
    /// Actual direct ODF parent kind.
    #[must_use]
    pub const fn parent(&self) -> CellKind {
        self.parent
    }

    /// Expanded ancestry from the document root to the direct parent.
    #[must_use]
    pub fn ancestors(&self) -> &[Ancestor] {
        &self.ancestors
    }

    /// Namespace context inherited by the owner.
    #[must_use]
    pub const fn namespace_context(&self) -> &NamespaceContext {
        &self.namespace_context
    }

    /// Whether any ancestor is outside the supported ODF ancestry.
    #[must_use]
    pub const fn has_foreign_ancestor(&self) -> bool {
        self.foreign_ancestor
    }

    /// Whether the owner occurs beneath an MCE branch.  Such branches are
    /// opaque; this flag is diagnostic only and never selects a branch.
    #[must_use]
    pub const fn has_markup_compatibility_ancestor(&self) -> bool {
        self.markup_compatibility_ancestor
    }
}

/// One ordered child of a typed detective owner.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum OrderedChild<'source> {
    /// A highlighted range at the given semantic vector position.
    HighlightedRange {
        /// Position in [`Detective::highlighted_ranges`].
        index: usize,
        /// Exact source span of the child element.
        raw: RawRange<'source>,
    },
    /// An operation at the given semantic vector position.
    Operation {
        /// Position in [`Detective::operations`].
        index: usize,
        /// Exact source span of the child element.
        raw: RawRange<'source>,
    },
}

impl<'source> OrderedChild<'source> {
    /// Exact source span of this child.
    #[must_use]
    pub const fn raw(self) -> RawRange<'source> {
        match self {
            Self::HighlightedRange { raw, .. } | Self::Operation { raw, .. } => raw,
        }
    }
}

/// Retained source-qualified state for a typed detective.
///
/// Keeping the complete value behind one `Arc` is deliberate.  Cloning a
/// typed view must not deep-clone its semantic vectors or source metadata
/// without charging a new budget reservation.  Standalone parsing installs
/// input and retained-memory reservation guards in this object before
/// returning the view; snapshot-owned views share the snapshot's reservations
/// instead.
#[derive(Debug)]
struct RetainedDetective<'source> {
    value: Detective,
    raw: RawRange<'source>,
    ordered: Vec<OrderedChild<'source>>,
    namespace_context: NamespaceContext,
    source_prefix: String,
    lexical_fidelity: LexicalFidelity,
    input_guard: Option<Arc<Reservation>>,
    memory_guard: Option<Arc<Reservation>>,
}

/// A typed detective value tied to its exact source owner.
///
/// The semantic value is the existing [`Detective`] model.  This wrapper only
/// adds source ranges, ordered child provenance, and lexical-fidelity status;
/// it does not define a second highlight or operation value model.  Cloning
/// the wrapper is shallow and retains the same budget guard.
#[derive(Clone, Debug)]
pub struct TypedDetective<'source> {
    retained: Arc<RetainedDetective<'source>>,
}

impl PartialEq for RetainedDetective<'_> {
    fn eq(&self, other: &Self) -> bool {
        self.value == other.value
            && self.raw == other.raw
            && self.ordered == other.ordered
            && self.namespace_context == other.namespace_context
            && self.source_prefix == other.source_prefix
            && self.lexical_fidelity == other.lexical_fidelity
    }
}

impl Eq for RetainedDetective<'_> {}

impl PartialEq for TypedDetective<'_> {
    fn eq(&self, other: &Self) -> bool {
        self.retained == other.retained
    }
}

impl Eq for TypedDetective<'_> {}

impl<'source> TypedDetective<'source> {
    /// Existing semantic detective value.
    #[must_use]
    pub fn value(&self) -> &Detective {
        &self.retained.value
    }

    /// Retained semantic detective value through the existing model API.
    ///
    /// This dereference is intentionally borrowed; callers must retain this
    /// typed wrapper while using any returned references.
    #[must_use]
    pub fn retained(&self) -> &Detective {
        self.value()
    }

    /// Exact owner source range.
    #[must_use]
    pub fn raw(&self) -> RawRange<'source> {
        self.retained.raw
    }

    /// Ordered highlighted-range/operation source children.
    #[must_use]
    pub fn ordered_children(&self) -> &[OrderedChild<'source>] {
        &self.retained.ordered
    }

    /// Inherited namespace context used for this owner.
    #[must_use]
    pub fn namespace_context(&self) -> &NamespaceContext {
        &self.retained.namespace_context
    }

    /// Source prefix used by the owner element.
    #[must_use]
    pub fn source_prefix(&self) -> &str {
        &self.retained.source_prefix
    }

    /// Whether canonical changed-owner output can retain source lexical data.
    #[must_use]
    pub fn lexical_fidelity(&self) -> LexicalFidelity {
        self.retained.lexical_fidelity
    }

    /// Number of bytes retained by the standalone parse guard, if present.
    #[must_use]
    pub fn retained_memory_bytes(&self) -> u64 {
        self.retained
            .memory_guard
            .as_ref()
            .map_or(0, |guard| guard.amount())
    }

    /// Number of input bytes retained by the standalone parse guard, if any.
    #[must_use]
    pub fn retained_input_bytes(&self) -> u64 {
        self.retained
            .input_guard
            .as_ref()
            .map_or(0, |guard| guard.amount())
    }

    fn with_guards(
        mut self,
        input_guard: Option<Arc<Reservation>>,
        memory_guard: Option<Arc<Reservation>>,
    ) -> Result<Self> {
        let retained = Arc::get_mut(&mut self.retained).ok_or_else(|| {
            unsupported("ODS detective typed state was unexpectedly shared while retaining memory")
        })?;
        retained.input_guard = input_guard;
        retained.memory_guard = memory_guard;
        Ok(self)
    }

    fn mark_preservation_required(&mut self) -> Result<()> {
        let retained = Arc::get_mut(&mut self.retained).ok_or_else(|| {
            unsupported("ODS detective typed state was unexpectedly shared while qualifying source")
        })?;
        retained.lexical_fidelity = LexicalFidelity::PreservationRequired;
        Ok(())
    }

    fn into_value(self) -> Result<Detective> {
        let retained = Arc::try_unwrap(self.retained).map_err(|_| {
            unsupported("ODS detective typed value is still borrowed while extracting its model")
        })?;
        let RetainedDetective { value, .. } = retained;
        Ok(value)
    }
}

impl Deref for TypedDetective<'_> {
    type Target = Detective;

    fn deref(&self) -> &Self::Target {
        self.value()
    }
}

/// Lexical data retained by a source owner.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum LexicalFidelity {
    /// The source is equivalent to the deterministic writer form for its
    /// inherited table prefix.
    Canonical,
    /// The source contains comments, processing instructions, noncanonical
    /// quoting/spacing, or another lexical form that this focused writer does
    /// not promise to preserve on a changed owner.
    PreservationRequired,
}

/// Why a recognized-looking owner is opaque to mutation.
#[derive(Clone, Debug, PartialEq, Eq)]
pub enum OpaqueReason {
    /// The owner is under a foreign wrapper or ancestor.
    ForeignAncestry,
    /// The owner occurs beneath an MCE branch; no branch is selected.
    MarkupCompatibilityBranch,
    /// More than one direct detective owner occurs in one physical cell.
    DuplicateDirectOwner,
    /// A direct cell child order makes a focused splice unsafe.
    InvalidCellChildOrder,
    /// An unsupported direct child makes insertion anchors ambiguous.
    UnsupportedCellChild,
    /// The owner failed the detective grammar.
    Malformed(String),
    /// The owner is valid but its source-local lexical data cannot be retained
    /// by the deterministic writer.
    PreservationRequired,
}

impl fmt::Display for OpaqueReason {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::ForeignAncestry => formatter.write_str("foreign owner ancestry"),
            Self::MarkupCompatibilityBranch => formatter.write_str("markup-compatibility branch"),
            Self::DuplicateDirectOwner => formatter.write_str("duplicate direct detective owner"),
            Self::InvalidCellChildOrder => formatter.write_str("invalid direct cell child order"),
            Self::UnsupportedCellChild => formatter.write_str("unsupported direct cell child"),
            Self::Malformed(message) => formatter.write_str(message),
            Self::PreservationRequired => {
                formatter.write_str("owner lexical data requires preservation")
            },
        }
    }
}

/// One source-qualified direct detective owner.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct DetectiveOwner<'source> {
    id: OwnerId,
    cell: CellId,
    raw: RawRange<'source>,
    qualification: SourceQualification,
    state: OwnerState<'source>,
}

impl<'source> DetectiveOwner<'source> {
    /// Stable owner index within its snapshot.
    #[must_use]
    pub const fn id(&self) -> OwnerId {
        self.id
    }

    /// Owning physical cell index.
    #[must_use]
    pub const fn cell(&self) -> CellId {
        self.cell
    }

    /// Exact source span.
    #[must_use]
    pub const fn raw(&self) -> RawRange<'source> {
        self.raw
    }

    /// Source qualification and inherited namespace context.
    #[must_use]
    pub const fn qualification(&self) -> &SourceQualification {
        &self.qualification
    }

    /// Typed value when the owner is schema-valid and editable as a focused
    /// owner.
    #[must_use]
    pub fn typed(&self) -> Option<&TypedDetective<'source>> {
        match &self.state {
            OwnerState::Typed(value) => Some(value),
            OwnerState::Opaque { .. } => None,
        }
    }

    /// Opaque diagnostic when the owner cannot be edited.
    #[must_use]
    pub fn opaque_reason(&self) -> Option<&OpaqueReason> {
        match &self.state {
            OwnerState::Typed(_) => None,
            OwnerState::Opaque { reason } => Some(reason),
        }
    }
}

#[derive(Clone, Debug, PartialEq, Eq)]
enum OwnerState<'source> {
    Typed(TypedDetective<'source>),
    Opaque { reason: OpaqueReason },
}

/// Stable physical-cell index in a source snapshot.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub struct CellId(usize);

impl CellId {
    /// Numeric index for diagnostics and deterministic ordering.
    #[must_use]
    pub const fn index(self) -> usize {
        self.0
    }
}

/// Stable direct-owner index in a source snapshot.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub struct OwnerId(usize);

impl OwnerId {
    /// Numeric index for diagnostics and deterministic ordering.
    #[must_use]
    pub const fn index(self) -> usize {
        self.0
    }
}

/// One physical cell's source qualification and direct-child sequence.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct CellSite<'source> {
    id: CellId,
    raw: RawRange<'source>,
    qname: String,
    empty: bool,
    kind: CellKind,
    qualification: SourceQualification,
    detective: Option<OwnerId>,
    direct_owner_count: usize,
    sequence_valid: bool,
    has_unsupported_child: bool,
    insertion_anchor: usize,
}

impl<'source> CellSite<'source> {
    /// Stable cell index.
    #[must_use]
    pub const fn id(&self) -> CellId {
        self.id
    }

    /// Exact physical cell span.
    #[must_use]
    pub const fn raw(&self) -> RawRange<'source> {
        self.raw
    }

    /// Source qualified name of the physical cell element.
    #[must_use]
    pub fn qname(&self) -> &str {
        &self.qname
    }

    /// Whether the physical cell was encoded with an empty-element tag.
    #[must_use]
    pub const fn is_empty_element(&self) -> bool {
        self.empty
    }

    /// Physical ODF cell role.
    #[must_use]
    pub const fn kind(&self) -> CellKind {
        self.kind
    }

    /// Cell ancestry and inherited namespace context.
    #[must_use]
    pub const fn qualification(&self) -> &SourceQualification {
        &self.qualification
    }

    /// Existing direct detective owner, if any.
    #[must_use]
    pub const fn detective_owner(&self) -> Option<OwnerId> {
        self.detective
    }

    /// Whether the physical cell has a schema-valid direct child ordering.
    #[must_use]
    pub const fn sequence_valid(&self) -> bool {
        self.sequence_valid
    }

    /// Whether an unknown/foreign direct element blocks an insertion anchor.
    #[must_use]
    pub const fn has_unsupported_child(&self) -> bool {
        self.has_unsupported_child
    }

    fn is_editable_anchor(&self) -> Result<()> {
        self.is_editable_owner()?;
        if self.has_unsupported_child {
            return Err(unsupported(
                "ODS detective insertion anchor is ambiguous around an unsupported sibling",
            ));
        }
        Ok(())
    }

    fn is_editable_owner(&self) -> Result<()> {
        if self.qualification.has_foreign_ancestor()
            || self.qualification.has_markup_compatibility_ancestor()
        {
            return Err(unsupported(
                "ODS detective cell has foreign or MCE ancestry",
            ));
        }
        if !self.sequence_valid {
            return Err(unsupported("ODS detective cell child order is malformed"));
        }
        Ok(())
    }
}

/// Immutable borrowed source index for detective owners.
pub struct Snapshot<'source> {
    source: &'source str,
    limits: Limits,
    owners: Vec<DetectiveOwner<'source>>,
    cells: Vec<CellSite<'source>>,
    _input_reservation: Option<Reservation>,
    _memory_reservation: Option<Reservation>,
}

impl fmt::Debug for Snapshot<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Snapshot")
            .field("source_bytes", &self.source.len())
            .field("owners", &self.owners.len())
            .field("cells", &self.cells.len())
            .field("limits", &self.limits)
            .finish_non_exhaustive()
    }
}

impl<'source> Snapshot<'source> {
    /// Parse and index `content.xml` under the caller's execution context.
    pub fn parse_with_context(
        source: &'source str,
        limits: Limits,
        context: &ExecutionContext,
    ) -> Result<Self> {
        limits.validate()?;
        context.check().map_err(map_execution)?;
        if source.len() > limits.max_input_bytes {
            return Err(limit("input bytes", source.len(), limits.max_input_bytes));
        }

        let mut ledger = Ledger::new(context, limits);
        let input_reservation = context
            .reserve(Resource::InputBytes, u64_len(source.len())?)
            .map_err(map_execution)?;
        let _parse_memory_floor = context
            .reserve(Resource::Memory, u64_len(source.len().min(4096))?)
            .map_err(map_execution)?;
        ledger.charge_work(u64_len(source.len())?)?;

        let scanned = scan_source(source, &mut ledger)?;
        if scanned.spans.len() > limits.max_events {
            return Err(limit("XML events", scanned.spans.len(), limits.max_events));
        }
        let (mut cells, cell_by_span) = build_cells(source, &scanned.spans, &mut ledger)?;
        let owners = build_owners(
            source,
            &scanned.spans,
            &cells,
            &cell_by_span,
            limits,
            &mut ledger,
        )?;
        attach_owner_ids(&mut cells, &owners)?;

        let mut memory_reservation = ledger.memory.take();
        // The retained index is bounded by the source/event limits.  Charge a
        // small per-record floor before publishing the vectors, and keep the
        // reservation alive with the snapshot.  This is deliberately a peak
        // bound rather than an allocator-size guess.
        let record_count = owners
            .len()
            .checked_add(cells.len())
            .ok_or_else(|| invalid("ODS detective owner count overflows"))?;
        if record_count != 0 {
            let bytes = u64_len(
                record_count
                    .checked_mul(std::mem::size_of::<usize>() * 4)
                    .ok_or_else(|| invalid("ODS detective index memory overflows"))?,
            )?;
            let reservation = context
                .reserve(Resource::Memory, bytes)
                .map_err(map_execution)?;
            merge_reservation(&mut memory_reservation, reservation)?;
        }

        Ok(Self {
            source,
            limits,
            owners,
            cells,
            _input_reservation: Some(input_reservation),
            _memory_reservation: memory_reservation,
        })
    }

    /// Borrowed source bytes used by the scan.
    #[must_use]
    pub const fn source(&self) -> &'source str {
        self.source
    }

    /// Limits captured by this snapshot.
    #[must_use]
    pub const fn limits(&self) -> Limits {
        self.limits
    }

    /// All direct physical detective owners in source order.
    #[must_use]
    pub fn owners(&self) -> &[DetectiveOwner<'source>] {
        &self.owners
    }

    /// All physical cells that can be selected by a caller-owned worksheet
    /// locator.
    #[must_use]
    pub fn cells(&self) -> &[CellSite<'source>] {
        &self.cells
    }

    /// Returns the direct owner attached to one physical cell.
    pub fn detective_owner(&self, cell: CellId) -> Result<Option<&DetectiveOwner<'source>>> {
        let site = self
            .cells
            .get(cell.0)
            .ok_or_else(|| invalid("ODS detective cell selector is out of range"))?;
        Ok(site.detective.and_then(|owner| self.owners.get(owner.0)))
    }

    /// Returns a typed detective value, refusing an opaque or malformed owner.
    pub fn detective(&self, cell: CellId) -> Result<Option<&Detective>> {
        let Some(owner) = self.detective_owner(cell)? else {
            return Ok(None);
        };
        let Some(typed) = owner.typed() else {
            let reason = owner
                .opaque_reason()
                .map(ToString::to_string)
                .unwrap_or_else(|| "unknown owner state".to_string());
            return Err(unsupported(format!(
                "ODS detective owner {} is opaque: {reason}",
                owner.id.index(),
            )));
        };
        Ok(Some(typed.value()))
    }

    /// Plan a focused replacement of one existing typed owner.
    pub fn plan_replace(
        &self,
        owner: OwnerId,
        value: &Detective,
        context: &ExecutionContext,
    ) -> Result<SplicePlan<'source>> {
        let owner = self
            .owners
            .get(owner.0)
            .ok_or_else(|| invalid("ODS detective owner selector is out of range"))?;
        let typed = owner.typed().ok_or_else(|| {
            unsupported(format!(
                "ODS detective owner {} is opaque and cannot be replaced",
                owner.id.index()
            ))
        })?;
        let cell = self
            .cells
            .get(owner.cell.0)
            .ok_or_else(|| invalid("ODS detective owner cell disappeared"))?;
        cell.is_editable_owner()?;
        if typed.value() == value {
            return SplicePlan::new(
                self.source,
                owner.raw,
                String::new(),
                PlanKind::NoOp,
                self.limits,
                context,
            );
        }
        if typed.lexical_fidelity() != LexicalFidelity::Canonical {
            return Err(unsupported(
                "ODS detective owner has source lexical data that the focused writer cannot retain",
            ));
        }
        ensure_model_limits(value, self.limits)?;
        let (_, rendered_len) = rendered_lengths(value, typed.source_prefix())?;
        preflight_candidate_lengths(
            self.source.len(),
            owner.raw.len(),
            rendered_len,
            self.limits,
        )?;
        let rendered = render_for_prefix(
            value,
            typed.source_prefix(),
            typed.namespace_context(),
            self.limits,
            context,
        )?;
        if rendered.text == typed.raw().as_str() {
            return SplicePlan::new(
                self.source,
                owner.raw,
                String::new(),
                PlanKind::NoOp,
                self.limits,
                context,
            );
        }
        SplicePlan::new_with_reservation(
            self.source,
            owner.raw,
            rendered.text,
            PlanKind::Replace,
            self.limits,
            context,
            Some(rendered.memory),
        )
    }

    /// Plan removal of one existing typed owner.  The surrounding cell and
    /// all unknown siblings remain byte-for-byte untouched.
    pub fn plan_remove(
        &self,
        owner: OwnerId,
        context: &ExecutionContext,
    ) -> Result<SplicePlan<'source>> {
        let owner = self
            .owners
            .get(owner.0)
            .ok_or_else(|| invalid("ODS detective owner selector is out of range"))?;
        if owner.typed().is_none() {
            return Err(unsupported(format!(
                "ODS detective owner {} is opaque and cannot be removed",
                owner.id.index()
            )));
        }
        let cell = self
            .cells
            .get(owner.cell.0)
            .ok_or_else(|| invalid("ODS detective owner cell disappeared"))?;
        cell.is_editable_owner()?;
        let typed = owner.typed().ok_or_else(|| {
            unsupported(format!(
                "ODS detective owner {} is opaque and cannot be removed",
                owner.id.index()
            ))
        })?;
        if typed.lexical_fidelity() != LexicalFidelity::Canonical {
            return Err(unsupported(
                "ODS detective owner has source lexical data that focused removal cannot retain",
            ));
        }
        SplicePlan::new(
            self.source,
            owner.raw,
            String::new(),
            PlanKind::Remove,
            self.limits,
            context,
        )
    }

    /// Plan insertion of an explicitly present, possibly empty detective
    /// container into a cell that has no direct detective owner.
    pub fn plan_insert(
        &self,
        cell: CellId,
        value: &Detective,
        context: &ExecutionContext,
    ) -> Result<SplicePlan<'source>> {
        let site = self
            .cells
            .get(cell.0)
            .ok_or_else(|| invalid("ODS detective cell selector is out of range"))?;
        site.is_editable_anchor()?;
        if site.detective.is_some() || site.direct_owner_count != 0 {
            return Err(unsupported(
                "ODS detective cell already has a direct detective owner",
            ));
        }
        let prefix = site
            .qualification
            .namespace_context
            .table_prefix()
            .ok_or_else(|| {
                unsupported("ODS detective cell has no inherited table namespace prefix")
            })?;
        ensure_model_limits(value, self.limits)?;
        let (_, rendered_len) = rendered_lengths(value, prefix)?;
        let replacement_len = if site.empty {
            expanded_empty_cell_len(site, rendered_len)?
        } else {
            rendered_len
        };
        preflight_candidate_lengths(
            self.source.len(),
            if site.empty { site.raw.len() } else { 0 },
            replacement_len,
            self.limits,
        )?;
        let rendered = render_for_prefix(
            value,
            prefix,
            &site.qualification.namespace_context,
            self.limits,
            context,
        )?;
        if site.empty {
            let (replacement, expansion_memory) = expand_empty_cell(site, &rendered.text, context)?;
            let mut precharged = Some(rendered.memory);
            merge_reservation(&mut precharged, expansion_memory)?;
            return SplicePlan::new_with_reservation(
                self.source,
                site.raw,
                replacement,
                PlanKind::Insert,
                self.limits,
                context,
                precharged,
            );
        }
        let insertion = insertion_anchor(site)?;
        SplicePlan::new_with_reservation(
            self.source,
            RawRange::new(self.source, insertion.clone())?,
            rendered.text,
            PlanKind::Insert,
            self.limits,
            context,
            Some(rendered.memory),
        )
    }
}

/// A checked, immutable XML splice plan owned by this codec.
///
/// Applying the plan returns a guarded candidate and never mutates the source.
/// The host transaction remains responsible for source-version fencing,
/// complete-package reopen, and publication atomicity.
pub struct SplicePlan<'source> {
    source: &'source str,
    range: Range<usize>,
    replacement: String,
    kind: PlanKind,
    output_len: usize,
    limits: Limits,
    _staged_memory: Option<Reservation>,
}

impl fmt::Debug for SplicePlan<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("SplicePlan")
            .field("range", &self.range)
            .field("replacement_bytes", &self.replacement.len())
            .field("kind", &self.kind)
            .field("output_len", &self.output_len)
            .finish_non_exhaustive()
    }
}

/// Planned focused operation.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum PlanKind {
    /// Semantic no-op; the original source is retained exactly by the host.
    NoOp,
    /// Replace one existing detective owner.
    Replace,
    /// Remove one existing detective owner.
    Remove,
    /// Insert one direct detective owner at a checked cell-content anchor.
    Insert,
}

/// Candidate bytes returned by a checked splice.
///
/// A no-op candidate borrows the source and therefore performs no output
/// allocation.  A changed candidate owns its output and retains both the
/// output-byte and transient-memory reservations until this wrapper is
/// dropped.  Callers should retain this value through publication; there is
/// intentionally no `into_string` escape hatch that could discard its guards.
#[derive(Clone, Debug)]
pub struct AppliedCandidate<'source> {
    text: CandidateText<'source>,
    output_guard: Option<Arc<Reservation>>,
    memory_guard: Option<Arc<Reservation>>,
}

#[derive(Clone, Debug)]
enum CandidateText<'source> {
    Borrowed(&'source str),
    Owned(Arc<str>),
}

impl AppliedCandidate<'_> {
    /// Candidate text for publication or readback.
    #[must_use]
    pub fn as_str(&self) -> &str {
        match &self.text {
            CandidateText::Borrowed(value) => value,
            CandidateText::Owned(value) => value,
        }
    }

    /// Whether this candidate shares the original source bytes.
    #[must_use]
    pub const fn is_borrowed(&self) -> bool {
        matches!(self.text, CandidateText::Borrowed(_))
    }

    /// Retained output-byte reservation, in bytes.
    #[must_use]
    pub fn retained_output_bytes(&self) -> u64 {
        self.output_guard.as_ref().map_or(0, |guard| guard.amount())
    }

    /// Retained transient-memory reservation, in bytes.
    #[must_use]
    pub fn retained_memory_bytes(&self) -> u64 {
        self.memory_guard.as_ref().map_or(0, |guard| guard.amount())
    }
}

impl AsRef<str> for AppliedCandidate<'_> {
    fn as_ref(&self) -> &str {
        self.as_str()
    }
}

impl Deref for AppliedCandidate<'_> {
    type Target = str;

    fn deref(&self) -> &Self::Target {
        self.as_str()
    }
}

impl PartialEq for AppliedCandidate<'_> {
    fn eq(&self, other: &Self) -> bool {
        self.as_str() == other.as_str()
    }
}

impl Eq for AppliedCandidate<'_> {}

impl PartialEq<str> for AppliedCandidate<'_> {
    fn eq(&self, other: &str) -> bool {
        self.as_str() == other
    }
}

impl PartialEq<AppliedCandidate<'_>> for str {
    fn eq(&self, other: &AppliedCandidate<'_>) -> bool {
        self == other.as_str()
    }
}

impl<'source> SplicePlan<'source> {
    fn new(
        source: &'source str,
        raw: RawRange<'source>,
        replacement: String,
        kind: PlanKind,
        limits: Limits,
        context: &ExecutionContext,
    ) -> Result<SplicePlan<'source>> {
        Self::new_with_reservation(source, raw, replacement, kind, limits, context, None)
    }

    fn new_with_reservation(
        source: &'source str,
        raw: RawRange<'source>,
        replacement: String,
        kind: PlanKind,
        limits: Limits,
        context: &ExecutionContext,
        mut precharged_memory: Option<Reservation>,
    ) -> Result<SplicePlan<'source>> {
        limits.validate()?;
        context.check().map_err(map_execution)?;
        let range = raw.range();
        if source.get(range.clone()) != Some(raw.as_str()) {
            return Err(invalid(
                "ODS detective splice source range is not from its snapshot",
            ));
        }
        if matches!(kind, PlanKind::NoOp) {
            drop(precharged_memory);
            return Ok(SplicePlan {
                source,
                range,
                replacement,
                kind,
                output_len: source.len(),
                limits,
                _staged_memory: None,
            });
        }
        let old_len = range
            .end
            .checked_sub(range.start)
            .ok_or_else(|| invalid("ODS detective splice range is invalid"))?;
        let output_len = source
            .len()
            .checked_sub(old_len)
            .and_then(|length| length.checked_add(replacement.len()))
            .ok_or_else(|| invalid("ODS detective candidate length overflows"))?;
        if output_len > limits.max_output_bytes {
            return Err(limit("output bytes", output_len, limits.max_output_bytes));
        }
        let scratch = output_len
            .checked_add(replacement.len())
            .ok_or_else(|| invalid("ODS detective candidate scratch size overflows"))?;
        if scratch > limits.max_scratch_bytes {
            return Err(limit("scratch bytes", scratch, limits.max_scratch_bytes));
        }
        context
            .consume(Resource::Work, u64_len(output_len)?)
            .map_err(map_execution)?;
        let staged_memory = context
            .reserve(Resource::Memory, u64_len(scratch)?)
            .map_err(map_execution)?;
        merge_reservation(&mut precharged_memory, staged_memory)?;
        Ok(SplicePlan {
            source,
            range,
            replacement,
            kind,
            output_len,
            limits,
            _staged_memory: precharged_memory,
        })
    }

    /// Planned operation kind.
    #[must_use]
    pub const fn kind(&self) -> PlanKind {
        self.kind
    }

    /// Exact source range replaced by this plan.
    #[must_use]
    pub fn range(&self) -> Range<usize> {
        self.range.clone()
    }

    /// Authored replacement bytes.  An empty replacement is meaningful for a
    /// removal and is empty for a semantic no-op as well; inspect [`Self::kind`].
    #[must_use]
    pub fn replacement(&self) -> &str {
        &self.replacement
    }

    /// Candidate byte length after applying this plan.
    #[must_use]
    pub const fn output_len(&self) -> usize {
        self.output_len
    }

    /// Whether this plan preserves the source exactly.
    #[must_use]
    pub const fn is_noop(&self) -> bool {
        matches!(self.kind, PlanKind::NoOp)
    }

    /// Apply against the exact source captured by this plan.
    pub fn apply(&self, context: &ExecutionContext) -> Result<AppliedCandidate<'source>> {
        self.apply_to(self.source, context)
    }

    /// Apply against a separately supplied source after an exact-byte check.
    /// This is the source-version handoff used by a host transaction; retain
    /// the returned wrapper through candidate publication.
    pub fn apply_to<'candidate>(
        &self,
        source: &'candidate str,
        context: &ExecutionContext,
    ) -> Result<AppliedCandidate<'candidate>> {
        context.check().map_err(map_execution)?;
        if source.as_bytes() != self.source.as_bytes() {
            return Err(Error::SourceChanged {
                expected: litchi_core::SourceVersion::new(0, 0),
                observed: litchi_core::SourceVersion::new(0, 1),
            });
        }
        if self.is_noop() {
            return Ok(AppliedCandidate {
                text: CandidateText::Borrowed(source),
                output_guard: None,
                memory_guard: None,
            });
        }
        let start = self.range.start;
        let end = self.range.end;
        if start > end || end > source.len() {
            return Err(invalid("ODS detective splice range is invalid"));
        }
        let old_len = end - start;
        let candidate_len = source
            .len()
            .checked_sub(old_len)
            .and_then(|length| length.checked_add(self.replacement.len()))
            .ok_or_else(|| invalid("ODS detective candidate length overflows"))?;
        if candidate_len != self.output_len || candidate_len > self.limits.max_output_bytes {
            return Err(limit(
                "output bytes",
                candidate_len,
                self.limits.max_output_bytes,
            ));
        }
        let output_bytes = u64_len(candidate_len)?;
        let scratch = candidate_len
            .checked_add(self.replacement.len())
            .ok_or_else(|| invalid("ODS detective candidate scratch size overflows"))?;
        if scratch > self.limits.max_scratch_bytes {
            return Err(limit(
                "scratch bytes",
                scratch,
                self.limits.max_scratch_bytes,
            ));
        }
        context
            .consume(Resource::Work, output_bytes)
            .map_err(map_execution)?;
        let output_guard = context
            .reserve(Resource::OutputBytes, output_bytes)
            .map_err(map_execution)?;
        let memory_guard = context
            .reserve(Resource::Memory, u64_len(scratch)?)
            .map_err(map_execution)?;
        let mut output = String::new();
        output
            .try_reserve_exact(candidate_len)
            .map_err(|source| allocation("ODS detective candidate output", source))?;
        output.push_str(
            source
                .get(..start)
                .ok_or_else(|| invalid("ODS detective splice prefix is invalid"))?,
        );
        output.push_str(&self.replacement);
        output.push_str(
            source
                .get(end..)
                .ok_or_else(|| invalid("ODS detective splice suffix is invalid"))?,
        );
        if output.len() != candidate_len {
            return Err(invalid(
                "ODS detective candidate length changed while splicing",
            ));
        }
        context.check().map_err(map_execution)?;
        let wrapper_bytes = arc_str_allocation_bytes(candidate_len)?
            .checked_add(
                arc_allocation_bytes::<Reservation>()?
                    .checked_mul(2)
                    .ok_or_else(|| invalid("ODS detective candidate guard size overflows"))?,
            )
            .ok_or_else(|| invalid("ODS detective candidate wrapper size overflows"))?;
        let wrapper_memory = context
            .reserve(Resource::Memory, wrapper_bytes)
            .map_err(map_execution)?;
        let mut retained_memory = Some(memory_guard);
        merge_reservation(&mut retained_memory, wrapper_memory)?;
        let owned = Arc::<str>::from(output);
        Ok(AppliedCandidate {
            text: CandidateText::Owned(owned),
            output_guard: Some(Arc::new(output_guard)),
            memory_guard: retained_memory.map(Arc::new),
        })
    }
}

/// Convenient name for a checked owner-authoring plan.
pub type AuthoringPlan<'source> = SplicePlan<'source>;
/// Convenient name for a checked owner-removal plan.
pub type RemovalPlan<'source> = SplicePlan<'source>;

/// Parse one standalone `table:detective` owner.
pub fn parse_detective(
    source: &str,
    limits: Limits,
    context: &ExecutionContext,
) -> Result<Detective> {
    parse_detective_with_namespace(source, 0, &NamespaceContext::default(), limits, context)
        .and_then(TypedDetective::into_value)
}

/// Parse one owner while supplying the namespace declarations inherited from
/// its physical source parent.  This accepts namespace aliases without
/// normalizing them into a second semantic value model.
pub fn parse_detective_with_namespace<'source>(
    source: &'source str,
    source_offset: usize,
    namespace_context: &NamespaceContext,
    limits: Limits,
    context: &ExecutionContext,
) -> Result<TypedDetective<'source>> {
    limits.validate()?;
    context.check().map_err(map_execution)?;
    if source.len() > limits.max_input_bytes {
        return Err(limit(
            "detective input bytes",
            source.len(),
            limits.max_input_bytes,
        ));
    }
    let input_reservation = context
        .reserve(Resource::InputBytes, u64_len(source.len())?)
        .map_err(map_execution)?;
    let mut ledger = Ledger::new(context, limits);
    ledger.charge_work(u64_len(source.len())?)?;
    let mut typed = parse_detective_inner(source, 0, namespace_context, limits, &mut ledger)?;
    if source_offset != 0 {
        typed = remap_typed_ranges(typed, source, source_offset, &mut ledger)?;
    }
    // The input and retained-memory Arc headers are part of the returned
    // typed allocation.  Charge both before moving the ledger token into the
    // wrapper, so clone/drop lifetimes release the exact standalone guards.
    let guard_bytes = arc_allocation_bytes::<Reservation>()?
        .checked_mul(2)
        .ok_or_else(|| invalid("ODS detective standalone guard size overflows"))?;
    ledger.reserve_memory(guard_bytes)?;
    let memory = ledger.memory.take().map(Arc::new);
    typed.with_guards(Some(Arc::new(input_reservation)), memory)
}

fn parse_detective_inner<'source>(
    source: &'source str,
    source_offset: usize,
    namespace_context: &NamespaceContext,
    limits: Limits,
    ledger: &mut Ledger<'_>,
) -> Result<TypedDetective<'source>> {
    ledger.begin_owner();
    let _owner_depth = ledger
        .context
        .reserve(Resource::Depth, 1)
        .map_err(map_execution)?;
    let mut reader = NsReader::from_str(source);
    reader.config_mut().check_end_names = true;
    reader.config_mut().trim_text(false);
    reader
        .resolver_mut()
        .set_max_declarations_per_element(limits.max_namespace_bindings);
    seed_resolver(reader.resolver_mut(), namespace_context)?;
    let mut buffer = Vec::new();
    let buffer_capacity = source.len().min(limits.max_text_bytes);
    let _buffer_memory = ledger
        .context
        .reserve(Resource::Memory, u64_len(buffer_capacity)?)
        .map_err(map_execution)?;
    buffer
        .try_reserve_exact(buffer_capacity)
        .map_err(|source| allocation("ODS detective owner reader buffer", source))?;
    let mut root_seen = false;
    let mut owner_start = 0usize;
    let mut owner_end = 0usize;
    let mut owner_prefix = String::new();
    let mut semantic = Detective::new();
    let mut highlighted_capacity = 0usize;
    let mut operation_capacity = 0usize;
    let mut ordered = Vec::new();
    let mut lexical_fidelity = LexicalFidelity::Canonical;
    let mut saw_root_end = false;
    let mut in_operations = false;

    loop {
        let event_start = position(&reader)?;
        let (resolved, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(xml_error)?;
        ledger.owner_event()?;
        let _owned_event_memory = ledger.reserve_event_memory(event.len())?;
        let element_namespace =
            if matches!(&event, Event::Start(_) | Event::Empty(_) | Event::End(_)) {
                resolve_namespace(&resolved, ledger)?
            } else {
                None
            };
        let event = event.into_owned();
        drop(resolved);
        match event {
            Event::Decl(declaration) => match declaration.xml_version().map_err(xml_error)? {
                XmlVersion::Explicit1_0 => {},
                XmlVersion::Implicit1_0 => {},
                XmlVersion::Explicit1_1 => {
                    return Err(unsupported("ODS detective codec accepts XML 1.0 only"));
                },
            },
            Event::Start(element) if !root_seen => {
                ensure_detective_name(element_namespace.as_deref(), &element)?;
                root_seen = true;
                owner_start = event_start;
                owner_prefix = prefix_string(element.name().prefix(), ledger)?;
                validate_owner_attributes(&reader, &element, limits, ledger)?;
            },
            Event::Empty(element) if !root_seen => {
                ensure_detective_name(element_namespace.as_deref(), &element)?;
                root_seen = true;
                owner_start = event_start;
                owner_end = position(&reader)?;
                owner_prefix = prefix_string(element.name().prefix(), ledger)?;
                validate_owner_attributes(&reader, &element, limits, ledger)?;
                saw_root_end = true;
            },
            Event::Start(element) if root_seen && !saw_root_end => {
                let child_start = event_start;
                let (child_kind, value) = parse_child_start(
                    &mut reader,
                    &element,
                    element_namespace.as_deref(),
                    limits,
                    ledger,
                )?;
                let child_end = position(&reader)?;
                let raw = RawRange::new(
                    source,
                    source_offset
                        .checked_add(child_start)
                        .ok_or_else(|| invalid("ODS detective child offset overflows"))?
                        ..source_offset
                            .checked_add(child_end)
                            .ok_or_else(|| invalid("ODS detective child end overflows"))?,
                )?;
                match (child_kind, value) {
                    (ChildKind::HighlightedRange, ChildValue::HighlightedRange(value)) => {
                        if in_operations {
                            return Err(invalid(
                                "ODS table:highlighted-range appears after table:operation",
                            ));
                        }
                        ensure_highlighted_range_capacity(
                            semantic.highlighted_ranges().len(),
                            limits,
                        )?;
                        let index = semantic.highlighted_ranges().len();
                        reserve_model_slot::<HighlightedRange>(
                            &semantic,
                            true,
                            &mut highlighted_capacity,
                            ledger,
                        )?;
                        semantic.add_highlighted_range(value);
                        reserve_ordered(&mut ordered, ledger)?;
                        ordered.push(OrderedChild::HighlightedRange { index, raw });
                    },
                    (ChildKind::Operation, ChildValue::Operation(value)) => {
                        ensure_operation_capacity(semantic.operations().len(), limits)?;
                        in_operations = true;
                        let index = semantic.operations().len();
                        reserve_model_slot::<Operation>(
                            &semantic,
                            false,
                            &mut operation_capacity,
                            ledger,
                        )?;
                        semantic.add_operation(value);
                        reserve_ordered(&mut ordered, ledger)?;
                        ordered.push(OrderedChild::Operation { index, raw });
                    },
                    _ => {
                        return Err(invalid(
                            "ODS detective child parser returned an invalid kind",
                        ));
                    },
                }
                lexical_fidelity = LexicalFidelity::PreservationRequired;
            },
            Event::Empty(element) if root_seen && !saw_root_end => {
                let child_start = event_start;
                let (child_kind, value) = parse_child_empty(
                    &reader,
                    &element,
                    element_namespace.as_deref(),
                    limits,
                    ledger,
                )?;
                let child_end = position(&reader)?;
                let raw = RawRange::new(
                    source,
                    source_offset
                        .checked_add(child_start)
                        .ok_or_else(|| invalid("ODS detective child offset overflows"))?
                        ..source_offset
                            .checked_add(child_end)
                            .ok_or_else(|| invalid("ODS detective child end overflows"))?,
                )?;
                match (child_kind, value) {
                    (ChildKind::HighlightedRange, ChildValue::HighlightedRange(value)) => {
                        if in_operations {
                            return Err(invalid(
                                "ODS table:highlighted-range appears after table:operation",
                            ));
                        }
                        ensure_highlighted_range_capacity(
                            semantic.highlighted_ranges().len(),
                            limits,
                        )?;
                        let index = semantic.highlighted_ranges().len();
                        reserve_model_slot::<HighlightedRange>(
                            &semantic,
                            true,
                            &mut highlighted_capacity,
                            ledger,
                        )?;
                        semantic.add_highlighted_range(value);
                        reserve_ordered(&mut ordered, ledger)?;
                        ordered.push(OrderedChild::HighlightedRange { index, raw });
                    },
                    (ChildKind::Operation, ChildValue::Operation(value)) => {
                        ensure_operation_capacity(semantic.operations().len(), limits)?;
                        in_operations = true;
                        let index = semantic.operations().len();
                        reserve_model_slot::<Operation>(
                            &semantic,
                            false,
                            &mut operation_capacity,
                            ledger,
                        )?;
                        semantic.add_operation(value);
                        reserve_ordered(&mut ordered, ledger)?;
                        ordered.push(OrderedChild::Operation { index, raw });
                    },
                    _ => {
                        return Err(invalid(
                            "ODS detective child parser returned an invalid kind",
                        ));
                    },
                }
            },
            Event::End(element) if root_seen && !saw_root_end => {
                if !is_detective_end(element_namespace.as_deref(), &element) {
                    return Err(invalid(
                        "ODS detective closing element does not match its owner",
                    ));
                }
                owner_end = position(&reader)?;
                saw_root_end = true;
            },
            Event::Text(text) if root_seen && !saw_root_end => {
                ensure_text_event_limit(text.len(), limits)?;
                ensure_whitespace_text(&text, XmlVersion::Explicit1_0, "detective")?;
                lexical_fidelity = LexicalFidelity::PreservationRequired;
            },
            Event::CData(text) if root_seen && !saw_root_end => {
                ensure_text_event_limit(text.len(), limits)?;
                ensure_whitespace_cdata(&text, XmlVersion::Explicit1_0, "detective CDATA")?;
                lexical_fidelity = LexicalFidelity::PreservationRequired;
            },
            Event::Comment(_) | Event::PI(_) if root_seen && !saw_root_end => {
                lexical_fidelity = LexicalFidelity::PreservationRequired;
            },
            Event::Eof if !root_seen => {
                return Err(invalid("ODS detective source has no table:detective root"));
            },
            Event::Eof if !saw_root_end => {
                return Err(invalid("unterminated ODS table:detective owner"));
            },
            Event::Eof => break,
            Event::Text(text) => {
                ensure_text_event_limit(text.len(), limits)?;
                if saw_root_end && !is_whitespace_event(source, event_start, position(&reader)?) {
                    return Err(invalid("ODS detective source has trailing content"));
                }
            },
            Event::CData(text) => {
                ensure_text_event_limit(text.len(), limits)?;
                if saw_root_end && !is_whitespace_event(source, event_start, position(&reader)?) {
                    return Err(invalid("ODS detective source has trailing content"));
                }
            },
            Event::Comment(_) | Event::PI(_) => {},
            Event::DocType(_) | Event::GeneralRef(_) => {
                if root_seen {
                    return Err(invalid(
                        "ODS detective owner contains unsupported XML content",
                    ));
                }
            },
            Event::Start(_) | Event::Empty(_) | Event::End(_) => {
                return Err(invalid("ODS detective owner contains an unsupported child"));
            },
        }
        if root_seen && saw_root_end {
            // The next event must be whitespace, a declaration that appeared
            // before the root was already handled, or EOF.  Non-whitespace
            // trailing text is rejected above.
        }
        buffer.clear();
    }

    if !root_seen || !saw_root_end {
        return Err(invalid("ODS detective owner is incomplete"));
    }
    if owner_end == 0 {
        owner_end = position(&reader)?;
    }
    let raw = RawRange::new(
        source,
        source_offset
            .checked_add(owner_start)
            .ok_or_else(|| invalid("ODS detective owner offset overflows"))?
            ..source_offset
                .checked_add(owner_end)
                .ok_or_else(|| invalid("ODS detective owner end overflows"))?,
    )?;
    reserve_namespace_context_copy(namespace_context, ledger)?;
    if owner_prefix.is_empty() {
        lexical_fidelity = LexicalFidelity::PreservationRequired;
    } else {
        let (canonical_len, rendered_len) = rendered_lengths(&semantic, &owner_prefix)?;
        if rendered_len > limits.max_output_bytes {
            return Err(limit(
                "detective owner output bytes",
                rendered_len,
                limits.max_output_bytes,
            ));
        }
        let render_scratch = if owner_prefix == "table" {
            canonical_len
        } else {
            canonical_len
                .checked_add(rendered_len)
                .ok_or_else(|| invalid("ODS detective owner render scratch overflows"))?
        };
        if render_scratch > limits.max_scratch_bytes {
            return Err(limit(
                "detective owner render scratch",
                render_scratch,
                limits.max_scratch_bytes,
            ));
        }
        ledger.charge_work(u64_len(rendered_len)?)?;
        let _render_memory = ledger
            .context
            .reserve(Resource::Memory, u64_len(render_scratch)?)
            .map_err(map_execution)?;
        let canonical = render_for_prefix_unchecked_with_capacities(
            &semantic,
            &owner_prefix,
            canonical_len,
            rendered_len,
        )?;
        if canonical.len() != rendered_len {
            return Err(limit(
                "detective owner output bytes",
                canonical.len(),
                rendered_len,
            ));
        }
        if raw.as_str() != canonical {
            lexical_fidelity = LexicalFidelity::PreservationRequired;
        }
    }
    ledger.charge_work(64)?;
    ledger.reserve_memory(arc_allocation_bytes::<RetainedDetective<'_>>()?)?;
    Ok(TypedDetective {
        retained: Arc::new(RetainedDetective {
            value: semantic,
            raw,
            ordered,
            namespace_context: namespace_context.clone(),
            source_prefix: owner_prefix,
            lexical_fidelity,
            input_guard: None,
            memory_guard: None,
        }),
    })
}

fn remap_typed_ranges<'source>(
    mut typed: TypedDetective<'source>,
    source: &'source str,
    source_offset: usize,
    ledger: &mut Ledger<'_>,
) -> Result<TypedDetective<'source>> {
    let retained = Arc::get_mut(&mut typed.retained).ok_or_else(|| {
        unsupported("ODS detective typed state was unexpectedly shared while remapping ranges")
    })?;
    let raw = retained.raw;
    let raw_start = source_offset
        .checked_add(raw.range().start)
        .ok_or_else(|| invalid("ODS detective owner source offset overflows"))?;
    let raw_end = source_offset
        .checked_add(raw.range().end)
        .ok_or_else(|| invalid("ODS detective owner source end overflows"))?;
    let raw = RawRange::new(source, raw_start..raw_end)?;
    let mut remapped = Vec::new();
    ledger.reserve_memory(u64_len(
        retained
            .ordered
            .len()
            .checked_mul(std::mem::size_of::<OrderedChild<'_>>())
            .ok_or_else(|| invalid("ODS detective ordered source range memory overflows"))?,
    )?)?;
    remapped
        .try_reserve_exact(retained.ordered.len())
        .map_err(|error| allocation("ODS detective ordered source ranges", error))?;
    for child in retained.ordered.iter().copied() {
        let child_raw = child.raw();
        let start = source_offset
            .checked_add(child_raw.range().start)
            .ok_or_else(|| invalid("ODS detective child source offset overflows"))?;
        let end = source_offset
            .checked_add(child_raw.range().end)
            .ok_or_else(|| invalid("ODS detective child source end overflows"))?;
        let child_raw = RawRange::new(source, start..end)?;
        let child = match child {
            OrderedChild::HighlightedRange { index, .. } => OrderedChild::HighlightedRange {
                index,
                raw: child_raw,
            },
            OrderedChild::Operation { index, .. } => OrderedChild::Operation {
                index,
                raw: child_raw,
            },
        };
        remapped.push(child);
    }
    retained.raw = raw;
    retained.ordered = remapped;
    Ok(typed)
}

#[derive(Clone, Copy)]
enum ChildKind {
    HighlightedRange,
    Operation,
}

enum ChildValue {
    HighlightedRange(HighlightedRange),
    Operation(Operation),
}

fn parse_child_start(
    reader: &mut NsReader<&[u8]>,
    element: &BytesStart<'_>,
    namespace: Option<&str>,
    limits: Limits,
    ledger: &mut Ledger<'_>,
) -> Result<(ChildKind, ChildValue)> {
    let kind = child_kind(namespace, element)?;
    let value = match kind {
        ChildKind::HighlightedRange => {
            parse_highlighted_attributes(reader, element, limits, ledger)?
        },
        ChildKind::Operation => parse_operation_attributes(reader, element, limits, ledger)?,
    };
    let expected = element.name();
    let mut buffer = Vec::new();
    let _child_buffer_memory = ledger
        .context
        .reserve(Resource::Memory, u64_len(limits.max_text_bytes)?)
        .map_err(map_execution)?;
    let _child_depth = ledger
        .context
        .reserve(Resource::Depth, 1)
        .map_err(map_execution)?;
    buffer
        .try_reserve_exact(limits.max_text_bytes)
        .map_err(|source| allocation("ODS detective child reader buffer", source))?;
    loop {
        let (resolved_end, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(xml_error)?;
        ledger.owner_event()?;
        match event {
            Event::End(end) if end.name() == expected => break,
            Event::Text(text) => {
                ensure_text_event_limit(text.len(), limits)?;
                ensure_whitespace_text(&text, XmlVersion::Explicit1_0, "detective child")?
            },
            Event::CData(text) => {
                ensure_text_event_limit(text.len(), limits)?;
                ensure_whitespace_cdata(&text, XmlVersion::Explicit1_0, "detective child CDATA")?
            },
            Event::Comment(_) | Event::PI(_) => {},
            Event::Start(_) | Event::Empty(_) | Event::End(_) => {
                let _ = resolved_end;
                return Err(invalid(
                    "ODS detective highlighted-range/operation must be empty",
                ));
            },
            Event::Eof => return Err(invalid("unterminated ODS detective child")),
            Event::Decl(_) | Event::DocType(_) | Event::GeneralRef(_) => {
                return Err(invalid(
                    "ODS detective child contains unsupported XML content",
                ));
            },
        }
        buffer.clear();
    }
    Ok((kind, value))
}

fn parse_child_empty(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    namespace: Option<&str>,
    limits: Limits,
    ledger: &mut Ledger<'_>,
) -> Result<(ChildKind, ChildValue)> {
    let kind = child_kind(namespace, element)?;
    let value = match kind {
        ChildKind::HighlightedRange => {
            parse_highlighted_attributes(reader, element, limits, ledger)?
        },
        ChildKind::Operation => parse_operation_attributes(reader, element, limits, ledger)?,
    };
    Ok((kind, value))
}

fn child_kind(namespace: Option<&str>, element: &BytesStart<'_>) -> Result<ChildKind> {
    if is_table_element(
        namespace,
        element.local_name().as_ref(),
        b"highlighted-range",
    ) {
        Ok(ChildKind::HighlightedRange)
    } else if is_table_element(namespace, element.local_name().as_ref(), b"operation") {
        Ok(ChildKind::Operation)
    } else {
        Err(invalid(
            "ODS table:detective contains an unsupported direct child",
        ))
    }
}

fn parse_highlighted_attributes(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    limits: Limits,
    ledger: &mut Ledger<'_>,
) -> Result<ChildValue> {
    let mut address = None;
    let mut direction = None;
    let mut contains_error = None;
    let mut marked_invalid = None;
    collect_attributes(reader, element, limits, ledger, |attribute| {
        if attribute.namespace == TABLE_NAMESPACE {
            match attribute.local.as_str() {
                "cell-range-address" => {
                    set_once(&mut address, attribute.value, "cell-range-address")?
                },
                "direction" => set_once(&mut direction, attribute.value, "direction")?,
                "contains-error" => {
                    let value = parse_bool(&attribute.value, "table:contains-error")?;
                    set_once(&mut contains_error, value, "contains-error")?;
                },
                "marked-invalid" => {
                    let value = parse_bool(&attribute.value, "table:marked-invalid")?;
                    set_once(&mut marked_invalid, value, "marked-invalid")?;
                },
                _ => {
                    return Err(unsupported(
                        "ODS highlighted-range has an unknown table attribute",
                    ));
                },
            }
        } else {
            return Err(unsupported(
                "ODS highlighted-range has an attribute outside the table namespace",
            ));
        }
        Ok(())
    })?;
    if let Some(marked_invalid) = marked_invalid {
        if address.is_some() || direction.is_some() || contains_error.is_some() {
            return Err(invalid(
                "ODS highlighted-range cannot combine marked-invalid with valid-range attributes",
            ));
        }
        return Ok(ChildValue::HighlightedRange(HighlightedRange::invalid(
            marked_invalid,
        )));
    }
    let direction = direction.ok_or_else(|| {
        invalid("ODS highlighted-range requires table:direction or table:marked-invalid")
    })?;
    let direction = Direction::parse(&direction)?;
    let value = HighlightedRange::valid(address, direction, contains_error)?;
    Ok(ChildValue::HighlightedRange(value))
}

fn parse_operation_attributes(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    limits: Limits,
    ledger: &mut Ledger<'_>,
) -> Result<ChildValue> {
    let mut name = None;
    let mut index = None;
    collect_attributes(reader, element, limits, ledger, |attribute| {
        if attribute.namespace != TABLE_NAMESPACE {
            return Err(unsupported(
                "ODS table:operation has an attribute outside the table namespace",
            ));
        }
        match attribute.local.as_str() {
            "name" => set_once(&mut name, attribute.value, "name")?,
            "index" => {
                let value = attribute.value.parse::<usize>().map_err(|_| {
                    invalid("ODS table:operation table:index must be a non-negative integer")
                })?;
                set_once(&mut index, value, "index")?;
            },
            _ => {
                return Err(unsupported(
                    "ODS table:operation has an unknown table attribute",
                ));
            },
        }
        Ok(())
    })?;
    let name = name.ok_or_else(|| invalid("ODS table:operation requires table:name"))?;
    let index = index.ok_or_else(|| invalid("ODS table:operation requires table:index"))?;
    let kind = OperationKind::parse(&name)?;
    Ok(ChildValue::Operation(Operation::new(kind, index)))
}

#[derive(Clone)]
struct ExpandedAttribute {
    namespace: String,
    local: String,
    value: String,
}

fn collect_attributes(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    limits: Limits,
    ledger: &mut Ledger<'_>,
    mut visit: impl FnMut(ExpandedAttribute) -> Result<()>,
) -> Result<()> {
    let mut seen = Vec::<(String, String)>::new();
    let mut namespace_declarations = Vec::<String>::new();
    for raw in element.attributes().with_checks(true) {
        ledger.check()?;
        let raw =
            raw.map_err(|error| invalid(format!("invalid ODS detective attribute: {error}")))?;
        let key = raw.key.as_ref();
        if key == b"xmlns" || key.starts_with(b"xmlns:") {
            let prefix = if key == b"xmlns" {
                String::new()
            } else {
                decode_with_ledger(key.get(6..).unwrap_or_default(), "namespace prefix", ledger)?
            };
            if namespace_declarations.iter().any(|value| value == &prefix) {
                return Err(invalid(
                    "duplicate XML namespace declaration on detective child",
                ));
            }
            reserve_vec_slot(
                &mut namespace_declarations,
                ledger,
                "ODS detective namespace declaration",
            )?;
            namespace_declarations.push(prefix);
            continue;
        }
        if raw.value.len() > limits.max_text_bytes {
            return Err(limit(
                "detective attribute bytes",
                raw.value.len(),
                limits.max_text_bytes,
            ));
        }
        let raw_value_bytes = raw.value.len();
        // `into_owned` may allocate while decoding entities or normalizing
        // whitespace.  Charge a bounded capacity before that allocation; a
        // decoded value longer than its source is charged for the excess
        // after conversion as a defensive check.
        ledger.reserve_memory(u64_len(raw_value_bytes)?)?;
        let value = raw
            .decoded_and_normalized_value(XmlVersion::Explicit1_0, reader.decoder())
            .map_err(|error| invalid(format!("invalid ODS detective attribute value: {error}")))?
            .into_owned();
        if value.len() > limits.max_text_bytes {
            return Err(limit(
                "detective decoded attribute bytes",
                value.len(),
                limits.max_text_bytes,
            ));
        }
        if value.len() > raw_value_bytes {
            ledger.reserve_memory(u64_len(value.len() - raw_value_bytes)?)?;
        }
        let (namespace, local) = reader.resolver().resolve_attribute(raw.key);
        let namespace = match namespace {
            ResolveResult::Bound(Namespace(uri)) => {
                decode_with_ledger(uri, "attribute namespace", ledger)?
            },
            ResolveResult::Unbound => String::new(),
            ResolveResult::Unknown(prefix) => {
                return Err(invalid(format!(
                    "ODS detective attribute uses unbound prefix '{}'",
                    String::from_utf8_lossy(prefix.as_ref())
                )));
            },
        };
        let local = decode_with_ledger(local.as_ref(), "attribute local name", ledger)?;
        if seen.iter().any(|(seen_namespace, seen_local)| {
            seen_namespace == &namespace && seen_local == &local
        }) {
            return Err(invalid(format!(
                "duplicate expanded ODS detective attribute '{}:{}'",
                namespace, local
            )));
        }
        let name_bytes = namespace
            .len()
            .checked_add(local.len())
            .ok_or_else(|| invalid("ODS detective expanded attribute name size overflows"))?;
        ledger.reserve_memory(u64_len(name_bytes)?)?;
        reserve_vec_slot(&mut seen, ledger, "ODS detective expanded attributes")?;
        seen.push((namespace.clone(), local.clone()));
        visit(ExpandedAttribute {
            namespace,
            local,
            value,
        })?;
    }
    Ok(())
}

fn validate_owner_attributes(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    limits: Limits,
    ledger: &mut Ledger<'_>,
) -> Result<()> {
    collect_attributes(reader, element, limits, ledger, |_attribute| {
        Err(unsupported("ODS table:detective must not have attributes"))
    })
}

fn scan_source<'source>(source: &'source str, ledger: &mut Ledger<'_>) -> Result<Scanned<'source>> {
    let limits = ledger.limits;
    let mut reader = NsReader::from_str(source);
    reader.config_mut().check_end_names = true;
    reader.config_mut().trim_text(false);
    reader
        .resolver_mut()
        .set_max_declarations_per_element(limits.max_namespace_bindings);
    let mut buffer = Vec::new();
    let buffer_capacity = source.len().min(limits.max_text_bytes);
    let _buffer_memory = ledger
        .context
        .reserve(Resource::Memory, u64_len(buffer_capacity)?)
        .map_err(map_execution)?;
    buffer
        .try_reserve_exact(buffer_capacity)
        .map_err(|source| allocation("ODS detective reader buffer", source))?;
    let mut spans = Vec::<Span>::new();
    let mut open = Vec::<OpenFrame>::new();
    let mut root_seen = false;
    let mut root_closed = false;
    let mut elements = 0usize;

    loop {
        let event_start = position(&reader)?;
        let (resolved, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(xml_error)?;
        ledger.event()?;
        let _owned_event_memory = ledger.reserve_event_memory(event.len())?;
        let element_namespace = if matches!(&event, Event::Start(_) | Event::Empty(_)) {
            resolve_namespace(&resolved, ledger)?
        } else {
            None
        };
        let event = event.into_owned();
        drop(resolved);
        let event_end = position(&reader)?;
        match event {
            Event::Decl(declaration) => match declaration.xml_version().map_err(xml_error)? {
                XmlVersion::Implicit1_0 | XmlVersion::Explicit1_0 => {},
                XmlVersion::Explicit1_1 => {
                    return Err(unsupported(
                        "ODS detective source codec accepts XML 1.0 only",
                    ));
                },
            },
            Event::Start(element) => {
                elements = elements
                    .checked_add(1)
                    .ok_or_else(|| invalid("ODS detective XML event count overflows"))?;
                if elements > limits.max_events {
                    return Err(limit("XML events", elements, limits.max_events));
                }
                if root_closed {
                    return Err(invalid("ODS detective source has multiple XML roots"));
                }
                let next_depth = open
                    .len()
                    .checked_add(1)
                    .ok_or_else(|| invalid("ODS detective XML depth overflows"))?;
                if next_depth > limits.max_depth {
                    return Err(limit("XML depth", next_depth, limits.max_depth));
                }
                if !root_seen {
                    root_seen = true;
                }
                let span = make_span(
                    &reader,
                    element_namespace,
                    &element,
                    event_start,
                    event_end,
                    false,
                    open.last().map(|frame| frame.index),
                    &open,
                    ledger,
                )?;
                let index = push_span(&mut spans, span, limits, ledger)?;
                if let Some(parent) = open.last().map(|frame| frame.index) {
                    reserve_vec_slot(
                        &mut spans[parent].children,
                        ledger,
                        "ODS detective child index",
                    )?;
                    spans[parent].children.push(index);
                }
                let depth_reservation = ledger
                    .context
                    .reserve(Resource::Depth, 1)
                    .map_err(map_execution)?;
                reserve_vec_slot(&mut open, ledger, "ODS detective XML stack")?;
                open.push(OpenFrame {
                    index,
                    depth_reservation,
                });
            },
            Event::Empty(element) => {
                elements = elements
                    .checked_add(1)
                    .ok_or_else(|| invalid("ODS detective XML event count overflows"))?;
                if elements > limits.max_events {
                    return Err(limit("XML events", elements, limits.max_events));
                }
                if root_closed {
                    return Err(invalid("ODS detective source has multiple XML roots"));
                }
                if !root_seen {
                    root_seen = true;
                    root_closed = true;
                }
                let span = make_span(
                    &reader,
                    element_namespace,
                    &element,
                    event_start,
                    event_end,
                    true,
                    open.last().map(|frame| frame.index),
                    &open,
                    ledger,
                )?;
                let index = push_span(&mut spans, span, limits, ledger)?;
                if let Some(parent) = open.last().map(|frame| frame.index) {
                    reserve_vec_slot(
                        &mut spans[parent].children,
                        ledger,
                        "ODS detective child index",
                    )?;
                    spans[parent].children.push(index);
                }
            },
            Event::End(element) => {
                let frame = open
                    .pop()
                    .ok_or_else(|| invalid("ODS detective XML element stack underflow"))?;
                let span = spans
                    .get_mut(frame.index)
                    .ok_or_else(|| invalid("ODS detective XML span disappeared"))?;
                if element.name().as_ref() != span.qname.as_bytes() {
                    return Err(invalid("ODS detective XML closing name does not match"));
                }
                span.close_start = event_start;
                span.end = event_end;
                drop(frame.depth_reservation);
                if open.is_empty() {
                    root_closed = true;
                }
            },
            Event::Eof => break,
            Event::Text(text) => {
                ensure_text_event_limit(text.len(), limits)?;
                if event_has_non_whitespace_text(&Event::Text(text))? {
                    mark_direct_cell_text(&mut spans, open.last().map(|frame| frame.index))?;
                }
            },
            Event::CData(text) => {
                ensure_text_event_limit(text.len(), limits)?;
                if event_has_non_whitespace_text(&Event::CData(text))? {
                    mark_direct_cell_text(&mut spans, open.last().map(|frame| frame.index))?;
                }
            },
            Event::Comment(_) | Event::PI(_) | Event::DocType(_) | Event::GeneralRef(_) => {},
        }
        buffer.clear();
    }
    if !open.is_empty() {
        return Err(invalid(
            "ODS detective source contains unclosed XML elements",
        ));
    }
    if !root_seen || !root_closed {
        return Err(invalid("ODS detective source has no complete XML root"));
    }
    Ok(Scanned::from_spans(spans))
}

struct Scanned<'source> {
    spans: Vec<Span>,
    _marker: std::marker::PhantomData<&'source str>,
}

impl<'source> Scanned<'source> {
    fn from_spans(spans: Vec<Span>) -> Self {
        Self {
            spans,
            _marker: std::marker::PhantomData,
        }
    }
}

struct OpenFrame {
    index: usize,
    depth_reservation: Reservation,
}

#[derive(Clone, Debug)]
struct Span {
    namespace: Option<String>,
    local: String,
    qname: String,
    start: usize,
    close_start: usize,
    end: usize,
    empty: bool,
    parent: Option<usize>,
    foreign_ancestor: bool,
    mce_ancestor: bool,
    namespace_context: Option<NamespaceContext>,
    raw_text_non_whitespace: bool,
    children: Vec<usize>,
}

#[allow(
    clippy::too_many_arguments,
    reason = "the source span constructor receives every bounded scanner fact explicitly"
)]
fn make_span(
    reader: &NsReader<&[u8]>,
    namespace: Option<String>,
    element: &BytesStart<'_>,
    start: usize,
    tag_end: usize,
    empty: bool,
    parent: Option<usize>,
    open: &[OpenFrame],
    ledger: &mut Ledger<'_>,
) -> Result<Span> {
    let local = decode_with_ledger(element.local_name().as_ref(), "element local name", ledger)?;
    let qname = decode_with_ledger(element.name().as_ref(), "element qualified name", ledger)?;
    let foreign_self = !is_supported_ancestor(namespace.as_deref(), local.as_str());
    let foreign_ancestor = foreign_self
        || open.iter().any(|frame| {
            // The parent span has already been materialized.  The caller uses the
            // stack only for depth; ancestry is corrected below from names in the
            // span table, so this branch is intentionally conservative for a
            // foreign current element.
            let _ = frame;
            false
        });
    let mce_self = namespace.as_deref() == Some(MC_NAMESPACE);
    let capture_context = is_cell_element(namespace.as_deref(), local.as_str())
        || is_detective_element(namespace.as_deref(), local.as_str());
    let namespace_context = if capture_context {
        Some(namespace_context(reader.resolver(), ledger.limits, ledger)?)
    } else {
        None
    };
    let _ = tag_end;
    Ok(Span {
        namespace,
        local,
        qname,
        start,
        close_start: tag_end,
        end: tag_end,
        empty,
        parent,
        foreign_ancestor,
        mce_ancestor: mce_self,
        namespace_context,
        raw_text_non_whitespace: false,
        children: Vec::new(),
    })
}

fn push_span(
    spans: &mut Vec<Span>,
    mut span: Span,
    limits: Limits,
    ledger: &mut Ledger<'_>,
) -> Result<usize> {
    // Resolve foreign ancestry after the direct parent is available.  This
    // keeps MCE/foreign wrappers opaque instead of making them transparent.
    if let Some(parent) = span.parent {
        let parent_span = spans
            .get(parent)
            .ok_or_else(|| invalid("ODS detective parent span is invalid"))?;
        span.foreign_ancestor |= parent_span.foreign_ancestor
            || !is_supported_ancestor(parent_span.namespace.as_deref(), parent_span.local.as_str());
        span.mce_ancestor |=
            parent_span.mce_ancestor || parent_span.namespace.as_deref() == Some(MC_NAMESPACE);
    }
    if spans.len() >= limits.max_events {
        return Err(limit(
            "XML spans",
            spans.len().saturating_add(1),
            limits.max_events,
        ));
    }
    ledger.reserve_memory(u64_len(128)?)?;
    ledger
        .context
        .consume(Resource::Objects, 1)
        .map_err(map_execution)?;
    reserve_vec_slot(spans, ledger, "ODS detective XML span index")?;
    let index = spans.len();
    spans.push(span);
    Ok(index)
}

fn build_cells<'source>(
    source: &'source str,
    spans: &[Span],
    ledger: &mut Ledger<'_>,
) -> Result<(Vec<CellSite<'source>>, Vec<Option<CellId>>)> {
    let mut cells = Vec::new();
    ledger.reserve_memory(u64_len(
        spans
            .len()
            .checked_mul(std::mem::size_of::<Option<CellId>>())
            .ok_or_else(|| invalid("ODS detective cell selector memory overflows"))?,
    )?)?;
    let mut cell_by_span = Vec::new();
    cell_by_span
        .try_reserve_exact(spans.len())
        .map_err(|source| allocation("ODS detective cell selector index", source))?;
    ledger
        .context
        .consume(Resource::Objects, u64_len(spans.len())?)
        .map_err(map_execution)?;
    cell_by_span.resize(spans.len(), None);
    for (span_index, span) in spans.iter().enumerate() {
        ledger.charge_work(8)?;
        let Some(kind) = cell_kind(span.namespace.as_deref(), span.local.as_str()) else {
            continue;
        };
        let parent = span.parent;
        let ancestors = ancestors_for_span(spans, parent, ledger)?;
        let context = if let Some(context) = span.namespace_context.as_ref() {
            reserve_namespace_context_copy(context, ledger)?;
            context.clone()
        } else {
            NamespaceContext::default()
        };
        let legal_ancestry = legal_cell_ancestry(spans, Some(span_index))?;
        let qualification = SourceQualification {
            parent: kind,
            ancestors,
            namespace_context: context,
            foreign_ancestor: span.foreign_ancestor || !legal_ancestry,
            markup_compatibility_ancestor: span.mce_ancestor,
        };
        let mut order = 0u8;
        let mut sequence_valid = true;
        let mut has_unsupported_child = false;
        let mut detective_count = 0usize;
        let mut annotation_count = 0usize;
        let mut source_count = 0usize;
        let mut insertion_anchor = span.close_start;
        if span.raw_text_non_whitespace {
            sequence_valid = false;
        }
        for child_index in &span.children {
            let child = spans
                .get(*child_index)
                .ok_or_else(|| invalid("ODS detective cell child span is invalid"))?;
            if !matches!(child_kind_for_cell(child), CellChildKind::Detective)
                && contains_detective_descendant(spans, *child_index, ledger)?
            {
                has_unsupported_child = true;
            }
            match child_kind_for_cell(child) {
                CellChildKind::Source => {
                    source_count = source_count.saturating_add(1);
                    if source_count > 1 {
                        sequence_valid = false;
                    }
                    if order > 0 {
                        sequence_valid = false;
                    }
                    order = order.max(1);
                },
                CellChildKind::Annotation => {
                    if order > 1 {
                        sequence_valid = false;
                    }
                    annotation_count = annotation_count.saturating_add(1);
                    if annotation_count > 1 {
                        sequence_valid = false;
                    }
                    order = order.max(2);
                },
                CellChildKind::Detective => {
                    if order > 2 {
                        sequence_valid = false;
                    }
                    order = order.max(3);
                    detective_count = detective_count.saturating_add(1);
                },
                CellChildKind::Text => {
                    order = order.max(4);
                    insertion_anchor = insertion_anchor.min(child.start);
                },
                CellChildKind::Unsupported => {
                    has_unsupported_child = true;
                },
            }
        }
        let id = CellId(cells.len());
        cell_by_span[span_index] = Some(id);
        let raw = RawRange::new(source, span.start..span.end)?;
        ledger.reserve_memory(u64_len(span.qname.len())?)?;
        reserve_vec_slot(&mut cells, ledger, "ODS detective cell index")?;
        ledger
            .context
            .consume(Resource::Objects, 1)
            .map_err(map_execution)?;
        cells.push(CellSite {
            id,
            raw,
            qname: span.qname.clone(),
            empty: span.empty,
            kind,
            qualification,
            detective: None,
            direct_owner_count: detective_count,
            sequence_valid,
            has_unsupported_child,
            insertion_anchor,
        });
    }
    Ok((cells, cell_by_span))
}

fn build_owners<'source>(
    source: &'source str,
    spans: &[Span],
    cells: &[CellSite<'source>],
    cell_by_span: &[Option<CellId>],
    limits: Limits,
    ledger: &mut Ledger<'_>,
) -> Result<Vec<DetectiveOwner<'source>>> {
    let mut owners = Vec::new();
    for span in spans {
        if !is_detective_element(span.namespace.as_deref(), span.local.as_str()) {
            continue;
        }
        let Some(parent_span) = span.parent else {
            continue;
        };
        let Some(cell) = cell_by_span.get(parent_span).and_then(|value| *value) else {
            // A detective under an opaque/foreign wrapper is intentionally not
            // made transparent and therefore is not an effective owner.
            continue;
        };
        if owners.len() >= limits.max_owners {
            return Err(limit(
                "detective owners",
                owners.len().saturating_add(1),
                limits.max_owners,
            ));
        }
        let site = cells
            .get(cell.0)
            .ok_or_else(|| invalid("ODS detective owner cell index is invalid"))?;
        let parent_span_index = parent_span;
        let parent_span = spans
            .get(parent_span_index)
            .ok_or_else(|| invalid("ODS detective owner parent span is invalid"))?;
        let ancestors = ancestors_for_span(spans, Some(parent_span_index), ledger)?;
        // Namespace declarations on the owner itself are resolved by the
        // owner parser, but the typed/qualification context must remain the
        // inherited context of the physical cell.  A local owner declaration
        // cannot silently become available after an owner-only splice.
        reserve_namespace_context_copy(&site.qualification.namespace_context, ledger)?;
        let context = site.qualification.namespace_context.clone();
        reserve_namespace_context_copy(&context, ledger)?;
        let qualification = SourceQualification {
            parent: site.kind,
            ancestors,
            namespace_context: context.clone(),
            foreign_ancestor: span.foreign_ancestor || parent_span.foreign_ancestor,
            markup_compatibility_ancestor: span.mce_ancestor || parent_span.mce_ancestor,
        };
        let raw = RawRange::new(source, span.start..span.end)?;
        ledger.reserve_memory(u64_len(span.qname.len())?)?;
        let duplicate = site.direct_owner_count > 1;
        let reason = if qualification.markup_compatibility_ancestor {
            Some(OpaqueReason::MarkupCompatibilityBranch)
        } else if qualification.foreign_ancestor {
            Some(OpaqueReason::ForeignAncestry)
        } else if duplicate {
            Some(OpaqueReason::DuplicateDirectOwner)
        } else if !site.sequence_valid {
            Some(OpaqueReason::InvalidCellChildOrder)
        } else {
            None
        };
        let state = if let Some(reason) = reason {
            OwnerState::Opaque { reason }
        } else {
            ledger.charge_work(64)?;
            match parse_detective_inner(raw.as_str(), 0, &context, limits, ledger) {
                Ok(typed) => match remap_typed_ranges(typed, source, span.start, ledger) {
                    Ok(mut typed) => {
                        // A table prefix declared only on the owner would be
                        // removed by a focused owner splice.  Keep the owner
                        // readable, but refuse a changed render unless the
                        // inherited cell context already carries that prefix.
                        if !site
                            .qualification
                            .namespace_context
                            .has_table_binding(typed.source_prefix())
                        {
                            typed.mark_preservation_required()?;
                        }
                        OwnerState::Typed(typed)
                    },
                    Err(error) => OwnerState::Opaque {
                        reason: OpaqueReason::Malformed(error.to_string()),
                    },
                },
                Err(error) => OwnerState::Opaque {
                    reason: OpaqueReason::Malformed(error.to_string()),
                },
            }
        };
        let id = OwnerId(owners.len());
        reserve_vec_slot(&mut owners, ledger, "ODS detective owner index")?;
        ledger
            .context
            .consume(Resource::Objects, 1)
            .map_err(map_execution)?;
        owners.push(DetectiveOwner {
            id,
            cell,
            raw,
            qualification,
            state,
        });
    }
    Ok(owners)
}

fn attach_owner_ids(cells: &mut [CellSite<'_>], owners: &[DetectiveOwner<'_>]) -> Result<()> {
    for owner in owners {
        let site = cells
            .get_mut(owner.cell.0)
            .ok_or_else(|| invalid("ODS detective owner cell disappeared while indexing"))?;
        if site.detective.is_none() {
            site.detective = Some(owner.id);
        }
    }
    Ok(())
}

fn insertion_anchor(site: &CellSite<'_>) -> Result<Range<usize>> {
    let insertion = site.insertion_anchor;
    Ok(insertion..insertion)
}

fn preflight_candidate_lengths(
    source_len: usize,
    old_len: usize,
    replacement_len: usize,
    limits: Limits,
) -> Result<(usize, usize)> {
    if old_len > source_len {
        return Err(invalid("ODS detective splice range exceeds its source"));
    }
    let output_len = source_len
        .checked_sub(old_len)
        .and_then(|length| length.checked_add(replacement_len))
        .ok_or_else(|| invalid("ODS detective candidate length overflows"))?;
    if output_len > limits.max_output_bytes {
        return Err(limit("output bytes", output_len, limits.max_output_bytes));
    }
    let scratch = output_len
        .checked_add(replacement_len)
        .ok_or_else(|| invalid("ODS detective candidate scratch size overflows"))?;
    if scratch > limits.max_scratch_bytes {
        return Err(limit("scratch bytes", scratch, limits.max_scratch_bytes));
    }
    Ok((output_len, scratch))
}

fn expanded_empty_cell_len(site: &CellSite<'_>, detective_len: usize) -> Result<usize> {
    let raw = site.raw.as_str();
    if !raw.ends_with("/>") {
        return Err(invalid(
            "ODS detective empty cell has no empty-element terminator",
        ));
    }
    raw.len()
        .checked_add(detective_len)
        .and_then(|length| length.checked_add(site.qname.len()))
        .and_then(|length| length.checked_add(2))
        .ok_or_else(|| invalid("ODS detective expanded cell length overflows"))
}

fn expand_empty_cell(
    site: &CellSite<'_>,
    detective: &str,
    context: &ExecutionContext,
) -> Result<(String, Reservation)> {
    let raw = site.raw.as_str();
    let slash = raw
        .rfind("/>")
        .ok_or_else(|| invalid("ODS detective empty cell has no empty-element terminator"))?;
    if slash + 2 != raw.len() {
        return Err(invalid(
            "ODS detective empty cell source range has trailing bytes",
        ));
    }
    let closing_len = site
        .qname
        .len()
        .checked_add(3)
        .ok_or_else(|| invalid("ODS detective expanded cell length overflows"))?;
    let capacity = raw
        .len()
        .checked_add(detective.len())
        .and_then(|length| length.checked_add(closing_len.saturating_sub(1)))
        .ok_or_else(|| invalid("ODS detective expanded cell length overflows"))?;
    let memory = context
        .reserve(Resource::Memory, u64_len(capacity)?)
        .map_err(map_execution)?;
    let mut expanded = String::new();
    expanded
        .try_reserve_exact(capacity)
        .map_err(|source| allocation("ODS detective expanded cell", source))?;
    expanded.push_str(&raw[..slash]);
    expanded.push('>');
    expanded.push_str(detective);
    expanded.push_str("</");
    expanded.push_str(&site.qname);
    expanded.push('>');
    if expanded.len() != capacity {
        return Err(invalid("ODS detective expanded cell length changed"));
    }
    Ok((expanded, memory))
}

fn contains_detective_descendant(
    spans: &[Span],
    root: usize,
    ledger: &mut Ledger<'_>,
) -> Result<bool> {
    ledger.charge_work(1)?;
    let span = spans
        .get(root)
        .ok_or_else(|| invalid("ODS detective descendant span is invalid"))?;
    for child in &span.children {
        let child_span = spans
            .get(*child)
            .ok_or_else(|| invalid("ODS detective descendant child span is invalid"))?;
        if is_detective_element(child_span.namespace.as_deref(), child_span.local.as_str())
            || contains_detective_descendant(spans, *child, ledger)?
        {
            return Ok(true);
        }
    }
    Ok(false)
}

fn ancestors_for_span(
    spans: &[Span],
    mut current: Option<usize>,
    ledger: &mut Ledger<'_>,
) -> Result<Vec<Ancestor>> {
    let mut reversed = Vec::new();
    while let Some(index) = current {
        let span = spans
            .get(index)
            .ok_or_else(|| invalid("ODS detective ancestry span is invalid"))?;
        let string_bytes = span
            .namespace
            .as_ref()
            .map_or(0, String::len)
            .checked_add(span.local.len())
            .ok_or_else(|| invalid("ODS detective ancestry string size overflows"))?;
        ledger.reserve_memory(u64_len(string_bytes)?)?;
        reserve_vec_slot(&mut reversed, ledger, "ODS detective ancestry")?;
        ledger
            .context
            .consume(Resource::Objects, 1)
            .map_err(map_execution)?;
        reversed.push(Ancestor {
            namespace: span.namespace.clone(),
            local_name: span.local.clone(),
        });
        current = span.parent;
    }
    reversed.reverse();
    Ok(reversed)
}

#[derive(Clone, Copy)]
enum CellChildKind {
    Source,
    Annotation,
    Detective,
    Text,
    Unsupported,
}

fn child_kind_for_cell(span: &Span) -> CellChildKind {
    if span.namespace.as_deref() == Some(TABLE_NAMESPACE) && span.local == "cell-range-source" {
        CellChildKind::Source
    } else if span.namespace.as_deref() == Some(OFFICE_NAMESPACE) && span.local == "annotation" {
        CellChildKind::Annotation
    } else if is_detective_element(span.namespace.as_deref(), span.local.as_str()) {
        CellChildKind::Detective
    } else if span.namespace.as_deref() == Some(TEXT_NAMESPACE)
        && is_legal_text_content_local(span.local.as_str())
    {
        CellChildKind::Text
    } else {
        CellChildKind::Unsupported
    }
}

fn cell_kind(namespace: Option<&str>, local: &str) -> Option<CellKind> {
    if namespace != Some(TABLE_NAMESPACE) {
        return None;
    }
    match local {
        "table-cell" => Some(CellKind::TableCell),
        "covered-table-cell" => Some(CellKind::CoveredTableCell),
        _ => None,
    }
}

fn is_cell_element(namespace: Option<&str>, local: &str) -> bool {
    cell_kind(namespace, local).is_some()
}

fn is_detective_element(namespace: Option<&str>, local: &str) -> bool {
    namespace == Some(TABLE_NAMESPACE) && local == "detective"
}

fn is_supported_ancestor(namespace: Option<&str>, local: &str) -> bool {
    match namespace {
        Some(OFFICE_NAMESPACE) => matches!(local, "document-content" | "body" | "spreadsheet"),
        Some(TABLE_NAMESPACE) => matches!(
            local,
            "table"
                | "table-row"
                | "table-header-rows"
                | "table-row-group"
                | "table-rows"
                | "table-cell"
                | "covered-table-cell"
                | "cell-range-source"
                | "detective"
        ),
        Some(TEXT_NAMESPACE) => is_legal_text_content_local(local),
        _ => false,
    }
}

fn is_legal_text_content_local(local: &str) -> bool {
    matches!(
        local,
        "p" | "h"
            | "span"
            | "s"
            | "tab"
            | "line-break"
            | "a"
            | "soft-page-break"
            | "bookmark"
            | "bookmark-start"
            | "bookmark-end"
            | "reference-mark"
            | "reference-mark-start"
            | "reference-mark-end"
            | "ruby"
            | "ruby-base"
            | "ruby-text"
            | "section"
            | "list"
            | "list-item"
            | "numbered-paragraph"
            | "change"
            | "change-start"
            | "change-end"
            | "page-number"
            | "page-count"
            | "page-variable"
            | "sheet-name"
            | "date"
            | "time"
            | "title"
            | "creator"
            | "author-name"
    )
}

fn reserve_namespace_context_copy(
    context: &NamespaceContext,
    ledger: &mut Ledger<'_>,
) -> Result<()> {
    ledger.reserve_memory(u64_len(namespace_context_bytes(context)?)?)?;
    ledger
        .context
        .consume(Resource::Objects, u64_len(context.bindings.len())?)
        .map_err(map_execution)
}

fn legal_cell_ancestry(spans: &[Span], mut current: Option<usize>) -> Result<bool> {
    let mut state = CellAncestryState::Cell;
    while let Some(index) = current {
        let span = spans
            .get(index)
            .ok_or_else(|| invalid("ODS detective cell ancestry span is invalid"))?;
        state = match (state, span.namespace.as_deref(), span.local.as_str()) {
            (CellAncestryState::Cell, Some(TABLE_NAMESPACE), "table-cell")
            | (CellAncestryState::Cell, Some(TABLE_NAMESPACE), "covered-table-cell") => {
                CellAncestryState::Row
            },
            (CellAncestryState::Row, Some(TABLE_NAMESPACE), "table-row") => {
                CellAncestryState::RowContainer
            },
            (CellAncestryState::RowContainer, Some(TABLE_NAMESPACE), "table")
            | (CellAncestryState::RowContainer, Some(TABLE_NAMESPACE), "table-row-group")
            | (CellAncestryState::RowContainer, Some(TABLE_NAMESPACE), "table-rows")
            | (CellAncestryState::RowContainer, Some(TABLE_NAMESPACE), "table-header-rows") => {
                if span.local == "table" {
                    CellAncestryState::Spreadsheet
                } else {
                    CellAncestryState::RowContainer
                }
            },
            (CellAncestryState::Spreadsheet, Some(OFFICE_NAMESPACE), "spreadsheet") => {
                CellAncestryState::Body
            },
            (CellAncestryState::Body, Some(OFFICE_NAMESPACE), "body") => {
                CellAncestryState::Document
            },
            (CellAncestryState::Document, Some(OFFICE_NAMESPACE), "document-content") => {
                CellAncestryState::Root
            },
            _ => return Ok(false),
        };
        current = span.parent;
    }
    Ok(matches!(state, CellAncestryState::Root))
}

#[derive(Clone, Copy)]
enum CellAncestryState {
    Cell,
    Row,
    RowContainer,
    Spreadsheet,
    Body,
    Document,
    Root,
}

fn ensure_detective_name(namespace: Option<&str>, element: &BytesStart<'_>) -> Result<()> {
    if is_detective_name(namespace, element) {
        Ok(())
    } else {
        Err(invalid("expected an ODS table:detective owner"))
    }
}

fn is_detective_name(namespace: Option<&str>, element: &BytesStart<'_>) -> bool {
    is_table_element(namespace, element.local_name().as_ref(), b"detective")
}

fn is_detective_end(namespace: Option<&str>, element: &BytesEnd<'_>) -> bool {
    is_table_element(
        namespace,
        element.name().local_name().as_ref(),
        b"detective",
    )
}

fn is_table_element(namespace: Option<&str>, local: &[u8], expected: &[u8]) -> bool {
    namespace == Some(TABLE_NAMESPACE) && local == expected
}

fn resolve_namespace(
    resolved: &ResolveResult<'_>,
    ledger: &mut Ledger<'_>,
) -> Result<Option<String>> {
    match resolved {
        ResolveResult::Bound(Namespace(uri)) => {
            Ok(Some(decode_with_ledger(uri, "element namespace", ledger)?))
        },
        ResolveResult::Unbound => Ok(None),
        ResolveResult::Unknown(prefix) => Err(invalid(format!(
            "ODS detective element uses unbound prefix '{}'",
            String::from_utf8_lossy(prefix.as_ref())
        ))),
    }
}

fn namespace_context(
    resolver: &NamespaceResolver,
    limits: Limits,
    ledger: &mut Ledger<'_>,
) -> Result<NamespaceContext> {
    let mut bindings = Vec::new();
    for (prefix, namespace) in resolver.bindings() {
        if bindings.len() >= limits.max_namespace_bindings {
            return Err(limit(
                "namespace bindings",
                bindings.len().saturating_add(1),
                limits.max_namespace_bindings,
            ));
        }
        let prefix = match prefix {
            PrefixDeclaration::Default => String::new(),
            PrefixDeclaration::Named(prefix) => {
                decode_with_ledger(prefix, "namespace prefix", ledger)?
            },
        };
        let uri = decode_with_ledger(namespace.0, "namespace URI", ledger)?;
        let bytes = prefix
            .len()
            .checked_add(uri.len())
            .ok_or_else(|| invalid("ODS namespace context size overflows"))?;
        ledger.reserve_memory(u64_len(
            bytes
                .checked_add(std::mem::size_of::<NamespaceBinding>())
                .ok_or_else(|| invalid("ODS namespace context allocation overflows"))?,
        )?)?;
        reserve_vec_slot(&mut bindings, ledger, "ODS detective namespace context")?;
        ledger
            .context
            .consume(Resource::Objects, 1)
            .map_err(map_execution)?;
        bindings.push(NamespaceBinding { prefix, uri });
    }
    Ok(NamespaceContext { bindings })
}

fn seed_resolver(resolver: &mut NamespaceResolver, context: &NamespaceContext) -> Result<()> {
    for binding in &context.bindings {
        let prefix = if binding.prefix.is_empty() {
            PrefixDeclaration::Default
        } else {
            PrefixDeclaration::Named(binding.prefix.as_bytes())
        };
        resolver
            .add(prefix, Namespace(binding.uri.as_bytes()))
            .map_err(|error| {
                invalid(format!("invalid inherited ODS namespace binding: {error}"))
            })?;
    }
    Ok(())
}

struct RenderedOwner {
    text: String,
    memory: Reservation,
}

fn render_for_prefix(
    value: &Detective,
    prefix: &str,
    context: &NamespaceContext,
    limits: Limits,
    context_budget: &ExecutionContext,
) -> Result<RenderedOwner> {
    if prefix.is_empty() || !context.has_table_binding(prefix) {
        return Err(unsupported(
            "ODS detective writer cannot retain the source table namespace prefix",
        ));
    }
    ensure_model_limits(value, limits)?;
    let (canonical_len, rendered_len) = rendered_lengths(value, prefix)?;
    if rendered_len > limits.max_output_bytes {
        return Err(limit(
            "detective owner output bytes",
            rendered_len,
            limits.max_output_bytes,
        ));
    }
    let scratch = if prefix == "table" {
        canonical_len
    } else {
        canonical_len
            .checked_add(rendered_len)
            .ok_or_else(|| invalid("ODS detective owner rendering scratch overflows"))?
    };
    if scratch > limits.max_scratch_bytes {
        return Err(limit(
            "detective owner rendering scratch",
            scratch,
            limits.max_scratch_bytes,
        ));
    }
    context_budget
        .consume(Resource::Work, u64_len(rendered_len)?)
        .map_err(map_execution)?;
    let memory = context_budget
        .reserve(Resource::Memory, u64_len(scratch)?)
        .map_err(map_execution)?;
    let rendered =
        render_for_prefix_unchecked_with_capacities(value, prefix, canonical_len, rendered_len)?;
    if rendered.len() != rendered_len {
        return Err(limit(
            "detective owner output bytes",
            rendered.len(),
            rendered_len,
        ));
    }
    Ok(RenderedOwner {
        text: rendered,
        memory,
    })
}

fn render_for_prefix_unchecked_with_capacities(
    value: &Detective,
    prefix: &str,
    canonical_len: usize,
    rendered_len: usize,
) -> Result<String> {
    let mut canonical = String::new();
    canonical
        .try_reserve_exact(canonical_len)
        .map_err(|source| allocation("ODS detective owner canonical rendering", source))?;
    write_detective(&mut canonical, value);
    if canonical.len() != canonical_len {
        return Err(invalid("ODS detective canonical rendering length changed"));
    }
    if prefix == "table" {
        return Ok(canonical);
    }
    let mut output = String::new();
    output
        .try_reserve_exact(rendered_len)
        .map_err(|source| allocation("ODS detective owner rendering", source))?;
    rewrite_table_qnames(&canonical, prefix, &mut output)?;
    if output.len() != rendered_len {
        return Err(invalid("ODS detective rendering length changed"));
    }
    Ok(output)
}

fn ensure_model_hard_counts(value: &Detective) -> Result<()> {
    if value.highlighted_ranges().len() > MAX_ITEMS {
        return Err(limit(
            "highlighted ranges",
            value.highlighted_ranges().len(),
            MAX_ITEMS,
        ));
    }
    if value.operations().len() > MAX_ITEMS {
        return Err(limit("operations", value.operations().len(), MAX_ITEMS));
    }
    Ok(())
}

fn ensure_model_limits(value: &Detective, limits: Limits) -> Result<()> {
    ensure_model_hard_counts(value)?;
    if value.highlighted_ranges().len() > limits.max_highlighted_ranges {
        return Err(limit(
            "highlighted ranges",
            value.highlighted_ranges().len(),
            limits.max_highlighted_ranges,
        ));
    }
    if value.operations().len() > limits.max_operations {
        return Err(limit(
            "operations",
            value.operations().len(),
            limits.max_operations,
        ));
    }
    Ok(())
}

fn rendered_lengths(value: &Detective, prefix: &str) -> Result<(usize, usize)> {
    ensure_model_hard_counts(value)?;
    if prefix.is_empty() {
        return Err(unsupported(
            "ODS detective owner has no source table prefix",
        ));
    }
    let canonical = exact_detective_render_len(value, "table")?;
    let rendered = exact_detective_render_len(value, prefix)?;
    Ok((canonical, rendered))
}

fn exact_detective_render_len(value: &Detective, prefix: &str) -> Result<usize> {
    let root = qname_len(prefix, "detective")?;
    let mut total = root
        .checked_add(2)
        .ok_or_else(|| invalid("ODS detective writer root tag size overflows"))?
        .checked_add(tag_end_len(root)?)
        .ok_or_else(|| invalid("ODS detective writer root size overflows"))?;
    for range in value.highlighted_ranges() {
        let element = qname_len(prefix, "highlighted-range")?;
        let mut child = tag_start_len(element)?;
        if let Some(address) = range.cell_range_address() {
            child = child
                .checked_add(attribute_len(
                    qname_len(prefix, "cell-range-address")?,
                    escaped_xml_len(address)?,
                )?)
                .ok_or_else(|| invalid("ODS detective writer range size overflows"))?;
        }
        if let Some(direction) = range.direction() {
            child = child
                .checked_add(attribute_len(
                    qname_len(prefix, "direction")?,
                    direction_text(direction).len(),
                )?)
                .ok_or_else(|| invalid("ODS detective writer range size overflows"))?;
            if let Some(contains_error) = range.contains_error() {
                child = child
                    .checked_add(attribute_len(
                        qname_len(prefix, "contains-error")?,
                        bool_text_len(contains_error),
                    )?)
                    .ok_or_else(|| invalid("ODS detective writer range size overflows"))?;
            }
        } else if let Some(marked_invalid) = range.marked_invalid() {
            child = child
                .checked_add(attribute_len(
                    qname_len(prefix, "marked-invalid")?,
                    bool_text_len(marked_invalid),
                )?)
                .ok_or_else(|| invalid("ODS detective writer range size overflows"))?;
        } else {
            return Err(invalid("ODS detective range has no representable state"));
        }
        child = child
            .checked_add(2)
            .ok_or_else(|| invalid("ODS detective writer range size overflows"))?;
        total = total
            .checked_add(child)
            .ok_or_else(|| invalid("ODS detective writer size overflows"))?;
    }
    for operation in value.operations() {
        let element = qname_len(prefix, "operation")?;
        let name_attribute = attribute_len(
            qname_len(prefix, "name")?,
            operation_text(operation.kind).len(),
        )?;
        let index_attribute =
            attribute_len(qname_len(prefix, "index")?, decimal_digits(operation.index))?;
        let child = tag_start_len(element)?
            .checked_add(name_attribute)
            .and_then(|length| length.checked_add(index_attribute))
            .and_then(|length| length.checked_add(2))
            .ok_or_else(|| invalid("ODS detective writer operation size overflows"))?;
        total = total
            .checked_add(child)
            .ok_or_else(|| invalid("ODS detective writer size overflows"))?;
    }
    Ok(total)
}

fn qname_len(prefix: &str, local: &str) -> Result<usize> {
    prefix
        .len()
        .checked_add(1)
        .and_then(|length| length.checked_add(local.len()))
        .ok_or_else(|| invalid("ODS detective writer QName size overflows"))
}

fn tag_start_len(qname: usize) -> Result<usize> {
    qname
        .checked_add(1)
        .ok_or_else(|| invalid("ODS detective writer tag size overflows"))
}

fn tag_end_len(qname: usize) -> Result<usize> {
    qname
        .checked_add(3)
        .ok_or_else(|| invalid("ODS detective writer closing tag size overflows"))
}

fn attribute_len(qname: usize, value: usize) -> Result<usize> {
    qname
        .checked_add(value)
        .and_then(|length| length.checked_add(4))
        .ok_or_else(|| invalid("ODS detective writer attribute size overflows"))
}

fn bool_text_len(value: bool) -> usize {
    if value { 4 } else { 5 }
}

fn direction_text(direction: Direction) -> &'static str {
    match direction {
        Direction::FromAnotherTable => "from-another-table",
        Direction::ToAnotherTable => "to-another-table",
        Direction::FromSameTable => "from-same-table",
    }
}

fn operation_text(kind: OperationKind) -> &'static str {
    match kind {
        OperationKind::TraceDependents => "trace-dependents",
        OperationKind::RemoveDependents => "remove-dependents",
        OperationKind::TracePrecedents => "trace-precedents",
        OperationKind::RemovePrecedents => "remove-precedents",
        OperationKind::TraceErrors => "trace-errors",
    }
}

fn escaped_xml_len(value: &str) -> Result<usize> {
    value.as_bytes().iter().try_fold(0usize, |total, byte| {
        let addition = match byte {
            b'&' => 5,
            b'<' | b'>' => 4,
            b'"' | b'\'' => 6,
            _ => 1,
        };
        total
            .checked_add(addition)
            .ok_or_else(|| invalid("ODS detective escaped attribute size overflows"))
    })
}

fn decimal_digits(mut value: usize) -> usize {
    let mut digits = 1;
    while value >= 10 {
        value /= 10;
        digits += 1;
    }
    digits
}

fn rewrite_table_qnames(canonical: &str, prefix: &str, output: &mut String) -> Result<()> {
    let bytes = canonical.as_bytes();
    let mut index = 0usize;
    while index < bytes.len() {
        if bytes[index] != b'<' {
            let next = bytes[index..]
                .iter()
                .position(|byte| *byte == b'<')
                .map_or(bytes.len(), |offset| index + offset);
            output.push_str(
                canonical
                    .get(index..next)
                    .ok_or_else(|| invalid("ODS detective writer token boundary is invalid"))?,
            );
            index = next;
            continue;
        }
        output.push('<');
        index += 1;
        if index >= bytes.len() {
            return Err(invalid("ODS detective writer has an incomplete tag"));
        }
        if bytes[index] == b'/' {
            output.push('/');
            index += 1;
        }
        if index >= bytes.len() {
            return Err(invalid("ODS detective writer has an incomplete tag name"));
        }
        if matches!(bytes[index], b'!' | b'?') {
            let end = bytes[index..]
                .iter()
                .position(|byte| *byte == b'>')
                .map(|offset| index + offset + 1)
                .ok_or_else(|| invalid("ODS detective writer tag has no end"))?;
            output.push_str(
                canonical
                    .get(index..end)
                    .ok_or_else(|| invalid("ODS detective writer token boundary is invalid"))?,
            );
            index = end;
            continue;
        }
        let name_end = qname_token_end(bytes, index);
        if name_end == index {
            return Err(invalid("ODS detective writer has an empty tag name"));
        }
        append_rewritten_qname(
            output,
            canonical
                .get(index..name_end)
                .ok_or_else(|| invalid("ODS detective writer tag name boundary is invalid"))?,
            prefix,
        );
        index = name_end;
        loop {
            if index >= bytes.len() {
                return Err(invalid("ODS detective writer tag has no end"));
            }
            if bytes[index] == b'>' {
                output.push('>');
                index += 1;
                break;
            }
            if bytes[index].is_ascii_whitespace() {
                let end = index
                    + bytes[index..]
                        .iter()
                        .take_while(|byte| byte.is_ascii_whitespace())
                        .count();
                output.push_str(canonical.get(index..end).ok_or_else(|| {
                    invalid("ODS detective writer whitespace boundary is invalid")
                })?);
                index = end;
                continue;
            }
            if bytes[index] == b'/' {
                output.push('/');
                index += 1;
                continue;
            }
            if matches!(bytes[index], b'"' | b'\'') {
                let quote = bytes[index];
                let end = bytes[index + 1..]
                    .iter()
                    .position(|byte| *byte == quote)
                    .map(|offset| index + offset + 2)
                    .ok_or_else(|| {
                        invalid("ODS detective writer attribute has no closing quote")
                    })?;
                output.push_str(
                    canonical
                        .get(index..end)
                        .ok_or_else(|| invalid("ODS detective writer value boundary is invalid"))?,
                );
                index = end;
                continue;
            }
            if bytes[index] == b'=' {
                output.push('=');
                index += 1;
                continue;
            }
            let attr_end = qname_token_end(bytes, index);
            let after = bytes[attr_end..]
                .iter()
                .position(|byte| !byte.is_ascii_whitespace())
                .map_or(bytes.len(), |offset| attr_end + offset);
            if after < bytes.len() && bytes[after] == b'=' {
                append_rewritten_qname(
                    output,
                    canonical.get(index..attr_end).ok_or_else(|| {
                        invalid("ODS detective writer attribute boundary is invalid")
                    })?,
                    prefix,
                );
            } else {
                output.push_str(
                    canonical
                        .get(index..attr_end)
                        .ok_or_else(|| invalid("ODS detective writer token boundary is invalid"))?,
                );
            }
            index = attr_end;
        }
    }
    Ok(())
}

fn qname_token_end(bytes: &[u8], start: usize) -> usize {
    bytes[start..]
        .iter()
        .position(|byte| byte.is_ascii_whitespace() || matches!(*byte, b'=' | b'/' | b'>'))
        .map_or(bytes.len(), |offset| start + offset)
}

fn append_rewritten_qname(output: &mut String, token: &str, prefix: &str) {
    if let Some(local) = token.strip_prefix("table:") {
        output.push_str(prefix);
        output.push(':');
        output.push_str(local);
    } else {
        output.push_str(token);
    }
}

fn ensure_whitespace_text(
    text: &quick_xml::events::BytesText<'_>,
    version: XmlVersion,
    owner: &str,
) -> Result<()> {
    let value = text
        .xml_content(version)
        .map_err(|error| invalid(format!("invalid ODS detective {owner} text: {error}")))?;
    ensure_whitespace_value(value.as_ref(), owner)
}

fn ensure_text_event_limit(bytes: usize, limits: Limits) -> Result<()> {
    if bytes > limits.max_text_bytes {
        return Err(limit("text bytes", bytes, limits.max_text_bytes));
    }
    Ok(())
}

fn ensure_whitespace_cdata(text: &BytesCData<'_>, version: XmlVersion, owner: &str) -> Result<()> {
    let value = text
        .xml_content(version)
        .map_err(|error| invalid(format!("invalid ODS detective {owner} CDATA: {error}")))?;
    ensure_whitespace_value(value.as_ref(), owner)
}

fn ensure_whitespace_value(value: &str, owner: &str) -> Result<()> {
    if value.trim().is_empty() {
        Ok(())
    } else {
        Err(invalid(format!("ODS detective {owner} must be empty")))
    }
}

fn is_whitespace_event(source: &str, start: usize, end: usize) -> bool {
    source
        .get(start..end)
        .is_some_and(|value| value.trim().is_empty())
}

fn event_has_non_whitespace_text(event: &Event<'_>) -> Result<bool> {
    let value = match event {
        Event::Text(text) => text
            .xml_content(XmlVersion::Explicit1_0)
            .map_err(|error| invalid(format!("invalid ODS cell text: {error}")))?,
        Event::CData(text) => text
            .xml_content(XmlVersion::Explicit1_0)
            .map_err(|error| invalid(format!("invalid ODS cell CDATA: {error}")))?,
        _ => return Ok(false),
    };
    Ok(!value.trim().is_empty())
}

fn mark_direct_cell_text(spans: &mut [Span], current: Option<usize>) -> Result<()> {
    let Some(index) = current else {
        return Ok(());
    };
    let span = spans
        .get_mut(index)
        .ok_or_else(|| invalid("ODS detective text span is invalid"))?;
    if is_cell_element(span.namespace.as_deref(), span.local.as_str()) {
        span.raw_text_non_whitespace = true;
    }
    Ok(())
}

fn set_once<T>(slot: &mut Option<T>, value: T, name: &str) -> Result<()> {
    if slot.replace(value).is_some() {
        return Err(invalid(format!(
            "duplicate ODS detective attribute '{name}'"
        )));
    }
    Ok(())
}

fn parse_bool(value: &str, name: &str) -> Result<bool> {
    match value {
        "true" | "1" => Ok(true),
        "false" | "0" => Ok(false),
        _ => Err(invalid(format!("invalid {name} boolean '{value}'"))),
    }
}

fn reserve_ordered(ordered: &mut Vec<OrderedChild<'_>>, ledger: &mut Ledger<'_>) -> Result<()> {
    if ordered.len()
        >= ledger
            .limits
            .max_highlighted_ranges
            .saturating_add(ledger.limits.max_operations)
    {
        return Err(limit(
            "detective children",
            ordered.len().saturating_add(1),
            ledger
                .limits
                .max_highlighted_ranges
                .saturating_add(ledger.limits.max_operations),
        ));
    }
    reserve_vec_slot(ordered, ledger, "ODS detective ordered children")?;
    ledger
        .context
        .consume(Resource::Objects, 1)
        .map_err(map_execution)?;
    Ok(())
}

fn reserve_model_slot<T>(
    value: &Detective,
    highlighted: bool,
    reserved_capacity: &mut usize,
    ledger: &mut Ledger<'_>,
) -> Result<()> {
    let current = if highlighted {
        value.highlighted_ranges().len()
    } else {
        value.operations().len()
    };
    let next = current
        .checked_add(1)
        .ok_or_else(|| invalid("ODS detective model vector capacity overflows"))?;
    if next <= *reserved_capacity {
        return Ok(());
    }
    let slots = next
        .checked_next_power_of_two()
        .and_then(|slots| slots.max(4).checked_sub(*reserved_capacity))
        .ok_or_else(|| invalid("ODS detective model vector capacity overflows"))?;
    let bytes = slots
        .checked_mul(std::mem::size_of::<T>())
        .ok_or_else(|| invalid("ODS detective model vector memory overflows"))?;
    ledger.reserve_memory(u64_len(bytes)?)?;
    *reserved_capacity = reserved_capacity
        .checked_add(slots)
        .ok_or_else(|| invalid("ODS detective model vector capacity overflows"))?;
    Ok(())
}

fn reserve_vec_slot<T>(
    values: &mut Vec<T>,
    ledger: &mut Ledger<'_>,
    resource: &'static str,
) -> Result<()> {
    let additional = if values.capacity() == values.len() {
        values
            .len()
            .checked_add(1)
            .ok_or_else(|| invalid("ODS detective vector capacity overflows"))?
    } else {
        1
    };
    let bytes = additional
        .checked_mul(std::mem::size_of::<T>())
        .ok_or_else(|| invalid("ODS detective vector memory overflows"))?;
    ledger.reserve_memory(u64_len(bytes)?)?;
    values
        .try_reserve_exact(additional)
        .map_err(|source| allocation(resource, source))?;
    Ok(())
}

fn ensure_highlighted_range_capacity(current: usize, limits: Limits) -> Result<()> {
    if current >= limits.max_highlighted_ranges {
        return Err(limit(
            "highlighted ranges",
            current.saturating_add(1),
            limits.max_highlighted_ranges,
        ));
    }
    Ok(())
}

fn ensure_operation_capacity(current: usize, limits: Limits) -> Result<()> {
    if current >= limits.max_operations {
        return Err(limit(
            "operations",
            current.saturating_add(1),
            limits.max_operations,
        ));
    }
    Ok(())
}

fn position(reader: &NsReader<&[u8]>) -> Result<usize> {
    usize::try_from(reader.buffer_position())
        .map_err(|_| invalid("ODS detective XML position overflows usize"))
}

fn prefix_string(
    prefix: Option<quick_xml::name::Prefix<'_>>,
    ledger: &mut Ledger<'_>,
) -> Result<String> {
    prefix.map_or_else(
        || Ok(String::new()),
        |prefix| decode_with_ledger(prefix.as_ref(), "element prefix", ledger),
    )
}

fn decode(bytes: &[u8], name: &str) -> Result<String> {
    String::from_utf8(bytes.to_vec())
        .map_err(|_| invalid(format!("ODS detective {name} is not valid UTF-8")))
}

fn decode_with_ledger(bytes: &[u8], name: &str, ledger: &mut Ledger<'_>) -> Result<String> {
    ledger.reserve_memory(u64_len(bytes.len())?)?;
    decode(bytes, name)
}

fn xml_error(error: quick_xml::Error) -> Error {
    invalid(format!("invalid ODS detective XML: {error}"))
}

fn u64_len(value: usize) -> Result<u64> {
    u64::try_from(value).map_err(|_| invalid("ODS detective size overflows u64"))
}

fn allocation(resource: &'static str, source: std::collections::TryReserveError) -> Error {
    Error::Allocation { resource, source }
}

fn invalid(message: impl Into<String>) -> Error {
    Error::InvalidFormat(message.into())
}

fn unsupported(message: impl Into<String>) -> Error {
    Error::Unsupported(message.into())
}

fn limit(resource: &'static str, observed: usize, maximum: usize) -> Error {
    Error::Unsupported(format!(
        "ODS detective {resource} limit exceeded: {observed} > {maximum}"
    ))
}

fn map_execution(error: ExecutionError) -> Error {
    match error {
        ExecutionError::ResourceLimit(limit) => Error::ResourceLimit(limit),
        ExecutionError::Cancelled => unsupported("ODS detective operation cancelled"),
        other => unsupported(format!(
            "ODS detective execution policy rejected operation: {other}"
        )),
    }
}

fn merge_reservation(slot: &mut Option<Reservation>, other: Reservation) -> Result<()> {
    if let Some(current) = slot.as_mut() {
        if let Err(other) = current.try_merge(other) {
            drop(other);
            return Err(unsupported("ODS detective reservation lineage mismatch"));
        }
    } else {
        *slot = Some(other);
    }
    Ok(())
}

struct Ledger<'context> {
    context: &'context ExecutionContext,
    limits: Limits,
    work: u64,
    events: usize,
    owner_events: usize,
    memory: Option<Reservation>,
}

impl<'context> Ledger<'context> {
    fn new(context: &'context ExecutionContext, limits: Limits) -> Self {
        Self {
            context,
            limits,
            work: 0,
            events: 0,
            owner_events: 0,
            memory: None,
        }
    }

    fn charge_work(&mut self, amount: u64) -> Result<()> {
        let next = self
            .work
            .checked_add(amount)
            .ok_or_else(|| invalid("ODS detective work counter overflows"))?;
        if next > self.limits.max_work_units {
            return Err(limit(
                "work units",
                usize::try_from(next.min(usize::MAX as u64)).unwrap_or(usize::MAX),
                usize::try_from(self.limits.max_work_units.min(usize::MAX as u64))
                    .unwrap_or(usize::MAX),
            ));
        }
        self.context
            .consume(Resource::Work, amount)
            .map_err(map_execution)?;
        self.work = next;
        Ok(())
    }

    fn event(&mut self) -> Result<()> {
        self.events = self
            .events
            .checked_add(1)
            .ok_or_else(|| invalid("ODS detective event counter overflows"))?;
        if self.events > self.limits.max_events {
            return Err(limit("XML events", self.events, self.limits.max_events));
        }
        self.charge_work(16)
    }

    fn begin_owner(&mut self) {
        self.owner_events = 0;
    }

    fn owner_event(&mut self) -> Result<()> {
        self.owner_events = self
            .owner_events
            .checked_add(1)
            .ok_or_else(|| invalid("ODS detective owner event counter overflows"))?;
        if self.owner_events > self.limits.max_events {
            return Err(limit(
                "detective owner events",
                self.owner_events,
                self.limits.max_events,
            ));
        }
        self.charge_work(16)
    }

    fn check(&self) -> Result<()> {
        self.context.check().map_err(map_execution)
    }

    fn reserve_memory(&mut self, amount: u64) -> Result<()> {
        if amount == 0 {
            return Ok(());
        }
        let reservation = self
            .context
            .reserve(Resource::Memory, amount)
            .map_err(map_execution)?;
        merge_reservation(&mut self.memory, reservation)
    }

    fn reserve_event_memory(&self, event_len: usize) -> Result<Option<Reservation>> {
        if event_len == 0 {
            return Ok(None);
        }
        let bytes = event_len
            .checked_add(64)
            .ok_or_else(|| invalid("ODS detective owned event size overflows"))?;
        self.context
            .reserve(Resource::Memory, u64_len(bytes)?)
            .map(Some)
            .map_err(map_execution)
    }
}

// `Scanned` is created in a helper to keep the source lifetime explicit while
// retaining the scan implementation's local allocations.
impl<'source> From<Vec<Span>> for Scanned<'source> {
    fn from(spans: Vec<Span>) -> Self {
        Self::from_spans(spans)
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use litchi_core::{Budget, CancellationSource, ExecutionLimits, Limits as CoreLimits, Profile};
    use std::num::{NonZeroU64, NonZeroUsize};

    fn context() -> ExecutionContext {
        let budget = Budget::root(
            "ods-detective-tests",
            CoreLimits::for_profile(Profile::TrustedBatch),
        );
        let (_source, token) = CancellationSource::pair();
        let execution = ExecutionLimits::new(
            NonZeroUsize::new(1).expect("non-zero test worker"),
            NonZeroUsize::new(1).expect("non-zero test task cap"),
            NonZeroU64::new(1024 * 1024).expect("non-zero test byte cap"),
            0,
        )
        .expect("test execution limits");
        ExecutionContext::new(budget, token, execution)
    }

    const PREFIX: &str = concat!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\"?>",
        "<o:document-content xmlns:o=\"urn:oasis:names:tc:opendocument:xmlns:office:1.0\" ",
        "xmlns:t=\"urn:oasis:names:tc:opendocument:xmlns:table:1.0\" ",
        "xmlns:text=\"urn:oasis:names:tc:opendocument:xmlns:text:1.0\">",
        "<o:body><o:spreadsheet><t:table t:name=\"S\"><t:table-row>",
        "<t:table-cell>"
    );
    const SUFFIX: &str =
        "</t:table-cell></t:table-row></t:table></o:spreadsheet></o:body></o:document-content>";

    #[test]
    fn parses_aliases_and_retains_ordered_typed_children() {
        let xml = format!(
            "{PREFIX}<t:detective><t:highlighted-range t:direction=\"from-same-table\" t:contains-error=\"1\"/><t:operation t:name=\"trace-errors\" t:index=\"7\"/></t:detective>{SUFFIX}"
        );
        let snapshot = Snapshot::parse_with_context(&xml, Limits::server(), &context())
            .expect("valid aliased detective should parse");
        assert_eq!(snapshot.owners().len(), 1);
        let owner = &snapshot.owners()[0];
        let typed = owner.typed().expect("owner should be typed");
        assert_eq!(typed.value().highlighted_ranges().len(), 1);
        assert_eq!(typed.value().operations()[0].index, 7);
        assert!(matches!(
            typed.ordered_children()[0],
            OrderedChild::HighlightedRange { index: 0, .. }
        ));
        assert!(matches!(
            typed.ordered_children()[1],
            OrderedChild::Operation { index: 0, .. }
        ));
        assert_eq!(typed.source_prefix(), "t");
        assert_eq!(typed.namespace_context().table_prefix(), Some("t"));
    }

    #[test]
    fn distinguishes_empty_and_absent_and_removal_preserves_siblings() {
        let xml = format!(
            "{PREFIX}<!--before--><t:detective></t:detective><!--keep--><text:p>after</text:p>{SUFFIX}"
        );
        let snapshot = Snapshot::parse_with_context(&xml, Limits::server(), &context())
            .expect("empty detective should parse");
        let owner = snapshot.owners()[0].id();
        let plan = snapshot
            .plan_remove(owner, &context())
            .expect("typed empty owner should be removable");
        let candidate = plan.apply(&context()).expect("removal should splice");
        assert!(candidate.retained_output_bytes() > 0);
        assert!(candidate.retained_memory_bytes() > 0);
        assert!(!candidate.contains("<t:detective"));
        assert!(candidate.contains("<!--before-->"));
        assert!(candidate.contains("<!--keep-->"));
        assert!(candidate.contains("before"));
        assert!(candidate.contains("after"));
    }

    #[test]
    fn rejects_malformed_ranges_duplicates_and_foreign_wrappers_for_editing() {
        let malformed = concat!(
            "<t:detective xmlns:t=\"urn:oasis:names:tc:opendocument:xmlns:table:1.0\">",
            "<t:highlighted-range t:marked-invalid=\"true\" t:direction=\"from-same-table\"/>",
            "</t:detective>"
        );
        assert!(parse_detective(malformed, Limits::server(), &context()).is_err());

        let xml = format!(
            "{PREFIX}<t:detective/><t:detective/><x:wrap xmlns:x=\"urn:vendor\"><t:detective/></x:wrap>{SUFFIX}"
        );
        let snapshot = Snapshot::parse_with_context(&xml, Limits::server(), &context())
            .expect("opaque duplicate owners remain inspectable");
        assert_eq!(snapshot.owners().len(), 2);
        assert!(
            snapshot
                .owners()
                .iter()
                .all(|owner| owner.opaque_reason().is_some())
        );
        assert!(
            snapshot
                .plan_remove(snapshot.owners()[0].id(), &context())
                .is_err()
        );
    }

    #[test]
    fn refuses_mce_branch_transparency_and_noncanonical_replacement() {
        let xml = format!(
            "{PREFIX}<t:detective>\n<t:operation t:name=\"trace-errors\" t:index=\"0\"/>\n</t:detective>{SUFFIX}"
        );
        let snapshot = Snapshot::parse_with_context(&xml, Limits::server(), &context())
            .expect("whitespace owner should remain readable");
        let owner = snapshot.owners()[0].id();
        let value = snapshot.owners()[0]
            .typed()
            .expect("typed source")
            .value()
            .clone();
        assert!(
            snapshot
                .plan_replace(owner, &value, &context())
                .expect("equal edit")
                .is_noop()
        );
        let noop = snapshot
            .plan_replace(owner, &value, &context())
            .expect("equal edit")
            .apply(&context())
            .expect("no-op application");
        assert!(noop.is_borrowed());
        assert_eq!(noop.retained_output_bytes(), 0);
        assert_eq!(noop.retained_memory_bytes(), 0);

        let mce = format!(
            "{PREFIX}<x:wrap xmlns:x=\"{MC_NAMESPACE}\"><x:Choice><t:detective/></x:Choice></x:wrap>{SUFFIX}"
        );
        let snapshot = Snapshot::parse_with_context(&mce, Limits::server(), &context())
            .expect("MCE branch remains opaque");
        assert!(snapshot.owners().is_empty());
        assert!(snapshot.cells()[0].has_unsupported_child());
        assert!(
            snapshot
                .plan_insert(CellId(0), &Detective::new(), &context())
                .is_err()
        );
    }

    #[test]
    fn indexes_covered_cells_and_resolves_text_aliases_for_insertion() {
        let text_binding = format!("xmlns:text=\"{TEXT_NAMESPACE}\">");
        let text_binding_with_alias =
            format!("xmlns:text=\"{TEXT_NAMESPACE}\" xmlns:x=\"{TEXT_NAMESPACE}\">");
        let prefix = PREFIX.replacen(&text_binding, &text_binding_with_alias, 1);
        let xml = format!("{prefix}<x:p>after</x:p>{SUFFIX}");
        let snapshot = Snapshot::parse_with_context(&xml, Limits::server(), &context())
            .expect("covered cell with an aliased text child should parse");
        assert_eq!(snapshot.cells().len(), 1);
        assert_eq!(snapshot.cells()[0].kind(), CellKind::TableCell);
        assert_eq!(
            snapshot.cells()[0].qualification().parent(),
            CellKind::TableCell
        );
        let plan = snapshot
            .plan_insert(CellId(0), &Detective::new(), &context())
            .expect("the aliased text child supplies a safe insertion slot");
        let candidate = plan.apply(&context()).expect("insertion should splice");
        assert!(candidate.find("<t:detective></t:detective>") < candidate.find("<x:p>"));

        let prefix_without_cell = PREFIX
            .strip_suffix("<t:table-cell>")
            .expect("test prefix has a physical-cell slot");
        let suffix_without_cell = SUFFIX
            .strip_prefix("</t:table-cell>")
            .expect("test suffix has a physical-cell close");
        let covered = format!(
            "{prefix_without_cell}<t:covered-table-cell><t:detective/></t:covered-table-cell>{suffix_without_cell}"
        );
        let covered = Snapshot::parse_with_context(&covered, Limits::server(), &context())
            .expect("covered cell owner should parse");
        assert_eq!(covered.cells()[0].kind(), CellKind::CoveredTableCell);
        assert_eq!(covered.owners()[0].cell(), covered.cells()[0].id());
    }

    #[test]
    fn expands_a_self_closing_cell_instead_of_inserting_a_sibling() {
        let prefix = PREFIX
            .strip_suffix("<t:table-cell>")
            .expect("test prefix has a physical-cell slot");
        let suffix = SUFFIX
            .strip_prefix("</t:table-cell>")
            .expect("test suffix has a physical-cell close");
        let xml = format!("{prefix}<t:table-cell/>{suffix}");
        let snapshot = Snapshot::parse_with_context(&xml, Limits::server(), &context())
            .expect("self-closing physical cell should parse");
        let plan = snapshot
            .plan_insert(CellId(0), &Detective::new(), &context())
            .expect("self-closing physical cell should expand");
        assert_eq!(plan.kind(), PlanKind::Insert);
        let candidate = plan.apply(&context()).expect("expanded cell should splice");
        assert!(candidate.contains("<t:table-cell><t:detective></t:detective></t:table-cell>"));
        assert!(!candidate.contains("<t:table-cell/><t:detective"));
    }

    #[test]
    fn removal_refuses_noncanonical_owner_lexical_data() {
        let xml = format!("{PREFIX}<t:detective/>{SUFFIX}");
        let snapshot = Snapshot::parse_with_context(&xml, Limits::server(), &context())
            .expect("self-closing owner should remain readable");
        let owner = snapshot.owners()[0].id();
        assert!(snapshot.plan_remove(owner, &context()).is_err());
    }

    #[test]
    fn rejects_per_owner_child_limits_on_read_and_render() {
        let limits = Limits::server()
            .with_item_limits(4, 1, 1)
            .expect("bounded owner item limits");
        let xml = format!(
            "{PREFIX}<t:detective><t:operation t:name=\"trace-errors\" t:index=\"0\"/><t:operation t:name=\"trace-errors\" t:index=\"1\"/></t:detective>{SUFFIX}"
        );
        let snapshot = Snapshot::parse_with_context(&xml, limits, &context())
            .expect("bounded scan keeps an opaque owner");
        assert!(snapshot.owners()[0].opaque_reason().is_some());
        let mut value = Detective::new();
        value
            .add_operation(Operation::new(OperationKind::TraceErrors, 0))
            .add_operation(Operation::new(OperationKind::TraceErrors, 1));
        let empty = format!("{PREFIX}{SUFFIX}");
        let empty = Snapshot::parse_with_context(&empty, limits, &context())
            .expect("empty cell should parse");
        assert!(empty.plan_insert(CellId(0), &value, &context()).is_err());
    }

    #[test]
    fn rewrites_only_qname_tokens_and_preserves_quoted_values() {
        let canonical = concat!(
            "<table:detective table:quoted=\"table:Bar\">",
            "<table:operation table:name=\"trace-errors\" table:index=\"0\"/>",
            "</table:detective>"
        );
        let mut output = String::new();
        rewrite_table_qnames(canonical, "t", &mut output).expect("QName rewrite");
        assert!(output.contains("<t:detective t:quoted=\"table:Bar\">"));
        assert!(output.contains("<t:operation t:name=\"trace-errors\" t:index=\"0\"/>"));
    }

    #[test]
    fn standalone_typed_and_rendered_candidates_retain_their_guards() {
        let namespace_context = NamespaceContext {
            bindings: vec![NamespaceBinding {
                prefix: "table".to_string(),
                uri: TABLE_NAMESPACE.to_string(),
            }],
        };
        let typed = parse_detective_with_namespace(
            "<table:detective/>",
            0,
            &namespace_context,
            Limits::server(),
            &context(),
        )
        .expect("standalone detective should parse");
        assert!(typed.retained_memory_bytes() > 0);
        assert_eq!(
            typed.retained_input_bytes(),
            "<table:detective/>".len() as u64
        );
        let cloned = typed.clone();
        drop(typed);
        assert!(cloned.retained_memory_bytes() > 0);
        assert_eq!(
            cloned.retained_input_bytes(),
            "<table:detective/>".len() as u64
        );

        let default_namespace = parse_detective_with_namespace(
            "<detective xmlns=\"urn:oasis:names:tc:opendocument:xmlns:table:1.0\"/>",
            0,
            &NamespaceContext::default(),
            Limits::server(),
            &context(),
        )
        .expect("default table namespace should remain readable");
        assert_eq!(default_namespace.source_prefix(), "");
        assert_eq!(
            default_namespace.lexical_fidelity(),
            LexicalFidelity::PreservationRequired
        );

        let mut value = Detective::new();
        value.add_operation(Operation::new(OperationKind::TraceErrors, 0));
        let (canonical_len, rendered_len) = rendered_lengths(&value, "alias").expect("lengths");
        let rendered = render_for_prefix_unchecked_with_capacities(
            &value,
            "alias",
            canonical_len,
            rendered_len,
        )
        .expect("exact render");
        assert_eq!(rendered.len(), rendered_len);
        assert_eq!(
            canonical_len,
            exact_detective_render_len(&value, "table").unwrap()
        );
    }

    #[test]
    fn standalone_input_reservation_accepts_exact_limit_and_rejects_one_under() {
        let source = "<table:detective/>";
        let namespace_context = NamespaceContext {
            bindings: vec![NamespaceBinding {
                prefix: "table".to_string(),
                uri: TABLE_NAMESPACE.to_string(),
            }],
        };
        let exact = Limits::new(source.len(), 4096, 16, 128, 100_000).expect("exact input profile");
        let typed =
            parse_detective_with_namespace(source, 0, &namespace_context, exact, &context())
                .expect("exact input limit should pass");
        assert_eq!(typed.retained_input_bytes(), source.len() as u64);
        let retained = typed.clone();
        drop(typed);
        assert_eq!(retained.retained_input_bytes(), source.len() as u64);
        drop(retained);

        let budget = Budget::root(
            "ods-detective-input-lifetime",
            CoreLimits::new(
                4 * 1024 * 1024,
                source.len() as u64,
                4 * 1024 * 1024,
                1_000_000,
                1_024,
                1_000_000,
            ),
        );
        let (_cancel_source, token) = CancellationSource::pair();
        let execution = ExecutionLimits::new(
            NonZeroUsize::new(1).expect("non-zero test worker"),
            NonZeroUsize::new(1).expect("non-zero test task cap"),
            NonZeroU64::new(1024 * 1024).expect("non-zero test byte cap"),
            0,
        )
        .expect("test execution limits");
        let bounded_context = ExecutionContext::new(budget, token, execution);
        let typed =
            parse_detective_with_namespace(source, 0, &namespace_context, exact, &bounded_context)
                .expect("bounded input reservation should parse");
        assert!(bounded_context.reserve(Resource::InputBytes, 1).is_err());
        let retained = typed.clone();
        drop(typed);
        assert!(bounded_context.reserve(Resource::InputBytes, 1).is_err());
        drop(retained);
        assert!(bounded_context.reserve(Resource::InputBytes, 1).is_ok());

        let under =
            Limits::new(source.len() - 1, 4096, 16, 128, 100_000).expect("one-under input profile");
        assert!(
            parse_detective_with_namespace(source, 0, &namespace_context, under, &context(),)
                .is_err()
        );
    }

    #[test]
    fn exact_render_lengths_cover_true_and_false_boolean_lexicals() {
        let mut true_value = Detective::new();
        true_value.add_highlighted_range(
            HighlightedRange::valid(None, Direction::FromSameTable, Some(true))
                .expect("valid true boolean range"),
        );
        let mut false_value = Detective::new();
        false_value.add_highlighted_range(
            HighlightedRange::valid(None, Direction::FromSameTable, Some(false))
                .expect("valid false boolean range"),
        );
        let (true_len, _) = rendered_lengths(&true_value, "table").expect("true length");
        let (false_len, _) = rendered_lengths(&false_value, "table").expect("false length");
        assert_eq!(false_len, true_len + 1);

        for value in [&true_value, &false_value] {
            let (canonical_len, rendered_len) = rendered_lengths(value, "alias").expect("lengths");
            let mut canonical = String::new();
            write_detective(&mut canonical, value);
            assert_eq!(canonical.len(), canonical_len);
            let rendered = render_for_prefix_unchecked_with_capacities(
                value,
                "alias",
                canonical_len,
                rendered_len,
            )
            .expect("exact render");
            assert_eq!(rendered.len(), rendered_len);
        }
    }

    #[test]
    fn rejects_nested_cells_and_rows_in_source_qualification() {
        let prefix = PREFIX
            .strip_suffix("<t:table-cell>")
            .expect("test prefix has a physical-cell slot");
        let suffix = SUFFIX
            .strip_prefix("</t:table-cell>")
            .expect("test suffix has a physical-cell close");
        for content in [
            "<t:table-cell><t:table-cell><t:detective/></t:table-cell></t:table-cell>",
            "<t:table-row><t:table-row><t:table-cell><t:detective/></t:table-cell></t:table-row></t:table-row>",
        ] {
            let xml = format!("{prefix}{content}{suffix}");
            let snapshot = Snapshot::parse_with_context(&xml, Limits::server(), &context())
                .expect("nested source remains inspectable");
            assert!(
                snapshot
                    .cells()
                    .iter()
                    .any(|cell| cell.qualification().has_foreign_ancestor())
            );
            assert!(snapshot.cells().iter().all(|cell| {
                snapshot
                    .plan_insert(cell.id(), &Detective::new(), &context())
                    .is_err()
            }));
        }
    }

    #[test]
    fn owner_local_namespace_alias_is_readable_but_not_editable() {
        let xml = format!("{PREFIX}<x:detective xmlns:x=\"{TABLE_NAMESPACE}\"/>{SUFFIX}");
        let snapshot = Snapshot::parse_with_context(&xml, Limits::server(), &context())
            .expect("owner-local namespace alias should remain indexed");
        let owner = &snapshot.owners()[0];
        let typed = owner.typed().expect("owner should remain readable");
        assert_eq!(typed.source_prefix(), "x");
        assert_eq!(
            owner.qualification().namespace_context().table_prefix(),
            Some("t")
        );
        assert!(snapshot.plan_remove(owner.id(), &context()).is_err());
    }

    #[test]
    fn refuses_raw_cell_text_unknown_text_qnames_duplicate_sources_and_bad_rows() {
        for content in [
            "raw",
            "<![CDATA[raw]]>",
            "<text:unknown xmlns:text=\"urn:oasis:names:tc:opendocument:xmlns:text:1.0\"/>",
            "<t:cell-range-source/><t:cell-range-source/>",
        ] {
            let xml = format!("{PREFIX}{content}{SUFFIX}");
            let snapshot = Snapshot::parse_with_context(&xml, Limits::server(), &context())
                .expect("bounded malformed cell content remains inspectable");
            assert!(
                !snapshot.cells()[0].sequence_valid()
                    || snapshot.cells()[0].has_unsupported_child()
            );
            assert!(
                snapshot
                    .plan_insert(CellId(0), &Detective::new(), &context())
                    .is_err()
            );
        }

        let xml = concat!(
            "<o:document-content xmlns:o=\"urn:oasis:names:tc:opendocument:xmlns:office:1.0\" ",
            "xmlns:t=\"urn:oasis:names:tc:opendocument:xmlns:table:1.0\">",
            "<o:body><o:spreadsheet><t:table><t:table-cell/></t:table></o:spreadsheet></o:body>",
            "</o:document-content>"
        );
        let snapshot = Snapshot::parse_with_context(xml, Limits::server(), &context())
            .expect("bad row ancestry should remain inspectable");
        assert!(snapshot.cells()[0].qualification().has_foreign_ancestor());
        assert!(
            snapshot
                .plan_insert(CellId(0), &Detective::new(), &context())
                .is_err()
        );
    }

    #[test]
    fn cancellation_is_checked_before_scan() {
        let budget = Budget::root(
            "ods-detective-cancel",
            CoreLimits::for_profile(Profile::Server),
        );
        let (source, token) = CancellationSource::pair();
        source.cancel();
        let execution = ExecutionLimits::new(
            NonZeroUsize::new(1).expect("worker"),
            NonZeroUsize::new(1).expect("tasks"),
            NonZeroU64::new(1).expect("bytes"),
            0,
        )
        .expect("execution limits");
        let context = ExecutionContext::new(budget, token, execution);
        assert!(
            Snapshot::parse_with_context("<t:detective/>", Limits::server(), &context).is_err()
        );
    }
}
