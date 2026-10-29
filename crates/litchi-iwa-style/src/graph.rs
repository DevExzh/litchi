//! A copy-on-write workspace over the parsed IWA components of one package,
//! for the object-graph edits of style variations: allocating, inserting and
//! removing style objects, creating and deleting component members, and the
//! `TSP.PackageMetadata` registry transitions (object UUIDs, component
//! external references, component registrations and the issued-identifier
//! watermark) that go with them.
//!
//! The owner supplies its parsed components through [`Components`]; the
//! workspace never re-reads package bytes. [`Workspace::finish`] returns the
//! [`Changes`]: the components to re-encode in place, the members to insert
//! and the members to delete. The owner re-encodes and reassembles its
//! package once (every component and member the edit did not touch keeps its
//! exact bytes).
//!
//! Registry transitions follow the migration host byte for byte (record
//! 0375a): a scalar is replaced in place, a repeated record is appended after
//! the component's last field, and a removal drops exactly the selected
//! record. New object UUIDs are derived deterministically (record 0298
//! §2.1) where the host draws random ones.

use std::cell::RefCell;
use std::collections::{BTreeMap, BTreeSet, HashMap, HashSet};
use std::fmt;
use std::mem::size_of;

use litchi_iwa_common::wire::{self, WireField};
use litchi_iwa_common::{WireLimits, decode_varint_from_bytes};
use litchi_iwa_core::{Archive, ArchiveObject, Limits, RawMessage};
use litchi_iwa_protos::package_metadata_codec::{
    self as registry_codec, Batch, ComponentSelector, RewriteError, RewriteOptions,
};

/// The package metadata component.
pub const METADATA_MEMBER: &str = "Index/Metadata.iwa";
/// The document metadata component (new standalone components go before it).
pub const DOCUMENT_METADATA_MEMBER: &str = "Index/DocumentMetadata.iwa";
const PACKAGE_METADATA_MESSAGE_TYPE: u32 = 11_006;
const REGISTRY_RECURSION_LIMIT: u32 = 64;

const METADATA_LAST_OBJECT_IDENTIFIER: u32 = 1;
const METADATA_COMPONENTS: u32 = 3;
const COMPONENT_IDENTIFIER: u32 = 1;
const COMPONENT_EXTERNAL_REFERENCES: u32 = 6;
const COMPONENT_OBJECT_UUIDS: u32 = 11;
const EXTERNAL_COMPONENT: u32 = 1;
const EXTERNAL_OBJECT: u32 = 2;
const EXTERNAL_IS_WEAK: u32 = 3;
const UUID_ENTRY_IDENTIFIER: u32 = 1;
const UUID_ENTRY_UUID: u32 = 2;

/// The parsed IWA components of one package, in package order.
pub trait Components {
    /// How many components there are.
    fn count(&self) -> usize;
    /// The member name and parsed archive of component `index`.
    fn component(&self, index: usize) -> Option<(&str, &Archive)>;
}

/// Resource dimensions owned by graph-level bounded allocation adapters.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum GraphLimitKind {
    /// Temporary scratch bytes reserved by a graph rewrite.
    ScratchBytes,
    /// Fallible allocation events admitted by a graph rewrite.
    AllocationEvents,
}

impl Components for [(String, Archive)] {
    fn count(&self) -> usize {
        self.len()
    }

    fn component(&self, index: usize) -> Option<(&str, &Archive)> {
        self.get(index)
            .map(|(name, archive)| (name.as_str(), archive))
    }
}

impl Components for Vec<(String, Archive)> {
    fn count(&self) -> usize {
        self.len()
    }

    fn component(&self, index: usize) -> Option<(&str, &Archive)> {
        self.as_slice().component(index)
    }
}

/// Why a graph edit failed. No change escapes a failed edit.
#[derive(Debug)]
#[non_exhaustive]
pub enum Error {
    /// An IWA object or archive operation failed.
    Archive(litchi_iwa_core::Error),
    /// A payload is not valid protobuf wire data.
    Wire(litchi_iwa_common::Error),
    /// A chart style codec failed.
    Codec(litchi_iwa_protos::chart_style_codec::Error),
    /// A graph adapter exhausted a resource that the wire layer does not
    /// model (for example scratch capacity or allocation events).
    Limit {
        /// Resource dimension.
        kind: GraphLimitKind,
        /// Attempted aggregate consumption.
        observed: usize,
        /// Configured maximum.
        maximum: usize,
    },
    /// A package metadata registry rewrite failed.
    Registry(RewriteError),
    /// Object `0` is missing (or the edit names no object).
    MissingObject(u64),
    /// Object `0` is stored by more than one component.
    AmbiguousObject(u64),
    /// The object graph or registry breaks an invariant the edit relies on.
    Invalid(&'static str),
    /// The package-wide object identifier space is exhausted.
    IdentifierOverflow,
}

impl fmt::Display for Error {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Archive(error) => write!(formatter, "IWA archive: {error}"),
            Self::Wire(error) => write!(formatter, "IWA wire: {error}"),
            Self::Codec(error) => write!(formatter, "chart style codec: {error}"),
            Self::Limit {
                kind,
                observed,
                maximum,
            } => write!(formatter, "graph {kind:?} limit: {observed} > {maximum}"),
            Self::Registry(error) => write!(formatter, "package metadata: {error}"),
            Self::MissingObject(identifier) => write!(formatter, "object {identifier} is missing"),
            Self::AmbiguousObject(identifier) => {
                write!(formatter, "object {identifier} is stored more than once")
            },
            Self::Invalid(reason) => formatter.write_str(reason),
            Self::IdentifierOverflow => formatter.write_str("object identifier overflow"),
        }
    }
}

impl std::error::Error for Error {}

impl From<litchi_iwa_core::Error> for Error {
    fn from(error: litchi_iwa_core::Error) -> Self {
        Self::Archive(error)
    }
}

impl From<litchi_iwa_common::Error> for Error {
    fn from(error: litchi_iwa_common::Error) -> Self {
        Self::Wire(error)
    }
}

impl From<litchi_iwa_protos::chart_style_codec::Error> for Error {
    fn from(error: litchi_iwa_protos::chart_style_codec::Error) -> Self {
        Self::Codec(error)
    }
}

impl From<RewriteError> for Error {
    fn from(error: RewriteError) -> Self {
        Self::Registry(error)
    }
}

/// Result of a graph edit.
pub type Result<T> = std::result::Result<T, Error>;

/// A member the edit creates.
#[derive(Debug, Clone)]
#[non_exhaustive]
pub struct Insertion {
    /// The new member name.
    pub name: String,
    /// Its archive.
    pub archive: Archive,
    /// The source member to insert it before (`None`: append).
    pub before: Option<String>,
}

/// What a finished edit changes.
#[derive(Debug, Default)]
#[non_exhaustive]
pub struct Changes {
    /// Existing components to re-encode, in package order.
    pub edits: Vec<(String, Archive)>,
    /// New members, in creation order.
    pub insertions: Vec<Insertion>,
    /// Source members to delete, in package order.
    pub deletions: Vec<String>,
}

#[derive(Debug, Default, Clone, Copy)]
struct FinishLedger {
    /// Every changed-map entry the finalizer must inspect, including entries
    /// later hidden by a member deletion.
    changed_entries_scanned: usize,
    /// Live changed source members that become `Changes::edits` entries.
    edit_count: usize,
    /// Deleted source members that become owned deletion names.
    deletion_count: usize,
    /// Bytes copied into those owned deletion/edit names.
    copy_bytes: usize,
}

impl Changes {
    /// Whether the edit changes nothing.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.edits.is_empty() && self.insertions.is_empty() && self.deletions.is_empty()
    }
}

/// A resource dimension owned by [`Workspace::finish_bounded`].
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum FinishLimitKind {
    /// Bounded source-name and collection traversal work.
    Work,
    /// Fallible vector and owned-name allocation events.
    Allocations,
    /// Bytes retained by the returned change vectors and copied names.
    RetainedBytes,
    /// Temporary bytes retained while constructing the returned changes.
    ScratchBytes,
}

impl fmt::Display for FinishLimitKind {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(match self {
            Self::Work => "finish work",
            Self::Allocations => "finish allocations",
            Self::RetainedBytes => "finish retained bytes",
            Self::ScratchBytes => "finish scratch bytes",
        })
    }
}

/// A typed failure from bounded workspace finalization.
#[derive(Debug)]
#[non_exhaustive]
pub enum FinishError {
    /// The staged workspace itself is invalid.
    Graph(Error),
    /// A finalization requirement overflowed its finite arithmetic.
    Overflow(&'static str),
    /// The caller's residual finalization profile was exhausted.
    Limit {
        /// Exhausted resource.
        kind: FinishLimitKind,
        /// Requested aggregate amount.
        observed: usize,
        /// Caller-supplied maximum.
        maximum: usize,
    },
    /// A fallible vector or owned-name reservation failed.
    Allocation {
        /// Requested elements or bytes.
        amount: usize,
    },
}

impl fmt::Display for FinishError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Graph(error) => write!(formatter, "workspace finalization: {error}"),
            Self::Overflow(reason) => {
                write!(formatter, "workspace finalization overflow: {reason}")
            },
            Self::Limit {
                kind,
                observed,
                maximum,
            } => write!(formatter, "{kind} limit: {observed} > {maximum}"),
            Self::Allocation { amount } => {
                write!(formatter, "workspace finalization allocation: {amount}")
            },
        }
    }
}

impl std::error::Error for FinishError {}

impl From<Error> for FinishError {
    fn from(error: Error) -> Self {
        Self::Graph(error)
    }
}

/// A resource dimension owned by bounded workspace construction.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum ConstructionLimitKind {
    /// Nonallocating component/object census and index population work.
    Work,
    /// HashMap/HashSet index reservations.
    Allocations,
    /// Logical retained bytes for the graph indexes.
    RetainedBytes,
    /// Temporary construction scratch bytes.
    ScratchBytes,
}

impl fmt::Display for ConstructionLimitKind {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(match self {
            Self::Work => "workspace construction work",
            Self::Allocations => "workspace construction allocations",
            Self::RetainedBytes => "workspace index retained bytes",
            Self::ScratchBytes => "workspace construction scratch bytes",
        })
    }
}

/// A typed failure from bounded workspace construction.
#[derive(Debug)]
#[non_exhaustive]
pub enum ConstructionError {
    /// The source census or index invariants are invalid.
    Graph(Error),
    /// A construction residual was exhausted.
    Limit {
        /// Exhausted resource.
        kind: ConstructionLimitKind,
        /// Requested aggregate amount.
        observed: usize,
        /// Caller-supplied maximum.
        maximum: usize,
    },
    /// An index reservation failed before source objects were indexed.
    Allocation {
        /// Requested entries.
        amount: usize,
    },
}

impl fmt::Display for ConstructionError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Graph(error) => write!(formatter, "workspace construction: {error}"),
            Self::Limit {
                kind,
                observed,
                maximum,
            } => write!(formatter, "{kind} limit: {observed} > {maximum}"),
            Self::Allocation { amount } => {
                write!(formatter, "workspace index allocation: {amount}")
            },
        }
    }
}

impl std::error::Error for ConstructionError {}

/// Residual limits for one bounded workspace census or index build.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct ConstructionLimits {
    max_work: usize,
    max_allocations: usize,
    max_retained_bytes: usize,
    max_scratch_bytes: usize,
}

impl ConstructionLimits {
    /// Construct a residual construction profile.
    #[must_use]
    pub const fn new(
        max_work: usize,
        max_allocations: usize,
        max_retained_bytes: usize,
        max_scratch_bytes: usize,
    ) -> Self {
        Self {
            max_work,
            max_allocations,
            max_retained_bytes,
            max_scratch_bytes,
        }
    }

    /// An unbounded profile for legacy tests.
    #[must_use]
    pub const fn unlimited() -> Self {
        Self::new(usize::MAX, usize::MAX, usize::MAX, usize::MAX)
    }

    #[must_use]
    pub const fn max_work(self) -> usize {
        self.max_work
    }

    #[must_use]
    pub const fn max_allocations(self) -> usize {
        self.max_allocations
    }

    #[must_use]
    pub const fn max_retained_bytes(self) -> usize {
        self.max_retained_bytes
    }

    #[must_use]
    pub const fn max_scratch_bytes(self) -> usize {
        self.max_scratch_bytes
    }
}

/// Exact structural requirements for one source census and graph-index build.
///
/// `component_view_bytes` and `component_view_allocations` describe the
/// caller-owned borrowed `Vec<(&str, &Archive)>` adapter when a facade uses
/// one. They are reported separately from the graph-owned indexes and are
/// never charged by [`Workspace::new_bounded`].
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct ConstructionRequirements {
    source_token: usize,
    fingerprint_a: u64,
    fingerprint_b: u64,
    component_count: usize,
    object_count: usize,
    source_work: usize,
    builder_work: usize,
    total_work: usize,
    index_allocations: usize,
    index_retained_bytes: usize,
    index_scratch_bytes: usize,
    component_view_bytes: usize,
    component_view_allocations: usize,
}

impl ConstructionRequirements {
    #[must_use]
    pub const fn component_count(self) -> usize {
        self.component_count
    }

    #[must_use]
    pub const fn object_count(self) -> usize {
        self.object_count
    }

    /// Work for one nonallocating source census or one index population pass.
    #[must_use]
    pub const fn source_work(self) -> usize {
        self.source_work
    }

    /// Aggregate work for `new_bounded`: its internal identity re-census plus
    /// the index population pass.
    #[must_use]
    pub const fn builder_work(self) -> usize {
        self.builder_work
    }

    /// Aggregate work for one external census plus `new_bounded`'s two
    /// internal passes.
    #[must_use]
    pub const fn total_work(self) -> usize {
        self.total_work
    }

    #[must_use]
    pub const fn index_allocations(self) -> usize {
        self.index_allocations
    }

    #[must_use]
    pub const fn index_retained_bytes(self) -> usize {
        self.index_retained_bytes
    }

    #[must_use]
    pub const fn index_scratch_bytes(self) -> usize {
        self.index_scratch_bytes
    }

    #[must_use]
    pub const fn component_view_bytes(self) -> usize {
        self.component_view_bytes
    }

    #[must_use]
    pub const fn component_view_allocations(self) -> usize {
        self.component_view_allocations
    }

    fn check(self, limits: ConstructionLimits) -> std::result::Result<(), ConstructionError> {
        for (kind, observed, maximum) in [
            (
                ConstructionLimitKind::Work,
                self.source_work,
                limits.max_work(),
            ),
            (
                ConstructionLimitKind::Allocations,
                self.index_allocations,
                limits.max_allocations(),
            ),
            (
                ConstructionLimitKind::RetainedBytes,
                self.index_retained_bytes,
                limits.max_retained_bytes(),
            ),
            (
                ConstructionLimitKind::ScratchBytes,
                self.index_scratch_bytes,
                limits.max_scratch_bytes(),
            ),
        ] {
            if observed > maximum {
                return Err(ConstructionError::Limit {
                    kind,
                    observed,
                    maximum,
                });
            }
        }
        Ok(())
    }

    fn check_builder(
        self,
        limits: ConstructionLimits,
    ) -> std::result::Result<(), ConstructionError> {
        if self.builder_work > limits.max_work() {
            return Err(ConstructionError::Limit {
                kind: ConstructionLimitKind::Work,
                observed: self.builder_work,
                maximum: limits.max_work(),
            });
        }
        self.check(ConstructionLimits::new(
            limits.max_work(),
            limits.max_allocations(),
            limits.max_retained_bytes(),
            limits.max_scratch_bytes(),
        ))
    }
}

/// Residual limits for the final transfer from a staged workspace into
/// [`Changes`].  Archive payloads are already owned by the workspace and are
/// transferred without copying; this profile covers only the final vectors,
/// copied member names, and their traversal.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct FinishLimits {
    max_work: usize,
    max_allocations: usize,
    max_retained_bytes: usize,
    max_scratch_bytes: usize,
}

impl FinishLimits {
    /// Construct a residual finalization profile.
    #[must_use]
    pub const fn new(
        max_work: usize,
        max_allocations: usize,
        max_retained_bytes: usize,
        max_scratch_bytes: usize,
    ) -> Self {
        Self {
            max_work,
            max_allocations,
            max_retained_bytes,
            max_scratch_bytes,
        }
    }

    /// An unbounded profile for legacy callers and tests.
    #[must_use]
    pub const fn unlimited() -> Self {
        Self::new(usize::MAX, usize::MAX, usize::MAX, usize::MAX)
    }

    #[must_use]
    pub const fn max_work(self) -> usize {
        self.max_work
    }

    #[must_use]
    pub const fn max_allocations(self) -> usize {
        self.max_allocations
    }

    #[must_use]
    pub const fn max_retained_bytes(self) -> usize {
        self.max_retained_bytes
    }

    #[must_use]
    pub const fn max_scratch_bytes(self) -> usize {
        self.max_scratch_bytes
    }
}

/// Exact source-derived requirements for one finalization pass.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct FinishRequirements {
    changed_entries_scanned: usize,
    edit_count: usize,
    insertion_count: usize,
    deletion_count: usize,
    copy_bytes: usize,
    work: usize,
    allocations: usize,
    retained_bytes: usize,
    scratch_bytes: usize,
    output_bytes: usize,
}

impl FinishRequirements {
    #[must_use]
    pub const fn changed_entries_scanned(self) -> usize {
        self.changed_entries_scanned
    }

    #[must_use]
    pub const fn edit_count(self) -> usize {
        self.edit_count
    }

    #[must_use]
    pub const fn insertion_count(self) -> usize {
        self.insertion_count
    }

    #[must_use]
    pub const fn deletion_count(self) -> usize {
        self.deletion_count
    }

    /// Bytes copied from source member names into owned `Changes` names.
    #[must_use]
    pub const fn copy_bytes(self) -> usize {
        self.copy_bytes
    }

    #[must_use]
    pub const fn work(self) -> usize {
        self.work
    }

    #[must_use]
    pub const fn allocations(self) -> usize {
        self.allocations
    }

    #[must_use]
    pub const fn retained_bytes(self) -> usize {
        self.retained_bytes
    }

    #[must_use]
    pub const fn scratch_bytes(self) -> usize {
        self.scratch_bytes
    }

    #[must_use]
    pub const fn output_bytes(self) -> usize {
        self.output_bytes
    }

    fn check(self, limits: FinishLimits) -> std::result::Result<(), FinishError> {
        for (kind, observed, maximum) in [
            (FinishLimitKind::Work, self.work, limits.max_work()),
            (
                FinishLimitKind::Allocations,
                self.allocations,
                limits.max_allocations(),
            ),
            (
                FinishLimitKind::RetainedBytes,
                self.retained_bytes,
                limits.max_retained_bytes(),
            ),
            (
                FinishLimitKind::ScratchBytes,
                self.scratch_bytes,
                limits.max_scratch_bytes(),
            ),
        ] {
            if observed > maximum {
                return Err(FinishError::Limit {
                    kind,
                    observed,
                    maximum,
                });
            }
        }
        Ok(())
    }
}

/// Actual resources consumed by one successful bounded finalization.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct FinishReport {
    requirements: FinishRequirements,
}

impl FinishReport {
    #[must_use]
    pub const fn requirements(self) -> FinishRequirements {
        self.requirements
    }

    #[must_use]
    pub const fn changed_entries_scanned(self) -> usize {
        self.requirements.changed_entries_scanned()
    }

    #[must_use]
    pub const fn edit_count(self) -> usize {
        self.requirements.edit_count()
    }

    #[must_use]
    pub const fn insertion_count(self) -> usize {
        self.requirements.insertion_count()
    }

    #[must_use]
    pub const fn deletion_count(self) -> usize {
        self.requirements.deletion_count()
    }

    #[must_use]
    pub const fn copy_bytes(self) -> usize {
        self.requirements.copy_bytes()
    }

    #[must_use]
    pub const fn work(self) -> usize {
        self.requirements.work()
    }

    #[must_use]
    pub const fn allocations(self) -> usize {
        self.requirements.allocations()
    }

    #[must_use]
    pub const fn retained_bytes(self) -> usize {
        self.requirements.retained_bytes()
    }

    #[must_use]
    pub const fn scratch_bytes(self) -> usize {
        self.requirements.scratch_bytes()
    }

    #[must_use]
    pub const fn output_bytes(self) -> usize {
        self.requirements.output_bytes()
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, PartialOrd, Ord)]
pub(crate) enum Location {
    Source(usize),
    Created(usize),
}

// These are logical accounting envelopes for the standard-library hash
// indexes, deliberately independent of package/source byte size.  They are
// shared with bounded callers rather than repeated as facade multipliers.
const INDEX_SLOT_BYTES: usize = 64;
const AMBIGUOUS_SLOT_BYTES: usize = 32;

fn source_token(source: &dyn Components) -> usize {
    let raw: *const dyn Components = source;
    raw as *const () as usize
}

fn fingerprint_u64(first: &mut u64, second: &mut u64, value: u64) {
    *first ^= value;
    *first = first.wrapping_mul(0x1000_0000_01b3);
    *second = second.rotate_left(13) ^ value.wrapping_mul(0x9e37_79b9_7f4a_7c15);
}

fn fingerprint_bytes(first: &mut u64, second: &mut u64, bytes: &[u8]) {
    for byte in bytes {
        fingerprint_u64(first, second, u64::from(*byte));
    }
}

/// The staged state of one graph edit.
pub struct Workspace<'s> {
    source: &'s dyn Components,
    limits: Limits,
    changed: BTreeMap<usize, Archive>,
    created: Vec<Insertion>,
    deleted: BTreeSet<usize>,
    index: HashMap<u64, Location>,
    ambiguous: HashSet<u64>,
    /// `(component, object, message)` of the registry payload.
    registry: Option<(usize, u64, usize)>,
    registry_requirements: RefCell<Option<Option<registry_codec::RewriteReport>>>,
    registry_root: RefCell<Option<Option<Root>>>,
    finish: FinishLedger,
    uuid_seed: u64,
    source_work: usize,
}

impl fmt::Debug for Workspace<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Workspace")
            .field("changed", &self.changed.len())
            .field("created", &self.created.len())
            .field("deleted", &self.deleted.len())
            .finish_non_exhaustive()
    }
}

impl<'s> Workspace<'s> {
    /// Census the borrowed component view without allocating a graph index.
    ///
    /// The caller supplies the residual census profile. `source_work` is the
    /// exact amount for this external census; `builder_work` covers the
    /// builder's internal re-census plus population; `total_work` covers all
    /// three actual passes.
    pub fn construction_requirements(
        source: &dyn Components,
        limits: ConstructionLimits,
    ) -> std::result::Result<ConstructionRequirements, ConstructionError> {
        let component_count = source.count();
        let mut object_count = 0usize;
        let mut source_work = 0usize;
        let mut fingerprint_a = 0xcbf2_9ce4_8422_2325;
        let mut fingerprint_b = 0x8422_2325_cbf2_9ce4;
        for position in 0..component_count {
            source_work =
                source_work
                    .checked_add(1)
                    .ok_or(ConstructionError::Graph(Error::Invalid(
                        "workspace census work overflow",
                    )))?;
            if source_work > limits.max_work() {
                return Err(ConstructionError::Limit {
                    kind: ConstructionLimitKind::Work,
                    observed: source_work,
                    maximum: limits.max_work(),
                });
            }
            let (name, archive) =
                source
                    .component(position)
                    .ok_or(ConstructionError::Graph(Error::Invalid(
                        "a component vanished during workspace census",
                    )))?;
            fingerprint_u64(&mut fingerprint_a, &mut fingerprint_b, position as u64);
            fingerprint_u64(&mut fingerprint_a, &mut fingerprint_b, name.len() as u64);
            fingerprint_bytes(&mut fingerprint_a, &mut fingerprint_b, name.as_bytes());
            fingerprint_u64(
                &mut fingerprint_a,
                &mut fingerprint_b,
                archive.objects.len() as u64,
            );
            source_work = source_work
                .checked_add(name.len())
                .ok_or(ConstructionError::Graph(Error::Invalid(
                    "workspace census work overflow",
                )))?;
            if source_work > limits.max_work() {
                return Err(ConstructionError::Limit {
                    kind: ConstructionLimitKind::Work,
                    observed: source_work,
                    maximum: limits.max_work(),
                });
            }
            for object in &archive.objects {
                let identifier = object
                    .archive_info
                    .identifier
                    .ok_or(ConstructionError::Graph(Error::Invalid(
                        "an IWA object has no identifier",
                    )))?;
                fingerprint_u64(&mut fingerprint_a, &mut fingerprint_b, identifier);
                fingerprint_u64(
                    &mut fingerprint_a,
                    &mut fingerprint_b,
                    object.messages.len() as u64,
                );
                source_work =
                    source_work
                        .checked_add(1)
                        .ok_or(ConstructionError::Graph(Error::Invalid(
                            "workspace census work overflow",
                        )))?;
                if source_work > limits.max_work() {
                    return Err(ConstructionError::Limit {
                        kind: ConstructionLimitKind::Work,
                        observed: source_work,
                        maximum: limits.max_work(),
                    });
                }
                source_work = source_work.checked_add(object.messages.len()).ok_or(
                    ConstructionError::Graph(Error::Invalid("workspace census work overflow")),
                )?;
                if source_work > limits.max_work() {
                    return Err(ConstructionError::Limit {
                        kind: ConstructionLimitKind::Work,
                        observed: source_work,
                        maximum: limits.max_work(),
                    });
                }
                for (message, raw) in object.messages.iter().enumerate() {
                    fingerprint_u64(&mut fingerprint_a, &mut fingerprint_b, message as u64);
                    fingerprint_u64(&mut fingerprint_a, &mut fingerprint_b, u64::from(raw.type_));
                    fingerprint_u64(
                        &mut fingerprint_a,
                        &mut fingerprint_b,
                        raw.data.len() as u64,
                    );
                }
                object_count =
                    object_count
                        .checked_add(1)
                        .ok_or(ConstructionError::Graph(Error::Invalid(
                            "workspace object count overflow",
                        )))?;
            }
        }
        let index_allocations = usize::from(object_count != 0).saturating_mul(2);
        let index_retained_bytes = object_count
            .checked_mul(INDEX_SLOT_BYTES.saturating_add(AMBIGUOUS_SLOT_BYTES))
            .ok_or(ConstructionError::Graph(Error::Invalid(
                "workspace index bytes overflow",
            )))?;
        let component_view_bytes = component_count
            .checked_mul(size_of::<(&str, &Archive)>())
            .ok_or(ConstructionError::Graph(Error::Invalid(
                "component view bytes overflow",
            )))?;
        let builder_work =
            source_work
                .checked_mul(2)
                .ok_or(ConstructionError::Graph(Error::Invalid(
                    "workspace builder work overflow",
                )))?;
        let total_work = source_work
            .checked_add(builder_work)
            .ok_or(ConstructionError::Graph(Error::Invalid(
                "workspace construction work overflow",
            )))?;
        let requirements = ConstructionRequirements {
            source_token: source_token(source),
            fingerprint_a,
            fingerprint_b,
            component_count,
            object_count,
            source_work,
            builder_work,
            total_work,
            index_allocations,
            index_retained_bytes,
            index_scratch_bytes: 0,
            component_view_bytes,
            component_view_allocations: usize::from(component_count != 0),
        };
        requirements.check(limits)?;
        Ok(requirements)
    }

    /// Build a workspace after a caller has admitted one exact construction
    /// census. The builder performs an internal identity re-census and then
    /// a population pass; no index grows beyond the returned requirements.
    pub fn new_bounded(
        source: &'s dyn Components,
        limits: Limits,
        requirements: ConstructionRequirements,
        construction_limits: ConstructionLimits,
    ) -> std::result::Result<Self, ConstructionError> {
        requirements.check_builder(construction_limits)?;
        let current = Self::construction_requirements(
            source,
            ConstructionLimits::new(
                construction_limits.max_work(),
                construction_limits.max_allocations(),
                construction_limits.max_retained_bytes(),
                construction_limits.max_scratch_bytes(),
            ),
        )?;
        if current != requirements {
            return Err(ConstructionError::Graph(Error::Invalid(
                "workspace source changed between census and build",
            )));
        }
        if source.count() != requirements.component_count {
            return Err(ConstructionError::Graph(Error::Invalid(
                "workspace component count changed between census and build",
            )));
        }
        let mut index = HashMap::new();
        let mut ambiguous = HashSet::new();
        if requirements.object_count != 0 {
            index.try_reserve(requirements.object_count).map_err(|_| {
                ConstructionError::Allocation {
                    amount: requirements.object_count,
                }
            })?;
            ambiguous
                .try_reserve(requirements.object_count)
                .map_err(|_| ConstructionError::Allocation {
                    amount: requirements.object_count,
                })?;
        }
        let mut registry = None;
        let mut source_objects = 0usize;
        let mut source_work = 0usize;
        let mut fingerprint_a = 0xcbf2_9ce4_8422_2325;
        let mut fingerprint_b = 0x8422_2325_cbf2_9ce4;
        for position in 0..requirements.component_count {
            if source.count() != requirements.component_count {
                return Err(ConstructionError::Graph(Error::Invalid(
                    "workspace component count changed during build",
                )));
            }
            source_work =
                source_work
                    .checked_add(1)
                    .ok_or(ConstructionError::Graph(Error::Invalid(
                        "workspace build work overflow",
                    )))?;
            let (name, archive) =
                source
                    .component(position)
                    .ok_or(ConstructionError::Graph(Error::Invalid(
                        "a component vanished during workspace build",
                    )))?;
            fingerprint_u64(&mut fingerprint_a, &mut fingerprint_b, position as u64);
            fingerprint_u64(&mut fingerprint_a, &mut fingerprint_b, name.len() as u64);
            fingerprint_bytes(&mut fingerprint_a, &mut fingerprint_b, name.as_bytes());
            fingerprint_u64(
                &mut fingerprint_a,
                &mut fingerprint_b,
                archive.objects.len() as u64,
            );
            source_work = source_work
                .checked_add(name.len())
                .ok_or(ConstructionError::Graph(Error::Invalid(
                    "workspace build work overflow",
                )))?;
            source_objects = source_objects.checked_add(archive.objects.len()).ok_or(
                ConstructionError::Graph(Error::Invalid("workspace object count overflow")),
            )?;
            if source_objects > requirements.object_count {
                return Err(ConstructionError::Graph(Error::Invalid(
                    "workspace object count grew during build",
                )));
            }
            for object in &archive.objects {
                source_work =
                    source_work
                        .checked_add(1)
                        .ok_or(ConstructionError::Graph(Error::Invalid(
                            "workspace build work overflow",
                        )))?;
                let identifier = object
                    .archive_info
                    .identifier
                    .ok_or(ConstructionError::Graph(Error::Invalid(
                        "an IWA object has no identifier",
                    )))?;
                fingerprint_u64(&mut fingerprint_a, &mut fingerprint_b, identifier);
                fingerprint_u64(
                    &mut fingerprint_a,
                    &mut fingerprint_b,
                    object.messages.len() as u64,
                );
                source_work = source_work.checked_add(object.messages.len()).ok_or(
                    ConstructionError::Graph(Error::Invalid("workspace build work overflow")),
                )?;
                for (message, raw) in object.messages.iter().enumerate() {
                    fingerprint_u64(&mut fingerprint_a, &mut fingerprint_b, message as u64);
                    fingerprint_u64(&mut fingerprint_a, &mut fingerprint_b, u64::from(raw.type_));
                    fingerprint_u64(
                        &mut fingerprint_a,
                        &mut fingerprint_b,
                        raw.data.len() as u64,
                    );
                }
                if index
                    .insert(identifier, Location::Source(position))
                    .is_some()
                {
                    ambiguous.insert(identifier);
                }
                if name == METADATA_MEMBER {
                    for (message, raw) in object.messages.iter().enumerate() {
                        if raw.type_ == PACKAGE_METADATA_MESSAGE_TYPE
                            && registry.replace((position, identifier, message)).is_some()
                        {
                            return Err(ConstructionError::Graph(Error::Invalid(
                                "the package metadata component holds several registries",
                            )));
                        }
                    }
                }
            }
        }
        if source_objects != requirements.object_count
            || source_work != requirements.source_work
            || source_token(source) != requirements.source_token
            || fingerprint_a != requirements.fingerprint_a
            || fingerprint_b != requirements.fingerprint_b
        {
            return Err(ConstructionError::Graph(Error::Invalid(
                "workspace source changed between census and build",
            )));
        }
        if source_work > construction_limits.max_work() {
            return Err(ConstructionError::Limit {
                kind: ConstructionLimitKind::Work,
                observed: source_work,
                maximum: construction_limits.max_work(),
            });
        }
        Ok(Self {
            source,
            limits,
            changed: BTreeMap::new(),
            created: Vec::new(),
            deleted: BTreeSet::new(),
            index,
            ambiguous,
            registry,
            registry_requirements: RefCell::new(None),
            registry_root: RefCell::new(None),
            finish: FinishLedger::default(),
            uuid_seed: 0,
            source_work: 0,
        })
    }

    /// Stage an edit over `source` under the archive `limits`.
    ///
    /// # Errors
    ///
    /// [`Error::Invalid`] for an object without an identifier or a package
    /// metadata component without exactly one registry payload.
    pub fn new(source: &'s dyn Components, limits: Limits) -> Result<Self> {
        let mut index = HashMap::new();
        let mut ambiguous = HashSet::new();
        let mut registry = None;
        let mut source_objects = 0usize;
        for position in 0..source.count() {
            let (name, archive) = source
                .component(position)
                .ok_or(Error::Invalid("a component vanished"))?;
            source_objects = source_objects.saturating_add(archive.objects.len());
            for object in &archive.objects {
                let identifier = object
                    .archive_info
                    .identifier
                    .ok_or(Error::Invalid("an IWA object has no identifier"))?;
                if index
                    .insert(identifier, Location::Source(position))
                    .is_some()
                {
                    ambiguous.insert(identifier);
                }
                if name == METADATA_MEMBER {
                    for (message, raw) in object.messages.iter().enumerate() {
                        if raw.type_ == PACKAGE_METADATA_MESSAGE_TYPE
                            && registry.replace((position, identifier, message)).is_some()
                        {
                            return Err(Error::Invalid(
                                "the package metadata component holds several registries",
                            ));
                        }
                    }
                }
            }
        }
        // Workspace construction already owns the parsed component bytes;
        // charge the bounded index census by component/object visits rather
        // than re-billing every borrowed payload byte to each style read.
        let source_work = source.count().saturating_add(source_objects);
        Ok(Self {
            source,
            limits,
            changed: BTreeMap::new(),
            created: Vec::new(),
            deleted: BTreeSet::new(),
            index,
            ambiguous,
            registry,
            registry_requirements: RefCell::new(None),
            registry_root: RefCell::new(None),
            finish: FinishLedger::default(),
            uuid_seed: 0,
            source_work,
        })
    }

    /// The archive limits of the edit.
    #[must_use]
    pub const fn limits(&self) -> Limits {
        self.limits
    }

    /// Clone the mutable graph staging state so a caller can publish no
    /// partial changes when a later admission fails.
    pub(crate) fn fork(&self) -> Self {
        Self {
            source: self.source,
            limits: self.limits,
            changed: self.changed.clone(),
            created: self.created.clone(),
            deleted: self.deleted.clone(),
            index: self.index.clone(),
            ambiguous: self.ambiguous.clone(),
            registry: self.registry,
            registry_requirements: RefCell::new(None),
            registry_root: RefCell::new(None),
            finish: self.finish,
            uuid_seed: self.uuid_seed,
            source_work: self.source_work,
        }
    }

    /// Constructor scan work for the legacy [`Self::new`] path.
    ///
    /// A workspace built with [`Self::new_bounded`] has already exposed its
    /// external census and the builder's two passes through
    /// [`ConstructionRequirements::total_work`] and therefore reports zero
    /// here; bounded callers must not charge construction as engine work.
    #[must_use]
    pub(crate) fn source_work(&self) -> usize {
        self.source_work
    }

    /// Exact requirements for transferring the staged workspace into
    /// [`Changes`].  The ledger is updated as source members first become
    /// copy-on-write entries or are deleted, so this query performs no
    /// source-name scan and allocates nothing.
    #[must_use = "inspect the bounded finalization requirements or handle overflow"]
    pub fn finish_requirements(&self) -> std::result::Result<FinishRequirements, FinishError> {
        let vector_allocations = usize::from(self.finish.edit_count != 0)
            .checked_add(usize::from(self.finish.deletion_count != 0))
            .ok_or(FinishError::Overflow("finish vector allocation count"))?;
        let allocations = vector_allocations
            .checked_add(self.finish.edit_count)
            .and_then(|value| value.checked_add(self.finish.deletion_count))
            .ok_or(FinishError::Overflow("finish allocation count"))?;
        let vector_bytes = self
            .finish
            .edit_count
            .checked_mul(size_of::<(String, Archive)>());
        let deletion_vector_bytes = self.finish.deletion_count.checked_mul(size_of::<String>());
        let vector_bytes = vector_bytes.ok_or(FinishError::Overflow("edit vector bytes"))?;
        let deletion_vector_bytes =
            deletion_vector_bytes.ok_or(FinishError::Overflow("deletion vector bytes"))?;
        let retained_bytes = vector_bytes
            .checked_add(deletion_vector_bytes)
            .and_then(|value| value.checked_add(self.finish.copy_bytes))
            .ok_or(FinishError::Overflow("finish retained bytes"))?;
        let work = self
            .finish
            .changed_entries_scanned
            .checked_add(self.finish.deletion_count)
            .and_then(|value| value.checked_add(self.finish.copy_bytes))
            .ok_or(FinishError::Overflow("finish work"))?;
        Ok(FinishRequirements {
            changed_entries_scanned: self.finish.changed_entries_scanned,
            edit_count: self.finish.edit_count,
            insertion_count: self.created.len(),
            deletion_count: self.finish.deletion_count,
            copy_bytes: self.finish.copy_bytes,
            work,
            allocations,
            retained_bytes,
            scratch_bytes: 0,
            output_bytes: 0,
        })
    }

    /// Finish the edit under an explicit residual profile.
    ///
    /// The final vectors and copied source member names are reserved only
    /// after the complete requirement check.  Existing archive payloads and
    /// created insertions are moved from the workspace, not copied.  The
    /// returned report is the exact logical footprint of this transfer and
    /// is suitable for charging the caller's transaction ledger.
    pub fn finish_bounded(
        self,
        limits: FinishLimits,
    ) -> std::result::Result<(Changes, FinishReport), FinishError> {
        let requirements = self.finish_requirements()?;
        requirements.check(limits)?;

        let mut edits = Vec::new();
        edits
            .try_reserve_exact(requirements.edit_count)
            .map_err(|_| FinishError::Allocation {
                amount: requirements.edit_count,
            })?;
        let mut deletions = Vec::new();
        deletions
            .try_reserve_exact(requirements.deletion_count)
            .map_err(|_| FinishError::Allocation {
                amount: requirements.deletion_count,
            })?;

        for (position, archive) in self.changed {
            if self.deleted.contains(&position) {
                continue;
            }
            let (name, _) =
                self.source
                    .component(position)
                    .ok_or(FinishError::Graph(Error::Invalid(
                        "a changed component vanished during finalization",
                    )))?;
            edits.push((owned_name(name)?, archive));
        }
        for position in self.deleted {
            let (name, _) =
                self.source
                    .component(position)
                    .ok_or(FinishError::Graph(Error::Invalid(
                        "a deleted component vanished during finalization",
                    )))?;
            deletions.push(owned_name(name)?);
        }
        let changes = Changes {
            edits,
            insertions: self.created,
            deletions,
        };
        Ok((changes, FinishReport { requirements }))
    }

    /// Finish the edit using the historical infallible API.
    ///
    /// New bounded facades should use [`Self::finish_requirements`] followed
    /// by [`Self::finish_bounded`] so finalization is admitted before names
    /// and change vectors are allocated.
    #[must_use]
    pub fn finish(self) -> Changes {
        self.finish_bounded(FinishLimits::unlimited())
            .expect("unbounded workspace finalization")
            .0
    }

    fn archive(&self, location: Location) -> Option<&Archive> {
        match location {
            Location::Source(position) => self
                .changed
                .get(&position)
                .or_else(|| self.source.component(position).map(|(_, archive)| archive)),
            Location::Created(position) => self.created.get(position).map(|entry| &entry.archive),
        }
    }

    fn archive_mut(&mut self, location: Location) -> Result<&mut Archive> {
        match location {
            Location::Source(position) => {
                if !self.changed.contains_key(&position) {
                    let (name, archive) = self
                        .source
                        .component(position)
                        .ok_or(Error::Invalid("a component vanished"))?;
                    self.changed.insert(position, archive.clone());
                    self.finish.changed_entries_scanned = self
                        .finish
                        .changed_entries_scanned
                        .checked_add(1)
                        .ok_or(Error::Invalid("finish entry count overflow"))?;
                    if !self.deleted.contains(&position) {
                        self.finish.edit_count = self
                            .finish
                            .edit_count
                            .checked_add(1)
                            .ok_or(Error::Invalid("finish edit count overflow"))?;
                        self.finish.copy_bytes = self
                            .finish
                            .copy_bytes
                            .checked_add(name.len())
                            .ok_or(Error::Invalid("finish name bytes overflow"))?;
                    }
                }
                self.changed
                    .get_mut(&position)
                    .ok_or(Error::Invalid("a component vanished"))
            },
            Location::Created(position) => self
                .created
                .get_mut(position)
                .map(|entry| &mut entry.archive)
                .ok_or(Error::Invalid("a created component vanished")),
        }
    }

    /// The member name of a location.
    pub(crate) fn member(&self, location: Location) -> &str {
        match location {
            Location::Source(position) => {
                self.source.component(position).map_or("", |(name, _)| name)
            },
            Location::Created(position) => self
                .created
                .get(position)
                .map_or("", |entry| entry.name.as_str()),
        }
    }

    /// Where object `identifier` is stored now.
    pub(crate) fn locate(&self, identifier: u64) -> Result<Location> {
        if self.ambiguous.contains(&identifier) {
            return Err(Error::AmbiguousObject(identifier));
        }
        self.index
            .get(&identifier)
            .copied()
            .ok_or(Error::MissingObject(identifier))
    }

    /// The location of member `name`, if it exists now.
    pub(crate) fn member_location(&self, name: &str) -> Option<Location> {
        if let Some(position) = self.created.iter().position(|entry| entry.name == name) {
            return Some(Location::Created(position));
        }
        (0..self.source.count())
            .find(|position| {
                !self.deleted.contains(position)
                    && self
                        .source
                        .component(*position)
                        .is_some_and(|(member, _)| member == name)
            })
            .map(Location::Source)
    }

    /// Whether member `name` exists now.
    pub(crate) fn has_member(&self, name: &str) -> bool {
        self.member_location(name).is_some()
    }

    /// Borrow object `identifier`.
    pub(crate) fn object(&self, identifier: u64) -> Result<&ArchiveObject> {
        let location = self.locate(identifier)?;
        self.archive(location)
            .and_then(|archive| archive.object(identifier))
            .ok_or(Error::MissingObject(identifier))
    }

    /// Mutably borrow object `identifier`, cloning its component first.
    pub(crate) fn object_mut(&mut self, identifier: u64) -> Result<&mut ArchiveObject> {
        let location = self.locate(identifier)?;
        self.archive_mut(location)?
            .object_mut(identifier)
            .ok_or(Error::MissingObject(identifier))
    }

    /// The archive at `location`.
    pub(crate) fn archive_at(&self, location: Location) -> Result<&Archive> {
        self.archive(location)
            .ok_or(Error::Invalid("a component vanished"))
    }

    /// Every live component location, in package order then creation order.
    pub(crate) fn locations(&self) -> Vec<Location> {
        (0..self.source.count())
            .filter(|position| !self.deleted.contains(position))
            .map(Location::Source)
            .chain((0..self.created.len()).map(Location::Created))
            .collect()
    }

    /// Borrow live component locations without allocating a collection.
    pub(crate) fn locations_iter(&self) -> impl Iterator<Item = Location> + '_ {
        (0..self.source.count())
            .filter(|position| !self.deleted.contains(position))
            .map(Location::Source)
            .chain((0..self.created.len()).map(Location::Created))
    }

    /// `(message index, payload)` of the unique `message_type` payload of
    /// object `identifier`.
    pub(crate) fn unique_payload(
        &self,
        identifier: u64,
        message_type: u32,
        reason: &'static str,
    ) -> Result<(usize, &[u8])> {
        unique_payload(self.object(identifier)?, message_type, reason)
    }

    /// Replace payload `message` of object `identifier`.
    pub(crate) fn replace_payload(
        &mut self,
        identifier: u64,
        message: usize,
        data: Vec<u8>,
    ) -> Result<()> {
        let limits = self.limits;
        let object = self.object_mut(identifier)?;
        let type_ = object
            .messages
            .get(message)
            .ok_or(Error::Invalid("a payload index changed"))?
            .type_;
        object.replace_message(message, RawMessage { type_, data }, limits)?;
        Ok(())
    }

    /// Insert `object` at the end of the component at `location`.
    pub(crate) fn insert(&mut self, location: Location, object: ArchiveObject) -> Result<()> {
        let identifier = object
            .archive_info
            .identifier
            .ok_or(Error::Invalid("a new object has no identifier"))?;
        if self.index.contains_key(&identifier) {
            return Err(Error::Invalid("a new object reuses an identifier"));
        }
        let limits = self.limits;
        self.archive_mut(location)?.insert_object(object, limits)?;
        self.index.insert(identifier, location);
        Ok(())
    }

    /// Remove object `identifier` from its component.
    pub(crate) fn remove(&mut self, identifier: u64) -> Result<ArchiveObject> {
        let location = self.locate(identifier)?;
        let removed = self
            .archive_mut(location)?
            .remove_object(identifier)
            .ok_or(Error::MissingObject(identifier))?;
        self.index.remove(&identifier);
        Ok(removed)
    }

    /// Create member `name` holding `archive`, inserted before `before`.
    pub(crate) fn create_member(
        &mut self,
        name: &str,
        archive: Archive,
        before: Option<&str>,
    ) -> Result<()> {
        if self.has_member(name) {
            return Err(Error::Invalid("a new member name is taken"));
        }
        let position = self.created.len();
        for object in &archive.objects {
            let identifier = object
                .archive_info
                .identifier
                .ok_or(Error::Invalid("a new object has no identifier"))?;
            if self
                .index
                .insert(identifier, Location::Created(position))
                .is_some()
            {
                return Err(Error::Invalid("a new object reuses an identifier"));
            }
        }
        self.created.push(Insertion {
            name: name.to_owned(),
            archive,
            before: before.map(str::to_owned),
        });
        Ok(())
    }

    /// Delete the source member at `location` with every object it holds.
    pub(crate) fn delete_member(&mut self, location: Location) -> Result<()> {
        let Location::Source(position) = location else {
            return Err(Error::Invalid(
                "a member created by the edit cannot be deleted",
            ));
        };
        let was_deleted = self.deleted.contains(&position);
        let was_changed = self.changed.contains_key(&position);
        let name_len = self
            .source
            .component(position)
            .ok_or(Error::Invalid("a component vanished"))?
            .0
            .len();
        let identifiers = self
            .archive_at(location)?
            .objects
            .iter()
            .filter_map(|object| object.archive_info.identifier)
            .collect::<Vec<_>>();
        for identifier in identifiers {
            self.index.remove(&identifier);
        }
        if self.deleted.insert(position) && !was_deleted {
            if was_changed {
                self.finish.edit_count = self
                    .finish
                    .edit_count
                    .checked_sub(1)
                    .ok_or(Error::Invalid("finish edit count underflow"))?;
                self.finish.copy_bytes = self
                    .finish
                    .copy_bytes
                    .checked_sub(name_len)
                    .ok_or(Error::Invalid("finish name bytes underflow"))?;
            }
            self.finish.deletion_count = self
                .finish
                .deletion_count
                .checked_add(1)
                .ok_or(Error::Invalid("finish deletion count overflow"))?;
            self.finish.copy_bytes = self
                .finish
                .copy_bytes
                .checked_add(name_len)
                .ok_or(Error::Invalid("finish name bytes overflow"))?;
        }
        Ok(())
    }

    /// The largest identifier of any live object (0 for none).
    pub(crate) fn max_object_identifier(&self) -> u64 {
        self.index.keys().copied().max().unwrap_or(0)
    }

    // --- The package metadata registry -----------------------------------

    /// Whether the package has a registry.
    pub(crate) const fn has_registry(&self) -> bool {
        self.registry.is_some()
    }

    fn registry_payload(&self) -> Result<Option<&[u8]>> {
        let Some((position, object, message)) = self.registry else {
            return Ok(None);
        };
        let archive = self.archive_at(Location::Source(position))?;
        let object = archive
            .object(object)
            .ok_or(Error::Invalid("the registry object vanished"))?;
        Ok(Some(
            object
                .messages
                .get(message)
                .ok_or(Error::Invalid("the registry payload vanished"))?
                .data
                .as_slice(),
        ))
    }

    fn set_registry_payload(&mut self, data: Vec<u8>) -> Result<()> {
        self.registry_requirements.borrow_mut().take();
        self.registry_root.borrow_mut().take();
        let (position, object, message) = self
            .registry
            .ok_or(Error::Invalid("the package has no registry"))?;
        let limits = self.limits;
        let archive = self.archive_mut(Location::Source(position))?;
        archive
            .object_mut(object)
            .ok_or(Error::Invalid("the registry object vanished"))?
            .replace_message(
                message,
                RawMessage {
                    type_: PACKAGE_METADATA_MESSAGE_TYPE,
                    data,
                },
                limits,
            )?;
        Ok(())
    }

    fn update_registry(
        &mut self,
        update: impl FnOnce(&[u8]) -> Result<Option<Vec<u8>>>,
    ) -> Result<()> {
        let Some(payload) = self.registry_payload()? else {
            return Ok(());
        };
        if let Some(data) = update(payload)? {
            self.set_registry_payload(data)?;
        }
        Ok(())
    }

    fn registry_options(&self) -> RewriteOptions {
        let message = self.limits.max_message_bytes().max(1);
        let input = message.min(WireLimits::MAX_INPUT_BYTES);
        let output = message.min(WireLimits::MAX_OUTPUT_BYTES);
        let records = WireLimits::MAX_FIELDS.min(input);
        RewriteOptions::new(
            input,
            output,
            WireLimits::MAX_FIELDS,
            WireLimits::MAX_REWRITE_WORK,
            REGISTRY_RECURSION_LIMIT,
            records,
            records,
            records,
        )
    }

    fn inspect(&self) -> Result<Option<Root>> {
        if let Some(cached) = self.registry_root.borrow().as_ref() {
            return Ok(cached.clone());
        }
        let Some(payload) = self.registry_payload()? else {
            *self.registry_root.borrow_mut() = Some(None);
            return Ok(None);
        };
        let mut root = Root::default();
        let inspection = registry_codec::inspect_package_metadata_with_visitor(
            payload,
            self.registry_options(),
            &mut root,
        )?;
        root.last_object_identifier = inspection.last_object_identifier();
        root.save_token = inspection.save_token();
        *self.registry_root.borrow_mut() = Some(Some(root.clone()));
        Ok(Some(root))
    }

    /// The registry's issued-identifier watermark.
    pub(crate) fn last_object_identifier(&self) -> Result<Option<u64>> {
        Ok(self.inspect()?.map(|root| root.last_object_identifier))
    }

    /// The registry's document save token.
    pub(crate) fn save_token(&self) -> Result<Option<u64>> {
        Ok(self.inspect()?.and_then(|root| root.save_token))
    }

    /// Whether the registry names a `DataMetadataMap`.
    pub(crate) fn has_data_metadata_map(&self) -> Result<bool> {
        Ok(self
            .inspect()?
            .is_some_and(|root| root.has_data_metadata_map))
    }

    /// The next free package-wide object identifier: one above every live
    /// object and every identifier the registry names.
    pub(crate) fn next_object_identifier(&self) -> Result<u64> {
        let mut maximum = self.max_object_identifier();
        if let Some(root) = self.inspect()? {
            maximum = maximum
                .max(root.last_object_identifier)
                .max(root.maximum_identifier);
        }
        maximum.checked_add(1).ok_or(Error::IdentifierOverflow)
    }

    /// Set the watermark (raising it through the registry codec, lowering
    /// it in place).
    pub(crate) fn set_last_object_identifier(&mut self, identifier: u64) -> Result<()> {
        let Some(current) = self.last_object_identifier()? else {
            return Ok(());
        };
        if identifier == current {
            return Ok(());
        }
        let options = self.registry_options();
        self.update_registry(|payload| {
            if identifier > current {
                let prepared = registry_codec::prepare_package_metadata_rewrite(
                    payload,
                    Batch::new(current, identifier, &[], &[]),
                    options,
                )?;
                let limits = prepared.execution_requirements().exact_limits();
                return Ok(Some(prepared.execute(limits)?.into_bytes()));
            }
            Ok(Some(wire::patch_varint_field(
                payload,
                METADATA_LAST_OBJECT_IDENTIFIER,
                true,
                Some(identifier),
            )?))
        })
    }

    /// Lower the watermark to the largest surviving object identifier when
    /// a removal took out the object that holds it.
    pub(crate) fn release_identifier_suffix(&mut self, removed: &[u64]) -> Result<()> {
        let Some(last) = self.last_object_identifier()? else {
            return Ok(());
        };
        if !removed.contains(&last) {
            return Ok(());
        }
        let maximum = self.max_object_identifier();
        if maximum >= last {
            return Err(Error::Invalid(
                "an object at or above the released watermark remains",
            ));
        }
        self.set_last_object_identifier(maximum)
    }

    /// The registered (current) component of member `name`.
    pub(crate) fn component_for_member(&self, name: &str) -> Result<Option<u64>> {
        let locator = name
            .strip_prefix("Index/")
            .and_then(|name| name.strip_suffix(".iwa"))
            .ok_or(Error::Invalid("a component member name is not Index/*.iwa"))?;
        let Some(root) = self.inspect()? else {
            return Ok(None);
        };
        let mut found = None;
        for (identifier, candidate) in &root.components {
            if candidate == locator && found.replace(*identifier).is_some() {
                return Err(Error::Invalid("several components share one locator"));
            }
        }
        Ok(found)
    }

    /// Exact prepared requirements for one package-registry inspection.
    pub(crate) fn registry_inspection_requirements(
        &self,
    ) -> Result<Option<registry_codec::RewriteReport>> {
        if let Some(cached) = self.registry_requirements.borrow().as_ref() {
            return Ok(*cached);
        }
        let Some(payload) = self.registry_payload()? else {
            *self.registry_requirements.borrow_mut() = Some(None);
            return Ok(None);
        };
        let report = registry_codec::inspect_package_metadata_requirements(
            payload,
            self.registry_options(),
        )?;
        *self.registry_requirements.borrow_mut() = Some(Some(report));
        Ok(Some(report))
    }

    /// Exact wire work required by one package-registry inspection.
    pub(crate) fn registry_inspection_work(&self) -> usize {
        // Callers use this as the admission for the next `inspect`-backed
        // lookup.  Once that lookup has populated the root cache, charging
        // the full two-pass codec report again would account preparation as
        // work that is not performed.  `set_registry_payload` clears both
        // caches, so every changed registry state is admitted again.
        if self.registry_root.borrow().is_some() {
            return 0;
        }
        self.registry_inspection_requirements()
            .ok()
            .flatten()
            .map_or(0, |report| report.work_bytes())
    }

    /// The current registry payload extent used by a bounded metadata plan.
    pub(crate) fn registry_payload_len(&self) -> Result<usize> {
        Ok(self.registry_payload()?.map_or(0, <[u8]>::len))
    }

    /// The effective locator of current component `component`.
    pub(crate) fn component_locator(&self, component: u64) -> Result<Option<String>> {
        Ok(self.inspect()?.and_then(|root| {
            root.components
                .into_iter()
                .find(|(identifier, _)| *identifier == component)
                .map(|(_, locator)| locator)
        }))
    }

    /// The object identifiers in the UUID map of `component`.
    pub(crate) fn component_uuid_objects(&self, component: u64) -> Result<HashSet<u64>> {
        let Some(payload) = self.registry_payload()? else {
            return Ok(HashSet::new());
        };
        let mut found = None;
        for data in repeated(payload, METADATA_COMPONENTS)? {
            if varint(data, COMPONENT_IDENTIFIER)? != Some(component) {
                continue;
            }
            let mut objects = HashSet::new();
            for entry in repeated(data, COMPONENT_OBJECT_UUIDS)? {
                let identifier = varint(entry, UUID_ENTRY_IDENTIFIER)?
                    .ok_or(Error::Invalid("a UUID entry has no object"))?;
                if !objects.insert(identifier) {
                    return Err(Error::Invalid("a UUID map repeats an object"));
                }
            }
            if found.replace(objects).is_some() {
                return Err(Error::Invalid("a component is registered twice"));
            }
        }
        found.ok_or(Error::Invalid("a component is not registered"))
    }

    fn existing_uuids(payload: &[u8]) -> Result<HashSet<(u64, u64)>> {
        let mut uuids = HashSet::new();
        for component in repeated(payload, METADATA_COMPONENTS)? {
            for entry in repeated(component, COMPONENT_OBJECT_UUIDS)? {
                if let Some(uuid) = repeated(entry, UUID_ENTRY_UUID)?.first() {
                    uuids.insert((varint(uuid, 1)?.unwrap_or(0), varint(uuid, 2)?.unwrap_or(0)));
                }
            }
        }
        Ok(uuids)
    }

    pub(crate) fn fresh_uuid(
        &mut self,
        identifier: u64,
        existing: &mut HashSet<(u64, u64)>,
    ) -> (u64, u64) {
        loop {
            self.uuid_seed = self.uuid_seed.wrapping_add(1);
            let lower = mix(identifier ^ 0x6c69_7463_6869_7374, self.uuid_seed);
            let upper = mix(lower, identifier.rotate_left(17) ^ 0x7374_796c_652d_7575);
            // An RFC 4122 version 4, variant 1 identifier.
            let upper = (upper & !0xf000) | 0x4000;
            let lower = (lower & !0xc0) | 0x80;
            if existing.insert((lower, upper)) {
                return (lower, upper);
            }
        }
    }

    /// Append UUID entries for `objects` to the map of `component`.
    pub(crate) fn add_object_uuids(&mut self, component: u64, objects: &[u64]) -> Result<()> {
        if objects.is_empty() || !self.has_registry() {
            return Ok(());
        }
        let requested = objects.iter().copied().collect::<HashSet<_>>();
        if requested.len() != objects.len() {
            return Err(Error::Invalid("a UUID allocation repeats an object"));
        }
        let payload = self.registry_payload()?.unwrap_or_default().to_vec();
        let mut existing = Self::existing_uuids(&payload)?;
        let mut entries = Vec::with_capacity(objects.len());
        for &identifier in objects {
            let (lower, upper) = self.fresh_uuid(identifier, &mut existing);
            let mut uuid = Vec::with_capacity(22);
            push_varint_field(&mut uuid, 1, lower);
            push_varint_field(&mut uuid, 2, upper);
            let mut entry = Vec::with_capacity(uuid.len() + 12);
            push_varint_field(&mut entry, UUID_ENTRY_IDENTIFIER, identifier);
            push_message_field(&mut entry, UUID_ENTRY_UUID, &uuid);
            entries.push(entry);
        }
        let mut count = 0usize;
        let data = wire::transform_length_delimited_fields_at_path(
            &payload,
            &[METADATA_COMPONENTS],
            |data| -> Result<Vec<u8>> {
                if varint(data, COMPONENT_IDENTIFIER)? != Some(component) {
                    return Ok(data.to_vec());
                }
                count += 1;
                for entry in repeated(data, COMPONENT_OBJECT_UUIDS)? {
                    if varint(entry, UUID_ENTRY_IDENTIFIER)?
                        .is_some_and(|identifier| requested.contains(&identifier))
                    {
                        return Err(Error::Invalid("an object already has a UUID"));
                    }
                }
                let mut output = data.to_vec();
                for entry in &entries {
                    output = wire::append_repeated_length_delimited_field(
                        &output,
                        COMPONENT_OBJECT_UUIDS,
                        entry,
                    )?;
                }
                Ok(output)
            },
        )?;
        if count != 1 {
            return Err(Error::Invalid("a component must be registered once"));
        }
        self.set_registry_payload(data)
    }

    /// Remove the UUID entries of `objects` from the map of `component`.
    pub(crate) fn remove_object_uuids(&mut self, component: u64, objects: &[u64]) -> Result<()> {
        if objects.is_empty() || !self.has_registry() {
            return Ok(());
        }
        let mut count = 0usize;
        self.update_registry(|payload| {
            let data = wire::transform_length_delimited_fields_at_path(
                payload,
                &[METADATA_COMPONENTS],
                |data| -> Result<Vec<u8>> {
                    if varint(data, COMPONENT_IDENTIFIER)? != Some(component) {
                        return Ok(data.to_vec());
                    }
                    count += 1;
                    let mut output = data.to_vec();
                    for &identifier in objects {
                        output = wire::remove_repeated_length_delimited_field_where(
                            &output,
                            COMPONENT_OBJECT_UUIDS,
                            |entry| -> Result<bool> {
                                Ok(varint(entry, UUID_ENTRY_IDENTIFIER)? == Some(identifier))
                            },
                        )?;
                    }
                    Ok(output)
                },
            )?;
            if count != 1 {
                return Err(Error::Invalid("a component must be registered once"));
            }
            Ok(Some(data))
        })
    }

    /// Append the external reference `source -> (target, object)` unless
    /// `source` already lists it.
    pub(crate) fn add_external_reference(
        &mut self,
        source: u64,
        target: u64,
        object: u64,
    ) -> Result<()> {
        let mut reference = Vec::with_capacity(22);
        push_varint_field(&mut reference, EXTERNAL_COMPONENT, target);
        push_varint_field(&mut reference, EXTERNAL_OBJECT, object);
        let mut sources = 0usize;
        self.update_registry(|payload| {
            let mut changed = false;
            let data = wire::transform_length_delimited_fields_at_path(
                payload,
                &[METADATA_COMPONENTS],
                |data| -> Result<Vec<u8>> {
                    if varint(data, COMPONENT_IDENTIFIER)? != Some(source) {
                        return Ok(data.to_vec());
                    }
                    sources += 1;
                    let mut existing = 0usize;
                    for candidate in repeated(data, COMPONENT_EXTERNAL_REFERENCES)? {
                        if external_matches(candidate, target, Some(object), true)? {
                            existing += 1;
                        }
                    }
                    if existing > 1 {
                        return Err(Error::Invalid("an external reference is repeated"));
                    }
                    if existing == 1 {
                        return Ok(data.to_vec());
                    }
                    changed = true;
                    Ok(wire::append_repeated_length_delimited_field(
                        data,
                        COMPONENT_EXTERNAL_REFERENCES,
                        &reference,
                    )?)
                },
            )?;
            if sources != 1 {
                return Err(Error::Invalid("a component must be registered once"));
            }
            Ok(changed.then_some(data))
        })
    }

    /// Remove the external reference `source -> (target, object)` if listed.
    pub(crate) fn remove_external_reference(
        &mut self,
        source: u64,
        target: u64,
        object: u64,
    ) -> Result<()> {
        let mut sources = 0usize;
        self.update_registry(|payload| {
            let mut changed = false;
            let data = wire::transform_length_delimited_fields_at_path(
                payload,
                &[METADATA_COMPONENTS],
                |data| -> Result<Vec<u8>> {
                    if varint(data, COMPONENT_IDENTIFIER)? != Some(source) {
                        return Ok(data.to_vec());
                    }
                    sources += 1;
                    let mut matches = 0usize;
                    for candidate in repeated(data, COMPONENT_EXTERNAL_REFERENCES)? {
                        if external_matches(candidate, target, Some(object), false)? {
                            matches += 1;
                        }
                    }
                    if matches > 1 {
                        return Err(Error::Invalid("an external reference is repeated"));
                    }
                    if matches == 0 {
                        return Ok(data.to_vec());
                    }
                    changed = true;
                    wire::remove_repeated_length_delimited_field_where(
                        data,
                        COMPONENT_EXTERNAL_REFERENCES,
                        |candidate| external_matches(candidate, target, Some(object), false),
                    )
                },
            )?;
            if sources != 1 {
                return Err(Error::Invalid("a component must be registered once"));
            }
            Ok(changed.then_some(data))
        })
    }

    /// Remove every component's external reference to `(target, object)`.
    pub(crate) fn remove_external_references_to(&mut self, target: u64, object: u64) -> Result<()> {
        self.update_registry(|payload| {
            let mut changed = false;
            let data = wire::transform_length_delimited_fields_at_path(
                payload,
                &[METADATA_COMPONENTS],
                |data| -> Result<Vec<u8>> {
                    let mut matches = 0usize;
                    for candidate in repeated(data, COMPONENT_EXTERNAL_REFERENCES)? {
                        if external_matches(candidate, target, Some(object), false)? {
                            matches += 1;
                        }
                    }
                    if matches > 1 {
                        return Err(Error::Invalid("an external reference is repeated"));
                    }
                    if matches == 0 {
                        return Ok(data.to_vec());
                    }
                    changed = true;
                    wire::remove_repeated_length_delimited_field_where(
                        data,
                        COMPONENT_EXTERNAL_REFERENCES,
                        |candidate| external_matches(candidate, target, Some(object), false),
                    )
                },
            )?;
            Ok(changed.then_some(data))
        })
    }

    /// Append an encoded `TSP.ComponentInfo` registration.
    pub(crate) fn add_component(&mut self, identifier: u64, encoded: &[u8]) -> Result<()> {
        self.update_registry(|payload| {
            if repeated(payload, METADATA_COMPONENTS)?
                .into_iter()
                .any(|component| {
                    varint(component, COMPONENT_IDENTIFIER).ok().flatten() == Some(identifier)
                })
            {
                return Err(Error::Invalid("a new component identifier is registered"));
            }
            Ok(Some(wire::append_repeated_length_delimited_field(
                payload,
                METADATA_COMPONENTS,
                encoded,
            )?))
        })
    }

    /// Remove the registration of `component` (at `locator`).
    pub(crate) fn remove_component(&mut self, component: u64, locator: &str) -> Result<()> {
        let options = self.registry_options();
        let options = RewriteOptions::new(
            options.max_input_bytes(),
            options.max_output_bytes(),
            options.max_fields(),
            options.max_work_bytes(),
            options.recursion_limit(),
            options.max_components(),
            options.max_references(),
            options.max_additions().max(1),
        );
        self.update_registry(|payload| {
            Ok(Some(
                registry_codec::rewrite_package_metadata_component_removal(
                    payload,
                    ComponentSelector::new(component, locator),
                    options,
                )?
                .into_bytes(),
            ))
        })
    }
}

/// What one strict registry inspection reports.
#[derive(Default, Clone)]
struct Root {
    last_object_identifier: u64,
    save_token: Option<u64>,
    maximum_identifier: u64,
    has_data_metadata_map: bool,
    /// `(identifier, effective locator)` of every current component.
    components: Vec<(u64, String)>,
}

impl registry_codec::PackageMetadataVisitor for Root {
    fn visit_component(
        &mut self,
        component: registry_codec::ComponentDescriptor<'_>,
    ) -> std::result::Result<(), RewriteError> {
        self.maximum_identifier = self.maximum_identifier.max(component.identifier());
        if component.is_current() {
            self.components.push((
                component.identifier(),
                component.effective_locator().to_owned(),
            ));
        }
        Ok(())
    }

    fn visit_object_uuid(
        &mut self,
        binding: registry_codec::ObjectUuidDescriptor<'_>,
    ) -> std::result::Result<(), RewriteError> {
        self.maximum_identifier = self.maximum_identifier.max(binding.object_identifier());
        Ok(())
    }

    fn visit_external_reference(
        &mut self,
        reference: registry_codec::ExternalReferenceDescriptor<'_>,
    ) -> std::result::Result<(), RewriteError> {
        if let Some(identifier) = reference.object_identifier() {
            self.maximum_identifier = self.maximum_identifier.max(identifier);
        }
        Ok(())
    }

    fn visit_data_reference_owner(
        &mut self,
        owner: registry_codec::DataReferenceOwnerDescriptor<'_>,
    ) -> std::result::Result<(), RewriteError> {
        self.maximum_identifier = self.maximum_identifier.max(owner.object_identifier());
        Ok(())
    }

    fn visit_ambiguous_object_identifier(
        &mut self,
        _component: registry_codec::ComponentDescriptor<'_>,
        identifier: u64,
    ) -> std::result::Result<(), RewriteError> {
        self.maximum_identifier = self.maximum_identifier.max(identifier);
        Ok(())
    }

    fn visit_data_metadata_map(
        &mut self,
        object_identifier: u64,
        _has_unknown_fields: bool,
    ) -> std::result::Result<(), RewriteError> {
        self.has_data_metadata_map = true;
        self.maximum_identifier = self.maximum_identifier.max(object_identifier);
        Ok(())
    }
}

/// `(message index, payload)` of the unique `message_type` payload.
pub(crate) fn unique_payload<'a>(
    object: &'a ArchiveObject,
    message_type: u32,
    reason: &'static str,
) -> Result<(usize, &'a [u8])> {
    let mut found = None;
    for (index, message) in object.messages.iter().enumerate() {
        if message.type_ == message_type && found.replace(index).is_some() {
            return Err(Error::Invalid(reason));
        }
    }
    let index = found.ok_or(Error::Invalid(reason))?;
    Ok((index, object.messages[index].data.as_slice()))
}

fn owned_name(name: &str) -> std::result::Result<String, FinishError> {
    let mut owned = String::new();
    owned
        .try_reserve_exact(name.len())
        .map_err(|_| FinishError::Allocation { amount: name.len() })?;
    owned.push_str(name);
    Ok(owned)
}

fn mix(value: u64, salt: u64) -> u64 {
    // SplitMix64 finaliser.
    let mut z = value
        .wrapping_add(salt.wrapping_mul(0x9e37_79b9_7f4a_7c15))
        .wrapping_add(0x9e37_79b9_7f4a_7c15);
    z = (z ^ (z >> 30)).wrapping_mul(0xbf58_476d_1ce4_e5b9);
    z = (z ^ (z >> 27)).wrapping_mul(0x94d0_49bb_1331_11eb);
    z ^ (z >> 31)
}

pub(crate) fn fields(data: &[u8]) -> Result<Vec<WireField>> {
    Ok(wire::parse_wire_fields(data, WireLimits::DEFAULT)?)
}

fn payload_of(data: &[u8], field: WireField) -> Result<&[u8]> {
    if field.wire_type() != 2 {
        return Err(Error::Invalid("a message field has the wrong wire type"));
    }
    Ok(field.payload(data)?)
}

/// Every payload of repeated length-delimited field `number`.
pub(crate) fn repeated(data: &[u8], number: u32) -> Result<Vec<&[u8]>> {
    fields(data)?
        .into_iter()
        .filter(|field| field.number() == number)
        .map(|field| payload_of(data, field))
        .collect()
}

/// The last value of varint field `number` (`None` when absent).
pub(crate) fn varint(data: &[u8], number: u32) -> Result<Option<u64>> {
    let mut value = None;
    for field in fields(data)? {
        if field.number() != number {
            continue;
        }
        if field.wire_type() != 0 {
            return Err(Error::Invalid("a varint field has the wrong wire type"));
        }
        let (decoded, _) = decode_varint_from_bytes(&data[field.key_end()..field.end()])
            .map_err(|_| Error::Invalid("a varint is malformed"))?;
        value = Some(decoded);
    }
    Ok(value)
}

fn external_matches(
    candidate: &[u8],
    target: u64,
    object: Option<u64>,
    require_strong_default: bool,
) -> Result<bool> {
    let matches = varint(candidate, EXTERNAL_COMPONENT)? == Some(target)
        && varint(candidate, EXTERNAL_OBJECT)? == object;
    if require_strong_default {
        return Ok(matches && varint(candidate, EXTERNAL_IS_WEAK)?.is_none());
    }
    Ok(matches)
}

pub(crate) fn push_varint(output: &mut Vec<u8>, value: u64) {
    litchi_iwa_common::encode_varint_into(output, value);
}

pub(crate) fn push_varint_field(output: &mut Vec<u8>, number: u32, value: u64) {
    push_varint(output, u64::from(number) << 3);
    push_varint(output, value);
}

pub(crate) fn push_message_field(output: &mut Vec<u8>, number: u32, payload: &[u8]) {
    push_varint(output, (u64::from(number) << 3) | 2);
    push_varint(output, payload.len() as u64);
    output.extend_from_slice(payload);
}

pub(crate) fn push_packed_field(output: &mut Vec<u8>, number: u32, values: &[u32]) {
    let mut packed = Vec::with_capacity(values.len() * 2);
    for &value in values {
        push_varint(&mut packed, u64::from(value));
    }
    push_message_field(output, number, &packed);
}

#[cfg(test)]
mod tests;
