//! Core-backed composition and non-applying merge plans for RTF edits.

use super::{Edit, Error, Limits, Operation, Snapshot};
use litchi_core::patch as core;
use serde_json::Value;
use std::cmp::Ordering;
use std::fmt;

pub use core::CompositionLimits;
pub use core::MergeChoice as MergeResolution;

#[derive(Clone)]
struct Lineage(Snapshot);

impl PartialEq for Lineage {
    fn eq(&self, other: &Self) -> bool {
        self.0.same_snapshot(&other.0)
    }
}

impl Eq for Lineage {}

/// One format-owned composition conflict.
#[derive(Debug, Clone, PartialEq, Eq)]
#[non_exhaustive]
pub enum CompositionConflict {
    /// Both branches used the same stable sub-edit identifier.
    DuplicateId(String),
    /// At least one branch writes a semantic facet used by the other.
    Effect {
        effect: String,
        left: String,
        right: String,
    },
    /// Two branches publish through incompatible RTF artifact domains.
    PublicationDomain { left: String, right: String },
    /// A newer common conflict kind not understood by this crate version.
    Unknown,
}

/// Deterministically ordered RTF composition conflicts.
pub type ConflictSet = core::ConflictSet<CompositionConflict>;

/// Failure while preparing, joining, or planning RTF sub-edits.
#[derive(Debug, Clone)]
#[non_exhaustive]
pub enum CompositionError {
    /// The common bounded composition layer rejected the request.
    Core(String),
    /// Independently prepared work overlaps.
    Conflicts(ConflictSet),
    /// The finished composition failed ordinary RTF validation.
    Edit(Error),
    /// A durable patch could not be admitted to the conservative RTF seam.
    Durable(String),
}

impl fmt::Display for CompositionError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Core(reason) => write!(formatter, "RTF composition failed: {reason}"),
            Self::Conflicts(conflicts) => {
                write!(
                    formatter,
                    "RTF composition has {} conflict(s)",
                    conflicts.len()
                )
            },
            Self::Edit(error) => error.fmt(formatter),
            Self::Durable(reason) => write!(formatter, "RTF durable composition failed: {reason}"),
        }
    }
}

impl std::error::Error for CompositionError {}

impl From<Error> for CompositionError {
    fn from(error: Error) -> Self {
        Self::Edit(error)
    }
}

/// Independently prepared work against one exact immutable RTF snapshot.
pub struct Prepared {
    inner: core::SubEdit<Lineage, Vec<Operation>>,
}

impl Prepared {
    /// Stable caller-selected identifier.
    #[must_use]
    pub fn id(&self) -> &str {
        self.inner.id()
    }

    /// Number of staged semantic operations.
    #[must_use]
    pub fn operation_count(&self) -> usize {
        self.inner.payload().len()
    }
}

/// A bounded deterministic collection of provably disjoint RTF sub-edits.
pub struct Composition {
    source: Snapshot,
    inner: core::JoinedSubEdits<Lineage, Vec<Operation>>,
}

impl fmt::Debug for Composition {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Composition")
            .field("sub_edits", &self.inner.len())
            .field("effects", &self.inner.total_effects())
            .finish_non_exhaustive()
    }
}

impl Composition {
    /// Starts an empty composition for one exact snapshot.
    #[must_use]
    pub fn new(source: &Snapshot, limits: CompositionLimits) -> Self {
        let lineage = Lineage(source.clone());
        Self {
            source: source.clone(),
            inner: core::JoinedSubEdits::new(lineage, limits),
        }
    }

    /// Number of accepted sub-edits.
    #[must_use]
    pub fn len(&self) -> usize {
        self.inner.len()
    }

    /// Whether no work has been accepted.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.inner.is_empty()
    }

    /// Joins one independent edit only when common effect analysis proves it disjoint.
    ///
    /// # Errors
    /// Returns bounded, lineage, duplicate-ID, or typed overlap details.
    pub fn join(&mut self, incoming: Prepared) -> Result<&mut Self, CompositionError> {
        for existing in self.inner.sub_edits() {
            if let Some(conflict) =
                publication_domain_conflict(existing.payload(), incoming.inner.payload())
            {
                return Err(CompositionError::Conflicts(ConflictSet::new([
                    conflict_for_domain(existing.id(), incoming.inner.id(), conflict),
                ])));
            }
            if body_interval_conflict(existing.payload(), incoming.inner.payload()) {
                return Err(CompositionError::Conflicts(ConflictSet::new([
                    CompositionConflict::Effect {
                        effect: "body:span:overlap".to_string(),
                        left: existing.id().to_string(),
                        right: incoming.inner.id().to_string(),
                    },
                ])));
            }
        }
        self.inner.join(incoming.inner).map_err(|error| {
            let (failure, _rejected) = error.into_parts();
            match failure {
                core::SubEditJoinFailure::Overlap(conflicts) => {
                    CompositionError::Conflicts(map_conflicts(&conflicts))
                },
                core::SubEditJoinFailure::DifferentLineage
                | core::SubEditJoinFailure::DifferentLimits
                | core::SubEditJoinFailure::DuplicateId
                | core::SubEditJoinFailure::Limit(_) => {
                    CompositionError::Core("common join refusal".to_string())
                },
                _ => CompositionError::Core("unknown common join failure".to_string()),
            }
        })?;
        Ok(self)
    }

    /// Converts the joined work to an ordinary immutable edit without applying it.
    #[must_use]
    pub fn into_edit(self) -> Edit {
        let operations = self
            .inner
            .into_sub_edits()
            .flat_map(core::SubEdit::into_payload)
            .collect::<Vec<_>>();
        let replacement_bytes = operations.iter().fold(0usize, |total, operation| {
            total.saturating_add(operation.replacement_bytes())
        });
        Edit {
            source: self.source,
            limits: Limits::new(operations.len()),
            operations,
            replacement_bytes,
        }
    }

    /// Validates and commits the joined work atomically.
    ///
    /// # Errors
    /// Returns the ordinary RTF transaction refusal.
    pub fn commit(self) -> Result<super::Commit, Error> {
        self.into_edit().commit()
    }
}

/// A non-mutating bounded three-way plan for two branches from one exact base.
pub struct MergePlan {
    source: Snapshot,
    inner: core::ThreeWayMergePlan<Lineage, Vec<Operation>>,
    conflicts: ConflictSet,
}

impl fmt::Debug for MergePlan {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("MergePlan")
            .field("conflicts", &self.conflicts)
            .finish_non_exhaustive()
    }
}

impl MergePlan {
    /// Plans a merge without applying either branch.
    ///
    /// # Errors
    /// Returns both-branch validation or finite-bound failures.
    pub fn new(left: Composition, right: Composition) -> Result<Self, CompositionError> {
        for left_edit in left.inner.sub_edits() {
            for right_edit in right.inner.sub_edits() {
                if let Some(conflict) =
                    publication_domain_conflict(left_edit.payload(), right_edit.payload())
                {
                    return Err(CompositionError::Conflicts(ConflictSet::new([
                        conflict_for_domain(left_edit.id(), right_edit.id(), conflict),
                    ])));
                }
                if body_interval_conflict(left_edit.payload(), right_edit.payload()) {
                    return Err(CompositionError::Conflicts(ConflictSet::new([
                        CompositionConflict::Effect {
                            effect: "body:span:overlap".to_string(),
                            left: left_edit.id().to_string(),
                            right: right_edit.id().to_string(),
                        },
                    ])));
                }
            }
        }
        let source = left.source.clone();
        let inner = core::ThreeWayMergePlan::new(left.inner, right.inner)
            .map_err(|error| CompositionError::Core(format!("{error:?}")))?;
        let conflicts = map_conflicts(inner.conflicts());
        Ok(Self {
            source,
            inner,
            conflicts,
        })
    }

    /// Typed conflicts without applying either branch.
    #[must_use]
    pub const fn conflicts(&self) -> &ConflictSet {
        &self.conflicts
    }

    /// Resolves the conservative conflicting branch group.
    pub fn resolve(&mut self, choice: MergeResolution) -> &mut Self {
        self.inner.resolve(choice);
        self
    }

    /// Finishes the plan into still-uncommitted joined work.
    ///
    /// # Errors
    /// Returns this plan unchanged while conflicts remain unresolved.
    pub fn finish(self) -> Result<Composition, Box<Self>> {
        let Self {
            source,
            inner,
            conflicts,
        } = self;
        match inner.finish() {
            Ok(joined) => Ok(Composition {
                source,
                inner: joined,
            }),
            Err(unresolved) => Err(Box::new(Self {
                source,
                inner: *unresolved,
                conflicts,
            })),
        }
    }
}

impl Edit {
    /// Wraps this independently prepared edit in the common bounded effect model.
    ///
    /// # Errors
    /// Returns an identifier, effect, or finite composition-bound error.
    pub fn into_sub_edit(
        self,
        identifier: impl Into<String>,
        limits: CompositionLimits,
    ) -> Result<Prepared, CompositionError> {
        let (reads, writes) = operation_effects(&self.operations);
        let lineage = Lineage(self.source.clone());
        let inner = core::SubEdit::new(lineage, limits, identifier, reads, writes, self.operations)
            .map_err(|error| CompositionError::Core(error.to_string()))?;
        Ok(Prepared { inner })
    }
}

fn operation_effects(operations: &[Operation]) -> (Vec<String>, Vec<String>) {
    let mut reads = Vec::new();
    let mut writes = Vec::new();
    for operation in operations {
        if operation.is_root_transfer() {
            reads.push("rtf:ordinary-root".to_string());
            writes.push("rtf:ordinary-root".to_string());
            continue;
        }
        reads.push("body:structure".to_string());
        writes.extend(operation.effect_keys());
    }
    (reads, writes)
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum PublicationDomain {
    Ordinary,
    Lifecycle,
    Destination,
    PicturePayload,
    PictureRemoval,
    RootTransfer,
}

impl PublicationDomain {
    const fn name(self) -> &'static str {
        match self {
            Self::Ordinary => "ordinary",
            Self::Lifecycle => "lifecycle",
            Self::Destination => "destination",
            Self::PicturePayload => "picture-payload",
            Self::PictureRemoval => "picture-removal",
            Self::RootTransfer => "root-transfer",
        }
    }
}

fn operation_publication_domain(operation: &Operation) -> PublicationDomain {
    match operation {
        Operation::Text {
            structural: true, ..
        }
        | Operation::InsertParagraph { .. }
        | Operation::RemoveParagraph { .. }
        | Operation::RestoreParagraph { .. }
        | Operation::MoveParagraph { .. } => PublicationDomain::Lifecycle,
        Operation::Text { .. }
        | Operation::Alignment { .. }
        | Operation::ParagraphLayout { .. }
        | Operation::Bold { .. }
        | Operation::Italic { .. }
        | Operation::Underline { .. }
        | Operation::FontSize { .. }
        | Operation::Strike { .. }
        | Operation::DoubleStrike { .. }
        | Operation::Hidden { .. }
        | Operation::SmallCaps { .. }
        | Operation::AllCaps { .. } => PublicationDomain::Ordinary,
        Operation::TableCellText { .. }
        | Operation::HeaderFooterText { .. }
        | Operation::AnnotationText { .. }
        | Operation::NoteText { .. }
        | Operation::ShapeText { .. } => PublicationDomain::Destination,
        Operation::PicturePayload(_) => PublicationDomain::PicturePayload,
        Operation::PictureRemoval(_) => PublicationDomain::PictureRemoval,
        Operation::RootTransfer { .. } => PublicationDomain::RootTransfer,
    }
}

const fn publication_domain_bit(domain: PublicationDomain) -> u8 {
    match domain {
        PublicationDomain::Ordinary => 1 << 0,
        PublicationDomain::Lifecycle => 1 << 1,
        PublicationDomain::Destination => 1 << 2,
        PublicationDomain::PicturePayload => 1 << 3,
        PublicationDomain::PictureRemoval => 1 << 4,
        PublicationDomain::RootTransfer => 1 << 5,
    }
}

fn publication_domain_mask(operations: &[Operation]) -> u8 {
    operations.iter().fold(0, |mask, operation| {
        mask | publication_domain_bit(operation_publication_domain(operation))
    })
}

fn publication_domain_conflict(
    left: &[Operation],
    right: &[Operation],
) -> Option<(PublicationDomain, PublicationDomain)> {
    let left_mask = publication_domain_mask(left);
    let right_mask = publication_domain_mask(right);
    if left_mask == 0 || right_mask == 0 {
        return None;
    }
    for left_domain in [
        PublicationDomain::Ordinary,
        PublicationDomain::Lifecycle,
        PublicationDomain::Destination,
        PublicationDomain::PicturePayload,
        PublicationDomain::PictureRemoval,
        PublicationDomain::RootTransfer,
    ] {
        if left_mask & publication_domain_bit(left_domain) == 0 {
            continue;
        }
        for right_domain in [
            PublicationDomain::Ordinary,
            PublicationDomain::Lifecycle,
            PublicationDomain::Destination,
            PublicationDomain::PicturePayload,
            PublicationDomain::PictureRemoval,
            PublicationDomain::RootTransfer,
        ] {
            if right_mask & publication_domain_bit(right_domain) != 0 && left_domain != right_domain
            {
                return Some((left_domain, right_domain));
            }
        }
    }
    None
}

fn conflict_for_domain(
    left_id: &str,
    right_id: &str,
    domains: (PublicationDomain, PublicationDomain),
) -> CompositionConflict {
    CompositionConflict::PublicationDomain {
        left: format!("{}:{left_id}", domains.0.name()),
        right: format!("{}:{right_id}", domains.1.name()),
    }
}

fn body_interval_conflict(left: &[Operation], right: &[Operation]) -> bool {
    // The legacy core-backed seam keeps this operation-level check bounded by
    // `CompositionLimits`; the durable seam below uses sorted interval
    // sweeps.  Keep this conservative fallback separate from the durable
    // admission path rather than exposing an unbounded last-writer merge.
    left.iter().any(|left_operation| {
        let Some(left_span) = body_span(left_operation) else {
            return false;
        };
        right.iter().any(|right_operation| {
            body_span(right_operation)
                .is_some_and(|right_span| spans_overlap(left_span, right_span))
        })
    })
}

fn body_span(operation: &Operation) -> Option<super::TextSpan> {
    match operation {
        Operation::Text { span, .. }
        | Operation::Bold { span, .. }
        | Operation::Italic { span, .. }
        | Operation::Underline { span, .. }
        | Operation::FontSize { span, .. }
        | Operation::Strike { span, .. }
        | Operation::DoubleStrike { span, .. }
        | Operation::Hidden { span, .. }
        | Operation::SmallCaps { span, .. }
        | Operation::AllCaps { span, .. }
        | Operation::InsertParagraph { span, .. } => Some(*span),
        _ => None,
    }
}

fn spans_overlap(left: super::TextSpan, right: super::TextSpan) -> bool {
    if left.is_empty() && right.is_empty() {
        return left.start == right.start;
    }
    if left.is_empty() {
        return (right.start..=right.end).contains(&left.start);
    }
    if right.is_empty() {
        return (left.start..=left.end).contains(&right.start);
    }
    left.start < right.end && right.start < left.end
}

fn map_conflicts(conflicts: &core::ConflictSet<core::SubEditConflict>) -> ConflictSet {
    ConflictSet::new(
        conflicts
            .conflicts()
            .iter()
            .map(|conflict| match conflict {
                core::SubEditConflict::DuplicateId(id) => {
                    CompositionConflict::DuplicateId(id.clone())
                },
                core::SubEditConflict::Effect(effect) => CompositionConflict::Effect {
                    effect: effect.effect().to_string(),
                    left: effect.left_id().to_string(),
                    right: effect.right_id().to_string(),
                },
                _ => CompositionConflict::Unknown,
            })
            .collect::<Vec<_>>(),
    )
}

type DurablePatch = core::Patch<core::Reversible>;
type DurableOperation = core::PatchOperation;

/// A conservative durable composition rooted at one exact RTF artifact.
///
/// The composition accepts only source-relative reversible patches whose
/// forward and inverse directions have both been checked against the same
/// source/target pair.  It does not concatenate inverse operations.  The
/// accepted forward operations are applied once to the base and a fresh exact
/// RTF patch is derived from that combined commit, so its inverse preconditions
/// refer to the combined target rather than to an individual branch.
pub struct DurableComposition {
    source: Snapshot,
    limits: core::PatchLimits,
    operations: Vec<DurableOperation>,
    domain: Option<DurableDomain>,
}

impl fmt::Debug for DurableComposition {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("DurableComposition")
            .field("operations", &self.operations.len())
            .field("domain", &self.domain)
            .finish_non_exhaustive()
    }
}

impl DurableComposition {
    /// Starts an empty durable composition for one exact base snapshot.
    #[must_use]
    pub fn new(source: &Snapshot, limits: core::PatchLimits) -> Self {
        Self {
            source: source.clone(),
            limits,
            operations: Vec::new(),
            domain: None,
        }
    }

    /// Exact immutable base snapshot used by this composition.
    #[must_use]
    pub const fn source(&self) -> &Snapshot {
        &self.source
    }

    /// Durable bounds required for every joined branch and the final patch.
    #[must_use]
    pub const fn limits(&self) -> core::PatchLimits {
        self.limits
    }

    /// Number of accepted forward operations.
    #[must_use]
    pub fn len(&self) -> usize {
        self.operations.len()
    }

    /// Whether no forward operation has been accepted.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.operations.is_empty()
    }

    /// Joins one already encoded reversible branch after exact source checks.
    ///
    /// The branch is applied and then reversed before it is admitted.  This
    /// validates both directions, catches forged source/target hashes, and
    /// leaves this composition unchanged on every refusal.
    ///
    /// # Errors
    /// Refuses different limits, a different source digest, unsupported or
    /// mixed publication domains, active/opaque/protected sources, malformed
    /// operations, or overlapping semantic effects.
    pub fn join(
        &mut self,
        patch: core::Patch<core::Reversible>,
    ) -> Result<&mut Self, CompositionError> {
        let incoming = validate_durable_branch(&self.source, self.limits, &patch)?;
        if incoming.is_empty() {
            return Ok(self);
        }
        let incoming_domain = durable_domain(&incoming)?;
        if let Some(existing_domain) = self.domain {
            if existing_domain != incoming_domain {
                return Err(CompositionError::Conflicts(ConflictSet::new([
                    durable_domain_conflict(existing_domain, incoming_domain),
                ])));
            }
        }
        if let Some(conflict) = plan_durable_conflicts(&self.operations, &incoming)?
            .2
            .into_iter()
            .next()
        {
            return Err(CompositionError::Conflicts(ConflictSet::new([conflict])));
        }
        let combined_count = self
            .operations
            .len()
            .checked_add(incoming.len())
            .ok_or_else(|| CompositionError::Durable("operation count overflow".to_string()))?;
        let mut bounded_operations = Vec::new();
        bounded_operations
            .try_reserve(combined_count)
            .map_err(|_error| {
                CompositionError::Durable("operation allocation failed".to_string())
            })?;
        bounded_operations.extend(self.operations.iter().cloned());
        bounded_operations.extend(incoming.iter().cloned());
        bounded_operations.sort_by(compare_durable_operations);
        validate_durable_forward_limit(self.limits, &bounded_operations)?;
        preflight_durable_reversible(&self.source, self.limits, &bounded_operations)?;
        self.operations
            .try_reserve(incoming.len())
            .map_err(|_error| {
                CompositionError::Durable("operation allocation failed".to_string())
            })?;
        self.operations.extend(incoming);
        self.domain = Some(incoming_domain);
        Ok(self)
    }

    /// Applies all accepted forward operations once and derives a fresh exact
    /// reversible durable patch from the combined commit.
    ///
    /// # Errors
    /// Returns a typed conflict, ordinary RTF validation error, or caller
    /// selected durable-limit error without publishing a partial result.
    pub fn finish(self) -> Result<core::Patch<core::Reversible>, CompositionError> {
        let mut operations = self.operations;
        operations.sort_by(compare_durable_operations);
        validate_durable_forward_limit(self.limits, &operations)?;
        let commit = commit_durable_operations(&self.source, &operations)?;
        let patch = commit
            .patch()
            .to_durable(self.limits)
            .map_err(|error| CompositionError::Durable(error.to_string()))?;
        patch
            .to_deterministic_json()
            .map_err(|error| CompositionError::Durable(error.to_string()))?;
        Ok(patch)
    }

    /// Alias for [`Self::finish`] emphasizing that the returned value remains
    /// an un-applied durable patch until a caller applies it to the exact base.
    pub fn to_durable(self) -> Result<core::Patch<core::Reversible>, CompositionError> {
        self.finish()
    }

    /// Applies the combined work and returns the exact in-memory commit.
    ///
    /// This is useful to callers that need the published snapshot alongside
    /// the newly derived durable patch.  No durable inverse is trusted here;
    /// callers can obtain it from `commit.patch().to_durable(...)`.
    pub fn commit(self) -> Result<super::Commit, CompositionError> {
        let mut operations = self.operations;
        operations.sort_by(compare_durable_operations);
        validate_durable_forward_limit(self.limits, &operations)?;
        let commit = commit_durable_operations(&self.source, &operations)?;
        preflight_durable_reversible(&self.source, self.limits, &operations)?;
        Ok(commit)
    }
}

/// A non-applying durable three-way plan for two independently validated
/// branches rooted at one exact RTF base.
pub struct DurableMergePlan {
    source: Snapshot,
    limits: core::PatchLimits,
    automatic: Vec<DurableOperation>,
    left_conflicts: Vec<DurableOperation>,
    right_conflicts: Vec<DurableOperation>,
    conflicts: ConflictSet,
    resolution: Option<MergeResolution>,
}

impl fmt::Debug for DurableMergePlan {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("DurableMergePlan")
            .field("automatic", &self.automatic.len())
            .field("conflicts", &self.conflicts)
            .field("resolution", &self.resolution)
            .finish_non_exhaustive()
    }
}

impl DurableMergePlan {
    /// Plans a deterministic three-way merge without applying either branch.
    ///
    /// Disjoint operations are accepted automatically.  Every overlapping
    /// pair is retained as one conservative conflict group and requires an
    /// explicit [`MergeResolution`] before [`Self::finish`] can yield work.
    pub fn new(
        left: DurableComposition,
        right: DurableComposition,
    ) -> Result<Self, CompositionError> {
        if left.source.source_bytes() != right.source.source_bytes() {
            return Err(CompositionError::Durable(
                "durable merge branches have different base artifacts".to_string(),
            ));
        }
        if left.limits != right.limits {
            return Err(CompositionError::Durable(
                "durable merge branches have different patch limits".to_string(),
            ));
        }
        let source = left.source.clone();
        let limits = left.limits;
        let mut automatic = Vec::new();
        let mut left_conflicts = Vec::new();
        let mut right_conflicts = Vec::new();
        automatic
            .try_reserve(left.operations.len().saturating_add(right.operations.len()))
            .map_err(|_error| {
                CompositionError::Durable("operation allocation failed".to_string())
            })?;
        left_conflicts
            .try_reserve(left.operations.len())
            .map_err(|_error| {
                CompositionError::Durable("operation allocation failed".to_string())
            })?;
        right_conflicts
            .try_reserve(right.operations.len())
            .map_err(|_error| {
                CompositionError::Durable("operation allocation failed".to_string())
            })?;
        let (left_conflicting, right_conflicting, mut conflicts) =
            plan_durable_conflicts(&left.operations, &right.operations)?;
        for (operation, conflicting) in left.operations.iter().zip(&left_conflicting) {
            if !*conflicting {
                automatic.push(operation.clone());
            } else {
                left_conflicts.push(operation.clone());
            }
        }
        for (operation, conflicting) in right.operations.iter().zip(&right_conflicting) {
            if !*conflicting {
                automatic.push(operation.clone());
            } else {
                right_conflicts.push(operation.clone());
            }
        }
        conflicts.sort_by(compare_conflicts);
        conflicts.dedup();
        Ok(Self {
            source,
            limits,
            automatic,
            left_conflicts,
            right_conflicts,
            conflicts: ConflictSet::new(conflicts),
            resolution: None,
        })
    }

    /// Typed deterministic conflicts requiring an explicit choice.
    #[must_use]
    pub const fn conflicts(&self) -> &ConflictSet {
        &self.conflicts
    }

    /// Resolves the complete conservative conflict group.
    pub fn resolve(&mut self, choice: MergeResolution) -> &mut Self {
        self.resolution = Some(choice);
        self
    }

    /// Yields uncommitted durable composition work after explicit resolution.
    ///
    /// If conflicts remain unresolved, the plan is returned unchanged in a
    /// recoverable box.
    pub fn finish(self) -> Result<DurableComposition, Box<Self>> {
        let Self {
            source,
            limits,
            mut automatic,
            left_conflicts,
            right_conflicts,
            conflicts,
            resolution,
        } = self;
        if !conflicts.is_empty() && resolution.is_none() {
            return Err(Box::new(Self {
                source,
                limits,
                automatic,
                left_conflicts,
                right_conflicts,
                conflicts,
                resolution,
            }));
        }
        let selected = match resolution {
            Some(MergeResolution::Left) => Some(left_conflicts.as_slice()),
            Some(MergeResolution::Right) => Some(right_conflicts.as_slice()),
            Some(MergeResolution::Neither) | None | Some(_) => None,
        };
        if durable_merge_limit_ok(&source, limits, &automatic, selected).is_err() {
            return Err(Box::new(Self {
                source,
                limits,
                automatic,
                left_conflicts,
                right_conflicts,
                conflicts,
                resolution,
            }));
        }
        match resolution {
            Some(MergeResolution::Left) => automatic.extend(left_conflicts),
            Some(MergeResolution::Right) => automatic.extend(right_conflicts),
            Some(MergeResolution::Neither) | None => {},
            Some(_) => {},
        }
        automatic.sort_by(compare_durable_operations);
        Ok(DurableComposition {
            source,
            limits,
            domain: automatic
                .first()
                .and_then(|operation| durable_domain(std::slice::from_ref(operation)).ok()),
            operations: automatic,
        })
    }
}

fn plan_durable_conflicts(
    left: &[DurableOperation],
    right: &[DurableOperation],
) -> Result<(Vec<bool>, Vec<bool>, Vec<CompositionConflict>), CompositionError> {
    let mut left_conflicting = Vec::new();
    left_conflicting
        .try_reserve(left.len())
        .map_err(|_error| CompositionError::Durable("conflict allocation failed".to_string()))?;
    left_conflicting.resize(left.len(), false);
    let mut right_conflicting = Vec::new();
    right_conflicting
        .try_reserve(right.len())
        .map_err(|_error| CompositionError::Durable("conflict allocation failed".to_string()))?;
    right_conflicting.resize(right.len(), false);
    let conflict_capacity = left
        .len()
        .checked_add(right.len())
        .ok_or_else(|| CompositionError::Durable("conflict count overflow".to_string()))?;
    let mut conflicts = Vec::new();
    conflicts
        .try_reserve(conflict_capacity)
        .map_err(|_error| CompositionError::Durable("conflict allocation failed".to_string()))?;

    let left_domain = if left.is_empty() {
        None
    } else {
        Some(durable_domain(left)?)
    };
    let right_domain = if right.is_empty() {
        None
    } else {
        Some(durable_domain(right)?)
    };
    if let (Some(left_domain), Some(right_domain)) = (left_domain, right_domain) {
        if left_domain != right_domain {
            left_conflicting.fill(true);
            right_conflicting.fill(true);
            conflicts.push(durable_domain_conflict(left_domain, right_domain));
            return Ok((left_conflicting, right_conflicting, conflicts));
        }
    }

    match left_domain {
        Some(DurableDomain::Ordinary) => {
            let mut left_intervals = Vec::new();
            let mut right_intervals = Vec::new();
            let mut left_alignments = Vec::new();
            let mut right_alignments = Vec::new();
            left_intervals.try_reserve(left.len()).map_err(|_error| {
                CompositionError::Durable("conflict allocation failed".to_string())
            })?;
            right_intervals.try_reserve(right.len()).map_err(|_error| {
                CompositionError::Durable("conflict allocation failed".to_string())
            })?;
            left_alignments.try_reserve(left.len()).map_err(|_error| {
                CompositionError::Durable("conflict allocation failed".to_string())
            })?;
            right_alignments
                .try_reserve(right.len())
                .map_err(|_error| {
                    CompositionError::Durable("conflict allocation failed".to_string())
                })?;
            for (index, operation) in left.iter().enumerate() {
                if operation.op == "paragraph-alignment.set" {
                    left_alignments.push((
                        index,
                        super::parse_paragraph_target(&operation.target)
                            .map_err(CompositionError::Edit)?,
                    ));
                } else {
                    left_intervals.push((
                        index,
                        durable_text_span(operation).ok_or_else(|| {
                            CompositionError::Durable(
                                "ordinary durable operation has no text span".to_string(),
                            )
                        })?,
                    ));
                }
            }
            for (index, operation) in right.iter().enumerate() {
                if operation.op == "paragraph-alignment.set" {
                    right_alignments.push((
                        index,
                        super::parse_paragraph_target(&operation.target)
                            .map_err(CompositionError::Edit)?,
                    ));
                } else {
                    right_intervals.push((
                        index,
                        durable_text_span(operation).ok_or_else(|| {
                            CompositionError::Durable(
                                "ordinary durable operation has no text span".to_string(),
                            )
                        })?,
                    ));
                }
            }
            // Empty body replacements are point insertions.  A point can
            // touch both non-empty intervals on either side of a boundary,
            // while a conventional two-pointer sweep would advance past the
            // first match and miss the second.  Group every body interval in
            // this rare shape conservatively; all are excluded from the
            // automatic merge and one typed conflict asks for an explicit
            // branch choice.  Alignment operations remain independently
            // comparable below.
            let has_empty_interval = left_intervals.iter().any(|(_, span)| span.is_empty())
                || right_intervals.iter().any(|(_, span)| span.is_empty());
            if has_empty_interval && !left_intervals.is_empty() && !right_intervals.is_empty() {
                for (index, _) in &left_intervals {
                    let marker = left_conflicting.get_mut(*index).ok_or_else(|| {
                        CompositionError::Durable(
                            "left conflict index is out of bounds".to_string(),
                        )
                    })?;
                    *marker = true;
                }
                for (index, _) in &right_intervals {
                    let marker = right_conflicting.get_mut(*index).ok_or_else(|| {
                        CompositionError::Durable(
                            "right conflict index is out of bounds".to_string(),
                        )
                    })?;
                    *marker = true;
                }
                let left_operation =
                    left_intervals
                        .first()
                        .map(|(index, _)| *index)
                        .ok_or_else(|| {
                            CompositionError::Durable("left interval group is empty".to_string())
                        })?;
                let right_operation = right_intervals
                    .first()
                    .map(|(index, _)| *index)
                    .ok_or_else(|| {
                        CompositionError::Durable("right interval group is empty".to_string())
                    })?;
                conflicts.push(durable_effect_conflict(
                    durable_operation_ref(left, left_operation)?,
                    durable_operation_ref(right, right_operation)?,
                    "character-or-text",
                ));
            }
            left_intervals.sort_unstable_by_key(|(_, span)| (span.start, span.end));
            right_intervals.sort_unstable_by_key(|(_, span)| (span.start, span.end));
            let mut left_index = 0;
            let mut right_index = 0;
            while let (Some((left_operation, left_span)), Some((right_operation, right_span))) = (
                left_intervals.get(left_index),
                right_intervals.get(right_index),
            ) {
                if spans_overlap(*left_span, *right_span) {
                    record_durable_conflict(
                        &mut left_conflicting,
                        &mut right_conflicting,
                        &mut conflicts,
                        *left_operation,
                        *right_operation,
                        durable_effect_conflict(
                            durable_operation_ref(left, *left_operation)?,
                            durable_operation_ref(right, *right_operation)?,
                            "character-or-text",
                        ),
                    )?;
                    match (
                        left_span.is_empty(),
                        right_span.is_empty(),
                        left_span.end.cmp(&right_span.end),
                    ) {
                        (true, false, _) => left_index += 1,
                        (false, true, _) => right_index += 1,
                        (true, true, _) | (_, _, Ordering::Equal) => {
                            left_index += 1;
                            right_index += 1;
                        },
                        (_, _, Ordering::Less) => left_index += 1,
                        (_, _, Ordering::Greater) => right_index += 1,
                    }
                } else if left_span.end < right_span.start
                    || (left_span.end == right_span.start
                        && !left_span.is_empty()
                        && !right_span.is_empty())
                {
                    left_index += 1;
                } else {
                    right_index += 1;
                }
            }
            left_alignments.sort_unstable_by_key(|(_, position)| *position);
            right_alignments.sort_unstable_by_key(|(_, position)| *position);
            let mut left_index = 0;
            let mut right_index = 0;
            while let (
                Some((left_operation, left_position)),
                Some((right_operation, right_position)),
            ) = (
                left_alignments.get(left_index),
                right_alignments.get(right_index),
            ) {
                match left_position.cmp(right_position) {
                    Ordering::Equal => {
                        record_durable_conflict(
                            &mut left_conflicting,
                            &mut right_conflicting,
                            &mut conflicts,
                            *left_operation,
                            *right_operation,
                            durable_effect_conflict(
                                durable_operation_ref(left, *left_operation)?,
                                durable_operation_ref(right, *right_operation)?,
                                "alignment",
                            ),
                        )?;
                        left_index += 1;
                        right_index += 1;
                    },
                    Ordering::Less => left_index += 1,
                    Ordering::Greater => right_index += 1,
                }
            }
        },
        Some(DurableDomain::Destination) => {
            let mut left_targets = Vec::new();
            let mut right_targets = Vec::new();
            left_targets.try_reserve(left.len()).map_err(|_error| {
                CompositionError::Durable("conflict allocation failed".to_string())
            })?;
            right_targets.try_reserve(right.len()).map_err(|_error| {
                CompositionError::Durable("conflict allocation failed".to_string())
            })?;
            for (index, operation) in left.iter().enumerate() {
                left_targets.push((index, durable_target_key(operation)?));
            }
            for (index, operation) in right.iter().enumerate() {
                right_targets.push((index, durable_target_key(operation)?));
            }
            left_targets.sort_unstable_by(|(_, left), (_, right)| left.cmp(right));
            right_targets.sort_unstable_by(|(_, left), (_, right)| left.cmp(right));
            let mut left_index = 0;
            let mut right_index = 0;
            while let (Some((left_operation, left_target)), Some((right_operation, right_target))) =
                (left_targets.get(left_index), right_targets.get(right_index))
            {
                match left_target.cmp(right_target) {
                    Ordering::Equal => {
                        record_durable_conflict(
                            &mut left_conflicting,
                            &mut right_conflicting,
                            &mut conflicts,
                            *left_operation,
                            *right_operation,
                            durable_effect_conflict(
                                durable_operation_ref(left, *left_operation)?,
                                durable_operation_ref(right, *right_operation)?,
                                "destination",
                            ),
                        )?;
                        left_index += 1;
                        right_index += 1;
                    },
                    Ordering::Less => left_index += 1,
                    Ordering::Greater => right_index += 1,
                }
            }
        },
        None => {},
    }
    Ok((left_conflicting, right_conflicting, conflicts))
}

fn record_durable_conflict(
    left_conflicting: &mut [bool],
    right_conflicting: &mut [bool],
    conflicts: &mut Vec<CompositionConflict>,
    left_index: usize,
    right_index: usize,
    conflict: CompositionConflict,
) -> Result<(), CompositionError> {
    let left_marker = left_conflicting.get_mut(left_index).ok_or_else(|| {
        CompositionError::Durable("left conflict index is out of bounds".to_string())
    })?;
    let right_marker = right_conflicting.get_mut(right_index).ok_or_else(|| {
        CompositionError::Durable("right conflict index is out of bounds".to_string())
    })?;
    *left_marker = true;
    *right_marker = true;
    conflicts
        .try_reserve(1)
        .map_err(|_error| CompositionError::Durable("conflict allocation failed".to_string()))?;
    conflicts.push(conflict);
    Ok(())
}

fn durable_operation_ref(
    operations: &[DurableOperation],
    index: usize,
) -> Result<&DurableOperation, CompositionError> {
    operations
        .get(index)
        .ok_or_else(|| CompositionError::Durable("operation index is out of bounds".to_string()))
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum DurableDomain {
    Ordinary,
    Destination,
}

#[derive(Debug, Clone, PartialEq, Eq, PartialOrd, Ord)]
enum DurableTargetKey {
    Table {
        root: (usize, usize, usize),
        nested: Vec<(usize, usize, usize)>,
    },
    Header {
        section: usize,
        kind: u8,
        paragraph: usize,
    },
    Indexed {
        kind: u8,
        index: usize,
    },
}

impl DurableDomain {
    const fn name(self) -> &'static str {
        match self {
            Self::Ordinary => "ordinary",
            Self::Destination => "destination",
        }
    }
}

fn durable_domain(operations: &[DurableOperation]) -> Result<DurableDomain, CompositionError> {
    let mut domain = None;
    for operation in operations {
        let current = match operation.op.as_str() {
            "body-text.replace"
            | "paragraph-alignment.set"
            | "character-bold.set"
            | "character-italic.set"
            | "character-underline.set"
            | "character-font-size.set"
            | "character-strike.set"
            | "character-double-strike.set"
            | "character-hidden.set"
            | "character-small-caps.set"
            | "character-all-caps.set" => DurableDomain::Ordinary,
            "table-cell-text.replace"
            | "header-footer-text.replace"
            | "annotation-text.replace"
            | "note-text.replace"
            | "shape-text.replace" => DurableDomain::Destination,
            "paragraph-layout.patch"
            | "paragraph.insert"
            | "paragraph.remove"
            | "paragraph.move"
            | "paragraph.split"
            | "paragraph.merge" => {
                return Err(CompositionError::Durable(
                    "paragraph lifecycle/layout durable operations are not composable".to_string(),
                ));
            },
            "picture-payload.replace" | "picture.remove" | "picture.insert-exact" => {
                return Err(CompositionError::Durable(
                    "picture durable operations are not composable".to_string(),
                ));
            },
            "field.transfer"
            | "nested-table.transfer"
            | "list.transfer"
            | "style.transfer"
            | "object.transfer"
            | "shape.transfer" => {
                return Err(CompositionError::Durable(
                    "root-transfer durable operations are not composable".to_string(),
                ));
            },
            _ => {
                return Err(CompositionError::Durable(
                    "unknown durable operation vocabulary".to_string(),
                ));
            },
        };
        if domain.is_some_and(|existing| existing != current) {
            return Err(CompositionError::Durable(
                "mixed durable publication domains are not composable".to_string(),
            ));
        }
        domain = Some(current);
    }
    domain.ok_or_else(|| CompositionError::Durable("empty durable branch".to_string()))
}

fn validate_durable_operation_set(operations: &[DurableOperation]) -> Result<(), CompositionError> {
    if operations.is_empty() {
        return Ok(());
    }
    let domain = durable_domain(operations)?;
    match domain {
        DurableDomain::Ordinary => {
            let mut intervals = Vec::new();
            let mut alignments = Vec::new();
            intervals.try_reserve(operations.len()).map_err(|_error| {
                CompositionError::Durable("operation allocation failed".to_string())
            })?;
            alignments.try_reserve(operations.len()).map_err(|_error| {
                CompositionError::Durable("operation allocation failed".to_string())
            })?;
            for (index, operation) in operations.iter().enumerate() {
                if operation.op == "paragraph-alignment.set" {
                    alignments.push((
                        index,
                        super::parse_paragraph_target(&operation.target)
                            .map_err(CompositionError::Edit)?,
                    ));
                } else {
                    intervals.push((
                        index,
                        durable_text_span(operation).ok_or_else(|| {
                            CompositionError::Durable(
                                "ordinary durable operation has no text span".to_string(),
                            )
                        })?,
                    ));
                }
            }
            intervals.sort_unstable_by_key(|(_, span)| (span.start, span.end));
            for pair in intervals.windows(2) {
                let [left, right] = pair else {
                    continue;
                };
                if spans_overlap(left.1, right.1) {
                    return Err(CompositionError::Conflicts(ConflictSet::new([
                        durable_operation_conflict(
                            durable_operation_ref(operations, left.0)?,
                            durable_operation_ref(operations, right.0)?,
                        )
                        .ok_or_else(|| {
                            CompositionError::Durable(
                                "ordinary durable overlap was not classified".to_string(),
                            )
                        })?,
                    ])));
                }
            }
            alignments.sort_unstable_by_key(|(_, position)| *position);
            for pair in alignments.windows(2) {
                let [left, right] = pair else {
                    continue;
                };
                if left.1 == right.1 {
                    return Err(CompositionError::Conflicts(ConflictSet::new([
                        durable_operation_conflict(
                            durable_operation_ref(operations, left.0)?,
                            durable_operation_ref(operations, right.0)?,
                        )
                        .ok_or_else(|| {
                            CompositionError::Durable(
                                "duplicate alignment was not classified".to_string(),
                            )
                        })?,
                    ])));
                }
            }
        },
        DurableDomain::Destination => {
            let mut targets = Vec::new();
            targets.try_reserve(operations.len()).map_err(|_error| {
                CompositionError::Durable("operation allocation failed".to_string())
            })?;
            for (index, operation) in operations.iter().enumerate() {
                targets.push((index, durable_target_key(operation)?));
            }
            targets.sort_unstable_by(|(_, left), (_, right)| left.cmp(right));
            for pair in targets.windows(2) {
                let [left, right] = pair else {
                    continue;
                };
                if left.1 == right.1 {
                    return Err(CompositionError::Conflicts(ConflictSet::new([
                        durable_operation_conflict(
                            durable_operation_ref(operations, left.0)?,
                            durable_operation_ref(operations, right.0)?,
                        )
                        .ok_or_else(|| {
                            CompositionError::Durable(
                                "duplicate destination was not classified".to_string(),
                            )
                        })?,
                    ])));
                }
            }
        },
    }
    Ok(())
}

fn durable_target_key(operation: &DurableOperation) -> Result<DurableTargetKey, CompositionError> {
    match operation.op.as_str() {
        "table-cell-text.replace" => {
            let path = super::parse_table_cell_target(&operation.target)
                .map_err(CompositionError::Edit)?;
            let mut nested = Vec::new();
            nested.try_reserve(path.nested.len()).map_err(|_error| {
                CompositionError::Durable("target allocation failed".to_string())
            })?;
            nested.extend(path.nested.iter().map(|coordinate| {
                (
                    coordinate.table_index,
                    coordinate.row_index,
                    coordinate.cell_index,
                )
            }));
            Ok(DurableTargetKey::Table {
                root: (
                    path.root.table_index,
                    path.root.row_index,
                    path.root.cell_index,
                ),
                nested,
            })
        },
        "header-footer-text.replace" => {
            let target = super::parse_header_footer_target(&operation.target)
                .map_err(CompositionError::Edit)?;
            Ok(DurableTargetKey::Header {
                section: target.section(),
                kind: header_footer_kind_key(target.kind()),
                paragraph: target.paragraph(),
            })
        },
        "annotation-text.replace" => Ok(DurableTargetKey::Indexed {
            kind: 0,
            index: super::parse_annotation_target(&operation.target)
                .map_err(CompositionError::Edit)?,
        }),
        "note-text.replace" => Ok(DurableTargetKey::Indexed {
            kind: 1,
            index: super::parse_note_target(&operation.target).map_err(CompositionError::Edit)?,
        }),
        "shape-text.replace" => Ok(DurableTargetKey::Indexed {
            kind: 2,
            index: super::parse_shape_target(&operation.target).map_err(CompositionError::Edit)?,
        }),
        _ => Err(CompositionError::Durable(
            "invalid durable destination target".to_string(),
        )),
    }
}

const fn header_footer_kind_key(kind: super::HeaderFooterType) -> u8 {
    match kind {
        super::HeaderFooterType::Header => 0,
        super::HeaderFooterType::Footer => 1,
        super::HeaderFooterType::HeaderFirst => 2,
        super::HeaderFooterType::FooterFirst => 3,
        super::HeaderFooterType::HeaderLeft => 4,
        super::HeaderFooterType::FooterLeft => 5,
        super::HeaderFooterType::HeaderRight => 6,
        super::HeaderFooterType::FooterRight => 7,
    }
}

fn validate_durable_forward_limit(
    limits: core::PatchLimits,
    operations: &[DurableOperation],
) -> Result<(), CompositionError> {
    let mut bounded_operations = Vec::new();
    bounded_operations
        .try_reserve(operations.len())
        .map_err(|_error| CompositionError::Durable("operation allocation failed".to_string()))?;
    bounded_operations.extend(operations.iter().cloned());
    core::Patch::<core::ForwardOnly>::new(
        limits,
        "litchi-rtf",
        bounded_operations,
        core::BlobBundle::new(limits.blobs()),
    )
    .map(|_patch| ())
    .map_err(|error| CompositionError::Durable(error.to_string()))
}

fn durable_merge_limit_ok(
    source: &Snapshot,
    limits: core::PatchLimits,
    automatic: &[DurableOperation],
    selected: Option<&[DurableOperation]>,
) -> Result<(), CompositionError> {
    let selected_len = selected.map_or(0, <[DurableOperation]>::len);
    let count = automatic
        .len()
        .checked_add(selected_len)
        .ok_or_else(|| CompositionError::Durable("operation count overflow".to_string()))?;
    let mut operations = Vec::new();
    operations
        .try_reserve(count)
        .map_err(|_error| CompositionError::Durable("operation allocation failed".to_string()))?;
    operations.extend(automatic.iter().cloned());
    if let Some(selected) = selected {
        operations.extend(selected.iter().cloned());
    }
    validate_durable_forward_limit(limits, &operations)?;
    preflight_durable_reversible(source, limits, &operations)
}

fn preflight_durable_reversible(
    source: &Snapshot,
    limits: core::PatchLimits,
    operations: &[DurableOperation],
) -> Result<(), CompositionError> {
    let commit = commit_durable_operations(source, operations)?;
    let patch = commit
        .patch()
        .to_durable(limits)
        .map_err(|error| CompositionError::Durable(error.to_string()))?;
    patch
        .to_deterministic_json()
        .map(|_json| ())
        .map_err(|error| CompositionError::Durable(error.to_string()))
}

fn durable_domain_conflict(left: DurableDomain, right: DurableDomain) -> CompositionConflict {
    CompositionConflict::PublicationDomain {
        left: left.name().to_string(),
        right: right.name().to_string(),
    }
}

fn validate_durable_branch(
    source: &Snapshot,
    limits: core::PatchLimits,
    patch: &DurablePatch,
) -> Result<Vec<DurableOperation>, CompositionError> {
    if patch.format() != "litchi-rtf" {
        return Err(CompositionError::Durable(
            "durable branch uses a different format".to_string(),
        ));
    }
    if patch.limits() != limits {
        return Err(CompositionError::Durable(
            "durable branch uses different patch limits".to_string(),
        ));
    }
    if !patch.blobs().is_empty() || !patch.inverse().blobs().is_empty() {
        return Err(CompositionError::Durable(
            "durable RTF composition does not accept blobs".to_string(),
        ));
    }
    if patch.operations().is_empty() {
        return Ok(Vec::new());
    }
    let _domain = durable_domain(patch.operations())?;
    validate_durable_operation_set(patch.operations())?;
    let source_bytes = source.source_bytes().ok_or_else(|| {
        CompositionError::Durable("base snapshot has no exact RTF bytes".to_string())
    })?;
    let source_hash = core::BlobId::of(source_bytes).as_hex();
    for operation in patch.operations() {
        let expected = operation
            .preconditions
            .get("artifact_sha256")
            .and_then(Value::as_str)
            .ok_or_else(|| {
                CompositionError::Durable("missing branch artifact digest".to_string())
            })?;
        if expected != source_hash {
            return Err(CompositionError::Edit(Error::PatchConflict));
        }
        validate_durable_operation_shape(source, operation)?;
    }
    let target = source
        .apply_durable(patch)
        .map_err(CompositionError::Edit)?;
    let inverse = patch.inverse();
    durable_domain(inverse.operations())?;
    validate_durable_inverse_pairs(patch.operations(), inverse.operations())?;
    for operation in inverse.operations() {
        validate_durable_operation_shape(&target, operation)?;
    }
    let restored = target
        .apply_durable(&inverse)
        .map_err(CompositionError::Edit)?;
    verify_durable_inverse_semantics(source, &restored, patch.operations())?;
    let mut operations = Vec::new();
    operations
        .try_reserve(patch.operations().len())
        .map_err(|_error| CompositionError::Durable("operation allocation failed".to_string()))?;
    operations.extend(patch.operations().iter().cloned());
    Ok(operations)
}

fn validate_durable_inverse_pairs(
    forward_operations: &[DurableOperation],
    inverse: &[DurableOperation],
) -> Result<(), CompositionError> {
    if forward_operations.len() != inverse.len() {
        return Err(CompositionError::Durable(
            "durable branch directions have different operation counts".to_string(),
        ));
    }
    for (forward_index, (forward, inverse)) in
        forward_operations.iter().rev().zip(inverse).enumerate()
    {
        let source_index = forward_operations
            .len()
            .saturating_sub(1)
            .saturating_sub(forward_index);
        if forward.op != inverse.op
            || !durable_inverse_targets_equal(forward, inverse, forward_operations, source_index)
            || !durable_inverse_values_match(forward, inverse)
        {
            return Err(CompositionError::Durable(
                "durable branch inverse does not exactly pair with its forward operation"
                    .to_string(),
            ));
        }
    }
    Ok(())
}

fn durable_inverse_targets_equal(
    forward: &DurableOperation,
    inverse: &DurableOperation,
    all_forward: &[DurableOperation],
    source_index: usize,
) -> bool {
    let target_span_len = match forward.op.as_str() {
        "body-text.replace" => forward.value.as_str().map(str::len),
        "character-bold.set"
        | "character-italic.set"
        | "character-underline.set"
        | "character-font-size.set"
        | "character-strike.set"
        | "character-double-strike.set"
        | "character-hidden.set"
        | "character-small-caps.set"
        | "character-all-caps.set" => None,
        _ => return durable_targets_equal(forward, inverse),
    };
    let Some(source_span) = super::parse_text_target(&forward.target).ok() else {
        return false;
    };
    let target_span_len =
        target_span_len.or_else(|| source_span.end.checked_sub(source_span.start));
    let Some(target_span_len) = target_span_len else {
        return false;
    };
    let mut projected_start = source_span.start;
    for (index, operation) in all_forward.iter().enumerate() {
        if index == source_index || operation.op != "body-text.replace" {
            continue;
        }
        let Some(other_span) = super::parse_text_target(&operation.target).ok() else {
            return false;
        };
        if other_span.end > source_span.start {
            continue;
        }
        let Some(other_len) = operation.value.as_str().map(str::len) else {
            return false;
        };
        let original_len = other_span.end.saturating_sub(other_span.start);
        if other_len >= original_len {
            projected_start = projected_start.saturating_add(other_len - original_len);
        } else {
            projected_start = projected_start.saturating_sub(original_len - other_len);
        }
    }
    let expected_span = super::TextSpan {
        start: projected_start,
        end: projected_start.saturating_add(target_span_len),
    };
    super::parse_text_target(&inverse.target).ok() == Some(expected_span)
}

fn durable_inverse_values_match(forward: &DurableOperation, inverse: &DurableOperation) -> bool {
    let key = match forward.op.as_str() {
        "body-text.replace"
        | "table-cell-text.replace"
        | "header-footer-text.replace"
        | "annotation-text.replace"
        | "note-text.replace"
        | "shape-text.replace" => "text",
        "paragraph-alignment.set" => "alignment",
        "character-bold.set" => "bold",
        "character-italic.set" => "italic",
        "character-underline.set" => "underline",
        "character-font-size.set" => "font_size_half_points",
        "character-strike.set" => "strike",
        "character-double-strike.set" => "double_strike",
        "character-hidden.set" => "hidden",
        "character-small-caps.set" => "small_caps",
        "character-all-caps.set" => "all_caps",
        _ => return false,
    };
    inverse.preconditions.get(key) == Some(&forward.value)
        && forward.preconditions.get(key) == Some(&inverse.value)
}

fn durable_targets_equal(left: &DurableOperation, right: &DurableOperation) -> bool {
    match (left.op.as_str(), right.op.as_str()) {
        ("body-text.replace", "body-text.replace")
        | ("character-bold.set", "character-bold.set")
        | ("character-italic.set", "character-italic.set")
        | ("character-underline.set", "character-underline.set")
        | ("character-font-size.set", "character-font-size.set")
        | ("character-strike.set", "character-strike.set")
        | ("character-double-strike.set", "character-double-strike.set")
        | ("character-hidden.set", "character-hidden.set")
        | ("character-small-caps.set", "character-small-caps.set")
        | ("character-all-caps.set", "character-all-caps.set") => {
            super::parse_text_target(&left.target).ok()
                == super::parse_text_target(&right.target).ok()
        },
        ("paragraph-alignment.set", "paragraph-alignment.set") => {
            super::parse_paragraph_target(&left.target).ok()
                == super::parse_paragraph_target(&right.target).ok()
        },
        ("table-cell-text.replace", "table-cell-text.replace") => {
            super::parse_table_cell_target(&left.target).ok()
                == super::parse_table_cell_target(&right.target).ok()
        },
        ("header-footer-text.replace", "header-footer-text.replace") => {
            super::parse_header_footer_target(&left.target).ok()
                == super::parse_header_footer_target(&right.target).ok()
        },
        ("annotation-text.replace", "annotation-text.replace") => {
            super::parse_annotation_target(&left.target).ok()
                == super::parse_annotation_target(&right.target).ok()
        },
        ("note-text.replace", "note-text.replace") => {
            super::parse_note_target(&left.target).ok()
                == super::parse_note_target(&right.target).ok()
        },
        ("shape-text.replace", "shape-text.replace") => {
            super::parse_shape_target(&left.target).ok()
                == super::parse_shape_target(&right.target).ok()
        },
        _ => false,
    }
}

fn verify_durable_inverse_semantics(
    source: &Snapshot,
    restored: &Snapshot,
    operations: &[DurableOperation],
) -> Result<(), CompositionError> {
    // Character/property rewrites are intentionally semantic: the RTF writer
    // may normalize control-word placement while retaining the same body.
    // Authentication remains exact because both directions have already been
    // checked against their artifact digests by `apply_durable`.
    if source.text() != restored.text() {
        return Err(CompositionError::Durable(
            "durable branch inverse did not restore body text".to_string(),
        ));
    }
    verify_durable_character_properties(source, restored)?;
    if operations
        .iter()
        .any(|operation| operation.op == "paragraph-alignment.set")
        && super::source_alignments(source) != super::source_alignments(restored)
    {
        return Err(CompositionError::Durable(
            "durable branch inverse did not restore paragraph alignment".to_string(),
        ));
    }
    for operation in operations {
        match operation.op.as_str() {
            "character-bold.set" => {
                let span =
                    super::parse_text_target(&operation.target).map_err(CompositionError::Edit)?;
                let expected = operation
                    .preconditions
                    .get("bold")
                    .and_then(Value::as_bool)
                    .ok_or_else(|| {
                        CompositionError::Durable(
                            "durable branch is missing bold state".to_string(),
                        )
                    })?;
                if super::bold_for_span(restored, span).map_err(CompositionError::Edit)? != expected
                {
                    return Err(CompositionError::Durable(
                        "durable branch inverse did not restore bold state".to_string(),
                    ));
                }
            },
            "character-italic.set" => {
                let span =
                    super::parse_text_target(&operation.target).map_err(CompositionError::Edit)?;
                let expected = operation
                    .preconditions
                    .get("italic")
                    .and_then(Value::as_bool)
                    .ok_or_else(|| {
                        CompositionError::Durable(
                            "durable branch is missing italic state".to_string(),
                        )
                    })?;
                if super::italic_for_span(restored, span).map_err(CompositionError::Edit)?
                    != expected
                {
                    return Err(CompositionError::Durable(
                        "durable branch inverse did not restore italic state".to_string(),
                    ));
                }
            },
            "character-underline.set" => {
                let span =
                    super::parse_text_target(&operation.target).map_err(CompositionError::Edit)?;
                let expected = operation
                    .preconditions
                    .get("underline")
                    .and_then(Value::as_str)
                    .ok_or_else(|| {
                        CompositionError::Durable(
                            "durable branch is missing underline state".to_string(),
                        )
                    })?;
                if super::underline_name(
                    super::underline_for_span(restored, span).map_err(CompositionError::Edit)?,
                ) != expected
                {
                    return Err(CompositionError::Durable(
                        "durable branch inverse did not restore underline state".to_string(),
                    ));
                }
            },
            "character-strike.set" => {
                let span =
                    super::parse_text_target(&operation.target).map_err(CompositionError::Edit)?;
                let expected = operation
                    .preconditions
                    .get("strike")
                    .and_then(Value::as_bool)
                    .ok_or_else(|| {
                        CompositionError::Durable(
                            "durable branch is missing strike state".to_string(),
                        )
                    })?;
                if super::strike_for_span(restored, span).map_err(CompositionError::Edit)?
                    != expected
                {
                    return Err(CompositionError::Durable(
                        "durable branch inverse did not restore strike state".to_string(),
                    ));
                }
            },
            "character-double-strike.set" => {
                let span =
                    super::parse_text_target(&operation.target).map_err(CompositionError::Edit)?;
                let expected = operation
                    .preconditions
                    .get("double_strike")
                    .and_then(Value::as_bool)
                    .ok_or_else(|| {
                        CompositionError::Durable(
                            "durable branch is missing double-strike state".to_string(),
                        )
                    })?;
                if super::double_strike_for_span(restored, span).map_err(CompositionError::Edit)?
                    != expected
                {
                    return Err(CompositionError::Durable(
                        "durable branch inverse did not restore double-strike state".to_string(),
                    ));
                }
            },
            "character-hidden.set" => {
                let span =
                    super::parse_text_target(&operation.target).map_err(CompositionError::Edit)?;
                let expected = operation
                    .preconditions
                    .get("hidden")
                    .and_then(Value::as_bool)
                    .ok_or_else(|| {
                        CompositionError::Durable(
                            "durable branch is missing hidden state".to_string(),
                        )
                    })?;
                if super::hidden_for_span(restored, span).map_err(CompositionError::Edit)?
                    != expected
                {
                    return Err(CompositionError::Durable(
                        "durable branch inverse did not restore hidden state".to_string(),
                    ));
                }
            },
            "character-small-caps.set" => {
                let span =
                    super::parse_text_target(&operation.target).map_err(CompositionError::Edit)?;
                let expected = operation
                    .preconditions
                    .get("small_caps")
                    .and_then(Value::as_bool)
                    .ok_or_else(|| {
                        CompositionError::Durable(
                            "durable branch is missing small-caps state".to_string(),
                        )
                    })?;
                if super::small_caps_for_span(restored, span).map_err(CompositionError::Edit)?
                    != expected
                {
                    return Err(CompositionError::Durable(
                        "durable branch inverse did not restore small-caps state".to_string(),
                    ));
                }
            },
            "character-all-caps.set" => {
                let span =
                    super::parse_text_target(&operation.target).map_err(CompositionError::Edit)?;
                let expected = operation
                    .preconditions
                    .get("all_caps")
                    .and_then(Value::as_bool)
                    .ok_or_else(|| {
                        CompositionError::Durable(
                            "durable branch is missing all-caps state".to_string(),
                        )
                    })?;
                if super::all_caps_for_span(restored, span).map_err(CompositionError::Edit)?
                    != expected
                {
                    return Err(CompositionError::Durable(
                        "durable branch inverse did not restore all-caps state".to_string(),
                    ));
                }
            },
            "character-font-size.set" => {
                let span =
                    super::parse_text_target(&operation.target).map_err(CompositionError::Edit)?;
                let expected = operation
                    .preconditions
                    .get("font_size_half_points")
                    .and_then(Value::as_u64)
                    .and_then(|value| u16::try_from(value).ok())
                    .and_then(std::num::NonZeroU16::new)
                    .ok_or_else(|| {
                        CompositionError::Durable(
                            "durable branch is missing font-size state".to_string(),
                        )
                    })?;
                if super::font_size_for_span(restored, span).map_err(CompositionError::Edit)?
                    != expected
                {
                    return Err(CompositionError::Durable(
                        "durable branch inverse did not restore font-size state".to_string(),
                    ));
                }
            },
            "table-cell-text.replace"
            | "header-footer-text.replace"
            | "annotation-text.replace"
            | "note-text.replace"
            | "shape-text.replace" => {
                let source_text =
                    durable_destination_text(source, operation).map_err(CompositionError::Edit)?;
                let restored_text = durable_destination_text(restored, operation)
                    .map_err(CompositionError::Edit)?;
                if source_text != restored_text {
                    return Err(CompositionError::Durable(
                        "durable branch inverse did not restore destination text".to_string(),
                    ));
                }
            },
            "body-text.replace" | "paragraph-alignment.set" => {},
            _ => {
                return Err(CompositionError::Durable(
                    "unsupported operation escaped durable inverse verification".to_string(),
                ));
            },
        }
    }
    Ok(())
}

fn verify_durable_character_properties(
    source: &Snapshot,
    restored: &Snapshot,
) -> Result<(), CompositionError> {
    let mut body_position = 0usize;
    for paragraph in source.body().paragraphs() {
        let mut run_position = body_position;
        for run in paragraph.runs() {
            let run_end = run_position.saturating_add(run.text().len());
            if run_position < run_end {
                let span = super::TextSpan {
                    start: run_position,
                    end: run_end,
                };
                if super::bold_for_span(restored, span).map_err(CompositionError::Edit)?
                    != run.format().bold()
                {
                    return Err(CompositionError::Durable(
                        "durable branch inverse did not restore body bold state".to_string(),
                    ));
                }
                if super::italic_for_span(restored, span).map_err(CompositionError::Edit)?
                    != run.format().italic()
                {
                    return Err(CompositionError::Durable(
                        "durable branch inverse did not restore body italic state".to_string(),
                    ));
                }
                if super::underline_for_span(restored, span).map_err(CompositionError::Edit)?
                    != run.format().underline()
                {
                    return Err(CompositionError::Durable(
                        "durable branch inverse did not restore body underline state".to_string(),
                    ));
                }
                if super::strike_for_span(restored, span).map_err(CompositionError::Edit)?
                    != run.format().raw().strike
                {
                    return Err(CompositionError::Durable(
                        "durable branch inverse did not restore body strike state".to_string(),
                    ));
                }
                if super::formatting_for_span(restored, span).map_err(CompositionError::Edit)?
                    != *run.format().raw()
                {
                    return Err(CompositionError::Durable(
                        "durable branch inverse did not restore full character formatting"
                            .to_string(),
                    ));
                }
            }
            run_position = run_end;
        }
        body_position = body_position
            .saturating_add(paragraph.len())
            .saturating_add(1);
    }
    Ok(())
}

fn durable_destination_text<'a>(
    source: &'a Snapshot,
    operation: &DurableOperation,
) -> Result<&'a str, Error> {
    match operation.op.as_str() {
        "table-cell-text.replace" => {
            let path = super::parse_table_cell_target(&operation.target)?;
            Ok(super::table_cell(source, &path)?.text())
        },
        "header-footer-text.replace" => {
            let target = super::parse_header_footer_target(&operation.target)?;
            Ok(super::header_footer(source, target)?
                .paragraphs
                .get(target.paragraph)
                .ok_or(Error::DestinationOutOfRange("header/footer paragraph"))?
                .text
                .as_ref())
        },
        "annotation-text.replace" => {
            let index = super::parse_annotation_target(&operation.target)?;
            Ok(super::annotation(source, index)?.text.as_ref())
        },
        "note-text.replace" => {
            let index = super::parse_note_target(&operation.target)?;
            Ok(super::note(source, index)?.content.as_ref())
        },
        "shape-text.replace" => {
            let index = super::parse_shape_target(&operation.target)?;
            Ok(super::shape(source, index)?.text.as_ref())
        },
        _ => Err(Error::DurablePatch(
            "invalid durable destination operation".to_string(),
        )),
    }
}

fn validate_durable_operation_shape(
    source: &Snapshot,
    operation: &DurableOperation,
) -> Result<(), CompositionError> {
    let expected_preconditions = match operation.op.as_str() {
        "body-text.replace"
        | "paragraph-alignment.set"
        | "character-bold.set"
        | "character-italic.set"
        | "character-underline.set"
        | "character-font-size.set"
        | "character-strike.set"
        | "character-double-strike.set"
        | "character-hidden.set"
        | "character-small-caps.set"
        | "character-all-caps.set"
        | "table-cell-text.replace"
        | "header-footer-text.replace"
        | "annotation-text.replace"
        | "note-text.replace"
        | "shape-text.replace" => 2,
        _ => return Ok(()),
    };
    if operation.preconditions.len() != expected_preconditions {
        return Err(CompositionError::Durable(
            "durable branch operation has an invalid precondition count".to_string(),
        ));
    }
    match operation.op.as_str() {
        "body-text.replace" => {
            let span =
                super::parse_text_target(&operation.target).map_err(CompositionError::Edit)?;
            super::validate_span(source.text(), span).map_err(CompositionError::Edit)?;
            let before = source.text().get(span.start..span.end).ok_or_else(|| {
                CompositionError::Durable("body target is outside the base".to_string())
            })?;
            if before.contains('\n')
                || operation
                    .value
                    .as_str()
                    .is_none_or(|replacement| replacement.contains('\n'))
            {
                return Err(CompositionError::Durable(
                    "body durable replacements cannot contain paragraph breaks".to_string(),
                ));
            }
        },
        "character-bold.set"
        | "character-italic.set"
        | "character-underline.set"
        | "character-font-size.set"
        | "character-strike.set"
        | "character-double-strike.set"
        | "character-hidden.set"
        | "character-small-caps.set"
        | "character-all-caps.set" => {
            let span =
                super::parse_text_target(&operation.target).map_err(CompositionError::Edit)?;
            super::validate_span(source.text(), span).map_err(CompositionError::Edit)?;
            if span.is_empty()
                || source
                    .text()
                    .get(span.start..span.end)
                    .is_none_or(|text| text.contains('\n'))
            {
                return Err(CompositionError::Durable(
                    "character durable ranges must be non-empty and paragraph-local".to_string(),
                ));
            }
        },
        "paragraph-alignment.set"
        | "table-cell-text.replace"
        | "header-footer-text.replace"
        | "annotation-text.replace"
        | "note-text.replace"
        | "shape-text.replace" => {},
        _ => {},
    }
    Ok(())
}

fn durable_operation_conflict(
    left: &DurableOperation,
    right: &DurableOperation,
) -> Option<CompositionConflict> {
    let left_domain = durable_operation_domain(left)?;
    let right_domain = durable_operation_domain(right)?;
    if left_domain != right_domain {
        return Some(durable_domain_conflict(left_domain, right_domain));
    }
    match left_domain {
        DurableDomain::Destination => {
            if durable_targets_equal(left, right) {
                Some(durable_effect_conflict(left, right, "destination"))
            } else {
                None
            }
        },
        DurableDomain::Ordinary => ordinary_durable_conflict(left, right),
    }
}

fn durable_operation_domain(operation: &DurableOperation) -> Option<DurableDomain> {
    match operation.op.as_str() {
        "body-text.replace"
        | "paragraph-alignment.set"
        | "character-bold.set"
        | "character-italic.set"
        | "character-underline.set"
        | "character-font-size.set"
        | "character-strike.set"
        | "character-double-strike.set"
        | "character-hidden.set"
        | "character-small-caps.set"
        | "character-all-caps.set" => Some(DurableDomain::Ordinary),
        "table-cell-text.replace"
        | "header-footer-text.replace"
        | "annotation-text.replace"
        | "note-text.replace"
        | "shape-text.replace" => Some(DurableDomain::Destination),
        _ => None,
    }
}

fn ordinary_durable_conflict(
    left: &DurableOperation,
    right: &DurableOperation,
) -> Option<CompositionConflict> {
    if left.op == "paragraph-alignment.set" && right.op == "paragraph-alignment.set" {
        return durable_targets_equal(left, right)
            .then(|| durable_effect_conflict(left, right, "alignment"));
    }
    if left.op == "paragraph-alignment.set" || right.op == "paragraph-alignment.set" {
        return None;
    }
    let left_span = durable_text_span(left)?;
    let right_span = durable_text_span(right)?;
    if spans_overlap(left_span, right_span) {
        Some(durable_effect_conflict(left, right, "character-or-text"))
    } else {
        None
    }
}

fn durable_text_span(operation: &DurableOperation) -> Option<super::TextSpan> {
    match operation.op.as_str() {
        "body-text.replace"
        | "character-bold.set"
        | "character-italic.set"
        | "character-underline.set"
        | "character-font-size.set"
        | "character-strike.set"
        | "character-double-strike.set"
        | "character-hidden.set"
        | "character-small-caps.set"
        | "character-all-caps.set" => super::parse_text_target(&operation.target).ok(),
        _ => None,
    }
}

fn durable_effect_conflict(
    left: &DurableOperation,
    right: &DurableOperation,
    effect: &str,
) -> CompositionConflict {
    CompositionConflict::Effect {
        effect: format!("rtf:durable:{effect}"),
        left: format!("{}:{}", left.op, left.target),
        right: format!("{}:{}", right.op, right.target),
    }
}

fn compare_durable_operations(left: &DurableOperation, right: &DurableOperation) -> Ordering {
    left.target
        .cmp(&right.target)
        .then_with(|| left.op.cmp(&right.op))
}

fn compare_conflicts(left: &CompositionConflict, right: &CompositionConflict) -> Ordering {
    match (left, right) {
        (CompositionConflict::DuplicateId(left), CompositionConflict::DuplicateId(right)) => {
            left.cmp(right)
        },
        (
            CompositionConflict::Effect {
                effect: left_effect,
                left: left_left,
                right: left_right,
            },
            CompositionConflict::Effect {
                effect: right_effect,
                left: right_left,
                right: right_right,
            },
        ) => left_effect
            .cmp(right_effect)
            .then_with(|| left_left.cmp(right_left))
            .then_with(|| left_right.cmp(right_right)),
        (
            CompositionConflict::PublicationDomain {
                left: left_left,
                right: left_right,
            },
            CompositionConflict::PublicationDomain {
                left: right_left,
                right: right_right,
            },
        ) => left_left
            .cmp(right_left)
            .then_with(|| left_right.cmp(right_right)),
        (CompositionConflict::Unknown, CompositionConflict::Unknown) => Ordering::Equal,
        (left, right) => conflict_kind(left).cmp(conflict_kind(right)),
    }
}

const fn conflict_kind(conflict: &CompositionConflict) -> &'static str {
    match conflict {
        CompositionConflict::DuplicateId(_) => "duplicate",
        CompositionConflict::Effect { .. } => "effect",
        CompositionConflict::PublicationDomain { .. } => "domain",
        CompositionConflict::Unknown => "unknown",
    }
}

fn commit_durable_operations(
    source: &Snapshot,
    operations: &[DurableOperation],
) -> Result<super::Commit, CompositionError> {
    let operation_limit = operations.len().max(1);
    let mut edit = source.edit_with_limits(Limits::new(operation_limit));
    for operation in operations {
        match operation.op.as_str() {
            "body-text.replace" => {
                let span =
                    super::parse_text_target(&operation.target).map_err(CompositionError::Edit)?;
                let replacement = operation.value.as_str().ok_or_else(|| {
                    CompositionError::Durable("body replacement is not text".to_string())
                })?;
                edit.replace_text(span, replacement)
                    .map_err(CompositionError::Edit)?;
            },
            "paragraph-alignment.set" => {
                let position = super::parse_paragraph_target(&operation.target)
                    .map_err(CompositionError::Edit)?;
                let alignment = operation
                    .value
                    .as_str()
                    .and_then(super::parse_alignment)
                    .ok_or_else(|| {
                        CompositionError::Durable("invalid paragraph alignment".to_string())
                    })?;
                edit.set_paragraph_alignment(position, alignment)
                    .map_err(CompositionError::Edit)?;
            },
            "character-bold.set" => {
                let span =
                    super::parse_text_target(&operation.target).map_err(CompositionError::Edit)?;
                let value = operation.value.as_bool().ok_or_else(|| {
                    CompositionError::Durable("bold value is not Boolean".to_string())
                })?;
                edit.set_text_bold(span, value)
                    .map_err(CompositionError::Edit)?;
            },
            "character-italic.set" => {
                let span =
                    super::parse_text_target(&operation.target).map_err(CompositionError::Edit)?;
                let value = operation.value.as_bool().ok_or_else(|| {
                    CompositionError::Durable("italic value is not Boolean".to_string())
                })?;
                edit.set_text_italic(span, value)
                    .map_err(CompositionError::Edit)?;
            },
            "character-underline.set" => {
                let span =
                    super::parse_text_target(&operation.target).map_err(CompositionError::Edit)?;
                let value = operation
                    .value
                    .as_str()
                    .and_then(super::parse_underline)
                    .ok_or_else(|| {
                        CompositionError::Durable("invalid underline value".to_string())
                    })?;
                edit.set_text_underline(span, value)
                    .map_err(CompositionError::Edit)?;
            },
            "character-font-size.set" => {
                let span =
                    super::parse_text_target(&operation.target).map_err(CompositionError::Edit)?;
                let value = operation
                    .value
                    .as_u64()
                    .and_then(|value| u16::try_from(value).ok())
                    .and_then(std::num::NonZeroU16::new)
                    .ok_or_else(|| {
                        CompositionError::Durable("invalid font-size value".to_string())
                    })?;
                edit.set_text_font_size(span, value)
                    .map_err(CompositionError::Edit)?;
            },
            "character-strike.set" => {
                let span =
                    super::parse_text_target(&operation.target).map_err(CompositionError::Edit)?;
                let value = operation
                    .value
                    .as_bool()
                    .ok_or_else(|| CompositionError::Durable("invalid strike value".to_string()))?;
                edit.set_text_strike(span, value)
                    .map_err(CompositionError::Edit)?;
            },
            "character-double-strike.set" => {
                let span =
                    super::parse_text_target(&operation.target).map_err(CompositionError::Edit)?;
                let value = operation.value.as_bool().ok_or_else(|| {
                    CompositionError::Durable("invalid double-strike value".to_string())
                })?;
                edit.set_text_double_strike(span, value)
                    .map_err(CompositionError::Edit)?;
            },
            "character-hidden.set" => {
                let span =
                    super::parse_text_target(&operation.target).map_err(CompositionError::Edit)?;
                let value = operation
                    .value
                    .as_bool()
                    .ok_or_else(|| CompositionError::Durable("invalid hidden value".to_string()))?;
                edit.set_text_hidden(span, value)
                    .map_err(CompositionError::Edit)?;
            },
            "character-small-caps.set" => {
                let span =
                    super::parse_text_target(&operation.target).map_err(CompositionError::Edit)?;
                let value = operation.value.as_bool().ok_or_else(|| {
                    CompositionError::Durable("invalid small-caps value".to_string())
                })?;
                edit.set_text_small_caps(span, value)
                    .map_err(CompositionError::Edit)?;
            },
            "character-all-caps.set" => {
                let span =
                    super::parse_text_target(&operation.target).map_err(CompositionError::Edit)?;
                let value = operation.value.as_bool().ok_or_else(|| {
                    CompositionError::Durable("invalid all-caps value".to_string())
                })?;
                edit.set_text_all_caps(span, value)
                    .map_err(CompositionError::Edit)?;
            },
            "table-cell-text.replace" => {
                let path = super::parse_table_cell_target(&operation.target)
                    .map_err(CompositionError::Edit)?;
                let value = operation.value.as_str().ok_or_else(|| {
                    CompositionError::Durable("table-cell replacement is not text".to_string())
                })?;
                edit.set_table_cell_text(path, value)
                    .map_err(CompositionError::Edit)?;
            },
            "header-footer-text.replace" => {
                let target = super::parse_header_footer_target(&operation.target)
                    .map_err(CompositionError::Edit)?;
                let value = operation.value.as_str().ok_or_else(|| {
                    CompositionError::Durable("header/footer replacement is not text".to_string())
                })?;
                edit.set_header_footer_text(target, value)
                    .map_err(CompositionError::Edit)?;
            },
            "annotation-text.replace" => {
                let index = super::parse_annotation_target(&operation.target)
                    .map_err(CompositionError::Edit)?;
                let value = operation.value.as_str().ok_or_else(|| {
                    CompositionError::Durable("annotation replacement is not text".to_string())
                })?;
                edit.set_annotation_text(index, value)
                    .map_err(CompositionError::Edit)?;
            },
            "note-text.replace" => {
                let index =
                    super::parse_note_target(&operation.target).map_err(CompositionError::Edit)?;
                let value = operation.value.as_str().ok_or_else(|| {
                    CompositionError::Durable("note replacement is not text".to_string())
                })?;
                edit.set_note_text(index, value)
                    .map_err(CompositionError::Edit)?;
            },
            "shape-text.replace" => {
                let index =
                    super::parse_shape_target(&operation.target).map_err(CompositionError::Edit)?;
                let value = operation.value.as_str().ok_or_else(|| {
                    CompositionError::Durable("shape replacement is not text".to_string())
                })?;
                edit.set_shape_text(index, value)
                    .map_err(CompositionError::Edit)?;
            },
            _ => {
                return Err(CompositionError::Durable(
                    "unsupported operation escaped durable admission".to_string(),
                ));
            },
        }
    }
    edit.commit().map_err(CompositionError::Edit)
}
