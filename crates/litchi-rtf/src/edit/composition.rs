//! Core-backed composition and non-applying merge plans for RTF edits.

use super::{Edit, Error, Limits, Operation, Snapshot};
use litchi_core::patch as core;
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
        if let Some(existing) = self.inner.sub_edits().find(|existing| {
            different_publication_domains(existing.payload(), incoming.inner.payload())
        }) {
            return Err(CompositionError::Conflicts(publication_conflict(
                existing.id(),
                incoming.inner.id(),
            )));
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
        if let Some((left_id, right_id)) = left.inner.sub_edits().find_map(|left_edit| {
            right
                .inner
                .sub_edits()
                .find(|right_edit| {
                    different_publication_domains(left_edit.payload(), right_edit.payload())
                })
                .map(|right_edit| (left_edit.id(), right_edit.id()))
        }) {
            return Err(CompositionError::Conflicts(publication_conflict(
                left_id, right_id,
            )));
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
        reads.push("body:structure".to_string());
        writes.extend(operation.effect_keys());
    }
    (reads, writes)
}

fn different_publication_domains(left: &[Operation], right: &[Operation]) -> bool {
    let left_destination = left.iter().any(Operation::is_destination);
    let right_destination = right.iter().any(Operation::is_destination);
    left_destination != right_destination
}

fn publication_conflict(left: &str, right: &str) -> ConflictSet {
    ConflictSet::new(vec![CompositionConflict::Effect {
        effect: "rtf:publication-domain".to_string(),
        left: left.to_string(),
        right: right.to_string(),
    }])
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
