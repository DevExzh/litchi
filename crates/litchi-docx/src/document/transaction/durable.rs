//! Durable semantic patches, disjoint composition, and bounded history.

use std::collections::{BTreeMap, BTreeSet, VecDeque};
use std::fmt;
use std::sync::Arc;

use litchi_core::Position;
use litchi_core::patch::{
    BlobBundle, BlobId, JoinedSubEdits, MergeChoice, Patch as CorePatch, PatchLimits,
    PatchOperation, Reversible, ReversibleOperation, SubEdit, ThreeWayMergeFailure,
    ThreeWayMergePlan,
};
use serde_json::Value;

#[cfg(test)]
use super::Commit;
use super::{
    CompositionLimits, Edit, Error as TransactionError, GraphChange, HistoryLimits,
    MAX_DOCUMENT_XML_BYTES, MAX_OPERATIONS, ManagedAdmission, Operation, Patch, ReplacedPart,
    Result as TransactionResult, RevisionAction, RevisionKind, Snapshot, SourceIdentity,
    TableCellAddress, TransferGraph, TransferPart, TransferRelationship, XmlStorage,
};

const MAX_TRANSFER_GRAPH_BYTES: usize = 64 * 1024 * 1024;

const FORMAT_NAME: &str = "litchi-docx/document";
const RESTORE_OPERATION: &str = "document.restore";
const RESTORE_TRANSFER_INSERT: &str = "document.restore-transfer.insert";
const RESTORE_TRANSFER_REMOVE: &str = "document.restore-transfer.remove";

#[cfg(test)]
#[path = "composition_effects_tests.rs"]
mod composition_effects_tests;

#[derive(Clone, Debug)]
enum Lineage {
    Unmanaged(Arc<Vec<u8>>),
    Identified {
        xml: Arc<Vec<u8>>,
        identity: Arc<SourceIdentity>,
    },
    Managed {
        storage: XmlStorage,
        identity: Arc<SourceIdentity>,
        _admission: Option<Arc<ManagedAdmission>>,
    },
}

impl PartialEq for Lineage {
    fn eq(&self, other: &Self) -> bool {
        match (self, other) {
            (Self::Unmanaged(left), Self::Unmanaged(right)) => left == right,
            (
                Self::Identified {
                    xml: left_xml,
                    identity: left_identity,
                },
                Self::Identified {
                    xml: right_xml,
                    identity: right_identity,
                },
            ) => left_identity == right_identity && left_xml == right_xml,
            (
                Self::Managed {
                    storage: left_storage,
                    identity: left_identity,
                    ..
                },
                Self::Managed {
                    storage: right_storage,
                    identity: right_identity,
                    ..
                },
            ) => left_identity == right_identity && left_storage.bytes() == right_storage.bytes(),
            _ => false,
        }
    }
}

impl Eq for Lineage {}

struct TransferGraphDecoder<'a> {
    bytes: &'a [u8],
    offset: usize,
}

/// One independently prepared DOCX document edit.
pub struct Prepared {
    inner: SubEdit<Lineage, Vec<Operation>>,
}

impl Prepared {
    /// Stable caller-selected composition identifier.
    #[must_use]
    pub fn identifier(&self) -> &str {
        self.inner.id()
    }

    /// Semantic operations retained by this prepared edit.
    #[must_use]
    pub fn operations(&self) -> &[Operation] {
        self.inner.payload()
    }
}

impl fmt::Debug for Prepared {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Prepared")
            .field("identifier_bytes", &self.identifier().len())
            .field("operations", &self.operations().len())
            .finish()
    }
}

/// Recoverable failure to join one prepared document edit.
pub struct JoinError {
    failure: Box<litchi_core::patch::SubEditJoinFailure>,
    rejected: Box<Prepared>,
}

impl JoinError {
    /// Structured common composition refusal.
    #[must_use]
    pub const fn failure(&self) -> &litchi_core::patch::SubEditJoinFailure {
        &self.failure
    }

    /// Recover the rejected prepared edit.
    #[must_use]
    pub fn into_rejected(self) -> Prepared {
        *self.rejected
    }
}

impl fmt::Debug for JoinError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("JoinError")
            .field("rejected", &self.rejected)
            .finish()
    }
}

/// Bounded deterministic composition of provably disjoint document edits.
pub struct Composition {
    source: Snapshot,
    joined: JoinedSubEdits<Lineage, Vec<Operation>>,
    artifact_fingerprint: Option<[u8; 32]>,
}

impl Composition {
    /// Borrow the exact immutable source used by this composition.
    #[must_use]
    pub const fn source(&self) -> &Snapshot {
        &self.source
    }

    /// Bind this plan to the complete package artifact that supplied its
    /// source document. Package owners use this proof to reject a package
    /// whose untouched parts differ while still accepting an independently
    /// reopened equal-content artifact.
    pub(crate) fn with_artifact_fingerprint(mut self, fingerprint: [u8; 32]) -> Self {
        self.artifact_fingerprint = Some(fingerprint);
        self
    }

    /// Join work only when its exact source and semantic effects are disjoint.
    ///
    /// # Errors
    ///
    /// Returns a structured common refusal while retaining the rejected edit.
    pub fn join(&mut self, incoming: Prepared) -> Result<&mut Self, JoinError> {
        if let Err(error) = self.joined.join(incoming.inner) {
            let (failure, rejected) = error.into_parts();
            return Err(JoinError {
                failure: Box::new(failure),
                rejected: Box::new(Prepared { inner: rejected }),
            });
        }
        Ok(self)
    }

    /// Number of accepted prepared edits.
    #[must_use]
    pub fn len(&self) -> usize {
        self.joined.len()
    }

    /// Whether no edit has been accepted.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.joined.is_empty()
    }

    /// Consume the plan into its exact source and one stable operation batch.
    ///
    /// Publication is owned by an eager or source-backed package. Keeping this
    /// extraction crate-private prevents a composition from publishing a raw
    /// document snapshot while retaining the one validation/commit pass of the
    /// planner.
    pub(crate) fn into_parts(self) -> (Snapshot, Vec<Operation>, Option<[u8; 32]>) {
        let artifact_fingerprint = self.artifact_fingerprint;
        let operations = self
            .joined
            .into_sub_edits()
            .flat_map(|prepared| prepared.into_payload())
            .collect();
        (self.source, operations, artifact_fingerprint)
    }
}

impl fmt::Debug for Composition {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Composition")
            .field("edits", &self.joined.len())
            .finish_non_exhaustive()
    }
}

/// Recoverable failure to plan two independently composed branches.
pub struct ThreeWayError {
    failure: ThreeWayMergeFailure,
    left: Box<Composition>,
    right: Box<Composition>,
}

impl ThreeWayError {
    /// Structured shared planning failure.
    #[must_use]
    pub const fn failure(&self) -> &ThreeWayMergeFailure {
        &self.failure
    }

    /// Recover both unchanged branches.
    #[must_use]
    pub fn into_branches(self) -> (Composition, Composition) {
        (*self.left, *self.right)
    }
}

impl fmt::Debug for ThreeWayError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("ThreeWayError")
            .finish_non_exhaustive()
    }
}

/// Non-mutating bounded three-way plan for two branches from one exact base.
pub struct ThreeWayPlan {
    source: Snapshot,
    inner: ThreeWayMergePlan<Lineage, Vec<Operation>>,
    artifact_fingerprint: Option<[u8; 32]>,
}

impl ThreeWayPlan {
    /// Deterministically ordered semantic overlaps requiring a choice.
    #[must_use]
    pub const fn conflicts(
        &self,
    ) -> &litchi_core::patch::ConflictSet<litchi_core::patch::SubEditConflict> {
        self.inner.conflicts()
    }

    /// Whether every branch edit is automatically disjoint.
    #[must_use]
    pub fn is_clean(&self) -> bool {
        self.inner.conflicts().is_empty()
    }

    /// Resolve the complete conservative conflict group.
    pub fn resolve(&mut self, choice: MergeChoice) -> &mut Self {
        self.inner.resolve(choice);
        self
    }

    /// Finish planning and retain complete staged work for atomic commit.
    ///
    /// # Errors
    ///
    /// Returns this unchanged plan while conflicts remain unresolved.
    pub fn finish(self) -> Result<Composition, Box<Self>> {
        let Self {
            source,
            inner,
            artifact_fingerprint,
        } = self;
        match inner.finish() {
            Ok(joined) => Ok(Composition {
                source,
                joined,
                artifact_fingerprint,
            }),
            Err(unresolved) => Err(Box::new(Self {
                source,
                inner: *unresolved,
                artifact_fingerprint,
            })),
        }
    }
}

impl fmt::Debug for ThreeWayPlan {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("ThreeWayPlan")
            .field("conflicts", &self.conflicts().len())
            .finish_non_exhaustive()
    }
}

/// Explicit byte-budgeted undo/redo history for document snapshots.
pub struct History {
    inner: litchi_core::patch::History<Snapshot>,
    undo_graph: VecDeque<Option<HistoryGraphTransition>>,
    redo_graph: VecDeque<Option<HistoryGraphTransition>>,
    pending_graph: Option<(HistoryGraphTransition, bool)>,
}

#[derive(Clone)]
pub(crate) struct HistoryGraphTransition {
    pub(crate) graph: Arc<TransferGraph>,
    pub(crate) before_digest: Arc<str>,
    pub(crate) after_digest: Arc<str>,
    pub(crate) forward_insert: bool,
}

impl History {
    /// Start history at one immutable snapshot.
    #[must_use]
    pub fn new(snapshot: Snapshot, limits: HistoryLimits) -> Self {
        Self {
            inner: litchi_core::patch::History::new(snapshot, limits),
            undo_graph: VecDeque::new(),
            redo_graph: VecDeque::new(),
            pending_graph: None,
        }
    }

    /// Current immutable snapshot.
    #[must_use]
    pub const fn current(&self) -> &Snapshot {
        self.inner.current()
    }

    /// Whether one undo transition is available.
    #[must_use]
    pub fn can_undo(&self) -> bool {
        self.inner.can_undo()
    }

    /// Whether one redo transition is available.
    #[must_use]
    pub fn can_redo(&self) -> bool {
        self.inner.can_redo()
    }

    /// Retained transition-byte weight.
    #[must_use]
    pub const fn retained_weight(&self) -> u64 {
        self.inner.retained_weight()
    }

    /// Configured finite retention bounds.
    #[must_use]
    pub const fn limits(&self) -> HistoryLimits {
        self.inner.limits()
    }

    pub(crate) fn ensure_can_record<P>(
        &self,
        commit: &crate::document::Commit<P>,
    ) -> TransactionResult<()> {
        if self.current().is_managed() || commit.patch().target().is_managed() {
            return Err(TransactionError::Document(crate::Error::internal(
                "document.history.ensure_can_record",
                "managed history metadata has no retained reservation ledger",
            )));
        }
        let bytes = history_weight(commit.patch().target(), commit.patch())?;
        if bytes > self.limits().max_weight() {
            return Err(litchi_core::patch::PatchError::HistoryWeight {
                observed: bytes,
                limit: self.limits().max_weight(),
            }
            .into());
        }
        Ok(())
    }

    pub(crate) fn take_graph_transition(&mut self) -> Option<(HistoryGraphTransition, bool)> {
        self.pending_graph.take()
    }

    /// Record a commit using its exact published XML and dependency-graph
    /// size as transition weight.
    ///
    /// # Errors
    ///
    /// Returns a history-weight error without changing history when the
    /// published transition alone exceeds the configured byte budget.
    pub fn record<P>(
        &mut self,
        commit: crate::document::Commit<P>,
    ) -> TransactionResult<Vec<Snapshot>> {
        if self.current().is_managed() || commit.patch().target().is_managed() {
            return Err(TransactionError::Document(crate::Error::internal(
                "document.history.record",
                "managed history metadata has no retained reservation ledger",
            )));
        }
        let weight = history_weight(commit.patch().target(), commit.patch())?;
        let graph = history_graph_transition(commit.patch().operations())?;
        let invalidated_redo = self.redo_graph.len();
        let discarded = self
            .inner
            .record(commit.patch().target().clone(), weight)
            .map_err(TransactionError::from)?;
        self.pending_graph = None;
        self.redo_graph.clear();
        self.undo_graph.push_back(graph);
        let evicted = discarded.len().saturating_sub(invalidated_redo);
        for _ in 0..evicted {
            let _discarded = self.undo_graph.pop_front();
        }
        Ok(discarded)
    }

    /// Record an internal codec transaction for the transaction module's
    /// differential tests. Public callers record the canonical owner commit
    /// through [`Self::record`].
    #[cfg(test)]
    pub(crate) fn record_transaction(
        &mut self,
        commit: Commit,
    ) -> TransactionResult<Vec<Snapshot>> {
        if self.current().is_managed() || commit.snapshot().is_managed() {
            return Err(TransactionError::Document(crate::Error::internal(
                "document.history.record_transaction",
                "managed history metadata has no retained reservation ledger",
            )));
        }
        let weight = history_weight(commit.snapshot(), commit.patch())?;
        let graph = history_graph_transition(commit.patch().operations())?;
        let invalidated_redo = self.redo_graph.len();
        let discarded = self
            .inner
            .record(commit.snapshot, weight)
            .map_err(TransactionError::from)?;
        self.pending_graph = None;
        self.redo_graph.clear();
        self.undo_graph.push_back(graph);
        let evicted = discarded.len().saturating_sub(invalidated_redo);
        for _ in 0..evicted {
            let _discarded = self.undo_graph.pop_front();
        }
        Ok(discarded)
    }

    /// Move one retained transition backward.
    pub fn undo(&mut self) -> bool {
        if !self.inner.undo() {
            return false;
        }
        let Some(graph) = self.undo_graph.pop_back() else {
            let _restored = self.inner.redo();
            return false;
        };
        self.pending_graph = graph.clone().map(|transition| (transition, false));
        self.redo_graph.push_back(graph);
        true
    }

    /// Move one retained transition forward.
    pub fn redo(&mut self) -> bool {
        if !self.inner.redo() {
            return false;
        }
        let Some(graph) = self.redo_graph.pop_back() else {
            let _restored = self.inner.undo();
            return false;
        };
        self.pending_graph = graph.clone().map(|transition| (transition, true));
        self.undo_graph.push_back(graph);
        true
    }
}

impl Snapshot {
    /// Start bounded deterministic composition against this exact source.
    #[must_use]
    pub fn compose(&self, limits: CompositionLimits) -> Composition {
        Composition {
            source: self.clone(),
            joined: JoinedSubEdits::new(self.lineage(), limits),
            artifact_fingerprint: None,
        }
    }

    /// Plan two independently composed branches without applying either one.
    ///
    /// # Errors
    ///
    /// Returns both branches intact when exact lineage or finite bounds differ,
    /// or when the shared planner cannot retain the complete conflict set.
    pub fn plan_three_way(
        &self,
        left: Composition,
        right: Composition,
    ) -> Result<ThreeWayPlan, ThreeWayError> {
        if !self.same_source(&left.source) || !self.same_source(&right.source) {
            return Err(ThreeWayError {
                failure: ThreeWayMergeFailure::DifferentLineage,
                left: Box::new(left),
                right: Box::new(right),
            });
        }
        let artifact_fingerprint = match (left.artifact_fingerprint, right.artifact_fingerprint) {
            (Some(left), Some(right)) if left == right => Some(left),
            (None, None) => None,
            _ => {
                return Err(ThreeWayError {
                    failure: ThreeWayMergeFailure::DifferentLineage,
                    left: Box::new(left),
                    right: Box::new(right),
                });
            },
        };
        let source = self.clone();
        let left_source = left.source.clone();
        let right_source = right.source.clone();
        match ThreeWayMergePlan::new(left.joined, right.joined) {
            Ok(inner) => Ok(ThreeWayPlan {
                source,
                inner,
                artifact_fingerprint,
            }),
            Err(error) => {
                let failure = error.failure().clone();
                let (left_joined, right_joined) = error.into_branches();
                Err(ThreeWayError {
                    failure,
                    left: Box::new(Composition {
                        source: left_source,
                        joined: left_joined,
                        artifact_fingerprint,
                    }),
                    right: Box::new(Composition {
                        source: right_source,
                        joined: right_joined,
                        artifact_fingerprint,
                    }),
                })
            },
        }
    }

    /// Start explicit bounded undo/redo history at this snapshot.
    #[must_use]
    pub fn history(&self, limits: HistoryLimits) -> History {
        History::new(self.clone(), limits)
    }

    /// Apply the supported durable DOCX semantic vocabulary to this exact
    /// source snapshot.
    ///
    /// # Errors
    ///
    /// Returns an error for the wrong format, blobs, malformed operations,
    /// stale artifact hashes, or failed semantic preconditions.
    pub fn apply_durable<Mode>(&self, patch: &CorePatch<Mode>) -> TransactionResult<Self> {
        if self.is_managed() {
            return Err(TransactionError::Document(crate::Error::internal(
                "document.apply_durable",
                "durable replay has no source-owner publication authority",
            )));
        }
        if patch.format() != FORMAT_NAME {
            return Err(invalid_durable("unsupported format"));
        }
        if patch.operations().is_empty() {
            return Err(invalid_durable("empty patch has no source precondition"));
        }
        if patch.operations().len() > MAX_OPERATIONS {
            return Err(TransactionError::Limit {
                kind: crate::document::LimitKind::DurableOperations,
                maximum: MAX_OPERATIONS,
                observed: patch.operations().len(),
            });
        }
        let expected_artifact = BlobId::of(self.xml_bytes()).as_hex();
        for operation in patch.operations() {
            let artifact = operation
                .preconditions
                .get("artifact_sha256")
                .and_then(Value::as_str)
                .ok_or_else(|| invalid_durable("missing artifact hash precondition"))?;
            if artifact != expected_artifact {
                return Err(TransactionError::PatchConflict);
            }
        }

        if patch
            .operations()
            .iter()
            .all(|operation| is_restore_operation(&operation.op))
        {
            return restore_snapshot(patch, self.source_identity());
        }
        if patch
            .operations()
            .iter()
            .any(|operation| is_restore_operation(&operation.op))
        {
            return Err(invalid_durable("invalid semantic blob bundle"));
        }
        validate_semantic_blobs(patch)?;
        let expected_target = common_target_artifact(patch)?;

        let mut edit = self.edit();
        for operation in patch.operations() {
            if operation.op == "document.noop"
                && operation.target == "document"
                && operation.value.is_null()
                && operation.preconditions.len() == 2
            {
                continue;
            }
            let semantic = parse_durable_operation(operation, patch.blobs())?;
            edit.apply_operation(&semantic)?;
        }
        let committed = edit.commit()?.snapshot;
        if BlobId::of(committed.xml_bytes()).as_hex() != expected_target {
            return Err(invalid_durable(
                "semantic replay produced an unexpected target",
            ));
        }
        Ok(committed)
    }

    fn lineage(&self) -> Lineage {
        match &self.xml {
            XmlStorage::Owned(xml) => Lineage::Unmanaged(Arc::clone(xml)),
            XmlStorage::OwnedWithIdentity { xml, identity } => Lineage::Identified {
                xml: Arc::clone(xml),
                identity: Arc::clone(identity),
            },
            XmlStorage::Managed(xml) => Lineage::Managed {
                storage: self.xml.clone(),
                identity: Arc::clone(&xml.identity),
                _admission: self.admission.clone(),
            },
            XmlStorage::Source { identity, .. } => Lineage::Managed {
                storage: self.xml.clone(),
                identity: Arc::clone(identity),
                _admission: self.admission.clone(),
            },
        }
    }
}

impl Edit {
    /// Convert independently staged work into a bounded composable edit.
    ///
    /// # Errors
    ///
    /// Returns a common composition identifier or effect-bound error.
    pub fn prepare(
        self,
        limits: CompositionLimits,
        identifier: impl Into<String>,
    ) -> TransactionResult<Prepared> {
        let (reads, writes) = operation_effects(&self.operations, self.base.paragraph_count());
        let inner = SubEdit::new(
            self.base.lineage(),
            limits,
            identifier,
            reads,
            writes,
            self.operations,
        )?;
        Ok(Prepared { inner })
    }
}

impl Patch {
    /// Convert this exact reversible patch to deterministic common JSON
    /// semantics without serializing native relationship or part identifiers.
    ///
    /// # Errors
    ///
    /// Returns an error when caller-selected bounds cannot retain the semantic
    /// operation vocabulary.
    pub fn to_durable(
        &self,
        limits: PatchLimits,
    ) -> Result<CorePatch<Reversible>, TransactionError> {
        if self.before.is_managed() || self.after.is_managed() {
            return Err(TransactionError::Document(crate::Error::internal(
                "document.patch.to_durable",
                "durable serialization has no managed owner reservation",
            )));
        }
        let forward_artifact = BlobId::of(self.before.xml_bytes()).as_hex();
        let reverse_artifact = BlobId::of(self.after.xml_bytes()).as_hex();
        let mut forward_blobs = BlobBundle::new(limits.blobs());
        let mut reverse_blobs = BlobBundle::new(limits.blobs());
        let source_blob = reverse_blobs.insert(self.before.xml_bytes())?;
        let operations = if self.operations.is_empty() {
            vec![ReversibleOperation::new(
                noop_operation(limits, &forward_artifact, &reverse_artifact)?,
                restore_operation(limits, &reverse_artifact, &source_blob)?,
            )]
        } else {
            self.operations
                .iter()
                .map(|operation| {
                    let inverse = restore_transfer_operation(
                        limits,
                        operation,
                        &reverse_artifact,
                        &source_blob,
                        &mut reverse_blobs,
                    )?
                    .unwrap_or(restore_operation(
                        limits,
                        &reverse_artifact,
                        &source_blob,
                    )?);
                    Ok(ReversibleOperation::new(
                        durable_operation(
                            limits,
                            operation,
                            &forward_artifact,
                            &reverse_artifact,
                            &mut forward_blobs,
                        )?,
                        inverse,
                    ))
                })
                .collect::<Result<Vec<_>, litchi_core::patch::PatchError>>()?
        };
        CorePatch::<Reversible>::new(
            limits,
            FORMAT_NAME,
            operations,
            forward_blobs,
            reverse_blobs,
        )
        .map_err(TransactionError::from)
    }
}

impl<'a> TransferGraphDecoder<'a> {
    fn new(bytes: &'a [u8]) -> TransactionResult<Self> {
        if !bytes.starts_with(b"LDXG1") {
            return Err(invalid_durable("invalid transfer graph header"));
        }
        Ok(Self { bytes, offset: 5 })
    }

    fn count(
        &mut self,
        kind: crate::document::LimitKind,
        maximum: usize,
    ) -> TransactionResult<usize> {
        let count_bytes = self.take(8)?;
        let encoded = u64::from_le_bytes(
            count_bytes
                .try_into()
                .map_err(|_error| invalid_durable("invalid transfer graph count"))?,
        );
        let count = usize::try_from(encoded).map_err(|_error| TransactionError::Limit {
            kind,
            observed: usize::MAX,
            maximum,
        })?;
        if count > maximum {
            return Err(TransactionError::Limit {
                kind,
                maximum,
                observed: count,
            });
        }
        Ok(count)
    }

    fn bytes(&mut self) -> TransactionResult<&'a [u8]> {
        let length = self.count(
            crate::document::LimitKind::FieldBytes,
            MAX_TRANSFER_GRAPH_BYTES,
        )?;
        self.take(length)
    }

    fn string(&mut self) -> TransactionResult<String> {
        std::str::from_utf8(self.bytes()?)
            .map(str::to_owned)
            .map_err(|_error| invalid_durable("transfer graph string is not UTF-8"))
    }

    fn relationship(&mut self) -> TransactionResult<TransferRelationship> {
        let id = self.string()?;
        let relationship_type = self.string()?;
        let target = self.string()?;
        let external = match self.take(1)? {
            [0] => false,
            [1] => true,
            _ => return Err(invalid_durable("invalid transfer relationship mode")),
        };
        Ok(TransferRelationship {
            id,
            relationship_type,
            target,
            external,
        })
    }

    fn take(&mut self, length: usize) -> TransactionResult<&'a [u8]> {
        let end = self
            .offset
            .checked_add(length)
            .ok_or_else(|| invalid_durable("transfer graph offset overflowed"))?;
        let value = self
            .bytes
            .get(self.offset..end)
            .ok_or_else(|| invalid_durable("truncated transfer graph"))?;
        self.offset = end;
        Ok(value)
    }

    fn finish(self) -> TransactionResult<()> {
        if self.offset != self.bytes.len() {
            return Err(invalid_durable("transfer graph has trailing bytes"));
        }
        Ok(())
    }
}

pub(crate) fn durable_transfer_operations<Mode>(
    patch: &CorePatch<Mode>,
) -> TransactionResult<Vec<Operation>> {
    patch
        .operations()
        .iter()
        .filter(|operation| {
            matches!(
                operation.op.as_str(),
                "paragraph.transfer.insert"
                    | "paragraph.transfer.remove"
                    | RESTORE_TRANSFER_INSERT
                    | RESTORE_TRANSFER_REMOVE
            )
        })
        .map(|operation| {
            if matches!(
                operation.op.as_str(),
                RESTORE_TRANSFER_INSERT | RESTORE_TRANSFER_REMOVE
            ) {
                parse_restore_transfer_operation(operation, patch.blobs())
            } else {
                parse_durable_operation(operation, patch.blobs())
            }
        })
        .collect()
}

fn history_weight(snapshot: &Snapshot, patch: &Patch) -> TransactionResult<u64> {
    let mut bytes =
        u64::try_from(snapshot.xml_bytes().len()).map_err(|_error| TransactionError::Limit {
            kind: crate::document::LimitKind::HistoryTransitionBytes,
            maximum: usize::MAX,
            observed: snapshot.xml_bytes().len(),
        })?;
    if let Some(transition) = history_graph_transition(patch.operations())? {
        for part in transition.graph.parts.iter() {
            let part_bytes =
                u64::try_from(part.blob.len()).map_err(|_error| TransactionError::Limit {
                    kind: crate::document::LimitKind::HistoryDependencyBytes,
                    maximum: usize::MAX,
                    observed: part.blob.len(),
                })?;
            bytes = bytes
                .checked_add(part_bytes)
                .ok_or(TransactionError::Limit {
                    kind: crate::document::LimitKind::HistoryTransitionBytes,
                    maximum: usize::MAX,
                    observed: usize::MAX,
                })?;
        }
    }
    Ok(bytes)
}

fn history_graph_transition(
    operations: &[Operation],
) -> TransactionResult<Option<HistoryGraphTransition>> {
    let mut selected = None;
    for operation in operations {
        let candidate = match operation {
            Operation::InsertTransferredParagraph {
                dependency_digest,
                inverse_dependency_digest,
                graph,
                ..
            } if !graph.is_empty() => Some(HistoryGraphTransition {
                graph: Arc::clone(graph),
                before_digest: Arc::clone(dependency_digest),
                after_digest: Arc::clone(inverse_dependency_digest),
                forward_insert: true,
            }),
            Operation::RemoveTransferredParagraph {
                dependency_digest,
                inverse_dependency_digest,
                graph,
                ..
            } if !graph.is_empty() => Some(HistoryGraphTransition {
                graph: Arc::clone(graph),
                before_digest: Arc::clone(dependency_digest),
                after_digest: Arc::clone(inverse_dependency_digest),
                forward_insert: false,
            }),
            Operation::InsertTransferredParagraph { .. }
            | Operation::RemoveTransferredParagraph { .. }
            | Operation::ReplaceParagraphText { .. }
            | Operation::ReplaceHyperlinkText { .. }
            | Operation::ReplaceRunText { .. }
            | Operation::ReplaceRunTextNode { .. }
            | Operation::MarkCachedRunFieldDirty { .. }
            | Operation::ReplaceSimpleFieldText { .. }
            | Operation::ReplaceComplexFieldText { .. }
            | Operation::ReplaceRevisionText { .. }
            | Operation::ApplyRevision { .. }
            | Operation::ReplaceContentControlText { .. }
            | Operation::ReplaceNestedContentControlText { .. }
            | Operation::ReplaceNestedContentControlHyperlinkText { .. }
            | Operation::ReplaceBlockContentControlParagraphText { .. }
            | Operation::ReplaceBlockContentControlParagraphHyperlinkText { .. }
            | Operation::ReplaceCellText { .. }
            | Operation::ReplaceCellParagraphText { .. }
            | Operation::ReplaceNestedCellParagraphText { .. }
            | Operation::ReplaceNestedCellParagraphHyperlinkText { .. }
            | Operation::InsertParagraph { .. }
            | Operation::RemoveParagraph { .. }
            | Operation::ReplaceRunProperties { .. }
            | Operation::ReplaceParagraphProperties { .. }
            | Operation::InsertParagraphFragment { .. }
            | Operation::RemoveParagraphFragment { .. }
            | Operation::InsertRun { .. }
            | Operation::RemoveRun { .. }
            | Operation::ReplaceElement { .. }
            | Operation::InsertElement { .. }
            | Operation::RemoveElement { .. }
            | Operation::ChangeGraph { .. } => None,
        };
        if let Some(selected_transition) = candidate {
            if selected.is_some() {
                return Err(invalid_durable(
                    "history cannot retain multiple dependency subgraphs",
                ));
            }
            selected = Some(selected_transition);
        }
    }
    Ok(selected)
}

fn position_path(path: &[Position]) -> String {
    path.iter()
        .map(|position| position.get().to_string())
        .collect::<Vec<_>>()
        .join(",")
}

fn table_cell_path(path: &[TableCellAddress]) -> String {
    path.iter()
        .map(|address| {
            format!(
                "{},{},{}",
                address.table.get(),
                address.row.get(),
                address.cell.get()
            )
        })
        .collect::<Vec<_>>()
        .join(";")
}

fn direct_table_cell_path(table: Position, row: Position, cell: Position) -> String {
    format!("{},{},{}", table.get(), row.get(), cell.get())
}

fn push_inline_control_ancestor_reads(
    paragraph: Position,
    controls: &[Position],
    reads: &mut Vec<String>,
) {
    for end in 1..controls.len() {
        reads.push(format!(
            "body/paragraph:{}/content-control-path:{}/text",
            paragraph.get(),
            position_path(&controls[..end])
        ));
    }
}

fn push_table_cell_scope_reads(path: &[TableCellAddress], reads: &mut Vec<String>) {
    for end in 1..=path.len() {
        reads.push(format!(
            "body/table-cell-path:{}/text",
            table_cell_path(&path[..end])
        ));
    }
}

fn operation_effects(
    operations: &[Operation],
    source_paragraphs: usize,
) -> (Vec<String>, Vec<String>) {
    let mut reads = Vec::new();
    let mut writes = Vec::new();
    for operation in operations {
        match operation {
            Operation::ReplaceParagraphText { position, .. } => {
                reads.push("body/paragraph-order".to_owned());
                writes.push(format!("body/paragraph:{}/all-text", position.get()));
            },
            Operation::ReplaceHyperlinkText {
                paragraph,
                hyperlink,
                ..
            } => {
                reads.push("body/paragraph-order".to_owned());
                reads.push(format!("body/paragraph:{}/all-text", paragraph.get()));
                writes.push(format!(
                    "body/paragraph:{}/hyperlink:{}/text",
                    paragraph.get(),
                    hyperlink.get()
                ));
            },
            Operation::ReplaceRunText { paragraph, run, .. } => {
                reads.push("body/paragraph-order".to_owned());
                reads.push(format!("body/paragraph:{}/all-text", paragraph.get()));
                writes.push(format!(
                    "body/paragraph:{}/run:{}/text",
                    paragraph.get(),
                    run.get()
                ));
            },
            Operation::ReplaceRunTextNode { paragraph, run, .. } => {
                reads.push("body/paragraph-order".to_owned());
                reads.push(format!("body/paragraph:{}/all-text", paragraph.get()));
                writes.push(format!(
                    "body/paragraph:{}/run:{}/text-node",
                    paragraph.get(),
                    run.get()
                ));
            },
            Operation::MarkCachedRunFieldDirty { paragraph, run, .. } => {
                reads.push("body/paragraph-order".to_owned());
                reads.push(format!("body/paragraph:{}/all-text", paragraph.get()));
                writes.push(format!(
                    "body/paragraph:{}/run:{}/cached-field-dirty",
                    paragraph.get(),
                    run.get()
                ));
            },
            Operation::ReplaceSimpleFieldText {
                paragraph, field, ..
            } => {
                reads.push("body/paragraph-order".to_owned());
                reads.push(format!("body/paragraph:{}/all-text", paragraph.get()));
                writes.push(format!(
                    "body/paragraph:{}/field:{}/result-text",
                    paragraph.get(),
                    field.get()
                ));
            },
            Operation::ReplaceComplexFieldText {
                paragraph, field, ..
            } => {
                reads.push("body/paragraph-order".to_owned());
                reads.push(format!("body/paragraph:{}/all-text", paragraph.get()));
                writes.push(format!(
                    "body/paragraph:{}/complex-field:{}/result-text",
                    paragraph.get(),
                    field.get()
                ));
            },
            Operation::ReplaceRevisionText {
                paragraph,
                kind,
                revision,
                ..
            } => {
                reads.push("body/paragraph-order".to_owned());
                reads.push(format!("body/paragraph:{}/all-text", paragraph.get()));
                writes.push(format!(
                    "body/paragraph:{}/revision:{kind:?}:{}/text",
                    paragraph.get(),
                    revision.get()
                ));
            },
            Operation::ApplyRevision { selector, .. } => {
                reads.push("body/paragraph-order".to_owned());
                let paragraph_source =
                    format!("body/paragraph:{}/all-text", selector.paragraph().get());
                reads.push(paragraph_source.clone());
                // The action replaces the complete owning paragraph. Every
                // other inline owner reads this paragraph-wide source key, so
                // retaining it as a write makes exact-key composition refuse
                // same-paragraph joins while preserving cross-paragraph joins.
                writes.push(paragraph_source);
                writes.push(format!(
                    "body/paragraph:{}/revision-structure",
                    selector.paragraph().get()
                ));
            },
            Operation::ReplaceContentControlText {
                paragraph, control, ..
            } => {
                reads.push("body/paragraph-order".to_owned());
                reads.push(format!("body/paragraph:{}/all-text", paragraph.get()));
                writes.push(format!(
                    "body/paragraph:{}/content-control-path:{}/text",
                    paragraph.get(),
                    position_path(&[*control])
                ));
            },
            Operation::ReplaceNestedContentControlText {
                paragraph,
                controls,
                ..
            } => {
                reads.push("body/paragraph-order".to_owned());
                reads.push(format!("body/paragraph:{}/all-text", paragraph.get()));
                push_inline_control_ancestor_reads(*paragraph, controls, &mut reads);
                writes.push(format!(
                    "body/paragraph:{}/content-control-path:{}/text",
                    paragraph.get(),
                    position_path(controls)
                ));
            },
            Operation::ReplaceNestedContentControlHyperlinkText {
                paragraph,
                controls,
                hyperlink,
                ..
            } => {
                let owner = format!(
                    "body/paragraph:{}/content-control-path:{}/text",
                    paragraph.get(),
                    position_path(controls)
                );
                reads.push("body/paragraph-order".to_owned());
                reads.push(format!("body/paragraph:{}/all-text", paragraph.get()));
                push_inline_control_ancestor_reads(*paragraph, controls, &mut reads);
                reads.push(owner.clone());
                writes.push(format!("{owner}/hyperlink:{}/text", hyperlink.get()));
            },
            Operation::ReplaceBlockContentControlParagraphText {
                controls,
                paragraph,
                ..
            } => writes.push(format!(
                "body/block-content-control-path:{}/paragraph:{}/text",
                position_path(controls),
                paragraph.get()
            )),
            Operation::ReplaceBlockContentControlParagraphHyperlinkText {
                controls,
                paragraph,
                hyperlink,
                ..
            } => {
                let owner = format!(
                    "body/block-content-control-path:{}/paragraph:{}/text",
                    position_path(controls),
                    paragraph.get()
                );
                reads.push(owner.clone());
                writes.push(format!("{owner}/hyperlink:{}/text", hyperlink.get()));
            },
            Operation::ReplaceCellText {
                table, row, cell, ..
            } => writes.push(format!(
                "body/table-cell-path:{}/text",
                direct_table_cell_path(*table, *row, *cell)
            )),
            Operation::ReplaceCellParagraphText {
                table,
                row,
                cell,
                paragraph,
                ..
            } => {
                reads.push(format!(
                    "body/table-cell-path:{}/text",
                    direct_table_cell_path(*table, *row, *cell)
                ));
                writes.push(format!(
                    "body/table-cell-path:{}/paragraph:{}/text",
                    direct_table_cell_path(*table, *row, *cell),
                    paragraph.get()
                ));
            },
            Operation::ReplaceNestedCellParagraphText {
                path, paragraph, ..
            } => {
                push_table_cell_scope_reads(path, &mut reads);
                writes.push(format!(
                    "body/table-cell-path:{}/paragraph:{}/text",
                    table_cell_path(path),
                    paragraph.get()
                ));
            },
            Operation::ReplaceNestedCellParagraphHyperlinkText {
                path,
                paragraph,
                hyperlink,
                ..
            } => {
                let owner = format!(
                    "body/table-cell-path:{}/paragraph:{}/text",
                    table_cell_path(path),
                    paragraph.get()
                );
                push_table_cell_scope_reads(path, &mut reads);
                reads.push(owner.clone());
                writes.push(format!("{owner}/hyperlink:{}/text", hyperlink.get()));
            },
            Operation::InsertParagraph { position, .. }
            | Operation::InsertTransferredParagraph { position, .. }
                if position.get() >= source_paragraphs =>
            {
                // Appending preserves every source paragraph selector. The
                // order read still conflicts with a shifting edit, while the
                // boundary write prevents two independent appends.
                reads.push("body/paragraph-order".to_owned());
                writes.push("body/paragraph-append-boundary".to_owned());
            },
            Operation::InsertParagraph { .. }
            | Operation::RemoveParagraph { .. }
            | Operation::InsertTransferredParagraph { .. }
            | Operation::RemoveTransferredParagraph { .. }
            | Operation::InsertParagraphFragment { .. }
            | Operation::RemoveParagraphFragment { .. } => {
                writes.push("body/paragraph-order".to_owned());
            },
            Operation::ReplaceRunProperties { paragraph, run, .. } => {
                // A property container changes neither the run's text nor
                // the paragraph's run order, so it composes with text edits
                // of the same paragraph; whole-paragraph rewrites (revision
                // dispositions, run insertion and removal) write the
                // paragraph's `all-text` key and conflict.
                reads.push("body/paragraph-order".to_owned());
                reads.push(format!("body/paragraph:{}/all-text", paragraph.get()));
                writes.push(format!(
                    "body/paragraph:{}/run:{}/properties",
                    paragraph.get(),
                    run.get()
                ));
            },
            Operation::ReplaceParagraphProperties { paragraph, .. } => {
                reads.push("body/paragraph-order".to_owned());
                reads.push(format!("body/paragraph:{}/all-text", paragraph.get()));
                writes.push(format!("body/paragraph:{}/properties", paragraph.get()));
            },
            Operation::InsertRun { paragraph, .. } | Operation::RemoveRun { paragraph, .. } => {
                // The run order of the paragraph changes: every inline
                // selector of the paragraph reads this key.
                reads.push("body/paragraph-order".to_owned());
                writes.push(format!("body/paragraph:{}/all-text", paragraph.get()));
            },
            Operation::ReplaceElement { path, .. } => {
                // A path-addressed element is reached through the body's
                // block order; a direct-body paragraph's whole text changes
                // with it.
                reads.push("body/paragraph-order".to_owned());
                if let Some(super::ElementStep::Paragraph(position)) = path.steps().first() {
                    writes.push(format!("body/paragraph:{}/all-text", position.get()));
                }
                writes.push(format!("body/element:{path}"));
            },
            Operation::InsertElement { container, .. }
            | Operation::RemoveElement { container, .. } => {
                writes.push("body/paragraph-order".to_owned());
                writes.push(format!("body/element:{container}/children"));
            },
            Operation::ChangeGraph { .. } => {
                writes.push("package/graph".to_owned());
            },
        }
    }
    (reads, writes)
}

fn revision_xml_value(xml: &Arc<Vec<u8>>) -> Result<Value, litchi_core::patch::PatchError> {
    String::from_utf8(xml.as_slice().to_vec())
        .map(Value::String)
        .map_err(|_error| litchi_core::patch::PatchError::InvalidText {
            field: "revision XML",
        })
}

fn durable_operation(
    limits: PatchLimits,
    operation: &Operation,
    artifact: &str,
    target_artifact: &str,
    blobs: &mut BlobBundle,
) -> Result<PatchOperation, litchi_core::patch::PatchError> {
    let mut preconditions = artifact_precondition(artifact);
    preconditions.insert(
        "target_sha256".to_owned(),
        Value::String(target_artifact.to_owned()),
    );
    let (name, target, value) = match operation {
        Operation::ReplaceParagraphText {
            position,
            before,
            after,
        } => {
            preconditions.insert("before".to_owned(), Value::String(before.clone()));
            (
                "paragraph.text.replace",
                format!("paragraph:{}", position.get()),
                Value::String(after.clone()),
            )
        },
        Operation::ReplaceHyperlinkText {
            paragraph,
            hyperlink,
            before,
            after,
        } => {
            preconditions.insert("before".to_owned(), Value::String(before.clone()));
            (
                "hyperlink.text.replace",
                format!(
                    "paragraph:{}/hyperlink:{}",
                    paragraph.get(),
                    hyperlink.get()
                ),
                Value::String(after.clone()),
            )
        },
        Operation::ReplaceRunText {
            paragraph,
            run,
            before,
            after,
        } => {
            preconditions.insert("before".to_owned(), Value::String(before.clone()));
            (
                "run.text.replace",
                format!("paragraph:{}/run:{}", paragraph.get(), run.get()),
                Value::String(after.clone()),
            )
        },
        Operation::ReplaceRunTextNode {
            paragraph,
            run,
            before,
            after,
        } => {
            preconditions.insert("before".to_owned(), markup_value(before, "run text node")?);
            (
                "run.text-node.replace",
                format!("paragraph:{}/run:{}", paragraph.get(), run.get()),
                markup_value(after, "run text node")?,
            )
        },
        Operation::MarkCachedRunFieldDirty {
            paragraph,
            run,
            before,
            after,
        } => {
            preconditions.insert(
                "before".to_owned(),
                markup_value(before, "cached field marker")?,
            );
            (
                "run.cached-field-dirty.mark",
                format!("paragraph:{}/run:{}", paragraph.get(), run.get()),
                markup_value(after, "cached field marker")?,
            )
        },
        Operation::ReplaceSimpleFieldText {
            paragraph,
            field,
            before,
            after,
        } => {
            preconditions.insert("before".to_owned(), Value::String(before.clone()));
            (
                "field.result-text.replace",
                format!("paragraph:{}/field:{}", paragraph.get(), field.get()),
                Value::String(after.clone()),
            )
        },
        Operation::ReplaceComplexFieldText {
            paragraph,
            field,
            before,
            after,
        } => {
            preconditions.insert("before".to_owned(), Value::String(before.clone()));
            (
                "field.complex-result-text.replace",
                format!(
                    "paragraph:{}/complex-field:{}",
                    paragraph.get(),
                    field.get()
                ),
                Value::String(after.clone()),
            )
        },
        Operation::ReplaceRevisionText {
            paragraph,
            kind,
            revision,
            before,
            after,
        } => {
            preconditions.insert("before".to_owned(), Value::String(before.clone()));
            (
                match kind {
                    RevisionKind::Insertion => "revision.insertion-text.replace",
                    RevisionKind::Deletion => "revision.deletion-text.replace",
                },
                format!("paragraph:{}/revision:{}", paragraph.get(), revision.get()),
                Value::String(after.clone()),
            )
        },
        Operation::ApplyRevision {
            selector,
            action,
            before,
            after,
        } => {
            preconditions.insert("before".to_owned(), revision_xml_value(before)?);
            (
                match (selector.kind(), action) {
                    (RevisionKind::Insertion, RevisionAction::Accept) => {
                        "revision.insertion.accept"
                    },
                    (RevisionKind::Insertion, RevisionAction::Reject) => {
                        "revision.insertion.reject"
                    },
                    (RevisionKind::Deletion, RevisionAction::Accept) => "revision.deletion.accept",
                    (RevisionKind::Deletion, RevisionAction::Reject) => "revision.deletion.reject",
                },
                format!(
                    "paragraph:{}/revision:{}",
                    selector.paragraph().get(),
                    selector.revision().get()
                ),
                revision_xml_value(after)?,
            )
        },
        Operation::ReplaceContentControlText {
            paragraph,
            control,
            before,
            after,
        } => {
            preconditions.insert("before".to_owned(), Value::String(before.clone()));
            (
                "content-control.text.replace",
                format!(
                    "paragraph:{}/content-control:{}",
                    paragraph.get(),
                    control.get()
                ),
                Value::String(after.clone()),
            )
        },
        Operation::ReplaceNestedContentControlText {
            paragraph,
            controls,
            before,
            after,
        } => {
            preconditions.insert("before".to_owned(), Value::String(before.clone()));
            (
                "content-control.nested-text.replace",
                format!(
                    "paragraph:{}/content-control-path:{}",
                    paragraph.get(),
                    position_path(controls)
                ),
                Value::String(after.clone()),
            )
        },
        Operation::ReplaceNestedContentControlHyperlinkText {
            paragraph,
            controls,
            hyperlink,
            before,
            after,
        } => {
            preconditions.insert("before".to_owned(), Value::String(before.clone()));
            (
                "content-control.nested-hyperlink-text.replace",
                format!(
                    "paragraph:{}/content-control-path:{}/hyperlink:{}",
                    paragraph.get(),
                    position_path(controls),
                    hyperlink.get()
                ),
                Value::String(after.clone()),
            )
        },
        Operation::ReplaceBlockContentControlParagraphText {
            controls,
            paragraph,
            before,
            after,
        } => {
            preconditions.insert("before".to_owned(), Value::String(before.clone()));
            (
                "content-control.block-paragraph-text.replace",
                format!(
                    "block-content-control-path:{}/paragraph:{}",
                    position_path(controls),
                    paragraph.get()
                ),
                Value::String(after.clone()),
            )
        },
        Operation::ReplaceBlockContentControlParagraphHyperlinkText {
            controls,
            paragraph,
            hyperlink,
            before,
            after,
        } => {
            preconditions.insert("before".to_owned(), Value::String(before.clone()));
            (
                "content-control.block-paragraph-hyperlink-text.replace",
                format!(
                    "block-content-control-path:{}/paragraph:{}/hyperlink:{}",
                    position_path(controls),
                    paragraph.get(),
                    hyperlink.get()
                ),
                Value::String(after.clone()),
            )
        },
        Operation::ReplaceCellText {
            table,
            row,
            cell,
            before,
            after,
        } => {
            preconditions.insert("before".to_owned(), Value::String(before.clone()));
            (
                "cell.text.replace",
                format!(
                    "table:{}/row:{}/cell:{}",
                    table.get(),
                    row.get(),
                    cell.get()
                ),
                Value::String(after.clone()),
            )
        },
        Operation::ReplaceCellParagraphText {
            table,
            row,
            cell,
            paragraph,
            before,
            after,
        } => {
            preconditions.insert("before".to_owned(), Value::String(before.clone()));
            (
                "cell.paragraph-text.replace",
                format!(
                    "table:{}/row:{}/cell:{}/paragraph:{}",
                    table.get(),
                    row.get(),
                    cell.get(),
                    paragraph.get()
                ),
                Value::String(after.clone()),
            )
        },
        Operation::ReplaceNestedCellParagraphText {
            path,
            paragraph,
            before,
            after,
        } => {
            preconditions.insert("before".to_owned(), Value::String(before.clone()));
            (
                "cell.nested-paragraph-text.replace",
                format!(
                    "table-cell-path:{}/paragraph:{}",
                    table_cell_path(path),
                    paragraph.get()
                ),
                Value::String(after.clone()),
            )
        },
        Operation::ReplaceNestedCellParagraphHyperlinkText {
            path,
            paragraph,
            hyperlink,
            before,
            after,
        } => {
            preconditions.insert("before".to_owned(), Value::String(before.clone()));
            (
                "cell.nested-paragraph-hyperlink-text.replace",
                format!(
                    "table-cell-path:{}/paragraph:{}/hyperlink:{}",
                    table_cell_path(path),
                    paragraph.get(),
                    hyperlink.get()
                ),
                Value::String(after.clone()),
            )
        },
        Operation::InsertParagraph { position, text } => (
            "paragraph.insert",
            format!("paragraph:{}", position.get()),
            Value::String(text.clone()),
        ),
        Operation::RemoveParagraph { position, text } => (
            "paragraph.remove",
            format!("paragraph:{}", position.get()),
            Value::String(text.clone()),
        ),
        Operation::InsertTransferredParagraph {
            position,
            xml,
            dependency_digest,
            inverse_dependency_digest,
            graph,
        } => {
            preconditions.insert(
                "dependency_sha256".to_owned(),
                Value::String(dependency_digest.to_string()),
            );
            preconditions.insert(
                "inverse_dependency_sha256".to_owned(),
                Value::String(inverse_dependency_digest.to_string()),
            );
            if !graph.is_empty() {
                let graph_identifier = blobs.insert(&encode_transfer_graph(graph)?)?;
                preconditions.insert(
                    "graph_sha256".to_owned(),
                    Value::String(graph_identifier.as_hex()),
                );
            }
            let identifier = blobs.insert(xml.as_slice())?;
            (
                "paragraph.transfer.insert",
                format!("paragraph:{}", position.get()),
                Value::String(identifier.as_hex()),
            )
        },
        Operation::RemoveTransferredParagraph {
            position,
            xml,
            dependency_digest,
            inverse_dependency_digest,
            graph,
        } => {
            preconditions.insert(
                "dependency_sha256".to_owned(),
                Value::String(dependency_digest.to_string()),
            );
            preconditions.insert(
                "inverse_dependency_sha256".to_owned(),
                Value::String(inverse_dependency_digest.to_string()),
            );
            if !graph.is_empty() {
                let graph_identifier = blobs.insert(&encode_transfer_graph(graph)?)?;
                preconditions.insert(
                    "graph_sha256".to_owned(),
                    Value::String(graph_identifier.as_hex()),
                );
            }
            let identifier = blobs.insert(xml.as_slice())?;
            (
                "paragraph.transfer.remove",
                format!("paragraph:{}", position.get()),
                Value::String(identifier.as_hex()),
            )
        },
        Operation::ReplaceRunProperties {
            paragraph,
            run,
            before,
            after,
        } => {
            preconditions.insert("before".to_owned(), markup_value(before, "run properties")?);
            (
                "run.properties.replace",
                format!("paragraph:{}/run:{}", paragraph.get(), run.get()),
                markup_value(after, "run properties")?,
            )
        },
        Operation::ReplaceParagraphProperties {
            paragraph,
            before,
            after,
        } => {
            preconditions.insert(
                "before".to_owned(),
                markup_value(before, "paragraph properties")?,
            );
            (
                "paragraph.properties.replace",
                format!("paragraph:{}", paragraph.get()),
                markup_value(after, "paragraph properties")?,
            )
        },
        Operation::InsertParagraphFragment {
            child,
            position,
            xml,
        } => (
            "paragraph.fragment.insert",
            format!("body-child:{}/paragraph:{}", child.get(), position.get()),
            markup_value(xml, "paragraph")?,
        ),
        Operation::RemoveParagraphFragment {
            child,
            position,
            xml,
        } => (
            "paragraph.fragment.remove",
            format!("body-child:{}/paragraph:{}", child.get(), position.get()),
            markup_value(xml, "paragraph")?,
        ),
        Operation::InsertRun {
            paragraph,
            child,
            run,
            xml,
        } => (
            "run.fragment.insert",
            format!(
                "paragraph:{}/child:{}/run:{}",
                paragraph.get(),
                child.get(),
                run.get()
            ),
            markup_value(xml, "run")?,
        ),
        Operation::RemoveRun {
            paragraph,
            child,
            run,
            xml,
        } => (
            "run.fragment.remove",
            format!(
                "paragraph:{}/child:{}/run:{}",
                paragraph.get(),
                child.get(),
                run.get()
            ),
            markup_value(xml, "run")?,
        ),
        Operation::ChangeGraph { change, forward } => {
            let identifier = blobs.insert(&encode_graph_change(change)?)?;
            (
                "graph.change",
                if *forward {
                    "package-graph:forward".to_owned()
                } else {
                    "package-graph:inverse".to_owned()
                },
                Value::String(identifier.as_hex()),
            )
        },
        Operation::ReplaceElement {
            path,
            before,
            after,
        } => {
            preconditions.insert("before".to_owned(), markup_value(before, "element")?);
            (
                "element.replace",
                format!("element:{path}"),
                markup_value(after, "element")?,
            )
        },
        Operation::InsertElement {
            container,
            child,
            xml,
        } => (
            "element.insert",
            format!("container:{container}/child:{}", child.get()),
            markup_value(xml, "element")?,
        ),
        Operation::RemoveElement {
            container,
            child,
            xml,
        } => (
            "element.remove",
            format!("container:{container}/child:{}", child.get()),
            markup_value(xml, "element")?,
        ),
    };
    PatchOperation::new(limits, name, target, preconditions, value)
}

/// Exact markup carried in a durable operation as a JSON string.
fn markup_value(
    xml: &Arc<Vec<u8>>,
    field: &'static str,
) -> Result<Value, litchi_core::patch::PatchError> {
    String::from_utf8(xml.as_slice().to_vec())
        .map(Value::String)
        .map_err(|_error| litchi_core::patch::PatchError::InvalidText { field })
}

fn noop_operation(
    limits: PatchLimits,
    artifact: &str,
    target_artifact: &str,
) -> Result<PatchOperation, litchi_core::patch::PatchError> {
    let mut preconditions = artifact_precondition(artifact);
    preconditions.insert(
        "target_sha256".to_owned(),
        Value::String(target_artifact.to_owned()),
    );
    PatchOperation::new(
        limits,
        "document.noop",
        "document",
        preconditions,
        Value::Null,
    )
}

fn restore_operation(
    limits: PatchLimits,
    artifact: &str,
    target: &BlobId,
) -> Result<PatchOperation, litchi_core::patch::PatchError> {
    PatchOperation::new(
        limits,
        RESTORE_OPERATION,
        "document",
        artifact_precondition(artifact),
        Value::String(target.as_hex()),
    )
}

fn restore_transfer_operation(
    limits: PatchLimits,
    operation: &Operation,
    artifact: &str,
    target: &BlobId,
    blobs: &mut BlobBundle,
) -> Result<Option<PatchOperation>, litchi_core::patch::PatchError> {
    let (name, dependency, inverse_dependency, graph) = match operation {
        Operation::InsertTransferredParagraph {
            dependency_digest,
            inverse_dependency_digest,
            graph,
            ..
        } if !graph.is_empty() => (
            RESTORE_TRANSFER_REMOVE,
            inverse_dependency_digest,
            dependency_digest,
            graph,
        ),
        Operation::RemoveTransferredParagraph {
            dependency_digest,
            inverse_dependency_digest,
            graph,
            ..
        } if !graph.is_empty() => (
            RESTORE_TRANSFER_INSERT,
            inverse_dependency_digest,
            dependency_digest,
            graph,
        ),
        Operation::InsertTransferredParagraph { .. }
        | Operation::RemoveTransferredParagraph { .. }
        | Operation::ReplaceParagraphText { .. }
        | Operation::ReplaceHyperlinkText { .. }
        | Operation::ReplaceRunText { .. }
        | Operation::ReplaceRunTextNode { .. }
        | Operation::MarkCachedRunFieldDirty { .. }
        | Operation::ReplaceSimpleFieldText { .. }
        | Operation::ReplaceComplexFieldText { .. }
        | Operation::ReplaceRevisionText { .. }
        | Operation::ApplyRevision { .. }
        | Operation::ReplaceContentControlText { .. }
        | Operation::ReplaceNestedContentControlText { .. }
        | Operation::ReplaceNestedContentControlHyperlinkText { .. }
        | Operation::ReplaceBlockContentControlParagraphText { .. }
        | Operation::ReplaceBlockContentControlParagraphHyperlinkText { .. }
        | Operation::ReplaceCellText { .. }
        | Operation::ReplaceCellParagraphText { .. }
        | Operation::ReplaceNestedCellParagraphText { .. }
        | Operation::ReplaceNestedCellParagraphHyperlinkText { .. }
        | Operation::InsertParagraph { .. }
        | Operation::RemoveParagraph { .. }
        | Operation::ReplaceRunProperties { .. }
        | Operation::ReplaceParagraphProperties { .. }
        | Operation::InsertParagraphFragment { .. }
        | Operation::RemoveParagraphFragment { .. }
        | Operation::InsertRun { .. }
        | Operation::RemoveRun { .. }
        | Operation::ReplaceElement { .. }
        | Operation::InsertElement { .. }
        | Operation::RemoveElement { .. }
        | Operation::ChangeGraph { .. } => return Ok(None),
    };
    let graph_identifier = blobs.insert(encode_transfer_graph(graph)?)?;
    let mut preconditions = artifact_precondition(artifact);
    preconditions.insert(
        "dependency_sha256".to_owned(),
        Value::String(dependency.to_string()),
    );
    preconditions.insert(
        "inverse_dependency_sha256".to_owned(),
        Value::String(inverse_dependency.to_string()),
    );
    preconditions.insert(
        "graph_sha256".to_owned(),
        Value::String(graph_identifier.as_hex()),
    );
    PatchOperation::new(
        limits,
        name,
        "document",
        preconditions,
        Value::String(target.as_hex()),
    )
    .map(Some)
}

fn is_restore_operation(name: &str) -> bool {
    matches!(
        name,
        RESTORE_OPERATION | RESTORE_TRANSFER_INSERT | RESTORE_TRANSFER_REMOVE
    )
}

fn artifact_precondition(artifact: &str) -> BTreeMap<String, Value> {
    BTreeMap::from([(
        "artifact_sha256".to_owned(),
        Value::String(artifact.to_owned()),
    )])
}

fn encode_transfer_graph(graph: &TransferGraph) -> Result<Vec<u8>, litchi_core::patch::PatchError> {
    let mut output = b"LDXG1".to_vec();
    put_transfer_count(&mut output, graph.main_relationships.len())?;
    for relationship in graph.main_relationships.iter() {
        encode_transfer_relationship(&mut output, relationship)?;
    }
    put_transfer_count(&mut output, graph.parts.len())?;
    for part in graph.parts.iter() {
        put_transfer_bytes(&mut output, part.name.as_bytes())?;
        put_transfer_bytes(&mut output, part.content_type.as_bytes())?;
        put_transfer_bytes(&mut output, part.blob.as_slice())?;
        put_transfer_count(&mut output, part.relationships.len())?;
        for relationship in part.relationships.iter() {
            encode_transfer_relationship(&mut output, relationship)?;
        }
    }
    Ok(output)
}

/// A package-graph change: its transfer-graph form, then its replaced
/// parts (name, before, after).
fn encode_graph_change(change: &GraphChange) -> Result<Vec<u8>, litchi_core::patch::PatchError> {
    let mut output = encode_transfer_graph(&TransferGraph {
        main_relationships: Arc::clone(&change.relationships),
        parts: Arc::clone(&change.parts),
    })?;
    put_transfer_count(&mut output, change.replaced.len())?;
    for replaced in change.replaced.iter() {
        put_transfer_bytes(&mut output, replaced.name.as_bytes())?;
        put_transfer_bytes(&mut output, replaced.before.as_slice())?;
        put_transfer_bytes(&mut output, replaced.after.as_slice())?;
    }
    Ok(output)
}

fn decode_graph_change(bytes: &[u8]) -> TransactionResult<GraphChange> {
    let mut decoder = TransferGraphDecoder::new(bytes)?;
    let main_count = decoder.count(crate::document::LimitKind::MainRelationships, 64)?;
    let mut relationships = Vec::with_capacity(main_count);
    for _ in 0..main_count {
        relationships.push(decoder.relationship()?);
    }
    let part_count = decoder.count(crate::document::LimitKind::Parts, 256)?;
    let mut parts = Vec::with_capacity(part_count);
    let mut total_bytes = 0usize;
    for _ in 0..part_count {
        let name = decoder.string()?;
        let content_type = decoder.string()?;
        let blob = decoder.bytes()?;
        total_bytes = checked_transfer_graph_bytes(total_bytes, blob.len())?;
        let blob = blob.to_vec();
        let relationship_count =
            decoder.count(crate::document::LimitKind::PartRelationships, 256)?;
        let mut part_relationships = Vec::with_capacity(relationship_count);
        for _ in 0..relationship_count {
            part_relationships.push(decoder.relationship()?);
        }
        parts.push(TransferPart {
            name,
            content_type,
            blob: Arc::new(blob),
            relationships: part_relationships.into(),
        });
    }
    let replaced_count = decoder.count(crate::document::LimitKind::Parts, 256)?;
    let mut replaced = Vec::with_capacity(replaced_count);
    for _ in 0..replaced_count {
        let name = decoder.string()?;
        let before = decoder.bytes()?;
        total_bytes = checked_transfer_graph_bytes(total_bytes, before.len())?;
        let before = Arc::new(before.to_vec());
        let after = decoder.bytes()?;
        total_bytes = checked_transfer_graph_bytes(total_bytes, after.len())?;
        replaced.push(ReplacedPart {
            name,
            before,
            after: Arc::new(after.to_vec()),
        });
    }
    decoder.finish()?;
    Ok(GraphChange {
        relationships: relationships.into(),
        parts: parts.into(),
        replaced: replaced.into(),
    })
}

fn encode_transfer_relationship(
    output: &mut Vec<u8>,
    relationship: &TransferRelationship,
) -> Result<(), litchi_core::patch::PatchError> {
    put_transfer_bytes(output, relationship.id.as_bytes())?;
    put_transfer_bytes(output, relationship.relationship_type.as_bytes())?;
    put_transfer_bytes(output, relationship.target.as_bytes())?;
    output.push(u8::from(relationship.external));
    Ok(())
}

fn put_transfer_count(
    output: &mut Vec<u8>,
    value: usize,
) -> Result<(), litchi_core::patch::PatchError> {
    let encoded_value =
        u64::try_from(value).map_err(|_error| litchi_core::patch::PatchError::Allocation)?;
    output.extend_from_slice(&encoded_value.to_le_bytes());
    Ok(())
}

fn put_transfer_bytes(
    output: &mut Vec<u8>,
    value: &[u8],
) -> Result<(), litchi_core::patch::PatchError> {
    put_transfer_count(output, value.len())?;
    output.extend_from_slice(value);
    Ok(())
}

fn decode_transfer_graph(bytes: &[u8]) -> TransactionResult<TransferGraph> {
    let mut decoder = TransferGraphDecoder::new(bytes)?;
    let main_count = decoder.count(crate::document::LimitKind::MainRelationships, 64)?;
    let mut main_relationships = Vec::new();
    main_relationships
        .try_reserve_exact(main_count)
        .map_err(|_error| invalid_durable("transfer graph allocation failed"))?;
    for _ in 0..main_count {
        main_relationships.push(decoder.relationship()?);
    }
    let part_count = decoder.count(crate::document::LimitKind::Parts, 256)?;
    let mut parts = Vec::new();
    parts
        .try_reserve_exact(part_count)
        .map_err(|_error| invalid_durable("transfer graph allocation failed"))?;
    let mut total_bytes = 0usize;
    for _ in 0..part_count {
        let name = decoder.string()?;
        let content_type = decoder.string()?;
        let blob = decoder.bytes()?;
        total_bytes = checked_transfer_graph_bytes(total_bytes, blob.len())?;
        let blob = blob.to_vec();
        let relationship_count =
            decoder.count(crate::document::LimitKind::PartRelationships, 256)?;
        let mut relationships = Vec::new();
        relationships
            .try_reserve_exact(relationship_count)
            .map_err(|_error| invalid_durable("transfer graph allocation failed"))?;
        for _ in 0..relationship_count {
            relationships.push(decoder.relationship()?);
        }
        parts.push(TransferPart {
            name,
            content_type,
            blob: Arc::new(blob),
            relationships: relationships.into(),
        });
    }
    decoder.finish()?;
    Ok(TransferGraph {
        main_relationships: main_relationships.into(),
        parts: parts.into(),
    })
}

fn parse_durable_operation(
    operation: &PatchOperation,
    blobs: &BlobBundle,
) -> TransactionResult<Operation> {
    let before = || {
        operation
            .preconditions
            .get("before")
            .and_then(Value::as_str)
            .map(str::to_owned)
            .ok_or_else(|| invalid_durable("missing semantic before precondition"))
    };
    let after = || {
        operation
            .value
            .as_str()
            .map(str::to_owned)
            .ok_or_else(|| invalid_durable("operation value must be a string"))
    };
    match operation.op.as_str() {
        "paragraph.text.replace" if operation.preconditions.len() == 3 => {
            Ok(Operation::ReplaceParagraphText {
                position: parse_single_target(&operation.target, "paragraph:")?,
                before: before()?,
                after: after()?,
            })
        },
        "hyperlink.text.replace" if operation.preconditions.len() == 3 => {
            let (paragraph, hyperlink) = parse_hyperlink_target(&operation.target)?;
            Ok(Operation::ReplaceHyperlinkText {
                paragraph,
                hyperlink,
                before: before()?,
                after: after()?,
            })
        },
        "run.text.replace" if operation.preconditions.len() == 3 => {
            let (paragraph, run) = parse_paragraph_child_target(&operation.target, "/run:")?;
            Ok(Operation::ReplaceRunText {
                paragraph,
                run,
                before: before()?,
                after: after()?,
            })
        },
        "run.text-node.replace" if operation.preconditions.len() == 3 => {
            let (paragraph, run) = parse_paragraph_child_target(&operation.target, "/run:")?;
            Ok(Operation::ReplaceRunTextNode {
                paragraph,
                run,
                before: Arc::new(before()?.into_bytes()),
                after: Arc::new(after()?.into_bytes()),
            })
        },
        "run.cached-field-dirty.mark" if operation.preconditions.len() == 3 => {
            let (paragraph, run) = parse_paragraph_child_target(&operation.target, "/run:")?;
            Ok(Operation::MarkCachedRunFieldDirty {
                paragraph,
                run,
                before: Arc::new(before()?.into_bytes()),
                after: Arc::new(after()?.into_bytes()),
            })
        },
        "field.result-text.replace" if operation.preconditions.len() == 3 => {
            let (paragraph, field) = parse_paragraph_child_target(&operation.target, "/field:")?;
            Ok(Operation::ReplaceSimpleFieldText {
                paragraph,
                field,
                before: before()?,
                after: after()?,
            })
        },
        "field.complex-result-text.replace" if operation.preconditions.len() == 3 => {
            let (paragraph, field) =
                parse_paragraph_child_target(&operation.target, "/complex-field:")?;
            Ok(Operation::ReplaceComplexFieldText {
                paragraph,
                field,
                before: before()?,
                after: after()?,
            })
        },
        "revision.insertion-text.replace" | "revision.deletion-text.replace"
            if operation.preconditions.len() == 3 =>
        {
            let (paragraph, revision) =
                parse_paragraph_child_target(&operation.target, "/revision:")?;
            Ok(Operation::ReplaceRevisionText {
                paragraph,
                kind: if operation.op == "revision.insertion-text.replace" {
                    RevisionKind::Insertion
                } else {
                    RevisionKind::Deletion
                },
                revision,
                before: before()?,
                after: after()?,
            })
        },
        "revision.insertion.accept"
        | "revision.insertion.reject"
        | "revision.deletion.accept"
        | "revision.deletion.reject"
            if operation.preconditions.len() == 3 =>
        {
            let (paragraph, revision) =
                parse_paragraph_child_target(&operation.target, "/revision:")?;
            let kind = if operation.op.starts_with("revision.insertion.") {
                RevisionKind::Insertion
            } else {
                RevisionKind::Deletion
            };
            let action = if operation.op.ends_with(".accept") {
                RevisionAction::Accept
            } else {
                RevisionAction::Reject
            };
            Ok(Operation::ApplyRevision {
                selector: super::RevisionSelector::new(paragraph, kind, revision),
                action,
                before: Arc::new(before()?.into_bytes()),
                after: Arc::new(after()?.into_bytes()),
            })
        },
        "content-control.text.replace" if operation.preconditions.len() == 3 => {
            let (paragraph, control) =
                parse_paragraph_child_target(&operation.target, "/content-control:")?;
            Ok(Operation::ReplaceContentControlText {
                paragraph,
                control,
                before: before()?,
                after: after()?,
            })
        },
        "content-control.nested-text.replace" if operation.preconditions.len() == 3 => {
            let (paragraph, controls) = parse_nested_control_target(&operation.target)?;
            Ok(Operation::ReplaceNestedContentControlText {
                paragraph,
                controls,
                before: before()?,
                after: after()?,
            })
        },
        "content-control.nested-hyperlink-text.replace" if operation.preconditions.len() == 3 => {
            let (owner, hyperlink) = parse_composite_hyperlink_target(&operation.target)?;
            let (paragraph, controls) = parse_nested_control_target(owner)?;
            Ok(Operation::ReplaceNestedContentControlHyperlinkText {
                paragraph,
                controls,
                hyperlink,
                before: before()?,
                after: after()?,
            })
        },
        "content-control.block-paragraph-text.replace" if operation.preconditions.len() == 3 => {
            let (controls, paragraph) = parse_block_control_target(&operation.target)?;
            Ok(Operation::ReplaceBlockContentControlParagraphText {
                controls,
                paragraph,
                before: before()?,
                after: after()?,
            })
        },
        "content-control.block-paragraph-hyperlink-text.replace"
            if operation.preconditions.len() == 3 =>
        {
            let (owner, hyperlink) = parse_composite_hyperlink_target(&operation.target)?;
            let (controls, paragraph) = parse_block_control_target(owner)?;
            Ok(
                Operation::ReplaceBlockContentControlParagraphHyperlinkText {
                    controls,
                    paragraph,
                    hyperlink,
                    before: before()?,
                    after: after()?,
                },
            )
        },
        "cell.text.replace" if operation.preconditions.len() == 3 => {
            let (table, row, cell) = parse_cell_target(&operation.target)?;
            Ok(Operation::ReplaceCellText {
                table,
                row,
                cell,
                before: before()?,
                after: after()?,
            })
        },
        "cell.paragraph-text.replace" if operation.preconditions.len() == 3 => {
            let (table, row, cell, paragraph) = parse_cell_paragraph_target(&operation.target)?;
            Ok(Operation::ReplaceCellParagraphText {
                table,
                row,
                cell,
                paragraph,
                before: before()?,
                after: after()?,
            })
        },
        "cell.nested-paragraph-text.replace" if operation.preconditions.len() == 3 => {
            let (path, paragraph) = parse_nested_cell_target(&operation.target)?;
            Ok(Operation::ReplaceNestedCellParagraphText {
                path,
                paragraph,
                before: before()?,
                after: after()?,
            })
        },
        "cell.nested-paragraph-hyperlink-text.replace" if operation.preconditions.len() == 3 => {
            let (owner, hyperlink) = parse_composite_hyperlink_target(&operation.target)?;
            let (path, paragraph) = parse_nested_cell_target(owner)?;
            Ok(Operation::ReplaceNestedCellParagraphHyperlinkText {
                path,
                paragraph,
                hyperlink,
                before: before()?,
                after: after()?,
            })
        },
        "paragraph.insert" if operation.preconditions.len() == 2 => {
            Ok(Operation::InsertParagraph {
                position: parse_single_target(&operation.target, "paragraph:")?,
                text: after()?,
            })
        },
        "paragraph.remove" if operation.preconditions.len() == 2 => {
            Ok(Operation::RemoveParagraph {
                position: parse_single_target(&operation.target, "paragraph:")?,
                text: after()?,
            })
        },
        "run.properties.replace" if operation.preconditions.len() == 3 => {
            let (paragraph, run) = parse_paragraph_child_target(&operation.target, "/run:")?;
            Ok(Operation::ReplaceRunProperties {
                paragraph,
                run,
                before: Arc::new(before()?.into_bytes()),
                after: Arc::new(after()?.into_bytes()),
            })
        },
        "paragraph.properties.replace" if operation.preconditions.len() == 3 => {
            Ok(Operation::ReplaceParagraphProperties {
                paragraph: parse_single_target(&operation.target, "paragraph:")?,
                before: Arc::new(before()?.into_bytes()),
                after: Arc::new(after()?.into_bytes()),
            })
        },
        "paragraph.fragment.insert" | "paragraph.fragment.remove"
            if operation.preconditions.len() == 2 =>
        {
            let rest = operation
                .target
                .strip_prefix("body-child:")
                .ok_or_else(|| invalid_durable("invalid paragraph fragment target"))?;
            let (child, position) = rest
                .split_once("/paragraph:")
                .ok_or_else(|| invalid_durable("invalid paragraph fragment target"))?;
            let (child, position) = (parse_position(child)?, parse_position(position)?);
            let xml = Arc::new(after()?.into_bytes());
            Ok(if operation.op == "paragraph.fragment.insert" {
                Operation::InsertParagraphFragment {
                    child,
                    position,
                    xml,
                }
            } else {
                Operation::RemoveParagraphFragment {
                    child,
                    position,
                    xml,
                }
            })
        },
        "graph.change" if operation.preconditions.len() == 2 => {
            let forward = match operation.target.as_str() {
                "package-graph:forward" => true,
                "package-graph:inverse" => false,
                _ => return Err(invalid_durable("invalid package graph target")),
            };
            let identifier = operation
                .value
                .as_str()
                .ok_or_else(|| invalid_durable("graph change value must be a blob identifier"))?;
            let blob_id = blobs
                .ids()
                .find(|candidate| candidate.as_hex() == identifier)
                .ok_or_else(|| invalid_durable("missing graph change blob"))?;
            let bytes = blobs
                .get(blob_id)
                .ok_or_else(|| invalid_durable("missing graph change blob"))?;
            Ok(Operation::ChangeGraph {
                change: Arc::new(decode_graph_change(bytes)?),
                forward,
            })
        },
        "element.replace" if operation.preconditions.len() == 3 => {
            let path = operation
                .target
                .strip_prefix("element:")
                .and_then(super::ElementPath::parse)
                .ok_or_else(|| invalid_durable("invalid element target"))?;
            Ok(Operation::ReplaceElement {
                path,
                before: Arc::new(before()?.into_bytes()),
                after: Arc::new(after()?.into_bytes()),
            })
        },
        "element.insert" | "element.remove" if operation.preconditions.len() == 2 => {
            let target = || invalid_durable("invalid element container target");
            let rest = operation
                .target
                .strip_prefix("container:")
                .ok_or_else(target)?;
            let (container, child) = rest.rsplit_once("/child:").ok_or_else(target)?;
            let container = super::ElementPath::parse(container).ok_or_else(target)?;
            let child = parse_position(child)?;
            let xml = Arc::new(after()?.into_bytes());
            Ok(if operation.op == "element.insert" {
                Operation::InsertElement {
                    container,
                    child,
                    xml,
                }
            } else {
                Operation::RemoveElement {
                    container,
                    child,
                    xml,
                }
            })
        },
        "run.fragment.insert" | "run.fragment.remove" if operation.preconditions.len() == 2 => {
            let target = || invalid_durable("invalid run fragment target");
            let rest = operation
                .target
                .strip_prefix("paragraph:")
                .ok_or_else(target)?;
            let (paragraph, rest) = rest.split_once("/child:").ok_or_else(target)?;
            let (child, run) = rest.split_once("/run:").ok_or_else(target)?;
            let (paragraph, child, run) = (
                parse_position(paragraph)?,
                parse_position(child)?,
                parse_position(run)?,
            );
            let xml = Arc::new(after()?.into_bytes());
            Ok(if operation.op == "run.fragment.insert" {
                Operation::InsertRun {
                    paragraph,
                    child,
                    run,
                    xml,
                }
            } else {
                Operation::RemoveRun {
                    paragraph,
                    child,
                    run,
                    xml,
                }
            })
        },
        "paragraph.transfer.insert" | "paragraph.transfer.remove"
            if matches!(operation.preconditions.len(), 4 | 5) =>
        {
            let identifier = operation
                .value
                .as_str()
                .ok_or_else(|| invalid_durable("transfer value must be a blob identifier"))?;
            let blob_id = blobs
                .ids()
                .find(|candidate| candidate.as_hex() == identifier)
                .ok_or_else(|| invalid_durable("missing transfer paragraph blob"))?;
            let bytes = blobs
                .get(blob_id)
                .ok_or_else(|| invalid_durable("missing transfer paragraph blob"))?;
            let xml = Arc::new(bytes.to_vec());
            let dependency_digest = operation
                .preconditions
                .get("dependency_sha256")
                .and_then(Value::as_str)
                .map(Arc::<str>::from)
                .ok_or_else(|| invalid_durable("missing transfer dependency precondition"))?;
            let inverse_dependency_digest = operation
                .preconditions
                .get("inverse_dependency_sha256")
                .and_then(Value::as_str)
                .map(Arc::<str>::from)
                .ok_or_else(|| {
                    invalid_durable("missing inverse transfer dependency precondition")
                })?;
            let graph = if let Some(graph_identifier) = operation
                .preconditions
                .get("graph_sha256")
                .and_then(Value::as_str)
            {
                let graph_blob_id = blobs
                    .ids()
                    .find(|candidate| candidate.as_hex() == graph_identifier)
                    .ok_or_else(|| invalid_durable("missing transfer graph blob"))?;
                let graph_bytes = blobs
                    .get(graph_blob_id)
                    .ok_or_else(|| invalid_durable("missing transfer graph blob"))?;
                Arc::new(decode_transfer_graph(graph_bytes)?)
            } else {
                Arc::new(TransferGraph::empty())
            };
            let position = parse_single_target(&operation.target, "paragraph:")?;
            if operation.op == "paragraph.transfer.insert" {
                Ok(Operation::InsertTransferredParagraph {
                    position,
                    xml,
                    dependency_digest,
                    inverse_dependency_digest,
                    graph,
                })
            } else {
                Ok(Operation::RemoveTransferredParagraph {
                    position,
                    xml,
                    dependency_digest,
                    inverse_dependency_digest,
                    graph,
                })
            }
        },
        _ => Err(invalid_durable("unsupported operation vocabulary")),
    }
}

fn parse_restore_transfer_operation(
    operation: &PatchOperation,
    blobs: &BlobBundle,
) -> TransactionResult<Operation> {
    if operation.target != "document" || operation.preconditions.len() != 4 {
        return Err(invalid_durable("invalid restore transfer operation"));
    }
    let dependency_digest = operation
        .preconditions
        .get("dependency_sha256")
        .and_then(Value::as_str)
        .map(Arc::<str>::from)
        .ok_or_else(|| invalid_durable("missing restore transfer dependency"))?;
    let inverse_dependency_digest = operation
        .preconditions
        .get("inverse_dependency_sha256")
        .and_then(Value::as_str)
        .map(Arc::<str>::from)
        .ok_or_else(|| invalid_durable("missing inverse restore transfer dependency"))?;
    let graph_identifier = operation
        .preconditions
        .get("graph_sha256")
        .and_then(Value::as_str)
        .ok_or_else(|| invalid_durable("missing restore transfer graph"))?;
    let blob_id = blobs
        .ids()
        .find(|candidate| candidate.as_hex() == graph_identifier)
        .ok_or_else(|| invalid_durable("missing restore transfer graph blob"))?;
    let graph =
        Arc::new(decode_transfer_graph(blobs.get(blob_id).ok_or_else(
            || invalid_durable("missing restore transfer graph blob"),
        )?)?);
    let xml = Arc::new(Vec::new());
    if operation.op == RESTORE_TRANSFER_INSERT {
        Ok(Operation::InsertTransferredParagraph {
            position: Position::new(0),
            xml,
            dependency_digest,
            inverse_dependency_digest,
            graph,
        })
    } else if operation.op == RESTORE_TRANSFER_REMOVE {
        Ok(Operation::RemoveTransferredParagraph {
            position: Position::new(0),
            xml,
            dependency_digest,
            inverse_dependency_digest,
            graph,
        })
    } else {
        Err(invalid_durable("invalid restore transfer vocabulary"))
    }
}

fn validate_semantic_blobs<Mode>(patch: &CorePatch<Mode>) -> TransactionResult<()> {
    let mut referenced = BTreeSet::new();
    for operation in patch.operations().iter().filter(|operation| {
        matches!(
            operation.op.as_str(),
            "paragraph.transfer.insert" | "paragraph.transfer.remove" | "graph.change"
        )
    }) {
        referenced.insert(
            operation
                .value
                .as_str()
                .map(str::to_owned)
                .ok_or_else(|| invalid_durable("transfer value must be a blob identifier"))?,
        );
        if let Some(identifier) = operation
            .preconditions
            .get("graph_sha256")
            .and_then(Value::as_str)
        {
            referenced.insert(identifier.to_owned());
        }
    }
    if referenced.len() != patch.blobs().len()
        || patch
            .blobs()
            .ids()
            .any(|identifier| !referenced.contains(&identifier.as_hex()))
    {
        return Err(invalid_durable("unreferenced semantic blob"));
    }
    Ok(())
}

fn common_target_artifact<Mode>(patch: &CorePatch<Mode>) -> TransactionResult<&str> {
    let mut expected = None;
    for operation in patch.operations() {
        let target = operation
            .preconditions
            .get("target_sha256")
            .and_then(Value::as_str)
            .ok_or_else(|| invalid_durable("missing target artifact precondition"))?;
        if expected.is_some_and(|value| value != target) {
            return Err(invalid_durable("inconsistent target artifact precondition"));
        }
        expected = Some(target);
    }
    expected.ok_or_else(|| invalid_durable("missing target artifact precondition"))
}

fn restore_snapshot<Mode>(
    patch: &CorePatch<Mode>,
    identity: Option<Arc<SourceIdentity>>,
) -> TransactionResult<Snapshot> {
    let identifier = patch.operations()[0]
        .value
        .as_str()
        .ok_or_else(|| invalid_durable("restore target must be a blob identifier"))?;
    if patch.operations().iter().any(|operation| {
        operation.target != "document"
            || (operation.op == RESTORE_OPERATION && operation.preconditions.len() != 1)
            || (matches!(
                operation.op.as_str(),
                RESTORE_TRANSFER_INSERT | RESTORE_TRANSFER_REMOVE
            ) && operation.preconditions.len() != 4)
            || operation.value.as_str() != Some(identifier)
    }) {
        return Err(invalid_durable("inconsistent restore operations"));
    }
    let blob_id = patch
        .blobs()
        .ids()
        .find(|candidate| candidate.as_hex() == identifier)
        .ok_or_else(|| invalid_durable("missing restore artifact"))?;
    let mut referenced = BTreeSet::from([identifier.to_owned()]);
    for operation in patch.operations() {
        if let Some(graph) = operation
            .preconditions
            .get("graph_sha256")
            .and_then(Value::as_str)
        {
            referenced.insert(graph.to_owned());
        }
    }
    if referenced.len() != patch.blobs().len()
        || patch
            .blobs()
            .ids()
            .any(|candidate| !referenced.contains(&candidate.as_hex()))
    {
        return Err(invalid_durable("restore has unreferenced artifacts"));
    }
    let bytes = patch
        .blobs()
        .get(blob_id)
        .ok_or_else(|| invalid_durable("missing restore artifact"))?;
    if bytes.len() > MAX_DOCUMENT_XML_BYTES {
        return Err(TransactionError::Limit {
            kind: crate::document::LimitKind::RestoreArtifactBytes,
            maximum: MAX_DOCUMENT_XML_BYTES,
            observed: bytes.len(),
        });
    }
    if let Some(identity) = identity {
        Snapshot::from_owned_xml_with_identity(Arc::new(bytes.to_vec()), identity)
    } else {
        Snapshot::from_xml(bytes.to_vec())
    }
}

fn parse_single_target(target: &str, prefix: &str) -> TransactionResult<Position> {
    parse_position(
        target
            .strip_prefix(prefix)
            .ok_or_else(|| invalid_durable("invalid operation target"))?,
    )
}

fn parse_hyperlink_target(target: &str) -> TransactionResult<(Position, Position)> {
    parse_paragraph_child_target(target, "/hyperlink:")
}

fn parse_paragraph_child_target(
    target: &str,
    separator: &str,
) -> TransactionResult<(Position, Position)> {
    let rest = target
        .strip_prefix("paragraph:")
        .ok_or_else(|| invalid_durable("invalid paragraph child target"))?;
    let (paragraph, child) = rest
        .split_once(separator)
        .ok_or_else(|| invalid_durable("invalid paragraph child target"))?;
    Ok((parse_position(paragraph)?, parse_position(child)?))
}

fn parse_cell_target(target: &str) -> TransactionResult<(Position, Position, Position)> {
    let rest = target
        .strip_prefix("table:")
        .ok_or_else(|| invalid_durable("invalid cell target"))?;
    let (table, row_and_cell) = rest
        .split_once("/row:")
        .ok_or_else(|| invalid_durable("invalid cell target"))?;
    let (row, cell) = row_and_cell
        .split_once("/cell:")
        .ok_or_else(|| invalid_durable("invalid cell target"))?;
    Ok((
        parse_position(table)?,
        parse_position(row)?,
        parse_position(cell)?,
    ))
}

fn parse_cell_paragraph_target(
    target: &str,
) -> TransactionResult<(Position, Position, Position, Position)> {
    let (cell_target, paragraph) = target
        .split_once("/paragraph:")
        .ok_or_else(|| invalid_durable("invalid cell paragraph target"))?;
    let (table, row, cell) = parse_cell_target(cell_target)?;
    Ok((table, row, cell, parse_position(paragraph)?))
}

fn parse_position_path(value: &str) -> TransactionResult<Arc<[Position]>> {
    if value.is_empty() {
        return Err(invalid_durable("empty selector path"));
    }
    let components = value.split(',').collect::<Vec<_>>();
    if components.len() > MAX_OPERATIONS {
        return Err(TransactionError::Limit {
            kind: crate::document::LimitKind::DurableSelectorPath,
            maximum: MAX_OPERATIONS,
            observed: components.len(),
        });
    }
    components
        .into_iter()
        .map(parse_position)
        .collect::<TransactionResult<Vec<_>>>()
        .map(Arc::from)
}

fn parse_composite_hyperlink_target(target: &str) -> TransactionResult<(&str, Position)> {
    let (owner, hyperlink) = target
        .rsplit_once("/hyperlink:")
        .ok_or_else(|| invalid_durable("invalid composite hyperlink target"))?;
    Ok((owner, parse_position(hyperlink)?))
}

fn parse_nested_control_target(target: &str) -> TransactionResult<(Position, Arc<[Position]>)> {
    let rest = target
        .strip_prefix("paragraph:")
        .ok_or_else(|| invalid_durable("invalid nested control target"))?;
    let (paragraph, path) = rest
        .split_once("/content-control-path:")
        .ok_or_else(|| invalid_durable("invalid nested control target"))?;
    Ok((parse_position(paragraph)?, parse_position_path(path)?))
}

fn parse_block_control_target(target: &str) -> TransactionResult<(Arc<[Position]>, Position)> {
    let rest = target
        .strip_prefix("block-content-control-path:")
        .ok_or_else(|| invalid_durable("invalid block control target"))?;
    let (path, paragraph) = rest
        .split_once("/paragraph:")
        .ok_or_else(|| invalid_durable("invalid block control target"))?;
    Ok((parse_position_path(path)?, parse_position(paragraph)?))
}

fn parse_nested_cell_target(
    target: &str,
) -> TransactionResult<(Arc<[TableCellAddress]>, Position)> {
    let rest = target
        .strip_prefix("table-cell-path:")
        .ok_or_else(|| invalid_durable("invalid nested cell target"))?;
    let (path, paragraph) = rest
        .split_once("/paragraph:")
        .ok_or_else(|| invalid_durable("invalid nested cell target"))?;
    if path.is_empty() {
        return Err(invalid_durable("empty nested table path"));
    }
    let steps = path.split(';').collect::<Vec<_>>();
    if steps.len() > MAX_OPERATIONS {
        return Err(TransactionError::Limit {
            kind: crate::document::LimitKind::DurableNestedTablePath,
            maximum: MAX_OPERATIONS,
            observed: steps.len(),
        });
    }
    let addresses = steps
        .into_iter()
        .map(|step| {
            let mut components = step.split(',');
            let table = components
                .next()
                .ok_or_else(|| invalid_durable("invalid nested table step"))?;
            let row = components
                .next()
                .ok_or_else(|| invalid_durable("invalid nested table step"))?;
            let cell = components
                .next()
                .ok_or_else(|| invalid_durable("invalid nested table step"))?;
            if components.next().is_some() {
                return Err(invalid_durable("invalid nested table step"));
            }
            Ok(TableCellAddress::new(
                parse_position(table)?,
                parse_position(row)?,
                parse_position(cell)?,
            ))
        })
        .collect::<TransactionResult<Vec<_>>>()?;
    Ok((addresses.into(), parse_position(paragraph)?))
}

fn parse_position(value: &str) -> TransactionResult<Position> {
    if value.is_empty()
        || !value.bytes().all(|byte| byte.is_ascii_digit())
        || (value != "0" && value.starts_with('0'))
    {
        return Err(invalid_durable("non-canonical numeric selector"));
    }
    value
        .parse::<usize>()
        .map(Position::new)
        .map_err(|_error| invalid_durable("selector exceeds this platform"))
}

fn invalid_durable(message: &str) -> TransactionError {
    TransactionError::InvalidDurable(message.to_owned())
}

fn checked_transfer_graph_bytes(current: usize, additional: usize) -> TransactionResult<usize> {
    let observed = current.saturating_add(additional);
    if observed > MAX_TRANSFER_GRAPH_BYTES {
        return Err(TransactionError::Limit {
            kind: crate::document::LimitKind::TransferGraphBytes,
            observed,
            maximum: MAX_TRANSFER_GRAPH_BYTES,
        });
    }
    Ok(observed)
}

#[cfg(test)]
mod limit_tests {
    use super::*;

    #[test]
    fn transfer_graph_aggregate_limit_reports_boundary_and_overflow() {
        assert_eq!(
            checked_transfer_graph_bytes(MAX_TRANSFER_GRAPH_BYTES - 1, 1).unwrap(),
            MAX_TRANSFER_GRAPH_BYTES
        );
        for (current, additional, expected) in [
            (MAX_TRANSFER_GRAPH_BYTES, 1, MAX_TRANSFER_GRAPH_BYTES + 1),
            (usize::MAX, 1, usize::MAX),
        ] {
            assert!(
                matches!(checked_transfer_graph_bytes(current, additional), Err(TransactionError::Limit { kind: crate::document::LimitKind::TransferGraphBytes, observed, maximum: MAX_TRANSFER_GRAPH_BYTES }) if observed == expected)
            );
        }
    }

    #[test]
    fn decoder_reports_typed_relationship_count_before_reading_records() {
        let mut bytes = b"LDXG1".to_vec();
        bytes.extend_from_slice(&65u64.to_le_bytes());
        let mut decoder = TransferGraphDecoder::new(&bytes).unwrap();
        assert!(matches!(
            decoder.count(crate::document::LimitKind::MainRelationships, 64),
            Err(TransactionError::Limit {
                kind: crate::document::LimitKind::MainRelationships,
                observed: 65,
                maximum: 64
            })
        ));
    }
}
