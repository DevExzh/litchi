//! Durable semantic patches, disjoint composition, and bounded history.

use std::collections::{BTreeMap, BTreeSet};
use std::fmt;
use std::sync::Arc;

use litchi_core::Position;
use litchi_core::patch::{
    BlobBundle, BlobId, JoinedSubEdits, MergeChoice, Patch as CorePatch, PatchLimits,
    PatchOperation, Reversible, ReversibleOperation, SubEdit, ThreeWayMergeFailure,
    ThreeWayMergePlan,
};
use serde_json::Value;

use super::{
    Commit, CompositionLimits, Edit, HistoryLimits, MAX_DOCUMENT_XML_BYTES, MAX_OPERATIONS,
    Operation, Patch, RevisionKind, Snapshot, TransactionError, TransactionResult,
};

const FORMAT_NAME: &str = "litchi-docx/document";
const RESTORE_OPERATION: &str = "document.restore";

#[derive(Clone, PartialEq, Eq)]
struct Lineage(Arc<Vec<u8>>);

/// One independently prepared DOCX document edit.
pub struct PreparedEdit {
    inner: SubEdit<Lineage, Vec<Operation>>,
}

impl PreparedEdit {
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

impl fmt::Debug for PreparedEdit {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("PreparedEdit")
            .field("identifier", &self.identifier())
            .field("operations", &self.operations())
            .finish()
    }
}

/// Recoverable failure to join one prepared document edit.
pub struct JoinError {
    failure: Box<litchi_core::patch::SubEditJoinFailure>,
    rejected: Box<PreparedEdit>,
}

impl JoinError {
    /// Structured common composition refusal.
    #[must_use]
    pub const fn failure(&self) -> &litchi_core::patch::SubEditJoinFailure {
        &self.failure
    }

    /// Recover the rejected prepared edit.
    #[must_use]
    pub fn into_rejected(self) -> PreparedEdit {
        *self.rejected
    }
}

impl fmt::Debug for JoinError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("JoinError")
            .field("failure", &self.failure)
            .field("rejected", &self.rejected)
            .finish()
    }
}

/// Bounded deterministic composition of provably disjoint document edits.
pub struct Composition {
    source: Snapshot,
    joined: JoinedSubEdits<Lineage, Vec<Operation>>,
}

impl Composition {
    /// Join work only when its exact source and semantic effects are disjoint.
    ///
    /// # Errors
    ///
    /// Returns a structured common refusal while retaining the rejected edit.
    pub fn join(&mut self, incoming: PreparedEdit) -> Result<&mut Self, JoinError> {
        if let Err(error) = self.joined.join(incoming.inner) {
            let (failure, rejected) = error.into_parts();
            return Err(JoinError {
                failure: Box::new(failure),
                rejected: Box::new(PreparedEdit { inner: rejected }),
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

    /// Commit every accepted edit atomically in stable identifier order.
    ///
    /// # Errors
    ///
    /// Returns a semantic precondition, bound, or XML validation error without
    /// changing the immutable source snapshot.
    pub fn commit(self) -> TransactionResult<Commit> {
        let mut edit = self.source.edit();
        for prepared in self.joined.into_sub_edits() {
            for operation in prepared.into_payload() {
                edit.apply_operation(&operation)?;
            }
        }
        edit.commit()
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
            .field("failure", &self.failure)
            .finish_non_exhaustive()
    }
}

/// Non-mutating bounded three-way plan for two branches from one exact base.
pub struct ThreeWayPlan {
    source: Snapshot,
    inner: ThreeWayMergePlan<Lineage, Vec<Operation>>,
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
        let Self { source, inner } = self;
        match inner.finish() {
            Ok(joined) => Ok(Composition { source, joined }),
            Err(unresolved) => Err(Box::new(Self {
                source,
                inner: *unresolved,
            })),
        }
    }
}

impl fmt::Debug for ThreeWayPlan {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("ThreeWayPlan")
            .field("conflicts", self.conflicts())
            .finish_non_exhaustive()
    }
}

/// Explicit byte-budgeted undo/redo history for document snapshots.
pub struct History {
    inner: litchi_core::patch::History<Snapshot>,
}

impl History {
    /// Start history at one immutable snapshot.
    #[must_use]
    pub fn new(snapshot: Snapshot, limits: HistoryLimits) -> Self {
        Self {
            inner: litchi_core::patch::History::new(snapshot, limits),
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

    pub(crate) fn ensure_can_record(&self, commit: &Commit) -> TransactionResult<()> {
        let bytes = u64::try_from(commit.snapshot().xml_bytes().len()).map_err(|_error| {
            TransactionError::Limit {
                resource: "history transition bytes",
                max: usize::MAX,
                actual: commit.snapshot().xml_bytes().len(),
            }
        })?;
        if bytes > self.limits().max_weight() {
            return Err(litchi_core::patch::PatchError::HistoryWeight {
                observed: bytes,
                limit: self.limits().max_weight(),
            }
            .into());
        }
        Ok(())
    }

    /// Record a commit using its exact published XML size as transition weight.
    ///
    /// # Errors
    ///
    /// Returns a history-weight error without changing history when the
    /// published snapshot alone exceeds the configured byte budget.
    pub fn record(&mut self, commit: Commit) -> TransactionResult<Vec<Snapshot>> {
        let weight = u64::try_from(commit.snapshot().xml_bytes().len()).map_err(|_error| {
            TransactionError::Limit {
                resource: "history transition bytes",
                max: usize::MAX,
                actual: commit.snapshot().xml_bytes().len(),
            }
        })?;
        self.inner
            .record(commit.snapshot, weight)
            .map_err(TransactionError::from)
    }

    /// Move one retained transition backward.
    pub fn undo(&mut self) -> bool {
        self.inner.undo()
    }

    /// Move one retained transition forward.
    pub fn redo(&mut self) -> bool {
        self.inner.redo()
    }
}

impl Snapshot {
    /// Start bounded deterministic composition against this exact source.
    #[must_use]
    pub fn compose(&self, limits: CompositionLimits) -> Composition {
        Composition {
            source: self.clone(),
            joined: JoinedSubEdits::new(self.lineage(), limits),
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
        let source = self.clone();
        let left_source = left.source.clone();
        let right_source = right.source.clone();
        match ThreeWayMergePlan::new(left.joined, right.joined) {
            Ok(inner) => Ok(ThreeWayPlan { source, inner }),
            Err(error) => {
                let failure = error.failure().clone();
                let (left_joined, right_joined) = error.into_branches();
                Err(ThreeWayError {
                    failure,
                    left: Box::new(Composition {
                        source: left_source,
                        joined: left_joined,
                    }),
                    right: Box::new(Composition {
                        source: right_source,
                        joined: right_joined,
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
        if patch.format() != FORMAT_NAME {
            return Err(invalid_durable("unsupported format"));
        }
        if patch.operations().is_empty() {
            return Err(invalid_durable("empty patch has no source precondition"));
        }
        if patch.operations().len() > MAX_OPERATIONS {
            return Err(TransactionError::Limit {
                resource: "durable operations",
                max: MAX_OPERATIONS,
                actual: patch.operations().len(),
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
                return Err(TransactionError::StaleSource);
            }
        }

        if patch
            .operations()
            .iter()
            .all(|operation| operation.op == RESTORE_OPERATION)
        {
            return restore_snapshot(patch);
        }
        if patch
            .operations()
            .iter()
            .any(|operation| operation.op == RESTORE_OPERATION)
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
        Lineage(Arc::clone(&self.xml))
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
    ) -> TransactionResult<PreparedEdit> {
        let (reads, writes) = operation_effects(&self.operations, self.base.paragraph_count());
        let inner = SubEdit::new(
            self.base.lineage(),
            limits,
            identifier,
            reads,
            writes,
            self.operations,
        )?;
        Ok(PreparedEdit { inner })
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
                    Ok(ReversibleOperation::new(
                        durable_operation(
                            limits,
                            operation,
                            &forward_artifact,
                            &reverse_artifact,
                            &mut forward_blobs,
                        )?,
                        restore_operation(limits, &reverse_artifact, &source_blob)?,
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
            Operation::ReplaceContentControlText {
                paragraph, control, ..
            } => {
                reads.push("body/paragraph-order".to_owned());
                reads.push(format!("body/paragraph:{}/all-text", paragraph.get()));
                writes.push(format!(
                    "body/paragraph:{}/content-control:{}/text",
                    paragraph.get(),
                    control.get()
                ));
            },
            Operation::ReplaceCellText {
                table, row, cell, ..
            } => writes.push(format!(
                "body/table:{}/row:{}/cell:{}/text",
                table.get(),
                row.get(),
                cell.get()
            )),
            Operation::ReplaceCellParagraphText {
                table,
                row,
                cell,
                paragraph,
                ..
            } => {
                reads.push(format!(
                    "body/table:{}/row:{}/cell:{}/text",
                    table.get(),
                    row.get(),
                    cell.get()
                ));
                writes.push(format!(
                    "body/table:{}/row:{}/cell:{}/paragraph:{}/text",
                    table.get(),
                    row.get(),
                    cell.get(),
                    paragraph.get()
                ));
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
            | Operation::RemoveTransferredParagraph { .. } => {
                writes.push("body/paragraph-order".to_owned());
            },
        }
    }
    (reads, writes)
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
        } => {
            preconditions.insert(
                "dependency_sha256".to_owned(),
                Value::String(dependency_digest.to_string()),
            );
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
        } => {
            preconditions.insert(
                "dependency_sha256".to_owned(),
                Value::String(dependency_digest.to_string()),
            );
            let identifier = blobs.insert(xml.as_slice())?;
            (
                "paragraph.transfer.remove",
                format!("paragraph:{}", position.get()),
                Value::String(identifier.as_hex()),
            )
        },
    };
    PatchOperation::new(limits, name, target, preconditions, value)
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

fn artifact_precondition(artifact: &str) -> BTreeMap<String, Value> {
    BTreeMap::from([(
        "artifact_sha256".to_owned(),
        Value::String(artifact.to_owned()),
    )])
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
        "field.result-text.replace" if operation.preconditions.len() == 3 => {
            let (paragraph, field) = parse_paragraph_child_target(&operation.target, "/field:")?;
            Ok(Operation::ReplaceSimpleFieldText {
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
        "paragraph.transfer.insert" | "paragraph.transfer.remove"
            if operation.preconditions.len() == 3 =>
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
            let position = parse_single_target(&operation.target, "paragraph:")?;
            if operation.op == "paragraph.transfer.insert" {
                Ok(Operation::InsertTransferredParagraph {
                    position,
                    xml,
                    dependency_digest,
                })
            } else {
                Ok(Operation::RemoveTransferredParagraph {
                    position,
                    xml,
                    dependency_digest,
                })
            }
        },
        _ => Err(invalid_durable("unsupported operation vocabulary")),
    }
}

fn validate_semantic_blobs<Mode>(patch: &CorePatch<Mode>) -> TransactionResult<()> {
    let referenced = patch
        .operations()
        .iter()
        .filter(|operation| {
            matches!(
                operation.op.as_str(),
                "paragraph.transfer.insert" | "paragraph.transfer.remove"
            )
        })
        .map(|operation| {
            operation
                .value
                .as_str()
                .map(str::to_owned)
                .ok_or_else(|| invalid_durable("transfer value must be a blob identifier"))
        })
        .collect::<TransactionResult<BTreeSet<_>>>()?;
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

fn restore_snapshot<Mode>(patch: &CorePatch<Mode>) -> TransactionResult<Snapshot> {
    if patch.blobs().len() != 1 {
        return Err(invalid_durable("restore requires exactly one artifact"));
    }
    let identifier = patch.operations()[0]
        .value
        .as_str()
        .ok_or_else(|| invalid_durable("restore target must be a blob identifier"))?;
    if patch.operations().iter().any(|operation| {
        operation.target != "document"
            || operation.preconditions.len() != 1
            || operation.value.as_str() != Some(identifier)
    }) {
        return Err(invalid_durable("inconsistent restore operations"));
    }
    let blob_id = patch
        .blobs()
        .ids()
        .next()
        .ok_or_else(|| invalid_durable("missing restore artifact"))?;
    if blob_id.as_hex() != identifier {
        return Err(invalid_durable("restore artifact identifier mismatch"));
    }
    let bytes = patch
        .blobs()
        .get(blob_id)
        .ok_or_else(|| invalid_durable("missing restore artifact"))?;
    if bytes.len() > MAX_DOCUMENT_XML_BYTES {
        return Err(TransactionError::Limit {
            resource: "restore artifact bytes",
            max: MAX_DOCUMENT_XML_BYTES,
            actual: bytes.len(),
        });
    }
    Snapshot::from_xml(bytes.to_vec())
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
