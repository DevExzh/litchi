//! Durable semantic patches, disjoint composition, and bounded history.

use std::collections::BTreeMap;
use std::fmt;
use std::sync::Arc;

use litchi_core::Position;
use litchi_core::patch::{
    BlobBundle, BlobId, JoinedSubEdits, Patch as CorePatch, PatchLimits, PatchOperation,
    Reversible, ReversibleOperation, SubEdit,
};
use serde_json::Value;

use super::{
    Commit, CompositionLimits, Edit, HistoryLimits, MAX_DOCUMENT_XML_BYTES, MAX_OPERATIONS,
    Operation, Patch, Snapshot, TransactionError, TransactionResult,
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
        if !patch.blobs().is_empty()
            || patch
                .operations()
                .iter()
                .any(|operation| operation.op == RESTORE_OPERATION)
        {
            return Err(invalid_durable("invalid semantic blob bundle"));
        }
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
            let semantic = parse_durable_operation(operation)?;
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
        let (reads, writes) = operation_effects(&self.operations);
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
                        durable_operation(limits, operation, &forward_artifact, &reverse_artifact)?,
                        restore_operation(limits, &reverse_artifact, &source_blob)?,
                    ))
                })
                .collect::<Result<Vec<_>, litchi_core::patch::PatchError>>()?
        };
        CorePatch::<Reversible>::new(
            limits,
            FORMAT_NAME,
            operations,
            BlobBundle::new(limits.blobs()),
            reverse_blobs,
        )
        .map_err(TransactionError::from)
    }
}

fn operation_effects(operations: &[Operation]) -> (Vec<String>, Vec<String>) {
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
            Operation::ReplaceCellText {
                table, row, cell, ..
            } => writes.push(format!(
                "body/table:{}/row:{}/cell:{}/text",
                table.get(),
                row.get(),
                cell.get()
            )),
            Operation::InsertParagraph { .. } | Operation::RemoveParagraph { .. } => {
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

fn parse_durable_operation(operation: &PatchOperation) -> TransactionResult<Operation> {
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
        _ => Err(invalid_durable("unsupported operation vocabulary")),
    }
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
    let rest = target
        .strip_prefix("paragraph:")
        .ok_or_else(|| invalid_durable("invalid hyperlink target"))?;
    let (paragraph, hyperlink) = rest
        .split_once("/hyperlink:")
        .ok_or_else(|| invalid_durable("invalid hyperlink target"))?;
    Ok((parse_position(paragraph)?, parse_position(hyperlink)?))
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
