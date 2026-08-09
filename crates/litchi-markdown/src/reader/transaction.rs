use std::fmt;
use std::sync::Arc;

use crate::escape;
use serde::{Deserialize, Serialize};

use super::model::{Dialect, Error, ReadLimits, Snapshot};

#[derive(Clone, Debug)]
enum Intent {
    Replace {
        position: usize,
        replacement: String,
    },
    Append {
        block: String,
    },
}

/// A bounded one-operation edit against an immutable Markdown snapshot.
#[derive(Debug)]
pub struct Edit<'snapshot> {
    source: &'snapshot Snapshot,
    intents: Vec<Intent>,
}

impl<'snapshot> Edit<'snapshot> {
    pub(crate) const fn new(source: &'snapshot Snapshot) -> Self {
        Self {
            source,
            intents: Vec::new(),
        }
    }

    /// Append exactly one parsed Markdown block.
    ///
    /// The adapter inserts the minimum deterministic blank-line separator:
    /// none for an empty source, two LF bytes after non-newline-terminated
    /// source, and one LF byte after newline-terminated source.
    ///
    /// # Errors
    ///
    /// Returns a typed error if an operation is already staged or `block` does
    /// not parse as exactly one top-level block under the snapshot policy.
    pub fn append_block(&mut self, block: &str) -> Result<&mut Self, Error> {
        self.ensure_capacity()?;
        validate_replacement(self.source, block)?;
        self.intents.push(Intent::Append {
            block: copy_source(block)?,
        });
        Ok(self)
    }

    /// Append literal text as one safely escaped paragraph block.
    ///
    /// # Errors
    ///
    /// Returns the same errors as [`Self::append_block`]. In particular, an
    /// empty value or literal blank-line sequence is refused because it cannot
    /// represent exactly one paragraph without changing the supplied text.
    pub fn append_text(&mut self, text: &str) -> Result<&mut Self, Error> {
        let escaped = escape::text(text);
        self.append_block(&escaped)
    }

    /// Atomically validate and publish the staged candidate.
    ///
    /// # Errors
    ///
    /// Returns a typed error without modifying the source when no operation is
    /// staged or the complete candidate violates its retained read policy.
    pub fn commit(self) -> Result<Commit, Error> {
        if self.intents.is_empty() {
            return Err(Error::NoStagedOperation);
        }
        validate_dependencies(self.source, &self.intents)?;
        let target = render(self.source, &self.intents)?;
        publish(self.source, &target, &self.intents)
    }

    fn ensure_capacity(&self) -> Result<(), Error> {
        if self.intents.len() == self.source.limits().max_operations {
            return Err(Error::OperationLimitExceeded {
                limit: self.source.limits().max_operations,
            });
        }
        Ok(())
    }

    /// Remove one selected top-level block without normalizing surrounding
    /// whitespace.
    ///
    /// # Errors
    ///
    /// Returns a typed missing-position or already-staged error.
    pub fn remove_block(&mut self, position: usize) -> Result<&mut Self, Error> {
        self.replace_block(position, "")
    }

    /// Replace one selected top-level block with zero or one parsed block.
    ///
    /// An empty replacement removes the selected block. Nonempty input must
    /// parse as exactly one top-level block, including a link definition.
    /// Untouched source bytes, including surrounding blank lines, remain exact.
    ///
    /// # Errors
    ///
    /// Returns a typed missing-position, already-staged, input, allocation, or
    /// resource-limit error.
    pub fn replace_block(
        &mut self,
        position: usize,
        replacement: &str,
    ) -> Result<&mut Self, Error> {
        self.ensure_capacity()?;
        if self.source.block(position).is_none() {
            return Err(Error::BlockNotFound { position });
        }
        if self.intents.iter().any(
            |intent| matches!(intent, Intent::Replace { position: existing, .. } if *existing == position),
        ) {
            return Err(Error::OverlappingOperation { position });
        }
        if !replacement.is_empty() {
            validate_replacement(self.source, replacement)?;
        }
        self.intents.push(Intent::Replace {
            position,
            replacement: copy_source(replacement)?,
        });
        Ok(self)
    }

    /// Replace one block with one safely escaped literal paragraph.
    ///
    /// Markdown delimiters in `text` cannot become active syntax. Use
    /// [`Self::replace_block`] when the replacement is intentionally Markdown.
    ///
    /// # Errors
    ///
    /// Returns the same errors as [`Self::replace_block`]. Empty text and a
    /// literal blank-line sequence are refused rather than treated as removal
    /// or silently collapsed; use [`Self::remove_block`] for removal.
    pub fn replace_block_with_text(
        &mut self,
        position: usize,
        text: &str,
    ) -> Result<&mut Self, Error> {
        let escaped = escape::text(text);
        if escaped.is_empty() {
            return Err(Error::ReplacementBlockCount { actual: 0 });
        }
        self.replace_block(position, &escaped)
    }
}

/// Diagnostics for one Markdown publication.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct Diagnostics {
    changed: bool,
    touched_blocks: usize,
    full_reparse_performed: bool,
}

impl Diagnostics {
    /// Whether the exact source changed.
    #[must_use]
    pub const fn changed(self) -> bool {
        self.changed
    }

    /// Whether the changed candidate required full parsing.
    #[must_use]
    pub const fn full_reparse_performed(self) -> bool {
        self.full_reparse_performed
    }

    /// Number of semantic blocks directly targeted by the operation.
    #[must_use]
    pub const fn touched_blocks(self) -> usize {
        self.touched_blocks
    }
}

/// A validated Markdown snapshot, reversible patch, and diagnostics.
#[derive(Debug)]
pub struct Commit {
    snapshot: Snapshot,
    patch: Patch,
    diagnostics: Diagnostics,
}

impl Commit {
    /// Publication diagnostics.
    #[must_use]
    pub const fn diagnostics(&self) -> Diagnostics {
        self.diagnostics
    }

    /// Reversible exact-source patch.
    #[must_use]
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Newly validated immutable snapshot.
    #[must_use]
    pub const fn snapshot(&self) -> &Snapshot {
        &self.snapshot
    }

    /// Consume this commit into its snapshot, patch, and diagnostics.
    #[must_use]
    pub fn into_parts(self) -> (Snapshot, Patch, Diagnostics) {
        (self.snapshot, self.patch, self.diagnostics)
    }
}

/// An in-memory reversible patch guarded by an exact complete before-image.
#[derive(Clone, PartialEq, Eq)]
pub struct Patch {
    before: Arc<str>,
    after: Arc<str>,
    dialect: Dialect,
    limits: ReadLimits,
    source_fingerprint: u64,
    target_fingerprint: u64,
    operation_count: usize,
    operations: Option<Box<[SemanticOperation]>>,
}

/// Bounds for deterministic durable patch JSON.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct PatchEnvelopeLimits {
    /// Maximum encoded JSON bytes.
    pub max_json_bytes: usize,
    /// Maximum bytes in either exact source image.
    pub max_source_bytes: usize,
}

impl PatchEnvelopeLimits {
    /// Conservative defaults for durable interchange.
    pub const DEFAULT: Self = Self {
        max_json_bytes: 32 * 1024 * 1024,
        max_source_bytes: 16 * 1024 * 1024,
    };
}

impl Default for PatchEnvelopeLimits {
    fn default() -> Self {
        Self::DEFAULT
    }
}

#[derive(Deserialize, Serialize)]
struct DurableEnvelope {
    version: u8,
    dialect: u8,
    max_source_bytes: usize,
    max_line_bytes: usize,
    max_events: usize,
    max_blocks: usize,
    max_nesting_depth: usize,
    max_operations: usize,
    before: String,
    after: String,
    operation_count: usize,
    operations: Option<Vec<SemanticOperation>>,
}

#[derive(Clone, Debug, Deserialize, PartialEq, Eq, Serialize)]
struct SemanticOperation {
    position: Option<usize>,
    replacement: String,
}

/// One conflicting immutable base-block position.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct Conflict {
    position: usize,
}

impl Conflict {
    /// Zero-based immutable base-block position written by both patches.
    #[must_use]
    pub const fn position(self) -> usize {
        self.position
    }
}

/// Deterministic source-ordered conflicts between independent patches.
#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct ConflictSet {
    conflicts: Box<[Conflict]>,
}

impl ConflictSet {
    /// Conflicts in ascending immutable base position.
    #[must_use]
    pub const fn conflicts(&self) -> &[Conflict] {
        &self.conflicts
    }

    /// Whether the set is empty.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.conflicts.is_empty()
    }
}

/// Failure to join two independently prepared patches.
#[derive(Debug, thiserror::Error)]
pub enum JoinError {
    /// Both patches write at least one base block differently.
    #[error("Markdown patches have {} conflicting base blocks", .0.conflicts.len())]
    Conflicts(ConflictSet),
    /// A source, limit, allocation, dependency, or candidate validation failed.
    #[error(transparent)]
    Validation(#[from] Error),
}

/// Non-mutating three-way merge result.
#[derive(Debug)]
pub struct MergePlan {
    commit: Option<Commit>,
    conflicts: ConflictSet,
}

impl MergePlan {
    /// Structured unresolved conflicts.
    #[must_use]
    pub const fn conflicts(&self) -> &ConflictSet {
        &self.conflicts
    }

    /// Validated merged candidate when no conflicts exist.
    #[must_use]
    pub const fn merged_commit(&self) -> Option<&Commit> {
        self.commit.as_ref()
    }

    /// Consume a conflict-free plan into its already validated commit.
    #[must_use]
    pub fn into_commit(self) -> Option<Commit> {
        self.commit
    }
}

impl Patch {
    /// Decode and fully validate deterministic durable patch JSON.
    ///
    /// # Errors
    ///
    /// Returns a typed size, schema, version, dialect, or Markdown validation
    /// error. Neither source image is published before both parse successfully.
    pub fn from_json(json: &str, limits: PatchEnvelopeLimits) -> Result<Self, Error> {
        if json.len() > limits.max_json_bytes {
            return Err(Error::PatchEnvelopeTooLarge {
                actual: json.len(),
                limit: limits.max_json_bytes,
            });
        }
        let envelope: DurableEnvelope =
            serde_json::from_str(json).map_err(|error| Error::InvalidPatchEnvelope {
                reason: error.to_string(),
            })?;
        if envelope.version != 1 {
            return Err(Error::InvalidPatchEnvelope {
                reason: format!("unsupported version {}", envelope.version),
            });
        }
        let dialect = match envelope.dialect {
            0 => Dialect::CommonMark,
            1 => Dialect::GitHubFlavored,
            value => {
                return Err(Error::InvalidPatchEnvelope {
                    reason: format!("unknown dialect {value}"),
                });
            },
        };
        for length in [envelope.before.len(), envelope.after.len()] {
            if length > limits.max_source_bytes {
                return Err(Error::PatchEnvelopeTooLarge {
                    actual: length,
                    limit: limits.max_source_bytes,
                });
            }
        }
        let read_limits = ReadLimits {
            max_source_bytes: envelope.max_source_bytes,
            max_line_bytes: envelope.max_line_bytes,
            max_events: envelope.max_events,
            max_blocks: envelope.max_blocks,
            max_nesting_depth: envelope.max_nesting_depth,
            max_operations: envelope.max_operations,
        };
        Snapshot::read_with(&envelope.before, dialect, read_limits)?;
        Snapshot::read_with(&envelope.after, dialect, read_limits)?;
        if let Some(operations) = envelope.operations.as_deref() {
            if envelope.operation_count != operations.len() {
                return Err(Error::InvalidPatchEnvelope {
                    reason: "operation count does not match operation array".to_owned(),
                });
            }
            let base = Snapshot::read_with(&envelope.before, dialect, read_limits)?;
            let replayed = replay_semantic_operations(&base, operations)?;
            if replayed.snapshot().source() != envelope.after {
                return Err(Error::InvalidPatchEnvelope {
                    reason: "semantic operations do not reproduce the after-image".to_owned(),
                });
            }
        }
        let source_fingerprint = fingerprint(&envelope.before);
        let target_fingerprint = fingerprint(&envelope.after);
        Ok(Self {
            before: Arc::from(envelope.before),
            after: Arc::from(envelope.after),
            dialect,
            limits: read_limits,
            source_fingerprint,
            target_fingerprint,
            operation_count: envelope.operation_count,
            operations: envelope.operations.map(Vec::into_boxed_slice),
        })
    }

    /// Return an exact inverse patch.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            before: Arc::clone(&self.after),
            after: Arc::clone(&self.before),
            dialect: self.dialect,
            limits: self.limits,
            source_fingerprint: self.target_fingerprint,
            target_fingerprint: self.source_fingerprint,
            operation_count: self.operation_count,
            operations: None,
        }
    }

    /// Join independently prepared patches against the same exact snapshot.
    ///
    /// Identical writes to one block are coalesced. Different writes to one
    /// block return a deterministic [`ConflictSet`]. Appends are independent
    /// and retain left-then-right patch order.
    ///
    /// # Errors
    ///
    /// Returns structured conflicts or a typed exact-source, operation-limit,
    /// dependency, allocation, or candidate-validation error.
    pub fn join(&self, other: &Self) -> Result<Commit, JoinError> {
        ensure_common_base(self, other)?;
        let (operations, conflicts) = joined_operations(self, other)?;
        if !conflicts.is_empty() {
            return Err(JoinError::Conflicts(conflicts));
        }
        let base = Snapshot::read_with(&self.before, self.dialect, self.limits)?;
        replay_semantic_operations(&base, &operations).map_err(JoinError::from)
    }

    /// Build a non-mutating three-way merge plan for two patches sharing this
    /// patch's immutable base snapshot.
    ///
    /// # Errors
    ///
    /// Returns a typed source or validation error. Semantic overlap is carried
    /// in the returned plan rather than published or raised as validation.
    pub fn plan_merge(&self, other: &Self) -> Result<MergePlan, Error> {
        ensure_common_base(self, other)?;
        let (operations, conflicts) = joined_operations(self, other)?;
        if !conflicts.is_empty() {
            return Ok(MergePlan {
                commit: None,
                conflicts,
            });
        }
        let base = Snapshot::read_with(&self.before, self.dialect, self.limits)?;
        let commit = replay_semantic_operations(&base, &operations)?;
        Ok(MergePlan {
            commit: Some(commit),
            conflicts,
        })
    }

    /// Whether the before- and after-images are byte-identical.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.before == self.after
    }

    /// Deterministic non-cryptographic fingerprint of the exact before-image.
    #[must_use]
    pub const fn source_fingerprint(&self) -> u64 {
        self.source_fingerprint
    }

    /// Deterministic non-cryptographic fingerprint of the exact after-image.
    #[must_use]
    pub const fn target_fingerprint(&self) -> u64 {
        self.target_fingerprint
    }

    /// Number of semantic operations represented by this patch.
    #[must_use]
    pub const fn operation_count(&self) -> usize {
        self.operation_count
    }

    /// Encode a stable-field-order, versioned, reversible JSON envelope.
    ///
    /// # Errors
    ///
    /// Returns a typed bound or serialization error.
    pub fn to_json(&self, limits: PatchEnvelopeLimits) -> Result<String, Error> {
        for length in [self.before.len(), self.after.len()] {
            if length > limits.max_source_bytes {
                return Err(Error::PatchEnvelopeTooLarge {
                    actual: length,
                    limit: limits.max_source_bytes,
                });
            }
        }
        let envelope = DurableEnvelope {
            version: 1,
            dialect: match self.dialect {
                Dialect::CommonMark => 0,
                Dialect::GitHubFlavored => 1,
            },
            max_source_bytes: self.limits.max_source_bytes,
            max_line_bytes: self.limits.max_line_bytes,
            max_events: self.limits.max_events,
            max_blocks: self.limits.max_blocks,
            max_nesting_depth: self.limits.max_nesting_depth,
            max_operations: self.limits.max_operations,
            before: self.before.to_string(),
            after: self.after.to_string(),
            operation_count: self.operation_count,
            operations: self.operations.as_deref().map(<[_]>::to_vec),
        };
        let json =
            serde_json::to_string(&envelope).map_err(|error| Error::InvalidPatchEnvelope {
                reason: error.to_string(),
            })?;
        if json.len() > limits.max_json_bytes {
            return Err(Error::PatchEnvelopeTooLarge {
                actual: json.len(),
                limit: limits.max_json_bytes,
            });
        }
        Ok(json)
    }
}

impl fmt::Debug for Patch {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Patch")
            .field("changed", &!self.is_empty())
            .field("dialect", &self.dialect)
            .finish_non_exhaustive()
    }
}

/// Resource policy for commit-coupled undo/redo history.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct HistoryLimits {
    /// Maximum retained commits across the undo and redo branches.
    pub max_entries: usize,
    /// Maximum exact before/after source bytes retained by all patches.
    pub max_patch_bytes: usize,
}

impl HistoryLimits {
    /// Conservative defaults for interactive editing.
    pub const DEFAULT: Self = Self {
        max_entries: 256,
        max_patch_bytes: 64 * 1024 * 1024,
    };
}

impl Default for HistoryLimits {
    fn default() -> Self {
        Self::DEFAULT
    }
}

/// Bounded undo/redo history coupled to validated Markdown commits.
#[derive(Debug)]
pub struct History {
    current: Snapshot,
    undo: Vec<Patch>,
    redo: Vec<Patch>,
    limits: HistoryLimits,
    retained_patch_bytes: usize,
}

impl History {
    /// Create an empty history around one immutable snapshot.
    ///
    /// # Errors
    ///
    /// Returns a typed limit error when either configured maximum is zero.
    pub fn new(current: Snapshot, limits: HistoryLimits) -> Result<Self, Error> {
        if limits.max_entries == 0 {
            return Err(Error::HistoryLimitExceeded {
                resource: "entry count",
                limit: 0,
            });
        }
        if limits.max_patch_bytes == 0 {
            return Err(Error::HistoryLimitExceeded {
                resource: "patch bytes",
                limit: 0,
            });
        }
        Ok(Self {
            current,
            undo: Vec::new(),
            redo: Vec::new(),
            limits,
            retained_patch_bytes: 0,
        })
    }

    /// Publish a commit whose exact before-image is the current snapshot.
    ///
    /// # Errors
    ///
    /// Returns a stale-source or history-limit error without changing history.
    pub fn apply(&mut self, commit: Commit) -> Result<(), Error> {
        if commit.patch.before != self.current.state.source {
            return Err(Error::PatchConflict);
        }
        if commit.patch.is_empty() {
            self.current = commit.snapshot;
            return Ok(());
        }
        let new_patch_bytes = patch_bytes(&commit.patch);
        let next_entries = self.undo.len().saturating_add(1);
        if next_entries > self.limits.max_entries {
            return Err(Error::HistoryLimitExceeded {
                resource: "entry count",
                limit: self.limits.max_entries,
            });
        }
        let released_redo: usize = self.redo.iter().map(patch_bytes).sum();
        let next_bytes = self
            .retained_patch_bytes
            .saturating_sub(released_redo)
            .saturating_add(new_patch_bytes);
        if next_bytes > self.limits.max_patch_bytes {
            return Err(Error::HistoryLimitExceeded {
                resource: "patch bytes",
                limit: self.limits.max_patch_bytes,
            });
        }
        self.undo
            .try_reserve(1)
            .map_err(|allocation_error| Error::Allocation {
                resource: "Markdown undo history",
                source: allocation_error,
            })?;
        self.redo.clear();
        self.retained_patch_bytes = next_bytes;
        self.undo.push(commit.patch);
        self.current = commit.snapshot;
        Ok(())
    }

    /// Current immutable snapshot.
    #[must_use]
    pub const fn current(&self) -> &Snapshot {
        &self.current
    }

    /// Redo one previously undone commit.
    ///
    /// # Errors
    ///
    /// Returns a typed patch or reparse error without partial publication.
    pub fn redo(&mut self) -> Result<bool, Error> {
        let Some(patch) = self.redo.last().cloned() else {
            return Ok(false);
        };
        let commit = self.current.apply(&patch)?;
        self.undo
            .try_reserve(1)
            .map_err(|allocation_error| Error::Allocation {
                resource: "Markdown undo history",
                source: allocation_error,
            })?;
        self.redo.pop();
        self.undo.push(patch);
        self.current = commit.snapshot;
        Ok(true)
    }

    /// Undo one committed edit through its exact inverse patch.
    ///
    /// # Errors
    ///
    /// Returns a typed patch or reparse error without partial publication.
    pub fn undo(&mut self) -> Result<bool, Error> {
        let Some(patch) = self.undo.last().cloned() else {
            return Ok(false);
        };
        let commit = self.current.apply(&patch.inverse())?;
        self.redo
            .try_reserve(1)
            .map_err(|allocation_error| Error::Allocation {
                resource: "Markdown redo history",
                source: allocation_error,
            })?;
        self.undo.pop();
        self.redo.push(patch);
        self.current = commit.snapshot;
        Ok(true)
    }
}

fn patch_bytes(patch: &Patch) -> usize {
    patch.before.len().saturating_add(patch.after.len())
}

fn ensure_common_base(left: &Patch, right: &Patch) -> Result<(), Error> {
    if left.before != right.before || left.dialect != right.dialect || left.limits != right.limits {
        return Err(Error::PatchConflict);
    }
    Ok(())
}

fn joined_operations(
    left: &Patch,
    right: &Patch,
) -> Result<(Vec<SemanticOperation>, ConflictSet), Error> {
    let left_operations =
        left.operations
            .as_deref()
            .ok_or_else(|| Error::InvalidPatchEnvelope {
                reason: "reverse-only patch has no joinable semantic operations".to_owned(),
            })?;
    let right_operations =
        right
            .operations
            .as_deref()
            .ok_or_else(|| Error::InvalidPatchEnvelope {
                reason: "reverse-only patch has no joinable semantic operations".to_owned(),
            })?;
    let combined = left_operations
        .len()
        .checked_add(right_operations.len())
        .ok_or(Error::OperationLimitExceeded {
            limit: left.limits.max_operations,
        })?;
    if combined > left.limits.max_operations {
        return Err(Error::OperationLimitExceeded {
            limit: left.limits.max_operations,
        });
    }
    let mut operations = Vec::new();
    operations
        .try_reserve_exact(combined)
        .map_err(|allocation_error| Error::Allocation {
            resource: "Markdown joined operations",
            source: allocation_error,
        })?;
    operations.extend_from_slice(left_operations);
    let mut conflicts = Vec::new();
    conflicts
        .try_reserve(right_operations.len())
        .map_err(|allocation_error| Error::Allocation {
            resource: "Markdown conflict set",
            source: allocation_error,
        })?;
    for candidate in right_operations {
        let Some(position) = candidate.position else {
            operations.push(candidate.clone());
            continue;
        };
        let existing = left_operations
            .iter()
            .find(|operation| operation.position == Some(position));
        match existing {
            Some(operation) if operation.replacement == candidate.replacement => {},
            Some(_) => conflicts.push(Conflict { position }),
            None => operations.push(candidate.clone()),
        }
    }
    conflicts.sort_unstable_by_key(|conflict| conflict.position);
    conflicts.dedup_by_key(|conflict| conflict.position);
    Ok((
        operations,
        ConflictSet {
            conflicts: conflicts.into_boxed_slice(),
        },
    ))
}

pub(crate) fn apply(source: &Snapshot, patch: &Patch) -> Result<Commit, Error> {
    if source.state.source != patch.before
        || source.dialect() != patch.dialect
        || source.limits() != patch.limits
    {
        return Err(Error::PatchConflict);
    }
    if patch.is_empty() {
        return Ok(Commit {
            snapshot: source.clone(),
            patch: patch.clone(),
            diagnostics: Diagnostics {
                changed: false,
                touched_blocks: 0,
                full_reparse_performed: false,
            },
        });
    }
    let snapshot = Snapshot::read_with(&patch.after, patch.dialect, patch.limits)?;
    Ok(Commit {
        snapshot,
        patch: patch.clone(),
        diagnostics: Diagnostics {
            changed: true,
            touched_blocks: patch.operation_count,
            full_reparse_performed: true,
        },
    })
}

fn fingerprint(source: &str) -> u64 {
    const OFFSET: u64 = 0xcbf2_9ce4_8422_2325;
    const PRIME: u64 = 0x0000_0100_0000_01b3;
    source.bytes().fold(OFFSET, |hash, byte| {
        (hash ^ u64::from(byte)).wrapping_mul(PRIME)
    })
}

fn copy_source(source: &str) -> Result<String, Error> {
    let mut retained = String::new();
    retained
        .try_reserve_exact(source.len())
        .map_err(|allocation_error| Error::Allocation {
            resource: "Markdown staged replacement",
            source: allocation_error,
        })?;
    retained.push_str(source);
    Ok(retained)
}

fn semantic_operations(intents: &[Intent]) -> Vec<SemanticOperation> {
    intents
        .iter()
        .map(|intent| match intent {
            Intent::Replace {
                position,
                replacement,
            } => SemanticOperation {
                position: Some(*position),
                replacement: replacement.clone(),
            },
            Intent::Append { block } => SemanticOperation {
                position: None,
                replacement: block.clone(),
            },
        })
        .collect()
}

fn replay_semantic_operations(
    source: &Snapshot,
    operations: &[SemanticOperation],
) -> Result<Commit, Error> {
    if operations.is_empty() {
        return Err(Error::InvalidPatchEnvelope {
            reason: "changed semantic patch has no operations".to_owned(),
        });
    }
    let mut edit = source.edit();
    for operation in operations {
        if let Some(position) = operation.position {
            edit.replace_block(position, &operation.replacement)?;
        } else {
            edit.append_block(&operation.replacement)?;
        }
    }
    edit.commit()
}

fn publish(source: &Snapshot, target: &str, intents: &[Intent]) -> Result<Commit, Error> {
    let operations = semantic_operations(intents);
    let operation_count = operations.len();
    if target == source.source() {
        let exact = Arc::clone(&source.state.source);
        let fingerprint = fingerprint(&exact);
        return Ok(Commit {
            snapshot: source.clone(),
            patch: Patch {
                before: Arc::clone(&exact),
                after: exact,
                dialect: source.dialect(),
                limits: source.limits(),
                source_fingerprint: fingerprint,
                target_fingerprint: fingerprint,
                operation_count,
                operations: Some(operations.clone().into_boxed_slice()),
            },
            diagnostics: Diagnostics {
                changed: false,
                touched_blocks: 0,
                full_reparse_performed: false,
            },
        });
    }
    let snapshot = Snapshot::read_with(target, source.dialect(), source.limits())?;
    let before = Arc::clone(&source.state.source);
    let after = Arc::clone(&snapshot.state.source);
    let source_fingerprint = fingerprint(&before);
    let target_fingerprint = fingerprint(&after);
    Ok(Commit {
        snapshot,
        patch: Patch {
            before,
            after,
            dialect: source.dialect(),
            limits: source.limits(),
            source_fingerprint,
            target_fingerprint,
            operation_count,
            operations: Some(operations.into_boxed_slice()),
        },
        diagnostics: Diagnostics {
            changed: true,
            touched_blocks: operation_count,
            full_reparse_performed: true,
        },
    })
}

fn render(source: &Snapshot, intents: &[Intent]) -> Result<String, Error> {
    let mut replacements: Vec<(usize, &str)> = intents
        .iter()
        .filter_map(|intent| match intent {
            Intent::Replace {
                position,
                replacement,
            } => Some((*position, replacement.as_str())),
            Intent::Append { .. } => None,
        })
        .collect();
    replacements.sort_unstable_by_key(|(position, _)| *position);
    let mut target = String::new();
    target
        .try_reserve_exact(source.source().len())
        .map_err(|allocation_error| Error::Allocation {
            resource: "Markdown edit candidate",
            source: allocation_error,
        })?;
    let mut cursor = 0usize;
    for (position, replacement) in replacements {
        let block = source
            .block(position)
            .ok_or(Error::BlockNotFound { position })?;
        let range = block.range();
        target.push_str(&source.source()[cursor..range.start]);
        target.push_str(replacement);
        if !replacement.is_empty()
            && range.end != source.source().len()
            && !replacement.ends_with(['\r', '\n'])
        {
            if block.source().ends_with("\r\n") {
                target.push_str("\r\n");
            } else if block.source().ends_with(['\r', '\n']) {
                target.push('\n');
            }
        }
        cursor = range.end;
    }
    target.push_str(&source.source()[cursor..]);
    for block in intents.iter().filter_map(|intent| match intent {
        Intent::Append { block } => Some(block.as_str()),
        Intent::Replace { .. } => None,
    }) {
        if !target.is_empty() {
            if target.ends_with('\n') {
                target.push('\n');
            } else {
                target.push_str("\n\n");
            }
        }
        target.push_str(block);
    }
    if target.len() > source.limits().max_source_bytes {
        return Err(Error::SourceTooLarge {
            actual: target.len(),
            limit: source.limits().max_source_bytes,
        });
    }
    Ok(target)
}

fn validate_replacement(source: &Snapshot, replacement: &str) -> Result<(), Error> {
    let parsed = Snapshot::read_with(replacement, source.dialect(), source.limits())?;
    let actual = parsed.blocks().len();
    if actual != 1 {
        return Err(Error::ReplacementBlockCount { actual });
    }
    Ok(())
}

fn validate_dependencies(source: &Snapshot, intents: &[Intent]) -> Result<(), Error> {
    for definition in source.references().filter(|reference| {
        matches!(
            reference.kind(),
            super::ReferenceKind::LinkDefinition | super::ReferenceKind::FootnoteDefinition
        )
    }) {
        let Some(label) = definition.label() else {
            continue;
        };
        let Some(definition_position) = containing_block(source, definition.range()) else {
            continue;
        };
        let Some(replacement) = replacement_at(intents, definition_position) else {
            continue;
        };
        if replacement_preserves_definition(source, replacement, definition.kind(), label)? {
            continue;
        }
        let dangling_use = source.references().any(|reference| {
            reference.label() == Some(label)
                && is_use_of(reference.kind(), definition.kind())
                && containing_block(source, reference.range())
                    .is_none_or(|position| replacement_at(intents, position).is_none())
        });
        if dangling_use {
            return Err(Error::ReferenceDependency {
                label: label.to_owned(),
            });
        }
    }
    Ok(())
}

fn containing_block(source: &Snapshot, range: std::ops::Range<usize>) -> Option<usize> {
    source.blocks().position(|block| {
        let block_range = block.range();
        block_range.start <= range.start && range.end <= block_range.end
    })
}

const fn is_use_of(kind: super::ReferenceKind, definition: super::ReferenceKind) -> bool {
    matches!(
        (kind, definition),
        (
            super::ReferenceKind::Link | super::ReferenceKind::Image,
            super::ReferenceKind::LinkDefinition
        ) | (
            super::ReferenceKind::Footnote,
            super::ReferenceKind::FootnoteDefinition
        )
    )
}

fn replacement_at(intents: &[Intent], position: usize) -> Option<&str> {
    intents.iter().find_map(|intent| match intent {
        Intent::Replace {
            position: candidate,
            replacement,
        } if *candidate == position => Some(replacement.as_str()),
        Intent::Replace { .. } | Intent::Append { .. } => None,
    })
}

fn replacement_preserves_definition(
    source: &Snapshot,
    replacement: &str,
    kind: super::ReferenceKind,
    label: &str,
) -> Result<bool, Error> {
    if replacement.is_empty() {
        return Ok(false);
    }
    let parsed = Snapshot::read_with(replacement, source.dialect(), source.limits())?;
    Ok(parsed
        .references()
        .any(|reference| reference.kind() == kind && reference.label() == Some(label)))
}
