//! Bounded, source-preserving edits for ordinary main-story DOC paragraphs.
//!
//! Length-changing replacements append Unicode text, rebuild the CLX and CHPX
//! FKPs, shift modeled main-story PLCFs, and update the FIB story length. The
//! transaction refuses structural text, tracked ranges, non-uniform character
//! formatting, interior position boundaries, and unmodeled dependencies before
//! publication. Text and direct-bold changes share one immutable
//! multi-operation transaction.

use crate::package::Error as PackageError;
use crate::tracked_revision::{Limits, RevisionEditor, RevisionKind};
use litchi_core::Position;
use litchi_core::patch::{
    BlobBundle, BlobId, CompositionError, JoinedSubEdits, PatchError, PatchLimits, PatchOperation,
    Reversible, ReversibleOperation, SubEdit,
};
pub use litchi_core::patch::{CompositionLimits, HistoryLimits, SubEditJoinFailure};
use std::collections::BTreeMap;
use std::io::Cursor;
use std::sync::Arc;

/// Bounded undo/redo history over immutable DOC snapshots.
pub type History = litchi_core::patch::History<Snapshot>;

/// Finite limits for one body transaction.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct TransactionLimits {
    operations: usize,
    replacement_units: usize,
    total_units: usize,
}

impl TransactionLimits {
    /// Creates explicit per-transaction bounds.
    #[must_use]
    pub const fn new(
        max_operations: usize,
        max_replacement_units: usize,
        max_total_replacement_units: usize,
    ) -> Self {
        Self {
            operations: max_operations,
            replacement_units: max_replacement_units,
            total_units: max_total_replacement_units,
        }
    }

    /// Maximum staged semantic operations.
    #[must_use]
    pub const fn max_operations(self) -> usize {
        self.operations
    }

    /// Maximum UTF-16 units in one replacement.
    #[must_use]
    pub const fn max_replacement_units(self) -> usize {
        self.replacement_units
    }

    /// Maximum aggregate UTF-16 replacement units.
    #[must_use]
    pub const fn max_total_replacement_units(self) -> usize {
        self.total_units
    }
}

impl Default for TransactionLimits {
    fn default() -> Self {
        Self::new(256, 1024 * 1024, 4 * 1024 * 1024)
    }
}

/// Main-story text visibility used for a review projection.
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq, Hash)]
pub enum Projection {
    /// Stored text, including both insertion and deletion redline text.
    #[default]
    All,
    /// Text visible after accepting insertion and deletion revisions.
    Accepted,
    /// Text visible after rejecting insertion and deletion revisions.
    Rejected,
}

/// A visible ordinary paragraph in the main document story.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Paragraph {
    position: Position,
    text: String,
}

impl Paragraph {
    /// Zero-based paragraph position in the selected projection.
    ///
    /// Constructing a [`Position`] is infallible. Resolving it against a
    /// snapshot collection is checked by [`Edit::replace_paragraph`].
    #[must_use]
    pub const fn position(&self) -> Position {
        self.position
    }

    /// Plain inert paragraph text, without the terminating paragraph mark.
    #[must_use]
    pub fn text(&self) -> &str {
        &self.text
    }
}

/// A reason why an edit is outside this intentionally small safe closure.
#[derive(Clone, Debug, PartialEq, Eq)]
#[non_exhaustive]
pub enum Refusal {
    /// The selected [`Position`] does not exist in the source body.
    ParagraphNotFound,
    /// Retained for source compatibility; length changes are now modeled by
    /// the bounded CLX/FKP transaction.
    LengthChange { expected: usize, actual: usize },
    /// The paragraph crosses pieces, which can have distinct encodings and PRMs.
    CrossesPieceBoundary,
    /// The selected paragraph is stored in an ANSI/compressed piece.
    CompressedPiece,
    /// Fields, object markers, cell markers, or other structural controls occur.
    StructuralContent,
    /// The paragraph intersects text-affecting tracked revisions.
    TrackedText,
    /// The requested replacement contains structural controls.
    ReplacementContainsStructuralContent,
    /// The source's review ranges overlap in a way this projection cannot prove.
    AmbiguousReviewRanges,
    /// An empty paragraph has no text run from which formatting can be copied.
    EmptyParagraph,
    /// Character formatting or another CP-bound dependency is not uniform or
    /// cannot be shifted without changing its meaning.
    FormattingDependency,
    /// A known CP-indexed table has coupled records outside the resize model.
    PositionDependency { fib_index: usize },
    /// The configured operation count was exhausted.
    OperationLimit { observed: usize, limit: usize },
    /// One or all replacement payloads exceed the configured UTF-16 bound.
    ReplacementLimit { observed: usize, limit: usize },
}

impl std::fmt::Display for Refusal {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::ParagraphNotFound => {
                formatter.write_str("body paragraph position is out of range")
            },
            Self::LengthChange { expected, actual } => write!(
                formatter,
                "replacement has {actual} UTF-16 units; this source paragraph has {expected}"
            ),
            Self::CrossesPieceBoundary => {
                formatter.write_str("body paragraph crosses DOC text pieces")
            },
            Self::CompressedPiece => {
                formatter.write_str("body paragraph is stored in a compressed DOC text piece")
            },
            Self::StructuralContent => {
                formatter.write_str("body paragraph contains DOC structural content")
            },
            Self::TrackedText => formatter.write_str("body paragraph intersects tracked text"),
            Self::ReplacementContainsStructuralContent => {
                formatter.write_str("replacement contains DOC structural content")
            },
            Self::AmbiguousReviewRanges => {
                formatter.write_str("tracked revision ranges overlap ambiguously")
            },
            Self::EmptyParagraph => {
                formatter.write_str("empty body paragraph has no editable text formatting run")
            },
            Self::FormattingDependency => formatter.write_str(
                "body paragraph has character formatting or CP-bound dependencies that cannot be resized losslessly",
            ),
            Self::PositionDependency { fib_index } => write!(
                formatter,
                "body length change depends on unmodeled CP-indexed FIB table {fib_index}"
            ),
            Self::OperationLimit { observed, limit } => write!(
                formatter,
                "body transaction requested {observed} operations; limit is {limit}"
            ),
            Self::ReplacementLimit { observed, limit } => write!(
                formatter,
                "body transaction requested {observed} UTF-16 replacement units; limit is {limit}"
            ),
        }
    }
}

/// Failure from a body-text transaction or source-checked patch.
#[derive(Debug)]
#[non_exhaustive]
pub enum Error {
    /// The DOC/CFB source or its required invariants is invalid.
    Invalid(PackageError),
    /// The request is valid in general but unsafe for this preservation seam.
    Refused(Refusal),
    /// A patch was presented with any snapshot other than its exact source.
    Conflict,
    /// A shared bounded composition rejected staged work.
    Composition(CompositionError),
    /// A durable semantic patch is malformed or exceeds its explicit limits.
    Durable(PatchError),
}

impl std::fmt::Display for Error {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Invalid(error) => error.fmt(formatter),
            Self::Refused(refusal) => refusal.fmt(formatter),
            Self::Conflict => formatter.write_str("body-text patch source conflict"),
            Self::Composition(error) => error.fmt(formatter),
            Self::Durable(error) => error.fmt(formatter),
        }
    }
}

impl std::error::Error for Error {
    fn source(&self) -> Option<&(dyn std::error::Error + 'static)> {
        match self {
            Self::Invalid(error) => Some(error),
            Self::Composition(error) => Some(error),
            Self::Durable(error) => Some(error),
            Self::Refused(_) | Self::Conflict => None,
        }
    }
}

impl From<CompositionError> for Error {
    fn from(error: CompositionError) -> Self {
        Self::Composition(error)
    }
}

impl From<PatchError> for Error {
    fn from(error: PatchError) -> Self {
        Self::Durable(error)
    }
}

/// Immutable, exact-source snapshot for the body-text transaction seam.
#[derive(Clone)]
pub struct Snapshot {
    source: Arc<[u8]>,
    limits: Limits,
    transaction_limits: TransactionLimits,
}

impl Snapshot {
    /// Opens an owned Word 97+ DOC source after validating its safe edit basis.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Invalid`] when the CFB, Word 97+ FIB, selected table
    /// stream, piece table, or FKP basis cannot support safe editing.
    pub fn open(input: impl Into<Vec<u8>>, limits: Limits) -> Result<Self> {
        Self::open_bounded(input, limits, TransactionLimits::default())
    }

    /// Opens an owned DOC source with explicit package and transaction bounds.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Invalid`] when the complete package or its safe edit
    /// basis cannot be reopened.
    pub fn open_bounded(
        input: impl Into<Vec<u8>>,
        limits: Limits,
        transaction_limits: TransactionLimits,
    ) -> Result<Self> {
        let bytes = input.into();
        RevisionEditor::open(bytes.clone(), limits).map_err(Error::Invalid)?;
        let mut package =
            crate::Package::from_reader(Cursor::new(bytes.clone())).map_err(Error::Invalid)?;
        package.document().map_err(Error::Invalid)?;
        Ok(Self {
            source: Arc::from(bytes.into_boxed_slice()),
            limits,
            transaction_limits,
        })
    }

    /// Parses a borrowed DOC source with the default resource limits.
    ///
    /// # Errors
    ///
    /// Returns the same validation failure as [`Self::open`].
    pub fn parse(input: &[u8]) -> Result<Self> {
        Self::open(input.to_vec(), Limits::default())
    }

    /// Opens an owned DOC source with the default resource limits.
    ///
    /// # Errors
    ///
    /// Returns the same validation failure as [`Self::open`].
    pub fn from_bytes(input: Vec<u8>) -> Result<Self> {
        Self::open(input, Limits::default())
    }

    /// Exact CFB source bytes retained for source checks and byte-exact no-ops.
    #[must_use]
    pub fn bytes(&self) -> &[u8] {
        &self.source
    }

    /// Shared ownership of the exact source allocation.
    #[must_use]
    pub fn bytes_shared(&self) -> Arc<[u8]> {
        Arc::clone(&self.source)
    }

    /// Stable first-stage fingerprint for diagnostics and stale-source checks.
    #[must_use]
    pub fn fingerprint(&self) -> u64 {
        fingerprint(&self.source)
    }

    /// Configured finite semantic-operation and text-payload bounds.
    #[must_use]
    pub const fn transaction_limits(&self) -> TransactionLimits {
        self.transaction_limits
    }

    /// Lists ordinary source-body paragraphs under the requested review projection.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Invalid`] for an invalid source or [`Error::Refused`]
    /// when tracked text ranges overlap ambiguously for the projection.
    pub fn paragraphs(&self, projection: Projection) -> Result<Vec<Paragraph>> {
        let editor = self.editor()?;
        projected_paragraphs(&editor, projection)
    }

    /// Starts a staged bounded body text-and-formatting transaction.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Invalid`] if the retained source no longer validates.
    pub fn edit(&self) -> Result<Edit> {
        Edit::new(self.clone())
    }

    /// Alias for [`Self::edit`].
    ///
    /// # Errors
    ///
    /// Returns the same failure as [`Self::edit`].
    pub fn transaction(&self) -> Result<Edit> {
        self.edit()
    }

    /// Starts a bounded disjoint composition for this exact artifact.
    #[must_use]
    pub fn compose(&self, limits: CompositionLimits) -> Composition {
        Composition {
            source: self.clone(),
            joined: JoinedSubEdits::new(self.lineage(), limits),
        }
    }

    /// Starts explicit bounded undo/redo history at this immutable snapshot.
    #[must_use]
    pub fn history(&self, limits: HistoryLimits) -> History {
        History::new(self.clone(), limits)
    }

    /// Prepares one independently composable paragraph replacement against
    /// this exact immutable artifact.
    ///
    /// # Errors
    ///
    /// Returns the same typed validation refusal as [`Edit::replace_paragraph`]
    /// or a common composition-bound error.
    pub fn prepare_replace(
        &self,
        limits: CompositionLimits,
        identifier: impl Into<String>,
        position: Position,
        replacement: impl Into<String>,
    ) -> Result<PreparedEdit> {
        let replacement = replacement.into();
        let mut validation = self.edit()?;
        validation.replace_paragraph(position, &replacement)?;
        PreparedEdit::new(
            self.lineage(),
            limits,
            identifier,
            PreparedOperation::Text {
                position,
                replacement,
            },
        )
    }

    /// Prepares one independently composable direct-bold change.
    ///
    /// # Errors
    ///
    /// Returns the same typed validation refusal as
    /// [`Edit::set_paragraph_bold`] or a common composition-bound error.
    pub fn prepare_bold(
        &self,
        limits: CompositionLimits,
        identifier: impl Into<String>,
        position: Position,
        enabled: bool,
    ) -> Result<PreparedEdit> {
        let mut validation = self.edit()?;
        validation.set_paragraph_bold(position, enabled)?;
        PreparedEdit::new(
            self.lineage(),
            limits,
            identifier,
            PreparedOperation::Bold { position, enabled },
        )
    }

    /// Applies supported durable body-text operations to this exact artifact.
    ///
    /// # Errors
    ///
    /// Returns an error for a foreign vocabulary, malformed selector,
    /// artifact or semantic precondition conflict, bound violation, or typed
    /// body-edit refusal.
    pub fn apply_durable<Mode>(&self, patch: &litchi_core::patch::Patch<Mode>) -> Result<Self> {
        if patch.format() != "litchi-doc-body" || !patch.blobs().is_empty() {
            return Err(invalid_durable_patch("unsupported format or blob bundle"));
        }
        if patch.operations().is_empty() {
            return Ok(self.clone());
        }
        let expected_artifact = BlobId::of(self.bytes()).as_hex();
        let mut edit = self.edit()?;
        for operation in patch.operations() {
            if operation.preconditions.len() != 2 {
                return Err(invalid_durable_patch(
                    "body operation must have exactly two preconditions",
                ));
            }
            let artifact = operation
                .preconditions
                .get("artifact_sha256")
                .and_then(serde_json::Value::as_str)
                .ok_or_else(|| invalid_durable_patch("missing artifact hash precondition"))?;
            if artifact != expected_artifact {
                return Err(Error::Conflict);
            }
            let position = parse_durable_target(&operation.target)?;
            match operation.op.as_str() {
                "body-text.set" => {
                    let expected = operation
                        .preconditions
                        .get("text")
                        .and_then(serde_json::Value::as_str)
                        .ok_or_else(|| invalid_durable_patch("missing text precondition"))?;
                    let paragraphs = source_paragraphs(&edit.editor)?;
                    let paragraph = paragraphs
                        .get(position.get())
                        .ok_or(Error::Refused(Refusal::ParagraphNotFound))?;
                    if paragraph.text != expected {
                        return Err(Error::Conflict);
                    }
                    let replacement = operation
                        .value
                        .as_str()
                        .ok_or_else(|| invalid_durable_patch("body text value is not a string"))?;
                    edit.replace_paragraph(position, replacement)?;
                },
                "body-bold.set" => {
                    let expected = parse_optional_bool(
                        operation
                            .preconditions
                            .get("bold")
                            .ok_or_else(|| invalid_durable_patch("missing bold precondition"))?,
                    )?;
                    let paragraphs = source_paragraphs(&edit.editor)?;
                    let paragraph = paragraphs
                        .get(position.get())
                        .ok_or(Error::Refused(Refusal::ParagraphNotFound))?;
                    let actual = edit
                        .editor
                        .uniform_bold_override(paragraph.start_cp, paragraph.end_cp)
                        .map_err(Error::Invalid)?
                        .ok_or(Error::Refused(Refusal::FormattingDependency))?;
                    if actual != expected {
                        return Err(Error::Conflict);
                    }
                    let value = parse_optional_bool(&operation.value)?;
                    edit.set_paragraph_bold_override(position, value)?;
                },
                _ => return Err(invalid_durable_patch("unsupported operation vocabulary")),
            }
        }
        edit.commit().map(|commit| commit.snapshot)
    }

    /// Exact source bytes. A snapshot has no implicit serialization step.
    #[must_use]
    pub fn finish(&self) -> Vec<u8> {
        self.source.as_ref().to_vec()
    }

    fn editor(&self) -> Result<RevisionEditor> {
        RevisionEditor::open(self.source.as_ref().to_vec(), self.limits).map_err(Error::Invalid)
    }

    fn lineage(&self) -> Lineage {
        Lineage(Arc::clone(&self.source))
    }
}

impl std::fmt::Debug for Snapshot {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("Snapshot")
            .field("bytes", &self.source.len())
            .field("fingerprint", &self.fingerprint())
            .field("limits", &self.limits)
            .field("transaction_limits", &self.transaction_limits)
            .finish()
    }
}

impl PartialEq for Snapshot {
    fn eq(&self, other: &Self) -> bool {
        self.source == other.source
    }
}

impl Eq for Snapshot {}

/// Clone-first staged text edit over one immutable source snapshot.
pub struct Edit {
    source: Snapshot,
    editor: RevisionEditor,
    changes: Vec<Change>,
    replacement_units: usize,
}

impl Edit {
    fn new(source: Snapshot) -> Result<Self> {
        let editor = source.editor()?;
        Ok(Self {
            source,
            editor,
            changes: Vec::new(),
            replacement_units: 0,
        })
    }

    /// Immutable source snapshot that authorizes this transaction.
    #[must_use]
    pub const fn source(&self) -> &Snapshot {
        &self.source
    }

    /// Replaces text in one ordinary source-body paragraph.
    ///
    /// The replacement can change UTF-16 length. Publication appends a Unicode
    /// piece, rebuilds CHPX FKPs and the CLX, shifts modeled main-story PLCFs,
    /// and updates the FIB story length.
    ///
    /// `position` is a format-neutral [`Position`]; its membership in this
    /// source body is checked here and an absent paragraph is reported as
    /// [`Refusal::ParagraphNotFound`].
    ///
    /// # Errors
    ///
    /// Returns a typed [`Refusal`] for every operation outside the proven
    /// bounded dependency closure and [`Error::Invalid`] for a failed package
    /// update.
    pub fn replace_paragraph(&mut self, position: Position, replacement: &str) -> Result<()> {
        let paragraphs = source_paragraphs(&self.editor)?;
        let paragraph = paragraphs
            .get(position.get())
            .ok_or(Error::Refused(Refusal::ParagraphNotFound))?;
        if has_structural_content(&paragraph.text) {
            return Err(Error::Refused(Refusal::StructuralContent));
        }
        if has_structural_content(replacement) {
            return Err(Error::Refused(
                Refusal::ReplacementContainsStructuralContent,
            ));
        }
        let actual = replacement.encode_utf16().count();
        self.ensure_replacement_capacity(actual)?;
        if paragraph.start_cp == paragraph.end_cp {
            return Err(Error::Refused(Refusal::EmptyParagraph));
        }
        if text_revision_intersects(&self.editor, paragraph.start_cp, paragraph.end_cp)? {
            return Err(Error::Refused(Refusal::TrackedText));
        }
        if paragraph.text == replacement {
            return Ok(());
        }
        if actual != paragraph.text.encode_utf16().count()
            && let Some(&fib_index) = self.editor.unmodeled_length_dependencies().first()
        {
            return Err(Error::Refused(Refusal::PositionDependency { fib_index }));
        }
        self.ensure_operation_capacity()?;
        if !self
            .editor
            .has_uniform_character_format(paragraph.start_cp, paragraph.end_cp)
            .map_err(Error::Invalid)?
        {
            return Err(Error::Refused(Refusal::FormattingDependency));
        }
        let before = paragraph.text.clone();
        self.editor
            .replace_plain_text(paragraph.start_cp, paragraph.end_cp, replacement)
            .map_err(Error::Invalid)?;
        self.replacement_units = self.replacement_units.saturating_add(actual);
        self.changes.push(Change::Text {
            position,
            before,
            after: replacement.to_string(),
        });
        Ok(())
    }

    /// Sets a direct bold override for one ordinary body paragraph in the
    /// same transaction as text replacements.
    ///
    /// # Errors
    ///
    /// Returns a typed refusal when the paragraph is absent, empty,
    /// structural, tracked, or has non-uniform direct bold semantics.
    pub fn set_paragraph_bold(&mut self, position: Position, enabled: bool) -> Result<()> {
        self.set_paragraph_bold_override(position, Some(enabled))
    }

    /// Discards staged changes and returns the original immutable snapshot.
    #[must_use]
    pub fn rollback(self) -> Snapshot {
        self.source
    }

    /// Publishes a validated snapshot and its reversible source-checked patch.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Invalid`] when the rendered candidate cannot be
    /// reopened with the original safety limits.
    pub fn commit(self) -> Result<Commit> {
        let bytes = self.editor.finish().map_err(Error::Invalid)?;
        let snapshot = if bytes == self.source.bytes() {
            self.source.clone()
        } else {
            Snapshot::open_bounded(bytes, self.source.limits, self.source.transaction_limits)?
        };
        let patch = Patch::new(self.source, snapshot.clone(), self.changes);
        Ok(Commit { snapshot, patch })
    }

    fn ensure_operation_capacity(&self) -> Result<()> {
        let observed = self.changes.len().saturating_add(1);
        let limit = self.source.transaction_limits.operations;
        if observed > limit {
            Err(Error::Refused(Refusal::OperationLimit { observed, limit }))
        } else {
            Ok(())
        }
    }

    fn ensure_replacement_capacity(&self, units: usize) -> Result<()> {
        let per_value = self.source.transaction_limits.replacement_units;
        if units > per_value {
            return Err(Error::Refused(Refusal::ReplacementLimit {
                observed: units,
                limit: per_value,
            }));
        }
        let total = self.replacement_units.saturating_add(units);
        let total_limit = self.source.transaction_limits.total_units;
        if total > total_limit {
            return Err(Error::Refused(Refusal::ReplacementLimit {
                observed: total,
                limit: total_limit,
            }));
        }
        Ok(())
    }

    fn set_paragraph_bold_override(
        &mut self,
        position: Position,
        value: Option<bool>,
    ) -> Result<()> {
        let paragraphs = source_paragraphs(&self.editor)?;
        let paragraph = paragraphs
            .get(position.get())
            .ok_or(Error::Refused(Refusal::ParagraphNotFound))?;
        if paragraph.start_cp == paragraph.end_cp {
            return Err(Error::Refused(Refusal::EmptyParagraph));
        }
        if has_structural_content(&paragraph.text) {
            return Err(Error::Refused(Refusal::StructuralContent));
        }
        if text_revision_intersects(&self.editor, paragraph.start_cp, paragraph.end_cp)? {
            return Err(Error::Refused(Refusal::TrackedText));
        }
        let before = self
            .editor
            .uniform_bold_override(paragraph.start_cp, paragraph.end_cp)
            .map_err(Error::Invalid)?
            .ok_or(Error::Refused(Refusal::FormattingDependency))?;
        if before == value {
            return Ok(());
        }
        self.ensure_operation_capacity()?;
        self.editor
            .set_character_bold_override(paragraph.start_cp, paragraph.end_cp, value)
            .map_err(Error::Invalid)?;
        self.changes.push(Change::Bold {
            position,
            before,
            after: value,
        });
        Ok(())
    }
}

/// Validated commit result for one body-text transaction.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Commit {
    snapshot: Snapshot,
    patch: Patch,
}

impl Commit {
    /// Whether a DOC byte changed.
    #[must_use]
    pub fn changed(&self) -> bool {
        !self.patch.is_noop()
    }

    /// Published post-edit snapshot.
    #[must_use]
    pub const fn snapshot(&self) -> &Snapshot {
        &self.snapshot
    }

    /// Reversible exact-source patch.
    #[must_use]
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Splits a commit into its snapshot and patch.
    #[must_use]
    pub fn into_parts(self) -> (Snapshot, Patch) {
        (self.snapshot, self.patch)
    }
}

/// In-memory reversible replacement guarded by exact source bytes.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Patch {
    before: Snapshot,
    after: Snapshot,
    before_fingerprint: u64,
    after_fingerprint: u64,
    changes: Vec<Change>,
}

impl Patch {
    fn new(before: Snapshot, after: Snapshot, changes: Vec<Change>) -> Self {
        Self {
            before_fingerprint: before.fingerprint(),
            after_fingerprint: after.fingerprint(),
            before,
            after,
            changes,
        }
    }

    /// Semantic text and formatting changes in application order.
    #[must_use]
    pub fn changes(&self) -> impl ExactSizeIterator<Item = ChangeRef<'_>> {
        self.changes.iter().map(Change::as_ref)
    }

    /// Exact source snapshot required for application.
    #[must_use]
    pub const fn before(&self) -> &Snapshot {
        &self.before
    }

    /// Snapshot produced by the transaction.
    #[must_use]
    pub const fn after(&self) -> &Snapshot {
        &self.after
    }

    /// Fast stale-source precheck; exact bytes remain authoritative.
    #[must_use]
    pub const fn source_fingerprint(&self) -> u64 {
        self.before_fingerprint
    }

    /// Target diagnostic fingerprint.
    #[must_use]
    pub const fn target_fingerprint(&self) -> u64 {
        self.after_fingerprint
    }

    /// Whether this patch preserves the exact artifact.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.before.bytes() == self.after.bytes()
    }

    /// Applies only to the exact source artifact.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Conflict`] unless `source` has byte-for-byte equality
    /// with this patch's captured source snapshot.
    pub fn apply(&self, source: &Snapshot) -> Result<Snapshot> {
        if source.fingerprint() != self.before_fingerprint || source.bytes() != self.before.bytes()
        {
            return Err(Error::Conflict);
        }
        Ok(if self.is_noop() {
            source.clone()
        } else {
            self.after.clone()
        })
    }

    /// Exact inverse patch.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            before: self.after.clone(),
            after: self.before.clone(),
            before_fingerprint: self.after_fingerprint,
            after_fingerprint: self.before_fingerprint,
            changes: self.changes.iter().rev().map(Change::inverse).collect(),
        }
    }

    /// Converts this exact-source patch to the shared bounded deterministic
    /// semantic envelope.
    ///
    /// # Errors
    ///
    /// Returns [`PatchError`] when the requested wire limits cannot represent
    /// every semantic operation and inverse.
    pub fn to_durable(
        &self,
        limits: PatchLimits,
    ) -> std::result::Result<litchi_core::patch::Patch<Reversible>, PatchError> {
        let before_artifact = BlobId::of(self.before.bytes()).as_hex();
        let after_artifact = BlobId::of(self.after.bytes()).as_hex();
        let operations = self
            .changes
            .iter()
            .map(|change| {
                let forward = durable_operation(limits, change, &before_artifact)?;
                let inverse = durable_operation(limits, &change.inverse(), &after_artifact)?;
                Ok(ReversibleOperation::new(forward, inverse))
            })
            .collect::<std::result::Result<Vec<_>, PatchError>>()?;
        litchi_core::patch::Patch::<Reversible>::new(
            limits,
            "litchi-doc-body",
            operations,
            BlobBundle::new(limits.blobs()),
            BlobBundle::new(limits.blobs()),
        )
    }
}

/// Borrowed semantic change carried by an in-memory patch.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
#[non_exhaustive]
pub enum ChangeRef<'a> {
    /// One ordinary paragraph text replacement.
    Text {
        position: Position,
        before: &'a str,
        after: &'a str,
    },
    /// One direct bold-override mutation.
    Bold {
        position: Position,
        before: Option<bool>,
        after: Option<bool>,
    },
}

#[derive(Clone, Debug, PartialEq, Eq)]
enum Change {
    Text {
        position: Position,
        before: String,
        after: String,
    },
    Bold {
        position: Position,
        before: Option<bool>,
        after: Option<bool>,
    },
}

impl Change {
    fn as_ref(&self) -> ChangeRef<'_> {
        match self {
            Self::Text {
                position,
                before,
                after,
            } => ChangeRef::Text {
                position: *position,
                before,
                after,
            },
            Self::Bold {
                position,
                before,
                after,
            } => ChangeRef::Bold {
                position: *position,
                before: *before,
                after: *after,
            },
        }
    }

    fn inverse(&self) -> Self {
        match self {
            Self::Text {
                position,
                before,
                after,
            } => Self::Text {
                position: *position,
                before: after.clone(),
                after: before.clone(),
            },
            Self::Bold {
                position,
                before,
                after,
            } => Self::Bold {
                position: *position,
                before: *after,
                after: *before,
            },
        }
    }
}

#[derive(Clone, PartialEq, Eq)]
struct Lineage(Arc<[u8]>);

#[derive(Clone, Debug, PartialEq, Eq)]
enum PreparedOperation {
    Text {
        position: Position,
        replacement: String,
    },
    Bold {
        position: Position,
        enabled: bool,
    },
}

impl PreparedOperation {
    fn position(&self) -> Position {
        match self {
            Self::Text { position, .. } | Self::Bold { position, .. } => *position,
        }
    }

    fn effect(&self) -> String {
        let facet = match self {
            Self::Text { .. } => "text",
            Self::Bold { .. } => "bold",
        };
        format!("body/paragraph:{}/{facet}", self.position().get())
    }
}

/// One independently prepared body-text or formatting edit.
pub struct PreparedEdit {
    inner: SubEdit<Lineage, PreparedOperation>,
}

impl PreparedEdit {
    fn new(
        lineage: Lineage,
        limits: CompositionLimits,
        identifier: impl Into<String>,
        operation: PreparedOperation,
    ) -> Result<Self> {
        let writes = [operation.effect()];
        let inner = SubEdit::new(
            lineage,
            limits,
            identifier,
            std::iter::empty(),
            writes,
            operation,
        )?;
        Ok(Self { inner })
    }

    /// Stable caller-selected composition identifier.
    #[must_use]
    pub fn identifier(&self) -> &str {
        self.inner.id()
    }

    /// Target paragraph position in the immutable source collection.
    #[must_use]
    pub fn position(&self) -> Position {
        self.inner.payload().position()
    }
}

impl std::fmt::Debug for PreparedEdit {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("PreparedEdit")
            .field("identifier", &self.identifier())
            .field("position", &self.position())
            .finish_non_exhaustive()
    }
}

/// Recoverable failure to join one independently prepared edit.
pub struct JoinError {
    failure: SubEditJoinFailure,
    rejected: PreparedEdit,
}

impl JoinError {
    /// Structured common composition refusal.
    #[must_use]
    pub const fn failure(&self) -> &SubEditJoinFailure {
        &self.failure
    }

    /// Recovers the rejected prepared edit.
    #[must_use]
    pub fn into_rejected(self) -> PreparedEdit {
        self.rejected
    }
}

impl std::fmt::Debug for JoinError {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("JoinError")
            .field("failure", &self.failure)
            .field("rejected", &self.rejected)
            .finish()
    }
}

/// Bounded deterministic composition of provably disjoint body edits.
pub struct Composition {
    source: Snapshot,
    joined: JoinedSubEdits<Lineage, PreparedOperation>,
}

impl Composition {
    /// Joins one edit when its exact artifact lineage and semantic facet are
    /// disjoint from every accepted edit.
    ///
    /// # Errors
    ///
    /// Returns a structured common composition refusal while retaining the
    /// rejected work.
    #[expect(
        clippy::result_large_err,
        reason = "composition refusals intentionally return the rejected prepared edit to the caller"
    )]
    pub fn join(&mut self, incoming: PreparedEdit) -> std::result::Result<&mut Self, JoinError> {
        if let Err(error) = self.joined.join(incoming.inner) {
            let (failure, rejected) = error.into_parts();
            return Err(JoinError {
                failure,
                rejected: PreparedEdit { inner: rejected },
            });
        }
        Ok(self)
    }

    /// Number of accepted independently prepared edits.
    #[must_use]
    pub fn len(&self) -> usize {
        self.joined.len()
    }

    /// Whether no edits have been accepted.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.joined.is_empty()
    }

    /// Commits all accepted edits atomically in stable identifier order.
    ///
    /// # Errors
    ///
    /// Returns the same staging, publication, and full-reopen errors as an
    /// ordinary multi-operation [`Edit`].
    pub fn commit(self) -> Result<Commit> {
        let mut edit = self.source.edit()?;
        for prepared in self.joined.into_sub_edits() {
            match prepared.into_payload() {
                PreparedOperation::Text {
                    position,
                    replacement,
                } => edit.replace_paragraph(position, &replacement)?,
                PreparedOperation::Bold { position, enabled } => {
                    edit.set_paragraph_bold(position, enabled)?;
                },
            }
        }
        edit.commit()
    }
}

impl std::fmt::Debug for Composition {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("Composition")
            .field("edits", &self.joined.len())
            .finish_non_exhaustive()
    }
}

fn durable_operation(
    limits: PatchLimits,
    change: &Change,
    artifact: &str,
) -> std::result::Result<PatchOperation, PatchError> {
    let mut preconditions = BTreeMap::new();
    preconditions.insert(
        "artifact_sha256".to_string(),
        serde_json::Value::String(artifact.to_string()),
    );
    let (op, position, value) = match change {
        Change::Text {
            position,
            before,
            after,
        } => {
            preconditions.insert(
                "text".to_string(),
                serde_json::Value::String(before.clone()),
            );
            (
                "body-text.set",
                *position,
                serde_json::Value::String(after.clone()),
            )
        },
        Change::Bold {
            position,
            before,
            after,
        } => {
            preconditions.insert("bold".to_string(), optional_bool_value(*before));
            ("body-bold.set", *position, optional_bool_value(*after))
        },
    };
    PatchOperation::new(
        limits,
        op,
        format!("body/paragraph:{}", position.get()),
        preconditions,
        value,
    )
}

fn optional_bool_value(value: Option<bool>) -> serde_json::Value {
    value.map_or(serde_json::Value::Null, serde_json::Value::Bool)
}

fn parse_optional_bool(value: &serde_json::Value) -> Result<Option<bool>> {
    match value {
        serde_json::Value::Null => Ok(None),
        serde_json::Value::Bool(value) => Ok(Some(*value)),
        _ => Err(invalid_durable_patch(
            "bold value must be a Boolean or null",
        )),
    }
}

fn parse_durable_target(target: &str) -> Result<Position> {
    let position = target
        .strip_prefix("body/paragraph:")
        .filter(|value| {
            !value.is_empty()
                && value.bytes().all(|byte| byte.is_ascii_digit())
                && (*value == "0" || !value.starts_with('0'))
        })
        .ok_or_else(|| invalid_durable_patch("invalid body paragraph target"))?;
    position
        .parse::<usize>()
        .map(Position::new)
        .map_err(|_error| invalid_durable_patch("body paragraph position exceeds this platform"))
}

fn invalid_durable_patch(message: &str) -> Error {
    Error::Invalid(PackageError::InvalidFormat(format!(
        "invalid DOC body durable patch: {message}"
    )))
}

type Result<T> = std::result::Result<T, Error>;

#[derive(Clone)]
struct SourceParagraph {
    start_cp: u32,
    end_cp: u32,
    text: String,
}

fn source_paragraphs(editor: &RevisionEditor) -> Result<Vec<SourceParagraph>> {
    let text = editor.main_story_text().map_err(Error::Invalid)?;
    let mut output = Vec::new();
    let mut start_cp = 0u32;
    let mut start_byte = 0usize;
    let mut cp = 0u32;
    for (byte, character) in text.char_indices() {
        let width = if character.len_utf16() == 1 { 1 } else { 2 };
        let next_cp = cp.checked_add(width).ok_or_else(|| {
            Error::Invalid(PackageError::Corrupted(
                "main-story CP overflow".to_string(),
            ))
        })?;
        if character == '\r' {
            if !editor.is_in_table_at_cp(cp).map_err(Error::Invalid)? {
                output.push(SourceParagraph {
                    start_cp,
                    end_cp: cp,
                    text: text[start_byte..byte].to_string(),
                });
            }
            start_cp = next_cp;
            start_byte = byte + character.len_utf8();
        } else if character == '\u{7}' {
            // A table cell marker is never an ordinary body paragraph.
            start_cp = next_cp;
            start_byte = byte + character.len_utf8();
        }
        cp = next_cp;
    }
    if cp != editor.main_story_cp_len() {
        return Err(Error::Invalid(PackageError::Corrupted(
            "decoded main story has an inconsistent CP length".to_string(),
        )));
    }
    Ok(output)
}

fn projected_paragraphs(editor: &RevisionEditor, projection: Projection) -> Result<Vec<Paragraph>> {
    let source = source_paragraphs(editor)?;
    if projection == Projection::All {
        return Ok(source
            .into_iter()
            .enumerate()
            .map(|(position, paragraph)| Paragraph {
                position: Position::new(position),
                text: paragraph.text,
            })
            .collect());
    }
    let hidden = hidden_ranges(editor, projection)?;
    let mut output = Vec::new();
    for paragraph in source {
        let text = project_text(&paragraph, &hidden)?;
        output.push(Paragraph {
            position: Position::new(output.len()),
            text,
        });
    }
    Ok(output)
}

fn hidden_ranges(editor: &RevisionEditor, projection: Projection) -> Result<Vec<(u32, u32)>> {
    let mut ranges = Vec::new();
    for revision in editor.revisions().map_err(Error::Invalid)? {
        let hide = matches!(
            (projection, revision.kind),
            (
                Projection::Accepted,
                RevisionKind::Deletion | RevisionKind::MoveFrom
            ) | (
                Projection::Rejected,
                RevisionKind::Insertion | RevisionKind::MoveTo
            )
        );
        if hide && revision.start_cp < revision.end_cp {
            ranges.push((revision.start_cp, revision.end_cp));
        }
    }
    ranges.sort_unstable();
    if ranges.windows(2).any(|pair| pair[0].1 > pair[1].0) {
        return Err(Error::Refused(Refusal::AmbiguousReviewRanges));
    }
    Ok(ranges)
}

fn project_text(paragraph: &SourceParagraph, hidden: &[(u32, u32)]) -> Result<String> {
    let mut output = String::new();
    let mut cp = paragraph.start_cp;
    for character in paragraph.text.chars() {
        let width = if character.len_utf16() == 1 { 1 } else { 2 };
        let end = cp.checked_add(width).ok_or_else(|| {
            Error::Invalid(PackageError::Corrupted(
                "projection CP overflow".to_string(),
            ))
        })?;
        if !hidden
            .iter()
            .any(|(start, finish)| *start < end && cp < *finish)
        {
            output.push(character);
        }
        cp = end;
    }
    Ok(output)
}

fn text_revision_intersects(editor: &RevisionEditor, start: u32, end: u32) -> Result<bool> {
    Ok(editor
        .revisions()
        .map_err(Error::Invalid)?
        .into_iter()
        .any(|revision| {
            matches!(
                revision.kind,
                RevisionKind::Insertion
                    | RevisionKind::Deletion
                    | RevisionKind::MoveFrom
                    | RevisionKind::MoveTo
            ) && revision.start_cp < end
                && start < revision.end_cp
        }))
}

fn has_structural_content(text: &str) -> bool {
    text.chars().any(|character| {
        matches!(character, '\r' | '\u{7}' | '\u{13}'..='\u{15}' | '\u{fffc}')
            || (character.is_control() && character != '\t')
    })
}

fn fingerprint(bytes: &[u8]) -> u64 {
    let mut value = 0xcbf2_9ce4_8422_2325u64;
    for byte in bytes {
        value ^= u64::from(*byte);
        value = value.wrapping_mul(0x0000_0100_0000_01b3);
    }
    value
}

#[cfg(test)]
mod tests {
    use super::{Error, Projection, Refusal, Snapshot, TransactionLimits};
    use crate::tracked_revision::Limits;
    use crate::writer::{CharacterFormatting, ParagraphFormatting, TextRevision, Writer};
    use litchi_core::Position;
    use litchi_core::patch::{
        BlobLimits, CompositionLimits, HistoryLimits, Patch, PatchLimits, Reversible,
        SubEditJoinFailure,
    };
    use std::io::Cursor;

    fn doc(paragraphs: &[&str]) -> Vec<u8> {
        let mut writer = Writer::new();
        for paragraph in paragraphs {
            writer
                .add_paragraph_runs(
                    vec![(paragraph.to_string(), CharacterFormatting::default())],
                    ParagraphFormatting::default(),
                )
                .expect("fixture paragraph must be valid");
        }
        let mut output = Cursor::new(Vec::new());
        writer
            .write_to(&mut output)
            .expect("fixture DOC must serialize");
        output.into_inner()
    }

    fn patch_limits() -> PatchLimits {
        PatchLimits::new(
            BlobLimits::new(0, 0, 0),
            128 * 1024,
            32,
            8,
            16 * 1024,
            64 * 1024,
        )
    }

    #[test]
    fn same_shape_body_edit_is_reversible_and_source_checked() {
        let source = Snapshot::parse(&doc(&["alpha", "bravo"])).expect("snapshot");
        assert_eq!(
            source
                .paragraphs(Projection::All)
                .expect("paragraphs")
                .iter()
                .map(|paragraph| paragraph.text())
                .collect::<Vec<_>>(),
            ["alpha", "bravo"]
        );
        assert_eq!(
            source.paragraphs(Projection::All).expect("paragraphs")[0].position(),
            Position::new(0)
        );

        let mut edit = source.edit().expect("edit");
        edit.replace_paragraph(Position::new(0), "omega")
            .expect("same shape edit");
        let commit = edit.commit().expect("commit");
        assert!(commit.changed());
        assert_eq!(
            commit
                .snapshot()
                .paragraphs(Projection::All)
                .expect("changed paragraphs")[0]
                .text(),
            "omega"
        );

        let applied = commit.patch().apply(&source).expect("exact source applies");
        assert_eq!(applied, *commit.snapshot());
        let restored = commit
            .patch()
            .inverse()
            .apply(&applied)
            .expect("inverse applies");
        assert_eq!(restored.bytes(), source.bytes());

        let other = Snapshot::open(doc(&["other"]), Limits::default()).expect("other source");
        assert!(matches!(commit.patch().apply(&other), Err(Error::Conflict)));
    }

    #[test]
    fn length_changes_publish_while_structural_changes_are_refused() {
        let source = Snapshot::parse(&doc(&["alpha"])).expect("snapshot");
        let mut edit = source.edit().expect("edit");
        edit.replace_paragraph(Position::new(0), "a much longer paragraph")
            .expect("length-changing edit");
        assert!(matches!(
            edit.replace_paragraph(Position::new(0), "a\rpha"),
            Err(Error::Refused(
                Refusal::ReplacementContainsStructuralContent
            ))
        ));
        assert!(matches!(
            edit.replace_paragraph(Position::new(1), "alpha"),
            Err(Error::Refused(Refusal::ParagraphNotFound))
        ));
        let commit = edit.commit().expect("changed commit");
        assert!(commit.changed());
        let paragraphs = commit
            .snapshot()
            .paragraphs(Projection::All)
            .expect("changed paragraphs");
        assert_eq!(paragraphs[0].text(), "a much longer paragraph");
    }

    #[test]
    fn accepted_and_rejected_projections_hide_text_revisions() {
        let mut writer = Writer::new();
        writer
            .add_paragraph_runs(
                vec![
                    ("kept ".to_string(), CharacterFormatting::default()),
                    (
                        "old".to_string(),
                        CharacterFormatting {
                            deletion_revision: Some(TextRevision::new("Reviewer")),
                            ..CharacterFormatting::default()
                        },
                    ),
                    (" new".to_string(), CharacterFormatting::default()),
                ],
                ParagraphFormatting::default(),
            )
            .expect("fixture paragraph");
        let mut output = Cursor::new(Vec::new());
        writer.write_to(&mut output).expect("fixture DOC");
        let snapshot = Snapshot::parse(&output.into_inner()).expect("snapshot");
        assert_eq!(
            snapshot.paragraphs(Projection::Accepted).expect("accepted")[0].text(),
            "kept  new"
        );
        assert_eq!(
            snapshot.paragraphs(Projection::Rejected).expect("rejected")[0].text(),
            "kept old new"
        );
    }

    #[test]
    fn multi_paragraph_text_and_bold_commit_reopens_and_inverts() {
        let source = Snapshot::parse(&doc(&["alpha", "bravo", "charlie"])).expect("snapshot");
        let mut edit = source.edit().expect("edit");
        edit.replace_paragraph(Position::new(0), "alpha expanded")
            .expect("grow first paragraph");
        edit.replace_paragraph(Position::new(2), "c")
            .expect("shrink third paragraph");
        edit.set_paragraph_bold(Position::new(0), true)
            .expect("bold first paragraph");
        let commit = edit.commit().expect("commit and full reopen");
        let texts = commit
            .snapshot()
            .paragraphs(Projection::All)
            .expect("readback")
            .into_iter()
            .map(|paragraph| paragraph.text().to_string())
            .collect::<Vec<_>>();
        assert_eq!(texts, ["alpha expanded", "bravo", "c"]);

        let mut package = crate::Package::from_reader(Cursor::new(commit.snapshot().finish()))
            .expect("CFB reopens");
        let document = package.document().expect("DOC reopens");
        let paragraphs = document.paragraphs().expect("paragraphs read back");
        assert_eq!(paragraphs[0].runs().expect("runs")[0].bold(), Some(true));

        let restored = commit
            .patch()
            .inverse()
            .apply(commit.snapshot())
            .expect("inverse applies");
        assert_eq!(restored.bytes(), source.bytes());
    }

    #[test]
    fn durable_patch_composition_and_history_are_bounded() {
        let source = Snapshot::parse(&doc(&["alpha", "bravo"])).expect("snapshot");
        let composition_limits = CompositionLimits::new(4, 2, 8, 8);
        let text = source
            .prepare_replace(
                composition_limits,
                "replace-alpha",
                Position::new(0),
                "alpha grows",
            )
            .expect("prepare text");
        let bold = source
            .prepare_bold(composition_limits, "bold-alpha", Position::new(0), true)
            .expect("prepare bold");
        let conflict = source
            .prepare_replace(
                composition_limits,
                "replace-alpha-again",
                Position::new(0),
                "other",
            )
            .expect("prepare conflicting text");
        let mut composition = source.compose(composition_limits);
        composition.join(text).expect("text joins");
        composition.join(bold).expect("disjoint bold joins");
        let failure = composition
            .join(conflict)
            .expect_err("same facet conflicts");
        assert!(matches!(failure.failure(), SubEditJoinFailure::Overlap(_)));

        let commit = composition.commit().expect("composition commits");
        let durable = commit
            .patch()
            .to_durable(patch_limits())
            .expect("durable patch");
        let wire = durable.to_deterministic_json().expect("canonical JSON");
        let decoded = Patch::<Reversible>::from_deterministic_json(&wire, patch_limits())
            .expect("durable decode");
        assert_eq!(
            decoded.to_deterministic_json().expect("durable re-encode"),
            wire
        );
        assert_eq!(
            source.apply_durable(&decoded).expect("durable apply"),
            *commit.snapshot()
        );
        let semantic_inverse = commit
            .snapshot()
            .apply_durable(&decoded.inverse())
            .expect("durable inverse");
        assert_eq!(
            semantic_inverse
                .paragraphs(Projection::All)
                .expect("inverse paragraphs")
                .iter()
                .map(|paragraph| paragraph.text())
                .collect::<Vec<_>>(),
            ["alpha", "bravo"]
        );
        let mut inverse_package =
            crate::Package::from_reader(Cursor::new(semantic_inverse.finish()))
                .expect("durable inverse CFB reopens");
        let inverse_document = inverse_package
            .document()
            .expect("durable inverse DOC reopens");
        assert_eq!(
            inverse_document.paragraphs().expect("inverse paragraphs")[0]
                .runs()
                .expect("inverse runs")[0]
                .bold(),
            None
        );

        let mut history = source.history(HistoryLimits::new(1, wire.len() as u64));
        history
            .record(commit.snapshot().clone(), wire.len() as u64)
            .expect("record history");
        assert!(history.undo());
        assert_eq!(history.current(), &source);
        assert!(history.redo());
        assert_eq!(history.current(), commit.snapshot());

        let mut too_small = source.history(HistoryLimits::new(1, wire.len() as u64 - 1));
        assert!(
            too_small
                .record(commit.snapshot().clone(), wire.len() as u64)
                .is_err()
        );
        assert_eq!(too_small.current(), &source);
    }

    #[test]
    fn transaction_limits_fail_before_mutation() {
        let source = Snapshot::open_bounded(
            doc(&["alpha", "bravo"]),
            Limits::default(),
            TransactionLimits::new(1, 5, 5),
        )
        .expect("bounded source");
        let mut edit = source.edit().expect("edit");
        assert!(matches!(
            edit.replace_paragraph(Position::new(0), "too long"),
            Err(Error::Refused(Refusal::ReplacementLimit { .. }))
        ));
        edit.replace_paragraph(Position::new(0), "short")
            .expect("within bounds");
        assert!(matches!(
            edit.set_paragraph_bold(Position::new(1), true),
            Err(Error::Refused(Refusal::OperationLimit { .. }))
        ));
    }
}
