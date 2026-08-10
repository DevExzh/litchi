//! Bounded immutable RTF body-text and paragraph-property transactions.
//!
//! Operations resolve against one immutable semantic body. Disjoint UTF-8
//! spans and paragraph-alignment facets compose in one atomic commit. The
//! exact-body closure deliberately excludes body-anchored opaque syntax,
//! tables, positioned content, and mixed formatting whose dependent ranges
//! cannot be updated losslessly yet. Canonical retained-story edits cover
//! checked table cells, headers/footers, comments, notes, and root shape text
//! frames while refusing unknown destinations and dependent positioned content.

use crate::{Alignment, Document, HeaderFooterType, RtfError, RtfWriter, TableCellPath};
use bumpalo::Bump;
use serde_json::Value;
use std::borrow::Cow;
use std::collections::BTreeMap;
use std::fmt;
use std::ops::Range;

mod composition;
mod transfer;
pub use composition::{
    Composition, CompositionConflict, CompositionError, CompositionLimits, ConflictSet, MergePlan,
    MergeResolution, Prepared,
};
pub use transfer::TransferPlan;

/// Immutable RTF snapshot used by the transaction API.
pub type Snapshot = Document;

/// Finite step and retained-weight bounds for [`History`].
pub use litchi_core::patch::HistoryLimits;

/// Commit-coupled bounded undo/redo history for immutable RTF snapshots.
pub struct History {
    inner: litchi_core::patch::History<Snapshot>,
}

impl History {
    /// Starts history at one immutable snapshot.
    #[must_use]
    pub fn new(current: Snapshot, limits: HistoryLimits) -> Self {
        Self {
            inner: litchi_core::patch::History::new(current, limits),
        }
    }

    /// Current immutable snapshot.
    #[must_use]
    pub const fn current(&self) -> &Snapshot {
        self.inner.current()
    }

    /// Whether an undo transition exists.
    #[must_use]
    pub fn can_undo(&self) -> bool {
        self.inner.can_undo()
    }

    /// Whether a redo transition exists.
    #[must_use]
    pub fn can_redo(&self) -> bool {
        self.inner.can_redo()
    }

    /// Commits an edit rooted at the current snapshot and records it atomically.
    ///
    /// # Errors
    /// Returns a source, edit, or history-budget error without changing history.
    pub fn commit(&mut self, edit: Edit) -> Result<Commit, Error> {
        if !edit.source().same_snapshot(self.current()) {
            return Err(Error::HistorySourceMismatch);
        }
        let commit = edit.commit()?;
        self.record_commit(&commit)?;
        Ok(commit)
    }

    /// Records an already validated commit rooted at the current snapshot.
    ///
    /// # Errors
    /// Returns a source or budget error without changing history.
    pub fn record_commit(&mut self, commit: &Commit) -> Result<(), Error> {
        if !commit.patch.before.same_snapshot(self.current()) {
            return Err(Error::HistorySourceMismatch);
        }
        if commit.diagnostics.changed {
            self.inner
                .record(commit.snapshot.clone(), commit.patch.history_weight())
                .map_err(|error| Error::History(error.to_string()))?;
        }
        Ok(())
    }

    /// Moves to the preceding retained snapshot.
    pub fn undo(&mut self) -> bool {
        self.inner.undo()
    }

    /// Moves to the following retained snapshot.
    pub fn redo(&mut self) -> bool {
        self.inner.redo()
    }
}

/// Default maximum number of operations staged in one edit.
pub const DEFAULT_MAX_OPERATIONS: usize = 256;

/// Finite limits for one RTF edit plan.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Limits {
    max_operations: usize,
}

impl Limits {
    /// Creates an operation bound. Zero permits only an empty commit.
    #[must_use]
    pub const fn new(max_operations: usize) -> Self {
        Self { max_operations }
    }

    /// Maximum staged semantic operations.
    #[must_use]
    pub const fn max_operations(self) -> usize {
        self.max_operations
    }
}

impl Default for Limits {
    fn default() -> Self {
        Self::new(DEFAULT_MAX_OPERATIONS)
    }
}

/// A checked half-open UTF-8 byte span in the semantic body text.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct TextSpan {
    start: usize,
    end: usize,
}

impl TextSpan {
    /// Creates an ordered half-open span. Document bounds and UTF-8 boundaries
    /// are checked when the span is staged.
    ///
    /// # Errors
    /// Returns [`Error::ReversedSpan`] when `start` is after `end`.
    pub const fn new(start: usize, end: usize) -> Result<Self, Error> {
        if start > end {
            return Err(Error::ReversedSpan { start, end });
        }
        Ok(Self { start, end })
    }

    /// Start UTF-8 byte position.
    #[must_use]
    pub const fn start(self) -> usize {
        self.start
    }

    /// Exclusive end UTF-8 byte position.
    #[must_use]
    pub const fn end(self) -> usize {
        self.end
    }

    /// Whether this span is an insertion point.
    #[must_use]
    pub const fn is_empty(self) -> bool {
        self.start == self.end
    }
}

/// Checked route to one paragraph in a retained section header or footer.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct HeaderFooterParagraph {
    section: usize,
    kind: HeaderFooterType,
    paragraph: usize,
}

impl HeaderFooterParagraph {
    /// Creates a selector whose existence is checked when an operation is staged.
    #[must_use]
    pub const fn new(section: usize, kind: HeaderFooterType, paragraph: usize) -> Self {
        Self {
            section,
            kind,
            paragraph,
        }
    }

    /// Zero-based section position.
    #[must_use]
    pub const fn section(self) -> usize {
        self.section
    }

    /// Native header/footer destination kind.
    #[must_use]
    pub const fn kind(self) -> HeaderFooterType {
        self.kind
    }

    /// Zero-based paragraph position within the destination.
    #[must_use]
    pub const fn paragraph(self) -> usize {
        self.paragraph
    }
}

/// Failure from an RTF transaction or patch application.
#[derive(Debug, Clone)]
#[non_exhaustive]
pub enum Error {
    /// A source-specific body feature cannot be rewritten through this seam.
    UnsupportedSource(&'static str),
    /// Compatibility variant retained from the former one-operation editor.
    OperationAlreadyStaged,
    /// The finite operation bound was exceeded.
    OperationLimit { observed: usize, limit: usize },
    /// Two staged operations have overlapping effects.
    Conflict { existing: usize, incoming: usize },
    /// Paragraph-structure and property changes cannot be proven independent.
    StructuralPropertyConflict,
    /// Canonical retained-destination work cannot share a transaction with a body splice.
    BodyDestinationConflict,
    /// Replacement text exceeds the source snapshot's retained resource profile.
    InputTooLarge { observed: usize, limit: usize },
    /// A span has its endpoints in reverse order.
    ReversedSpan { start: usize, end: usize },
    /// A span is outside the semantic body.
    SpanOutOfRange { end: usize, length: usize },
    /// A span endpoint is not a UTF-8 boundary.
    SpanNotOnCharacterBoundary { position: usize },
    /// A checked paragraph position does not exist in the source body story.
    ParagraphOutOfRange { position: usize, count: usize },
    /// A checked retained-destination selector does not exist.
    DestinationOutOfRange(&'static str),
    /// A durable operation's semantic expectation is stale.
    StalePrecondition(&'static str),
    /// A durable patch uses an unsupported format-owned vocabulary.
    DurablePatch(String),
    /// An edit or commit does not originate from the history's current snapshot.
    HistorySourceMismatch,
    /// The common bounded history refused a transition.
    History(String),
    /// Candidate parsing or validation failed.
    Rtf(RtfError),
    /// Candidate transport construction failed before publication.
    Write(String),
    /// The patch was applied to bytes other than the snapshot that created it.
    PatchConflict,
}

impl fmt::Display for Error {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::UnsupportedSource(reason) => {
                write!(formatter, "unsupported RTF edit source: {reason}")
            },
            Self::OperationAlreadyStaged => {
                formatter.write_str("RTF edit already has a staged operation")
            },
            Self::OperationLimit { observed, limit } => write!(
                formatter,
                "RTF edit operation limit exceeded: observed {observed}, limit {limit}"
            ),
            Self::Conflict { existing, incoming } => write!(
                formatter,
                "RTF edit operation {incoming} conflicts with operation {existing}"
            ),
            Self::StructuralPropertyConflict => formatter
                .write_str("RTF paragraph-structure and property changes cannot compose safely"),
            Self::BodyDestinationConflict => formatter.write_str(
                "RTF body splices and canonical retained-destination edits cannot compose safely",
            ),
            Self::InputTooLarge { observed, limit } => write!(
                formatter,
                "replacement body text exceeds the source limit: observed {observed}, limit {limit}"
            ),
            Self::ReversedSpan { start, end } => {
                write!(formatter, "RTF text span {start}..{end} is reversed")
            },
            Self::SpanOutOfRange { end, length } => write!(
                formatter,
                "RTF text span ends at {end}, beyond body length {length}"
            ),
            Self::SpanNotOnCharacterBoundary { position } => write!(
                formatter,
                "RTF text span position {position} is not a UTF-8 boundary"
            ),
            Self::ParagraphOutOfRange { position, count } => write!(
                formatter,
                "RTF body paragraph position {position} is outside 0..{count}"
            ),
            Self::DestinationOutOfRange(destination) => {
                write!(
                    formatter,
                    "RTF retained destination does not exist: {destination}"
                )
            },
            Self::StalePrecondition(reason) => {
                write!(formatter, "stale RTF patch precondition: {reason}")
            },
            Self::DurablePatch(reason) => write!(formatter, "invalid durable RTF patch: {reason}"),
            Self::HistorySourceMismatch => {
                formatter.write_str("RTF history source is not its current snapshot")
            },
            Self::History(reason) => write!(formatter, "RTF history refused commit: {reason}"),
            Self::Rtf(error) => error.fmt(formatter),
            Self::Write(error) => write!(formatter, "RTF candidate construction failed: {error}"),
            Self::PatchConflict => {
                formatter.write_str("RTF patch source does not match its expected snapshot")
            },
        }
    }
}

impl std::error::Error for Error {}

impl From<RtfError> for Error {
    fn from(error: RtfError) -> Self {
        Self::Rtf(error)
    }
}

/// Detached, bounded edit of an immutable snapshot.
pub struct Edit {
    source: Snapshot,
    limits: Limits,
    operations: Vec<Operation>,
    replacement_bytes: usize,
}

#[derive(Debug, Clone)]
enum Operation {
    Text {
        span: TextSpan,
        before: String,
        after: String,
        structural: bool,
    },
    Alignment {
        position: usize,
        before: Alignment,
        after: Alignment,
    },
    Bold {
        span: TextSpan,
        before: bool,
        after: bool,
    },
    InsertParagraph {
        position: usize,
        span: TextSpan,
        text: String,
    },
    TableCellText {
        path: TableCellPath,
        before: String,
        after: String,
    },
    HeaderFooterText {
        target: HeaderFooterParagraph,
        before: String,
        after: String,
    },
    AnnotationText {
        index: usize,
        before: String,
        after: String,
    },
    NoteText {
        index: usize,
        before: String,
        after: String,
    },
    ShapeText {
        index: usize,
        before: String,
        after: String,
    },
    RootTransfer {
        vocabulary: &'static str,
        effect: String,
        before: Vec<u8>,
        after: Vec<u8>,
    },
}

impl Operation {
    fn replacement_bytes(&self) -> usize {
        match self {
            Self::Text { after, .. }
            | Self::TableCellText { after, .. }
            | Self::HeaderFooterText { after, .. }
            | Self::AnnotationText { after, .. }
            | Self::NoteText { after, .. }
            | Self::ShapeText { after, .. } => after.len(),
            Self::InsertParagraph { text, .. } => text.len().saturating_add(1),
            Self::RootTransfer { after, .. } => after.len(),
            Self::Alignment { .. } | Self::Bold { .. } => 0,
        }
    }

    fn effect_keys(&self) -> Vec<String> {
        match self {
            Self::Text { structural, .. } if *structural => {
                vec!["body:structure".to_string()]
            },
            Self::Text { .. } => vec!["body:text".to_string()],
            Self::Alignment { position, .. } => {
                vec![format!("body:paragraph:{position}:alignment")]
            },
            Self::Bold { span, .. } => {
                vec![format!("body:character:{}-{}:bold", span.start, span.end)]
            },
            Self::InsertParagraph { .. } => vec!["body:structure".to_string()],
            Self::TableCellText { path, .. } => vec![table_cell_effect(path)],
            Self::HeaderFooterText { target, .. } => vec![header_footer_effect(*target)],
            Self::AnnotationText { index, .. } => vec![annotation_effect(*index)],
            Self::NoteText { index, .. } => vec![note_effect(*index)],
            Self::ShapeText { index, .. } => vec![shape_effect(*index)],
            Self::RootTransfer { effect, .. } => vec![effect.clone()],
        }
    }

    const fn span(&self) -> Option<TextSpan> {
        match self {
            Self::Text { span, .. }
            | Self::Bold { span, .. }
            | Self::InsertParagraph { span, .. } => Some(*span),
            Self::Alignment { .. }
            | Self::TableCellText { .. }
            | Self::HeaderFooterText { .. }
            | Self::AnnotationText { .. }
            | Self::NoteText { .. }
            | Self::ShapeText { .. }
            | Self::RootTransfer { .. } => None,
        }
    }

    const fn is_property(&self) -> bool {
        matches!(self, Self::Alignment { .. } | Self::Bold { .. })
    }

    const fn is_destination(&self) -> bool {
        matches!(
            self,
            Self::TableCellText { .. }
                | Self::HeaderFooterText { .. }
                | Self::AnnotationText { .. }
                | Self::NoteText { .. }
                | Self::ShapeText { .. }
                | Self::RootTransfer { .. }
        )
    }

    const fn is_root_transfer(&self) -> bool {
        matches!(self, Self::RootTransfer { .. })
    }
}

impl Edit {
    pub(crate) fn new(source: Snapshot) -> Self {
        Self::new_with_limits(source, Limits::default())
    }

    pub(crate) fn new_with_limits(source: Snapshot, limits: Limits) -> Self {
        Self {
            source,
            limits,
            operations: Vec::new(),
            replacement_bytes: 0,
        }
    }

    pub(crate) fn stage_root_transfer(
        &mut self,
        vocabulary: &'static str,
        effect: String,
        after: Vec<u8>,
    ) -> Result<&mut Self, Error> {
        self.ensure_operation_room()?;
        if !self.operations.is_empty() {
            return Err(Error::BodyDestinationConflict);
        }
        if !self.source.opaque().is_empty() {
            return Err(Error::UnsupportedSource(
                "ordinary-root transfer refuses unknown target destinations",
            ));
        }
        let before = self
            .source
            .source_bytes()
            .ok_or(Error::UnsupportedSource("snapshot has no exact RTF source"))?
            .to_vec();
        self.charge_replacement(after.len())?;
        self.operations.push(Operation::RootTransfer {
            vocabulary,
            effect,
            before,
            after,
        });
        Ok(self)
    }

    /// Returns the immutable source snapshot.
    #[must_use]
    pub const fn source(&self) -> &Snapshot {
        &self.source
    }

    /// Number of staged operations.
    #[must_use]
    pub fn operation_count(&self) -> usize {
        self.operations.len()
    }

    /// Stages replacement of the complete ordinary body story.
    ///
    /// A newline creates an RTF paragraph break. This operation conflicts with
    /// every other text-span operation, but independent paragraph properties
    /// may compose when paragraph structure is unchanged.
    ///
    /// # Errors
    /// Returns an error for invalid bounds, conflicts, or retained limits.
    pub fn replace_body_text(
        &mut self,
        replacement: impl Into<String>,
    ) -> Result<&mut Self, Error> {
        let span = TextSpan::new(0, self.source.text().len())?;
        self.replace_text(span, replacement)
    }

    /// Stages replacement of one checked zero-based ordinary body paragraph.
    ///
    /// Selection resolves against the immutable source snapshot. Newlines in
    /// the replacement create paragraph breaks, and therefore cannot compose
    /// with paragraph-property operations in the same edit.
    ///
    /// # Errors
    /// Returns an error for an invalid selector, conflicts, or retained limits.
    pub fn replace_paragraph_text(
        &mut self,
        position: usize,
        replacement: impl Into<String>,
    ) -> Result<&mut Self, Error> {
        let range = paragraph_range(&self.source, position)?;
        self.replace_text(TextSpan::new(range.start, range.end)?, replacement)
    }

    /// Stages one UTF-8 semantic body splice.
    ///
    /// Every selector resolves against the immutable base text. Disjoint spans
    /// compose regardless of staging order; overlap returns a deterministic
    /// conflict instead of applying last-writer-wins behavior.
    ///
    /// # Errors
    /// Returns an error for invalid UTF-8 boundaries, conflicts, or limits.
    pub fn replace_text(
        &mut self,
        span: TextSpan,
        replacement: impl Into<String>,
    ) -> Result<&mut Self, Error> {
        self.ensure_body_compatible()?;
        self.ensure_operation_room()?;
        let body = self.source.text();
        validate_span(body, span)?;
        let after = replacement.into();
        let before = body
            .get(span.start..span.end)
            .ok_or(Error::SpanNotOnCharacterBoundary {
                position: span.start,
            })?
            .to_string();
        let structural = before.contains('\n') || after.contains('\n');
        if structural
            && self
                .operations
                .iter()
                .any(|operation| matches!(operation, Operation::Alignment { .. }))
        {
            return Err(Error::StructuralPropertyConflict);
        }
        let incoming = self.operations.len();
        for (existing, operation) in self.operations.iter().enumerate() {
            if operation
                .span()
                .is_some_and(|existing_span| spans_conflict(existing_span, span))
            {
                return Err(Error::Conflict { existing, incoming });
            }
        }
        self.charge_replacement(after.len())?;
        self.operations.push(Operation::Text {
            span,
            before,
            after,
            structural,
        });
        Ok(self)
    }

    /// Stages bold state for one non-empty UTF-8 body span.
    ///
    /// The selected source range must have one effective bold state and may
    /// not consume a paragraph boundary. Unknown or mixed character ranges
    /// are refused rather than normalized.
    ///
    /// # Errors
    /// Returns an error for invalid geometry, conflicts, structure changes,
    /// mixed formatting, or finite bounds.
    pub fn set_text_bold(&mut self, span: TextSpan, bold: bool) -> Result<&mut Self, Error> {
        self.ensure_body_compatible()?;
        self.ensure_operation_room()?;
        let body = self.source.text();
        validate_span(body, span)?;
        if span.is_empty()
            || body
                .get(span.start..span.end)
                .is_some_and(|text| text.contains('\n'))
        {
            return Err(Error::UnsupportedSource(
                "bold edits require non-empty text within one paragraph",
            ));
        }
        if self.operations.iter().any(|operation| {
            matches!(
                operation,
                Operation::Text {
                    structural: true,
                    ..
                } | Operation::InsertParagraph { .. }
            )
        }) {
            return Err(Error::StructuralPropertyConflict);
        }
        let incoming = self.operations.len();
        for (existing, operation) in self.operations.iter().enumerate() {
            if operation
                .span()
                .is_some_and(|existing_span| spans_conflict(existing_span, span))
            {
                return Err(Error::Conflict { existing, incoming });
            }
        }
        let before = bold_for_span(&self.source, span)?;
        self.operations.push(Operation::Bold {
            span,
            before,
            after: bold,
        });
        Ok(self)
    }

    /// Inserts one ordinary paragraph immediately after a checked paragraph.
    ///
    /// This is the transaction's explicit structural class. The inserted text
    /// cannot contain another paragraph break; multiple disjoint insertions
    /// may compose, while paragraph/character property edits conflict.
    ///
    /// # Errors
    /// Returns an error for an invalid selector, newline-bearing input,
    /// conflicts, or finite bounds.
    pub fn insert_paragraph_after(
        &mut self,
        position: usize,
        input: impl Into<String>,
    ) -> Result<&mut Self, Error> {
        self.ensure_body_compatible()?;
        self.ensure_operation_room()?;
        if self.operations.iter().any(Operation::is_property) {
            return Err(Error::StructuralPropertyConflict);
        }
        let text = input.into();
        if text.contains('\n') {
            return Err(Error::UnsupportedSource(
                "one structural insertion authors exactly one paragraph",
            ));
        }
        let range = paragraph_range(&self.source, position)?;
        let span = TextSpan::new(range.end, range.end)?;
        let incoming = self.operations.len();
        for (existing, operation) in self.operations.iter().enumerate() {
            if operation
                .span()
                .is_some_and(|existing_span| spans_conflict(existing_span, span))
            {
                return Err(Error::Conflict { existing, incoming });
            }
        }
        self.charge_replacement(text.len().saturating_add(1))?;
        self.operations.push(Operation::InsertParagraph {
            position,
            span,
            text,
        });
        Ok(self)
    }

    fn remove_paragraph_after(
        &mut self,
        position: usize,
        expected: &str,
    ) -> Result<&mut Self, Error> {
        let owner = paragraph_range(&self.source, position)?;
        let removed = paragraph_range(&self.source, position.saturating_add(1))?;
        let actual = self
            .source
            .text()
            .get(removed.clone())
            .ok_or(Error::StalePrecondition("inserted paragraph disappeared"))?;
        if actual != expected {
            return Err(Error::StalePrecondition("inserted paragraph text differs"));
        }
        self.replace_text(TextSpan::new(owner.end, removed.end)?, "")
    }

    /// Stages the local alignment of one checked base-snapshot paragraph.
    ///
    /// This is a property facet, not a raw control-word API. The commit writes
    /// the RTF 1.9.1 `ql`, `qr`, `qc`, or `qj` selector and verifies semantic
    /// readback before publication.
    ///
    /// # Errors
    /// Returns an error for an invalid selector, a duplicate facet, structural
    /// text changes in the same transaction, or retained limits.
    pub fn set_paragraph_alignment(
        &mut self,
        position: usize,
        alignment: Alignment,
    ) -> Result<&mut Self, Error> {
        self.ensure_body_compatible()?;
        self.ensure_operation_room()?;
        if self.operations.iter().any(|operation| {
            matches!(
                operation,
                Operation::Text {
                    structural: true,
                    ..
                } | Operation::InsertParagraph { .. }
            )
        }) {
            return Err(Error::StructuralPropertyConflict);
        }
        let paragraphs = self.source.body().paragraphs().collect::<Vec<_>>();
        let count = paragraphs.len();
        let paragraph = paragraphs
            .get(position)
            .ok_or(Error::ParagraphOutOfRange { position, count })?;
        let incoming = self.operations.len();
        if let Some(existing) = self.operations.iter().position(|operation| {
            matches!(
                operation,
                Operation::Alignment {
                    position: existing_position,
                    ..
                } if *existing_position == position
            )
        }) {
            return Err(Error::Conflict { existing, incoming });
        }
        self.operations.push(Operation::Alignment {
            position,
            before: paragraph.format().alignment(),
            after: alignment,
        });
        Ok(self)
    }

    /// Stages replacement of one retained table-cell text story.
    ///
    /// The complete cell path is resolved against the immutable base. Cell
    /// drawings, fields, nested tables, and positional events remain retained;
    /// their positions must still validate against the replacement. Canonical
    /// publication refuses every snapshot containing unknown destinations.
    ///
    /// # Errors
    /// Returns an error for an invalid path, dependent story geometry,
    /// duplicate destination, mixed body work, or finite limits.
    pub fn set_table_cell_text(
        &mut self,
        path: TableCellPath,
        input: impl Into<String>,
    ) -> Result<&mut Self, Error> {
        self.ensure_destination_compatible()?;
        self.ensure_operation_room()?;
        let after = input.into();
        let before = table_cell(&self.source, &path)?.text().to_string();
        let effect = table_cell_effect(&path);
        self.ensure_unique_destination(&effect)?;

        // Exercise the retained model's dependency validation before staging.
        let mut candidate = table_cell(&self.source, &path)?.clone();
        candidate.set_text(Cow::Owned(after.clone()))?;
        self.charge_replacement(after.len())?;
        self.operations.push(Operation::TableCellText {
            path,
            before,
            after,
        });
        Ok(self)
    }

    /// Stages replacement of one retained header/footer paragraph.
    ///
    /// The paragraph's formatting is preserved. Destinations with drawings,
    /// fields, or other positional story events are refused because changing
    /// text could stale dependent offsets.
    ///
    /// # Errors
    /// Returns an error for an invalid selector, dependent story content,
    /// duplicate destination, mixed body work, or finite limits.
    pub fn set_header_footer_text(
        &mut self,
        target: HeaderFooterParagraph,
        input: impl Into<String>,
    ) -> Result<&mut Self, Error> {
        self.ensure_destination_compatible()?;
        self.ensure_operation_room()?;
        let after = input.into();
        if after.contains('\n') {
            return Err(Error::UnsupportedSource(
                "one header/footer paragraph cannot contain a paragraph break",
            ));
        }
        let header_footer = header_footer(&self.source, target)?;
        if !header_footer.shapes.is_empty()
            || !header_footer.shape_groups.is_empty()
            || !header_footer.story_events.is_empty()
        {
            return Err(Error::UnsupportedSource(
                "header/footer text has dependent positioned content",
            ));
        }
        let before = header_footer
            .paragraphs
            .get(target.paragraph)
            .ok_or(Error::DestinationOutOfRange("header/footer paragraph"))?
            .text
            .to_string();
        let effect = header_footer_effect(target);
        self.ensure_unique_destination(&effect)?;
        self.charge_replacement(after.len())?;
        self.operations.push(Operation::HeaderFooterText {
            target,
            before,
            after,
        });
        Ok(self)
    }

    /// Stages replacement of one retained comment body.
    ///
    /// The comment's range identity, author metadata, and body anchor are
    /// preserved. Positioned drawings or fields are refused because replacing
    /// the complete text could stale their offsets.
    ///
    /// # Errors
    /// Returns an error for an invalid index, dependent story content,
    /// duplicate destination, mixed body work, or finite limits.
    pub fn set_annotation_text(
        &mut self,
        index: usize,
        input: impl Into<String>,
    ) -> Result<&mut Self, Error> {
        self.ensure_destination_compatible()?;
        self.ensure_operation_room()?;
        let after = input.into();
        let annotation = annotation(&self.source, index)?;
        if !annotation.shapes.is_empty()
            || !annotation.shape_groups.is_empty()
            || !annotation.story_events.is_empty()
        {
            return Err(Error::UnsupportedSource(
                "annotation text has dependent positioned content",
            ));
        }
        let mut candidate = annotation.clone();
        candidate.text = Cow::Owned(after.clone());
        candidate.validate()?;
        let before = annotation.text.to_string();
        let effect = annotation_effect(index);
        self.ensure_unique_destination(&effect)?;
        self.charge_replacement(after.len())?;
        self.operations.push(Operation::AnnotationText {
            index,
            before,
            after,
        });
        Ok(self)
    }

    /// Stages replacement of one retained footnote or endnote body.
    ///
    /// The note kind, reference mark, formatting, and main-story anchor are
    /// preserved. Positioned drawings or fields are refused because replacing
    /// the complete text could stale their offsets.
    ///
    /// # Errors
    /// Returns an error for an invalid index, dependent story content,
    /// duplicate destination, mixed body work, or finite limits.
    pub fn set_note_text(
        &mut self,
        index: usize,
        input: impl Into<String>,
    ) -> Result<&mut Self, Error> {
        self.ensure_destination_compatible()?;
        self.ensure_operation_room()?;
        let after = input.into();
        let note = note(&self.source, index)?;
        if !note.shapes.is_empty() || !note.shape_groups.is_empty() || !note.story_events.is_empty()
        {
            return Err(Error::UnsupportedSource(
                "note text has dependent positioned content",
            ));
        }
        let mut candidate = note.clone();
        candidate.content = Cow::Owned(after.clone());
        candidate.validate()?;
        let before = note.content.to_string();
        let effect = note_effect(index);
        self.ensure_unique_destination(&effect)?;
        self.charge_replacement(after.len())?;
        self.operations.push(Operation::NoteText {
            index,
            before,
            after,
        });
        Ok(self)
    }

    /// Stages replacement of one retained root shape text frame.
    ///
    /// Geometry, scalar properties, formatting, name, and body position are
    /// preserved. Shapes without an existing text destination and shapes with
    /// hyperlinks, legacy fallbacks, or nested positioned story content are
    /// refused.
    ///
    /// # Errors
    /// Returns an error for an invalid index, active/dependent drawing content,
    /// duplicate destination, mixed body work, or finite limits.
    pub fn set_shape_text(
        &mut self,
        index: usize,
        input: impl Into<String>,
    ) -> Result<&mut Self, Error> {
        self.ensure_destination_compatible()?;
        self.ensure_operation_room()?;
        let after = input.into();
        let shape = shape(&self.source, index)?;
        if !shape.text_destination_present {
            return Err(Error::UnsupportedSource(
                "shape text editing requires an existing text destination",
            ));
        }
        if shape.result.is_some()
            || !shape.text_shapes.is_empty()
            || !shape.text_shape_groups.is_empty()
            || !shape.text_drawing_order.is_empty()
            || !shape.text_story_events.is_empty()
        {
            return Err(Error::UnsupportedSource(
                "shape text has dependent positioned or fallback content",
            ));
        }
        if shape_has_active_link(shape) {
            return Err(Error::UnsupportedSource(
                "shape text editing refuses active hyperlink metadata",
            ));
        }
        let mut candidate = shape.clone();
        candidate.set_text(Cow::Owned(after.clone()));
        candidate.validate()?;
        let before = shape.text.to_string();
        let effect = shape_effect(index);
        self.ensure_unique_destination(&effect)?;
        self.charge_replacement(after.len())?;
        self.operations.push(Operation::ShapeText {
            index,
            before,
            after,
        });
        Ok(self)
    }

    fn ensure_body_compatible(&self) -> Result<(), Error> {
        if self.operations.iter().any(Operation::is_destination) {
            return Err(Error::BodyDestinationConflict);
        }
        Ok(())
    }

    fn ensure_destination_compatible(&self) -> Result<(), Error> {
        if self
            .operations
            .iter()
            .any(|operation| !operation.is_destination())
        {
            return Err(Error::BodyDestinationConflict);
        }
        Ok(())
    }

    fn ensure_unique_destination(&self, effect: &str) -> Result<(), Error> {
        let incoming = self.operations.len();
        if let Some(existing) = self.operations.iter().position(|operation| {
            operation
                .effect_keys()
                .iter()
                .any(|candidate| candidate == effect)
        }) {
            return Err(Error::Conflict { existing, incoming });
        }
        Ok(())
    }

    fn ensure_operation_room(&self) -> Result<(), Error> {
        let observed = self.operations.len().saturating_add(1);
        if observed > self.limits.max_operations {
            return Err(Error::OperationLimit {
                observed,
                limit: self.limits.max_operations,
            });
        }
        Ok(())
    }

    fn charge_replacement(&mut self, bytes: usize) -> Result<(), Error> {
        let observed = self.replacement_bytes.saturating_add(bytes);
        let limit = self.source.limits().max_source_bytes();
        if observed > limit {
            return Err(Error::InputTooLarge { observed, limit });
        }
        self.replacement_bytes = observed;
        Ok(())
    }

    /// Validates and publishes the complete candidate atomically.
    ///
    /// # Errors
    /// Returns an error when the changed source is outside the supported
    /// closure or candidate validation/readback fails.
    pub fn commit(self) -> Result<Commit, Error> {
        let operation_count = self.operations.len();
        if operation_count == 0 {
            return Ok(Commit::new(
                self.source.clone(),
                self.source,
                false,
                0,
                Vec::new(),
            ));
        }

        let destination_count = self
            .operations
            .iter()
            .filter(|operation| operation.is_destination())
            .count();
        let root_transfer_count = self
            .operations
            .iter()
            .filter(|operation| operation.is_root_transfer())
            .count();
        if root_transfer_count != 0 {
            if root_transfer_count != 1 || operation_count != 1 {
                return Err(Error::BodyDestinationConflict);
            }
            return commit_root_transfer(self, operation_count);
        }
        if destination_count == operation_count {
            return commit_destinations(self, operation_count);
        }
        if destination_count != 0 {
            return Err(Error::BodyDestinationConflict);
        }

        let (replacement, projected_spans) = project_text(&self.source, &self.operations)?;
        let property_operation = self.operations.iter().any(|operation| {
            matches!(
                operation,
                Operation::Alignment { .. } | Operation::Bold { .. }
            )
        });
        let mut alignments = if property_operation {
            source_alignments(&self.source)
        } else {
            Vec::new()
        };
        let base_bold = if property_operation {
            base_bold_for_edit(&self.source, &self.operations)?
        } else {
            false
        };
        let mut projected_bold_ranges = Vec::new();
        for operation in &self.operations {
            match operation {
                Operation::Alignment {
                    position, after, ..
                } => {
                    let count = alignments.len();
                    let slot = alignments
                        .get_mut(*position)
                        .ok_or(Error::ParagraphOutOfRange {
                            position: *position,
                            count,
                        })?;
                    *slot = *after;
                },
                Operation::Bold { span, after, .. } => {
                    projected_bold_ranges
                        .push((project_base_span(*span, &self.operations)?, *after));
                },
                Operation::Text { .. } | Operation::InsertParagraph { .. } => {},
                Operation::TableCellText { .. }
                | Operation::HeaderFooterText { .. }
                | Operation::AnnotationText { .. }
                | Operation::NoteText { .. }
                | Operation::ShapeText { .. } => {
                    return Err(Error::BodyDestinationConflict);
                },
                Operation::RootTransfer { .. } => return Err(Error::BodyDestinationConflict),
            }
        }
        let original_alignments = if property_operation {
            source_alignments(&self.source)
        } else {
            Vec::new()
        };
        let has_bold_delta = self.operations.iter().any(|operation| {
            matches!(operation, Operation::Bold { before, after, .. } if before != after)
        });
        let did_change = replacement != self.source.text()
            || alignments != original_alignments
            || has_bold_delta;
        let semantic_delta = semantic_changes(&self.operations, &projected_spans);
        if !did_change {
            return Ok(Commit::new(
                self.source.clone(),
                self.source,
                false,
                operation_count,
                semantic_delta,
            ));
        }

        let source_bytes = self
            .source
            .source_bytes()
            .ok_or(Error::UnsupportedSource("snapshot has no exact RTF source"))?;
        if crate::compressed::is_compressed_rtf(source_bytes) {
            return Err(Error::UnsupportedSource(
                "compressed RTF needs a transport-aware rewrite",
            ));
        }
        let has_bold_operation = self
            .operations
            .iter()
            .any(|operation| matches!(operation, Operation::Bold { .. }));
        if has_bold_operation {
            self.source
                .model()
                .plain_body_bold_editability()
                .map_err(Error::UnsupportedSource)?;
        } else {
            self.source
                .model()
                .plain_body_text_editability()
                .map_err(Error::UnsupportedSource)?;
        }
        let span =
            ordinary_body_source_span(source_bytes, self.source.text(), self.source.limits())?;
        validate_opaque_preservation(&self.source, source_bytes, &span)?;
        let replacement_bytes = if property_operation {
            encoded_body_with_properties(
                &replacement,
                &alignments,
                base_bold,
                &projected_bold_ranges,
                self.source.limits(),
            )?
        } else {
            encoded_body_text(&replacement, self.source.limits())?
        };
        let bytes = splice_body(source_bytes, span, &replacement_bytes, self.source.limits())?;
        let snapshot = Snapshot::from_bytes_with_limits(&bytes, self.source.limits())?;
        if snapshot.text() != replacement {
            return Err(Error::UnsupportedSource(
                "candidate body text did not survive RTF validation",
            ));
        }
        if property_operation && source_alignments(&snapshot) != alignments {
            return Err(Error::UnsupportedSource(
                "candidate paragraph alignment did not survive RTF validation",
            ));
        }
        for (bold_span, expected) in projected_bold_ranges {
            if bold_for_span(&snapshot, bold_span)? != expected {
                return Err(Error::UnsupportedSource(
                    "candidate bold property did not survive RTF validation",
                ));
            }
        }
        Ok(Commit::new(
            self.source,
            snapshot,
            true,
            operation_count,
            semantic_delta,
        ))
    }
}

fn commit_destinations(edit: Edit, operation_count: usize) -> Result<Commit, Error> {
    let source_bytes = edit
        .source
        .source_bytes()
        .ok_or(Error::UnsupportedSource("snapshot has no exact RTF source"))?;
    if crate::compressed::is_compressed_rtf(source_bytes) {
        return Err(Error::UnsupportedSource(
            "compressed RTF needs a transport-aware rewrite",
        ));
    }
    if !edit.source.opaque().is_empty() {
        return Err(Error::UnsupportedSource(
            "canonical destination edits refuse unknown RTF destinations",
        ));
    }
    let semantic_delta = semantic_changes(&edit.operations, &[]);
    if semantic_delta.is_empty() {
        return Ok(Commit::new(
            edit.source.clone(),
            edit.source,
            false,
            operation_count,
            semantic_delta,
        ));
    }

    let mut model =
        crate::document::RtfDocument::parse_bytes_with_limits(source_bytes, edit.source.limits())?;
    for operation in &edit.operations {
        match operation {
            Operation::TableCellText { path, after, .. } => {
                model
                    .table_cell_mut(path)?
                    .set_text(Cow::Owned(after.clone()))?;
            },
            Operation::HeaderFooterText { target, after, .. } => {
                let paragraph = model_header_footer_paragraph_mut(&mut model, *target)?;
                paragraph.text = Cow::Owned(after.clone());
            },
            Operation::AnnotationText { index, after, .. } => {
                model.set_annotation_text(*index, Cow::Owned(after.clone()))?;
            },
            Operation::NoteText { index, after, .. } => {
                model.set_note_content(*index, Cow::Owned(after.clone()))?;
            },
            Operation::ShapeText { index, after, .. } => {
                model.set_body_shape_text(*index, Cow::Owned(after.clone()))?;
            },
            Operation::Text { .. }
            | Operation::Alignment { .. }
            | Operation::Bold { .. }
            | Operation::InsertParagraph { .. }
            | Operation::RootTransfer { .. } => return Err(Error::BodyDestinationConflict),
        }
    }

    let mut bytes = Vec::new();
    RtfWriter::new(&mut bytes)
        .write_document(&model)
        .map_err(|error| Error::Write(error.to_string()))?;
    let limit = edit.source.limits().max_source_bytes();
    if bytes.len() > limit {
        return Err(Error::InputTooLarge {
            observed: bytes.len(),
            limit,
        });
    }
    let snapshot = Snapshot::from_bytes_with_limits(&bytes, edit.source.limits())?;
    if snapshot.text() != edit.source.text() {
        return Err(Error::UnsupportedSource(
            "canonical destination edit changed the ordinary body story",
        ));
    }
    for operation in &edit.operations {
        match operation {
            Operation::TableCellText { path, after, .. } => {
                if table_cell(&snapshot, path)?.text() != after {
                    return Err(Error::UnsupportedSource(
                        "table-cell text did not survive RTF validation",
                    ));
                }
            },
            Operation::HeaderFooterText { target, after, .. } => {
                let actual = header_footer(&snapshot, *target)?
                    .paragraphs
                    .get(target.paragraph)
                    .ok_or(Error::DestinationOutOfRange("header/footer paragraph"))?
                    .text
                    .as_ref();
                if actual != after {
                    return Err(Error::UnsupportedSource(
                        "header/footer text did not survive RTF validation",
                    ));
                }
            },
            Operation::AnnotationText { index, after, .. } => {
                if annotation(&snapshot, *index)?.text != after.as_str() {
                    return Err(Error::UnsupportedSource(
                        "annotation text did not survive RTF validation",
                    ));
                }
            },
            Operation::NoteText { index, after, .. } => {
                if note(&snapshot, *index)?.content != after.as_str() {
                    return Err(Error::UnsupportedSource(
                        "note text did not survive RTF validation",
                    ));
                }
            },
            Operation::ShapeText { index, after, .. } => {
                if shape(&snapshot, *index)?.text != after.as_str() {
                    return Err(Error::UnsupportedSource(
                        "shape text did not survive RTF validation",
                    ));
                }
            },
            Operation::Text { .. }
            | Operation::Alignment { .. }
            | Operation::Bold { .. }
            | Operation::InsertParagraph { .. }
            | Operation::RootTransfer { .. } => return Err(Error::BodyDestinationConflict),
        }
    }
    Ok(Commit::new(
        edit.source,
        snapshot,
        true,
        operation_count,
        semantic_delta,
    ))
}

fn commit_root_transfer(edit: Edit, operation_count: usize) -> Result<Commit, Error> {
    let operation = edit
        .operations
        .first()
        .ok_or(Error::UnsupportedSource("missing ordinary-root transfer"))?;
    let Operation::RootTransfer {
        before,
        after,
        vocabulary: _,
        effect: _,
    } = operation
    else {
        return Err(Error::BodyDestinationConflict);
    };
    let source = edit
        .source
        .source_bytes()
        .ok_or(Error::UnsupportedSource("snapshot has no exact RTF source"))?;
    if source != before {
        return Err(Error::PatchConflict);
    }
    if crate::compressed::is_compressed_rtf(source) {
        return Err(Error::UnsupportedSource(
            "compressed RTF needs a transport-aware rewrite",
        ));
    }
    if !edit.source.opaque().is_empty() {
        return Err(Error::UnsupportedSource(
            "ordinary-root transfer refuses unknown target destinations",
        ));
    }
    let snapshot = Snapshot::from_bytes_with_limits(after, edit.source.limits())?;
    if !snapshot.opaque().is_empty() {
        return Err(Error::UnsupportedSource(
            "ordinary-root transfer produced unknown destinations",
        ));
    }
    let semantic_delta = semantic_changes(&edit.operations, &[]);
    let changed = source != after;
    Ok(Commit::new(
        edit.source,
        snapshot,
        changed,
        operation_count,
        semantic_delta,
    ))
}

fn table_cell<'a>(
    source: &'a Snapshot,
    path: &TableCellPath,
) -> Result<&'a crate::Cell<'a>, Error> {
    let root = path.root;
    let mut cell = source
        .tables()
        .get(root.table_index)
        .and_then(|table| table.rows().get(root.row_index))
        .and_then(|row| row.cells().get(root.cell_index))
        .ok_or(Error::DestinationOutOfRange("table cell"))?;
    for coordinate in &path.nested {
        cell = cell
            .nested_tables()
            .get(coordinate.table_index)
            .and_then(|nested| nested.table.rows().get(coordinate.row_index))
            .and_then(|row| row.cells().get(coordinate.cell_index))
            .ok_or(Error::DestinationOutOfRange("nested table cell"))?;
    }
    Ok(cell)
}

fn header_footer(
    source: &Snapshot,
    target: HeaderFooterParagraph,
) -> Result<&crate::HeaderFooter<'_>, Error> {
    source
        .sections()
        .get(target.section)
        .ok_or(Error::DestinationOutOfRange("section"))?
        .headers_footers
        .iter()
        .find(|candidate| candidate.header_type == target.kind)
        .ok_or(Error::DestinationOutOfRange("header/footer"))
}

fn annotation(source: &Snapshot, index: usize) -> Result<&crate::Annotation<'_>, Error> {
    source
        .annotations()
        .get(index)
        .ok_or(Error::DestinationOutOfRange("annotation"))
}

fn note(source: &Snapshot, index: usize) -> Result<&crate::Note<'_>, Error> {
    source
        .notes()
        .get(index)
        .ok_or(Error::DestinationOutOfRange("note"))
}

fn shape(source: &Snapshot, index: usize) -> Result<&crate::Shape<'_>, Error> {
    source
        .shapes()
        .get(index)
        .ok_or(Error::DestinationOutOfRange("shape"))
}

fn shape_has_active_link(shape: &crate::Shape<'_>) -> bool {
    shape
        .properties
        .iter()
        .any(|property| property.hyperlink.is_some())
        || shape.text_shapes.iter().any(shape_has_active_link)
        || shape
            .text_shape_groups
            .iter()
            .any(shape_group_has_active_link)
}

fn shape_group_has_active_link(group: &crate::ShapeGroup<'_>) -> bool {
    group.shapes.iter().any(shape_has_active_link)
        || group.groups.iter().any(shape_group_has_active_link)
}

fn model_header_footer_paragraph_mut<'a>(
    model: &'a mut crate::document::RtfDocument<'static>,
    target: HeaderFooterParagraph,
) -> Result<&'a mut crate::HeaderFooterParagraph<'static>, Error> {
    model
        .sections_mut()
        .get_mut(target.section)
        .ok_or(Error::DestinationOutOfRange("section"))?
        .headers_footers
        .iter_mut()
        .find(|candidate| candidate.header_type == target.kind)
        .ok_or(Error::DestinationOutOfRange("header/footer"))?
        .paragraphs
        .get_mut(target.paragraph)
        .ok_or(Error::DestinationOutOfRange("header/footer paragraph"))
}

fn table_cell_effect(path: &TableCellPath) -> String {
    let mut effect = format!(
        "table:{}:row:{}:cell:{}",
        path.root.table_index, path.root.row_index, path.root.cell_index
    );
    for coordinate in &path.nested {
        effect.push_str(":nested:");
        effect.push_str(&coordinate.table_index.to_string());
        effect.push(':');
        effect.push_str(&coordinate.row_index.to_string());
        effect.push(':');
        effect.push_str(&coordinate.cell_index.to_string());
    }
    effect.push_str(":text");
    effect
}

fn header_footer_effect(target: HeaderFooterParagraph) -> String {
    format!(
        "section:{}:{}:paragraph:{}:text",
        target.section,
        header_footer_kind_name(target.kind),
        target.paragraph
    )
}

fn annotation_effect(index: usize) -> String {
    format!("body:annotation:{index}:text")
}

fn note_effect(index: usize) -> String {
    format!("body:note:{index}:text")
}

fn shape_effect(index: usize) -> String {
    format!("body:shape:{index}:text")
}

const fn header_footer_kind_name(kind: HeaderFooterType) -> &'static str {
    match kind {
        HeaderFooterType::Header => "header",
        HeaderFooterType::Footer => "footer",
        HeaderFooterType::HeaderFirst => "header-first",
        HeaderFooterType::FooterFirst => "footer-first",
        HeaderFooterType::HeaderLeft => "header-left",
        HeaderFooterType::FooterLeft => "footer-left",
        HeaderFooterType::HeaderRight => "header-right",
        HeaderFooterType::FooterRight => "footer-right",
    }
}

fn validate_span(body: &str, span: TextSpan) -> Result<(), Error> {
    if span.end > body.len() {
        return Err(Error::SpanOutOfRange {
            end: span.end,
            length: body.len(),
        });
    }
    if !body.is_char_boundary(span.start) {
        return Err(Error::SpanNotOnCharacterBoundary {
            position: span.start,
        });
    }
    if !body.is_char_boundary(span.end) {
        return Err(Error::SpanNotOnCharacterBoundary { position: span.end });
    }
    Ok(())
}

fn spans_conflict(left: TextSpan, right: TextSpan) -> bool {
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

fn paragraph_range(source: &Snapshot, position: usize) -> Result<Range<usize>, Error> {
    let mut start = 0usize;
    for (paragraph_position, paragraph) in source.body().paragraphs().enumerate() {
        let end = start
            .checked_add(paragraph.len())
            .ok_or(Error::InputTooLarge {
                observed: usize::MAX,
                limit: source.limits().max_source_bytes(),
            })?;
        if paragraph_position == position {
            return Ok(start..end);
        }
        start = end.checked_add(1).ok_or(Error::InputTooLarge {
            observed: usize::MAX,
            limit: source.limits().max_source_bytes(),
        })?;
    }
    let count = source.paragraph_count();
    Err(Error::ParagraphOutOfRange { position, count })
}

fn project_text(
    source: &Snapshot,
    operations: &[Operation],
) -> Result<(String, Vec<(usize, TextSpan)>), Error> {
    let mut text_operations = operations
        .iter()
        .enumerate()
        .filter_map(|(operation_index, operation)| match operation {
            Operation::Text {
                span,
                after,
                before: _,
                structural: _,
            } => Some((operation_index, *span, after.as_str(), false)),
            Operation::InsertParagraph { span, text, .. } => {
                Some((operation_index, *span, text.as_str(), true))
            },
            Operation::Alignment { .. }
            | Operation::Bold { .. }
            | Operation::TableCellText { .. }
            | Operation::HeaderFooterText { .. }
            | Operation::AnnotationText { .. }
            | Operation::NoteText { .. }
            | Operation::ShapeText { .. }
            | Operation::RootTransfer { .. } => None,
        })
        .collect::<Vec<_>>();
    text_operations.sort_unstable_by_key(|(_, span, _, _)| (span.start, span.end));
    let source_text = source.text();
    let mut final_len = source_text.len();
    for (_, span, replacement, structural) in &text_operations {
        let replacement_len = replacement.len().saturating_add(usize::from(*structural));
        final_len = final_len
            .checked_sub(span.end - span.start)
            .and_then(|length| length.checked_add(replacement_len))
            .ok_or(Error::InputTooLarge {
                observed: usize::MAX,
                limit: source.limits().max_source_bytes(),
            })?;
    }
    let limit = source.limits().max_source_bytes();
    if final_len > limit {
        return Err(Error::InputTooLarge {
            observed: final_len,
            limit,
        });
    }
    let mut output = String::new();
    output
        .try_reserve_exact(final_len)
        .map_err(|_error| Error::Write("could not reserve replacement body text".to_string()))?;
    let mut cursor = 0usize;
    let mut projected = Vec::new();
    projected
        .try_reserve_exact(text_operations.len())
        .map_err(|_error| Error::Write("could not reserve projected text spans".to_string()))?;
    for (index, span, replacement, structural) in text_operations {
        output.push_str(
            source_text
                .get(cursor..span.start)
                .ok_or(Error::UnsupportedSource(
                    "staged text spans are not ordered UTF-8 boundaries",
                ))?,
        );
        let projected_start = output.len();
        if structural {
            output.push('\n');
        }
        output.push_str(replacement);
        projected.push((
            index,
            TextSpan {
                start: projected_start,
                end: output.len(),
            },
        ));
        cursor = span.end;
    }
    output.push_str(source_text.get(cursor..).ok_or(Error::UnsupportedSource(
        "staged text span ends outside the body",
    ))?);
    Ok((output, projected))
}

fn source_alignments(source: &Snapshot) -> Vec<Alignment> {
    source
        .body()
        .paragraphs()
        .map(|paragraph| paragraph.format().alignment())
        .collect()
}

fn uniform_body_bold(source: &Snapshot) -> Result<bool, Error> {
    let mut value = None;
    for run in source.body().runs() {
        let bold = run.format().bold();
        if value.is_some_and(|existing| existing != bold) {
            return Err(Error::UnsupportedSource(
                "the body has mixed character formatting",
            ));
        }
        value = Some(bold);
    }
    Ok(value.unwrap_or(false))
}

fn base_bold_for_edit(source: &Snapshot, operations: &[Operation]) -> Result<bool, Error> {
    let bold_spans = operations
        .iter()
        .filter_map(|operation| match operation {
            Operation::Bold { span, .. } => Some(*span),
            Operation::Text { .. }
            | Operation::Alignment { .. }
            | Operation::InsertParagraph { .. }
            | Operation::TableCellText { .. }
            | Operation::HeaderFooterText { .. }
            | Operation::AnnotationText { .. }
            | Operation::NoteText { .. }
            | Operation::ShapeText { .. }
            | Operation::RootTransfer { .. } => None,
        })
        .collect::<Vec<_>>();
    if bold_spans.is_empty() {
        return uniform_body_bold(source);
    }
    let mut body_position = 0usize;
    let mut base = None;
    for paragraph in source.body().paragraphs() {
        let mut run_position = body_position;
        for run in paragraph.runs() {
            let run_span = TextSpan {
                start: run_position,
                end: run_position.saturating_add(run.text().len()),
            };
            if !bold_spans
                .iter()
                .any(|span| spans_conflict(*span, run_span))
            {
                let value = run.format().bold();
                if base.is_some_and(|existing| existing != value) {
                    return Err(Error::UnsupportedSource(
                        "unselected body text has mixed bold state",
                    ));
                }
                base = Some(value);
            }
            run_position = run_span.end;
        }
        body_position = body_position
            .saturating_add(paragraph.len())
            .saturating_add(1);
    }
    Ok(base.unwrap_or_else(|| {
        operations
            .iter()
            .find_map(|operation| match operation {
                Operation::Bold { after, .. } => Some(*after),
                Operation::Text { .. }
                | Operation::Alignment { .. }
                | Operation::InsertParagraph { .. }
                | Operation::TableCellText { .. }
                | Operation::HeaderFooterText { .. }
                | Operation::AnnotationText { .. }
                | Operation::NoteText { .. }
                | Operation::ShapeText { .. }
                | Operation::RootTransfer { .. } => None,
            })
            .unwrap_or(false)
    }))
}

fn bold_for_span(source: &Snapshot, span: TextSpan) -> Result<bool, Error> {
    validate_span(source.text(), span)?;
    let mut body_position = 0usize;
    let mut covered = 0usize;
    let mut value = None;
    for paragraph in source.body().paragraphs() {
        let mut run_position = body_position;
        for run in paragraph.runs() {
            let run_end = run_position.saturating_add(run.text().len());
            let start = run_position.max(span.start);
            let end = run_end.min(span.end);
            if start < end {
                let bold = run.format().bold();
                if value.is_some_and(|existing| existing != bold) {
                    return Err(Error::UnsupportedSource(
                        "the selected character span has mixed bold state",
                    ));
                }
                value = Some(bold);
                covered = covered.saturating_add(end - start);
            }
            run_position = run_end;
        }
        body_position = body_position
            .saturating_add(paragraph.len())
            .saturating_add(1);
    }
    if covered != span.end.saturating_sub(span.start) {
        return Err(Error::UnsupportedSource(
            "the selected character span crosses non-text inline content",
        ));
    }
    value.ok_or(Error::UnsupportedSource(
        "the selected character span has no text run",
    ))
}

fn project_base_span(span: TextSpan, operations: &[Operation]) -> Result<TextSpan, Error> {
    Ok(TextSpan {
        start: project_base_position(span.start, operations)?,
        end: project_base_position(span.end, operations)?,
    })
}

fn project_base_position(position: usize, operations: &[Operation]) -> Result<usize, Error> {
    let mut changes = operations
        .iter()
        .filter_map(|operation| match operation {
            Operation::Text { span, after, .. } => Some((*span, after.len())),
            Operation::InsertParagraph { span, text, .. } => {
                Some((*span, text.len().saturating_add(1)))
            },
            Operation::Alignment { .. }
            | Operation::Bold { .. }
            | Operation::TableCellText { .. }
            | Operation::HeaderFooterText { .. }
            | Operation::AnnotationText { .. }
            | Operation::NoteText { .. }
            | Operation::ShapeText { .. }
            | Operation::RootTransfer { .. } => None,
        })
        .collect::<Vec<_>>();
    changes.sort_unstable_by_key(|(span, _)| (span.start, span.end));
    let mut source_cursor = 0usize;
    let mut projected_cursor = 0usize;
    for (span, replacement_len) in changes {
        if position <= span.start {
            return projected_cursor
                .checked_add(position.saturating_sub(source_cursor))
                .ok_or(Error::InputTooLarge {
                    observed: usize::MAX,
                    limit: usize::MAX,
                });
        }
        projected_cursor = projected_cursor
            .checked_add(span.start.saturating_sub(source_cursor))
            .and_then(|value| value.checked_add(replacement_len))
            .ok_or(Error::InputTooLarge {
                observed: usize::MAX,
                limit: usize::MAX,
            })?;
        source_cursor = span.end;
    }
    projected_cursor
        .checked_add(position.saturating_sub(source_cursor))
        .ok_or(Error::InputTooLarge {
            observed: usize::MAX,
            limit: usize::MAX,
        })
}

fn ordinary_body_source_span(
    source: &[u8],
    semantic_text: &str,
    limits: crate::ParseLimits,
) -> Result<Range<usize>, Error> {
    let lexical = if source.is_ascii() {
        std::str::from_utf8(source)
            .map(str::to_owned)
            .map_err(|error| Error::Write(error.to_string()))?
    } else {
        source.iter().map(|byte| char::from(*byte)).collect()
    };
    let arena = Bump::new();
    let mut lexer = crate::lexer::Lexer::new_with_limits(&lexical, &arena, limits);
    let (tokens, spans) = lexer.tokenize_with_spans()?;
    let mut depth = 0usize;
    let mut start = None;
    let mut end = None;
    for (token, span) in tokens.iter().zip(&spans) {
        match token {
            crate::lexer::Token::OpenBrace => {
                if depth == 1 && start.is_some() {
                    return Err(Error::UnsupportedSource(
                        "the body source is not one contiguous root-level span",
                    ));
                }
                depth = depth.checked_add(1).ok_or(Error::UnsupportedSource(
                    "RTF group nesting overflowed while locating the body",
                ))?;
            },
            crate::lexer::Token::CloseBrace => {
                if depth == 1 {
                    end = Some(span.start);
                    break;
                }
                depth = depth.checked_sub(1).ok_or(Error::UnsupportedSource(
                    "RTF group nesting underflowed while locating the body",
                ))?;
            },
            crate::lexer::Token::Text(_) if depth == 1 && start.is_none() => {
                start = Some(span.start);
            },
            crate::lexer::Token::Binary(_) if depth == 1 && start.is_some() => {
                return Err(Error::UnsupportedSource(
                    "the body source contains binary data",
                ));
            },
            crate::lexer::Token::Control(_)
            | crate::lexer::Token::Text(_)
            | crate::lexer::Token::Binary(_) => {},
        }
    }
    let root_end = end.ok_or(Error::UnsupportedSource(
        "RTF root group has no closing boundary",
    ))?;
    match start {
        Some(start_offset) => Ok(start_offset..root_end),
        None if semantic_text.is_empty() => Ok(root_end..root_end),
        None => Err(Error::UnsupportedSource(
            "the body has no literal source span for a lossless replacement",
        )),
    }
}

fn validate_opaque_preservation(
    source: &Snapshot,
    source_bytes: &[u8],
    body_span: &Range<usize>,
) -> Result<(), Error> {
    for node in source.opaque() {
        if !matches!(node.anchor(), crate::opaque::Anchor::Body(_)) {
            continue;
        }
        let opaque = node.source();
        if opaque.is_empty() {
            return Err(Error::UnsupportedSource(
                "an empty body-anchored opaque node cannot be located",
            ));
        }
        let retained_before_body = source_bytes
            .get(..body_span.start)
            .is_some_and(|prefix| prefix.windows(opaque.len()).any(|window| window == opaque));
        if !retained_before_body {
            return Err(Error::UnsupportedSource(
                "body-anchored opaque syntax is retained losslessly but not editable here",
            ));
        }
    }
    Ok(())
}

fn encoded_body_text(text: &str, limits: crate::ParseLimits) -> Result<Vec<u8>, Error> {
    let required = encoded_body_len(text)?;
    if required > limits.max_source_bytes() {
        return Err(Error::InputTooLarge {
            observed: required,
            limit: limits.max_source_bytes(),
        });
    }
    let mut output = Vec::new();
    output
        .try_reserve_exact(required)
        .map_err(|_error| Error::Write("could not reserve replacement RTF bytes".to_string()))?;
    RtfWriter::new(&mut output)
        .write_text(text)
        .map_err(|error| Error::Write(error.to_string()))?;
    Ok(output)
}

fn encoded_body_with_properties(
    text: &str,
    alignments: &[Alignment],
    base_bold: bool,
    bold_changes: &[(TextSpan, bool)],
    limits: crate::ParseLimits,
) -> Result<Vec<u8>, Error> {
    let extra = alignments
        .len()
        .checked_mul(4)
        .and_then(|bytes| bytes.checked_add(bold_changes.len().saturating_mul(6)))
        .ok_or(Error::InputTooLarge {
            observed: usize::MAX,
            limit: limits.max_source_bytes(),
        })?;
    let required = encoded_body_len(text)?
        .checked_add(extra)
        .ok_or(Error::InputTooLarge {
            observed: usize::MAX,
            limit: limits.max_source_bytes(),
        })?;
    if required > limits.max_source_bytes() {
        return Err(Error::InputTooLarge {
            observed: required,
            limit: limits.max_source_bytes(),
        });
    }
    let mut output = Vec::new();
    output
        .try_reserve_exact(required)
        .map_err(|_error| Error::Write("could not reserve formatted RTF bytes".to_string()))?;
    let mut paragraph = 0usize;
    let mut paragraph_start = 0usize;
    loop {
        let paragraph_end = text
            .get(paragraph_start..)
            .and_then(|remainder| remainder.find('\n'))
            .map_or(text.len(), |offset| paragraph_start.saturating_add(offset));
        if paragraph_start == text.len() && text.ends_with('\n') {
            break;
        }
        write_alignment(&mut output, alignments.get(paragraph).copied())?;
        let mut cursor = paragraph_start;
        let mut paragraph_changes = bold_changes
            .iter()
            .copied()
            .filter(|(span, _)| span.start >= paragraph_start && span.end <= paragraph_end)
            .collect::<Vec<_>>();
        paragraph_changes.sort_unstable_by_key(|(span, _)| (span.start, span.end));
        for (span, bold) in paragraph_changes {
            write_encoded_fragment(&mut output, text, cursor..span.start)?;
            write_bold(&mut output, bold);
            write_encoded_fragment(&mut output, text, span.start..span.end)?;
            write_bold(&mut output, base_bold);
            cursor = span.end;
        }
        write_encoded_fragment(&mut output, text, cursor..paragraph_end)?;
        paragraph = paragraph.saturating_add(1);
        if paragraph_end == text.len() {
            break;
        }
        output.extend_from_slice(br"\par ");
        paragraph_start = paragraph_end.saturating_add(1);
    }
    if paragraph != alignments.len() {
        return Err(Error::StructuralPropertyConflict);
    }
    Ok(output)
}

fn write_encoded_fragment(
    output: &mut Vec<u8>,
    text: &str,
    range: Range<usize>,
) -> Result<(), Error> {
    let fragment = text.get(range).ok_or(Error::UnsupportedSource(
        "property span is not a UTF-8 text boundary",
    ))?;
    RtfWriter::new(output)
        .write_text(fragment)
        .map_err(|error| Error::Write(error.to_string()))
}

fn write_bold(output: &mut Vec<u8>, bold: bool) {
    output.extend_from_slice(if bold { br"\b " } else { br"\b0 " });
}

fn write_alignment(output: &mut Vec<u8>, alignment: Option<Alignment>) -> Result<(), Error> {
    let bytes = match alignment.ok_or(Error::StructuralPropertyConflict)? {
        Alignment::Left => br"\ql ".as_slice(),
        Alignment::Right => br"\qr ".as_slice(),
        Alignment::Center => br"\qc ".as_slice(),
        Alignment::Justify => br"\qj ".as_slice(),
    };
    output.extend_from_slice(bytes);
    Ok(())
}

fn encoded_body_len(text: &str) -> Result<usize, Error> {
    text.chars().try_fold(0usize, |total, character| {
        let width = match character {
            '\\' | '{' | '}' => 2,
            '\n' | '\t' => 5,
            value if (value as u32) < 0x20 => 4,
            value if value.is_ascii() => 1,
            _ => 10,
        };
        total.checked_add(width).ok_or(Error::InputTooLarge {
            observed: usize::MAX,
            limit: usize::MAX,
        })
    })
}

fn splice_body(
    source: &[u8],
    span: Range<usize>,
    replacement: &[u8],
    limits: crate::ParseLimits,
) -> Result<Vec<u8>, Error> {
    let retained = source
        .len()
        .checked_sub(span.end.saturating_sub(span.start))
        .ok_or(Error::UnsupportedSource("body source span is invalid"))?;
    let total = retained
        .checked_add(replacement.len())
        .ok_or(Error::InputTooLarge {
            observed: usize::MAX,
            limit: limits.max_source_bytes(),
        })?;
    if total > limits.max_source_bytes() {
        return Err(Error::InputTooLarge {
            observed: total,
            limit: limits.max_source_bytes(),
        });
    }
    let mut output = Vec::new();
    output
        .try_reserve_exact(total)
        .map_err(|_error| Error::Write("could not reserve candidate RTF bytes".to_string()))?;
    output.extend_from_slice(source.get(..span.start).ok_or(Error::UnsupportedSource(
        "body source span starts outside the document",
    ))?);
    output.extend_from_slice(replacement);
    output.extend_from_slice(source.get(span.end..).ok_or(Error::UnsupportedSource(
        "body source span ends outside the document",
    ))?);
    Ok(output)
}

/// Deterministic facts about a published transaction.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Diagnostics {
    operation_count: usize,
    changed: bool,
}

impl Diagnostics {
    /// Number of staged semantic operations represented by the commit.
    #[must_use]
    pub const fn operation_count(self) -> usize {
        self.operation_count
    }

    /// Whether the transaction published a distinct snapshot.
    #[must_use]
    pub const fn changed(self) -> bool {
        self.changed
    }
}

/// Result of an atomically validated RTF edit.
pub struct Commit {
    snapshot: Snapshot,
    patch: Patch,
    diagnostics: Diagnostics,
}

impl Commit {
    fn new(
        before: Snapshot,
        snapshot: Snapshot,
        did_change: bool,
        operation_count: usize,
        semantic_delta: Vec<Change>,
    ) -> Self {
        Self {
            patch: Patch {
                before,
                after: snapshot.clone(),
                changes: semantic_delta.into_boxed_slice(),
            },
            snapshot,
            diagnostics: Diagnostics {
                operation_count,
                changed: did_change,
            },
        }
    }

    /// Returns the published immutable snapshot.
    #[must_use]
    pub const fn snapshot(&self) -> &Snapshot {
        &self.snapshot
    }

    /// Returns the exact-source-checked reversible patch.
    #[must_use]
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Returns deterministic commit diagnostics.
    #[must_use]
    pub const fn diagnostics(&self) -> Diagnostics {
        self.diagnostics
    }

    /// Consumes the commit and returns its immutable snapshot.
    #[must_use]
    pub fn into_snapshot(self) -> Snapshot {
        self.snapshot
    }
}

#[derive(Debug, Clone)]
enum Change {
    Text {
        span: TextSpan,
        after_span: TextSpan,
        before: String,
        after: String,
    },
    Alignment {
        position: usize,
        before: Alignment,
        after: Alignment,
    },
    Bold {
        span: TextSpan,
        after_span: TextSpan,
        before: bool,
        after: bool,
    },
    InsertParagraph {
        position: usize,
        span: TextSpan,
        after_span: TextSpan,
        text: String,
        removing: bool,
    },
    TableCellText {
        path: TableCellPath,
        before: String,
        after: String,
    },
    HeaderFooterText {
        target: HeaderFooterParagraph,
        before: String,
        after: String,
    },
    AnnotationText {
        index: usize,
        before: String,
        after: String,
    },
    NoteText {
        index: usize,
        before: String,
        after: String,
    },
    ShapeText {
        index: usize,
        before: String,
        after: String,
    },
    RootTransfer {
        vocabulary: &'static str,
        effect: String,
        before: Vec<u8>,
        after: Vec<u8>,
    },
}

impl Change {
    fn inverse(&self) -> Self {
        match self {
            Self::Text {
                span,
                after_span,
                before,
                after,
            } => Self::Text {
                span: *after_span,
                after_span: *span,
                before: after.clone(),
                after: before.clone(),
            },
            Self::Alignment {
                position,
                before,
                after,
            } => Self::Alignment {
                position: *position,
                before: *after,
                after: *before,
            },
            Self::Bold {
                span,
                after_span,
                before,
                after,
            } => Self::Bold {
                span: *after_span,
                after_span: *span,
                before: *after,
                after: *before,
            },
            Self::InsertParagraph {
                position,
                span,
                after_span,
                text,
                removing,
            } => Self::InsertParagraph {
                position: *position,
                span: *after_span,
                after_span: *span,
                text: text.clone(),
                removing: !removing,
            },
            Self::TableCellText {
                path,
                before,
                after,
            } => Self::TableCellText {
                path: path.clone(),
                before: after.clone(),
                after: before.clone(),
            },
            Self::HeaderFooterText {
                target,
                before,
                after,
            } => Self::HeaderFooterText {
                target: *target,
                before: after.clone(),
                after: before.clone(),
            },
            Self::AnnotationText {
                index,
                before,
                after,
            } => Self::AnnotationText {
                index: *index,
                before: after.clone(),
                after: before.clone(),
            },
            Self::NoteText {
                index,
                before,
                after,
            } => Self::NoteText {
                index: *index,
                before: after.clone(),
                after: before.clone(),
            },
            Self::ShapeText {
                index,
                before,
                after,
            } => Self::ShapeText {
                index: *index,
                before: after.clone(),
                after: before.clone(),
            },
            Self::RootTransfer {
                vocabulary,
                effect,
                before,
                after,
            } => Self::RootTransfer {
                vocabulary,
                effect: effect.clone(),
                before: after.clone(),
                after: before.clone(),
            },
        }
    }
}

fn semantic_changes(
    operations: &[Operation],
    projected_spans: &[(usize, TextSpan)],
) -> Vec<Change> {
    operations
        .iter()
        .enumerate()
        .filter_map(|(operation_index, operation)| match operation {
            Operation::Text {
                span,
                before,
                after,
                structural: _,
            } if before != after => {
                projected_spans
                    .iter()
                    .find_map(|(projected_index, projected)| {
                        (*projected_index == operation_index).then_some(Change::Text {
                            span: *span,
                            after_span: *projected,
                            before: before.clone(),
                            after: after.clone(),
                        })
                    })
            },
            Operation::Alignment {
                position,
                before,
                after,
            } if before != after => Some(Change::Alignment {
                position: *position,
                before: *before,
                after: *after,
            }),
            Operation::Bold {
                span,
                before,
                after,
            } if before != after => Some(Change::Bold {
                span: *span,
                after_span: project_base_span(*span, operations).ok()?,
                before: *before,
                after: *after,
            }),
            Operation::InsertParagraph {
                position,
                span,
                text,
            } => projected_spans
                .iter()
                .find_map(|(projected_index, projected)| {
                    (*projected_index == operation_index).then_some(Change::InsertParagraph {
                        position: *position,
                        span: *span,
                        after_span: *projected,
                        text: text.clone(),
                        removing: false,
                    })
                }),
            Operation::TableCellText {
                path,
                before,
                after,
            } if before != after => Some(Change::TableCellText {
                path: path.clone(),
                before: before.clone(),
                after: after.clone(),
            }),
            Operation::HeaderFooterText {
                target,
                before,
                after,
            } if before != after => Some(Change::HeaderFooterText {
                target: *target,
                before: before.clone(),
                after: after.clone(),
            }),
            Operation::AnnotationText {
                index,
                before,
                after,
            } if before != after => Some(Change::AnnotationText {
                index: *index,
                before: before.clone(),
                after: after.clone(),
            }),
            Operation::NoteText {
                index,
                before,
                after,
            } if before != after => Some(Change::NoteText {
                index: *index,
                before: before.clone(),
                after: after.clone(),
            }),
            Operation::ShapeText {
                index,
                before,
                after,
            } if before != after => Some(Change::ShapeText {
                index: *index,
                before: before.clone(),
                after: after.clone(),
            }),
            Operation::RootTransfer {
                vocabulary,
                effect,
                before,
                after,
            } if before != after => Some(Change::RootTransfer {
                vocabulary,
                effect: effect.clone(),
                before: before.clone(),
                after: after.clone(),
            }),
            Operation::Text { .. }
            | Operation::Alignment { .. }
            | Operation::Bold { .. }
            | Operation::TableCellText { .. }
            | Operation::HeaderFooterText { .. }
            | Operation::AnnotationText { .. }
            | Operation::NoteText { .. }
            | Operation::ShapeText { .. }
            | Operation::RootTransfer { .. } => None,
        })
        .collect()
}

/// In-memory exact-source patch with a durable shared semantic projection.
#[derive(Clone)]
pub struct Patch {
    before: Snapshot,
    after: Snapshot,
    changes: Box<[Change]>,
}

impl Patch {
    /// Applies this patch only to the exact source bytes from which it was made.
    ///
    /// # Errors
    /// Returns [`Error::PatchConflict`] when the supplied source bytes differ.
    pub fn apply(&self, source: &Snapshot) -> Result<Snapshot, Error> {
        if source.source_bytes() != self.before.source_bytes() {
            return Err(Error::PatchConflict);
        }
        Ok(self.after.clone())
    }

    /// Returns the patch that restores the accepted source bytes.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            before: self.after.clone(),
            after: self.before.clone(),
            changes: self.changes.iter().rev().map(Change::inverse).collect(),
        }
    }

    /// Conservative byte weight suitable for the common bounded history.
    #[must_use]
    pub fn history_weight(&self) -> u64 {
        let before = self.before.source_bytes().map_or(0, <[u8]>::len);
        let after = self.after.source_bytes().map_or(0, <[u8]>::len);
        u64::try_from(before)
            .unwrap_or(u64::MAX)
            .saturating_add(u64::try_from(after).unwrap_or(u64::MAX))
    }

    /// Converts this patch to the shared versioned deterministic-JSON patch.
    ///
    /// # Errors
    /// Returns an error when caller-selected wire limits cannot represent the
    /// semantic operations or their preconditions.
    pub fn to_durable(
        &self,
        limits: litchi_core::patch::PatchLimits,
    ) -> Result<
        litchi_core::patch::Patch<litchi_core::patch::Reversible>,
        litchi_core::patch::PatchError,
    > {
        use litchi_core::patch::{BlobBundle, ReversibleOperation};

        let before =
            self.before
                .source_bytes()
                .ok_or(litchi_core::patch::PatchError::InvalidText {
                    field: "RTF source artifact",
                })?;
        let after =
            self.after
                .source_bytes()
                .ok_or(litchi_core::patch::PatchError::InvalidText {
                    field: "RTF target artifact",
                })?;
        let operations = self
            .changes
            .iter()
            .map(|change| {
                let forward = durable_operation(limits, change, before)?;
                let inverse = durable_operation(limits, &change.inverse(), after)?;
                Ok(ReversibleOperation::new(forward, inverse))
            })
            .collect::<Result<Vec<_>, litchi_core::patch::PatchError>>()?;
        litchi_core::patch::Patch::<litchi_core::patch::Reversible>::new(
            limits,
            "litchi-rtf",
            operations,
            BlobBundle::new(limits.blobs()),
            BlobBundle::new(limits.blobs()),
        )
    }
}

fn durable_operation(
    limits: litchi_core::patch::PatchLimits,
    change: &Change,
    source: &[u8],
) -> Result<litchi_core::patch::PatchOperation, litchi_core::patch::PatchError> {
    let mut preconditions = BTreeMap::new();
    preconditions.insert(
        "artifact_sha256".to_string(),
        Value::String(litchi_core::patch::BlobId::of(source).as_hex()),
    );
    match change {
        Change::Text {
            span,
            before,
            after,
            after_span: _,
        } => {
            preconditions.insert("text".to_string(), Value::String(before.clone()));
            litchi_core::patch::PatchOperation::new(
                limits,
                "body-text.replace",
                format!("body:utf8:{}-{}", span.start, span.end),
                preconditions,
                Value::String(after.clone()),
            )
        },
        Change::Alignment {
            position,
            before,
            after,
        } => {
            preconditions.insert(
                "alignment".to_string(),
                Value::String(alignment_name(*before).to_string()),
            );
            litchi_core::patch::PatchOperation::new(
                limits,
                "paragraph-alignment.set",
                format!("body:paragraph:{position}"),
                preconditions,
                Value::String(alignment_name(*after).to_string()),
            )
        },
        Change::Bold {
            span,
            before,
            after,
            after_span: _,
        } => {
            preconditions.insert("bold".to_string(), Value::Bool(*before));
            litchi_core::patch::PatchOperation::new(
                limits,
                "character-bold.set",
                format!("body:utf8:{}-{}", span.start, span.end),
                preconditions,
                Value::Bool(*after),
            )
        },
        Change::InsertParagraph {
            position,
            text,
            removing,
            span: _,
            after_span: _,
        } => {
            preconditions.insert(
                "text".to_string(),
                Value::String(if *removing {
                    text.clone()
                } else {
                    String::new()
                }),
            );
            litchi_core::patch::PatchOperation::new(
                limits,
                if *removing {
                    "paragraph.remove-after"
                } else {
                    "paragraph.insert-after"
                },
                format!("body:paragraph:{position}"),
                preconditions,
                if *removing {
                    Value::Null
                } else {
                    Value::String(text.clone())
                },
            )
        },
        Change::TableCellText {
            path,
            before,
            after,
        } => {
            preconditions.insert("text".to_string(), Value::String(before.clone()));
            litchi_core::patch::PatchOperation::new(
                limits,
                "table-cell-text.replace",
                table_cell_effect(path),
                preconditions,
                Value::String(after.clone()),
            )
        },
        Change::HeaderFooterText {
            target,
            before,
            after,
        } => {
            preconditions.insert("text".to_string(), Value::String(before.clone()));
            litchi_core::patch::PatchOperation::new(
                limits,
                "header-footer-text.replace",
                header_footer_effect(*target),
                preconditions,
                Value::String(after.clone()),
            )
        },
        Change::AnnotationText {
            index,
            before,
            after,
        } => {
            preconditions.insert("text".to_string(), Value::String(before.clone()));
            litchi_core::patch::PatchOperation::new(
                limits,
                "annotation-text.replace",
                annotation_effect(*index),
                preconditions,
                Value::String(after.clone()),
            )
        },
        Change::NoteText {
            index,
            before,
            after,
        } => {
            preconditions.insert("text".to_string(), Value::String(before.clone()));
            litchi_core::patch::PatchOperation::new(
                limits,
                "note-text.replace",
                note_effect(*index),
                preconditions,
                Value::String(after.clone()),
            )
        },
        Change::ShapeText {
            index,
            before,
            after,
        } => {
            preconditions.insert("text".to_string(), Value::String(before.clone()));
            litchi_core::patch::PatchOperation::new(
                limits,
                "shape-text.replace",
                shape_effect(*index),
                preconditions,
                Value::String(after.clone()),
            )
        },
        Change::RootTransfer {
            vocabulary,
            effect,
            before: _,
            after,
        } => {
            preconditions.insert("feature".to_string(), Value::String(effect.clone()));
            litchi_core::patch::PatchOperation::new(
                limits,
                *vocabulary,
                effect.clone(),
                preconditions,
                Value::String(hex_encode(after)),
            )
        },
    }
}

pub(crate) fn apply_durable<Mode>(
    source: &Snapshot,
    patch: &litchi_core::patch::Patch<Mode>,
) -> Result<Snapshot, Error> {
    if patch.format() != "litchi-rtf" || !patch.blobs().is_empty() {
        return Err(Error::DurablePatch(
            "unsupported format or blob bundle".to_string(),
        ));
    }
    if patch.operations().is_empty() {
        return Ok(source.clone());
    }
    let source_bytes = source
        .source_bytes()
        .ok_or(Error::UnsupportedSource("snapshot has no exact RTF source"))?;
    let source_hash = litchi_core::patch::BlobId::of(source_bytes).as_hex();
    let mut edit = source.edit();
    for operation in patch.operations() {
        if operation.preconditions.len() != 2 {
            return Err(Error::DurablePatch(
                "operation must contain exactly two preconditions".to_string(),
            ));
        }
        let expected_hash = operation
            .preconditions
            .get("artifact_sha256")
            .and_then(Value::as_str)
            .ok_or_else(|| Error::DurablePatch("missing artifact digest".to_string()))?;
        if expected_hash != source_hash {
            return Err(Error::PatchConflict);
        }
        match operation.op.as_str() {
            "body-text.replace" => {
                let span = parse_text_target(&operation.target)?;
                validate_span(source.text(), span)?;
                let expected = operation
                    .preconditions
                    .get("text")
                    .and_then(Value::as_str)
                    .ok_or_else(|| Error::DurablePatch("missing text precondition".to_string()))?;
                let actual = source.text().get(span.start..span.end).ok_or(
                    Error::SpanNotOnCharacterBoundary {
                        position: span.start,
                    },
                )?;
                if actual != expected {
                    return Err(Error::StalePrecondition("body text differs"));
                }
                let replacement = operation.value.as_str().ok_or_else(|| {
                    Error::DurablePatch("text value must be a string".to_string())
                })?;
                edit.replace_text(span, replacement)?;
            },
            "paragraph-alignment.set" => {
                let position = parse_paragraph_target(&operation.target)?;
                let expected = operation
                    .preconditions
                    .get("alignment")
                    .and_then(Value::as_str)
                    .ok_or_else(|| {
                        Error::DurablePatch("missing alignment precondition".to_string())
                    })?;
                let paragraphs = source.body().paragraphs().collect::<Vec<_>>();
                let count = paragraphs.len();
                let actual = paragraphs
                    .get(position)
                    .ok_or(Error::ParagraphOutOfRange { position, count })?
                    .format()
                    .alignment();
                if alignment_name(actual) != expected {
                    return Err(Error::StalePrecondition("paragraph alignment differs"));
                }
                let replacement = operation
                    .value
                    .as_str()
                    .and_then(parse_alignment)
                    .ok_or_else(|| Error::DurablePatch("invalid alignment value".to_string()))?;
                edit.set_paragraph_alignment(position, replacement)?;
            },
            "character-bold.set" => {
                let span = parse_text_target(&operation.target)?;
                let expected = operation
                    .preconditions
                    .get("bold")
                    .and_then(Value::as_bool)
                    .ok_or_else(|| Error::DurablePatch("missing bold precondition".to_string()))?;
                if bold_for_span(source, span)? != expected {
                    return Err(Error::StalePrecondition("character bold state differs"));
                }
                let replacement = operation
                    .value
                    .as_bool()
                    .ok_or_else(|| Error::DurablePatch("bold value must be Boolean".to_string()))?;
                edit.set_text_bold(span, replacement)?;
            },
            "paragraph.insert-after" => {
                let position = parse_paragraph_target(&operation.target)?;
                let replacement = operation.value.as_str().ok_or_else(|| {
                    Error::DurablePatch("inserted paragraph text must be a string".to_string())
                })?;
                edit.insert_paragraph_after(position, replacement)?;
            },
            "paragraph.remove-after" => {
                let position = parse_paragraph_target(&operation.target)?;
                let expected = operation
                    .preconditions
                    .get("text")
                    .and_then(Value::as_str)
                    .ok_or_else(|| {
                        Error::DurablePatch("missing inserted paragraph precondition".to_string())
                    })?;
                if !operation.value.is_null() {
                    return Err(Error::DurablePatch(
                        "paragraph removal value must be null".to_string(),
                    ));
                }
                edit.remove_paragraph_after(position, expected)?;
            },
            "table-cell-text.replace" => {
                let cell_path = parse_table_cell_target(&operation.target)?;
                let expected = operation
                    .preconditions
                    .get("text")
                    .and_then(Value::as_str)
                    .ok_or_else(|| {
                        Error::DurablePatch("missing table-cell text precondition".to_string())
                    })?;
                if table_cell(source, &cell_path)?.text() != expected {
                    return Err(Error::StalePrecondition("table-cell text differs"));
                }
                let replacement = operation.value.as_str().ok_or_else(|| {
                    Error::DurablePatch("table-cell text value must be a string".to_string())
                })?;
                edit.set_table_cell_text(cell_path, replacement)?;
            },
            "header-footer-text.replace" => {
                let target = parse_header_footer_target(&operation.target)?;
                let expected = operation
                    .preconditions
                    .get("text")
                    .and_then(Value::as_str)
                    .ok_or_else(|| {
                        Error::DurablePatch("missing header/footer text precondition".to_string())
                    })?;
                let actual = header_footer(source, target)?
                    .paragraphs
                    .get(target.paragraph)
                    .ok_or(Error::DestinationOutOfRange("header/footer paragraph"))?
                    .text
                    .as_ref();
                if actual != expected {
                    return Err(Error::StalePrecondition("header/footer text differs"));
                }
                let replacement = operation.value.as_str().ok_or_else(|| {
                    Error::DurablePatch("header/footer text value must be a string".to_string())
                })?;
                edit.set_header_footer_text(target, replacement)?;
            },
            "annotation-text.replace" => {
                let index = parse_annotation_target(&operation.target)?;
                let expected = operation
                    .preconditions
                    .get("text")
                    .and_then(Value::as_str)
                    .ok_or_else(|| {
                        Error::DurablePatch("missing annotation text precondition".to_string())
                    })?;
                if annotation(source, index)?.text != expected {
                    return Err(Error::StalePrecondition("annotation text differs"));
                }
                let replacement = operation.value.as_str().ok_or_else(|| {
                    Error::DurablePatch("annotation text value must be a string".to_string())
                })?;
                edit.set_annotation_text(index, replacement)?;
            },
            "note-text.replace" => {
                let index = parse_note_target(&operation.target)?;
                let expected = operation
                    .preconditions
                    .get("text")
                    .and_then(Value::as_str)
                    .ok_or_else(|| {
                        Error::DurablePatch("missing note text precondition".to_string())
                    })?;
                if note(source, index)?.content != expected {
                    return Err(Error::StalePrecondition("note text differs"));
                }
                let replacement = operation.value.as_str().ok_or_else(|| {
                    Error::DurablePatch("note text value must be a string".to_string())
                })?;
                edit.set_note_text(index, replacement)?;
            },
            "shape-text.replace" => {
                let index = parse_shape_target(&operation.target)?;
                let expected = operation
                    .preconditions
                    .get("text")
                    .and_then(Value::as_str)
                    .ok_or_else(|| {
                        Error::DurablePatch("missing shape text precondition".to_string())
                    })?;
                if shape(source, index)?.text != expected {
                    return Err(Error::StalePrecondition("shape text differs"));
                }
                let replacement = operation.value.as_str().ok_or_else(|| {
                    Error::DurablePatch("shape text value must be a string".to_string())
                })?;
                edit.set_shape_text(index, replacement)?;
            },
            vocabulary if is_root_transfer_vocabulary(vocabulary) => {
                let expected_feature = operation
                    .preconditions
                    .get("feature")
                    .and_then(Value::as_str)
                    .ok_or_else(|| {
                        Error::DurablePatch("missing feature precondition".to_string())
                    })?;
                if expected_feature != operation.target {
                    return Err(Error::StalePrecondition("ordinary-root feature differs"));
                }
                let encoded = operation.value.as_str().ok_or_else(|| {
                    Error::DurablePatch("ordinary-root value must be hexadecimal".to_string())
                })?;
                let after = hex_decode(encoded, source.limits().max_source_bytes())?;
                edit.stage_root_transfer(
                    root_transfer_vocabulary(vocabulary)?,
                    operation.target.clone(),
                    after,
                )?;
            },
            _ => {
                return Err(Error::DurablePatch(
                    "unsupported operation vocabulary".to_string(),
                ));
            },
        }
    }
    edit.commit().map(Commit::into_snapshot)
}

fn is_root_transfer_vocabulary(vocabulary: &str) -> bool {
    matches!(
        vocabulary,
        "field.transfer"
            | "nested-table.transfer"
            | "list.transfer"
            | "style.transfer"
            | "object.transfer"
            | "shape.transfer"
    )
}

fn root_transfer_vocabulary(vocabulary: &str) -> Result<&'static str, Error> {
    match vocabulary {
        "field.transfer" => Ok("field.transfer"),
        "nested-table.transfer" => Ok("nested-table.transfer"),
        "list.transfer" => Ok("list.transfer"),
        "style.transfer" => Ok("style.transfer"),
        "object.transfer" => Ok("object.transfer"),
        "shape.transfer" => Ok("shape.transfer"),
        _ => Err(Error::DurablePatch(
            "unsupported ordinary-root vocabulary".to_string(),
        )),
    }
}

fn hex_encode(bytes: &[u8]) -> String {
    let mut output = String::with_capacity(bytes.len().saturating_mul(2));
    for byte in bytes {
        output.push(hex_character(byte >> 4));
        output.push(hex_character(byte & 0x0f));
    }
    output
}

fn hex_decode(input: &str, limit: usize) -> Result<Vec<u8>, Error> {
    if !input.len().is_multiple_of(2) {
        return Err(Error::DurablePatch(
            "ordinary-root hexadecimal value has odd length".to_string(),
        ));
    }
    let observed = input.len() / 2;
    if observed > limit {
        return Err(Error::InputTooLarge { observed, limit });
    }
    let mut output = Vec::with_capacity(observed);
    let mut digits = input.bytes();
    while let Some(high) = digits.next() {
        let low = digits.next().ok_or_else(|| {
            Error::DurablePatch("ordinary-root hexadecimal value has odd length".to_string())
        })?;
        let high_nibble = hex_digit(high)?;
        let low_nibble = hex_digit(low)?;
        output.push((high_nibble << 4) | low_nibble);
    }
    Ok(output)
}

fn hex_character(value: u8) -> char {
    match value {
        0..=9 => char::from(b'0' + value),
        _ => char::from(b'a' + value - 10),
    }
}

fn hex_digit(value: u8) -> Result<u8, Error> {
    match value {
        b'0'..=b'9' => Ok(value - b'0'),
        b'a'..=b'f' => Ok(value - b'a' + 10),
        b'A'..=b'F' => Ok(value - b'A' + 10),
        _ => Err(Error::DurablePatch(
            "ordinary-root value contains non-hexadecimal data".to_string(),
        )),
    }
}

fn parse_text_target(target: &str) -> Result<TextSpan, Error> {
    let coordinates = target
        .strip_prefix("body:utf8:")
        .ok_or_else(|| Error::DurablePatch("invalid text target".to_string()))?;
    let (start_text, end_text) = coordinates
        .split_once('-')
        .ok_or_else(|| Error::DurablePatch("invalid text target span".to_string()))?;
    let start = start_text
        .parse::<usize>()
        .map_err(|_error| Error::DurablePatch("invalid text target start".to_string()))?;
    let end = end_text
        .parse::<usize>()
        .map_err(|_error| Error::DurablePatch("invalid text target end".to_string()))?;
    TextSpan::new(start, end)
}

fn parse_paragraph_target(target: &str) -> Result<usize, Error> {
    target
        .strip_prefix("body:paragraph:")
        .ok_or_else(|| Error::DurablePatch("invalid paragraph target".to_string()))?
        .parse::<usize>()
        .map_err(|_error| Error::DurablePatch("invalid paragraph position".to_string()))
}

fn parse_table_cell_target(target: &str) -> Result<TableCellPath, Error> {
    let mut parts = target.split(':');
    if parts.next() != Some("table") {
        return Err(Error::DurablePatch("invalid table-cell target".to_string()));
    }
    let table_index = parse_target_position(parts.next(), "table position")?;
    if parts.next() != Some("row") {
        return Err(Error::DurablePatch(
            "invalid table-cell row target".to_string(),
        ));
    }
    let row_index = parse_target_position(parts.next(), "row position")?;
    if parts.next() != Some("cell") {
        return Err(Error::DurablePatch(
            "invalid table-cell column target".to_string(),
        ));
    }
    let cell_index = parse_target_position(parts.next(), "cell position")?;
    let mut path = TableCellPath::outer(table_index, row_index, cell_index);
    loop {
        match parts.next() {
            Some("nested") => {
                let nested_table = parse_target_position(parts.next(), "nested table position")?;
                let nested_row = parse_target_position(parts.next(), "nested row position")?;
                let nested_cell = parse_target_position(parts.next(), "nested cell position")?;
                path = path.with_nested(crate::TableCellCoordinate {
                    table_index: nested_table,
                    row_index: nested_row,
                    cell_index: nested_cell,
                });
            },
            Some("text") if parts.next().is_none() => return Ok(path),
            _ => {
                return Err(Error::DurablePatch(
                    "invalid nested table-cell target".to_string(),
                ));
            },
        }
    }
}

fn parse_header_footer_target(target: &str) -> Result<HeaderFooterParagraph, Error> {
    let mut parts = target.split(':');
    if parts.next() != Some("section") {
        return Err(Error::DurablePatch(
            "invalid header/footer section target".to_string(),
        ));
    }
    let section = parse_target_position(parts.next(), "section position")?;
    let kind = parts
        .next()
        .and_then(parse_header_footer_kind)
        .ok_or_else(|| Error::DurablePatch("invalid header/footer kind".to_string()))?;
    if parts.next() != Some("paragraph") {
        return Err(Error::DurablePatch(
            "invalid header/footer paragraph target".to_string(),
        ));
    }
    let paragraph = parse_target_position(parts.next(), "header/footer paragraph position")?;
    if parts.next() != Some("text") || parts.next().is_some() {
        return Err(Error::DurablePatch(
            "invalid header/footer text target".to_string(),
        ));
    }
    Ok(HeaderFooterParagraph::new(section, kind, paragraph))
}

fn parse_annotation_target(target: &str) -> Result<usize, Error> {
    parse_indexed_story_target(target, "annotation")
}

fn parse_note_target(target: &str) -> Result<usize, Error> {
    parse_indexed_story_target(target, "note")
}

fn parse_shape_target(target: &str) -> Result<usize, Error> {
    parse_indexed_story_target(target, "shape")
}

fn parse_indexed_story_target(target: &str, owner: &'static str) -> Result<usize, Error> {
    let mut parts = target.split(':');
    if parts.next() != Some("body") || parts.next() != Some(owner) {
        return Err(Error::DurablePatch(format!("invalid {owner} text target")));
    }
    let index = parse_target_position(parts.next(), "story position")?;
    if parts.next() != Some("text") || parts.next().is_some() {
        return Err(Error::DurablePatch(format!("invalid {owner} text target")));
    }
    Ok(index)
}

fn parse_target_position(value: Option<&str>, name: &'static str) -> Result<usize, Error> {
    value
        .ok_or_else(|| Error::DurablePatch(format!("missing {name}")))?
        .parse::<usize>()
        .map_err(|_error| Error::DurablePatch(format!("invalid {name}")))
}

fn parse_header_footer_kind(value: &str) -> Option<HeaderFooterType> {
    match value {
        "header" => Some(HeaderFooterType::Header),
        "footer" => Some(HeaderFooterType::Footer),
        "header-first" => Some(HeaderFooterType::HeaderFirst),
        "footer-first" => Some(HeaderFooterType::FooterFirst),
        "header-left" => Some(HeaderFooterType::HeaderLeft),
        "footer-left" => Some(HeaderFooterType::FooterLeft),
        "header-right" => Some(HeaderFooterType::HeaderRight),
        "footer-right" => Some(HeaderFooterType::FooterRight),
        _ => None,
    }
}

const fn alignment_name(alignment: Alignment) -> &'static str {
    match alignment {
        Alignment::Left => "left",
        Alignment::Right => "right",
        Alignment::Center => "center",
        Alignment::Justify => "justify",
    }
}

fn parse_alignment(value: &str) -> Option<Alignment> {
    match value {
        "left" => Some(Alignment::Left),
        "right" => Some(Alignment::Right),
        "center" => Some(Alignment::Center),
        "justify" => Some(Alignment::Justify),
        _ => None,
    }
}
