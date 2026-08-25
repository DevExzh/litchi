//! Bounded immutable RTF body-text and paragraph-property transactions.
//!
//! Operations resolve against one immutable semantic body. Disjoint UTF-8
//! spans, paragraph alignment, and dependency-free local layout facets compose
//! in bounded atomic commits. Text edits retain a deliberately narrow uniform
//! formatting closure. Layout edits instead rewrite checked paragraph/run
//! snapshots and verify every unrelated property after reopening. Both seams
//! refuse body-anchored opaque syntax, tables, and positioned content.
//! Canonical retained-story edits cover checked table cells, headers/footers,
//! comments, notes, and root shape text frames while refusing unknown
//! destinations and dependent positioned content.

use crate::{
    Alignment, Document, HeaderFooterType, RtfError, RtfWriter, TableCellPath, UnderlineStyle,
};
use bumpalo::Bump;
use serde_json::Value;
use std::borrow::Cow;
use std::collections::BTreeMap;
use std::fmt;
use std::io::{self, Write as _};
use std::ops::Range;

mod composition;
mod picture_payload;
mod transfer;
pub use composition::{
    Composition, CompositionConflict, CompositionError, CompositionLimits, ConflictSet,
    DurableComposition, DurableMergePlan, MergePlan, MergeResolution, Prepared,
};
pub use picture_payload::{
    MAX_PICTURE_PAYLOAD_OPERATIONS, MAX_PICTURE_REMOVAL_OPERATIONS, PicturePayloadReplacement,
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

/// Exact bytes inserted for a raw ordinary-paragraph split.
const PARAGRAPH_SPLIT_BYTES: &[u8] = br"\par ";

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

/// One checked replacement in a bounded body-paragraph batch.
///
/// The zero-based paragraph position is resolved against the immutable source
/// snapshot when the complete batch is staged.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ParagraphTextReplacement {
    position: usize,
    replacement: String,
}

/// The dependency-free local layout facets of one ordinary body paragraph.
///
/// Values are the effective explicit RTF values after parsing. Zero spacing
/// or indentation and `false` pagination flags represent the cleared state.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct ParagraphLayout {
    space_before: i32,
    space_after: i32,
    left_indent: i32,
    right_indent: i32,
    first_line_indent: i32,
    keep_together: bool,
    keep_with_next: bool,
    page_break_before: bool,
}

impl ParagraphLayout {
    pub(crate) fn from_raw(paragraph: &crate::types::Paragraph) -> Self {
        Self {
            space_before: paragraph.spacing.before,
            space_after: paragraph.spacing.after,
            left_indent: paragraph.indentation.left,
            right_indent: paragraph.indentation.right,
            first_line_indent: paragraph.indentation.first_line,
            keep_together: paragraph.keep_together,
            keep_with_next: paragraph.keep_next,
            page_break_before: paragraph.page_break_before,
        }
    }

    /// Space before the paragraph in twips.
    #[must_use]
    pub const fn space_before(self) -> i32 {
        self.space_before
    }

    /// Space after the paragraph in twips.
    #[must_use]
    pub const fn space_after(self) -> i32 {
        self.space_after
    }

    /// Physical left indentation in twips.
    #[must_use]
    pub const fn left_indent(self) -> i32 {
        self.left_indent
    }

    /// Physical right indentation in twips.
    #[must_use]
    pub const fn right_indent(self) -> i32 {
        self.right_indent
    }

    /// First-line indentation in twips. Negative values are hanging indents.
    #[must_use]
    pub const fn first_line_indent(self) -> i32 {
        self.first_line_indent
    }

    /// Whether the paragraph requests staying on one page.
    #[must_use]
    pub const fn keep_together(self) -> bool {
        self.keep_together
    }

    /// Whether the paragraph requests staying with its successor.
    #[must_use]
    pub const fn keep_with_next(self) -> bool {
        self.keep_with_next
    }

    /// Whether the paragraph requests a page break before itself.
    #[must_use]
    pub const fn page_break_before(self) -> bool {
        self.page_break_before
    }
}

/// Typed partial update for dependency-free local paragraph layout.
///
/// An omitted field is preserved. Supplying zero for spacing or indentation,
/// or `false` for a pagination flag, clears that explicit semantic value.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub struct ParagraphLayoutPatch {
    space_before: Option<i32>,
    space_after: Option<i32>,
    left_indent: Option<i32>,
    right_indent: Option<i32>,
    first_line_indent: Option<i32>,
    keep_together: Option<bool>,
    keep_with_next: Option<bool>,
    page_break_before: Option<bool>,
}

impl ParagraphLayoutPatch {
    /// Creates an empty delta. At least one field is required when staging.
    #[must_use]
    pub const fn new() -> Self {
        Self {
            space_before: None,
            space_after: None,
            left_indent: None,
            right_indent: None,
            first_line_indent: None,
            keep_together: None,
            keep_with_next: None,
            page_break_before: None,
        }
    }

    #[must_use]
    pub const fn with_space_before(mut self, twips: i32) -> Self {
        self.space_before = Some(twips);
        self
    }

    /// Clears explicit space before the paragraph.
    #[must_use]
    pub const fn clear_space_before(self) -> Self {
        self.with_space_before(0)
    }

    #[must_use]
    pub const fn with_space_after(mut self, twips: i32) -> Self {
        self.space_after = Some(twips);
        self
    }

    /// Clears explicit space after the paragraph.
    #[must_use]
    pub const fn clear_space_after(self) -> Self {
        self.with_space_after(0)
    }

    #[must_use]
    pub const fn with_left_indent(mut self, twips: i32) -> Self {
        self.left_indent = Some(twips);
        self
    }

    /// Clears explicit physical left indentation.
    #[must_use]
    pub const fn clear_left_indent(self) -> Self {
        self.with_left_indent(0)
    }

    #[must_use]
    pub const fn with_right_indent(mut self, twips: i32) -> Self {
        self.right_indent = Some(twips);
        self
    }

    /// Clears explicit physical right indentation.
    #[must_use]
    pub const fn clear_right_indent(self) -> Self {
        self.with_right_indent(0)
    }

    #[must_use]
    pub const fn with_first_line_indent(mut self, twips: i32) -> Self {
        self.first_line_indent = Some(twips);
        self
    }

    /// Clears explicit first-line or hanging indentation.
    #[must_use]
    pub const fn clear_first_line_indent(self) -> Self {
        self.with_first_line_indent(0)
    }

    #[must_use]
    pub const fn with_keep_together(mut self, keep: bool) -> Self {
        self.keep_together = Some(keep);
        self
    }

    /// Clears the local keep-together request.
    #[must_use]
    pub const fn clear_keep_together(self) -> Self {
        self.with_keep_together(false)
    }

    #[must_use]
    pub const fn with_keep_with_next(mut self, keep: bool) -> Self {
        self.keep_with_next = Some(keep);
        self
    }

    /// Clears the local keep-with-next request.
    #[must_use]
    pub const fn clear_keep_with_next(self) -> Self {
        self.with_keep_with_next(false)
    }

    #[must_use]
    pub const fn with_page_break_before(mut self, page_break: bool) -> Self {
        self.page_break_before = Some(page_break);
        self
    }

    /// Clears the local page-break-before request.
    #[must_use]
    pub const fn clear_page_break_before(self) -> Self {
        self.with_page_break_before(false)
    }

    const fn fields(self) -> LayoutFields {
        let mut fields = 0u8;
        if self.space_before.is_some() {
            fields |= LayoutFields::SPACE_BEFORE;
        }
        if self.space_after.is_some() {
            fields |= LayoutFields::SPACE_AFTER;
        }
        if self.left_indent.is_some() {
            fields |= LayoutFields::LEFT_INDENT;
        }
        if self.right_indent.is_some() {
            fields |= LayoutFields::RIGHT_INDENT;
        }
        if self.first_line_indent.is_some() {
            fields |= LayoutFields::FIRST_LINE_INDENT;
        }
        if self.keep_together.is_some() {
            fields |= LayoutFields::KEEP_TOGETHER;
        }
        if self.keep_with_next.is_some() {
            fields |= LayoutFields::KEEP_WITH_NEXT;
        }
        if self.page_break_before.is_some() {
            fields |= LayoutFields::PAGE_BREAK_BEFORE;
        }
        LayoutFields(fields)
    }

    fn apply(self, layout: &mut ParagraphLayout) {
        if let Some(value) = self.space_before {
            layout.space_before = value;
        }
        if let Some(value) = self.space_after {
            layout.space_after = value;
        }
        if let Some(value) = self.left_indent {
            layout.left_indent = value;
        }
        if let Some(value) = self.right_indent {
            layout.right_indent = value;
        }
        if let Some(value) = self.first_line_indent {
            layout.first_line_indent = value;
        }
        if let Some(value) = self.keep_together {
            layout.keep_together = value;
        }
        if let Some(value) = self.keep_with_next {
            layout.keep_with_next = value;
        }
        if let Some(value) = self.page_break_before {
            layout.page_break_before = value;
        }
    }
}

/// One immutable-source selector and typed local paragraph-layout delta.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct ParagraphLayoutUpdate {
    position: usize,
    patch: ParagraphLayoutPatch,
}

impl ParagraphLayoutUpdate {
    #[must_use]
    pub const fn new(position: usize, patch: ParagraphLayoutPatch) -> Self {
        Self { position, patch }
    }

    #[must_use]
    pub const fn position(self) -> usize {
        self.position
    }

    #[must_use]
    pub const fn patch(self) -> ParagraphLayoutPatch {
        self.patch
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
struct LayoutFields(u8);

impl LayoutFields {
    const SPACE_BEFORE: u8 = 1 << 0;
    const SPACE_AFTER: u8 = 1 << 1;
    const LEFT_INDENT: u8 = 1 << 2;
    const RIGHT_INDENT: u8 = 1 << 3;
    const FIRST_LINE_INDENT: u8 = 1 << 4;
    const KEEP_TOGETHER: u8 = 1 << 5;
    const KEEP_WITH_NEXT: u8 = 1 << 6;
    const PAGE_BREAK_BEFORE: u8 = 1 << 7;

    const fn is_empty(self) -> bool {
        self.0 == 0
    }

    const fn overlaps(self, other: Self) -> bool {
        self.0 & other.0 != 0
    }
}

fn layout_effect_keys(position: usize, fields: LayoutFields) -> Vec<String> {
    let mut effects = Vec::new();
    for (bit, name) in [
        (LayoutFields::SPACE_BEFORE, "space-before"),
        (LayoutFields::SPACE_AFTER, "space-after"),
        (LayoutFields::LEFT_INDENT, "left-indent"),
        (LayoutFields::RIGHT_INDENT, "right-indent"),
        (LayoutFields::FIRST_LINE_INDENT, "first-line-indent"),
        (LayoutFields::KEEP_TOGETHER, "keep-together"),
        (LayoutFields::KEEP_WITH_NEXT, "keep-with-next"),
        (LayoutFields::PAGE_BREAK_BEFORE, "page-break-before"),
    ] {
        if fields.0 & bit != 0 {
            effects.push(format!("body:paragraph:{position}:layout:{name}"));
        }
    }
    effects
}

fn layout_operation_conflicts(
    operation: &Operation,
    position: usize,
    fields: LayoutFields,
) -> bool {
    matches!(
        operation,
        Operation::ParagraphLayout {
            position: existing_position,
            fields: existing_fields,
            ..
        } if *existing_position == position && existing_fields.overlaps(fields)
    )
}

fn ensure_paragraph_layout_source(source: &Snapshot) -> Result<(), Error> {
    let source_bytes = source
        .source_bytes()
        .ok_or(Error::UnsupportedSource("snapshot has no exact RTF source"))?;
    if crate::compressed::is_compressed_rtf(source_bytes) {
        return Err(Error::UnsupportedSource(
            "compressed RTF needs a transport-aware rewrite",
        ));
    }
    if !source_bytes.is_ascii() {
        return Err(Error::UnsupportedSource(
            "paragraph-layout editing refuses non-ASCII transport encodings",
        ));
    }
    if !source.opaque().is_empty() {
        return Err(Error::UnsupportedSource(
            "paragraph-layout editing refuses unknown RTF syntax",
        ));
    }
    source
        .model()
        .local_paragraph_property_editability()
        .map_err(Error::UnsupportedSource)?;
    for paragraph in source.body().paragraphs() {
        if paragraph
            .inlines()
            .any(|inline| matches!(inline, crate::text::Inline::Break(crate::text::Break::Line)))
        {
            return Err(Error::UnsupportedSource(
                "paragraph-layout editing refuses inline line breaks",
            ));
        }
        let raw = paragraph.format().raw();
        if raw.paragraph_style.is_some()
            || raw.list_override.is_some()
            || raw.list_level.is_some()
            || raw.legacy_numbering.is_some()
        {
            return Err(Error::UnsupportedSource(
                "paragraph-layout editing refuses dependent style or list references",
            ));
        }
    }
    Ok(())
}

impl ParagraphTextReplacement {
    /// Creates one source-relative paragraph replacement.
    #[must_use]
    pub fn new(position: usize, replacement: impl Into<String>) -> Self {
        Self {
            position,
            replacement: replacement.into(),
        }
    }

    /// Zero-based source paragraph position.
    #[must_use]
    pub const fn position(&self) -> usize {
        self.position
    }

    /// Replacement paragraph text.
    #[must_use]
    pub fn replacement(&self) -> &str {
        &self.replacement
    }
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
    /// A paragraph batch must contain at least one replacement.
    EmptyParagraphBatch,
    /// Paragraph batch selectors must be strictly increasing and unique.
    ParagraphBatchOutOfOrder { previous: usize, incoming: usize },
    /// A paragraph-layout batch must contain at least one update.
    EmptyParagraphLayoutBatch,
    /// A paragraph-layout delta must name at least one field.
    EmptyParagraphLayoutPatch { position: usize },
    /// Paragraph-layout batch selectors must be strictly increasing and unique.
    ParagraphLayoutBatchOutOfOrder { previous: usize, incoming: usize },
    /// Two staged operations have overlapping effects.
    Conflict { existing: usize, incoming: usize },
    /// Paragraph-structure and property changes cannot be proven independent.
    StructuralPropertyConflict,
    /// Local paragraph-layout updates cannot share a transaction with body text or character work.
    ParagraphLayoutTextConflict,
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
    /// A paragraph split offset is outside the selected paragraph.
    ParagraphSplitOffsetOutOfRange {
        position: usize,
        offset: usize,
        length: usize,
    },
    /// A final paragraph split at its end has no exact terminal boundary.
    ParagraphSplitAtEndRequiresBoundary { position: usize },
    /// A paragraph merge must name two consecutive source paragraphs.
    ParagraphMergeNonAdjacent { first: usize, second: usize },
    /// A checked retained-destination selector does not exist.
    DestinationOutOfRange(&'static str),
    /// A checked standalone picture selector does not exist.
    PictureOutOfRange { position: usize, count: usize },
    /// Replacement picture bytes must keep the exact payload length.
    PicturePayloadSizeMismatch {
        position: usize,
        expected: usize,
        observed: usize,
    },
    /// A picture batch must contain at least one replacement.
    EmptyPicturePayloadBatch,
    /// Picture batch selectors must be strictly increasing and unique.
    PicturePayloadBatchOutOfOrder { previous: usize, incoming: usize },
    /// A picture-removal batch must contain at least one selector.
    EmptyPictureRemovalBatch,
    /// Picture-removal selectors must be strictly increasing and unique.
    PictureRemovalBatchOutOfOrder { previous: usize, incoming: usize },
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
    /// The source declares active document protection, so changed publication is refused.
    ProtectedDocument {
        protection_type: crate::ProtectionType,
    },
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
            Self::EmptyParagraphBatch => {
                formatter.write_str("RTF paragraph replacement batch must not be empty")
            },
            Self::ParagraphBatchOutOfOrder { previous, incoming } => write!(
                formatter,
                "RTF paragraph replacement positions must be strictly increasing: {incoming} follows {previous}"
            ),
            Self::EmptyParagraphLayoutBatch => {
                formatter.write_str("RTF paragraph-layout update batch must not be empty")
            },
            Self::EmptyParagraphLayoutPatch { position } => write!(
                formatter,
                "RTF paragraph-layout update at position {position} has no fields"
            ),
            Self::ParagraphLayoutBatchOutOfOrder { previous, incoming } => write!(
                formatter,
                "RTF paragraph-layout positions must be strictly increasing: {incoming} follows {previous}"
            ),
            Self::Conflict { existing, incoming } => write!(
                formatter,
                "RTF edit operation {incoming} conflicts with operation {existing}"
            ),
            Self::StructuralPropertyConflict => formatter
                .write_str("RTF paragraph-structure and property changes cannot compose safely"),
            Self::ParagraphLayoutTextConflict => formatter.write_str(
                "RTF paragraph-layout updates cannot compose with body text or character changes",
            ),
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
            Self::ParagraphSplitOffsetOutOfRange {
                position,
                offset,
                length,
            } => write!(
                formatter,
                "RTF paragraph {position} split offset {offset} is outside 0..{length}"
            ),
            Self::ParagraphSplitAtEndRequiresBoundary { position } => write!(
                formatter,
                "RTF paragraph {position} split at its end requires an exact terminal boundary"
            ),
            Self::ParagraphMergeNonAdjacent { first, second } => write!(
                formatter,
                "RTF paragraphs {first} and {second} are not adjacent source paragraphs"
            ),
            Self::DestinationOutOfRange(destination) => {
                write!(
                    formatter,
                    "RTF retained destination does not exist: {destination}"
                )
            },
            Self::PictureOutOfRange { position, count } => write!(
                formatter,
                "RTF standalone picture position {position} is outside 0..{count}"
            ),
            Self::PicturePayloadSizeMismatch {
                position,
                expected,
                observed,
            } => write!(
                formatter,
                "RTF picture {position} payload length must remain {expected} bytes, observed {observed}"
            ),
            Self::EmptyPicturePayloadBatch => {
                formatter.write_str("RTF picture payload replacement batch must not be empty")
            },
            Self::PicturePayloadBatchOutOfOrder { previous, incoming } => write!(
                formatter,
                "RTF picture positions must be strictly increasing: {incoming} follows {previous}"
            ),
            Self::EmptyPictureRemovalBatch => {
                formatter.write_str("RTF picture removal batch must not be empty")
            },
            Self::PictureRemovalBatchOutOfOrder { previous, incoming } => write!(
                formatter,
                "RTF picture removal positions must be strictly increasing: {incoming} follows {previous}"
            ),
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
            Self::ProtectedDocument { protection_type } => write!(
                formatter,
                "RTF document protection ({}) refuses changed publication",
                protection_type_name(*protection_type)
            ),
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
        raw_structure: Option<RawParagraphOperation>,
    },
    Alignment {
        position: usize,
        before: Alignment,
        after: Alignment,
    },
    ParagraphLayout {
        position: usize,
        fields: LayoutFields,
        before: ParagraphLayout,
        after: ParagraphLayout,
    },
    Bold {
        span: TextSpan,
        before: bool,
        after: bool,
    },
    Italic {
        span: TextSpan,
        before: bool,
        after: bool,
    },
    Underline {
        span: TextSpan,
        before: UnderlineStyle,
        after: UnderlineStyle,
    },
    Strike {
        span: TextSpan,
        before: bool,
        after: bool,
    },
    InsertParagraph {
        position: usize,
        span: TextSpan,
        text: String,
    },
    RemoveParagraph {
        position: usize,
        text: String,
    },
    RestoreParagraph {
        position: usize,
        text: String,
    },
    MoveParagraph {
        position: usize,
        final_position: usize,
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
    PicturePayload(picture_payload::StagedPicturePayload),
    PictureRemoval(picture_payload::StagedPictureRemoval),
    RootTransfer {
        vocabulary: &'static str,
        effect: String,
        before: Vec<u8>,
        after: Vec<u8>,
    },
}

/// A source-proven ordinary-paragraph boundary edit.
///
/// The semantic `Text` operation carrying this marker is deliberately kept
/// separate from the normal writer-backed text path. Its source ranges are
/// proven against the immutable snapshot before staging and checked again
/// before publication, allowing the commit to splice only `\\par` bytes.
#[derive(Debug, Clone)]
enum RawParagraphOperation {
    Split {
        position: usize,
        offset: usize,
        source_offset: usize,
        boundary: Vec<u8>,
        before: String,
    },
    Merge {
        position: usize,
        boundary: Range<usize>,
        boundary_bytes: Vec<u8>,
        left: String,
        right: String,
    },
}

impl Operation {
    fn replacement_bytes(&self) -> usize {
        match self {
            Self::Text {
                raw_structure: Some(RawParagraphOperation::Split { boundary, .. }),
                ..
            } => boundary.len(),
            Self::Text { after, .. }
            | Self::TableCellText { after, .. }
            | Self::HeaderFooterText { after, .. }
            | Self::AnnotationText { after, .. }
            | Self::NoteText { after, .. }
            | Self::ShapeText { after, .. } => after.len(),
            Self::InsertParagraph { text, .. } => text.len().saturating_add(1),
            Self::RestoreParagraph { text, .. } => text.len().saturating_add(1),
            Self::RemoveParagraph { .. } | Self::MoveParagraph { .. } => 0,
            Self::PicturePayload(operation) => operation.after.len(),
            Self::PictureRemoval(_) => 0,
            Self::RootTransfer { after, .. } => after.len(),
            Self::ParagraphLayout { before, after, .. } => {
                let _ = (before, after);
                0
            },
            Self::Alignment { .. }
            | Self::Bold { .. }
            | Self::Italic { .. }
            | Self::Underline { .. }
            | Self::Strike { .. } => 0,
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
            Self::ParagraphLayout {
                position, fields, ..
            } => layout_effect_keys(*position, *fields),
            Self::Bold { span, .. } => {
                vec![format!("body:character:{}-{}:bold", span.start, span.end)]
            },
            Self::Italic { span, .. } => {
                vec![format!("body:character:{}-{}:italic", span.start, span.end)]
            },
            Self::Underline { span, .. } => {
                vec![format!(
                    "body:character:{}-{}:underline",
                    span.start, span.end
                )]
            },
            Self::Strike { span, .. } => {
                vec![format!("body:character:{}-{}:strike", span.start, span.end)]
            },
            Self::InsertParagraph { .. }
            | Self::RemoveParagraph { .. }
            | Self::RestoreParagraph { .. }
            | Self::MoveParagraph { .. } => vec!["body:structure".to_string()],
            Self::TableCellText { path, .. } => vec![table_cell_effect(path)],
            Self::HeaderFooterText { target, .. } => vec![header_footer_effect(*target)],
            Self::AnnotationText { index, .. } => vec![annotation_effect(*index)],
            Self::NoteText { index, .. } => vec![note_effect(*index)],
            Self::ShapeText { index, .. } => vec![shape_effect(*index)],
            Self::PicturePayload(operation) => {
                vec![format!("body:picture:{}:payload", operation.position)]
            },
            Self::PictureRemoval(operation) => {
                vec![format!("body:picture:{}", operation.position)]
            },
            Self::RootTransfer { effect, .. } => vec![effect.clone()],
        }
    }

    const fn span(&self) -> Option<TextSpan> {
        match self {
            Self::Text { span, .. }
            | Self::Bold { span, .. }
            | Self::Italic { span, .. }
            | Self::Underline { span, .. }
            | Self::Strike { span, .. }
            | Self::InsertParagraph { span, .. } => Some(*span),
            Self::Alignment { .. }
            | Self::ParagraphLayout { .. }
            | Self::RemoveParagraph { .. }
            | Self::RestoreParagraph { .. }
            | Self::MoveParagraph { .. }
            | Self::TableCellText { .. }
            | Self::HeaderFooterText { .. }
            | Self::AnnotationText { .. }
            | Self::NoteText { .. }
            | Self::ShapeText { .. }
            | Self::PicturePayload(_)
            | Self::PictureRemoval(_)
            | Self::RootTransfer { .. } => None,
        }
    }

    const fn is_property(&self) -> bool {
        matches!(
            self,
            Self::Alignment { .. }
                | Self::ParagraphLayout { .. }
                | Self::Bold { .. }
                | Self::Italic { .. }
                | Self::Underline { .. }
                | Self::Strike { .. }
        )
    }

    const fn is_destination(&self) -> bool {
        matches!(
            self,
            Self::TableCellText { .. }
                | Self::HeaderFooterText { .. }
                | Self::AnnotationText { .. }
                | Self::NoteText { .. }
                | Self::ShapeText { .. }
                | Self::PicturePayload(_)
                | Self::PictureRemoval(_)
                | Self::RootTransfer { .. }
        )
    }

    const fn is_root_transfer(&self) -> bool {
        matches!(self, Self::RootTransfer { .. })
    }

    const fn is_picture_payload(&self) -> bool {
        matches!(self, Self::PicturePayload(_))
    }

    const fn is_picture_removal(&self) -> bool {
        matches!(self, Self::PictureRemoval(_))
    }

    const fn is_lifecycle(&self) -> bool {
        matches!(
            self,
            Self::Text {
                raw_structure: Some(_),
                ..
            } | Self::RemoveParagraph { .. }
                | Self::RestoreParagraph { .. }
                | Self::MoveParagraph { .. }
        )
    }

    const fn is_raw_paragraph_structure(&self) -> bool {
        matches!(
            self,
            Self::Text {
                raw_structure: Some(_),
                ..
            }
        )
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

    /// Atomically stages a bounded, source-ordered batch of paragraph replacements.
    ///
    /// The batch must be non-empty and its zero-based paragraph positions must
    /// be strictly increasing. Every selector and conflict is preflighted with
    /// one forward paragraph traversal before any operation is appended, so a
    /// late failure leaves this edit unchanged. Newlines retain the scalar
    /// replacement semantics and therefore form structural operations.
    ///
    /// # Errors
    /// Returns an error for an empty or unordered batch, an invalid selector,
    /// a conflict, a structural/property conflict, or retained limits.
    pub fn replace_body_paragraph_texts(
        &mut self,
        replacements: &[ParagraphTextReplacement],
    ) -> Result<&mut Self, Error> {
        self.ensure_body_compatible()?;
        if self
            .operations
            .iter()
            .any(|operation| matches!(operation, Operation::ParagraphLayout { .. }))
        {
            return Err(Error::ParagraphLayoutTextConflict);
        }
        if replacements.is_empty() {
            return Err(Error::EmptyParagraphBatch);
        }
        for (previous, incoming) in replacements.iter().zip(replacements.iter().skip(1)) {
            let previous = previous.position;
            let incoming = incoming.position;
            if incoming <= previous {
                return Err(Error::ParagraphBatchOutOfOrder { previous, incoming });
            }
        }

        self.ensure_operation_room_for(replacements.len())?;
        let replacement_bytes = replacements.iter().fold(0usize, |total, replacement| {
            total.saturating_add(replacement.replacement.len())
        });
        let observed_replacement_bytes = self.replacement_bytes.saturating_add(replacement_bytes);
        let replacement_limit = self.source.limits().max_source_bytes();
        if observed_replacement_bytes > replacement_limit {
            return Err(Error::InputTooLarge {
                observed: observed_replacement_bytes,
                limit: replacement_limit,
            });
        }

        let has_property_operation = self.operations.iter().any(Operation::is_property);
        let mut staged = Vec::with_capacity(replacements.len());
        let mut replacements = replacements.iter().peekable();
        let mut paragraph_start = 0usize;
        for (position, paragraph) in self.source.body().paragraphs().enumerate() {
            let paragraph_end =
                paragraph_start
                    .checked_add(paragraph.len())
                    .ok_or(Error::InputTooLarge {
                        observed: usize::MAX,
                        limit: replacement_limit,
                    })?;
            if replacements
                .peek()
                .is_some_and(|replacement| replacement.position == position)
            {
                let replacement = replacements.next().ok_or(Error::UnsupportedSource(
                    "paragraph replacement cursor became inconsistent",
                ))?;
                let span = TextSpan::new(paragraph_start, paragraph_end)?;
                let before = self
                    .source
                    .text()
                    .get(span.start..span.end)
                    .ok_or(Error::SpanNotOnCharacterBoundary {
                        position: span.start,
                    })?
                    .to_string();
                let structural = before.contains('\n') || replacement.replacement.contains('\n');
                if structural && has_property_operation {
                    return Err(Error::StructuralPropertyConflict);
                }
                let incoming = self.operations.len().saturating_add(staged.len());
                for (existing, operation) in self.operations.iter().enumerate() {
                    if operation
                        .span()
                        .is_some_and(|existing_span| spans_conflict(existing_span, span))
                    {
                        return Err(Error::Conflict { existing, incoming });
                    }
                }
                staged.push(Operation::Text {
                    span,
                    before,
                    after: replacement.replacement.clone(),
                    structural,
                    raw_structure: None,
                });
            }
            paragraph_start = paragraph_end.checked_add(1).ok_or(Error::InputTooLarge {
                observed: usize::MAX,
                limit: replacement_limit,
            })?;
        }
        if let Some(replacement) = replacements.next() {
            return Err(Error::ParagraphOutOfRange {
                position: replacement.position,
                count: self.source.paragraph_count(),
            });
        }

        self.replacement_bytes = observed_replacement_bytes;
        self.operations.append(&mut staged);
        Ok(self)
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
        if self
            .operations
            .iter()
            .any(|operation| matches!(operation, Operation::ParagraphLayout { .. }))
        {
            return Err(Error::ParagraphLayoutTextConflict);
        }
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
        if structural && self.operations.iter().any(Operation::is_property) {
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
            raw_structure: None,
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
        if self
            .operations
            .iter()
            .any(|operation| matches!(operation, Operation::ParagraphLayout { .. }))
        {
            return Err(Error::ParagraphLayoutTextConflict);
        }
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
                    | Operation::RemoveParagraph { .. }
                    | Operation::RestoreParagraph { .. }
                    | Operation::MoveParagraph { .. }
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

    /// Stages italic state for one non-empty UTF-8 body span.
    ///
    /// The selected source range must have one effective italic state and may
    /// not consume a paragraph boundary. Unknown or mixed character ranges
    /// are refused rather than normalized.
    ///
    /// # Errors
    /// Returns an error for invalid geometry, conflicts, structure changes,
    /// mixed formatting, or finite bounds.
    pub fn set_text_italic(&mut self, span: TextSpan, italic: bool) -> Result<&mut Self, Error> {
        self.ensure_body_compatible()?;
        self.ensure_operation_room()?;
        if self
            .operations
            .iter()
            .any(|operation| matches!(operation, Operation::ParagraphLayout { .. }))
        {
            return Err(Error::ParagraphLayoutTextConflict);
        }
        let body = self.source.text();
        validate_span(body, span)?;
        if span.is_empty()
            || body
                .get(span.start..span.end)
                .is_some_and(|text| text.contains('\n'))
        {
            return Err(Error::UnsupportedSource(
                "italic edits require non-empty text within one paragraph",
            ));
        }
        if self.operations.iter().any(|operation| {
            matches!(
                operation,
                Operation::Text {
                    structural: true,
                    ..
                } | Operation::InsertParagraph { .. }
                    | Operation::RemoveParagraph { .. }
                    | Operation::RestoreParagraph { .. }
                    | Operation::MoveParagraph { .. }
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
        let before = italic_for_span(&self.source, span)?;
        self.operations.push(Operation::Italic {
            span,
            before,
            after: italic,
        });
        Ok(self)
    }

    /// Stages one exact underline style for a non-empty UTF-8 body span.
    ///
    /// The selected source range must have one effective underline style and
    /// may not consume a paragraph boundary. Unknown or mixed character
    /// ranges are refused rather than normalized.
    ///
    /// # Errors
    /// Returns an error for invalid geometry, conflicts, structure changes,
    /// mixed formatting, or finite bounds.
    pub fn set_text_underline(
        &mut self,
        span: TextSpan,
        underline: UnderlineStyle,
    ) -> Result<&mut Self, Error> {
        self.ensure_body_compatible()?;
        self.ensure_operation_room()?;
        if self
            .operations
            .iter()
            .any(|operation| matches!(operation, Operation::ParagraphLayout { .. }))
        {
            return Err(Error::ParagraphLayoutTextConflict);
        }
        let body = self.source.text();
        validate_span(body, span)?;
        if span.is_empty()
            || body
                .get(span.start..span.end)
                .is_some_and(|text| text.contains('\n'))
        {
            return Err(Error::UnsupportedSource(
                "underline edits require non-empty text within one paragraph",
            ));
        }
        if self.operations.iter().any(|operation| {
            matches!(
                operation,
                Operation::Text {
                    structural: true,
                    ..
                } | Operation::InsertParagraph { .. }
                    | Operation::RemoveParagraph { .. }
                    | Operation::RestoreParagraph { .. }
                    | Operation::MoveParagraph { .. }
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
        let before = underline_for_span(&self.source, span)?;
        self.operations.push(Operation::Underline {
            span,
            before,
            after: underline,
        });
        Ok(self)
    }

    /// Stages one exact single-strike state for a non-empty UTF-8 body span.
    ///
    /// The selected source range must have one raw single-strike state. Double
    /// strikethrough is retained as an independent formatting facet and is
    /// never normalized by this operation.
    ///
    /// # Errors
    /// Returns an error for invalid geometry, conflicts, structure changes,
    /// mixed formatting, or finite bounds.
    pub fn set_text_strike(&mut self, span: TextSpan, strike: bool) -> Result<&mut Self, Error> {
        self.ensure_body_compatible()?;
        self.ensure_operation_room()?;
        if self
            .operations
            .iter()
            .any(|operation| matches!(operation, Operation::ParagraphLayout { .. }))
        {
            return Err(Error::ParagraphLayoutTextConflict);
        }
        let body = self.source.text();
        validate_span(body, span)?;
        if span.is_empty()
            || body
                .get(span.start..span.end)
                .is_some_and(|text| text.contains('\n'))
        {
            return Err(Error::UnsupportedSource(
                "strike edits require non-empty text within one paragraph",
            ));
        }
        if self.operations.iter().any(|operation| {
            matches!(
                operation,
                Operation::Text {
                    structural: true,
                    ..
                } | Operation::InsertParagraph { .. }
                    | Operation::RemoveParagraph { .. }
                    | Operation::RestoreParagraph { .. }
                    | Operation::MoveParagraph { .. }
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
        let before = strike_for_span(&self.source, span)?;
        self.operations.push(Operation::Strike {
            span,
            before,
            after: strike,
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

    /// Removes one ordinary body paragraph selected against the immutable base.
    ///
    /// The selector is the paragraph's zero-based position before this edit.
    /// Removing the sole paragraph produces an empty visible body story. This
    /// lifecycle operation is intentionally exclusive: it cannot compose with
    /// another body, property, destination, or structural operation in the same
    /// edit, so staging order never becomes last-writer-wins behavior.
    ///
    /// # Errors
    /// Returns an error for an invalid selector, conflicting staged work, or
    /// the finite operation bound. Changed publication additionally refuses
    /// non-plain, ambiguous, compressed, opaque, or protected sources.
    pub fn remove_paragraph(&mut self, position: usize) -> Result<&mut Self, Error> {
        let text = paragraph_text_at(&self.source, position)?.to_string();
        self.ensure_lifecycle_room()?;
        self.operations
            .push(Operation::RemoveParagraph { position, text });
        Ok(self)
    }

    /// Moves one ordinary body paragraph to a final zero-based position.
    ///
    /// Both positions resolve against the immutable base paragraph list.
    /// `final_position` is the paragraph's ordinal in the completed list, after
    /// removal and reinsertion. Equal positions are an exact no-op that shares
    /// the source snapshot. Like removal, a move is an exclusive structural
    /// lifecycle operation and cannot compose with other staged work.
    ///
    /// # Errors
    /// Returns an error when either position is outside the immutable base,
    /// conflicting work is already staged, or finite limits are exceeded.
    /// Changed publication additionally refuses non-plain, ambiguous,
    /// compressed, opaque, or protected sources.
    pub fn move_paragraph(
        &mut self,
        position: usize,
        final_position: usize,
    ) -> Result<&mut Self, Error> {
        let count = self.source.paragraph_count();
        let text = paragraph_text_at(&self.source, position)?.to_string();
        if final_position >= count {
            return Err(Error::ParagraphOutOfRange {
                position: final_position,
                count,
            });
        }
        self.ensure_lifecycle_room()?;
        self.operations.push(Operation::MoveParagraph {
            position,
            final_position,
            text,
        });
        Ok(self)
    }

    /// Splits one ordinary body paragraph at a checked UTF-8 byte offset.
    ///
    /// The offset is measured in the selected paragraph's semantic UTF-8
    /// text, and must lie on a Unicode scalar boundary. This deliberately
    /// narrow transaction accepts only an ASCII, ungrouped, ordinary source
    /// with uniform formatting; publication inserts exactly one `\\par`
    /// control sequence at the proven source boundary and preserves every
    /// unrelated source byte.
    ///
    /// # Errors
    /// Returns an error for an invalid selector or offset, unsupported body
    /// syntax, incompatible formatting, a protected source, or finite limits.
    pub fn split_paragraph(&mut self, position: usize, offset: usize) -> Result<&mut Self, Error> {
        self.split_paragraph_with_boundary(position, offset, PARAGRAPH_SPLIT_BYTES)
    }

    fn split_paragraph_with_boundary(
        &mut self,
        position: usize,
        offset: usize,
        boundary: &[u8],
    ) -> Result<&mut Self, Error> {
        self.ensure_lifecycle_room()?;
        let semantic_range = paragraph_range(&self.source, position)?;
        let semantic_text =
            self.source
                .text()
                .get(semantic_range.clone())
                .ok_or(Error::UnsupportedSource(
                    "ordinary paragraph selector did not resolve to UTF-8 text",
                ))?;
        if offset > semantic_text.len() {
            return Err(Error::ParagraphSplitOffsetOutOfRange {
                position,
                offset,
                length: semantic_text.len(),
            });
        }
        if !semantic_text.is_char_boundary(offset) {
            return Err(Error::SpanNotOnCharacterBoundary {
                position: semantic_range.start.saturating_add(offset),
            });
        }
        let map = ordinary_paragraph_source_map(&self.source)?;
        let paragraph = map
            .paragraphs
            .get(position)
            .ok_or(Error::ParagraphOutOfRange {
                position,
                count: map.paragraphs.len(),
            })?;
        let text =
            self.source
                .text()
                .get(paragraph.text.clone())
                .ok_or(Error::UnsupportedSource(
                    "ordinary paragraph source map has an invalid semantic range",
                ))?;
        if paragraph.text != semantic_range || text != semantic_text {
            return Err(Error::StalePrecondition(
                "ordinary paragraph source map does not match the semantic selector",
            ));
        }
        if offset == semantic_text.len() && paragraph.boundary_after.is_none() {
            return Err(Error::ParagraphSplitAtEndRequiresBoundary { position });
        }
        validate_paragraph_boundary_bytes(boundary)?;
        let source_offset =
            paragraph
                .source
                .start
                .checked_add(offset)
                .ok_or(Error::InputTooLarge {
                    observed: usize::MAX,
                    limit: self.source.limits().max_source_bytes(),
                })?;
        if source_offset > paragraph.source.end {
            return Err(Error::UnsupportedSource(
                "ordinary paragraph split escaped its proven source range",
            ));
        }
        let before = clone_bounded_text(
            text,
            self.source.limits(),
            "could not reserve ordinary paragraph text",
        )?;
        let after = clone_bounded_text(
            "\n",
            self.source.limits(),
            "could not reserve ordinary paragraph break",
        )?;
        let boundary = clone_bounded_bytes(
            boundary,
            self.source.limits(),
            "could not reserve ordinary paragraph boundary",
        )?;
        self.operations.try_reserve(1).map_err(|_error| {
            Error::Write("could not reserve ordinary paragraph operation".to_string())
        })?;
        self.charge_replacement(boundary.len())?;
        self.operations.push(Operation::Text {
            span: TextSpan {
                start: paragraph.text.start.saturating_add(offset),
                end: paragraph.text.start.saturating_add(offset),
            },
            before: String::new(),
            after,
            structural: true,
            raw_structure: Some(RawParagraphOperation::Split {
                position,
                offset,
                source_offset,
                boundary,
                before,
            }),
        });
        Ok(self)
    }

    /// Alias for [`Self::split_paragraph`] emphasizing the checked offset.
    pub fn split_paragraph_at(
        &mut self,
        position: usize,
        offset: usize,
    ) -> Result<&mut Self, Error> {
        self.split_paragraph(position, offset)
    }

    /// Merges two adjacent ordinary body paragraphs.
    ///
    /// `first` and `second` are zero-based positions in the immutable source
    /// snapshot. Only the exact source span of their `\\par` boundary is
    /// removed; all text, controls, and bytes outside that boundary remain
    /// untouched.
    ///
    /// # Errors
    /// Returns an error for invalid selectors, non-adjacent positions,
    /// unsupported body syntax, incompatible formatting, a protected source,
    /// or finite limits.
    pub fn merge_paragraphs(&mut self, first: usize, second: usize) -> Result<&mut Self, Error> {
        self.ensure_lifecycle_room()?;
        let map = ordinary_paragraph_source_map(&self.source)?;
        let left = map
            .paragraphs
            .get(first)
            .ok_or(Error::ParagraphOutOfRange {
                position: first,
                count: map.paragraphs.len(),
            })?;
        let right = map
            .paragraphs
            .get(second)
            .ok_or(Error::ParagraphOutOfRange {
                position: second,
                count: map.paragraphs.len(),
            })?;
        let expected_second = first
            .checked_add(1)
            .ok_or(Error::ParagraphMergeNonAdjacent { first, second })?;
        if second != expected_second {
            return Err(Error::ParagraphMergeNonAdjacent { first, second });
        }
        let boundary = left.boundary_after.clone().ok_or(Error::UnsupportedSource(
            "ordinary paragraph has no exact source boundary",
        ))?;
        let source_bytes = self
            .source
            .source_bytes()
            .ok_or(Error::UnsupportedSource("snapshot has no exact RTF source"))?;
        let boundary_bytes = source_bytes
            .get(boundary.clone())
            .ok_or(Error::UnsupportedSource(
                "ordinary paragraph boundary is outside the source",
            ))?;
        let left_text =
            self.source
                .text()
                .get(left.text.clone())
                .ok_or(Error::UnsupportedSource(
                    "ordinary paragraph source map has an invalid left range",
                ))?;
        let right_text =
            self.source
                .text()
                .get(right.text.clone())
                .ok_or(Error::UnsupportedSource(
                    "ordinary paragraph source map has an invalid right range",
                ))?;
        // `ordinary_paragraph_source_map` has already proved uniform body
        // formatting through `plain_body_text_editability`.
        self.operations.try_reserve(1).map_err(|_error| {
            Error::Write("could not reserve ordinary paragraph operation".to_string())
        })?;
        let left_owned = clone_bounded_text(
            left_text,
            self.source.limits(),
            "could not reserve left ordinary paragraph text",
        )?;
        let right_owned = clone_bounded_text(
            right_text,
            self.source.limits(),
            "could not reserve right ordinary paragraph text",
        )?;
        let before = clone_bounded_text(
            "\n",
            self.source.limits(),
            "could not reserve ordinary paragraph boundary",
        )?;
        let boundary_bytes = clone_bounded_bytes(
            boundary_bytes,
            self.source.limits(),
            "could not reserve ordinary paragraph boundary bytes",
        )?;
        self.operations.push(Operation::Text {
            span: TextSpan {
                start: left.text.end,
                end: left.text.end.saturating_add(1),
            },
            before,
            after: String::new(),
            structural: true,
            raw_structure: Some(RawParagraphOperation::Merge {
                position: first,
                boundary,
                boundary_bytes,
                left: left_owned,
                right: right_owned,
            }),
        });
        Ok(self)
    }

    /// Merges one ordinary paragraph with its immediate successor.
    pub fn merge_paragraph_with_next(&mut self, position: usize) -> Result<&mut Self, Error> {
        let second = position.checked_add(1).ok_or(Error::ParagraphOutOfRange {
            position,
            count: self.source.paragraph_count(),
        })?;
        self.merge_paragraphs(position, second)
    }

    fn restore_paragraph(&mut self, position: usize, text: &str) -> Result<&mut Self, Error> {
        if text.contains('\n') {
            return Err(Error::UnsupportedSource(
                "one restored paragraph cannot contain an ambiguous newline",
            ));
        }
        let count = self.source.paragraph_count();
        if position > count {
            return Err(Error::ParagraphOutOfRange { position, count });
        }
        self.ensure_lifecycle_room()?;
        self.charge_replacement(text.len().saturating_add(1))?;
        self.operations.push(Operation::RestoreParagraph {
            position,
            text: text.to_string(),
        });
        Ok(self)
    }

    fn ensure_lifecycle_room(&self) -> Result<(), Error> {
        self.ensure_body_compatible()?;
        self.ensure_operation_room()?;
        if self.operations.iter().any(Operation::is_property) {
            return Err(Error::StructuralPropertyConflict);
        }
        if !self.operations.is_empty() {
            return Err(Error::Conflict {
                existing: 0,
                incoming: self.operations.len(),
            });
        }
        Ok(())
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
                    | Operation::RemoveParagraph { .. }
                    | Operation::RestoreParagraph { .. }
                    | Operation::MoveParagraph { .. }
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

    /// Stages one typed dependency-free local paragraph-layout delta.
    ///
    /// Selection and the complete before-state resolve against the immutable
    /// source snapshot. Zero numeric values and false flags clear the selected
    /// facet. Alignment may compose independently; body text, character, and
    /// structural work is deliberately refused by this lossless seam.
    ///
    /// # Errors
    /// Returns an error for an empty delta, invalid selector, overlapping
    /// facet, unsupported source closure, or retained operation bound.
    pub fn patch_paragraph_layout(
        &mut self,
        position: usize,
        patch: ParagraphLayoutPatch,
    ) -> Result<&mut Self, Error> {
        self.patch_body_paragraph_layouts(&[ParagraphLayoutUpdate::new(position, patch)])
    }

    /// Atomically stages a bounded source-ordered paragraph-layout batch.
    ///
    /// Every selector, dependency, limit, and effect conflict is preflighted
    /// before any operation is appended. Updates must be strictly ordered by
    /// source paragraph position. Different facets, including multiple
    /// independently prepared deltas for different paragraphs, compose.
    ///
    /// # Errors
    /// Returns an error for an empty/unordered batch or delta, unsupported
    /// source, invalid selector, conflict, mixed text/character work, or bound.
    pub fn patch_body_paragraph_layouts(
        &mut self,
        updates: &[ParagraphLayoutUpdate],
    ) -> Result<&mut Self, Error> {
        self.ensure_body_compatible()?;
        if updates.is_empty() {
            return Err(Error::EmptyParagraphLayoutBatch);
        }
        for (previous, incoming) in updates.iter().zip(updates.iter().skip(1)) {
            if incoming.position <= previous.position {
                return Err(Error::ParagraphLayoutBatchOutOfOrder {
                    previous: previous.position,
                    incoming: incoming.position,
                });
            }
        }
        if self.operations.iter().any(|operation| {
            matches!(
                operation,
                Operation::Text { .. }
                    | Operation::Bold { .. }
                    | Operation::Italic { .. }
                    | Operation::Underline { .. }
                    | Operation::Strike { .. }
                    | Operation::InsertParagraph { .. }
            )
        }) {
            return Err(Error::ParagraphLayoutTextConflict);
        }
        self.ensure_operation_room_for(updates.len())?;
        ensure_paragraph_layout_source(&self.source)?;

        let mut staged = Vec::new();
        staged.try_reserve_exact(updates.len()).map_err(|_error| {
            Error::Write("could not reserve paragraph-layout batch".to_string())
        })?;
        let mut updates = updates.iter().peekable();
        for (position, paragraph) in self.source.body().paragraphs().enumerate() {
            if updates
                .peek()
                .is_some_and(|update| update.position == position)
            {
                let update = updates.next().ok_or(Error::UnsupportedSource(
                    "paragraph-layout selector cursor became inconsistent",
                ))?;
                let fields = update.patch.fields();
                if fields.is_empty() {
                    return Err(Error::EmptyParagraphLayoutPatch { position });
                }
                let before = ParagraphLayout::from_raw(paragraph.format().raw());
                let mut after = before;
                update.patch.apply(&mut after);
                let incoming = self.operations.len().saturating_add(staged.len());
                for (existing, operation) in self.operations.iter().enumerate() {
                    if layout_operation_conflicts(operation, position, fields) {
                        return Err(Error::Conflict { existing, incoming });
                    }
                }
                staged.push(Operation::ParagraphLayout {
                    position,
                    fields,
                    before,
                    after,
                });
            }
        }
        if let Some(update) = updates.next() {
            return Err(Error::ParagraphOutOfRange {
                position: update.position,
                count: self.source.paragraph_count(),
            });
        }
        self.operations.append(&mut staged);
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
        if let Some(existing) = self.operations.iter().position(Operation::is_lifecycle) {
            return Err(Error::Conflict {
                existing,
                incoming: self.operations.len(),
            });
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
        self.ensure_operation_room_for(1)
    }

    fn ensure_operation_room_for(&self, additional: usize) -> Result<(), Error> {
        let observed = self.operations.len().saturating_add(additional);
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

        if self
            .operations
            .iter()
            .any(Operation::is_raw_paragraph_structure)
        {
            if operation_count != 1 {
                return Err(Error::BodyDestinationConflict);
            }
            return commit_raw_paragraph_structure(self, operation_count);
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
        let picture_payload_count = self
            .operations
            .iter()
            .filter(|operation| operation.is_picture_payload())
            .count();
        let picture_removal_count = self
            .operations
            .iter()
            .filter(|operation| operation.is_picture_removal())
            .count();
        if picture_removal_count != 0 {
            if picture_removal_count != operation_count {
                return Err(Error::BodyDestinationConflict);
            }
            return picture_payload::commit_removals(self, operation_count);
        }
        if picture_payload_count != 0 {
            if picture_payload_count != operation_count {
                return Err(Error::BodyDestinationConflict);
            }
            return picture_payload::commit(self, operation_count);
        }
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
        let layout_operation = self
            .operations
            .iter()
            .any(|operation| matches!(operation, Operation::ParagraphLayout { .. }));
        let property_operation = self.operations.iter().any(|operation| {
            matches!(
                operation,
                Operation::Alignment { .. }
                    | Operation::ParagraphLayout { .. }
                    | Operation::Bold { .. }
                    | Operation::Italic { .. }
                    | Operation::Underline { .. }
                    | Operation::Strike { .. }
            )
        });
        let mut alignments = if property_operation {
            source_alignments(&self.source)
        } else {
            Vec::new()
        };
        let base_bold = if property_operation && !layout_operation {
            base_bold_for_edit(&self.source, &self.operations)?
        } else {
            false
        };
        let base_italic = if property_operation && !layout_operation {
            base_italic_for_edit(&self.source, &self.operations)?
        } else {
            false
        };
        let base_underline = if property_operation && !layout_operation {
            base_underline_for_edit(&self.source, &self.operations)?
        } else {
            UnderlineStyle::None
        };
        let base_strike = if property_operation && !layout_operation {
            base_strike_for_edit(&self.source, &self.operations)?
        } else {
            false
        };
        let mut baseline = self
            .source
            .body()
            .runs()
            .next()
            .map_or_else(crate::types::Formatting::default, |run| *run.format().raw());
        baseline.bold = base_bold;
        baseline.italic = base_italic;
        baseline.underline = base_underline;
        baseline.strike = base_strike;
        let mut projected_bold_ranges = Vec::new();
        let mut projected_italic_ranges = Vec::new();
        let mut projected_underline_ranges = Vec::new();
        let mut projected_strike_ranges = Vec::new();
        let mut paragraph_properties = if layout_operation {
            source_paragraph_properties(&self.source)?
        } else {
            Vec::new()
        };
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
                    if layout_operation {
                        let count = paragraph_properties.len();
                        let paragraph = paragraph_properties.get_mut(*position).ok_or(
                            Error::ParagraphOutOfRange {
                                position: *position,
                                count,
                            },
                        )?;
                        paragraph.alignment = *after;
                    }
                },
                Operation::ParagraphLayout {
                    position,
                    fields,
                    after,
                    ..
                } => {
                    let count = paragraph_properties.len();
                    let paragraph = paragraph_properties.get_mut(*position).ok_or(
                        Error::ParagraphOutOfRange {
                            position: *position,
                            count,
                        },
                    )?;
                    apply_layout_to_raw(paragraph, *after, *fields);
                },
                Operation::Bold { span, after, .. } => {
                    projected_bold_ranges
                        .push((project_base_span(*span, &self.operations)?, *after));
                },
                Operation::Italic { span, after, .. } => {
                    projected_italic_ranges
                        .push((project_base_span(*span, &self.operations)?, *after));
                },
                Operation::Underline { span, after, .. } => {
                    projected_underline_ranges
                        .push((project_base_span(*span, &self.operations)?, *after));
                },
                Operation::Strike { span, after, .. } => {
                    projected_strike_ranges
                        .push((project_base_span(*span, &self.operations)?, *after));
                },
                Operation::Text { .. }
                | Operation::InsertParagraph { .. }
                | Operation::RemoveParagraph { .. }
                | Operation::RestoreParagraph { .. }
                | Operation::MoveParagraph { .. } => {},
                Operation::TableCellText { .. }
                | Operation::HeaderFooterText { .. }
                | Operation::AnnotationText { .. }
                | Operation::NoteText { .. }
                | Operation::ShapeText { .. }
                | Operation::PicturePayload(_)
                | Operation::PictureRemoval(_) => {
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
        let has_italic_delta = self.operations.iter().any(|operation| {
            matches!(operation, Operation::Italic { before, after, .. } if before != after)
        });
        let has_underline_delta = self.operations.iter().any(|operation| {
            matches!(operation, Operation::Underline { before, after, .. } if before != after)
        });
        let has_strike_delta = self.operations.iter().any(|operation| {
            matches!(operation, Operation::Strike { before, after, .. } if before != after)
        });
        let has_layout_delta = self.operations.iter().any(|operation| {
            matches!(operation, Operation::ParagraphLayout { before, after, .. } if before != after)
        });
        let did_change = replacement != self.source.text()
            || alignments != original_alignments
            || has_layout_delta
            || has_bold_delta
            || has_italic_delta
            || has_underline_delta
            || has_strike_delta;
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
        ensure_changed_publication_allowed(&self.source)?;

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
        let has_italic_operation = self
            .operations
            .iter()
            .any(|operation| matches!(operation, Operation::Italic { .. }));
        let has_underline_operation = self
            .operations
            .iter()
            .any(|operation| matches!(operation, Operation::Underline { .. }));
        let has_strike_operation = self
            .operations
            .iter()
            .any(|operation| matches!(operation, Operation::Strike { .. }));
        if has_italic_operation && !source_bytes.is_ascii() {
            return Err(Error::UnsupportedSource(
                "italic edits refuse non-ASCII transport encodings",
            ));
        }
        if has_underline_operation && !source_bytes.is_ascii() {
            return Err(Error::UnsupportedSource(
                "underline edits refuse non-ASCII transport encodings",
            ));
        }
        if has_strike_operation && !source_bytes.is_ascii() {
            return Err(Error::UnsupportedSource(
                "strike edits refuse non-ASCII transport encodings",
            ));
        }
        if layout_operation {
            self.source
                .model()
                .local_paragraph_property_editability()
                .map_err(Error::UnsupportedSource)?;
        } else if has_bold_operation
            || has_italic_operation
            || has_underline_operation
            || has_strike_operation
        {
            if has_bold_operation
                && !has_italic_operation
                && !has_underline_operation
                && !has_strike_operation
            {
                self.source
                    .model()
                    .plain_body_bold_editability()
                    .map_err(Error::UnsupportedSource)?;
            } else {
                plain_body_character_editability(
                    &self.source,
                    has_bold_operation,
                    has_italic_operation,
                    has_underline_operation,
                    has_strike_operation,
                )
                .map_err(Error::UnsupportedSource)?;
            }
        } else {
            self.source
                .model()
                .plain_body_text_editability()
                .map_err(Error::UnsupportedSource)?;
        }
        let span = retained_or_located_body_source_span(&self.source, source_bytes)?;
        let replacement_bytes = if layout_operation {
            if replacement != self.source.text()
                || has_bold_operation
                || has_italic_operation
                || has_underline_operation
                || has_strike_operation
            {
                return Err(Error::ParagraphLayoutTextConflict);
            }
            ensure_paragraph_layout_source(&self.source)?;
            encoded_body_with_paragraph_properties(
                &self.source,
                &paragraph_properties,
                self.source.limits(),
            )?
        } else if property_operation {
            encoded_body_with_properties(
                &replacement,
                &alignments,
                baseline,
                base_bold,
                &projected_bold_ranges,
                base_italic,
                &projected_italic_ranges,
                base_underline,
                &projected_underline_ranges,
                base_strike,
                &projected_strike_ranges,
                self.source.limits(),
            )?
        } else {
            encoded_body_text(&replacement, self.source.limits())?
        };
        let bytes = splice_body(source_bytes, span, &replacement_bytes, self.source.limits())?;
        let snapshot = Snapshot::from_bytes_with_limits(&bytes, self.source.limits())?;
        validate_opaque_preservation(&self.source, &snapshot, &self.operations)?;
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
        if layout_operation {
            if source_paragraph_properties(&snapshot)? != paragraph_properties {
                return Err(Error::UnsupportedSource(
                    "candidate paragraph properties did not survive RTF validation",
                ));
            }
            if !same_source_runs(&snapshot, &self.source) {
                return Err(Error::UnsupportedSource(
                    "candidate character runs did not survive paragraph-layout publication",
                ));
            }
        }
        for (bold_span, expected) in projected_bold_ranges {
            if bold_for_span(&snapshot, bold_span)? != expected {
                return Err(Error::UnsupportedSource(
                    "candidate bold property did not survive RTF validation",
                ));
            }
        }
        for (italic_span, expected) in projected_italic_ranges {
            if italic_for_span(&snapshot, italic_span)? != expected {
                return Err(Error::UnsupportedSource(
                    "candidate italic property did not survive RTF validation",
                ));
            }
        }
        for (underline_span, expected) in projected_underline_ranges {
            if underline_for_span(&snapshot, underline_span)? != expected {
                return Err(Error::UnsupportedSource(
                    "candidate underline property did not survive RTF validation",
                ));
            }
        }
        for (strike_span, expected) in projected_strike_ranges {
            if strike_for_span(&snapshot, strike_span)? != expected {
                return Err(Error::UnsupportedSource(
                    "candidate strike property did not survive RTF validation",
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

fn commit_raw_paragraph_structure(edit: Edit, operation_count: usize) -> Result<Commit, Error> {
    ensure_changed_publication_allowed(&edit.source)?;
    let source_bytes = edit
        .source
        .source_bytes()
        .ok_or(Error::UnsupportedSource("snapshot has no exact RTF source"))?;
    let map = ordinary_paragraph_source_map(&edit.source)?;
    let operation = edit.operations.first().ok_or(Error::UnsupportedSource(
        "missing ordinary paragraph operation",
    ))?;
    let (snapshot, semantic_delta) =
        match operation {
            Operation::Text {
                raw_structure:
                    Some(RawParagraphOperation::Split {
                        position,
                        offset,
                        source_offset,
                        boundary,
                        before,
                    }),
                ..
            } => {
                let paragraph =
                    map.paragraphs
                        .get(*position)
                        .ok_or(Error::ParagraphOutOfRange {
                            position: *position,
                            count: map.paragraphs.len(),
                        })?;
                let actual = edit.source.text().get(paragraph.text.clone()).ok_or(
                    Error::UnsupportedSource(
                        "ordinary paragraph split selector lost its source text",
                    ),
                )?;
                if actual != before {
                    return Err(Error::StalePrecondition("ordinary paragraph text differs"));
                }
                if *offset > actual.len() || !actual.is_char_boundary(*offset) {
                    return Err(Error::SpanNotOnCharacterBoundary {
                        position: paragraph.text.start.saturating_add(*offset),
                    });
                }
                let expected_source_offset =
                    paragraph
                        .source
                        .start
                        .checked_add(*offset)
                        .ok_or(Error::InputTooLarge {
                            observed: usize::MAX,
                            limit: edit.source.limits().max_source_bytes(),
                        })?;
                if expected_source_offset != *source_offset
                    || source_bytes.get(*source_offset..*source_offset) != Some(&[])
                {
                    return Err(Error::StalePrecondition(
                        "ordinary paragraph split source boundary differs",
                    ));
                }
                let bytes = splice_body(
                    source_bytes,
                    *source_offset..*source_offset,
                    boundary,
                    edit.source.limits(),
                )?;
                let snapshot = Snapshot::from_bytes_with_limits(&bytes, edit.source.limits())?;
                ordinary_paragraph_source_map(&snapshot)?;
                let (left, right) = actual.split_at(*offset);
                if snapshot.paragraph_count() != edit.source.paragraph_count().saturating_add(1)
                    || paragraph_text_at(&snapshot, *position)? != left
                    || paragraph_text_at(&snapshot, position.saturating_add(1))? != right
                {
                    return Err(Error::UnsupportedSource(
                        "split candidate failed ordinary paragraph readback",
                    ));
                }
                let semantic_delta = vec![Change::SplitParagraph {
                    position: *position,
                    offset: *offset,
                    before: before.clone(),
                    boundary: boundary.clone(),
                    left: left.to_string(),
                    right: right.to_string(),
                }];
                (snapshot, semantic_delta)
            },
            Operation::Text {
                raw_structure:
                    Some(RawParagraphOperation::Merge {
                        position,
                        boundary,
                        boundary_bytes,
                        left,
                        right,
                    }),
                ..
            } => {
                let paragraph =
                    map.paragraphs
                        .get(*position)
                        .ok_or(Error::ParagraphOutOfRange {
                            position: *position,
                            count: map.paragraphs.len(),
                        })?;
                let next_position = position.saturating_add(1);
                let next = map
                    .paragraphs
                    .get(next_position)
                    .ok_or(Error::ParagraphOutOfRange {
                        position: next_position,
                        count: map.paragraphs.len(),
                    })?;
                if paragraph.boundary_after.as_ref() != Some(boundary)
                    || boundary.start != paragraph.source.end
                    || boundary.end > next.source.start
                    || source_bytes.get(boundary.clone()) != Some(boundary_bytes.as_slice())
                    || edit.source.text().get(paragraph.text.clone()) != Some(left.as_str())
                    || edit.source.text().get(next.text.clone()) != Some(right.as_str())
                {
                    return Err(Error::StalePrecondition(
                        "ordinary paragraph merge boundary or text differs",
                    ));
                }
                let bytes = splice_body(source_bytes, boundary.clone(), &[], edit.source.limits())?;
                let snapshot = Snapshot::from_bytes_with_limits(&bytes, edit.source.limits())?;
                let merged = merged_paragraph_text(left, right, edit.source.limits())?;
                if snapshot.paragraph_count() != edit.source.paragraph_count().saturating_sub(1)
                    || paragraph_text_at(&snapshot, *position)? != merged
                {
                    return Err(Error::UnsupportedSource(
                        "merge candidate failed ordinary paragraph readback",
                    ));
                }
                let semantic_delta = vec![Change::MergeParagraph {
                    position: *position,
                    boundary: boundary_bytes.clone(),
                    left: left.clone(),
                    right: right.clone(),
                }];
                (snapshot, semantic_delta)
            },
            Operation::Text { .. } => return Err(Error::BodyDestinationConflict),
            _ => return Err(Error::BodyDestinationConflict),
        };
    Ok(Commit::new(
        edit.source,
        snapshot,
        true,
        operation_count,
        semantic_delta,
    ))
}

fn merged_paragraph_text(
    left: &str,
    right: &str,
    limits: crate::ParseLimits,
) -> Result<String, Error> {
    let length = left
        .len()
        .checked_add(right.len())
        .ok_or(Error::InputTooLarge {
            observed: usize::MAX,
            limit: limits.max_source_bytes(),
        })?;
    if length > limits.max_source_bytes() {
        return Err(Error::InputTooLarge {
            observed: length,
            limit: limits.max_source_bytes(),
        });
    }
    let mut merged = String::new();
    merged
        .try_reserve_exact(length)
        .map_err(|_error| Error::Write("could not reserve merged paragraph text".to_string()))?;
    merged.push_str(left);
    merged.push_str(right);
    Ok(merged)
}

fn clone_bounded_text(
    text: &str,
    limits: crate::ParseLimits,
    allocation_context: &'static str,
) -> Result<String, Error> {
    if text.len() > limits.max_source_bytes() {
        return Err(Error::InputTooLarge {
            observed: text.len(),
            limit: limits.max_source_bytes(),
        });
    }
    let mut owned = String::new();
    owned
        .try_reserve_exact(text.len())
        .map_err(|_error| Error::Write(allocation_context.to_string()))?;
    owned.push_str(text);
    Ok(owned)
}

fn clone_bounded_bytes(
    bytes: &[u8],
    limits: crate::ParseLimits,
    allocation_context: &'static str,
) -> Result<Vec<u8>, Error> {
    if bytes.len() > limits.max_source_bytes() {
        return Err(Error::InputTooLarge {
            observed: bytes.len(),
            limit: limits.max_source_bytes(),
        });
    }
    let mut owned = Vec::new();
    owned
        .try_reserve_exact(bytes.len())
        .map_err(|_error| Error::Write(allocation_context.to_string()))?;
    owned.extend_from_slice(bytes);
    Ok(owned)
}

fn validate_paragraph_boundary_bytes(boundary: &[u8]) -> Result<(), Error> {
    let keyword_is_par = match boundary.get(..4) {
        Some([slash, p, a, r]) => {
            *slash == b'\\'
                && p.eq_ignore_ascii_case(&b'p')
                && a.eq_ignore_ascii_case(&b'a')
                && r.eq_ignore_ascii_case(&b'r')
        },
        _ => false,
    };
    let valid_delimiter = matches!(boundary.get(4..), Some([]) | Some([b' ']));
    if !keyword_is_par || !valid_delimiter {
        return Err(Error::UnsupportedSource(
            "ordinary paragraph boundary is not an exact \\par control",
        ));
    }
    Ok(())
}

fn commit_destinations(edit: Edit, operation_count: usize) -> Result<Commit, Error> {
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
    ensure_changed_publication_allowed(&edit.source)?;
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
            | Operation::ParagraphLayout { .. }
            | Operation::Bold { .. }
            | Operation::Italic { .. }
            | Operation::Underline { .. }
            | Operation::Strike { .. }
            | Operation::InsertParagraph { .. }
            | Operation::RemoveParagraph { .. }
            | Operation::RestoreParagraph { .. }
            | Operation::MoveParagraph { .. }
            | Operation::PicturePayload(_)
            | Operation::PictureRemoval(_)
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
            | Operation::ParagraphLayout { .. }
            | Operation::Bold { .. }
            | Operation::Italic { .. }
            | Operation::Underline { .. }
            | Operation::Strike { .. }
            | Operation::InsertParagraph { .. }
            | Operation::RemoveParagraph { .. }
            | Operation::RestoreParagraph { .. }
            | Operation::MoveParagraph { .. }
            | Operation::PicturePayload(_)
            | Operation::PictureRemoval(_)
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
    let semantic_delta = semantic_changes(&edit.operations, &[]);
    if source == after {
        return Ok(Commit::new(
            edit.source.clone(),
            edit.source,
            false,
            operation_count,
            semantic_delta,
        ));
    }
    ensure_changed_publication_allowed(&edit.source)?;
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
    Ok(Commit::new(
        edit.source,
        snapshot,
        true,
        operation_count,
        semantic_delta,
    ))
}

fn ensure_changed_publication_allowed(source: &Snapshot) -> Result<(), Error> {
    let protection = source.model().protection();
    if protection.is_protected() {
        return Err(Error::ProtectedDocument {
            protection_type: protection.protection_type(),
        });
    }
    Ok(())
}

const fn protection_type_name(protection_type: crate::ProtectionType) -> &'static str {
    match protection_type {
        crate::ProtectionType::None => "none",
        crate::ProtectionType::ReadOnly => "read-only",
        crate::ProtectionType::RevisionTracking => "revision tracking",
        crate::ProtectionType::Comments => "comments",
        crate::ProtectionType::Forms => "forms",
        crate::ProtectionType::All => "all changes",
    }
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

#[derive(Debug, Clone)]
struct OrdinaryParagraphSource {
    text: Range<usize>,
    source: Range<usize>,
    boundary_after: Option<Range<usize>>,
}

#[derive(Debug, Clone)]
struct OrdinaryParagraphSourceMap {
    paragraphs: Vec<OrdinaryParagraphSource>,
}

/// Proves the deliberately narrow source closure used by paragraph
/// split/merge and maps semantic paragraph offsets to exact source offsets.
///
/// The source map accepts literal ASCII text and `\\par` controls only. This
/// makes every retained semantic byte correspond to exactly one source byte;
/// encoded characters, visible control symbols, nested groups, and all other
/// controls fail closed rather than being guessed around.
fn ordinary_paragraph_source_map(source: &Snapshot) -> Result<OrdinaryParagraphSourceMap, Error> {
    let source_bytes = source
        .source_bytes()
        .ok_or(Error::UnsupportedSource("snapshot has no exact RTF source"))?;
    if crate::compressed::is_compressed_rtf(source_bytes) {
        return Err(Error::UnsupportedSource(
            "ordinary paragraph structure editing refuses compressed RTF",
        ));
    }
    if !source_bytes.is_ascii() {
        return Err(Error::UnsupportedSource(
            "ordinary paragraph structure editing requires an ASCII RTF transport",
        ));
    }
    if !source.opaque().is_empty() || source.model().unknown_syntax_markers() != 0 {
        return Err(Error::UnsupportedSource(
            "ordinary paragraph structure editing refuses unknown RTF syntax",
        ));
    }
    if !source.model().external_references().is_empty()
        || source.model().mail_merge().is_some()
        || source.model().xsl_transform().is_some()
        || source.model().xsl_transform_usage().is_requested()
    {
        return Err(Error::UnsupportedSource(
            "ordinary paragraph structure editing refuses external or transformation content",
        ));
    }
    if !source.tables().is_empty()
        || !source.fields().is_empty()
        || !source.objects().is_empty()
        || !source.pictures().is_empty()
        || !source.shapes().is_empty()
        || !source.model().form_fields().is_empty()
        || !source.model().bookmarks().bookmarks().is_empty()
        || !source.revisions().is_empty()
        || !source.annotations().is_empty()
        || !source.notes().is_empty()
        || !source.model().math_zones().is_empty()
        || !source.model().custom_xml_tags().is_empty()
        || !source.model().protection_ranges().is_empty()
        || !source.model().editable_regions().is_empty()
        || !source.model().body_story_events().is_empty()
    {
        return Err(Error::UnsupportedSource(
            "ordinary paragraph structure editing refuses dependent destinations or review content",
        ));
    }
    source
        .model()
        .plain_body_text_editability()
        .map_err(Error::UnsupportedSource)?;
    let body_span = if !source.text().starts_with('\n') {
        source
            .model()
            .ordinary_body_source_span()
            .ok_or(Error::UnsupportedSource(
                "ordinary paragraph source span is not losslessly proven",
            ))?
    } else {
        ordinary_paragraph_body_source_span(source_bytes, source.text(), source.limits())?
    };
    if body_span.start >= body_span.end || body_span.end > source_bytes.len() {
        return Err(Error::UnsupportedSource(
            "ordinary paragraph source span is malformed",
        ));
    }
    let body_bytes = source_bytes
        .get(body_span.clone())
        .ok_or(Error::UnsupportedSource(
            "ordinary paragraph source span is outside the source",
        ))?;
    let starts_with_par = body_bytes.get(..4).is_some_and(|prefix| {
        prefix.first() == Some(&b'\\')
            && prefix
                .get(1)
                .is_some_and(|byte| byte.eq_ignore_ascii_case(&b'p'))
            && prefix
                .get(2)
                .is_some_and(|byte| byte.eq_ignore_ascii_case(&b'a'))
            && prefix
                .get(3)
                .is_some_and(|byte| byte.eq_ignore_ascii_case(&b'r'))
    });
    let has_preceding_delimiter = body_span.start == 0
        || source_bytes
            .get(body_span.start.saturating_sub(1))
            .is_some_and(|byte| byte.is_ascii_whitespace());
    if semantic_text_starts_with_empty_paragraph(source.text())
        && starts_with_par
        && !has_preceding_delimiter
    {
        return Err(Error::UnsupportedSource(
            "leading empty paragraph has no safe control-word delimiter",
        ));
    }
    let body_input = std::str::from_utf8(body_bytes)
        .map_err(|_error| Error::UnsupportedSource("ordinary paragraph source is not UTF-8"))?;
    let arena = Bump::new();
    let mut lexer = crate::lexer::Lexer::new_with_limits(body_input, &arena, source.limits());
    let (tokens, spans) = lexer.tokenize_with_spans()?;
    let mut paragraphs = Vec::new();
    paragraphs
        .try_reserve_exact(source.paragraph_count())
        .map_err(|_error| Error::Write("could not reserve paragraph source map".to_string()))?;
    let semantic_text = source.text();
    let mut semantic_position = 0usize;
    let mut paragraph_start = 0usize;
    let mut source_text_start = None;
    let mut previous_boundary_end = body_span.start;

    for (token, relative_span) in tokens.iter().zip(&spans) {
        let absolute_span = relative_span
            .start
            .checked_add(body_span.start)
            .and_then(|start| {
                relative_span
                    .end
                    .checked_add(body_span.start)
                    .map(|end| start..end)
            })
            .ok_or(Error::InputTooLarge {
                observed: usize::MAX,
                limit: source.limits().max_source_bytes(),
            })?;
        if absolute_span.start < previous_boundary_end {
            return Err(Error::UnsupportedSource(
                "ordinary paragraph source tokens overlap",
            ));
        }
        match token {
            crate::lexer::Token::Text(value) => {
                let source_fragment =
                    source_bytes
                        .get(absolute_span.clone())
                        .ok_or(Error::UnsupportedSource(
                            "ordinary paragraph text token is outside the source",
                        ))?;
                if value.is_empty() {
                    if !source_fragment.is_empty() {
                        return Err(Error::UnsupportedSource(
                            "ordinary paragraph source contains an unmodeled control boundary",
                        ));
                    }
                    continue;
                }
                if !value.is_ascii() || source_fragment != value.as_bytes() {
                    return Err(Error::UnsupportedSource(
                        "ordinary paragraph source text is not one-to-one ASCII text",
                    ));
                }
                if source_text_start.is_none() {
                    source_text_start = Some(absolute_span.start);
                }
                semantic_position =
                    semantic_position
                        .checked_add(value.len())
                        .ok_or(Error::InputTooLarge {
                            observed: usize::MAX,
                            limit: source.limits().max_source_bytes(),
                        })?;
            },
            crate::lexer::Token::Control(crate::lexer::ControlWord::Par) => {
                let source_start = source_text_start.unwrap_or(absolute_span.start);
                let source_end = absolute_span.start;
                let text_range = paragraph_start..semantic_position;
                let source_range = source_start..source_end;
                if source_range.len() != text_range.len()
                    || source_bytes.get(source_range.clone())
                        != semantic_text.as_bytes().get(text_range.clone())
                {
                    return Err(Error::UnsupportedSource(
                        "ordinary paragraph boundary is not a literal source boundary",
                    ));
                }
                paragraphs.push(OrdinaryParagraphSource {
                    text: text_range,
                    source: source_range,
                    boundary_after: Some(absolute_span.clone()),
                });
                semantic_position =
                    semantic_position
                        .checked_add(1)
                        .ok_or(Error::InputTooLarge {
                            observed: usize::MAX,
                            limit: source.limits().max_source_bytes(),
                        })?;
                paragraph_start = semantic_position;
                source_text_start = None;
                previous_boundary_end = absolute_span.end;
            },
            crate::lexer::Token::OpenBrace
            | crate::lexer::Token::CloseBrace
            | crate::lexer::Token::Control(_)
            | crate::lexer::Token::Binary(_) => {
                return Err(Error::UnsupportedSource(
                    "ordinary paragraph source contains a group, binary payload, or non-paragraph control",
                ));
            },
        }
    }

    if let Some(source_start) = source_text_start {
        let text_range = paragraph_start..semantic_position;
        let source_range = source_start..body_span.end;
        if source_range.len() != text_range.len()
            || source_bytes.get(source_range.clone())
                != semantic_text.as_bytes().get(text_range.clone())
        {
            return Err(Error::UnsupportedSource(
                "ordinary paragraph final source range is not literal text",
            ));
        }
        paragraphs.push(OrdinaryParagraphSource {
            text: text_range,
            source: source_range,
            boundary_after: None,
        });
    }
    if semantic_position != semantic_text.len()
        || paragraphs.len() != source.paragraph_count()
        || paragraphs.is_empty()
    {
        return Err(Error::UnsupportedSource(
            "ordinary paragraph source map does not cover the complete body story",
        ));
    }
    Ok(OrdinaryParagraphSourceMap { paragraphs })
}

fn semantic_text_starts_with_empty_paragraph(text: &str) -> bool {
    text.starts_with('\n')
}

/// Locates the contiguous root-level ordinary body span, including leading
/// `\\par` controls that represent empty paragraphs. The retained parser span
/// intentionally starts at the first literal text, which is insufficient for
/// replaying a split at offset zero on the first paragraph.
fn ordinary_paragraph_body_source_span(
    source: &[u8],
    semantic_text: &str,
    limits: crate::ParseLimits,
) -> Result<Range<usize>, Error> {
    let lexical = std::str::from_utf8(source)
        .map_err(|_error| Error::UnsupportedSource("ordinary paragraph source is not UTF-8"))?;
    let arena = Bump::new();
    let mut lexer = crate::lexer::Lexer::new_with_limits(lexical, &arena, limits);
    let (tokens, spans) = lexer.tokenize_with_spans()?;
    let mut depth = 0usize;
    let mut start = None;
    let mut end = None;
    for (token, span) in tokens.iter().zip(&spans) {
        match token {
            crate::lexer::Token::OpenBrace => {
                if depth == 1 && start.is_some() {
                    return Err(Error::UnsupportedSource(
                        "the ordinary paragraph body is not one contiguous root-level span",
                    ));
                }
                depth = depth.checked_add(1).ok_or(Error::UnsupportedSource(
                    "RTF group nesting overflowed while locating the ordinary body",
                ))?;
            },
            crate::lexer::Token::CloseBrace => {
                if depth == 1 {
                    end = Some(span.start);
                    break;
                }
                depth = depth.checked_sub(1).ok_or(Error::UnsupportedSource(
                    "RTF group nesting underflowed while locating the ordinary body",
                ))?;
            },
            crate::lexer::Token::Text(_)
            | crate::lexer::Token::Control(crate::lexer::ControlWord::Par)
                if depth == 1 && start.is_none() =>
            {
                start = Some(span.start);
            },
            crate::lexer::Token::Binary(_) if depth == 1 && start.is_some() => {
                return Err(Error::UnsupportedSource(
                    "the ordinary paragraph body contains binary data",
                ));
            },
            crate::lexer::Token::Control(_)
            | crate::lexer::Token::Text(_)
            | crate::lexer::Token::Binary(_) => {},
        }
    }
    let root_end = end.ok_or(Error::UnsupportedSource(
        "RTF root group has no closing ordinary-body boundary",
    ))?;
    match start {
        Some(start_offset) => Ok(start_offset..root_end),
        None if semantic_text.is_empty() => Ok(root_end..root_end),
        None => Err(Error::UnsupportedSource(
            "the body has no literal source span for a lossless paragraph edit",
        )),
    }
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

fn paragraph_text_at(source: &Snapshot, position: usize) -> Result<&str, Error> {
    let range = paragraph_range(source, position)?;
    source.text().get(range).ok_or(Error::UnsupportedSource(
        "paragraph selector did not resolve to one UTF-8 body range",
    ))
}

fn project_text(
    source: &Snapshot,
    operations: &[Operation],
) -> Result<(String, Vec<(usize, TextSpan)>), Error> {
    if let [
        operation @ (Operation::RemoveParagraph { .. }
        | Operation::RestoreParagraph { .. }
        | Operation::MoveParagraph { .. }),
    ] = operations
    {
        return Ok((project_lifecycle_text(source, operation)?, Vec::new()));
    }
    let mut text_operations = operations
        .iter()
        .enumerate()
        .filter_map(|(operation_index, operation)| match operation {
            Operation::Text {
                span,
                after,
                before: _,
                structural: _,
                raw_structure: None,
            } => Some((operation_index, *span, after.as_str(), false)),
            Operation::Text {
                raw_structure: Some(_),
                ..
            } => None,
            Operation::InsertParagraph { span, text, .. } => {
                Some((operation_index, *span, text.as_str(), true))
            },
            Operation::Alignment { .. }
            | Operation::ParagraphLayout { .. }
            | Operation::Bold { .. }
            | Operation::Italic { .. }
            | Operation::Underline { .. }
            | Operation::Strike { .. }
            | Operation::RemoveParagraph { .. }
            | Operation::RestoreParagraph { .. }
            | Operation::MoveParagraph { .. }
            | Operation::TableCellText { .. }
            | Operation::HeaderFooterText { .. }
            | Operation::AnnotationText { .. }
            | Operation::NoteText { .. }
            | Operation::ShapeText { .. }
            | Operation::PicturePayload(_)
            | Operation::PictureRemoval(_)
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

fn project_lifecycle_text(source: &Snapshot, operation: &Operation) -> Result<String, Error> {
    if matches!(
        operation,
        Operation::MoveParagraph {
            position,
            final_position,
            ..
        } if position == final_position
    ) {
        return Ok(source.text().to_string());
    }
    let mut paragraphs = plain_lifecycle_paragraphs(source)?;
    match operation {
        Operation::RemoveParagraph { position, text } => {
            let actual = paragraphs
                .get(*position)
                .ok_or(Error::ParagraphOutOfRange {
                    position: *position,
                    count: paragraphs.len(),
                })?;
            if actual != text {
                return Err(Error::StalePrecondition("paragraph text differs"));
            }
            paragraphs.remove(*position);
        },
        Operation::RestoreParagraph { position, text } => {
            if *position > paragraphs.len() {
                return Err(Error::ParagraphOutOfRange {
                    position: *position,
                    count: paragraphs.len(),
                });
            }
            paragraphs.insert(*position, text.clone());
        },
        Operation::MoveParagraph {
            position,
            final_position,
            text,
        } => {
            let count = paragraphs.len();
            if *position >= count {
                return Err(Error::ParagraphOutOfRange {
                    position: *position,
                    count,
                });
            }
            if *final_position >= count {
                return Err(Error::ParagraphOutOfRange {
                    position: *final_position,
                    count,
                });
            }
            let actual = paragraphs
                .get(*position)
                .ok_or(Error::ParagraphOutOfRange {
                    position: *position,
                    count,
                })?;
            if actual != text {
                return Err(Error::StalePrecondition("moved paragraph text differs"));
            }
            let paragraph = paragraphs.remove(*position);
            paragraphs.insert(*final_position, paragraph);
        },
        Operation::Text { .. }
        | Operation::Alignment { .. }
        | Operation::ParagraphLayout { .. }
        | Operation::Bold { .. }
        | Operation::Italic { .. }
        | Operation::Underline { .. }
        | Operation::Strike { .. }
        | Operation::InsertParagraph { .. }
        | Operation::TableCellText { .. }
        | Operation::HeaderFooterText { .. }
        | Operation::AnnotationText { .. }
        | Operation::NoteText { .. }
        | Operation::ShapeText { .. }
        | Operation::PicturePayload(_)
        | Operation::PictureRemoval(_)
        | Operation::RootTransfer { .. } => {
            return Err(Error::UnsupportedSource(
                "non-lifecycle operation entered lifecycle projection",
            ));
        },
    }
    Ok(paragraphs.join("\n"))
}

fn plain_lifecycle_paragraphs(source: &Snapshot) -> Result<Vec<String>, Error> {
    let source_bytes = source
        .source_bytes()
        .ok_or(Error::UnsupportedSource("snapshot has no exact RTF source"))?;
    if crate::compressed::is_compressed_rtf(source_bytes) {
        return Err(Error::UnsupportedSource(
            "compressed RTF needs a transport-aware rewrite",
        ));
    }
    if !source_bytes.is_ascii() {
        return Err(Error::UnsupportedSource(
            "paragraph lifecycle editing refuses non-ASCII transport encodings",
        ));
    }
    if !source.opaque().is_empty() {
        return Err(Error::UnsupportedSource(
            "paragraph lifecycle editing refuses unknown RTF syntax",
        ));
    }
    source
        .model()
        .plain_body_text_editability()
        .map_err(Error::UnsupportedSource)?;
    if source.model().retained_blocks().iter().any(|block| {
        block.formatting != crate::types::Formatting::default()
            || block.paragraph != crate::types::Paragraph::default()
    }) {
        return Err(Error::UnsupportedSource(
            "paragraph lifecycle editing requires default body formatting",
        ));
    }
    let paragraphs = source
        .body()
        .paragraphs()
        .map(|paragraph| paragraph.to_text())
        .collect::<Vec<_>>();
    if paragraphs.iter().any(|paragraph| paragraph.contains('\n'))
        || paragraphs.join("\n") != source.text()
    {
        return Err(Error::UnsupportedSource(
            "paragraph lifecycle editing requires an unambiguous ordinary body story",
        ));
    }
    Ok(paragraphs)
}

fn source_alignments(source: &Snapshot) -> Vec<Alignment> {
    source
        .body()
        .paragraphs()
        .map(|paragraph| paragraph.format().alignment())
        .collect()
}

fn source_paragraph_properties(source: &Snapshot) -> Result<Vec<crate::types::Paragraph>, Error> {
    let mut properties = Vec::new();
    properties
        .try_reserve_exact(source.paragraph_count())
        .map_err(|_error| Error::Write("could not reserve paragraph properties".to_string()))?;
    properties.extend(
        source
            .body()
            .paragraphs()
            .map(|paragraph| *paragraph.format().raw()),
    );
    Ok(properties)
}

fn same_source_runs(left: &Snapshot, right: &Snapshot) -> bool {
    let mut left = left.body().runs();
    let mut right = right.body().runs();
    loop {
        match (left.next(), right.next()) {
            (Some(left), Some(right))
                if left.text() == right.text() && left.format().raw() == right.format().raw() => {},
            (None, None) => return true,
            _ => return false,
        }
    }
}

fn apply_layout_to_raw(
    paragraph: &mut crate::types::Paragraph,
    layout: ParagraphLayout,
    fields: LayoutFields,
) {
    if fields.0 & LayoutFields::SPACE_BEFORE != 0 {
        paragraph.spacing.before = layout.space_before;
    }
    if fields.0 & LayoutFields::SPACE_AFTER != 0 {
        paragraph.spacing.after = layout.space_after;
    }
    if fields.0 & LayoutFields::LEFT_INDENT != 0 {
        paragraph.indentation.left = layout.left_indent;
    }
    if fields.0 & LayoutFields::RIGHT_INDENT != 0 {
        paragraph.indentation.right = layout.right_indent;
    }
    if fields.0 & LayoutFields::FIRST_LINE_INDENT != 0 {
        paragraph.indentation.first_line = layout.first_line_indent;
    }
    if fields.0 & LayoutFields::KEEP_TOGETHER != 0 {
        paragraph.keep_together = layout.keep_together;
    }
    if fields.0 & LayoutFields::KEEP_WITH_NEXT != 0 {
        paragraph.keep_next = layout.keep_with_next;
    }
    if fields.0 & LayoutFields::PAGE_BREAK_BEFORE != 0 {
        paragraph.page_break_before = layout.page_break_before;
    }
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

fn plain_body_character_editability(
    source: &Snapshot,
    allow_mixed_bold: bool,
    allow_mixed_italic: bool,
    allow_mixed_underline: bool,
    allow_mixed_strike: bool,
) -> Result<(), &'static str> {
    source.model().local_paragraph_property_editability()?;
    let mut paragraph_format = None;
    let mut character_format = None;
    for paragraph in source.body().paragraphs() {
        let raw_paragraph = *paragraph.format().raw();
        if paragraph_format.is_some_and(|existing| existing != raw_paragraph) {
            return Err("the body has mixed run or paragraph formatting");
        }
        paragraph_format = Some(raw_paragraph);
        for run in paragraph.runs() {
            let mut raw_character = *run.format().raw();
            if allow_mixed_bold {
                raw_character.bold = false;
            }
            if allow_mixed_italic {
                raw_character.italic = false;
            }
            if allow_mixed_underline {
                raw_character.underline = UnderlineStyle::None;
            }
            if allow_mixed_strike {
                raw_character.strike = false;
            }
            if character_format.is_some_and(|existing| existing != raw_character) {
                return Err("the body has mixed run or paragraph formatting");
            }
            character_format = Some(raw_character);
        }
    }
    Ok(())
}

fn span_fully_covered(span: TextSpan, selected: &[TextSpan]) -> bool {
    let covered = selected.iter().fold(0usize, |total, selected| {
        let start = span.start.max(selected.start);
        let end = span.end.min(selected.end);
        total.saturating_add(end.saturating_sub(start))
    });
    covered >= span.end.saturating_sub(span.start)
}

fn base_bold_for_edit(source: &Snapshot, operations: &[Operation]) -> Result<bool, Error> {
    let bold_spans = operations
        .iter()
        .filter_map(|operation| match operation {
            Operation::Bold { span, .. } => Some(*span),
            Operation::Text { .. }
            | Operation::Alignment { .. }
            | Operation::ParagraphLayout { .. }
            | Operation::Italic { .. }
            | Operation::Underline { .. }
            | Operation::Strike { .. }
            | Operation::InsertParagraph { .. }
            | Operation::RemoveParagraph { .. }
            | Operation::RestoreParagraph { .. }
            | Operation::MoveParagraph { .. }
            | Operation::TableCellText { .. }
            | Operation::HeaderFooterText { .. }
            | Operation::AnnotationText { .. }
            | Operation::NoteText { .. }
            | Operation::ShapeText { .. }
            | Operation::PicturePayload(_)
            | Operation::PictureRemoval(_)
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
            if !span_fully_covered(run_span, &bold_spans) {
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
                | Operation::ParagraphLayout { .. }
                | Operation::Italic { .. }
                | Operation::Underline { .. }
                | Operation::Strike { .. }
                | Operation::InsertParagraph { .. }
                | Operation::RemoveParagraph { .. }
                | Operation::RestoreParagraph { .. }
                | Operation::MoveParagraph { .. }
                | Operation::TableCellText { .. }
                | Operation::HeaderFooterText { .. }
                | Operation::AnnotationText { .. }
                | Operation::NoteText { .. }
                | Operation::ShapeText { .. }
                | Operation::PicturePayload(_)
                | Operation::PictureRemoval(_)
                | Operation::RootTransfer { .. } => None,
            })
            .unwrap_or(false)
    }))
}

fn uniform_body_italic(source: &Snapshot) -> Result<bool, Error> {
    let mut value = None;
    for run in source.body().runs() {
        let italic = run.format().italic();
        if value.is_some_and(|existing| existing != italic) {
            return Err(Error::UnsupportedSource(
                "the body has mixed character formatting",
            ));
        }
        value = Some(italic);
    }
    Ok(value.unwrap_or(false))
}

fn base_italic_for_edit(source: &Snapshot, operations: &[Operation]) -> Result<bool, Error> {
    let italic_spans = operations
        .iter()
        .filter_map(|operation| match operation {
            Operation::Italic { span, .. } => Some(*span),
            Operation::Text { .. }
            | Operation::Alignment { .. }
            | Operation::ParagraphLayout { .. }
            | Operation::Bold { .. }
            | Operation::Underline { .. }
            | Operation::Strike { .. }
            | Operation::InsertParagraph { .. }
            | Operation::RemoveParagraph { .. }
            | Operation::RestoreParagraph { .. }
            | Operation::MoveParagraph { .. }
            | Operation::TableCellText { .. }
            | Operation::HeaderFooterText { .. }
            | Operation::AnnotationText { .. }
            | Operation::NoteText { .. }
            | Operation::ShapeText { .. }
            | Operation::PicturePayload(_)
            | Operation::PictureRemoval(_)
            | Operation::RootTransfer { .. } => None,
        })
        .collect::<Vec<_>>();
    if italic_spans.is_empty() {
        return uniform_body_italic(source);
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
            if !span_fully_covered(run_span, &italic_spans) {
                let value = run.format().italic();
                if base.is_some_and(|existing| existing != value) {
                    return Err(Error::UnsupportedSource(
                        "unselected body text has mixed italic state",
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
                Operation::Italic { after, .. } => Some(*after),
                Operation::Text { .. }
                | Operation::Alignment { .. }
                | Operation::ParagraphLayout { .. }
                | Operation::Bold { .. }
                | Operation::Underline { .. }
                | Operation::Strike { .. }
                | Operation::InsertParagraph { .. }
                | Operation::RemoveParagraph { .. }
                | Operation::RestoreParagraph { .. }
                | Operation::MoveParagraph { .. }
                | Operation::TableCellText { .. }
                | Operation::HeaderFooterText { .. }
                | Operation::AnnotationText { .. }
                | Operation::NoteText { .. }
                | Operation::ShapeText { .. }
                | Operation::PicturePayload(_)
                | Operation::PictureRemoval(_)
                | Operation::RootTransfer { .. } => None,
            })
            .unwrap_or(false)
    }))
}

fn uniform_body_underline(source: &Snapshot) -> Result<UnderlineStyle, Error> {
    let mut value = None;
    for run in source.body().runs() {
        let underline = run.format().underline();
        if value.is_some_and(|existing| existing != underline) {
            return Err(Error::UnsupportedSource(
                "the body has mixed character formatting",
            ));
        }
        value = Some(underline);
    }
    Ok(value.unwrap_or(UnderlineStyle::None))
}

fn base_underline_for_edit(
    source: &Snapshot,
    operations: &[Operation],
) -> Result<UnderlineStyle, Error> {
    let underline_spans = operations
        .iter()
        .filter_map(|operation| match operation {
            Operation::Underline { span, .. } => Some(*span),
            Operation::Text { .. }
            | Operation::Alignment { .. }
            | Operation::ParagraphLayout { .. }
            | Operation::Bold { .. }
            | Operation::Italic { .. }
            | Operation::Strike { .. }
            | Operation::InsertParagraph { .. }
            | Operation::RemoveParagraph { .. }
            | Operation::RestoreParagraph { .. }
            | Operation::MoveParagraph { .. }
            | Operation::TableCellText { .. }
            | Operation::HeaderFooterText { .. }
            | Operation::AnnotationText { .. }
            | Operation::NoteText { .. }
            | Operation::ShapeText { .. }
            | Operation::PicturePayload(_)
            | Operation::PictureRemoval(_)
            | Operation::RootTransfer { .. } => None,
        })
        .collect::<Vec<_>>();
    if underline_spans.is_empty() {
        return uniform_body_underline(source);
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
            if !span_fully_covered(run_span, &underline_spans) {
                let value = run.format().underline();
                if base.is_some_and(|existing| existing != value) {
                    return Err(Error::UnsupportedSource(
                        "unselected body text has mixed underline state",
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
                Operation::Underline { after, .. } => Some(*after),
                Operation::Text { .. }
                | Operation::Alignment { .. }
                | Operation::ParagraphLayout { .. }
                | Operation::Bold { .. }
                | Operation::Italic { .. }
                | Operation::Strike { .. }
                | Operation::InsertParagraph { .. }
                | Operation::RemoveParagraph { .. }
                | Operation::RestoreParagraph { .. }
                | Operation::MoveParagraph { .. }
                | Operation::TableCellText { .. }
                | Operation::HeaderFooterText { .. }
                | Operation::AnnotationText { .. }
                | Operation::NoteText { .. }
                | Operation::ShapeText { .. }
                | Operation::PicturePayload(_)
                | Operation::PictureRemoval(_)
                | Operation::RootTransfer { .. } => None,
            })
            .unwrap_or(UnderlineStyle::None)
    }))
}

fn uniform_body_strike(source: &Snapshot) -> Result<bool, Error> {
    let mut value = None;
    for run in source.body().runs() {
        let strike = run.format().raw().strike;
        if value.is_some_and(|existing| existing != strike) {
            return Err(Error::UnsupportedSource(
                "the body has mixed character formatting",
            ));
        }
        value = Some(strike);
    }
    Ok(value.unwrap_or(false))
}

fn base_strike_for_edit(source: &Snapshot, operations: &[Operation]) -> Result<bool, Error> {
    let strike_spans = operations
        .iter()
        .filter_map(|operation| match operation {
            Operation::Strike { span, .. } => Some(*span),
            Operation::Text { .. }
            | Operation::Alignment { .. }
            | Operation::ParagraphLayout { .. }
            | Operation::Bold { .. }
            | Operation::Italic { .. }
            | Operation::Underline { .. }
            | Operation::InsertParagraph { .. }
            | Operation::RemoveParagraph { .. }
            | Operation::RestoreParagraph { .. }
            | Operation::MoveParagraph { .. }
            | Operation::TableCellText { .. }
            | Operation::HeaderFooterText { .. }
            | Operation::AnnotationText { .. }
            | Operation::NoteText { .. }
            | Operation::ShapeText { .. }
            | Operation::PicturePayload(_)
            | Operation::PictureRemoval(_)
            | Operation::RootTransfer { .. } => None,
        })
        .collect::<Vec<_>>();
    if strike_spans.is_empty() {
        return uniform_body_strike(source);
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
            if !span_fully_covered(run_span, &strike_spans) {
                let value = run.format().raw().strike;
                if base.is_some_and(|existing| existing != value) {
                    return Err(Error::UnsupportedSource(
                        "unselected body text has mixed strike state",
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
                Operation::Strike { after, .. } => Some(*after),
                Operation::Text { .. }
                | Operation::Alignment { .. }
                | Operation::ParagraphLayout { .. }
                | Operation::Bold { .. }
                | Operation::Italic { .. }
                | Operation::Underline { .. }
                | Operation::InsertParagraph { .. }
                | Operation::RemoveParagraph { .. }
                | Operation::RestoreParagraph { .. }
                | Operation::MoveParagraph { .. }
                | Operation::TableCellText { .. }
                | Operation::HeaderFooterText { .. }
                | Operation::AnnotationText { .. }
                | Operation::NoteText { .. }
                | Operation::ShapeText { .. }
                | Operation::PicturePayload(_)
                | Operation::PictureRemoval(_)
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

fn italic_for_span(source: &Snapshot, span: TextSpan) -> Result<bool, Error> {
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
                let italic = run.format().italic();
                if value.is_some_and(|existing| existing != italic) {
                    return Err(Error::UnsupportedSource(
                        "the selected character span has mixed italic state",
                    ));
                }
                value = Some(italic);
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

fn underline_for_span(source: &Snapshot, span: TextSpan) -> Result<UnderlineStyle, Error> {
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
                let underline = run.format().underline();
                if value.is_some_and(|existing| existing != underline) {
                    return Err(Error::UnsupportedSource(
                        "the selected character span has mixed underline state",
                    ));
                }
                value = Some(underline);
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

fn strike_for_span(source: &Snapshot, span: TextSpan) -> Result<bool, Error> {
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
                let strike = run.format().raw().strike;
                if value.is_some_and(|existing| existing != strike) {
                    return Err(Error::UnsupportedSource(
                        "the selected character span has mixed strike state",
                    ));
                }
                value = Some(strike);
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

fn formatting_for_span(
    source: &Snapshot,
    span: TextSpan,
) -> Result<crate::types::Formatting, Error> {
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
                let formatting = *run.format().raw();
                if value.is_some_and(|existing| existing != formatting) {
                    return Err(Error::UnsupportedSource(
                        "the selected character span has mixed formatting",
                    ));
                }
                value = Some(formatting);
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
            | Operation::ParagraphLayout { .. }
            | Operation::Bold { .. }
            | Operation::Italic { .. }
            | Operation::Underline { .. }
            | Operation::Strike { .. }
            | Operation::RemoveParagraph { .. }
            | Operation::RestoreParagraph { .. }
            | Operation::MoveParagraph { .. }
            | Operation::TableCellText { .. }
            | Operation::HeaderFooterText { .. }
            | Operation::AnnotationText { .. }
            | Operation::NoteText { .. }
            | Operation::ShapeText { .. }
            | Operation::PicturePayload(_)
            | Operation::PictureRemoval(_)
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

#[inline(never)]
fn retained_or_located_body_source_span(
    source: &Snapshot,
    source_bytes: &[u8],
) -> Result<Range<usize>, Error> {
    match source.model().ordinary_body_source_span() {
        Some(span)
            if source_bytes.is_ascii()
                && span.start <= span.end
                && span.end <= source_bytes.len() =>
        {
            Ok(span)
        },
        Some(_) | None => ordinary_body_source_span(source_bytes, source.text(), source.limits()),
    }
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
    target: &Snapshot,
    operations: &[Operation],
) -> Result<(), Error> {
    let source_nodes = source
        .opaque()
        .iter()
        .filter(|node| matches!(node.anchor(), crate::opaque::Anchor::Body(_)))
        .collect::<Vec<_>>();
    let target_nodes = target
        .opaque()
        .iter()
        .filter(|node| matches!(node.anchor(), crate::opaque::Anchor::Body(_)))
        .collect::<Vec<_>>();
    if source_nodes.len() != target_nodes.len() {
        return Err(Error::UnsupportedSource(
            "body-anchored opaque syntax changed during publication",
        ));
    }
    for (source_node, target_node) in source_nodes.iter().zip(&target_nodes) {
        let crate::opaque::Anchor::Body(source_position) = source_node.anchor() else {
            return Err(Error::UnsupportedSource(
                "source opaque anchor is not a body position",
            ));
        };
        let crate::opaque::Anchor::Body(target_position) = target_node.anchor() else {
            return Err(Error::UnsupportedSource(
                "target opaque anchor is not a body position",
            ));
        };
        let projected_position = project_base_position(source_position, operations)?;
        if target_position != projected_position
            || target_node.kind() != source_node.kind()
            || target_node.source() != source_node.source()
        {
            return Err(Error::UnsupportedSource(
                "body-anchored opaque syntax changed during publication",
            ));
        }
        if source_node.source().is_empty() {
            return Err(Error::UnsupportedSource(
                "an empty body-anchored opaque node cannot be located",
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

#[derive(Clone, Copy)]
enum CharacterPropertyChange {
    Bold(bool),
    Italic(bool),
    Underline(UnderlineStyle),
    Strike(bool),
}

fn encoded_body_with_properties(
    text: &str,
    alignments: &[Alignment],
    baseline: crate::types::Formatting,
    base_bold: bool,
    bold_changes: &[(TextSpan, bool)],
    base_italic: bool,
    italic_changes: &[(TextSpan, bool)],
    base_underline: UnderlineStyle,
    underline_changes: &[(TextSpan, UnderlineStyle)],
    base_strike: bool,
    strike_changes: &[(TextSpan, bool)],
    limits: crate::ParseLimits,
) -> Result<Vec<u8>, Error> {
    let extra = alignments
        .len()
        .checked_mul(4)
        .and_then(|bytes| bytes.checked_add(bold_changes.len().saturating_mul(6)))
        .and_then(|bytes| bytes.checked_add(italic_changes.len().saturating_mul(6)))
        .and_then(|bytes| bytes.checked_add(underline_changes.len().saturating_mul(12)))
        .and_then(|bytes| bytes.checked_add(strike_changes.len().saturating_mul(6)))
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
        output.extend_from_slice(br"\plain ");
        RtfWriter::new(&mut output)
            .write_formatting(&baseline)
            .map_err(|error| Error::Write(error.to_string()))?;
        write_alignment(&mut output, alignments.get(paragraph).copied())?;
        let mut cursor = paragraph_start;
        let mut paragraph_changes = bold_changes
            .iter()
            .copied()
            .filter(|(span, _)| span.start >= paragraph_start && span.end <= paragraph_end)
            .map(|(span, value)| (span, CharacterPropertyChange::Bold(value)))
            .chain(
                italic_changes
                    .iter()
                    .copied()
                    .filter(|(span, _)| span.start >= paragraph_start && span.end <= paragraph_end)
                    .map(|(span, value)| (span, CharacterPropertyChange::Italic(value))),
            )
            .chain(
                underline_changes
                    .iter()
                    .copied()
                    .filter(|(span, _)| span.start >= paragraph_start && span.end <= paragraph_end)
                    .map(|(span, value)| (span, CharacterPropertyChange::Underline(value))),
            )
            .chain(
                strike_changes
                    .iter()
                    .copied()
                    .filter(|(span, _)| span.start >= paragraph_start && span.end <= paragraph_end)
                    .map(|(span, value)| (span, CharacterPropertyChange::Strike(value))),
            )
            .collect::<Vec<_>>();
        paragraph_changes.sort_unstable_by_key(|(span, _)| (span.start, span.end));
        for (span, change) in paragraph_changes {
            write_encoded_fragment(&mut output, text, cursor..span.start)?;
            match change {
                CharacterPropertyChange::Bold(value) => {
                    write_bold(&mut output, value);
                    write_encoded_fragment(&mut output, text, span.start..span.end)?;
                    write_bold(&mut output, base_bold);
                },
                CharacterPropertyChange::Italic(value) => {
                    // The parser flushes body text at bold controls but not at
                    // italic controls. Reasserting the current bold state around
                    // this italic fragment creates a run boundary without
                    // changing the effective formatting or source envelope.
                    write_bold(&mut output, base_bold);
                    write_italic(&mut output, value);
                    write_encoded_fragment(&mut output, text, span.start..span.end)?;
                    write_bold(&mut output, base_bold);
                    write_italic(&mut output, base_italic);
                },
                CharacterPropertyChange::Underline(value) => {
                    // Underline controls do not by themselves flush all parser
                    // run state. Bold controls force the same safe boundary as
                    // the existing italic property path.
                    write_bold(&mut output, base_bold);
                    write_underline(&mut output, value);
                    write_encoded_fragment(&mut output, text, span.start..span.end)?;
                    write_bold(&mut output, base_bold);
                    write_underline(&mut output, base_underline);
                },
                CharacterPropertyChange::Strike(value) => {
                    // Strike controls do not by themselves flush all parser
                    // run state. Bold controls force the same safe boundary
                    // while the baseline's double-strike facet remains intact.
                    write_bold(&mut output, base_bold);
                    write_strike(&mut output, value);
                    write_encoded_fragment(&mut output, text, span.start..span.end)?;
                    write_bold(&mut output, base_bold);
                    write_strike(&mut output, base_strike);
                },
            }
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

fn encoded_body_with_paragraph_properties(
    source: &Snapshot,
    properties: &[crate::types::Paragraph],
    limits: crate::ParseLimits,
) -> Result<Vec<u8>, Error> {
    let mut output = BoundedVec::new(limits.max_source_bytes());
    let mut body_position = 0usize;
    let mut paragraph_count = 0usize;
    for (position, paragraph) in source.body().paragraphs().enumerate() {
        let properties = properties
            .get(position)
            .ok_or(Error::StructuralPropertyConflict)?;
        let paragraph_end =
            body_position
                .checked_add(paragraph.len())
                .ok_or(Error::InputTooLarge {
                    observed: usize::MAX,
                    limit: limits.max_source_bytes(),
                })?;
        let terminated = source.text().as_bytes().get(paragraph_end) == Some(&b'\n');
        let mut runs = paragraph.runs().peekable();
        if runs.peek().is_none() {
            write_bounded(&mut output, br"\plain\pard ")?;
            let mut writer = RtfWriter::new(&mut output);
            writer
                .write_paragraph_properties(properties)
                .map_err(|error| Error::Write(error.to_string()))?;
            if terminated {
                write_bounded(&mut output, br"\par ")?;
            }
        } else {
            while let Some(run) = runs.next() {
                write_bounded(&mut output, br"\plain\pard ")?;
                let mut writer = RtfWriter::new(&mut output);
                writer
                    .write_formatting(run.format().raw())
                    .and_then(|()| writer.write_paragraph_properties(properties))
                    .map_err(|error| Error::Write(error.to_string()))?;
                write_bounded(&mut output, b" ")?;
                RtfWriter::new(&mut output)
                    .write_text(run.text())
                    .map_err(|error| Error::Write(error.to_string()))?;
                if terminated && runs.peek().is_none() {
                    write_bounded(&mut output, br"\par ")?;
                }
            }
        }
        body_position = paragraph_end.saturating_add(usize::from(terminated));
        paragraph_count = paragraph_count.saturating_add(1);
    }
    if paragraph_count != properties.len() {
        return Err(Error::StructuralPropertyConflict);
    }
    Ok(output.into_inner())
}

fn write_bounded(output: &mut BoundedVec, bytes: &[u8]) -> Result<(), Error> {
    output
        .write_all(bytes)
        .map_err(|error| Error::Write(error.to_string()))
}

struct BoundedVec {
    bytes: Vec<u8>,
    limit: usize,
}

impl BoundedVec {
    const fn new(limit: usize) -> Self {
        Self {
            bytes: Vec::new(),
            limit,
        }
    }

    fn into_inner(self) -> Vec<u8> {
        self.bytes
    }
}

impl io::Write for BoundedVec {
    fn write(&mut self, input: &[u8]) -> io::Result<usize> {
        let observed = self
            .bytes
            .len()
            .checked_add(input.len())
            .ok_or_else(|| io::Error::other("RTF paragraph-layout output size overflow"))?;
        if observed > self.limit {
            return Err(io::Error::other(
                "RTF paragraph-layout output exceeds the source limit",
            ));
        }
        self.bytes
            .try_reserve(input.len())
            .map_err(|_error| io::Error::other("could not reserve paragraph-layout output"))?;
        self.bytes.extend_from_slice(input);
        Ok(input.len())
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
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

fn write_italic(output: &mut Vec<u8>, italic: bool) {
    output.extend_from_slice(if italic { br"\i " } else { br"\i0 " });
}

fn write_underline(output: &mut Vec<u8>, underline: UnderlineStyle) {
    output.extend_from_slice(match underline {
        UnderlineStyle::None => br"\ulnone ",
        UnderlineStyle::Single => br"\ul ",
        UnderlineStyle::Double => br"\uldb ",
        UnderlineStyle::Dotted => br"\uld ",
        UnderlineStyle::Dashed => br"\uldash ",
        UnderlineStyle::DashDot => br"\uldashd ",
        UnderlineStyle::DashDotDot => br"\uldashdd ",
        UnderlineStyle::Words => br"\ulw ",
        UnderlineStyle::Thick => br"\ulth ",
        UnderlineStyle::Wave => br"\ulwave ",
        UnderlineStyle::Hairline => br"\ulhair ",
        UnderlineStyle::ThickDotted => br"\ulthd ",
        UnderlineStyle::ThickDashed => br"\ulthdash ",
        UnderlineStyle::ThickDashDot => br"\ulthdashd ",
        UnderlineStyle::ThickDashDotDot => br"\ulthdashdd ",
        UnderlineStyle::ThickLongDash => br"\ulthldash ",
        UnderlineStyle::LongDash => br"\ulldash ",
        UnderlineStyle::HeavyWave => br"\ulhwave ",
        UnderlineStyle::DoubleWave => br"\ululdbwave ",
    });
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

fn write_strike(output: &mut Vec<u8>, strike: bool) {
    if strike {
        output.extend_from_slice(br"\strike ");
    } else {
        output.extend_from_slice(br"\strike0 ");
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
    ParagraphLayout {
        position: usize,
        fields: LayoutFields,
        before: ParagraphLayout,
        after: ParagraphLayout,
    },
    Bold {
        span: TextSpan,
        after_span: TextSpan,
        before: bool,
        after: bool,
    },
    Italic {
        span: TextSpan,
        after_span: TextSpan,
        before: bool,
        after: bool,
    },
    Underline {
        span: TextSpan,
        after_span: TextSpan,
        before: UnderlineStyle,
        after: UnderlineStyle,
    },
    Strike {
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
    RemoveParagraph {
        position: usize,
        text: String,
        restoring: bool,
    },
    MoveParagraph {
        position: usize,
        final_position: usize,
        text: String,
    },
    SplitParagraph {
        position: usize,
        offset: usize,
        before: String,
        boundary: Vec<u8>,
        left: String,
        right: String,
    },
    MergeParagraph {
        position: usize,
        boundary: Vec<u8>,
        left: String,
        right: String,
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
    PicturePayload(picture_payload::StagedPicturePayload),
    PictureRemoval {
        position: usize,
        group_start: usize,
        group: Vec<u8>,
        removing: bool,
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
            Self::ParagraphLayout {
                position,
                fields,
                before,
                after,
            } => Self::ParagraphLayout {
                position: *position,
                fields: *fields,
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
            Self::Italic {
                span,
                after_span,
                before,
                after,
            } => Self::Italic {
                span: *after_span,
                after_span: *span,
                before: *after,
                after: *before,
            },
            Self::Underline {
                span,
                after_span,
                before,
                after,
            } => Self::Underline {
                span: *after_span,
                after_span: *span,
                before: *after,
                after: *before,
            },
            Self::Strike {
                span,
                after_span,
                before,
                after,
            } => Self::Strike {
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
            Self::RemoveParagraph {
                position,
                text,
                restoring,
            } => Self::RemoveParagraph {
                position: *position,
                text: text.clone(),
                restoring: !restoring,
            },
            Self::MoveParagraph {
                position,
                final_position,
                text,
            } => Self::MoveParagraph {
                position: *final_position,
                final_position: *position,
                text: text.clone(),
            },
            Self::SplitParagraph {
                position,
                boundary,
                left,
                right,
                ..
            } => Self::MergeParagraph {
                position: *position,
                boundary: boundary.clone(),
                left: left.clone(),
                right: right.clone(),
            },
            Self::MergeParagraph {
                position,
                boundary,
                left,
                right,
            } => Self::SplitParagraph {
                position: *position,
                offset: left.len(),
                boundary: boundary.clone(),
                before: {
                    let mut before = String::new();
                    before.push_str(left);
                    before.push_str(right);
                    before
                },
                left: left.clone(),
                right: right.clone(),
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
            Self::PicturePayload(operation) => Self::PicturePayload(operation.inverse()),
            Self::PictureRemoval {
                position,
                group_start,
                group,
                removing,
            } => Self::PictureRemoval {
                position: *position,
                group_start: *group_start,
                group: group.clone(),
                removing: !removing,
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
                raw_structure: None,
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
            Operation::ParagraphLayout {
                position,
                fields,
                before,
                after,
            } if before != after => Some(Change::ParagraphLayout {
                position: *position,
                fields: *fields,
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
            Operation::Italic {
                span,
                before,
                after,
            } if before != after => Some(Change::Italic {
                span: *span,
                after_span: project_base_span(*span, operations).ok()?,
                before: *before,
                after: *after,
            }),
            Operation::Underline {
                span,
                before,
                after,
            } if before != after => Some(Change::Underline {
                span: *span,
                after_span: project_base_span(*span, operations).ok()?,
                before: *before,
                after: *after,
            }),
            Operation::Strike {
                span,
                before,
                after,
            } if before != after => Some(Change::Strike {
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
            Operation::RemoveParagraph { position, text } => Some(Change::RemoveParagraph {
                position: *position,
                text: text.clone(),
                restoring: false,
            }),
            Operation::RestoreParagraph { position, text } => Some(Change::RemoveParagraph {
                position: *position,
                text: text.clone(),
                restoring: true,
            }),
            Operation::MoveParagraph {
                position,
                final_position,
                text,
            } if position != final_position => Some(Change::MoveParagraph {
                position: *position,
                final_position: *final_position,
                text: text.clone(),
            }),
            Operation::Text {
                raw_structure:
                    Some(RawParagraphOperation::Split {
                        position,
                        offset,
                        boundary,
                        before,
                        ..
                    }),
                ..
            } => {
                let (left, right) = (before.get(..*offset)?, before.get(*offset..)?);
                Some(Change::SplitParagraph {
                    position: *position,
                    offset: *offset,
                    before: before.clone(),
                    boundary: boundary.clone(),
                    left: left.to_string(),
                    right: right.to_string(),
                })
            },
            Operation::Text {
                raw_structure:
                    Some(RawParagraphOperation::Merge {
                        position,
                        boundary_bytes,
                        left,
                        right,
                        ..
                    }),
                ..
            } => Some(Change::MergeParagraph {
                position: *position,
                boundary: boundary_bytes.clone(),
                left: left.clone(),
                right: right.clone(),
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
            Operation::PicturePayload(operation) if operation.before != operation.after => {
                Some(Change::PicturePayload(operation.clone()))
            },
            Operation::PictureRemoval(operation) => Some(operation.inverse_change().inverse()),
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
            | Operation::ParagraphLayout { .. }
            | Operation::Bold { .. }
            | Operation::Italic { .. }
            | Operation::Underline { .. }
            | Operation::Strike { .. }
            | Operation::MoveParagraph { .. }
            | Operation::TableCellText { .. }
            | Operation::HeaderFooterText { .. }
            | Operation::AnnotationText { .. }
            | Operation::NoteText { .. }
            | Operation::ShapeText { .. }
            | Operation::PicturePayload(_)
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
    /// Paragraph-layout changes remain source-bound because canonical RTF
    /// rewrites cannot yet prove an imported exact artifact byte-for-byte;
    /// callers retain exact undo through [`Self::inverse`] instead.
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

        if self
            .changes
            .iter()
            .any(|change| matches!(change, Change::ParagraphLayout { .. }))
        {
            return Err(litchi_core::patch::PatchError::InvalidText {
                field: "RTF paragraph-layout durable patches are not supported",
            });
        }

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
                let forward = durable_operation(limits, change, before, after)?;
                let inverse_change = change.inverse();
                let inverse = durable_operation(limits, &inverse_change, after, before)?;
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
    target: &[u8],
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
        Change::ParagraphLayout {
            position: _,
            fields: _,
            before: _,
            after: _,
        } => Err(litchi_core::patch::PatchError::InvalidText {
            field: "RTF paragraph-layout durable patches are not supported",
        }),
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
        Change::Italic {
            span,
            before,
            after,
            after_span: _,
        } => {
            preconditions.insert("italic".to_string(), Value::Bool(*before));
            litchi_core::patch::PatchOperation::new(
                limits,
                "character-italic.set",
                format!("body:utf8:{}-{}", span.start, span.end),
                preconditions,
                Value::Bool(*after),
            )
        },
        Change::Underline {
            span,
            before,
            after,
            after_span: _,
        } => {
            preconditions.insert(
                "underline".to_string(),
                Value::String(underline_name(*before).to_string()),
            );
            litchi_core::patch::PatchOperation::new(
                limits,
                "character-underline.set",
                format!("body:utf8:{}-{}", span.start, span.end),
                preconditions,
                Value::String(underline_name(*after).to_string()),
            )
        },
        Change::Strike {
            span,
            before,
            after,
            after_span: _,
        } => {
            preconditions.insert("strike".to_string(), Value::Bool(*before));
            litchi_core::patch::PatchOperation::new(
                limits,
                "character-strike.set",
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
        Change::RemoveParagraph {
            position,
            text,
            restoring,
        } => {
            preconditions.insert(
                "text".to_string(),
                if *restoring {
                    Value::Null
                } else {
                    Value::String(text.clone())
                },
            );
            litchi_core::patch::PatchOperation::new(
                limits,
                if *restoring {
                    "paragraph.insert"
                } else {
                    "paragraph.remove"
                },
                format!("body:paragraph:{position}"),
                preconditions,
                if *restoring {
                    Value::String(text.clone())
                } else {
                    Value::Null
                },
            )
        },
        Change::MoveParagraph {
            position,
            final_position,
            text,
        } => {
            preconditions.insert("text".to_string(), Value::String(text.clone()));
            litchi_core::patch::PatchOperation::new(
                limits,
                "paragraph.move",
                format!("body:paragraph:{position}"),
                preconditions,
                Value::String(final_position.to_string()),
            )
        },
        Change::SplitParagraph {
            position,
            offset,
            before,
            boundary,
            left: _,
            right: _,
        } => {
            preconditions.insert("text".to_string(), Value::String(before.clone()));
            preconditions.insert("offset".to_string(), Value::from(*offset as u64));
            preconditions.insert("boundary".to_string(), Value::String(hex_encode(boundary)));
            preconditions.insert(
                "result_artifact_sha256".to_string(),
                Value::String(litchi_core::patch::BlobId::of(target).as_hex()),
            );
            preconditions.insert(
                "boundary_mode".to_string(),
                Value::String(
                    if boundary == PARAGRAPH_SPLIT_BYTES {
                        "canonical"
                    } else {
                        "exact"
                    }
                    .to_string(),
                ),
            );
            litchi_core::patch::PatchOperation::new(
                limits,
                "paragraph.split",
                format!("body:paragraph:{position}"),
                preconditions,
                Value::Null,
            )
        },
        Change::MergeParagraph {
            position,
            boundary,
            left,
            right,
        } => {
            preconditions.insert("left".to_string(), Value::String(left.clone()));
            preconditions.insert("right".to_string(), Value::String(right.clone()));
            preconditions.insert("boundary".to_string(), Value::String(hex_encode(boundary)));
            preconditions.insert(
                "result_artifact_sha256".to_string(),
                Value::String(litchi_core::patch::BlobId::of(target).as_hex()),
            );
            litchi_core::patch::PatchOperation::new(
                limits,
                "paragraph.merge",
                format!("body:paragraph:{position}"),
                preconditions,
                Value::Null,
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
        Change::PicturePayload(operation) => {
            picture_payload::durable_operation(limits, operation, source)
        },
        Change::PictureRemoval {
            position,
            group_start,
            group,
            removing,
        } => picture_payload::durable_removal_operation(
            limits,
            *position,
            *group_start,
            group,
            *removing,
            source,
        ),
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
    if patch.format() != "litchi-rtf" {
        return Err(Error::DurablePatch(
            "unsupported durable patch format".to_string(),
        ));
    }
    if patch.operations().is_empty() {
        if !patch.blobs().is_empty() {
            return Err(Error::DurablePatch(
                "empty patch has an unreferenced blob bundle".to_string(),
            ));
        }
        return Ok(source.clone());
    }
    if !patch.blobs().is_empty() {
        return Err(Error::DurablePatch(
            "RTF durable operations do not accept artifact blobs".to_string(),
        ));
    }
    let source_bytes = source
        .source_bytes()
        .ok_or(Error::UnsupportedSource("snapshot has no exact RTF source"))?;
    let source_hash = litchi_core::patch::BlobId::of(source_bytes).as_hex();
    let picture_operation_count = patch
        .operations()
        .iter()
        .filter(|operation| operation.op == "picture-payload.replace")
        .count();
    if picture_operation_count != 0 {
        if picture_operation_count != patch.operations().len() {
            return Err(Error::DurablePatch(
                "picture payload operations cannot compose with other RTF vocabularies".to_string(),
            ));
        }
        return picture_payload::apply_durable_patch(source, patch.operations(), &source_hash);
    }
    let picture_removal_count = patch
        .operations()
        .iter()
        .filter(|operation| operation.op == "picture.remove")
        .count();
    if picture_removal_count != 0 {
        if picture_removal_count != patch.operations().len() {
            return Err(Error::DurablePatch(
                "picture removal operations cannot compose with other RTF vocabularies".to_string(),
            ));
        }
        return picture_payload::apply_durable_removal_patch(
            source,
            patch.operations(),
            &source_hash,
        );
    }
    let picture_insertion_count = patch
        .operations()
        .iter()
        .filter(|operation| operation.op == "picture.insert-exact")
        .count();
    if picture_insertion_count != 0 {
        if picture_insertion_count != patch.operations().len() {
            return Err(Error::DurablePatch(
                "exact picture insertion operations cannot compose with other RTF vocabularies"
                    .to_string(),
            ));
        }
        return picture_payload::apply_durable_insertion_patch(
            source,
            patch.operations(),
            &source_hash,
        );
    }
    // A durable operation may be admitted under a caller-selected patch
    // bound larger than the ordinary edit default.  Bind replay to the
    // validated patch's operation count so a valid 257+ operation patch does
    // not fail merely because `Snapshot::edit()` defaults to 256.
    let mut edit = source.edit_with_limits(Limits::new(patch.operations().len()));
    let mut expected_result_hash = None;
    for operation in patch.operations() {
        let expected_preconditions = match operation.op.as_str() {
            "paragraph.split" => 6,
            "paragraph.merge" => 5,
            _ => 2,
        };
        if operation.preconditions.len() != expected_preconditions {
            return Err(Error::DurablePatch(
                "operation has an invalid precondition count".to_string(),
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
            "paragraph-layout.patch" => {
                return Err(Error::DurablePatch(
                    "paragraph-layout durable patches are not supported".to_string(),
                ));
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
            "character-italic.set" => {
                let span = parse_text_target(&operation.target)?;
                let expected = operation
                    .preconditions
                    .get("italic")
                    .and_then(Value::as_bool)
                    .ok_or_else(|| {
                        Error::DurablePatch("missing italic precondition".to_string())
                    })?;
                if italic_for_span(source, span)? != expected {
                    return Err(Error::StalePrecondition("character italic state differs"));
                }
                let replacement = operation.value.as_bool().ok_or_else(|| {
                    Error::DurablePatch("italic value must be Boolean".to_string())
                })?;
                edit.set_text_italic(span, replacement)?;
            },
            "character-underline.set" => {
                let span = parse_text_target(&operation.target)?;
                let expected = operation
                    .preconditions
                    .get("underline")
                    .and_then(Value::as_str)
                    .ok_or_else(|| {
                        Error::DurablePatch("missing underline precondition".to_string())
                    })?;
                if underline_name(underline_for_span(source, span)?) != expected {
                    return Err(Error::StalePrecondition(
                        "character underline state differs",
                    ));
                }
                let replacement = operation
                    .value
                    .as_str()
                    .and_then(parse_underline)
                    .ok_or_else(|| {
                        Error::DurablePatch("underline value must be a string".to_string())
                    })?;
                edit.set_text_underline(span, replacement)?;
            },
            "character-strike.set" => {
                let span = parse_text_target(&operation.target)?;
                let expected = operation
                    .preconditions
                    .get("strike")
                    .and_then(Value::as_bool)
                    .ok_or_else(|| {
                        Error::DurablePatch("missing strike precondition".to_string())
                    })?;
                if strike_for_span(source, span)? != expected {
                    return Err(Error::StalePrecondition("character strike state differs"));
                }
                let replacement = operation.value.as_bool().ok_or_else(|| {
                    Error::DurablePatch("strike value must be Boolean".to_string())
                })?;
                edit.set_text_strike(span, replacement)?;
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
            "paragraph.remove" => {
                let position = parse_paragraph_target(&operation.target)?;
                let expected = operation
                    .preconditions
                    .get("text")
                    .and_then(Value::as_str)
                    .ok_or_else(|| {
                        Error::DurablePatch("missing removed paragraph precondition".to_string())
                    })?;
                if paragraph_text_at(source, position)? != expected {
                    return Err(Error::StalePrecondition("paragraph text differs"));
                }
                if !operation.value.is_null() {
                    return Err(Error::DurablePatch(
                        "paragraph removal value must be null".to_string(),
                    ));
                }
                edit.remove_paragraph(position)?;
            },
            "paragraph.insert" => {
                let position = parse_paragraph_target(&operation.target)?;
                if !operation
                    .preconditions
                    .get("text")
                    .is_some_and(Value::is_null)
                {
                    return Err(Error::DurablePatch(
                        "paragraph insertion precondition must be null".to_string(),
                    ));
                }
                let text = operation.value.as_str().ok_or_else(|| {
                    Error::DurablePatch("inserted paragraph text must be a string".to_string())
                })?;
                edit.restore_paragraph(position, text)?;
            },
            "paragraph.move" => {
                let position = parse_paragraph_target(&operation.target)?;
                let expected = operation
                    .preconditions
                    .get("text")
                    .and_then(Value::as_str)
                    .ok_or_else(|| {
                        Error::DurablePatch("missing moved paragraph precondition".to_string())
                    })?;
                if paragraph_text_at(source, position)? != expected {
                    return Err(Error::StalePrecondition("moved paragraph text differs"));
                }
                let final_position = operation
                    .value
                    .as_str()
                    .ok_or_else(|| {
                        Error::DurablePatch("paragraph final position must be a string".to_string())
                    })?
                    .parse::<usize>()
                    .map_err(|_error| {
                        Error::DurablePatch("invalid paragraph final position".to_string())
                    })?;
                edit.move_paragraph(position, final_position)?;
            },
            "paragraph.split" => {
                let position = parse_paragraph_target(&operation.target)?;
                if !operation.value.is_null() {
                    return Err(Error::DurablePatch(
                        "paragraph split value must be null".to_string(),
                    ));
                }
                let expected = operation
                    .preconditions
                    .get("text")
                    .and_then(Value::as_str)
                    .ok_or_else(|| {
                        Error::DurablePatch("missing split paragraph text precondition".to_string())
                    })?;
                let offset = operation
                    .preconditions
                    .get("offset")
                    .and_then(Value::as_u64)
                    .and_then(|value| usize::try_from(value).ok())
                    .ok_or_else(|| {
                        Error::DurablePatch("invalid split paragraph offset".to_string())
                    })?;
                let encoded_boundary = operation
                    .preconditions
                    .get("boundary")
                    .and_then(Value::as_str)
                    .ok_or_else(|| {
                        Error::DurablePatch("missing split paragraph boundary".to_string())
                    })?;
                let boundary = hex_decode(encoded_boundary, source.limits().max_source_bytes())?;
                let result_hash = operation
                    .preconditions
                    .get("result_artifact_sha256")
                    .and_then(Value::as_str)
                    .ok_or_else(|| {
                        Error::DurablePatch(
                            "missing split paragraph result artifact digest".to_string(),
                        )
                    })?;
                if expected_result_hash.is_some() {
                    return Err(Error::DurablePatch(
                        "multiple structural durable results cannot compose".to_string(),
                    ));
                }
                expected_result_hash = Some(result_hash);
                let boundary_mode = operation
                    .preconditions
                    .get("boundary_mode")
                    .and_then(Value::as_str)
                    .ok_or_else(|| {
                        Error::DurablePatch("missing split paragraph boundary mode".to_string())
                    })?;
                let expected_boundary_mode = if boundary == PARAGRAPH_SPLIT_BYTES {
                    "canonical"
                } else {
                    "exact"
                };
                if boundary_mode != expected_boundary_mode {
                    return Err(Error::StalePrecondition(
                        "ordinary paragraph split boundary mode differs",
                    ));
                }
                if paragraph_text_at(source, position)? != expected {
                    return Err(Error::StalePrecondition(
                        "ordinary paragraph split text differs",
                    ));
                }
                if boundary_mode == "canonical" {
                    edit.split_paragraph(position, offset)?;
                } else {
                    edit.split_paragraph_with_boundary(position, offset, &boundary)?;
                }
                let generated_boundary =
                    edit.operations
                        .last()
                        .and_then(|operation| match operation {
                            Operation::Text {
                                raw_structure:
                                    Some(RawParagraphOperation::Split {
                                        position: generated_position,
                                        offset: generated_offset,
                                        boundary: generated_boundary,
                                        ..
                                    }),
                                ..
                            } if *generated_position == position && *generated_offset == offset => {
                                Some(generated_boundary.as_slice())
                            },
                            _ => None,
                        });
                if generated_boundary != Some(boundary.as_slice()) {
                    return Err(Error::StalePrecondition(
                        "ordinary paragraph split boundary differs",
                    ));
                }
            },
            "paragraph.merge" => {
                let position = parse_paragraph_target(&operation.target)?;
                if !operation.value.is_null() {
                    return Err(Error::DurablePatch(
                        "paragraph merge value must be null".to_string(),
                    ));
                }
                let expected_left = operation
                    .preconditions
                    .get("left")
                    .and_then(Value::as_str)
                    .ok_or_else(|| {
                        Error::DurablePatch("missing merge left paragraph precondition".to_string())
                    })?;
                let expected_right = operation
                    .preconditions
                    .get("right")
                    .and_then(Value::as_str)
                    .ok_or_else(|| {
                        Error::DurablePatch(
                            "missing merge right paragraph precondition".to_string(),
                        )
                    })?;
                let encoded_boundary = operation
                    .preconditions
                    .get("boundary")
                    .and_then(Value::as_str)
                    .ok_or_else(|| {
                        Error::DurablePatch("missing merge paragraph boundary".to_string())
                    })?;
                let boundary = hex_decode(encoded_boundary, source.limits().max_source_bytes())?;
                let result_hash = operation
                    .preconditions
                    .get("result_artifact_sha256")
                    .and_then(Value::as_str)
                    .ok_or_else(|| {
                        Error::DurablePatch(
                            "missing merge paragraph result artifact digest".to_string(),
                        )
                    })?;
                if expected_result_hash.is_some() {
                    return Err(Error::DurablePatch(
                        "multiple structural durable results cannot compose".to_string(),
                    ));
                }
                expected_result_hash = Some(result_hash);
                let second = position.checked_add(1).ok_or(Error::ParagraphOutOfRange {
                    position,
                    count: source.paragraph_count(),
                })?;
                if paragraph_text_at(source, position)? != expected_left
                    || paragraph_text_at(source, second)? != expected_right
                {
                    return Err(Error::StalePrecondition(
                        "ordinary paragraph merge text differs",
                    ));
                }
                let map = ordinary_paragraph_source_map(source)?;
                let actual_boundary = map
                    .paragraphs
                    .get(position)
                    .and_then(|paragraph| paragraph.boundary_after.as_ref())
                    .and_then(|span| source.source_bytes()?.get(span.clone()))
                    .ok_or(Error::UnsupportedSource(
                        "ordinary paragraph merge has no exact source boundary",
                    ))?;
                if actual_boundary != boundary.as_slice() {
                    return Err(Error::StalePrecondition(
                        "ordinary paragraph merge boundary bytes differ",
                    ));
                }
                edit.merge_paragraphs(position, second)?;
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
            "picture-payload.replace" => {
                picture_payload::apply_durable_operation(source, &mut edit, operation)?;
            },
            "picture.remove" | "picture.insert-exact" => {
                return Err(Error::DurablePatch(
                    "picture lifecycle operation escaped its atomic batch".to_string(),
                ));
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
    let snapshot = edit.commit()?.into_snapshot();
    if let Some(expected_hash) = expected_result_hash {
        let result_bytes = snapshot
            .source_bytes()
            .ok_or(Error::UnsupportedSource("snapshot has no exact RTF source"))?;
        let actual_hash = litchi_core::patch::BlobId::of(result_bytes).as_hex();
        if actual_hash != expected_hash {
            return Err(Error::StalePrecondition("durable paragraph result differs"));
        }
    }
    Ok(snapshot)
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
    let mut output = Vec::new();
    output
        .try_reserve_exact(observed)
        .map_err(|_error| Error::Write("could not reserve hexadecimal bytes".to_string()))?;
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

const fn underline_name(underline: UnderlineStyle) -> &'static str {
    match underline {
        UnderlineStyle::None => "none",
        UnderlineStyle::Single => "single",
        UnderlineStyle::Double => "double",
        UnderlineStyle::Dotted => "dotted",
        UnderlineStyle::Dashed => "dashed",
        UnderlineStyle::DashDot => "dash-dot",
        UnderlineStyle::DashDotDot => "dash-dot-dot",
        UnderlineStyle::Words => "words",
        UnderlineStyle::Thick => "thick",
        UnderlineStyle::Wave => "wave",
        UnderlineStyle::Hairline => "hairline",
        UnderlineStyle::ThickDotted => "thick-dotted",
        UnderlineStyle::ThickDashed => "thick-dashed",
        UnderlineStyle::ThickDashDot => "thick-dash-dot",
        UnderlineStyle::ThickDashDotDot => "thick-dash-dot-dot",
        UnderlineStyle::ThickLongDash => "thick-long-dash",
        UnderlineStyle::LongDash => "long-dash",
        UnderlineStyle::HeavyWave => "heavy-wave",
        UnderlineStyle::DoubleWave => "double-wave",
    }
}

fn parse_underline(value: &str) -> Option<UnderlineStyle> {
    match value {
        "none" => Some(UnderlineStyle::None),
        "single" => Some(UnderlineStyle::Single),
        "double" => Some(UnderlineStyle::Double),
        "dotted" => Some(UnderlineStyle::Dotted),
        "dashed" => Some(UnderlineStyle::Dashed),
        "dash-dot" => Some(UnderlineStyle::DashDot),
        "dash-dot-dot" => Some(UnderlineStyle::DashDotDot),
        "words" => Some(UnderlineStyle::Words),
        "thick" => Some(UnderlineStyle::Thick),
        "wave" => Some(UnderlineStyle::Wave),
        "hairline" => Some(UnderlineStyle::Hairline),
        "thick-dotted" => Some(UnderlineStyle::ThickDotted),
        "thick-dashed" => Some(UnderlineStyle::ThickDashed),
        "thick-dash-dot" => Some(UnderlineStyle::ThickDashDot),
        "thick-dash-dot-dot" => Some(UnderlineStyle::ThickDashDotDot),
        "thick-long-dash" => Some(UnderlineStyle::ThickLongDash),
        "long-dash" => Some(UnderlineStyle::LongDash),
        "heavy-wave" => Some(UnderlineStyle::HeavyWave),
        "double-wave" => Some(UnderlineStyle::DoubleWave),
        _ => None,
    }
}
