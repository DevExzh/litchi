//! Bounded immutable RTF body-text and paragraph-property transactions.
//!
//! Operations resolve against one immutable semantic body. Disjoint UTF-8
//! spans and paragraph-alignment facets compose in one atomic commit. The
//! safe closure deliberately excludes body-anchored opaque syntax, tables,
//! positioned content, and mixed formatting whose dependent ranges cannot be
//! updated losslessly yet. Header destinations and all bytes outside the
//! rewritten root-body span remain exact.

use crate::{Alignment, Document, RtfError, RtfWriter};
use bumpalo::Bump;
use serde_json::Value;
use std::collections::BTreeMap;
use std::fmt;
use std::ops::Range;

/// Immutable RTF snapshot used by the transaction API.
pub type Snapshot = Document;

/// Explicit budgeted undo/redo retention for immutable RTF snapshots.
pub type History = litchi_core::patch::History<Snapshot>;

/// Finite step and retained-weight bounds for [`History`].
pub use litchi_core::patch::HistoryLimits;

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
    /// A durable operation's semantic expectation is stale.
    StalePrecondition(&'static str),
    /// A durable patch uses an unsupported format-owned vocabulary.
    DurablePatch(String),
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
            Self::StalePrecondition(reason) => {
                write!(formatter, "stale RTF patch precondition: {reason}")
            },
            Self::DurablePatch(reason) => write!(formatter, "invalid durable RTF patch: {reason}"),
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
            if let Operation::Text {
                span: existing_span,
                ..
            } = operation
                && spans_conflict(*existing_span, span)
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
        self.ensure_operation_room()?;
        if self.operations.iter().any(|operation| {
            matches!(
                operation,
                Operation::Text {
                    structural: true,
                    ..
                }
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

        let (replacement, projected_spans) = project_text(&self.source, &self.operations)?;
        let mut alignments = source_alignments(&self.source);
        let mut property_operation = false;
        for operation in &self.operations {
            if let Operation::Alignment {
                position, after, ..
            } = operation
            {
                let count = alignments.len();
                let slot = alignments
                    .get_mut(*position)
                    .ok_or(Error::ParagraphOutOfRange {
                        position: *position,
                        count,
                    })?;
                *slot = *after;
                property_operation = true;
            }
        }
        let original_alignments = source_alignments(&self.source);
        let did_change = replacement != self.source.text() || alignments != original_alignments;
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
        self.source
            .model()
            .plain_body_text_editability()
            .map_err(Error::UnsupportedSource)?;
        let span =
            ordinary_body_source_span(source_bytes, self.source.text(), self.source.limits())?;
        validate_opaque_preservation(&self.source, source_bytes, &span)?;
        let replacement_bytes = if property_operation {
            encoded_body_with_alignments(&replacement, &alignments, self.source.limits())?
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
        Ok(Commit::new(
            self.source,
            snapshot,
            true,
            operation_count,
            semantic_delta,
        ))
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
    let mut selected = None;
    for (paragraph_position, paragraph) in source.body().paragraphs().enumerate() {
        let end = start
            .checked_add(paragraph.len())
            .ok_or(Error::InputTooLarge {
                observed: usize::MAX,
                limit: source.limits().max_source_bytes(),
            })?;
        if paragraph_position == position {
            selected = Some(start..end);
            break;
        }
        start = end.checked_add(1).ok_or(Error::InputTooLarge {
            observed: usize::MAX,
            limit: source.limits().max_source_bytes(),
        })?;
    }
    let count = source.paragraph_count();
    selected.ok_or(Error::ParagraphOutOfRange { position, count })
}

fn project_text(
    source: &Snapshot,
    operations: &[Operation],
) -> Result<(String, Vec<(usize, TextSpan)>), Error> {
    let mut text_operations = operations
        .iter()
        .enumerate()
        .filter_map(|(index, operation)| match operation {
            Operation::Text {
                span,
                after,
                before: _,
                structural: _,
            } => Some((index, *span, after.as_str())),
            Operation::Alignment { .. } => None,
        })
        .collect::<Vec<_>>();
    text_operations.sort_unstable_by_key(|(_, span, _)| (span.start, span.end));
    let source_text = source.text();
    let mut final_len = source_text.len();
    for (_, span, replacement) in &text_operations {
        final_len = final_len
            .checked_sub(span.end - span.start)
            .and_then(|length| length.checked_add(replacement.len()))
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
    for (index, span, replacement) in text_operations {
        output.push_str(
            source_text
                .get(cursor..span.start)
                .ok_or(Error::UnsupportedSource(
                    "staged text spans are not ordered UTF-8 boundaries",
                ))?,
        );
        let projected_start = output.len();
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

fn encoded_body_with_alignments(
    text: &str,
    alignments: &[Alignment],
    limits: crate::ParseLimits,
) -> Result<Vec<u8>, Error> {
    let extra = alignments
        .len()
        .checked_mul(4)
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
    let mut start = 0usize;
    for (newline, _) in text.match_indices('\n') {
        write_alignment(&mut output, alignments.get(paragraph).copied())?;
        let fragment = text.get(start..newline).ok_or(Error::UnsupportedSource(
            "paragraph boundary is not a UTF-8 text boundary",
        ))?;
        RtfWriter::new(&mut output)
            .write_text(fragment)
            .map_err(|error| Error::Write(error.to_string()))?;
        output.extend_from_slice(br"\par ");
        paragraph = paragraph.saturating_add(1);
        start = newline.saturating_add(1);
    }
    if start < text.len() || !text.ends_with('\n') {
        write_alignment(&mut output, alignments.get(paragraph).copied())?;
        let fragment = text.get(start..).ok_or(Error::UnsupportedSource(
            "paragraph boundary is not a UTF-8 text boundary",
        ))?;
        RtfWriter::new(&mut output)
            .write_text(fragment)
            .map_err(|error| Error::Write(error.to_string()))?;
        paragraph = paragraph.saturating_add(1);
    }
    if paragraph != alignments.len() {
        return Err(Error::StructuralPropertyConflict);
    }
    Ok(output)
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
        .filter_map(|(index, operation)| match operation {
            Operation::Text {
                span,
                before,
                after,
                structural: _,
            } if before != after => {
                projected_spans
                    .iter()
                    .find_map(|(projected_index, projected)| {
                        (*projected_index == index).then_some(Change::Text {
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
            Operation::Text { .. } | Operation::Alignment { .. } => None,
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
            _ => {
                return Err(Error::DurablePatch(
                    "unsupported operation vocabulary".to_string(),
                ));
            },
        }
    }
    edit.commit().map(Commit::into_snapshot)
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
