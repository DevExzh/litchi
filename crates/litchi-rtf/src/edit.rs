//! Bounded immutable RTF body-text transactions.
//!
//! This seam deliberately covers one ordinary body-text replacement. It does
//! not expose the retained mutable parser tree, and it refuses changed sources
//! containing opaque syntax or body structure whose positions would need a
//! richer editor to update safely.

use crate::{Document, RtfError, RtfWriter};
use bumpalo::Bump;
use std::fmt;
use std::ops::Range;

/// Immutable RTF snapshot used by the body-text transaction API.
pub type Snapshot = Document;

/// Failure from a body-text transaction or patch application.
#[derive(Debug, Clone)]
#[non_exhaustive]
pub enum Error {
    /// A source-specific body feature cannot be rewritten through this narrow seam.
    UnsupportedSource(&'static str),
    /// Only one bounded semantic operation is accepted by an edit.
    OperationAlreadyStaged,
    /// Replacement text exceeds the source snapshot's retained resource profile.
    InputTooLarge { observed: usize, limit: usize },
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
            Self::InputTooLarge { observed, limit } => write!(
                formatter,
                "replacement body text exceeds the source limit: observed {observed}, limit {limit}"
            ),
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

/// Detached, one-operation edit of an immutable snapshot.
pub struct Edit {
    source: Snapshot,
    replacement: Option<String>,
}

impl Edit {
    pub(crate) fn new(source: Snapshot) -> Self {
        Self {
            source,
            replacement: None,
        }
    }

    /// Stages replacement of the complete ordinary body story.
    ///
    /// A newline creates an RTF paragraph break. Changed documents with
    /// opaque syntax, tables, positioned body content, or mixed run/paragraph
    /// formatting are refused instead of losing or approximating that content.
    /// # Errors
    /// Returns an error if another operation is staged or the replacement
    /// exceeds the retained resource profile.
    pub fn replace_body_text(
        &mut self,
        replacement: impl Into<String>,
    ) -> Result<&mut Self, Error> {
        if self.replacement.is_some() {
            return Err(Error::OperationAlreadyStaged);
        }
        let text = replacement.into();
        let limit = self.source.limits().max_source_bytes();
        if text.len() > limit {
            return Err(Error::InputTooLarge {
                observed: text.len(),
                limit,
            });
        }
        self.replacement = Some(text);
        Ok(self)
    }

    /// Validates and publishes the candidate atomically.
    ///
    /// # Errors
    /// Returns an error when the source is outside this seam's supported
    /// closure or candidate validation fails.
    pub fn commit(self) -> Result<Commit, Error> {
        let Some(replacement) = self.replacement else {
            return Ok(Commit::new(self.source.clone(), self.source, false, 0));
        };
        if replacement == self.source.text() {
            return Ok(Commit::new(self.source.clone(), self.source, false, 1));
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
        let replacement_bytes = encoded_body_text(&replacement, self.source.limits())?;
        let bytes = splice_body(source_bytes, span, &replacement_bytes, self.source.limits())?;
        let snapshot = Snapshot::from_bytes_with_limits(&bytes, self.source.limits())?;
        if snapshot.text() != replacement {
            return Err(Error::UnsupportedSource(
                "candidate body text did not survive RTF validation",
            ));
        }
        Ok(Commit::new(self.source, snapshot, true, 1))
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
        .map_err(|_err| Error::Write("could not reserve replacement RTF bytes".to_string()))?;
    RtfWriter::new(&mut output)
        .write_text(text)
        .map_err(|error| Error::Write(error.to_string()))?;
    Ok(output)
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
        .map_err(|_err| Error::Write("could not reserve candidate RTF bytes".to_string()))?;
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

/// Result of an atomically validated body-text edit.
pub struct Commit {
    snapshot: Snapshot,
    patch: Patch,
    diagnostics: Diagnostics,
}

impl Commit {
    fn new(before: Snapshot, snapshot: Snapshot, changed: bool, operation_count: usize) -> Self {
        Self {
            patch: Patch {
                before,
                after: snapshot.clone(),
            },
            snapshot,
            diagnostics: Diagnostics {
                operation_count,
                changed,
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

/// In-memory exact-source-checked reversible RTF patch.
#[derive(Clone)]
pub struct Patch {
    before: Snapshot,
    after: Snapshot,
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
        }
    }
}
