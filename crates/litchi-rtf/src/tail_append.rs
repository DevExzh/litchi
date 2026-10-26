//! Preservation-first logical appends to an existing RTF document.
//!
//! This module is deliberately separate from [`crate::streaming`].  A
//! streaming writer owns a new forward-only document; this transaction owns an
//! immutable existing snapshot and can publish only a newly validated artifact
//! whose inserted bytes are placed immediately before the exact root closing
//! group.  The source is never normalized or mutated.

mod limits;
pub use limits::{EditOptions, SemanticLimits, SemanticLimitsBuilder};

use crate::Package;
use litchi_core::{
    ExecutionContext, ExecutionError, ReadAt, Reservation, Resource, SourceVersion, patch::BlobId,
};
use serde_json::{Map, Value};
use sha2::{Digest as _, Sha256};
use std::collections::BTreeSet;
use std::fmt;
use std::io::{self, Write};
use std::sync::Arc;

/// Stable resource dimension for a logical-tail transaction refusal.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum LimitKind {
    /// Admitted paragraphs.
    Paragraphs,
    /// Admitted runs.
    Runs,
    /// Admitted input bytes.
    InputBytes,
    /// Admitted inserted bytes.
    InsertedBytes,
    /// Admitted output bytes.
    OutputBytes,
    /// Admitted source bytes.
    SourceBytes,
    /// Admitted patch bytes.
    PatchBytes,
    /// Admitted patch offset.
    PatchOffset,
    /// Admitted after bytes.
    AfterBytes,
    /// Admitted before bytes.
    BeforeBytes,
    /// Admitted memory.
    MemoryBytes,
}

/// The only story currently accepted by the logical-tail splice.
///
/// Keeping the selector explicit makes the API selector-first and leaves room
/// for future, independently proved story tails without making a body append
/// look like a whole-document rewrite.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum TailSelector {
    /// The ordinary root document body story.
    Body,
}

/// A plain run supplied to a bounded tail append.
///
/// Runs carry no formatting or destination semantics.  Their boundaries are
/// retained in the authored payload as `\\plain` groups, so callers may submit
/// one paragraph as one run or as several adjacent runs without asking this
/// narrow API to infer formatting dependencies.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct PlainRun<'a> {
    text: &'a str,
}

impl<'a> PlainRun<'a> {
    /// Creates one borrowed plain run.
    #[must_use]
    pub const fn new(text: &'a str) -> Self {
        Self { text }
    }

    /// Borrow the run's source text.
    #[must_use]
    pub const fn text(self) -> &'a str {
        self.text
    }
}

/// One plain paragraph made of zero or more plain runs.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct PlainParagraph<'a> {
    runs: &'a [PlainRun<'a>],
}

impl<'a> PlainParagraph<'a> {
    /// Creates one borrowed paragraph from its ordered runs.
    #[must_use]
    pub const fn new(runs: &'a [PlainRun<'a>]) -> Self {
        Self { runs }
    }

    /// Borrow the paragraph's ordered plain runs.
    #[must_use]
    pub const fn runs(self) -> &'a [PlainRun<'a>] {
        self.runs
    }
}

/// Hard publication-window bounds for an existing-document logical-tail
/// append.
///
/// The semantic [`Commit`] path intentionally retains its complete
/// validated candidate.  A [`TailAppendPublicationPlan`] instead emits the
/// exact source prefix, bounded inserted span, and exact source suffix
/// directly to a sequential sink.  `max_window_bytes` bounds each source or
/// inserted window offered to the sink, while `max_write_bytes` is the hard
/// maximum passed to one `Write::write` call.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct TailAppendPublicationLimits {
    /// Maximum bytes retained in one publication window.
    pub max_window_bytes: usize,
    /// Maximum bytes offered to one sequential sink write.
    pub max_write_bytes: usize,
}

impl TailAppendPublicationLimits {
    /// Creates explicit finite publication-window bounds.
    #[must_use]
    pub const fn new(max_window_bytes: usize, max_write_bytes: usize) -> Self {
        Self {
            max_window_bytes,
            max_write_bytes,
        }
    }

    pub(crate) fn validate(self) -> Result<(), TailAppendPublicationError> {
        if self.max_window_bytes == 0 {
            return Err(TailAppendPublicationError::InvalidLimits(
                "publication window must be positive",
            ));
        }
        if self.max_write_bytes == 0 {
            return Err(TailAppendPublicationError::InvalidLimits(
                "publication write cap must be positive",
            ));
        }
        if self.max_write_bytes > self.max_window_bytes {
            return Err(TailAppendPublicationError::InvalidLimits(
                "publication write cap must not exceed the window",
            ));
        }
        Ok(())
    }
}

impl Default for TailAppendPublicationLimits {
    fn default() -> Self {
        Self::new(16 * 1024, 16 * 1024)
    }
}

/// Exact progress known after a non-atomic sequential publication failure.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum TailAppendOutputProgress {
    /// The sink accepted no bytes.
    Untouched,
    /// The sink accepted an exact prefix of the planned artifact.
    Prefix {
        /// Bytes definitely accepted by the sink.
        accepted: u64,
        /// Complete planned artifact length.
        expected: u64,
    },
    /// Every planned artifact byte was accepted but flush failed.
    CompleteUnflushed {
        /// Complete planned artifact length.
        bytes: u64,
    },
    /// Every planned artifact byte was accepted but a post-write proof failed.
    CompleteUnverified {
        /// Complete planned artifact length.
        bytes: u64,
    },
    /// The sink reported more bytes than were offered, so exact progress is
    /// unknowable beyond this point.
    Indeterminate {
        /// Bytes definitely accepted before the invalid report.
        accepted_before: u64,
    },
}

/// Failure from bounded publication of an existing-document logical tail.
#[derive(Debug)]
#[non_exhaustive]
pub enum TailAppendPublicationError {
    /// The publication plan could not be constructed or its source proof
    /// failed before output began.
    Plan(Error),
    /// The caller supplied an invalid publication window configuration.
    InvalidLimits(&'static str),
    /// The caller supplied a different immutable source snapshot.
    SourceVersionChanged {
        /// Version captured while planning.
        expected: SourceVersion,
        /// Version observed at publication.
        observed: SourceVersion,
    },
    /// The source bytes no longer match the exact planning fingerprint.
    SourceFingerprintChanged {
        /// Fingerprint captured while planning.
        expected: BlobId,
        /// Fingerprint observed at publication.
        observed: BlobId,
    },
    /// The source length changed while a source-backed plan was being
    /// validated or published.
    SourceLengthChanged {
        /// Length captured while planning.
        expected: u64,
        /// Length observed at the check.
        observed: u64,
    },
    /// The source-backed plan observed a different canonical SHA-256 digest.
    SourceDigestChanged {
        /// Digest captured while planning.
        expected: String,
        /// Digest observed at the check.
        observed: String,
    },
    /// A positional source failed while it was being read.
    ///
    /// The original [`io::Error`] is retained so callers can inspect its
    /// source chain and raw operating-system error.  A conforming
    /// [`Write::write`] implementation reports zero accepted bytes alongside
    /// `Err`; the `written` count therefore remains exact.  Only a sink that
    /// violates that contract by over-reporting bytes produces indeterminate
    /// progress.
    Source {
        /// Original source I/O error, including any raw OS error.
        error: io::Error,
        /// Bytes accepted before the source failure.
        written: u64,
    },
    /// A caller execution context rejected a check or resource reservation.
    Execution {
        /// Underlying cancellation or hierarchical budget failure.
        error: ExecutionError,
        /// Bytes accepted before the failure.
        written: u64,
    },
    /// A local publication bound was exceeded.
    LimitExceeded {
        /// Stable resource name.
        resource: &'static str,
        /// Observed value.
        observed: u64,
        /// Configured maximum.
        limit: u64,
        /// Bytes accepted before the failure.
        written: u64,
    },
    /// A sequential sink failed after accepting the reported bytes.
    Sink {
        /// Underlying I/O category.
        kind: io::ErrorKind,
        /// Stable sink diagnostic.
        message: String,
        /// Bytes accepted before the failure.
        written: u64,
    },
    /// A sink failure or post-write proof failure left a typed partial output.
    IncompleteOutput {
        /// Exact progress classification.
        progress: TailAppendOutputProgress,
        /// Underlying failure cause.
        source: Box<TailAppendPublicationError>,
    },
}

impl TailAppendPublicationError {
    /// Bytes definitely accepted before this failure.
    #[must_use]
    pub const fn written(&self) -> u64 {
        match self {
            Self::Plan(_) | Self::InvalidLimits(_) => 0,
            Self::SourceVersionChanged { .. }
            | Self::SourceFingerprintChanged { .. }
            | Self::SourceLengthChanged { .. }
            | Self::SourceDigestChanged { .. } => 0,
            Self::Source { written, .. } => *written,
            Self::Execution { written, .. }
            | Self::LimitExceeded { written, .. }
            | Self::Sink { written, .. } => *written,
            Self::IncompleteOutput { progress, .. } => match progress {
                TailAppendOutputProgress::Untouched => 0,
                TailAppendOutputProgress::Prefix { accepted, .. } => *accepted,
                TailAppendOutputProgress::CompleteUnflushed { bytes }
                | TailAppendOutputProgress::CompleteUnverified { bytes } => *bytes,
                TailAppendOutputProgress::Indeterminate { accepted_before } => *accepted_before,
            },
        }
    }

    /// Returns exact output progress when publication reached the sink.
    #[must_use]
    pub const fn progress(&self) -> Option<TailAppendOutputProgress> {
        match self {
            Self::IncompleteOutput { progress, .. } => Some(*progress),
            _ => None,
        }
    }
}

impl fmt::Display for TailAppendPublicationError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Plan(error) => write!(formatter, "RTF tail publication plan failed: {error}"),
            Self::InvalidLimits(reason) => {
                write!(formatter, "invalid RTF tail publication limits: {reason}")
            },
            Self::SourceVersionChanged { expected, observed } => write!(
                formatter,
                "RTF tail publication source version changed from {expected:?} to {observed:?}"
            ),
            Self::SourceFingerprintChanged { .. } => {
                formatter.write_str("RTF tail publication source fingerprint changed")
            },
            Self::SourceLengthChanged { expected, observed } => write!(
                formatter,
                "RTF source length changed from {expected} to {observed}"
            ),
            Self::SourceDigestChanged { .. } => {
                formatter.write_str("RTF source-backed publication source digest changed")
            },
            Self::Source { error, written } => write!(
                formatter,
                "RTF tail publication source failed ({}) after {written} bytes: {error}",
                error.kind()
            ),
            Self::Execution { error, written } => write!(
                formatter,
                "RTF tail publication execution failed after {written} bytes: {error}"
            ),
            Self::LimitExceeded {
                resource,
                observed,
                limit,
                written,
            } => write!(
                formatter,
                "RTF tail publication limit exceeded for {resource}: observed {observed}, limit {limit}, output {written}"
            ),
            Self::Sink {
                kind,
                message,
                written,
            } => write!(
                formatter,
                "RTF tail publication sink failed ({kind}) after {written} bytes: {message}"
            ),
            Self::IncompleteOutput { progress, source } => write!(
                formatter,
                "incomplete RTF tail publication output ({progress:?}): {source}"
            ),
        }
    }
}

impl std::error::Error for TailAppendPublicationError {
    fn source(&self) -> Option<&(dyn std::error::Error + 'static)> {
        match self {
            Self::Plan(error) => Some(error),
            Self::IncompleteOutput { source, .. } => Some(source),
            Self::InvalidLimits(_)
            | Self::SourceVersionChanged { .. }
            | Self::SourceFingerprintChanged { .. }
            | Self::SourceLengthChanged { .. }
            | Self::SourceDigestChanged { .. }
            | Self::Execution { .. }
            | Self::LimitExceeded { .. }
            | Self::Sink { .. } => None,
            Self::Source { error, .. } => Some(error),
        }
    }
}

impl From<Error> for TailAppendPublicationError {
    fn from(error: Error) -> Self {
        Self::Plan(error)
    }
}

/// Evidence returned after bounded direct publication succeeds.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct TailAppendPublicationReport {
    bytes: usize,
    source_bytes: usize,
    inserted_bytes: usize,
    source_version: SourceVersion,
    source_fingerprint: BlobId,
    writes: usize,
    largest_write: usize,
}

impl TailAppendPublicationReport {
    /// Complete output bytes accepted by the sink.
    #[must_use]
    pub const fn bytes(&self) -> usize {
        self.bytes
    }

    /// Exact source bytes emitted before insertion.
    #[must_use]
    pub const fn source_bytes(&self) -> usize {
        self.source_bytes
    }

    /// Exact inserted bytes emitted between source prefix and suffix.
    #[must_use]
    pub const fn inserted_bytes(&self) -> usize {
        self.inserted_bytes
    }

    /// Source version checked before and after publication.
    #[must_use]
    pub const fn source_version(&self) -> SourceVersion {
        self.source_version
    }

    /// Exact source fingerprint checked before and after publication.
    #[must_use]
    pub fn source_fingerprint(&self) -> &BlobId {
        &self.source_fingerprint
    }

    /// Number of sink write calls, excluding flush.
    #[must_use]
    pub const fn writes(&self) -> usize {
        self.writes
    }

    /// Largest byte slice offered to one sink write call.
    #[must_use]
    pub const fn largest_write(&self) -> usize {
        self.largest_write
    }
}

/// Failure from a bounded logical-tail transaction or its durable patch.
#[derive(Debug, Clone)]
#[non_exhaustive]
pub enum Error {
    /// A well-formed request cannot be carried through this edit route.
    Refused(Refusal),
    /// A submitted batch contains no paragraphs/runs where one is required.
    EmptyInput,
    /// A configured bound exceeds the supported addressable maximum.
    InvalidLimit {
        kind: LimitKind,
        value: usize,
        maximum: usize,
    },
    /// A finite owner count was exceeded.
    Limit {
        kind: LimitKind,
        observed: usize,
        maximum: usize,
    },
    /// A submitted text scalar is not plain-text-safe for this API.
    InvalidText(&'static str),
    /// A fallible destination reservation failed.
    AllocationFailed {
        resource: &'static str,
        requested: usize,
    },
    /// Candidate transport parsing or semantic readback failed.
    Rtf(crate::Error),
    /// A durable wire envelope is malformed or exceeds its caller bound.
    DurablePatch(&'static str),
    /// A patch was applied to a source with a different exact artifact digest.
    PatchConflict,
    /// A sequential sink failed after accepting this many bytes.
    Sink { kind: io::ErrorKind, written: usize },
}

/// Shared, content-free refusal details for this family's errors.
pub use litchi_core::edit::{Refusal, RefusalKind};

impl Error {
    pub(crate) const fn unsupported_source(reason: &'static str) -> Self {
        Self::Refused(Refusal::new(
            RefusalKind::UnsupportedContent,
            "tail_append",
            reason,
        ))
    }

    pub(crate) const fn unsupported_operation(reason: &'static str) -> Self {
        Self::Refused(Refusal::new(
            RefusalKind::UnsupportedOperation,
            "tail_append",
            reason,
        ))
    }
}

impl fmt::Display for Error {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Refused(refusal) => refusal.fmt(formatter),
            Self::EmptyInput => formatter.write_str("RTF logical-tail input must not be empty"),
            Self::InvalidLimit {
                kind,
                value,
                maximum,
            } => write!(
                formatter,
                "invalid RTF tail {kind:?} limit: {value}, maximum {maximum}"
            ),
            Self::Limit {
                kind,
                observed,
                maximum,
            } => write!(
                formatter,
                "RTF logical-tail {kind:?} limit exceeded: observed {observed}, maximum {maximum}"
            ),
            Self::InvalidText(reason) => {
                write!(formatter, "invalid RTF logical-tail text: {reason}")
            },
            Self::AllocationFailed {
                resource,
                requested,
            } => write!(
                formatter,
                "RTF logical-tail allocation failed for {resource}: requested {requested} bytes"
            ),
            Self::Rtf(error) => error.fmt(formatter),
            Self::DurablePatch(reason) => {
                write!(formatter, "invalid RTF logical-tail patch: {reason}")
            },
            Self::PatchConflict => formatter.write_str(
                "RTF logical-tail patch source does not match its exact expected artifact",
            ),
            Self::Sink { kind, written } => {
                write!(
                    formatter,
                    "RTF logical-tail sink failed ({kind}) after {written} bytes"
                )
            },
        }
    }
}

impl std::error::Error for Error {}

impl From<crate::Error> for Error {
    fn from(error: crate::Error) -> Self {
        Self::Rtf(error)
    }
}

#[derive(Debug, Clone)]
struct OwnedParagraph {
    runs: Vec<String>,
}

/// A detached, selector-first append transaction.
///
/// ```
/// use litchi_rtf::{Package, tail_append::{Edit, TailSelector}};
/// let source = Package::parse(r"{\rtf1 Before}")?;
/// let mut edit: Edit = source.edit_tail_append(TailSelector::Body);
/// edit.append_text_paragraphs(&["After"])?;
/// let commit = edit.commit()?;
/// assert!(commit.diagnostics().changed());
/// let (package, patch) = commit.into_parts();
/// assert!(!patch.is_noop());
/// assert_eq!(source.apply_tail_append(&patch)?.into_package().text(), package.text());
/// # Ok::<(), Box<dyn std::error::Error>>(())
/// ```
pub struct Edit {
    source: Package,
    selector: TailSelector,
    limits: SemanticLimits,
    paragraphs: Vec<OwnedParagraph>,
    run_count: usize,
    input_bytes: usize,
}

impl Edit {
    /// Starts a bounded body-tail transaction rooted at an immutable snapshot.
    #[must_use]
    pub fn new(source: &Package, selector: TailSelector) -> Self {
        Self {
            source: source.clone(),
            selector,
            limits: SemanticLimits::default(),
            paragraphs: Vec::new(),
            run_count: 0,
            input_bytes: 0,
        }
    }

    /// Starts a transaction with explicit append and wire bounds.
    #[must_use]
    pub fn new_with_options(
        source: &Package,
        selector: TailSelector,
        options: EditOptions,
    ) -> Self {
        let mut edit = Self::new(source, selector);
        edit.limits = options.semantic_limits_value();
        edit
    }

    /// The policy admitted for this edit.
    #[must_use]
    pub fn options(&self) -> EditOptions {
        EditOptions::new().semantic_limits(self.limits)
    }

    /// Immutable snapshot against which selectors and source bytes resolve.
    #[must_use]
    pub const fn source(&self) -> &Package {
        &self.source
    }

    /// Selector resolved by this transaction.
    #[must_use]
    pub const fn selector(&self) -> TailSelector {
        self.selector
    }

    /// Number of staged paragraphs.
    #[must_use]
    pub fn paragraph_count(&self) -> usize {
        self.paragraphs.len()
    }

    /// Number of staged runs.
    #[must_use]
    pub const fn run_count(&self) -> usize {
        self.run_count
    }

    /// UTF-8 bytes copied from the caller's submitted text.
    #[must_use]
    pub const fn input_bytes(&self) -> usize {
        self.input_bytes
    }

    /// Stages one or more borrowed plain paragraphs atomically.
    ///
    /// A paragraph may contain multiple plain runs.  This method performs all
    /// shape, text, owner, and allocation preflight before changing the edit.
    /// A second staging call is rejected so the transaction has one explicit
    /// bounded batch and no order-dependent merge behavior.
    pub fn append_paragraphs(
        &mut self,
        paragraphs: &[PlainParagraph<'_>],
    ) -> Result<&mut Self, Error> {
        if !self.paragraphs.is_empty() {
            return Err(Error::unsupported_operation(
                "the edit already has a staged batch",
            ));
        }
        if paragraphs.is_empty() {
            // An empty append is an exact semantic no-op.  Keeping the edit
            // unstaged lets commit share the original snapshot and bytes.
            return Ok(self);
        }
        let (observed_paragraphs, observed_runs, observed_input_bytes) =
            preflight_plain_paragraphs(paragraphs, self.limits)?;

        let mut owned = Vec::new();
        owned
            .try_reserve(observed_paragraphs)
            .map_err(|_error| Error::AllocationFailed {
                resource: "tail paragraphs",
                requested: observed_paragraphs.saturating_mul(size_of::<OwnedParagraph>()),
            })?;
        for paragraph in paragraphs {
            let mut runs = Vec::new();
            runs.try_reserve(paragraph.runs().len())
                .map_err(|_error| Error::AllocationFailed {
                    resource: "tail runs",
                    requested: paragraph.runs().len().saturating_mul(size_of::<String>()),
                })?;
            for run in paragraph.runs() {
                let mut text = String::new();
                text.try_reserve(run.text().len())
                    .map_err(|_error| Error::AllocationFailed {
                        resource: "tail text",
                        requested: run.text().len(),
                    })?;
                text.push_str(run.text());
                runs.push(text);
            }
            owned.push(OwnedParagraph { runs });
        }
        self.paragraphs = owned;
        self.run_count = observed_runs;
        self.input_bytes = observed_input_bytes;
        Ok(self)
    }

    /// Alias that makes the plain-text ownership boundary explicit at call
    /// Convenience staging method for one plain-text run per paragraph.
    pub fn append_text_paragraphs(&mut self, paragraphs: &[&str]) -> Result<&mut Self, Error> {
        if !self.paragraphs.is_empty() {
            return Err(Error::unsupported_operation(
                "the edit already has a staged batch",
            ));
        }
        if paragraphs.is_empty() {
            return Ok(self);
        }
        preflight_plain_text_paragraphs(paragraphs, self.limits)?;
        let mut runs = Vec::new();
        runs.try_reserve(paragraphs.len())
            .map_err(|_error| Error::AllocationFailed {
                resource: "tail input descriptors",
                requested: paragraphs.len().saturating_mul(size_of::<PlainRun<'_>>()),
            })?;
        let mut descriptors = Vec::new();
        descriptors
            .try_reserve(paragraphs.len())
            .map_err(|_error| Error::AllocationFailed {
                resource: "tail input descriptors",
                requested: paragraphs
                    .len()
                    .saturating_mul(size_of::<PlainParagraph<'_>>()),
            })?;
        for text in paragraphs {
            runs.push(PlainRun::new(text));
        }
        // The complete run vector was reserved before it was populated, so its
        // address is stable while descriptors borrow it for this call.
        for run in &runs {
            descriptors.push(PlainParagraph::new(std::slice::from_ref(run)));
        }
        self.append_paragraphs(&descriptors)
    }

    /// Convenience staging method for one paragraph made from plain runs.
    pub fn append_runs(&mut self, runs: &[PlainRun<'_>]) -> Result<&mut Self, Error> {
        if runs.is_empty() {
            return Ok(self);
        }
        let paragraph = PlainParagraph::new(runs);
        self.append_paragraphs(std::slice::from_ref(&paragraph))
    }

    /// Converts this staged transaction into a publication-only plan with an
    /// explicit hard source/write window.
    ///
    /// This constructor uses the transaction's finite [`SemanticLimits`]
    /// as its allocation boundary.  Call
    /// [`Self::publication_plan_with_context`] when planning must
    /// also reserve execution-budget memory before the encoded insertion is
    /// allocated.
    pub fn publication_plan(
        self,
        publication_limits: TailAppendPublicationLimits,
    ) -> Result<TailAppendPublicationPlan, Error> {
        self.publication_plan_inner(publication_limits, None)
            .map_err(|error| match error {
                TailAppendPublicationError::Plan(error) => error,
                TailAppendPublicationError::InvalidLimits(_)
                | TailAppendPublicationError::SourceVersionChanged { .. }
                | TailAppendPublicationError::SourceFingerprintChanged { .. }
                | TailAppendPublicationError::SourceLengthChanged { .. }
                | TailAppendPublicationError::SourceDigestChanged { .. }
                | TailAppendPublicationError::Source { .. }
                | TailAppendPublicationError::Execution { .. }
                | TailAppendPublicationError::LimitExceeded { .. }
                | TailAppendPublicationError::Sink { .. }
                | TailAppendPublicationError::IncompleteOutput { .. } => Error::unsupported_source(
                    "publication planning failed outside the source proof",
                ),
            })
    }

    /// Context-aware variant with explicit bounded source/write windows.
    pub fn publication_plan_with_context(
        self,
        publication_limits: TailAppendPublicationLimits,
        context: &ExecutionContext,
    ) -> Result<TailAppendPublicationPlan, TailAppendPublicationError> {
        publication_limits.validate()?;
        self.publication_plan_inner(publication_limits, Some(context))
    }

    fn publication_plan_inner(
        self,
        publication_limits: TailAppendPublicationLimits,
        planning_context: Option<&ExecutionContext>,
    ) -> Result<TailAppendPublicationPlan, TailAppendPublicationError> {
        if self.selector != TailSelector::Body {
            return Err(Error::Refused(Refusal::new(
                RefusalKind::UnsupportedTarget,
                "tail_append",
                "the selector is not supported by this edit family",
            ))
            .into());
        }
        let source_bytes = self
            .source
            .source_bytes()
            .ok_or(Error::unsupported_operation(
                "snapshot has no exact RTF source",
            ))?;
        let source_length = source_bytes.len();
        if source_length > self.limits.max_output_bytes {
            return Err(Error::Limit {
                kind: LimitKind::OutputBytes,
                observed: source_length,
                maximum: self.limits.max_output_bytes,
            }
            .into());
        }
        let source_fingerprint = BlobId::of(source_bytes);
        let source_version = self.source.source_version();

        let (root_close, inserted_len, input_bytes, paragraphs, runs, ends_with_par) =
            if self.paragraphs.is_empty() {
                // An empty append remains an exact identity transition, but the
                // publication-only path still proves that the source is eligible
                // for direct output before it touches the sink.  The established
                // semantic commit path retains its broader exact-no-op behavior.
                let proof = prove_splice_source(&self.source, self.selector)?;
                (source_length, 0, 0, 0, 0, proof.ends_with_par)
            } else {
                let proof = prove_splice_source(&self.source, self.selector)?;
                let inserted_len =
                    encoded_inserted_len(&self.source, &self.paragraphs, proof.ends_with_par)?;
                (
                    proof.root_close,
                    inserted_len,
                    self.input_bytes,
                    self.paragraphs.len(),
                    self.run_count,
                    proof.ends_with_par,
                )
            };
        if inserted_len > self.limits.max_inserted_bytes {
            return Err(Error::Limit {
                kind: LimitKind::InsertedBytes,
                observed: inserted_len,
                maximum: self.limits.max_inserted_bytes,
            }
            .into());
        }
        let output_length = source_length
            .checked_add(inserted_len)
            .ok_or(Error::Limit {
                kind: LimitKind::OutputBytes,
                observed: usize::MAX,
                maximum: self.limits.max_output_bytes,
            })?;
        if output_length > self.limits.max_output_bytes {
            return Err(Error::Limit {
                kind: LimitKind::OutputBytes,
                observed: output_length,
                maximum: self.limits.max_output_bytes,
            }
            .into());
        }
        if output_length > self.source.limits().max_input_bytes() {
            return Err(Error::Limit {
                kind: LimitKind::SourceBytes,
                observed: output_length,
                maximum: self.source.limits().max_input_bytes(),
            }
            .into());
        }

        let planning_memory = if let Some(context) = planning_context {
            let window = publication_limits.max_window_bytes.min(output_length);
            let amount = window.checked_add(inserted_len).ok_or({
                TailAppendPublicationError::Plan(Error::Limit {
                    kind: LimitKind::MemoryBytes,
                    observed: usize::MAX,
                    maximum: self.limits.max_output_bytes,
                })
            })?;
            context
                .check()
                .map_err(|error| publication_execution(error, 0))?;
            Some(Arc::new(
                context
                    .reserve(Resource::Memory, u64::try_from(amount).unwrap_or(u64::MAX))
                    .map_err(|error| publication_execution(error, 0))?,
            ))
        } else {
            None
        };

        let inserted = if inserted_len == 0 {
            Box::new([])
        } else {
            encode_inserted(&self.source, &self.paragraphs, ends_with_par, inserted_len)?
                .into_boxed_slice()
        };
        if inserted.len() != inserted_len {
            return Err(Error::unsupported_source(
                "encoded insertion length changed after bounded preflight",
            )
            .into());
        }
        Ok(TailAppendPublicationPlan {
            source: self.source,
            selector: self.selector,
            append_limits: self.limits,
            publication_limits,
            root_close,
            inserted,
            planning_memory,
            source_version,
            source_fingerprint,
            source_bytes: source_length,
            input_bytes,
            paragraphs,
            runs,
        })
    }

    /// Atomically validates and publishes the candidate artifact.
    pub fn commit(self) -> Result<Commit, Error> {
        if self.selector != TailSelector::Body {
            return Err(Error::Refused(Refusal::new(
                RefusalKind::UnsupportedTarget,
                "tail_append",
                "the selector is not supported by this edit family",
            )));
        }
        let operation_count = usize::from(!self.paragraphs.is_empty());
        if self.paragraphs.is_empty() {
            let source_bytes = self.source.source_bytes().map_or(0, <[u8]>::len);
            let patch = Patch::no_op(self.source.clone(), self.selector);
            return Ok(Commit {
                snapshot: self.source,
                patch,
                diagnostics: Diagnostics {
                    changed: false,
                    operation_count,
                    paragraphs: 0,
                    runs: 0,
                    input_bytes: 0,
                    source_bytes,
                    inserted_bytes: 0,
                    output_bytes: source_bytes,
                },
            });
        }
        let proof = prove_splice_source(&self.source, self.selector)?;
        let source_bytes = self
            .source
            .source_bytes()
            .ok_or(Error::unsupported_operation(
                "snapshot has no exact RTF source",
            ))?;
        let inserted_len =
            encoded_inserted_len(&self.source, &self.paragraphs, proof.ends_with_par)?;
        if inserted_len > self.limits.max_inserted_bytes {
            return Err(Error::Limit {
                kind: LimitKind::InsertedBytes,
                observed: inserted_len,
                maximum: self.limits.max_inserted_bytes,
            });
        }
        let output_len = source_bytes
            .len()
            .checked_add(inserted_len)
            .ok_or(Error::Limit {
                kind: LimitKind::OutputBytes,
                observed: usize::MAX,
                maximum: self.limits.max_output_bytes,
            })?;
        if output_len > self.limits.max_output_bytes {
            return Err(Error::Limit {
                kind: LimitKind::OutputBytes,
                observed: output_len,
                maximum: self.limits.max_output_bytes,
            });
        }
        if output_len > self.source.limits().max_input_bytes() {
            return Err(Error::Limit {
                kind: LimitKind::SourceBytes,
                observed: output_len,
                maximum: self.source.limits().max_input_bytes(),
            });
        }
        let inserted = encode_inserted(
            &self.source,
            &self.paragraphs,
            proof.ends_with_par,
            inserted_len,
        )?;
        let mut output = Vec::new();
        output
            .try_reserve(output_len)
            .map_err(|_error| Error::AllocationFailed {
                resource: "tail candidate",
                requested: output_len,
            })?;
        let prefix = source_bytes
            .get(..proof.root_close)
            .ok_or(Error::unsupported_source(
                "root-close proof is outside the source",
            ))?;
        let suffix = source_bytes
            .get(proof.root_close..)
            .ok_or(Error::unsupported_source(
                "root-close proof is outside the source",
            ))?;
        output.extend_from_slice(prefix);
        output.extend_from_slice(&inserted);
        output.extend_from_slice(suffix);

        let candidate = Package::from_bytes_with_options(&output, self.source.options())?;
        verify_candidate(
            &self.source,
            &candidate,
            &self.paragraphs,
            proof.ends_with_par,
        )?;
        let patch = Patch {
            before: self.source.clone(),
            after: candidate.clone(),
            selector: self.selector,
            root_close: proof.root_close,
            inserted: InsertedSpan {
                offset: proof.root_close,
                len: inserted_len,
            },
            direction: Direction::Append,
        };
        Ok(Commit {
            snapshot: candidate,
            patch,
            diagnostics: Diagnostics {
                changed: true,
                operation_count,
                paragraphs: self.paragraphs.len(),
                runs: self.run_count,
                input_bytes: self.input_bytes,
                source_bytes: source_bytes.len(),
                inserted_bytes: output_len.saturating_sub(source_bytes.len()),
                output_bytes: output_len,
            },
        })
    }
}

/// Fully validated, publication-only logical-tail plan.
///
/// The plan owns the immutable source handle and only the bounded encoded
/// insertion.  It never constructs or retains a complete target artifact;
/// [`Self::write_to`] emits source prefix, insertion, and source suffix in
/// bounded windows after rechecking source version and exact bytes.
#[derive(Debug, Clone)]
pub struct TailAppendPublicationPlan {
    source: Package,
    selector: TailSelector,
    append_limits: SemanticLimits,
    publication_limits: TailAppendPublicationLimits,
    root_close: usize,
    inserted: Box<[u8]>,
    planning_memory: Option<Arc<Reservation>>,
    source_version: SourceVersion,
    source_fingerprint: BlobId,
    source_bytes: usize,
    input_bytes: usize,
    paragraphs: usize,
    runs: usize,
}

impl TailAppendPublicationPlan {
    /// Immutable source snapshot retained by this plan.
    #[must_use]
    pub const fn source(&self) -> &Package {
        &self.source
    }

    /// Selector retained by this plan.
    #[must_use]
    pub const fn selector(&self) -> TailSelector {
        self.selector
    }

    /// Whether this plan is an exact source identity publication.
    #[must_use]
    pub const fn is_noop(&self) -> bool {
        self.inserted.is_empty() && self.root_close == self.source_bytes
    }

    /// Exact source byte length captured during planning.
    #[must_use]
    pub const fn source_bytes(&self) -> usize {
        self.source_bytes
    }

    /// Exact encoded insertion byte length.
    #[must_use]
    pub const fn inserted_bytes(&self) -> usize {
        self.inserted.len()
    }

    /// Exact target byte length emitted by this plan.
    #[must_use]
    pub const fn output_bytes(&self) -> usize {
        self.source_bytes.saturating_add(self.inserted.len())
    }

    /// UTF-8 bytes copied from caller-owned paragraph runs during staging.
    #[must_use]
    pub const fn input_bytes(&self) -> usize {
        self.input_bytes
    }

    /// Number of paragraphs represented by the bounded insertion.
    #[must_use]
    pub const fn paragraphs(&self) -> usize {
        self.paragraphs
    }

    /// Number of plain runs represented by the bounded insertion.
    #[must_use]
    pub const fn runs(&self) -> usize {
        self.runs
    }

    /// Exact root-close insertion offset, or the source length for a no-op.
    #[must_use]
    pub const fn root_close(&self) -> usize {
        self.root_close
    }

    /// Source version captured during planning.
    #[must_use]
    pub const fn source_version(&self) -> SourceVersion {
        self.source_version
    }

    /// Exact source fingerprint captured during planning.
    #[must_use]
    pub fn source_fingerprint(&self) -> &BlobId {
        &self.source_fingerprint
    }

    /// Publication window bounds retained by this plan.
    #[must_use]
    pub const fn publication_limits(&self) -> TailAppendPublicationLimits {
        self.publication_limits
    }

    /// Emits the plan against its retained immutable source.
    pub fn write_to<W: Write>(
        &self,
        sink: &mut W,
        context: &ExecutionContext,
    ) -> Result<TailAppendPublicationReport, TailAppendPublicationError> {
        self.write_to_source(&self.source, sink, context)
    }

    /// Alias for [`Self::write_to`].
    pub fn publish_to<W: Write>(
        &self,
        sink: &mut W,
        context: &ExecutionContext,
    ) -> Result<TailAppendPublicationReport, TailAppendPublicationError> {
        self.write_to(sink, context)
    }

    /// Alias for [`Self::write_to_source`] for callers that make the source
    /// authorization boundary explicit at the call site.
    pub fn publish_to_source<W: Write>(
        &self,
        source: &Package,
        sink: &mut W,
        context: &ExecutionContext,
    ) -> Result<TailAppendPublicationReport, TailAppendPublicationError> {
        self.write_to_source(source, sink, context)
    }

    /// Emits the plan against an explicitly supplied source snapshot.
    ///
    /// The supplied source must be the exact immutable snapshot captured by
    /// the plan.  A foreign or stale source is rejected before the first sink
    /// write, even when its bytes happen to be equal.
    pub fn write_to_source<W: Write>(
        &self,
        source: &Package,
        sink: &mut W,
        context: &ExecutionContext,
    ) -> Result<TailAppendPublicationReport, TailAppendPublicationError> {
        self.publication_limits.validate()?;
        if self.selector != TailSelector::Body {
            return Err(TailAppendPublicationError::Plan(Error::Refused(
                Refusal::new(
                    RefusalKind::UnsupportedTarget,
                    "tail_append",
                    "the selector is not supported by this edit family",
                ),
            )));
        }
        let source_bytes = source
            .source_bytes()
            .ok_or(TailAppendPublicationError::Plan(
                Error::unsupported_operation("snapshot has no exact RTF source"),
            ))?;
        let observed_version = source.source_version();
        if observed_version != self.source_version {
            return Err(TailAppendPublicationError::SourceVersionChanged {
                expected: self.source_version,
                observed: observed_version,
            });
        }
        let observed_fingerprint = BlobId::of(source_bytes);
        if source_bytes.len() != self.source_bytes
            || observed_fingerprint != self.source_fingerprint
        {
            return Err(TailAppendPublicationError::SourceFingerprintChanged {
                expected: self.source_fingerprint.clone(),
                observed: observed_fingerprint,
            });
        }
        if self.inserted.len() > self.append_limits.max_inserted_bytes {
            return Err(TailAppendPublicationError::LimitExceeded {
                resource: "inserted bytes",
                observed: u64::try_from(self.inserted.len()).unwrap_or(u64::MAX),
                limit: u64::try_from(self.append_limits.max_inserted_bytes).unwrap_or(u64::MAX),
                written: 0,
            });
        }
        let output_length = source_bytes.len().checked_add(self.inserted.len()).ok_or(
            TailAppendPublicationError::LimitExceeded {
                resource: "output bytes",
                observed: u64::MAX,
                limit: u64::try_from(self.append_limits.max_output_bytes).unwrap_or(u64::MAX),
                written: 0,
            },
        )?;
        if output_length != self.output_bytes() {
            return Err(TailAppendPublicationError::Plan(Error::unsupported_source(
                "publication output length no longer matches its proof",
            )));
        }
        if output_length > self.append_limits.max_output_bytes {
            return Err(TailAppendPublicationError::LimitExceeded {
                resource: "output bytes",
                observed: u64::try_from(output_length).unwrap_or(u64::MAX),
                limit: u64::try_from(self.append_limits.max_output_bytes).unwrap_or(u64::MAX),
                written: 0,
            });
        }
        if output_length > source.limits().max_input_bytes() {
            return Err(TailAppendPublicationError::LimitExceeded {
                resource: "source bytes",
                observed: u64::try_from(output_length).unwrap_or(u64::MAX),
                limit: u64::try_from(source.limits().max_input_bytes()).unwrap_or(u64::MAX),
                written: 0,
            });
        }
        if self.root_close > source_bytes.len() {
            return Err(TailAppendPublicationError::Plan(Error::unsupported_source(
                "publication root-close proof is outside the source",
            )));
        }
        let expected = u64::try_from(output_length).map_err(|_error| {
            TailAppendPublicationError::LimitExceeded {
                resource: "output bytes",
                observed: u64::MAX,
                limit: u64::try_from(self.append_limits.max_output_bytes).unwrap_or(u64::MAX),
                written: 0,
            }
        })?;
        context
            .check()
            .map_err(|error| publication_execution(error, 0))?;

        let window = self.publication_limits.max_window_bytes.min(output_length);
        let memory_amount = if self.planning_memory.is_some() {
            0
        } else {
            window.checked_add(self.inserted.len()).ok_or(
                TailAppendPublicationError::LimitExceeded {
                    resource: "memory",
                    observed: u64::MAX,
                    limit: u64::try_from(window).unwrap_or(u64::MAX),
                    written: 0,
                },
            )?
        };
        let mut memory = Some(
            context
                .reserve(
                    Resource::Memory,
                    u64::try_from(memory_amount).unwrap_or(u64::MAX),
                )
                .map_err(|error| publication_execution(error, 0))?,
        );
        let mut input = Some(
            context
                .reserve(Resource::InputBytes, expected)
                .map_err(|error| publication_execution(error, 0))?,
        );
        let mut output = Some(
            context
                .reserve(Resource::OutputBytes, expected)
                .map_err(|error| publication_execution(error, 0))?,
        );
        let mut work = Some(
            context
                .reserve(Resource::Work, expected)
                .map_err(|error| publication_execution(error, 0))?,
        );

        let root_close = if self.is_noop() {
            source_bytes.len()
        } else {
            self.root_close
        };
        let prefix = source_bytes
            .get(..root_close)
            .ok_or(TailAppendPublicationError::Plan(Error::unsupported_source(
                "publication prefix proof is outside the source",
            )))?;
        let suffix = source_bytes
            .get(root_close..)
            .ok_or(TailAppendPublicationError::Plan(Error::unsupported_source(
                "publication suffix proof is outside the source",
            )))?;
        let segments = [prefix, self.inserted.as_ref(), suffix];
        let chunk_limit = self
            .publication_limits
            .max_window_bytes
            .min(self.publication_limits.max_write_bytes);
        let mut accepted = 0usize;
        let mut writes = 0usize;
        let mut largest_write = 0usize;
        for segment in segments {
            if let Err(error) = write_publication_segment(
                sink,
                segment,
                chunk_limit,
                context,
                None,
                &mut accepted,
                &mut writes,
                &mut largest_write,
            ) {
                settle_publication_reservations(
                    &mut memory,
                    &mut input,
                    &mut output,
                    &mut work,
                    u64::try_from(accepted).unwrap_or(u64::MAX),
                );
                return Err(with_publication_progress(
                    error,
                    u64::try_from(accepted).unwrap_or(u64::MAX),
                    expected,
                ));
            }
        }
        if let Err(error) = sink.flush() {
            settle_publication_reservations(
                &mut memory,
                &mut input,
                &mut output,
                &mut work,
                u64::try_from(accepted).unwrap_or(u64::MAX),
            );
            return Err(TailAppendPublicationError::IncompleteOutput {
                progress: TailAppendOutputProgress::CompleteUnflushed { bytes: expected },
                source: Box::new(TailAppendPublicationError::Sink {
                    kind: error.kind(),
                    message: error.to_string(),
                    written: expected,
                }),
            });
        }
        if let Err(error) = context.check() {
            settle_publication_reservations(
                &mut memory,
                &mut input,
                &mut output,
                &mut work,
                u64::try_from(accepted).unwrap_or(u64::MAX),
            );
            return Err(TailAppendPublicationError::IncompleteOutput {
                progress: TailAppendOutputProgress::CompleteUnverified { bytes: expected },
                source: Box::new(publication_execution(error, expected)),
            });
        }
        let observed_version = source.source_version();
        if observed_version != self.source_version {
            settle_publication_reservations(
                &mut memory,
                &mut input,
                &mut output,
                &mut work,
                u64::try_from(accepted).unwrap_or(u64::MAX),
            );
            return Err(TailAppendPublicationError::IncompleteOutput {
                progress: TailAppendOutputProgress::CompleteUnverified { bytes: expected },
                source: Box::new(TailAppendPublicationError::SourceVersionChanged {
                    expected: self.source_version,
                    observed: observed_version,
                }),
            });
        }
        let observed_fingerprint = BlobId::of(source_bytes);
        if observed_fingerprint != self.source_fingerprint {
            settle_publication_reservations(
                &mut memory,
                &mut input,
                &mut output,
                &mut work,
                u64::try_from(accepted).unwrap_or(u64::MAX),
            );
            return Err(TailAppendPublicationError::IncompleteOutput {
                progress: TailAppendOutputProgress::CompleteUnverified { bytes: expected },
                source: Box::new(TailAppendPublicationError::SourceFingerprintChanged {
                    expected: self.source_fingerprint.clone(),
                    observed: observed_fingerprint,
                }),
            });
        }
        settle_publication_reservations(&mut memory, &mut input, &mut output, &mut work, expected);
        Ok(TailAppendPublicationReport {
            bytes: output_length,
            source_bytes: self.source_bytes,
            inserted_bytes: self.inserted.len(),
            source_version: self.source_version,
            source_fingerprint: self.source_fingerprint.clone(),
            writes,
            largest_write,
        })
    }
}

/// A compact capability proving that an exact RTF tail splice was authorized
/// by the existing parser/security checks.
///
/// The proof contains no parsed document or candidate artifact. It is issued
/// only by the constructors on this type, all of which call the same
/// conservative proof used by [`Edit`].
#[derive(Debug, Clone)]
pub(crate) struct TailAppendSourceProof {
    selector: TailSelector,
    source_version: SourceVersion,
    source_bytes: usize,
    source_digest: String,
    root_close: usize,
    ends_with_par: bool,
    body_nonempty: bool,
    source_limit: usize,
    capability: TailAppendSourceCapability,
}

#[derive(Debug, Clone, Copy)]
struct TailAppendSourceCapability;

impl TailAppendSourceProof {
    /// Issues a proof for `source`, binding the proof to that positional
    /// source's declared revision and exact bytes.
    ///
    /// This finite convenience scan does not claim cancellation or budget
    /// enforcement. Use [`Self::from_document_for_source_with_context`] when
    /// the digest scan must participate in an [`ExecutionContext`]. The
    /// provider must be immutable while the proof and any derived plan live,
    /// and its revision token must be monotonic and never reused (including
    /// after a mutate-and-restore ABA sequence). A version and digest bind
    /// content/revision, not object identity; equal-byte providers with the
    /// same declared version are intentionally equivalent.
    #[cfg_attr(
        not(test),
        allow(
            dead_code,
            reason = "the source-backed engine's full surface is exercised by its internal tests"
        )
    )]
    pub(crate) fn from_document_for_source(
        document: &Package,
        selector: TailSelector,
        source: &dyn ReadAt,
    ) -> Result<Self, TailAppendPublicationError> {
        Self::from_document_for_source_inner(document, selector, source, None)
    }

    /// Issues a source proof while charging the bounded digest scan to
    /// `context` and checking cancellation between positional reads.
    pub(crate) fn from_document_for_source_with_context(
        document: &Package,
        selector: TailSelector,
        source: &dyn ReadAt,
        context: &ExecutionContext,
    ) -> Result<Self, TailAppendPublicationError> {
        Self::from_document_for_source_inner(document, selector, source, Some(context))
    }

    fn from_document_for_source_inner(
        document: &Package,
        selector: TailSelector,
        source: &dyn ReadAt,
        scan_context: Option<&ExecutionContext>,
    ) -> Result<Self, TailAppendPublicationError> {
        let splice =
            prove_splice_source(document, selector).map_err(TailAppendPublicationError::Plan)?;
        let document_bytes = document
            .source_bytes()
            .ok_or(TailAppendPublicationError::Plan(
                Error::unsupported_operation("snapshot has no exact RTF source"),
            ))?;
        let expected_version = source
            .version()
            .map_err(|error| source_publication_io(error, 0))?;
        let observed_length = source
            .len()
            .map_err(|error| source_publication_io(error, 0))?;
        let expected_length = u64::try_from(document_bytes.len()).unwrap_or(u64::MAX);
        if observed_length != expected_length {
            return Err(TailAppendPublicationError::SourceLengthChanged {
                expected: expected_length,
                observed: observed_length,
            });
        }
        let scan_reservations = if let Some(context) = scan_context {
            let window = TailAppendPublicationLimits::default()
                .max_window_bytes
                .min(document_bytes.len());
            let memory = context
                .reserve(Resource::Memory, u64::try_from(window).unwrap_or(u64::MAX))
                .map_err(|error| publication_execution(error, 0))?;
            let input = context
                .reserve(Resource::InputBytes, expected_length)
                .map_err(|error| publication_execution(error, 0))?;
            let work = context
                .reserve(Resource::Work, expected_length)
                .map_err(|error| publication_execution(error, 0))?;
            Some((memory, input, work))
        } else {
            None
        };
        let mut source_scanned = 0usize;
        let mut inserted_scanned = 0usize;
        let digest_result = source_and_target_digest(
            source,
            document_bytes.len(),
            document_bytes.len(),
            &[],
            TailAppendPublicationLimits::default().max_window_bytes,
            scan_context,
            &mut source_scanned,
            &mut inserted_scanned,
        );
        let (observed_digest, _) = match digest_result {
            Ok(digests) => digests,
            Err(error) => {
                if let Some((memory, input, work)) = scan_reservations {
                    let _ = memory.commit(0);
                    let units = source_scan_units(source_scanned, inserted_scanned);
                    let _ = input.commit(units);
                    let _ = work.commit(units);
                }
                return Err(source_scan_publication_error(error, 0));
            },
        };
        if let Some((memory, input, work)) = scan_reservations {
            let _ = memory.commit(0);
            let units = source_scan_units(source_scanned, inserted_scanned);
            let _ = input.commit(units);
            let _ = work.commit(units);
        }
        let expected_digest = BlobId::of(document_bytes).as_hex();
        if observed_digest != expected_digest {
            return Err(TailAppendPublicationError::SourceDigestChanged {
                expected: expected_digest,
                observed: observed_digest,
            });
        }
        let final_version = source
            .version()
            .map_err(|error| source_publication_io(error, 0))?;
        if final_version != expected_version {
            return Err(TailAppendPublicationError::SourceVersionChanged {
                expected: expected_version,
                observed: final_version,
            });
        }
        let final_length = source
            .len()
            .map_err(|error| source_publication_io(error, 0))?;
        if final_length != expected_length {
            return Err(TailAppendPublicationError::SourceLengthChanged {
                expected: expected_length,
                observed: final_length,
            });
        }
        Ok(Self {
            selector,
            source_version: expected_version,
            source_bytes: document_bytes.len(),
            source_digest: expected_digest,
            root_close: splice.root_close,
            ends_with_par: splice.ends_with_par,
            body_nonempty: !document.body().is_empty(),
            source_limit: document.limits().max_input_bytes(),
            capability: TailAppendSourceCapability,
        })
    }

    /// Source revision token authorized by this proof.
    #[must_use]
    pub(crate) const fn source_version(&self) -> SourceVersion {
        self.source_version
    }

    /// Exact source length authorized by this proof.
    #[must_use]
    pub(crate) const fn source_bytes(&self) -> usize {
        self.source_bytes
    }

    /// Whether the existing body ends in a parameterless `\\par` control.
    #[cfg_attr(
        not(test),
        allow(
            dead_code,
            reason = "the source-backed engine's full surface is exercised by its internal tests"
        )
    )]
    #[must_use]
    pub(crate) const fn ends_with_par(&self) -> bool {
        self.ends_with_par
    }
}

/// A source-backed append transaction that never retains the parsed
/// [`Package`] or a complete source/target candidate.
pub(crate) struct TailAppendSourceEdit {
    source: Arc<dyn ReadAt>,
    proof: TailAppendSourceProof,
    selector: TailSelector,
    limits: SemanticLimits,
    paragraphs: Vec<OwnedParagraph>,
    run_count: usize,
    input_bytes: usize,
}

impl fmt::Debug for TailAppendSourceEdit {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("TailAppendSourceEdit")
            .field("proof", &self.proof)
            .field("selector", &self.selector)
            .field("limits", &self.limits)
            .field("paragraphs", &self.paragraphs.len())
            .field("run_count", &self.run_count)
            .field("input_bytes", &self.input_bytes)
            .finish()
    }
}

impl TailAppendSourceEdit {
    /// Starts a source-backed append from a previously issued proof.
    #[must_use]
    pub(crate) fn new(
        source: Arc<dyn ReadAt>,
        proof: TailAppendSourceProof,
        selector: TailSelector,
    ) -> Self {
        Self {
            source,
            proof,
            selector,
            limits: SemanticLimits::default(),
            paragraphs: Vec::new(),
            run_count: 0,
            input_bytes: 0,
        }
    }

    /// Starts a source-backed append with explicit finite append bounds.
    #[must_use]
    pub(crate) fn new_with_options(
        source: Arc<dyn ReadAt>,
        proof: TailAppendSourceProof,
        selector: TailSelector,
        options: EditOptions,
    ) -> Self {
        let mut edit = Self::new(source, proof, selector);
        edit.limits = options.semantic_limits_value();
        edit
    }

    /// The policy admitted for this edit.
    #[must_use]
    pub(crate) fn options(&self) -> EditOptions {
        EditOptions::new().semantic_limits(self.limits)
    }

    /// Number of staged paragraphs.
    #[must_use]
    pub(crate) fn paragraph_count(&self) -> usize {
        self.paragraphs.len()
    }

    /// Number of staged runs.
    #[must_use]
    pub(crate) const fn run_count(&self) -> usize {
        self.run_count
    }

    /// Stages one bounded batch of borrowed plain paragraphs.
    pub(crate) fn append_paragraphs(
        &mut self,
        paragraphs: &[PlainParagraph<'_>],
    ) -> Result<&mut Self, Error> {
        if !self.paragraphs.is_empty() {
            return Err(Error::unsupported_operation(
                "the edit already has a staged batch",
            ));
        }
        if paragraphs.is_empty() {
            return Ok(self);
        }
        let (owned, run_count, input_bytes) = stage_plain_paragraphs(paragraphs, self.limits)?;
        self.paragraphs = owned;
        self.run_count = run_count;
        self.input_bytes = input_bytes;
        Ok(self)
    }

    /// Convenience staging method for one plain run per paragraph.
    pub(crate) fn append_text_paragraphs(
        &mut self,
        paragraphs: &[&str],
    ) -> Result<&mut Self, Error> {
        if !self.paragraphs.is_empty() {
            return Err(Error::unsupported_operation(
                "the edit already has a staged batch",
            ));
        }
        if paragraphs.is_empty() {
            return Ok(self);
        }
        preflight_plain_text_paragraphs(paragraphs, self.limits)?;
        let mut runs = Vec::new();
        runs.try_reserve(paragraphs.len())
            .map_err(|_error| Error::AllocationFailed {
                resource: "tail input descriptors",
                requested: paragraphs.len().saturating_mul(size_of::<PlainRun<'_>>()),
            })?;
        let mut descriptors = Vec::new();
        descriptors
            .try_reserve(paragraphs.len())
            .map_err(|_error| Error::AllocationFailed {
                resource: "tail input descriptors",
                requested: paragraphs
                    .len()
                    .saturating_mul(size_of::<PlainParagraph<'_>>()),
            })?;
        for text in paragraphs {
            runs.push(PlainRun::new(text));
        }
        for run in &runs {
            descriptors.push(PlainParagraph::new(std::slice::from_ref(run)));
        }
        self.append_paragraphs(&descriptors)
    }

    /// Convenience staging method for one paragraph made from plain runs.
    pub(crate) fn append_runs(&mut self, runs: &[PlainRun<'_>]) -> Result<&mut Self, Error> {
        if runs.is_empty() {
            return Ok(self);
        }
        let paragraph = PlainParagraph::new(runs);
        self.append_paragraphs(std::slice::from_ref(&paragraph))
    }

    /// Builds a source-backed publication plan with explicit source windows.
    #[cfg_attr(
        not(test),
        allow(
            dead_code,
            reason = "the source-backed engine's full surface is exercised by its internal tests"
        )
    )]
    pub(crate) fn publication_plan(
        self,
        publication_limits: TailAppendPublicationLimits,
    ) -> Result<TailAppendSourcePublicationPlan, TailAppendPublicationError> {
        self.publication_plan_inner(publication_limits, None)
    }

    /// Context-aware source-backed planning with explicit source windows.
    pub(crate) fn publication_plan_with_context(
        self,
        publication_limits: TailAppendPublicationLimits,
        context: &ExecutionContext,
    ) -> Result<TailAppendSourcePublicationPlan, TailAppendPublicationError> {
        publication_limits.validate()?;
        self.publication_plan_inner(publication_limits, Some(context))
    }

    fn publication_plan_inner(
        self,
        publication_limits: TailAppendPublicationLimits,
        planning_context: Option<&ExecutionContext>,
    ) -> Result<TailAppendSourcePublicationPlan, TailAppendPublicationError> {
        publication_limits.validate()?;
        let _capability = self.proof.capability;
        if self.selector != TailSelector::Body || self.proof.selector != TailSelector::Body {
            return Err(TailAppendPublicationError::Plan(Error::Refused(
                Refusal::new(
                    RefusalKind::UnsupportedTarget,
                    "tail_append",
                    "the selector is not supported by this edit family",
                ),
            )));
        }
        if self.selector != self.proof.selector {
            return Err(TailAppendPublicationError::Plan(Error::Refused(
                Refusal::new(
                    RefusalKind::UnsupportedTarget,
                    "tail_append",
                    "the selector is not supported by this edit family",
                ),
            )));
        }
        let source_version = self
            .source
            .version()
            .map_err(|error| source_publication_io(error, 0))?;
        if source_version != self.proof.source_version {
            return Err(TailAppendPublicationError::SourceVersionChanged {
                expected: self.proof.source_version,
                observed: source_version,
            });
        }
        let source_length = self
            .source
            .len()
            .map_err(|error| source_publication_io(error, 0))?;
        let expected_length = u64::try_from(self.proof.source_bytes).unwrap_or(u64::MAX);
        if source_length != expected_length {
            return Err(TailAppendPublicationError::SourceLengthChanged {
                expected: expected_length,
                observed: source_length,
            });
        }
        let source_length_usize = usize::try_from(source_length).map_err(|_error| {
            TailAppendPublicationError::Plan(Error::Limit {
                kind: LimitKind::SourceBytes,
                observed: usize::MAX,
                maximum: self.proof.source_limit,
            })
        })?;
        let root_close = if self.paragraphs.is_empty() {
            source_length_usize
        } else {
            self.proof.root_close
        };
        if root_close > source_length_usize {
            return Err(TailAppendPublicationError::Plan(Error::unsupported_source(
                "source-backed root-close proof is outside the source",
            )));
        }
        let inserted_len = if self.paragraphs.is_empty() {
            0
        } else {
            encoded_inserted_len_for_body(
                self.proof.body_nonempty,
                &self.paragraphs,
                self.proof.ends_with_par,
            )?
        };
        if inserted_len > self.limits.max_inserted_bytes {
            return Err(TailAppendPublicationError::Plan(Error::Limit {
                kind: LimitKind::InsertedBytes,
                observed: inserted_len,
                maximum: self.limits.max_inserted_bytes,
            }));
        }
        let output_length = source_length_usize.checked_add(inserted_len).ok_or(
            TailAppendPublicationError::Plan(Error::Limit {
                kind: LimitKind::OutputBytes,
                observed: usize::MAX,
                maximum: self.limits.max_output_bytes,
            }),
        )?;
        if output_length > self.limits.max_output_bytes {
            return Err(TailAppendPublicationError::Plan(Error::Limit {
                kind: LimitKind::OutputBytes,
                observed: output_length,
                maximum: self.limits.max_output_bytes,
            }));
        }
        if output_length > self.proof.source_limit {
            return Err(TailAppendPublicationError::Plan(Error::Limit {
                kind: LimitKind::SourceBytes,
                observed: output_length,
                maximum: self.proof.source_limit,
            }));
        }
        let mut planning_input = None;
        let mut planning_work = None;
        let mut planning_memory = if let Some(context) = planning_context {
            context
                .check()
                .map_err(|error| publication_execution(error, 0))?;
            let amount = publication_limits
                .max_window_bytes
                .min(source_length_usize)
                .checked_add(inserted_len)
                .ok_or(TailAppendPublicationError::Plan(Error::Limit {
                    kind: LimitKind::MemoryBytes,
                    observed: usize::MAX,
                    maximum: self.limits.max_output_bytes,
                }))?;
            let scan_amount = source_length_usize.checked_add(inserted_len).ok_or(
                TailAppendPublicationError::Plan(Error::Limit {
                    kind: LimitKind::InputBytes,
                    observed: usize::MAX,
                    maximum: self.limits.max_input_bytes,
                }),
            )?;
            let memory = context
                .reserve(Resource::Memory, u64::try_from(amount).unwrap_or(u64::MAX))
                .map_err(|error| publication_execution(error, 0))?;
            planning_input = Some(
                context
                    .reserve(
                        Resource::InputBytes,
                        u64::try_from(scan_amount).unwrap_or(u64::MAX),
                    )
                    .map_err(|error| publication_execution(error, 0))?,
            );
            planning_work = Some(
                context
                    .reserve(
                        Resource::Work,
                        u64::try_from(scan_amount).unwrap_or(u64::MAX),
                    )
                    .map_err(|error| publication_execution(error, 0))?,
            );
            Some(memory)
        } else {
            None
        };
        let inserted = if inserted_len == 0 {
            Box::new([])
        } else {
            encode_inserted_for_body(
                self.proof.body_nonempty,
                &self.paragraphs,
                self.proof.ends_with_par,
                inserted_len,
            )?
            .into_boxed_slice()
        };
        if inserted.len() != inserted_len {
            return Err(TailAppendPublicationError::Plan(Error::unsupported_source(
                "source-backed encoded insertion changed after preflight",
            )));
        }
        let mut source_scanned = 0usize;
        let mut inserted_scanned = 0usize;
        let digest_result = source_and_target_digest(
            self.source.as_ref(),
            source_length_usize,
            root_close,
            &inserted,
            publication_limits.max_window_bytes,
            planning_context,
            &mut source_scanned,
            &mut inserted_scanned,
        );
        let (source_digest, target_digest) = match digest_result {
            Ok(digests) => digests,
            Err(error) => {
                if let Some(reservation) = planning_memory.take() {
                    let _ = reservation.commit(0);
                }
                if let Some(reservation) = planning_input.take() {
                    let _ = reservation.commit(source_scan_units(source_scanned, inserted_scanned));
                }
                if let Some(reservation) = planning_work.take() {
                    let _ = reservation.commit(source_scan_units(source_scanned, inserted_scanned));
                }
                return Err(source_scan_publication_error(error, 0));
            },
        };
        if let Some(reservation) = planning_input.take() {
            let _ = reservation.commit(
                u64::try_from(source_scanned.saturating_add(inserted_scanned)).unwrap_or(u64::MAX),
            );
        }
        if let Some(reservation) = planning_work.take() {
            let _ = reservation.commit(
                u64::try_from(source_scanned.saturating_add(inserted_scanned)).unwrap_or(u64::MAX),
            );
        }
        if source_digest != self.proof.source_digest {
            return Err(TailAppendPublicationError::SourceDigestChanged {
                expected: self.proof.source_digest.clone(),
                observed: source_digest,
            });
        }
        let final_version = self
            .source
            .version()
            .map_err(|error| source_publication_io(error, 0))?;
        if final_version != source_version {
            return Err(TailAppendPublicationError::SourceVersionChanged {
                expected: source_version,
                observed: final_version,
            });
        }
        let final_length = self
            .source
            .len()
            .map_err(|error| source_publication_io(error, 0))?;
        if final_length != source_length {
            return Err(TailAppendPublicationError::SourceLengthChanged {
                expected: source_length,
                observed: final_length,
            });
        }
        if let Some(reservation) = planning_memory.take() {
            let _ = reservation.commit(0);
        }
        let source_digest = self.proof.source_digest.clone();
        Ok(TailAppendSourcePublicationPlan {
            source: self.source,
            proof: self.proof,
            selector: self.selector,
            append_limits: self.limits,
            publication_limits,
            root_close,
            inserted,
            direction: Direction::Append,
            source_digest,
            target_digest,
            source_bytes: source_length_usize,
            input_bytes: self.input_bytes,
            paragraphs: self.paragraphs.len(),
            runs: self.run_count,
        })
    }
}

/// Fully validated source-backed publication plan.
///
/// Its retained state is only an `Arc<dyn ReadAt>`, private proof metadata,
/// bounded insertion bytes, and source/target digests.  The source-window plus
/// insertion plan bound is `max_window_bytes + inserted_bytes`; the initial
/// parser's document retention is outside this publication plan's bound.
///
/// Plans are intentionally not [`Clone`]. Each publication reserves its own
/// source window, so concurrent or repeated calls are independently bounded
/// by the caller's execution budget.
pub(crate) struct TailAppendSourcePublicationPlan {
    source: Arc<dyn ReadAt>,
    proof: TailAppendSourceProof,
    selector: TailSelector,
    append_limits: SemanticLimits,
    publication_limits: TailAppendPublicationLimits,
    root_close: usize,
    inserted: Box<[u8]>,
    direction: Direction,
    source_digest: String,
    target_digest: String,
    source_bytes: usize,
    input_bytes: usize,
    paragraphs: usize,
    runs: usize,
}

impl fmt::Debug for TailAppendSourcePublicationPlan {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("TailAppendSourcePublicationPlan")
            .field("proof", &self.proof)
            .field("selector", &self.selector)
            .field("append_limits", &self.append_limits)
            .field("publication_limits", &self.publication_limits)
            .field("root_close", &self.root_close)
            .field("inserted_bytes", &self.inserted.len())
            .field("direction", &self.direction)
            .field("source_digest", &self.source_digest)
            .field("target_digest", &self.target_digest)
            .field("source_bytes", &self.source_bytes)
            .field("input_bytes", &self.input_bytes)
            .field("paragraphs", &self.paragraphs)
            .field("runs", &self.runs)
            .finish()
    }
}

impl TailAppendSourcePublicationPlan {
    /// Whether this plan is an exact source identity publication.
    #[must_use]
    pub(crate) const fn is_noop(&self) -> bool {
        match self.direction {
            Direction::Append => self.inserted.is_empty() && self.root_close == self.source_bytes,
            Direction::Remove => false,
        }
    }

    /// Exact source byte length captured during planning.
    #[must_use]
    pub(crate) const fn source_bytes(&self) -> usize {
        self.source_bytes
    }

    /// Exact inserted byte length.
    #[must_use]
    pub(crate) const fn inserted_bytes(&self) -> usize {
        self.inserted.len()
    }

    /// Exact output byte length.
    #[must_use]
    pub(crate) const fn output_bytes(&self) -> usize {
        match self.direction {
            Direction::Append => self.source_bytes.saturating_add(self.inserted.len()),
            Direction::Remove => self.source_bytes.saturating_sub(self.inserted.len()),
        }
    }

    /// UTF-8 bytes copied while staging.
    #[must_use]
    pub(crate) const fn input_bytes(&self) -> usize {
        self.input_bytes
    }

    /// Number of staged paragraphs.
    #[must_use]
    pub(crate) const fn paragraphs(&self) -> usize {
        self.paragraphs
    }

    /// Number of staged runs.
    #[must_use]
    pub(crate) const fn runs(&self) -> usize {
        self.runs
    }

    /// Canonical source digest captured during planning.
    #[must_use]
    pub(crate) fn source_digest(&self) -> &str {
        &self.source_digest
    }

    /// Canonical target digest computed during planning.
    #[must_use]
    pub(crate) fn target_digest(&self) -> &str {
        &self.target_digest
    }

    /// Publication window bounds retained by this plan.
    #[cfg_attr(
        not(test),
        allow(
            dead_code,
            reason = "the source-backed engine's full surface is exercised by its internal tests"
        )
    )]
    #[must_use]
    pub(crate) const fn publication_limits(&self) -> TailAppendPublicationLimits {
        self.publication_limits
    }

    /// Rebind this validated append plan to an independently opened equal
    /// source. Inverse replay removes the exact inserted span and swaps the
    /// authenticated source/target digests; no complete artifact is retained.
    pub(crate) fn rebase(
        &self,
        source: Arc<dyn ReadAt>,
        proof: TailAppendSourceProof,
        publication_limits: TailAppendPublicationLimits,
        inverse: bool,
    ) -> Result<Self, TailAppendPublicationError> {
        publication_limits.validate()?;
        let source_bytes = if inverse {
            self.output_bytes()
        } else {
            self.source_bytes
        };
        if proof.source_bytes != source_bytes {
            return Err(TailAppendPublicationError::SourceLengthChanged {
                expected: u64::try_from(source_bytes).unwrap_or(u64::MAX),
                observed: u64::try_from(proof.source_bytes).unwrap_or(u64::MAX),
            });
        }
        let source_digest = if inverse {
            &self.target_digest
        } else {
            &self.source_digest
        };
        let target_digest = if inverse {
            &self.source_digest
        } else {
            &self.target_digest
        };
        if proof.source_digest.as_str() != source_digest {
            return Err(TailAppendPublicationError::SourceDigestChanged {
                expected: source_digest.clone(),
                observed: proof.source_digest.clone(),
            });
        }
        let removed_end = self.root_close.checked_add(self.inserted.len()).ok_or(
            TailAppendPublicationError::Plan(Error::unsupported_source(
                "source-backed replay span overflows",
            )),
        )?;
        if self.root_close > source_bytes || (inverse && removed_end > source_bytes) {
            return Err(TailAppendPublicationError::Plan(Error::unsupported_source(
                "source-backed replay span is outside the recipient source",
            )));
        }
        Ok(Self {
            source,
            proof,
            selector: self.selector,
            append_limits: self.append_limits,
            publication_limits,
            root_close: self.root_close,
            inserted: self.inserted.clone(),
            direction: if inverse {
                Direction::Remove
            } else {
                self.direction
            },
            source_digest: source_digest.clone(),
            target_digest: target_digest.clone(),
            source_bytes,
            input_bytes: self.input_bytes,
            paragraphs: self.paragraphs,
            runs: self.runs,
        })
    }

    /// Emits against the plan's retained positional source.
    ///
    /// This method proves bytes and source lineage only; callers that need
    /// semantic readback should explicitly reopen the accepted output.
    pub(crate) fn write_to<W: Write>(
        &self,
        sink: &mut W,
        context: &ExecutionContext,
    ) -> Result<crate::source_backed::Report, TailAppendPublicationError> {
        self.write_to_source(self.source.as_ref(), sink, context)
    }

    /// Emits against an explicit positional source. A source with a different
    /// declared version, length, or digest is rejected before the first sink
    /// write. `SourceVersion` plus the digest is a content/revision binding,
    /// not an object-identity check: an equal-byte provider with the same
    /// declared version is intentionally indistinguishable. The provider
    /// must remain immutable for the plan's lifetime and use a monotonic,
    /// never-reused revision token, including across mutate-and-restore ABA
    /// changes.
    ///
    /// Sink progress follows the standard [`Write::write`] contract: when a
    /// sink returns `Err`, it has accepted zero bytes for that call. A sink
    /// that reports more bytes than offered violates that contract and is the
    /// only case classified as indeterminate progress.
    ///
    /// The execution context charges `InputBytes` and `Work` for the exact
    /// cumulative bounded passes: source/target preflight, emission, and the
    /// final source/target verification.
    pub(crate) fn write_to_source<W: Write>(
        &self,
        source: &dyn ReadAt,
        sink: &mut W,
        context: &ExecutionContext,
    ) -> Result<crate::source_backed::Report, TailAppendPublicationError> {
        self.publication_limits.validate()?;
        if self.selector != TailSelector::Body || self.proof.selector != TailSelector::Body {
            return Err(TailAppendPublicationError::Plan(Error::Refused(
                Refusal::new(
                    RefusalKind::UnsupportedTarget,
                    "tail_append",
                    "the selector is not supported by this edit family",
                ),
            )));
        }
        let expected = u64::try_from(self.source_bytes).unwrap_or(u64::MAX);
        let observed_version = source
            .version()
            .map_err(|error| source_publication_io(error, 0))?;
        if observed_version != self.proof.source_version {
            return Err(TailAppendPublicationError::SourceVersionChanged {
                expected: self.proof.source_version,
                observed: observed_version,
            });
        }
        let observed_length = source
            .len()
            .map_err(|error| source_publication_io(error, 0))?;
        if observed_length != expected {
            return Err(TailAppendPublicationError::SourceLengthChanged {
                expected,
                observed: observed_length,
            });
        }
        let source_length = usize::try_from(observed_length).map_err(|_error| {
            TailAppendPublicationError::LimitExceeded {
                resource: "source bytes",
                observed: u64::MAX,
                limit: u64::try_from(self.proof.source_limit).unwrap_or(u64::MAX),
                written: 0,
            }
        })?;
        let removed_end = self.root_close.checked_add(self.inserted.len()).ok_or(
            TailAppendPublicationError::Plan(Error::unsupported_source(
                "source-backed publication span overflows",
            )),
        )?;
        if source_length != self.source_bytes
            || self.root_close > source_length
            || (self.direction == Direction::Remove && removed_end > source_length)
        {
            return Err(TailAppendPublicationError::Plan(Error::unsupported_source(
                "source-backed publication bounds no longer match the proof",
            )));
        }
        let output_length = match self.direction {
            Direction::Append => source_length.checked_add(self.inserted.len()),
            Direction::Remove => source_length.checked_sub(self.inserted.len()),
        }
        .ok_or(TailAppendPublicationError::LimitExceeded {
            resource: "output bytes",
            observed: u64::MAX,
            limit: u64::try_from(self.append_limits.max_output_bytes).unwrap_or(u64::MAX),
            written: 0,
        })?;
        let expected_output = u64::try_from(output_length).map_err(|_error| {
            TailAppendPublicationError::LimitExceeded {
                resource: "output bytes",
                observed: u64::MAX,
                limit: u64::try_from(self.append_limits.max_output_bytes).unwrap_or(u64::MAX),
                written: 0,
            }
        })?;
        if output_length != self.output_bytes() {
            return Err(TailAppendPublicationError::Plan(Error::unsupported_source(
                "source-backed output length no longer matches its plan",
            )));
        }
        if output_length > self.append_limits.max_output_bytes {
            return Err(TailAppendPublicationError::LimitExceeded {
                resource: "output bytes",
                observed: expected_output,
                limit: u64::try_from(self.append_limits.max_output_bytes).unwrap_or(u64::MAX),
                written: 0,
            });
        }
        if output_length > self.proof.source_limit {
            return Err(TailAppendPublicationError::LimitExceeded {
                resource: "source bytes",
                observed: expected_output,
                limit: u64::try_from(self.proof.source_limit).unwrap_or(u64::MAX),
                written: 0,
            });
        }
        let chunk_limit = self
            .publication_limits
            .max_window_bytes
            .min(self.publication_limits.max_write_bytes);
        let memory_amount = self
            .publication_limits
            .max_window_bytes
            .min(source_length)
            .checked_add(self.inserted.len())
            .ok_or(TailAppendPublicationError::LimitExceeded {
                resource: "memory",
                observed: u64::MAX,
                limit: u64::try_from(self.publication_limits.max_window_bytes).unwrap_or(u64::MAX),
                written: 0,
            })?;
        let expected_work =
            expected_output
                .checked_mul(3)
                .ok_or(TailAppendPublicationError::LimitExceeded {
                    resource: "work",
                    observed: u64::MAX,
                    limit: u64::try_from(self.append_limits.max_output_bytes).unwrap_or(u64::MAX),
                    written: 0,
                })?;
        let mut memory = Some(
            context
                .reserve(
                    Resource::Memory,
                    u64::try_from(memory_amount).unwrap_or(u64::MAX),
                )
                .map_err(|error| publication_execution(error, 0))?,
        );
        let mut input = Some(
            context
                .reserve(Resource::InputBytes, expected_work)
                .map_err(|error| publication_execution(error, 0))?,
        );
        let mut output = Some(
            context
                .reserve(Resource::OutputBytes, expected_output)
                .map_err(|error| publication_execution(error, 0))?,
        );
        let mut work = Some(
            context
                .reserve(Resource::Work, expected_work)
                .map_err(|error| publication_execution(error, 0))?,
        );
        context
            .check()
            .map_err(|error| publication_execution(error, 0))?;
        let mut source_scanned = 0usize;
        let mut inserted_scanned = 0usize;
        let preflight = source_and_target_digest_with_direction(
            source,
            source_length,
            self.root_close,
            &self.inserted,
            self.direction,
            self.publication_limits.max_window_bytes,
            Some(context),
            &mut source_scanned,
            &mut inserted_scanned,
        );
        let (observed_digest, observed_target) = match preflight {
            Ok(digests) => digests,
            Err(error) => {
                settle_source_publication_reservations(
                    &mut memory,
                    &mut input,
                    &mut output,
                    &mut work,
                    source_scan_units(source_scanned, inserted_scanned),
                    0,
                );
                return Err(source_scan_publication_error(error, 0));
            },
        };
        if observed_digest != self.source_digest {
            settle_source_publication_reservations(
                &mut memory,
                &mut input,
                &mut output,
                &mut work,
                source_scan_units(source_scanned, inserted_scanned),
                0,
            );
            return Err(TailAppendPublicationError::SourceDigestChanged {
                expected: self.source_digest.clone(),
                observed: observed_digest,
            });
        }
        if observed_target != self.target_digest {
            settle_source_publication_reservations(
                &mut memory,
                &mut input,
                &mut output,
                &mut work,
                source_scan_units(source_scanned, inserted_scanned),
                0,
            );
            return Err(TailAppendPublicationError::SourceDigestChanged {
                expected: self.target_digest.clone(),
                observed: observed_target,
            });
        }
        let stable_version = source
            .version()
            .map_err(|error| source_publication_io(error, 0));
        let stable_version = match stable_version {
            Ok(version) => version,
            Err(error) => {
                settle_source_publication_reservations(
                    &mut memory,
                    &mut input,
                    &mut output,
                    &mut work,
                    source_scan_units(source_scanned, inserted_scanned),
                    0,
                );
                return Err(error);
            },
        };
        if stable_version != self.proof.source_version {
            settle_source_publication_reservations(
                &mut memory,
                &mut input,
                &mut output,
                &mut work,
                source_scan_units(source_scanned, inserted_scanned),
                0,
            );
            return Err(TailAppendPublicationError::SourceVersionChanged {
                expected: self.proof.source_version,
                observed: stable_version,
            });
        }
        let stable_length = source
            .len()
            .map_err(|error| source_publication_io(error, 0));
        let stable_length = match stable_length {
            Ok(length) => length,
            Err(error) => {
                settle_source_publication_reservations(
                    &mut memory,
                    &mut input,
                    &mut output,
                    &mut work,
                    source_scan_units(source_scanned, inserted_scanned),
                    0,
                );
                return Err(error);
            },
        };
        if stable_length != expected {
            settle_source_publication_reservations(
                &mut memory,
                &mut input,
                &mut output,
                &mut work,
                source_scan_units(source_scanned, inserted_scanned),
                0,
            );
            return Err(TailAppendPublicationError::SourceLengthChanged {
                expected,
                observed: stable_length,
            });
        }
        let mut source_window = Vec::new();
        let source_window_len = chunk_limit.min(source_length);
        if source_window.try_reserve_exact(source_window_len).is_err() {
            settle_source_publication_reservations(
                &mut memory,
                &mut input,
                &mut output,
                &mut work,
                source_scan_units(source_scanned, inserted_scanned),
                0,
            );
            return Err(TailAppendPublicationError::Plan(Error::AllocationFailed {
                resource: "source publication window",
                requested: source_window_len,
            }));
        }
        source_window.resize(source_window_len, 0);

        let mut accepted = 0usize;
        let mut writes = 0usize;
        let mut largest_write = 0usize;
        let root_close = if self.is_noop() {
            source_length
        } else {
            self.root_close
        };
        let write_result = write_source_segment(
            source,
            0,
            root_close,
            &mut source_window,
            sink,
            chunk_limit,
            context,
            &mut source_scanned,
            &mut accepted,
            &mut writes,
            &mut largest_write,
        )
        .and_then(|()| match self.direction {
            Direction::Append => write_publication_segment(
                sink,
                &self.inserted,
                chunk_limit,
                context,
                Some(&mut inserted_scanned),
                &mut accepted,
                &mut writes,
                &mut largest_write,
            ),
            Direction::Remove => Ok(()),
        })
        .and_then(|()| {
            write_source_segment(
                source,
                if self.direction == Direction::Remove {
                    removed_end
                } else {
                    root_close
                },
                source_length,
                &mut source_window,
                sink,
                chunk_limit,
                context,
                &mut source_scanned,
                &mut accepted,
                &mut writes,
                &mut largest_write,
            )
        });
        if let Err(error) = write_result {
            settle_source_publication_reservations(
                &mut memory,
                &mut input,
                &mut output,
                &mut work,
                source_scan_units(source_scanned, inserted_scanned),
                u64::try_from(accepted).unwrap_or(u64::MAX),
            );
            return Err(with_publication_progress(
                error,
                u64::try_from(accepted).unwrap_or(u64::MAX),
                expected_output,
            ));
        }
        if let Err(error) = sink.flush() {
            settle_source_publication_reservations(
                &mut memory,
                &mut input,
                &mut output,
                &mut work,
                source_scan_units(source_scanned, inserted_scanned),
                u64::try_from(accepted).unwrap_or(u64::MAX),
            );
            return Err(TailAppendPublicationError::IncompleteOutput {
                progress: TailAppendOutputProgress::CompleteUnflushed {
                    bytes: expected_output,
                },
                source: Box::new(TailAppendPublicationError::Sink {
                    kind: error.kind(),
                    message: error.to_string(),
                    written: expected_output,
                }),
            });
        }
        if let Err(error) = context.check() {
            settle_source_publication_reservations(
                &mut memory,
                &mut input,
                &mut output,
                &mut work,
                source_scan_units(source_scanned, inserted_scanned),
                u64::try_from(accepted).unwrap_or(u64::MAX),
            );
            return Err(TailAppendPublicationError::IncompleteOutput {
                progress: TailAppendOutputProgress::CompleteUnverified {
                    bytes: expected_output,
                },
                source: Box::new(publication_execution(error, expected_output)),
            });
        }
        let final_version = match source.version() {
            Ok(version) => version,
            Err(error) => {
                settle_source_publication_reservations(
                    &mut memory,
                    &mut input,
                    &mut output,
                    &mut work,
                    source_scan_units(source_scanned, inserted_scanned),
                    u64::try_from(accepted).unwrap_or(u64::MAX),
                );
                return Err(incomplete_source_failure(error, expected_output));
            },
        };
        if final_version != self.proof.source_version {
            settle_source_publication_reservations(
                &mut memory,
                &mut input,
                &mut output,
                &mut work,
                source_scan_units(source_scanned, inserted_scanned),
                u64::try_from(accepted).unwrap_or(u64::MAX),
            );
            return Err(TailAppendPublicationError::IncompleteOutput {
                progress: TailAppendOutputProgress::CompleteUnverified {
                    bytes: expected_output,
                },
                source: Box::new(TailAppendPublicationError::SourceVersionChanged {
                    expected: self.proof.source_version,
                    observed: final_version,
                }),
            });
        }
        let final_length = match source.len() {
            Ok(length) => length,
            Err(error) => {
                settle_source_publication_reservations(
                    &mut memory,
                    &mut input,
                    &mut output,
                    &mut work,
                    source_scan_units(source_scanned, inserted_scanned),
                    u64::try_from(accepted).unwrap_or(u64::MAX),
                );
                return Err(incomplete_source_failure(error, expected_output));
            },
        };
        if final_length != expected {
            settle_source_publication_reservations(
                &mut memory,
                &mut input,
                &mut output,
                &mut work,
                source_scan_units(source_scanned, inserted_scanned),
                u64::try_from(accepted).unwrap_or(u64::MAX),
            );
            return Err(TailAppendPublicationError::IncompleteOutput {
                progress: TailAppendOutputProgress::CompleteUnverified {
                    bytes: expected_output,
                },
                source: Box::new(TailAppendPublicationError::SourceLengthChanged {
                    expected,
                    observed: final_length,
                }),
            });
        }
        drop(source_window);
        let final_digest = match source_and_target_digest_with_direction(
            source,
            source_length,
            self.root_close,
            &self.inserted,
            self.direction,
            self.publication_limits.max_window_bytes,
            Some(context),
            &mut source_scanned,
            &mut inserted_scanned,
        ) {
            Ok((digest, target)) => {
                if digest != self.source_digest {
                    settle_source_publication_reservations(
                        &mut memory,
                        &mut input,
                        &mut output,
                        &mut work,
                        source_scan_units(source_scanned, inserted_scanned),
                        u64::try_from(accepted).unwrap_or(u64::MAX),
                    );
                    return Err(TailAppendPublicationError::IncompleteOutput {
                        progress: TailAppendOutputProgress::CompleteUnverified {
                            bytes: expected_output,
                        },
                        source: Box::new(TailAppendPublicationError::SourceDigestChanged {
                            expected: self.source_digest.clone(),
                            observed: digest,
                        }),
                    });
                }
                if target != self.target_digest {
                    settle_source_publication_reservations(
                        &mut memory,
                        &mut input,
                        &mut output,
                        &mut work,
                        source_scan_units(source_scanned, inserted_scanned),
                        u64::try_from(accepted).unwrap_or(u64::MAX),
                    );
                    return Err(TailAppendPublicationError::IncompleteOutput {
                        progress: TailAppendOutputProgress::CompleteUnverified {
                            bytes: expected_output,
                        },
                        source: Box::new(TailAppendPublicationError::SourceDigestChanged {
                            expected: self.target_digest.clone(),
                            observed: target,
                        }),
                    });
                }
                digest
            },
            Err(SourceScanFailure::Execution(error)) => {
                settle_source_publication_reservations(
                    &mut memory,
                    &mut input,
                    &mut output,
                    &mut work,
                    source_scan_units(source_scanned, inserted_scanned),
                    u64::try_from(accepted).unwrap_or(u64::MAX),
                );
                return Err(TailAppendPublicationError::IncompleteOutput {
                    progress: TailAppendOutputProgress::CompleteUnverified {
                        bytes: expected_output,
                    },
                    source: Box::new(publication_execution(error, expected_output)),
                });
            },
            Err(SourceScanFailure::Io(error)) => {
                settle_source_publication_reservations(
                    &mut memory,
                    &mut input,
                    &mut output,
                    &mut work,
                    source_scan_units(source_scanned, inserted_scanned),
                    u64::try_from(accepted).unwrap_or(u64::MAX),
                );
                return Err(incomplete_source_failure(error, expected_output));
            },
            Err(SourceScanFailure::Allocation {
                resource,
                requested,
            }) => {
                settle_source_publication_reservations(
                    &mut memory,
                    &mut input,
                    &mut output,
                    &mut work,
                    source_scan_units(source_scanned, inserted_scanned),
                    u64::try_from(accepted).unwrap_or(u64::MAX),
                );
                return Err(TailAppendPublicationError::IncompleteOutput {
                    progress: TailAppendOutputProgress::CompleteUnverified {
                        bytes: expected_output,
                    },
                    source: Box::new(TailAppendPublicationError::Plan(Error::AllocationFailed {
                        resource,
                        requested,
                    })),
                });
            },
        };
        settle_source_publication_reservations(
            &mut memory,
            &mut input,
            &mut output,
            &mut work,
            source_scan_units(source_scanned, inserted_scanned),
            u64::try_from(accepted).unwrap_or(u64::MAX),
        );
        Ok(crate::source_backed::Report {
            bytes: output_length,
            source_bytes: self.source_bytes,
            inserted_bytes: if self.direction == Direction::Append {
                self.inserted.len()
            } else {
                0
            },
            source_version: self.proof.source_version,
            source_digest: final_digest,
            target_digest: self.target_digest.clone(),
            writes,
            largest_write,
        })
    }

    /// Converts this source-backed plan to the existing safe durable codec.
    /// The codec remains document-authorized on application; parsing JSON
    /// alone never creates a source proof.
    #[cfg_attr(
        not(test),
        allow(
            dead_code,
            reason = "the source-backed engine's full surface is exercised by its internal tests"
        )
    )]
    pub(crate) fn to_durable(&self) -> Result<DurableTailAppendPatch, Error> {
        let wire = DurableTailAppendPatch {
            selector: self.selector,
            direction: self.direction,
            before_digest: self.source_digest.clone(),
            after_digest: self.target_digest.clone(),
            before_bytes: self.source_bytes,
            after_bytes: self.output_bytes(),
            root_close: if self.inserted.is_empty() {
                0
            } else {
                self.root_close
            },
            inserted: self.inserted.clone(),
            limits: self.append_limits,
        };
        validate_wire_fields(&wire)?;
        validate_wire_bounds(
            wire.direction,
            wire.before_bytes,
            wire.after_bytes,
            wire.root_close,
            wire.inserted.len(),
            wire.limits,
        )?;
        let _ = wire.to_deterministic_json()?;
        Ok(wire)
    }
}

fn preflight_plain_paragraphs(
    paragraphs: &[PlainParagraph<'_>],
    limits: SemanticLimits,
) -> Result<(usize, usize, usize), Error> {
    let observed_paragraphs = paragraphs.len();
    if observed_paragraphs > limits.max_paragraphs {
        return Err(Error::Limit {
            kind: LimitKind::Paragraphs,
            observed: observed_paragraphs,
            maximum: limits.max_paragraphs,
        });
    }
    let mut observed_runs = 0usize;
    let mut observed_input_bytes = 0usize;
    for paragraph in paragraphs {
        for run in paragraph.runs() {
            observed_runs = observed_runs.checked_add(1).ok_or(Error::Limit {
                kind: LimitKind::Runs,
                observed: usize::MAX,
                maximum: limits.max_runs,
            })?;
            validate_plain_text(run.text())?;
            observed_input_bytes =
                observed_input_bytes
                    .checked_add(run.text().len())
                    .ok_or(Error::Limit {
                        kind: LimitKind::InputBytes,
                        observed: usize::MAX,
                        maximum: limits.max_input_bytes,
                    })?;
        }
    }
    if observed_runs > limits.max_runs {
        return Err(Error::Limit {
            kind: LimitKind::Runs,
            observed: observed_runs,
            maximum: limits.max_runs,
        });
    }
    if observed_input_bytes > limits.max_input_bytes {
        return Err(Error::Limit {
            kind: LimitKind::InputBytes,
            observed: observed_input_bytes,
            maximum: limits.max_input_bytes,
        });
    }
    Ok((observed_paragraphs, observed_runs, observed_input_bytes))
}

fn preflight_plain_text_paragraphs(
    paragraphs: &[&str],
    limits: SemanticLimits,
) -> Result<(), Error> {
    let observed_paragraphs = paragraphs.len();
    if observed_paragraphs > limits.max_paragraphs {
        return Err(Error::Limit {
            kind: LimitKind::Paragraphs,
            observed: observed_paragraphs,
            maximum: limits.max_paragraphs,
        });
    }
    if observed_paragraphs > limits.max_runs {
        return Err(Error::Limit {
            kind: LimitKind::Runs,
            observed: observed_paragraphs,
            maximum: limits.max_runs,
        });
    }
    let mut observed_input_bytes = 0usize;
    for text in paragraphs {
        validate_plain_text(text)?;
        observed_input_bytes =
            observed_input_bytes
                .checked_add(text.len())
                .ok_or(Error::Limit {
                    kind: LimitKind::InputBytes,
                    observed: usize::MAX,
                    maximum: limits.max_input_bytes,
                })?;
    }
    if observed_input_bytes > limits.max_input_bytes {
        return Err(Error::Limit {
            kind: LimitKind::InputBytes,
            observed: observed_input_bytes,
            maximum: limits.max_input_bytes,
        });
    }
    Ok(())
}

fn stage_plain_paragraphs(
    paragraphs: &[PlainParagraph<'_>],
    limits: SemanticLimits,
) -> Result<(Vec<OwnedParagraph>, usize, usize), Error> {
    let (observed_paragraphs, observed_runs, observed_input_bytes) =
        preflight_plain_paragraphs(paragraphs, limits)?;
    let mut owned = Vec::new();
    owned
        .try_reserve(observed_paragraphs)
        .map_err(|_error| Error::AllocationFailed {
            resource: "tail paragraphs",
            requested: observed_paragraphs.saturating_mul(size_of::<OwnedParagraph>()),
        })?;
    for paragraph in paragraphs {
        let mut runs = Vec::new();
        runs.try_reserve(paragraph.runs().len())
            .map_err(|_error| Error::AllocationFailed {
                resource: "tail runs",
                requested: paragraph.runs().len().saturating_mul(size_of::<String>()),
            })?;
        for run in paragraph.runs() {
            let mut text = String::new();
            text.try_reserve(run.text().len())
                .map_err(|_error| Error::AllocationFailed {
                    resource: "tail text",
                    requested: run.text().len(),
                })?;
            text.push_str(run.text());
            runs.push(text);
        }
        owned.push(OwnedParagraph { runs });
    }
    Ok((owned, observed_runs, observed_input_bytes))
}

#[derive(Debug)]
enum SourceScanFailure {
    Io(io::Error),
    Execution(ExecutionError),
    Allocation {
        resource: &'static str,
        requested: usize,
    },
}

fn source_publication_io(error: io::Error, written: u64) -> TailAppendPublicationError {
    TailAppendPublicationError::Source { error, written }
}

fn source_scan_publication_error(
    error: SourceScanFailure,
    written: u64,
) -> TailAppendPublicationError {
    match error {
        SourceScanFailure::Io(error) => source_publication_io(error, written),
        SourceScanFailure::Execution(error) => publication_execution(error, written),
        SourceScanFailure::Allocation {
            resource,
            requested,
        } => TailAppendPublicationError::Plan(Error::AllocationFailed {
            resource,
            requested,
        }),
    }
}

fn incomplete_source_failure(error: io::Error, written: u64) -> TailAppendPublicationError {
    TailAppendPublicationError::IncompleteOutput {
        progress: TailAppendOutputProgress::CompleteUnverified { bytes: written },
        source: Box::new(source_publication_io(error, written)),
    }
}

fn read_source_window(
    source: &dyn ReadAt,
    offset: usize,
    output: &mut [u8],
    context: Option<&ExecutionContext>,
    source_scanned: &mut usize,
) -> Result<(), SourceScanFailure> {
    let mut filled = 0usize;
    while filled < output.len() {
        if let Some(context) = context {
            context.check().map_err(SourceScanFailure::Execution)?;
        }
        let source_offset = u64::try_from(offset.checked_add(filled).ok_or_else(|| {
            SourceScanFailure::Io(io::Error::new(
                io::ErrorKind::InvalidInput,
                "source offset overflow",
            ))
        })?)
        .map_err(|_error| {
            SourceScanFailure::Io(io::Error::new(
                io::ErrorKind::InvalidInput,
                "source offset does not fit u64",
            ))
        })?;
        let remaining = output.get_mut(filled..).ok_or_else(|| {
            SourceScanFailure::Io(io::Error::new(
                io::ErrorKind::InvalidData,
                "source window is outside its allocation",
            ))
        })?;
        match source.read_at(source_offset, remaining) {
            Ok(0) => {
                return Err(SourceScanFailure::Io(io::Error::new(
                    io::ErrorKind::UnexpectedEof,
                    "positional source ended before the requested range",
                )));
            },
            Ok(count) if count > remaining.len() => {
                return Err(SourceScanFailure::Io(io::Error::new(
                    io::ErrorKind::InvalidData,
                    "positional source reported more bytes than requested",
                )));
            },
            Ok(count) => {
                filled = filled.saturating_add(count);
                *source_scanned = (*source_scanned).saturating_add(count);
            },
            Err(error) if error.kind() == io::ErrorKind::Interrupted => {},
            Err(error) => return Err(SourceScanFailure::Io(error)),
        }
    }
    Ok(())
}

fn source_and_target_digest(
    source: &dyn ReadAt,
    source_length: usize,
    root_close: usize,
    inserted: &[u8],
    window_limit: usize,
    context: Option<&ExecutionContext>,
    source_scanned: &mut usize,
    inserted_scanned: &mut usize,
) -> Result<(String, String), SourceScanFailure> {
    source_and_target_digest_with_direction(
        source,
        source_length,
        root_close,
        inserted,
        Direction::Append,
        window_limit,
        context,
        source_scanned,
        inserted_scanned,
    )
}

fn source_and_target_digest_with_direction(
    source: &dyn ReadAt,
    source_length: usize,
    root_close: usize,
    inserted: &[u8],
    direction: Direction,
    window_limit: usize,
    context: Option<&ExecutionContext>,
    source_scanned: &mut usize,
    inserted_scanned: &mut usize,
) -> Result<(String, String), SourceScanFailure> {
    if window_limit == 0 || root_close > source_length {
        return Err(SourceScanFailure::Io(io::Error::new(
            io::ErrorKind::InvalidInput,
            "source digest window or root offset is invalid",
        )));
    }
    let removed_end = root_close.checked_add(inserted.len()).ok_or_else(|| {
        SourceScanFailure::Io(io::Error::new(
            io::ErrorKind::InvalidInput,
            "source digest span overflows",
        ))
    })?;
    if direction == Direction::Remove && removed_end > source_length {
        return Err(SourceScanFailure::Io(io::Error::new(
            io::ErrorKind::InvalidInput,
            "source digest removal is outside the source",
        )));
    }
    let mut source_hasher = Sha256::new();
    let mut target_hasher = Sha256::new();
    if source_length == 0 {
        if direction == Direction::Append {
            target_hasher.update(inserted);
            *inserted_scanned = (*inserted_scanned).saturating_add(inserted.len());
        }
        return Ok((sha256_hex(source_hasher), sha256_hex(target_hasher)));
    }
    let window_limit = window_limit.min(source_length);
    let mut window = Vec::new();
    window
        .try_reserve_exact(window_limit)
        .map_err(|_error| SourceScanFailure::Allocation {
            resource: "source digest window",
            requested: window_limit,
        })?;
    window.resize(window_limit, 0);
    match direction {
        Direction::Append => {
            hash_source_range(
                source,
                0,
                root_close,
                &mut window,
                &mut source_hasher,
                &mut target_hasher,
                None,
                context,
                source_scanned,
            )?;
            target_hasher.update(inserted);
            *inserted_scanned = (*inserted_scanned).saturating_add(inserted.len());
            hash_source_range(
                source,
                root_close,
                source_length,
                &mut window,
                &mut source_hasher,
                &mut target_hasher,
                None,
                context,
                source_scanned,
            )?;
        },
        Direction::Remove => {
            hash_source_range(
                source,
                0,
                source_length,
                &mut window,
                &mut source_hasher,
                &mut target_hasher,
                Some((root_close, removed_end)),
                context,
                source_scanned,
            )?;
        },
    }
    Ok((sha256_hex(source_hasher), sha256_hex(target_hasher)))
}

fn hash_source_range(
    source: &dyn ReadAt,
    start: usize,
    end: usize,
    window: &mut [u8],
    source_hasher: &mut Sha256,
    target_hasher: &mut Sha256,
    target_skip: Option<(usize, usize)>,
    context: Option<&ExecutionContext>,
    source_scanned: &mut usize,
) -> Result<(), SourceScanFailure> {
    let mut offset = start;
    while offset < end {
        let count = end.saturating_sub(offset).min(window.len());
        let chunk = window.get_mut(..count).ok_or_else(|| {
            SourceScanFailure::Io(io::Error::new(
                io::ErrorKind::InvalidData,
                "source digest range exceeds its window",
            ))
        })?;
        read_source_window(source, offset, chunk, context, source_scanned)?;
        source_hasher.update(&*chunk);
        if let Some((skip_start, skip_end)) = target_skip {
            let chunk_end = offset.saturating_add(count);
            if skip_end <= offset || skip_start >= chunk_end {
                target_hasher.update(&*chunk);
            } else {
                let left_end = skip_start.min(chunk_end);
                if left_end > offset {
                    let left = chunk.get(..left_end - offset).ok_or_else(|| {
                        SourceScanFailure::Io(io::Error::new(
                            io::ErrorKind::InvalidData,
                            "source digest range exceeds its window",
                        ))
                    })?;
                    target_hasher.update(left);
                }
                let right_start = skip_end.max(offset);
                if right_start < chunk_end {
                    let right = chunk.get(right_start - offset..).ok_or_else(|| {
                        SourceScanFailure::Io(io::Error::new(
                            io::ErrorKind::InvalidData,
                            "source digest range exceeds its window",
                        ))
                    })?;
                    target_hasher.update(right);
                }
            }
        } else {
            target_hasher.update(&*chunk);
        }
        offset = offset.saturating_add(count);
    }
    Ok(())
}

fn sha256_hex(hasher: Sha256) -> String {
    let digest = hasher.finalize();
    let mut text = String::with_capacity(64);
    for byte in digest {
        use std::fmt::Write as _;
        let _ = write!(text, "{byte:02x}");
    }
    text
}

fn write_source_segment<W: Write>(
    source: &dyn ReadAt,
    start: usize,
    end: usize,
    window: &mut [u8],
    sink: &mut W,
    chunk_limit: usize,
    context: &ExecutionContext,
    source_scanned: &mut usize,
    accepted: &mut usize,
    writes: &mut usize,
    largest_write: &mut usize,
) -> Result<(), PublicationWriteFailure> {
    if start >= end {
        return Ok(());
    }
    let mut offset = start;
    while offset < end {
        context
            .check()
            .map_err(PublicationWriteFailure::Execution)?;
        let count = end.saturating_sub(offset).min(chunk_limit);
        let chunk = window
            .get_mut(..count)
            .ok_or(PublicationWriteFailure::Allocation {
                resource: "source publication window",
                requested: count,
            })?;
        read_source_window(source, offset, chunk, Some(context), source_scanned).map_err(
            |error| match error {
                SourceScanFailure::Io(error) => PublicationWriteFailure::Source(error),
                SourceScanFailure::Execution(error) => PublicationWriteFailure::Execution(error),
                SourceScanFailure::Allocation {
                    resource,
                    requested,
                } => PublicationWriteFailure::Allocation {
                    resource,
                    requested,
                },
            },
        )?;
        write_publication_segment(
            sink,
            chunk,
            chunk_limit,
            context,
            None,
            accepted,
            writes,
            largest_write,
        )?;
        offset = offset.saturating_add(count);
    }
    Ok(())
}

#[derive(Debug)]
enum PublicationWriteFailure {
    Execution(ExecutionError),
    Source(io::Error),
    Allocation {
        resource: &'static str,
        requested: usize,
    },
    Sink {
        kind: io::ErrorKind,
        message: String,
    },
    Overreported,
}

fn write_publication_segment<W: Write>(
    sink: &mut W,
    bytes: &[u8],
    chunk_limit: usize,
    context: &ExecutionContext,
    mut processed: Option<&mut usize>,
    accepted: &mut usize,
    writes: &mut usize,
    largest_write: &mut usize,
) -> Result<(), PublicationWriteFailure> {
    if bytes.is_empty() {
        return Ok(());
    }
    let mut offset = 0usize;
    while offset < bytes.len() {
        context
            .check()
            .map_err(PublicationWriteFailure::Execution)?;
        let end = offset
            .checked_add(chunk_limit)
            .unwrap_or(usize::MAX)
            .min(bytes.len());
        let window = bytes
            .get(offset..end)
            .ok_or(PublicationWriteFailure::Sink {
                kind: io::ErrorKind::InvalidData,
                message: "publication window is outside the source span".to_string(),
            })?;
        let mut window_offset = 0usize;
        while window_offset < window.len() {
            context
                .check()
                .map_err(PublicationWriteFailure::Execution)?;
            let remaining = window
                .get(window_offset..)
                .ok_or(PublicationWriteFailure::Sink {
                    kind: io::ErrorKind::InvalidData,
                    message: "publication write window is outside the source span".to_string(),
                })?;
            *writes = writes.saturating_add(1);
            *largest_write = (*largest_write).max(remaining.len());
            match sink.write(remaining) {
                Ok(0) => {
                    return Err(PublicationWriteFailure::Sink {
                        kind: io::ErrorKind::WriteZero,
                        message: "publication sink returned zero progress".to_string(),
                    });
                },
                Ok(count) if count > remaining.len() => {
                    return Err(PublicationWriteFailure::Overreported);
                },
                Ok(count) => {
                    window_offset = window_offset.saturating_add(count);
                    *accepted = accepted.saturating_add(count);
                    if let Some(processed) = processed.as_deref_mut() {
                        *processed = (*processed).saturating_add(count);
                    }
                    context
                        .check()
                        .map_err(PublicationWriteFailure::Execution)?;
                },
                Err(error) if error.kind() == io::ErrorKind::Interrupted => {},
                Err(error) => {
                    return Err(PublicationWriteFailure::Sink {
                        kind: error.kind(),
                        message: error.to_string(),
                    });
                },
            }
        }
        offset = end;
    }
    Ok(())
}

fn settle_publication_reservations(
    memory: &mut Option<Reservation>,
    input: &mut Option<Reservation>,
    output: &mut Option<Reservation>,
    work: &mut Option<Reservation>,
    accepted: u64,
) {
    settle_source_publication_reservations(memory, input, output, work, accepted, accepted);
}

fn settle_source_publication_reservations(
    memory: &mut Option<Reservation>,
    input: &mut Option<Reservation>,
    output: &mut Option<Reservation>,
    work: &mut Option<Reservation>,
    input_work: u64,
    output_bytes: u64,
) {
    if let Some(reservation) = memory.take() {
        let _ = reservation.commit(0);
    }
    if let Some(reservation) = input.take() {
        let _ = reservation.commit(input_work);
    }
    if let Some(reservation) = output.take() {
        let _ = reservation.commit(output_bytes);
    }
    if let Some(reservation) = work.take() {
        let _ = reservation.commit(input_work);
    }
}

fn source_scan_units(source_scanned: usize, inserted_scanned: usize) -> u64 {
    u64::try_from(source_scanned.saturating_add(inserted_scanned)).unwrap_or(u64::MAX)
}

fn publication_execution(error: ExecutionError, written: u64) -> TailAppendPublicationError {
    TailAppendPublicationError::Execution { error, written }
}

fn with_publication_progress(
    error: PublicationWriteFailure,
    accepted: u64,
    expected: u64,
) -> TailAppendPublicationError {
    let indeterminate = matches!(&error, PublicationWriteFailure::Overreported);
    let error = match error {
        PublicationWriteFailure::Execution(error) => publication_execution(error, accepted),
        PublicationWriteFailure::Source(error) => source_publication_io(error, accepted),
        PublicationWriteFailure::Allocation {
            resource,
            requested,
        } => TailAppendPublicationError::Plan(Error::AllocationFailed {
            resource,
            requested,
        }),
        PublicationWriteFailure::Sink { kind, message } => TailAppendPublicationError::Sink {
            kind,
            message,
            written: accepted,
        },
        PublicationWriteFailure::Overreported => TailAppendPublicationError::Sink {
            kind: io::ErrorKind::InvalidData,
            message: "publication sink reported more bytes than offered".to_string(),
            written: accepted,
        },
    };
    if accepted == 0 && !indeterminate {
        return error;
    }
    let progress = if indeterminate {
        TailAppendOutputProgress::Indeterminate {
            accepted_before: accepted,
        }
    } else if accepted == expected {
        TailAppendOutputProgress::CompleteUnverified { bytes: accepted }
    } else {
        TailAppendOutputProgress::Prefix { accepted, expected }
    };
    TailAppendPublicationError::IncompleteOutput {
        progress,
        source: Box::new(error),
    }
}

/// Deterministic facts about a logical-tail commit.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct Diagnostics {
    changed: bool,
    operation_count: usize,
    paragraphs: usize,
    runs: usize,
    input_bytes: usize,
    source_bytes: usize,
    inserted_bytes: usize,
    output_bytes: usize,
}

impl Diagnostics {
    /// Whether a distinct candidate artifact was published.
    #[must_use]
    pub const fn changed(self) -> bool {
        self.changed
    }

    /// Number of staged append batches, or one for a changed patch replay.
    #[must_use]
    pub const fn operation_count(self) -> usize {
        self.operation_count
    }

    /// Number of caller-staged paragraphs; zero when replaying a patch.
    #[must_use]
    pub const fn paragraphs(self) -> usize {
        self.paragraphs
    }

    /// Number of caller-staged runs; zero when replaying a patch.
    #[must_use]
    pub const fn runs(self) -> usize {
        self.runs
    }

    /// UTF-8 bytes copied from caller-owned paragraph runs.
    #[must_use]
    pub const fn input_bytes(self) -> usize {
        self.input_bytes
    }

    /// Exact source artifact size before the change.
    #[must_use]
    pub const fn source_bytes(self) -> usize {
        self.source_bytes
    }

    /// Exact inserted byte count; zero for an inverse removal.
    #[must_use]
    pub const fn inserted_bytes(self) -> usize {
        self.inserted_bytes
    }

    /// Exact candidate artifact size.
    #[must_use]
    pub const fn output_bytes(self) -> usize {
        self.output_bytes
    }
}

/// Result of a validated logical-tail append.
pub struct Commit {
    snapshot: Package,
    patch: Patch,
    diagnostics: Diagnostics,
}

impl Commit {
    pub(crate) fn from_replay(package: Package, patch: Patch) -> Self {
        let source_bytes = patch.before.source_bytes().map_or(0, <[u8]>::len);
        let output_bytes = patch.after.source_bytes().map_or(source_bytes, <[u8]>::len);
        Self {
            snapshot: package,
            diagnostics: Diagnostics {
                changed: !patch.is_noop(),
                operation_count: usize::from(!patch.is_noop()),
                paragraphs: 0,
                runs: 0,
                input_bytes: 0,
                source_bytes,
                inserted_bytes: if patch.direction == Direction::Append {
                    patch.inserted.len
                } else {
                    0
                },
                output_bytes,
            },
            patch,
        }
    }

    /// Published immutable candidate snapshot.
    #[must_use]
    pub const fn package(&self) -> &Package {
        &self.snapshot
    }

    /// Consume the commit and return its published package.
    #[must_use]
    pub fn into_package(self) -> Package {
        self.snapshot
    }

    /// Split the published package and its reversible patch.
    #[must_use]
    pub fn into_parts(self) -> (Package, Patch) {
        (self.snapshot, self.patch)
    }

    /// Exact-source-checked reversible in-memory patch.
    #[must_use]
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Deterministic commit diagnostics.
    #[must_use]
    pub const fn diagnostics(&self) -> Diagnostics {
        self.diagnostics
    }

    /// Writes the accepted artifact sequentially without requiring a seekable
    /// sink.  The source snapshot and patch remain unchanged on partial sink
    /// failure.
    pub fn write_to<W: Write>(&self, sink: &mut W, limits: SemanticLimits) -> Result<usize, Error> {
        let bytes = self
            .snapshot
            .source_bytes()
            .ok_or(Error::unsupported_operation(
                "snapshot has no exact RTF source",
            ))?;
        if bytes.len() > limits.max_output_bytes {
            return Err(Error::Limit {
                kind: LimitKind::OutputBytes,
                observed: bytes.len(),
                maximum: limits.max_output_bytes,
            });
        }
        write_sequential(sink, bytes)
    }
}

/// Exact-source-checked reversible logical-tail patch.
#[derive(Clone)]
pub struct Patch {
    before: Package,
    after: Package,
    selector: TailSelector,
    root_close: usize,
    inserted: InsertedSpan,
    direction: Direction,
}

impl Patch {
    pub(crate) fn apply_commit(&self, source: &Package) -> Result<Commit, Error> {
        let package = self.apply(source)?;
        Ok(Commit::from_replay(package, self.clone()))
    }

    fn no_op(source: Package, selector: TailSelector) -> Self {
        Self {
            before: source.clone(),
            after: source,
            selector,
            root_close: 0,
            inserted: InsertedSpan::EMPTY,
            direction: Direction::Append,
        }
    }

    /// Applies this patch only to the exact source artifact that created it.
    pub(crate) fn apply(&self, source: &Package) -> Result<Package, Error> {
        if source.source_bytes() != self.before.source_bytes() {
            return Err(Error::PatchConflict);
        }
        if self.is_noop() {
            return Ok(source.clone());
        }
        if source.options() == self.after.options() {
            return Ok(self.after.clone());
        }
        let bytes = self.after.source_bytes().ok_or(Error::unsupported_source(
            "patch target has no retained source artifact",
        ))?;
        Ok(Package::from_bytes_with_options(bytes, source.options())?)
    }

    /// Whether this patch preserves the exact source artifact.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.before.source_bytes() == self.after.source_bytes()
    }

    /// Returns the exact inverse patch.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            before: self.after.clone(),
            after: self.before.clone(),
            selector: self.selector,
            root_close: self.root_close,
            inserted: self.inserted,
            direction: if self.inserted.is_empty() {
                Direction::Append
            } else {
                self.direction.inverse()
            },
        }
    }

    /// Selector retained by this patch.
    #[must_use]
    pub const fn selector(&self) -> TailSelector {
        self.selector
    }

    /// Number of exact inserted bytes represented by this patch.
    #[must_use]
    pub const fn inserted_bytes(&self) -> usize {
        self.inserted.len
    }

    /// Converts this patch to a deterministic bounded JSON envelope.
    pub fn to_durable(&self, limits: SemanticLimits) -> Result<DurableTailAppendPatch, Error> {
        let before = self.before.source_bytes().ok_or(Error::unsupported_source(
            "patch source has no exact RTF source",
        ))?;
        let after = self.after.source_bytes().ok_or(Error::unsupported_source(
            "patch target has no exact RTF source",
        ))?;
        let wire = DurableTailAppendPatch {
            selector: self.selector,
            direction: self.direction,
            before_digest: BlobId::of(before).as_hex(),
            after_digest: BlobId::of(after).as_hex(),
            before_bytes: before.len(),
            after_bytes: after.len(),
            root_close: self.root_close,
            inserted: self.inserted_source_bytes()?.to_vec().into_boxed_slice(),
            limits,
        };
        validate_wire_fields(&wire)?;
        validate_artifact_splice(
            before,
            after,
            wire.direction,
            wire.before_bytes,
            wire.after_bytes,
            wire.root_close,
            &wire.inserted,
            wire.limits,
        )?;
        let _ = wire.to_deterministic_json()?;
        Ok(wire)
    }

    fn inserted_source_bytes(&self) -> Result<&[u8], Error> {
        if self.inserted.offset != self.root_close {
            return Err(Error::PatchConflict);
        }
        let source = match self.direction {
            Direction::Append => self.after.source_bytes(),
            Direction::Remove => self.before.source_bytes(),
        }
        .ok_or(Error::unsupported_source(
            "tail patch source has no exact RTF source",
        ))?;
        let end = self
            .inserted
            .offset
            .checked_add(self.inserted.len)
            .ok_or(Error::PatchConflict)?;
        source
            .get(self.inserted.offset..end)
            .ok_or(Error::PatchConflict)
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
struct InsertedSpan {
    offset: usize,
    len: usize,
}

impl InsertedSpan {
    const EMPTY: Self = Self { offset: 0, len: 0 };

    const fn is_empty(self) -> bool {
        self.len == 0
    }
}

/// Durable, source-digest-checked logical-tail patch.
#[derive(Clone)]
pub struct DurableTailAppendPatch {
    selector: TailSelector,
    direction: Direction,
    before_digest: String,
    after_digest: String,
    before_bytes: usize,
    after_bytes: usize,
    root_close: usize,
    inserted: Box<[u8]>,
    limits: SemanticLimits,
}

impl DurableTailAppendPatch {
    /// Serializes the patch in deterministic JSON and enforces its wire cap.
    pub fn to_deterministic_json(&self) -> Result<Vec<u8>, Error> {
        validate_wire_fields(self)?;
        validate_wire_bounds(
            self.direction,
            self.before_bytes,
            self.after_bytes,
            self.root_close,
            self.inserted.len(),
            self.limits,
        )?;
        let encoded_len = deterministic_json_len(self)?;
        if encoded_len > self.limits.max_patch_bytes {
            return Err(Error::Limit {
                kind: LimitKind::PatchBytes,
                observed: encoded_len,
                maximum: self.limits.max_patch_bytes,
            });
        }
        let encoded = hex_encode(&self.inserted)?;
        let mut object = Map::new();
        object.insert(
            "after".to_string(),
            Value::String(self.after_digest.clone()),
        );
        object.insert("after_bytes".to_string(), number(self.after_bytes)?);
        object.insert(
            "before".to_string(),
            Value::String(self.before_digest.clone()),
        );
        object.insert("before_bytes".to_string(), number(self.before_bytes)?);
        object.insert(
            "direction".to_string(),
            Value::String(self.direction.as_str().to_string()),
        );
        object.insert(
            "format".to_string(),
            Value::String("litchi-rtf-tail-append".to_string()),
        );
        object.insert("inserted_hex".to_string(), Value::String(encoded));
        object.insert("offset".to_string(), number(self.root_close)?);
        object.insert("selector".to_string(), Value::String("body".to_string()));
        object.insert("version".to_string(), Value::from(1_u64));
        // Sorted members whatever `serde_json` features are unified (0349).
        let bytes = litchi_core::canonical_json::to_vec(&Value::Object(object));
        if bytes.len() != encoded_len {
            return Err(Error::DurablePatch(
                "deterministic JSON size preflight disagrees with serializer",
            ));
        }
        if bytes.len() > self.limits.max_patch_bytes {
            return Err(Error::Limit {
                kind: LimitKind::PatchBytes,
                observed: bytes.len(),
                maximum: self.limits.max_patch_bytes,
            });
        }
        Ok(bytes)
    }

    /// Parses a canonical deterministic JSON envelope under explicit bounds.
    pub fn from_deterministic_json(bytes: &[u8], limits: SemanticLimits) -> Result<Self, Error> {
        if bytes.len() > limits.max_patch_bytes {
            return Err(Error::Limit {
                kind: LimitKind::PatchBytes,
                observed: bytes.len(),
                maximum: limits.max_patch_bytes,
            });
        }
        let value: Value = serde_json::from_slice(bytes)
            .map_err(|_error| Error::DurablePatch("JSON is malformed"))?;
        let canonical = litchi_core::canonical_json::to_vec(&value);
        if canonical != bytes {
            return Err(Error::DurablePatch("JSON is not canonical"));
        }
        let Value::Object(object) = value else {
            return Err(Error::DurablePatch("patch root must be an object"));
        };
        let expected = [
            "after",
            "after_bytes",
            "before",
            "before_bytes",
            "direction",
            "format",
            "inserted_hex",
            "offset",
            "selector",
            "version",
        ];
        let keys = object.keys().map(String::as_str).collect::<BTreeSet<_>>();
        if keys.len() != expected.len() || expected.iter().any(|key| !keys.contains(key)) {
            return Err(Error::DurablePatch("patch fields are not exact"));
        }
        if object.get("format") != Some(&Value::String("litchi-rtf-tail-append".to_string()))
            || object.get("selector") != Some(&Value::String("body".to_string()))
            || object.get("version") != Some(&Value::from(1_u64))
        {
            return Err(Error::DurablePatch("unsupported patch format or version"));
        }
        let before_digest = string_field(&object, "before")?.to_string();
        let after_digest = string_field(&object, "after")?.to_string();
        validate_digest(&before_digest)?;
        validate_digest(&after_digest)?;
        let before_bytes = number_field(&object, "before_bytes")?;
        let after_bytes = number_field(&object, "after_bytes")?;
        let direction = Direction::parse(string_field(&object, "direction")?)?;
        let root_close = number_field(&object, "offset")?;
        let encoded = string_field(&object, "inserted_hex")?;
        let inserted = hex_decode(encoded, limits.max_inserted_bytes)?;
        let patch = Self {
            selector: TailSelector::Body,
            direction,
            before_digest,
            after_digest,
            before_bytes,
            after_bytes,
            root_close,
            inserted: inserted.into_boxed_slice(),
            limits,
        };
        validate_wire_fields(&patch)?;
        validate_wire_bounds(
            patch.direction,
            patch.before_bytes,
            patch.after_bytes,
            patch.root_close,
            patch.inserted.len(),
            patch.limits,
        )?;
        // Re-run the complete bounded wire check, including output size.
        let _ = patch.to_deterministic_json()?;
        Ok(patch)
    }

    /// Returns the exact reverse operation.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            selector: self.selector,
            direction: if self.inserted.is_empty() {
                Direction::Append
            } else {
                self.direction.inverse()
            },
            before_digest: self.after_digest.clone(),
            after_digest: self.before_digest.clone(),
            before_bytes: self.after_bytes,
            after_bytes: self.before_bytes,
            root_close: self.root_close,
            inserted: self.inserted.clone(),
            limits: self.limits,
        }
    }

    /// Applies this patch to an exact source artifact and reopens the result.
    pub fn apply(&self, source: &Package) -> Result<Package, Error> {
        if self.selector != TailSelector::Body {
            return Err(Error::Refused(Refusal::new(
                RefusalKind::UnsupportedTarget,
                "tail_append",
                "the selector is not supported by this edit family",
            )));
        }
        validate_wire_fields(self)?;
        validate_wire_bounds(
            self.direction,
            self.before_bytes,
            self.after_bytes,
            self.root_close,
            self.inserted.len(),
            self.limits,
        )?;
        let source_bytes = source.source_bytes().ok_or(Error::unsupported_operation(
            "snapshot has no exact RTF source",
        ))?;
        if self.direction == Direction::Append && source_bytes.len() > self.limits.max_output_bytes
        {
            return Err(Error::Limit {
                kind: LimitKind::OutputBytes,
                observed: source_bytes.len(),
                maximum: self.limits.max_output_bytes,
            });
        }
        if source_bytes.len() != self.before_bytes {
            return Err(Error::PatchConflict);
        }
        if BlobId::of(source_bytes).as_hex() != self.before_digest {
            return Err(Error::PatchConflict);
        }
        if self.inserted.is_empty() && self.before_digest == self.after_digest {
            // Empty append patches are exact identity transitions, but the
            // source still needs the same complete splice proof as a changed
            // patch. This prevents a forged equal-digest envelope from using
            // the identity fast path to bypass malformed/protected input.
            let _ = prove_splice_source(source, self.selector)?;
            return Ok(source.clone());
        }
        let proof = prove_splice_source(source, self.selector)?;
        let expected_root_close = match self.direction {
            Direction::Append => self.root_close,
            Direction::Remove => self
                .root_close
                .checked_add(self.inserted.len())
                .ok_or(Error::PatchConflict)?,
        };
        if expected_root_close != proof.root_close {
            return Err(Error::PatchConflict);
        }
        let output =
            match self.direction {
                Direction::Append => {
                    if self.root_close > source_bytes.len() {
                        return Err(Error::PatchConflict);
                    }
                    let mut output = Vec::new();
                    let output_len = source_bytes.len().checked_add(self.inserted.len()).ok_or(
                        Error::Limit {
                            kind: LimitKind::OutputBytes,
                            observed: usize::MAX,
                            maximum: self.limits.max_output_bytes,
                        },
                    )?;
                    if output_len != self.after_bytes {
                        return Err(Error::PatchConflict);
                    }
                    if output_len > source.limits().max_input_bytes() {
                        return Err(Error::Limit {
                            kind: LimitKind::SourceBytes,
                            observed: output_len,
                            maximum: source.limits().max_input_bytes(),
                        });
                    }
                    if output_len > self.limits.max_output_bytes {
                        return Err(Error::Limit {
                            kind: LimitKind::OutputBytes,
                            observed: output_len,
                            maximum: self.limits.max_output_bytes,
                        });
                    }
                    output
                        .try_reserve(output_len)
                        .map_err(|_error| Error::AllocationFailed {
                            resource: "durable tail candidate",
                            requested: output_len,
                        })?;
                    let prefix = source_bytes
                        .get(..self.root_close)
                        .ok_or(Error::PatchConflict)?;
                    let suffix = source_bytes
                        .get(self.root_close..)
                        .ok_or(Error::PatchConflict)?;
                    output.extend_from_slice(prefix);
                    output.extend_from_slice(&self.inserted);
                    output.extend_from_slice(suffix);
                    output
                },
                Direction::Remove => {
                    let end = self
                        .root_close
                        .checked_add(self.inserted.len())
                        .ok_or(Error::PatchConflict)?;
                    if end > source_bytes.len()
                        || source_bytes.get(self.root_close..end) != Some(&self.inserted)
                    {
                        return Err(Error::PatchConflict);
                    }
                    let output_len = source_bytes
                        .len()
                        .checked_sub(self.inserted.len())
                        .ok_or(Error::PatchConflict)?;
                    if output_len > self.limits.max_output_bytes {
                        return Err(Error::Limit {
                            kind: LimitKind::OutputBytes,
                            observed: output_len,
                            maximum: self.limits.max_output_bytes,
                        });
                    }
                    if output_len != self.after_bytes {
                        return Err(Error::PatchConflict);
                    }
                    let mut output = Vec::new();
                    output
                        .try_reserve(output_len)
                        .map_err(|_error| Error::AllocationFailed {
                            resource: "durable inverse candidate",
                            requested: output_len,
                        })?;
                    let prefix = source_bytes
                        .get(..self.root_close)
                        .ok_or(Error::PatchConflict)?;
                    let suffix = source_bytes.get(end..).ok_or(Error::PatchConflict)?;
                    output.extend_from_slice(prefix);
                    output.extend_from_slice(suffix);
                    output
                },
            };
        if BlobId::of(&output).as_hex() != self.after_digest {
            return Err(Error::PatchConflict);
        }
        Ok(Package::from_bytes_with_options(&output, source.options())?)
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
enum Direction {
    Append,
    Remove,
}

impl Direction {
    const fn as_str(self) -> &'static str {
        match self {
            Self::Append => "append",
            Self::Remove => "remove",
        }
    }

    const fn inverse(self) -> Self {
        match self {
            Self::Append => Self::Remove,
            Self::Remove => Self::Append,
        }
    }

    fn parse(value: &str) -> Result<Self, Error> {
        match value {
            "append" => Ok(Self::Append),
            "remove" => Ok(Self::Remove),
            _ => Err(Error::DurablePatch("invalid patch direction")),
        }
    }
}

fn validate_wire_fields(patch: &DurableTailAppendPatch) -> Result<(), Error> {
    validate_digest(&patch.before_digest)?;
    validate_digest(&patch.after_digest)?;
    if patch.inserted.is_empty() {
        if patch.direction != Direction::Append
            || patch.root_close != 0
            || patch.before_digest != patch.after_digest
        {
            return Err(Error::DurablePatch(
                "empty durable patch is not the canonical no-op shape",
            ));
        }
    } else if patch.before_digest == patch.after_digest {
        return Err(Error::DurablePatch(
            "changed durable patch has equal source and target digests",
        ));
    }
    Ok(())
}

fn validate_wire_bounds(
    direction: Direction,
    before_bytes: usize,
    after_bytes: usize,
    root_close: usize,
    inserted_len: usize,
    limits: SemanticLimits,
) -> Result<(), Error> {
    if inserted_len > limits.max_inserted_bytes {
        return Err(Error::Limit {
            kind: LimitKind::InsertedBytes,
            observed: inserted_len,
            maximum: limits.max_inserted_bytes,
        });
    }
    if root_close > limits.max_output_bytes {
        return Err(Error::Limit {
            kind: LimitKind::PatchOffset,
            observed: root_close,
            maximum: limits.max_output_bytes,
        });
    }
    if after_bytes > limits.max_output_bytes {
        return Err(Error::Limit {
            kind: LimitKind::OutputBytes,
            observed: after_bytes,
            maximum: limits.max_output_bytes,
        });
    }
    match direction {
        Direction::Append => {
            let expected_after = before_bytes.checked_add(inserted_len).ok_or(Error::Limit {
                kind: LimitKind::AfterBytes,
                observed: usize::MAX,
                maximum: limits.max_output_bytes,
            })?;
            if after_bytes != expected_after {
                return Err(Error::DurablePatch(
                    "append before/after byte sizes do not match inserted bytes",
                ));
            }
            if root_close > before_bytes {
                return Err(Error::DurablePatch(
                    "append offset exceeds before artifact size",
                ));
            }
        },
        Direction::Remove => {
            let expected_before = after_bytes.checked_add(inserted_len).ok_or(Error::Limit {
                kind: LimitKind::BeforeBytes,
                observed: usize::MAX,
                maximum: limits.max_output_bytes,
            })?;
            if before_bytes != expected_before {
                return Err(Error::DurablePatch(
                    "remove before/after byte sizes do not match inserted bytes",
                ));
            }
            if root_close > after_bytes {
                return Err(Error::DurablePatch(
                    "remove offset exceeds after artifact size",
                ));
            }
        },
    }
    let end = root_close.checked_add(inserted_len).ok_or(Error::Limit {
        kind: LimitKind::OutputBytes,
        observed: usize::MAX,
        maximum: limits.max_output_bytes,
    })?;
    if direction == Direction::Append && end > limits.max_output_bytes {
        return Err(Error::Limit {
            kind: LimitKind::OutputBytes,
            observed: end,
            maximum: limits.max_output_bytes,
        });
    }
    if inserted_len == 0 && (direction != Direction::Append || before_bytes != after_bytes) {
        return Err(Error::DurablePatch(
            "empty durable patch must use canonical append identity shape",
        ));
    }
    Ok(())
}

fn validate_artifact_splice(
    before: &[u8],
    after: &[u8],
    direction: Direction,
    before_bytes: usize,
    after_bytes: usize,
    root_close: usize,
    inserted: &[u8],
    limits: SemanticLimits,
) -> Result<(), Error> {
    validate_wire_bounds(
        direction,
        before_bytes,
        after_bytes,
        root_close,
        inserted.len(),
        limits,
    )?;
    if direction == Direction::Append && before.len() > limits.max_output_bytes {
        return Err(Error::Limit {
            kind: LimitKind::OutputBytes,
            observed: before.len(),
            maximum: limits.max_output_bytes,
        });
    }
    if after.len() > limits.max_output_bytes {
        return Err(Error::Limit {
            kind: LimitKind::OutputBytes,
            observed: after.len(),
            maximum: limits.max_output_bytes,
        });
    }
    if before.len() != before_bytes || after.len() != after_bytes {
        return Err(Error::DurablePatch(
            "artifact byte sizes do not match durable wire sizes",
        ));
    }
    match direction {
        Direction::Append => {
            let expected_after = before
                .len()
                .checked_add(inserted.len())
                .ok_or(Error::Limit {
                    kind: LimitKind::OutputBytes,
                    observed: usize::MAX,
                    maximum: limits.max_output_bytes,
                })?;
            if after.len() != expected_after {
                return Err(Error::DurablePatch(
                    "append source/target sizes do not match inserted bytes",
                ));
            }
            let end = root_close
                .checked_add(inserted.len())
                .ok_or(Error::DurablePatch("append offset overflows inserted span"))?;
            if root_close > before.len() || end > after.len() {
                return Err(Error::DurablePatch(
                    "append offset is outside the source/target artifacts",
                ));
            }
            if after.get(..root_close) != before.get(..root_close)
                || after.get(root_close..end) != Some(inserted)
                || after.get(end..) != before.get(root_close..)
            {
                return Err(Error::DurablePatch(
                    "append artifacts do not contain the exact inserted span",
                ));
            }
        },
        Direction::Remove => {
            let expected_before = after
                .len()
                .checked_add(inserted.len())
                .ok_or(Error::Limit {
                    kind: LimitKind::OutputBytes,
                    observed: usize::MAX,
                    maximum: limits.max_output_bytes,
                })?;
            if before.len() != expected_before {
                return Err(Error::DurablePatch(
                    "remove source/target sizes do not match inserted bytes",
                ));
            }
            let end = root_close
                .checked_add(inserted.len())
                .ok_or(Error::DurablePatch("remove offset overflows inserted span"))?;
            if root_close > after.len() || end > before.len() {
                return Err(Error::DurablePatch(
                    "remove offset is outside the source/target artifacts",
                ));
            }
            if before.get(root_close..end) != Some(inserted)
                || before.get(..root_close) != after.get(..root_close)
                || before.get(end..) != after.get(root_close..)
            {
                return Err(Error::DurablePatch(
                    "remove artifacts do not contain the exact inserted span",
                ));
            }
        },
    }
    Ok(())
}

fn deterministic_json_len(patch: &DurableTailAppendPatch) -> Result<usize, Error> {
    for value in [patch.after_bytes, patch.before_bytes, patch.root_close] {
        let _ = u64::try_from(value).map_err(|_error| {
            Error::DurablePatch("patch number exceeds the durable integer range")
        })?;
    }
    let encoded_len = patch
        .inserted
        .len()
        .checked_mul(2)
        .ok_or(Error::AllocationFailed {
            resource: "durable inserted hex",
            requested: usize::MAX,
        })?;
    let offset_len = decimal_len_usize(patch.root_close);
    let parts = [
        b"{\"after\":\"".len(),
        patch.after_digest.len(),
        b"\",\"after_bytes\":".len(),
        decimal_len_usize(patch.after_bytes),
        b",\"before\":\"".len(),
        patch.before_digest.len(),
        b"\",\"before_bytes\":".len(),
        decimal_len_usize(patch.before_bytes),
        b",\"direction\":\"".len(),
        patch.direction.as_str().len(),
        b"\",\"format\":\"litchi-rtf-tail-append\",\"inserted_hex\":\"".len(),
        encoded_len,
        b"\",\"offset\":".len(),
        offset_len,
        b",\"selector\":\"body\",\"version\":1}".len(),
    ];
    parts.into_iter().try_fold(0usize, |total, part| {
        total.checked_add(part).ok_or(Error::AllocationFailed {
            resource: "durable JSON size",
            requested: usize::MAX,
        })
    })
}

fn decimal_len_usize(mut value: usize) -> usize {
    let mut length = 1usize;
    while value >= 10 {
        value /= 10;
        length = length.saturating_add(1);
    }
    length
}

#[derive(Debug, Clone, Copy)]
struct SpliceProof {
    root_close: usize,
    ends_with_par: bool,
}

fn prove_splice_source(source: &Package, selector: TailSelector) -> Result<SpliceProof, Error> {
    if selector != TailSelector::Body {
        return Err(Error::Refused(Refusal::new(
            RefusalKind::UnsupportedTarget,
            "tail_append",
            "the selector is not supported by this edit family",
        )));
    }
    let provenance = source.model().parse_provenance();
    if !provenance.syntax_valid || !provenance.root_valid || !provenance.document_valid {
        return Err(Error::unsupported_source(
            "parser provenance does not prove a complete document",
        ));
    }
    if source.model().protection().is_protected() {
        return Err(Error::Refused(Refusal::new(
            RefusalKind::Protected,
            "tail_append",
            "document protection refuses changed publication",
        )));
    }
    if source.model().unknown_syntax_markers() != 0 {
        return Err(Error::unsupported_source(
            "dropped unknown syntax prevents a safe root-tail splice",
        ));
    }
    reject_active_or_external_surfaces(source)?;
    if crate::compressed::is_compressed_rtf(source.source_bytes().ok_or(
        Error::unsupported_operation("snapshot has no exact RTF source"),
    )?) {
        return Err(Error::unsupported_source(
            "compressed RTF is not an exact uncompressed splice source",
        ));
    }
    let bytes = source.source_bytes().ok_or(Error::unsupported_operation(
        "snapshot has no exact RTF source",
    ))?;
    if !bytes.is_ascii() {
        return Err(Error::unsupported_source(
            "non-ASCII transport requires an explicit code-page writer",
        ));
    }
    let root_close = scan_root_close(bytes)?;
    let ends_with_par = ends_with_parameterless_par(source);
    Ok(SpliceProof {
        root_close,
        ends_with_par,
    })
}

fn reject_active_or_external_surfaces(source: &Package) -> Result<(), Error> {
    let model = source.model();
    if !model.external_references().is_empty()
        || model.mail_merge().is_some()
        || model.xsl_transform().is_some()
        || model.xsl_transform_usage().is_requested()
    {
        return Err(Error::unsupported_source(
            "external document or transformation content is present",
        ));
    }
    if !model.form_fields().is_empty() || !model.objects().is_empty() {
        return Err(Error::unsupported_source(
            "active form or embedded-object content is present",
        ));
    }
    if model
        .pictures()
        .iter()
        .any(|picture| picture.image_type == crate::ImageType::Unknown)
    {
        return Err(Error::unsupported_source(
            "an unsupported picture payload is present",
        ));
    }
    for safety in model.field_safety() {
        match safety {
            crate::validation::FieldSafety::Neutral => {},
            crate::validation::FieldSafety::External
            | crate::validation::FieldSafety::ExternalUnknown
            | crate::validation::FieldSafety::ExternalAndActive
            | crate::validation::FieldSafety::ExternalAndActiveUnknown => {
                return Err(Error::unsupported_source(
                    "external field content is present",
                ));
            },
            crate::validation::FieldSafety::Active
            | crate::validation::FieldSafety::ActiveUnknown => {
                return Err(Error::unsupported_source("active field content is present"));
            },
        }
    }
    Ok(())
}

fn scan_root_close(bytes: &[u8]) -> Result<usize, Error> {
    if bytes.first() != Some(&b'{') {
        return Err(Error::unsupported_source(
            "source does not begin with one root group",
        ));
    }
    let mut depth = 1usize;
    let mut index = 1usize;
    while index < bytes.len() {
        let byte = bytes.get(index).copied().ok_or(Error::unsupported_source(
            "root scanner offset is outside the source",
        ))?;
        match byte {
            b'{' => {
                depth = depth
                    .checked_add(1)
                    .ok_or(Error::unsupported_source("root nesting depth overflowed"))?;
                index += 1;
            },
            b'}' => {
                depth = depth.checked_sub(1).ok_or(Error::unsupported_source(
                    "root group closed before its opening brace",
                ))?;
                if depth == 0 {
                    if bytes
                        .get(index + 1..)
                        .ok_or(Error::unsupported_source(
                            "root scanner suffix is outside the source",
                        ))?
                        .iter()
                        .any(|byte| !byte.is_ascii_whitespace())
                    {
                        return Err(Error::unsupported_source(
                            "trailing non-whitespace follows the root group",
                        ));
                    }
                    return Ok(index);
                }
                index += 1;
            },
            b'\\' => {
                index = skip_control(bytes, index)?;
            },
            _ => index += 1,
        }
    }
    Err(Error::unsupported_source(
        "root group does not have one exact closing brace",
    ))
}

fn skip_control(bytes: &[u8], slash: usize) -> Result<usize, Error> {
    let next = slash
        .checked_add(1)
        .ok_or(Error::unsupported_source("control offset overflowed"))?;
    let Some(&byte) = bytes.get(next) else {
        return Err(Error::unsupported_source(
            "source ends after a control slash",
        ));
    };
    if matches!(byte, b'\\' | b'{' | b'}' | b'*') {
        return Ok(next + 1);
    }
    if byte == b'\'' {
        let end = next
            .checked_add(3)
            .ok_or(Error::unsupported_source("hex control offset overflowed"))?;
        let digits = bytes
            .get(next + 1..next + 3)
            .ok_or(Error::unsupported_source("malformed hexadecimal control"))?;
        if !digits.first().is_some_and(u8::is_ascii_hexdigit)
            || !digits.get(1).is_some_and(u8::is_ascii_hexdigit)
        {
            return Err(Error::unsupported_source("malformed hexadecimal control"));
        }
        return Ok(end);
    }
    if !byte.is_ascii_alphabetic() {
        return Ok(next + 1);
    }
    let mut end = next;
    while end < bytes.len() && bytes.get(end).is_some_and(u8::is_ascii_alphabetic) {
        end += 1;
    }
    let name = bytes
        .get(next..end)
        .ok_or(Error::unsupported_source("control word span is invalid"))?;
    if name.eq_ignore_ascii_case(b"bin") {
        return Err(Error::unsupported_source(
            "binary destinations are outside the exact ASCII splice proof",
        ));
    }
    if end < bytes.len()
        && bytes
            .get(end)
            .is_some_and(|byte| matches!(byte, b'-' | b'+' | b'0'..=b'9'))
    {
        end += 1;
        while end < bytes.len() && bytes.get(end).is_some_and(u8::is_ascii_digit) {
            end += 1;
        }
    }
    if end < bytes.len() && bytes.get(end) == Some(&b' ') {
        end += 1;
    }
    Ok(end)
}

fn ends_with_parameterless_par(source: &Package) -> bool {
    matches!(
        source.body().inlines().last(),
        Some(crate::text::Inline::Break(crate::text::Break::Paragraph))
    )
}

fn validate_plain_text(text: &str) -> Result<(), Error> {
    for character in text.chars() {
        if character == '\n' || character == '\r' {
            return Err(Error::InvalidText(
                "paragraph separators must be expressed by the paragraph API",
            ));
        }
        if character != '\t' && character.is_control() {
            return Err(Error::InvalidText(
                "control characters other than tab are not plain text",
            ));
        }
    }
    Ok(())
}

fn encode_inserted(
    source: &Package,
    paragraphs: &[OwnedParagraph],
    ends_with_par: bool,
    estimated: usize,
) -> Result<Vec<u8>, Error> {
    encode_inserted_for_body(
        !source.body().is_empty(),
        paragraphs,
        ends_with_par,
        estimated,
    )
}

fn encode_inserted_for_body(
    body_nonempty: bool,
    paragraphs: &[OwnedParagraph],
    ends_with_par: bool,
    estimated: usize,
) -> Result<Vec<u8>, Error> {
    let separator = body_nonempty && !ends_with_par;
    let mut encoded = Vec::new();
    encoded
        .try_reserve(estimated)
        .map_err(|_error| Error::AllocationFailed {
            resource: "tail encoded payload",
            requested: estimated,
        })?;
    if separator {
        encoded.extend_from_slice(b"\\par ");
    }
    for paragraph in paragraphs {
        // `plain` resets character properties and `pard` resets paragraph
        // properties inherited from the prior root-body tail.  Both are
        // authored controls inside the changed closure; every pre-existing
        // byte remains outside that closure.
        encoded.extend_from_slice(b"\\pard ");
        for run in &paragraph.runs {
            // The source may have left a non-default `\\ucN` state at the
            // root tail.  Reset it locally so each `\\uN?` fallback is
            // consumed exactly once during candidate readback.
            encoded.extend_from_slice(b"{\\plain \\uc1 ");
            encode_plain_text(run, &mut encoded);
            encoded.push(b'}');
        }
        encoded.extend_from_slice(b"\\par ");
    }
    Ok(encoded)
}

fn encoded_inserted_len(
    source: &Package,
    paragraphs: &[OwnedParagraph],
    ends_with_par: bool,
) -> Result<usize, Error> {
    encoded_inserted_len_for_body(!source.body().is_empty(), paragraphs, ends_with_par)
}

fn encoded_inserted_len_for_body(
    body_nonempty: bool,
    paragraphs: &[OwnedParagraph],
    ends_with_par: bool,
) -> Result<usize, Error> {
    let separator = body_nonempty && !ends_with_par;
    let mut length =
        usize::from(separator)
            .checked_mul(b"\\par ".len())
            .ok_or(Error::AllocationFailed {
                resource: "tail encoded estimate",
                requested: usize::MAX,
            })?;
    for paragraph in paragraphs {
        length = length
            .checked_add(b"\\pard ".len())
            .and_then(|value| value.checked_add(b"\\par ".len()))
            .ok_or(Error::AllocationFailed {
                resource: "tail encoded estimate",
                requested: usize::MAX,
            })?;
        for run in &paragraph.runs {
            let run_length = encoded_plain_text_len(run)?;
            length = length
                .checked_add(b"{\\plain \\uc1 ".len())
                .and_then(|value| value.checked_add(1))
                .and_then(|value| value.checked_add(run_length))
                .ok_or(Error::AllocationFailed {
                    resource: "tail encoded estimate",
                    requested: usize::MAX,
                })?;
        }
    }
    Ok(length)
}

fn encoded_plain_text_len(text: &str) -> Result<usize, Error> {
    let mut length = 0usize;
    for character in text.chars() {
        let addition = match character {
            '\\' | '{' | '}' => 2,
            '\t' => b"\\tab ".len(),
            character if character.is_ascii() && !character.is_control() => 1,
            character => {
                let mut units = [0u16; 2];
                let count = character.encode_utf16(&mut units).len();
                let mut units_length = 0usize;
                for unit in units.get(..count).into_iter().flatten().copied() {
                    let signed = i32::from(i16::from_ne_bytes(unit.to_ne_bytes()));
                    units_length = units_length
                        .checked_add(3)
                        .and_then(|value| value.checked_add(decimal_len(signed)))
                        .ok_or(Error::AllocationFailed {
                            resource: "tail encoded estimate",
                            requested: usize::MAX,
                        })?;
                }
                units_length
            },
        };
        length = length
            .checked_add(addition)
            .ok_or(Error::AllocationFailed {
                resource: "tail encoded estimate",
                requested: usize::MAX,
            })?;
    }
    Ok(length)
}

fn encode_plain_text(text: &str, output: &mut Vec<u8>) {
    for character in text.chars() {
        match character {
            '\\' => output.extend_from_slice(b"\\\\"),
            '{' => output.extend_from_slice(b"\\{"),
            '}' => output.extend_from_slice(b"\\}"),
            '\t' => output.extend_from_slice(b"\\tab "),
            character if character.is_ascii() && !character.is_control() => {
                output.push(character as u8)
            },
            character => {
                let mut units = [0u16; 2];
                let count = character.encode_utf16(&mut units).len();
                for unit in units.get(..count).into_iter().flatten().copied() {
                    let signed = i32::from(i16::from_ne_bytes(unit.to_ne_bytes()));
                    output.extend_from_slice(b"\\u");
                    append_decimal(output, signed);
                }
                // Keep surrogate controls adjacent. With `\\uc1`, the
                // parser consumes one fallback character for every adjacent
                // `\\uN` control. A fallback between a pair would make each
                // half an invalid lone surrogate during readback.
                for _ in 0..count {
                    output.push(b'?');
                }
            },
        }
    }
}

fn append_decimal(output: &mut Vec<u8>, value: i32) {
    let mut buffer = [0u8; 12];
    let mut index = buffer.len();
    let negative = value < 0;
    let mut magnitude = i64::from(value).unsigned_abs();
    loop {
        index -= 1;
        if let Some(slot) = buffer.get_mut(index) {
            *slot = b'0' + u8::try_from(magnitude % 10).unwrap_or(0);
        }
        magnitude /= 10;
        if magnitude == 0 {
            break;
        }
    }
    if negative {
        index -= 1;
        if let Some(slot) = buffer.get_mut(index) {
            *slot = b'-';
        }
    }
    if let Some(encoded) = buffer.get(index..) {
        output.extend_from_slice(encoded);
    }
}

fn decimal_len(value: i32) -> usize {
    let mut magnitude = i64::from(value).unsigned_abs();
    let mut length = 1usize;
    while magnitude >= 10 {
        magnitude /= 10;
        length = length.saturating_add(1);
    }
    if value < 0 {
        length.saturating_add(1)
    } else {
        length
    }
}

fn verify_candidate(
    source: &Package,
    candidate: &Package,
    paragraphs: &[OwnedParagraph],
    ends_with_par: bool,
) -> Result<(), Error> {
    let source_text = source.text();
    let candidate_text = candidate.text();
    if !candidate_text.starts_with(source_text) {
        return Err(Error::unsupported_source(
            "candidate changed the pre-existing body text",
        ));
    }
    let suffix = candidate_text
        .get(source_text.len()..)
        .ok_or(Error::unsupported_source(
            "candidate text prefix is invalid",
        ))?;
    let separator = !source.body().is_empty() && !ends_with_par;
    let mut expected_len = usize::from(separator);
    for paragraph in paragraphs {
        for run in &paragraph.runs {
            expected_len = expected_len
                .checked_add(run.len())
                .ok_or(Error::AllocationFailed {
                    resource: "tail readback text",
                    requested: usize::MAX,
                })?;
        }
        expected_len = expected_len.checked_add(1).ok_or(Error::AllocationFailed {
            resource: "tail readback text",
            requested: usize::MAX,
        })?;
    }
    let mut expected = String::new();
    expected
        .try_reserve(expected_len)
        .map_err(|_error| Error::AllocationFailed {
            resource: "tail readback text",
            requested: expected_len,
        })?;
    if separator {
        expected.push('\n');
    }
    for paragraph in paragraphs {
        for run in &paragraph.runs {
            expected.push_str(run);
        }
        expected.push('\n');
    }
    if suffix != expected {
        return Err(Error::unsupported_source(
            "candidate plain paragraphs did not survive RTF readback",
        ));
    }
    Ok(())
}

fn write_sequential<W: Write>(sink: &mut W, bytes: &[u8]) -> Result<usize, Error> {
    let mut written = 0usize;
    while written < bytes.len() {
        let remaining = bytes.get(written..).ok_or(Error::Sink {
            kind: io::ErrorKind::InvalidData,
            written,
        })?;
        match sink.write(remaining) {
            Ok(0) => {
                return Err(Error::Sink {
                    kind: io::ErrorKind::WriteZero,
                    written,
                });
            },
            Ok(count) => {
                if count > bytes.len().saturating_sub(written) {
                    return Err(Error::Sink {
                        kind: io::ErrorKind::InvalidData,
                        written,
                    });
                }
                written = written.saturating_add(count);
            },
            Err(error) if error.kind() == io::ErrorKind::Interrupted => {},
            Err(error) => {
                return Err(Error::Sink {
                    kind: error.kind(),
                    written,
                });
            },
        }
    }
    sink.flush().map_err(|error| Error::Sink {
        kind: error.kind(),
        written,
    })?;
    Ok(written)
}

fn hex_encode(bytes: &[u8]) -> Result<String, Error> {
    let encoded_len = bytes.len().checked_mul(2).ok_or(Error::AllocationFailed {
        resource: "durable inserted hex",
        requested: usize::MAX,
    })?;
    let mut text = String::new();
    text.try_reserve(encoded_len)
        .map_err(|_error| Error::AllocationFailed {
            resource: "durable inserted hex",
            requested: encoded_len,
        })?;
    const HEX: &[u8; 16] = b"0123456789abcdef";
    for byte in bytes {
        let high = HEX
            .get((byte >> 4) as usize)
            .copied()
            .ok_or(Error::DurablePatch("hex encoder index overflowed"))?;
        let low = HEX
            .get((byte & 0x0f) as usize)
            .copied()
            .ok_or(Error::DurablePatch("hex encoder index overflowed"))?;
        text.push(high as char);
        text.push(low as char);
    }
    Ok(text)
}

fn hex_decode(text: &str, max_bytes: usize) -> Result<Vec<u8>, Error> {
    if text.len() % 2 != 0 {
        return Err(Error::DurablePatch("inserted hex has odd length"));
    }
    let observed = text.len() / 2;
    if observed > max_bytes {
        return Err(Error::Limit {
            kind: LimitKind::InsertedBytes,
            observed,
            maximum: max_bytes,
        });
    }
    let mut bytes = Vec::new();
    bytes
        .try_reserve(observed)
        .map_err(|_error| Error::AllocationFailed {
            resource: "durable inserted bytes",
            requested: observed,
        })?;
    let raw = text.as_bytes();
    for pair in raw.as_chunks::<2>().0.iter() {
        let high = hex_value(
            *pair
                .first()
                .ok_or(Error::DurablePatch("invalid inserted hex"))?,
        )
        .ok_or(Error::DurablePatch("invalid inserted hex"))?;
        let low = hex_value(
            *pair
                .get(1)
                .ok_or(Error::DurablePatch("invalid inserted hex"))?,
        )
        .ok_or(Error::DurablePatch("invalid inserted hex"))?;
        bytes.push((high << 4) | low);
    }
    Ok(bytes)
}

const fn hex_value(byte: u8) -> Option<u8> {
    match byte {
        b'0'..=b'9' => Some(byte - b'0'),
        b'a'..=b'f' => Some(byte - b'a' + 10),
        b'A'..=b'F' => Some(byte - b'A' + 10),
        _ => None,
    }
}

fn string_field<'a>(object: &'a Map<String, Value>, key: &str) -> Result<&'a str, Error> {
    object
        .get(key)
        .and_then(Value::as_str)
        .ok_or(Error::DurablePatch("patch string field is missing"))
}

fn number(value: usize) -> Result<Value, Error> {
    let value = u64::try_from(value)
        .map_err(|_error| Error::DurablePatch("patch offset exceeds the durable integer range"))?;
    Ok(Value::from(value))
}

fn number_field(object: &Map<String, Value>, key: &str) -> Result<usize, Error> {
    let value = object
        .get(key)
        .and_then(Value::as_u64)
        .ok_or(Error::DurablePatch("patch offset is missing"))?;
    usize::try_from(value).map_err(|_error| Error::DurablePatch("patch offset overflows usize"))
}

fn validate_digest(value: &str) -> Result<(), Error> {
    if value.len() != 64
        || !value
            .as_bytes()
            .iter()
            .all(|byte| byte.is_ascii_digit() || matches!(byte, b'a'..=b'f'))
    {
        return Err(Error::DurablePatch(
            "patch digest is not canonical lowercase SHA-256 hex",
        ));
    }
    Ok(())
}
