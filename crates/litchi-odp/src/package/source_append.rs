//! Bounded append of one plain slide to a [`SourceBackedPresentation`]
//! (record 0358a; the shape of DOCX record 0355a §1).
//!
//! [`SourceBackedPresentation::edit_tail_append`] starts an [`Edit`] holding
//! only the bounded title and body; [`Edit::prepare`] scans the source once
//! and returns a [`Plan`] with its [`SourceProof`]; [`Plan::commit`] (or
//! [`Edit::commit`]) returns a [`Commit`] holding the new presentation, the
//! reversible [`Patch`] and the [`Diagnostics`]. Nothing is written until the
//! committed presentation's
//! [`write_to`](SourceBackedPresentation::write_to) or
//! [`save`](SourceBackedPresentation::save), which replays `content.xml` with
//! the page inserted into the preserved ZIP layout and copies every other
//! member raw. The commit never constructs a complete `content.xml` or
//! candidate package; reading or editing the committed presentation
//! materializes its output once.
//!
//! ```no_run
//! use litchi_odp::SourceBackedPresentation;
//!
//! # fn main() -> Result<(), Box<dyn std::error::Error>> {
//! let presentation = SourceBackedPresentation::open("deck.odp")?;
//! let commit = presentation.edit_tail_append("Title", "Body")?.commit()?;
//! commit.package().save("deck-appended.odp")?;
//! # Ok(())
//! # }
//! ```

use std::{fmt, sync::Arc};

use super::source::SourceBackedPresentation;
pub use litchi_odf_common::LimitKind;

use litchi_core::{
    CancellationToken, Error as CoreError, ExecutionError, Resource, Result, SourceVersion,
};
use litchi_odf_common::{
    Error as OdfError,
    constants::ODF_CONTENT,
    core::source_publication::{
        SourceContentInsertionCallbackError, SourceContentInsertionError,
        SourceContentInsertionPlan, SourceContentScanError, scan_source_content,
    },
    core::{
        AuthoredXmlFragment, SourceBackedPackage, SourceContentPublicationError,
        SourceContentPublicationOptions, SourceContentPublicationProgress, SourceMemberReaderError,
        XmlStreamLimits,
        private::{BindingTracker, XmlStreamEvent},
    },
};
use quick_xml::{
    XmlVersion,
    events::{BytesStart, Event},
    name::{Namespace, ResolveResult},
};

const OFFICE_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const DRAW_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:drawing:1.0";
const GENERATED_DOCUMENT_PREFIX: &[u8] = br#"<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:presentation="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"><office:body><office:presentation>"#;
const GENERATED_DOCUMENT_SUFFIX: &[u8] =
    b"</office:presentation></office:body></office:document-content>";
const GENERATED_READBACK_TEXT_GEOMETRIC_MULTIPLIER: usize = 2;
const GENERATED_READBACK_TEXT_PARALLEL_STRINGS: usize = 3;
const GENERATED_FRAGMENT_GROWTH_MULTIPLIER: usize = 2;
const GENERATED_READBACK_PARSER_BUFFER_MULTIPLIER: usize = 1;
const GENERATED_READBACK_DECODED_TOKEN_MULTIPLIER: usize = 2;
const GENERATED_READBACK_FIXED_STATE_BYTES: usize = 64 * 1024;
const MAX_TEXT_SPACE_COUNT: usize = 1_000_000;
const MAX_SLIDES: usize = 65_536;
const MAX_TEXT_BYTES: usize = 16 * 1024 * 1024;
const MAX_FRAGMENT_BYTES: usize = 512 * 1024;
const MAX_CONTENT_XML_BYTES: u64 = 256 * 1024 * 1024;

/// Finite bounds for a source-backed one-slide append.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Limits {
    /// Maximum combined UTF-8 bytes in the submitted title and body.
    pub max_text_bytes: usize,
    /// Maximum generated page bytes retained by the plan.
    pub max_fragment_bytes: usize,
    /// Maximum direct pages accepted in the source presentation.
    pub max_slides: usize,
    /// Maximum decoded `content.xml` bytes accepted by the scanner.
    pub max_content_xml_bytes: u64,
}

impl Limits {
    /// Create explicit finite append bounds.
    #[must_use]
    pub const fn new(
        max_text_bytes: usize,
        max_fragment_bytes: usize,
        max_slides: usize,
        max_content_xml_bytes: u64,
    ) -> Self {
        Self {
            max_text_bytes,
            max_fragment_bytes,
            max_slides,
            max_content_xml_bytes,
        }
    }
}

impl Default for Limits {
    fn default() -> Self {
        Self::new(
            MAX_TEXT_BYTES,
            MAX_FRAGMENT_BYTES,
            MAX_SLIDES,
            MAX_CONTENT_XML_BYTES,
        )
    }
}

/// Failure from the format-owned source append proof or publication handoff.
#[derive(Debug)]
#[non_exhaustive]
pub enum Error {
    /// The source or authored request is not accepted by the ODP proof.
    Invalid(CoreError),
    /// A verified source member reader failed while scanning the source.
    SourceMember(SourceMemberReaderError<CoreError>),
    /// The common insertion owner rejected the bounded plan before output.
    Plan(SourceContentPublicationError),
    /// The common insertion owner reported a typed sink/source failure after
    /// or during publication.
    Publication(SourceContentInsertionError),
    /// The positional source changed after the append captured its proof.
    SourceChanged {
        /// Source identity and revision captured before the operation.
        expected: SourceVersion,
        /// Source identity and revision observed during the operation.
        observed: SourceVersion,
        /// Sink progress observed when the source became stale.
        progress: SourceContentPublicationProgress,
        /// The typed reader or publication diagnostic retained below the
        /// freshness classification.
        source: Option<Box<SourceChangedSource>>,
    },
    /// A patch was applied to a presentation it was not made on.
    PatchConflict {
        /// The replay operation that detected the conflict.
        operation: &'static str,
        /// Why the base does not match.
        reason: &'static str,
    },
}

/// Typed diagnostic retained by [`Error::SourceChanged`].
#[derive(Debug)]
#[non_exhaustive]
pub enum SourceChangedSource {
    /// The common core freshness error observed before source publication.
    Core(CoreError),
    /// The verified source-member reader that observed the stale source.
    SourceMember(SourceMemberReaderError<CoreError>),
    /// The common source-content publication failure that observed it.
    Publication(SourceContentPublicationError),
    /// The common source-content insertion failure that observed it, keeping
    /// replay and callback diagnostics on the effective/read path.
    Insertion(SourceContentInsertionError),
}

impl fmt::Display for SourceChangedSource {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Core(error) => error.fmt(formatter),
            Self::SourceMember(error) => error.fmt(formatter),
            Self::Publication(error) => error.fmt(formatter),
            Self::Insertion(error) => error.fmt(formatter),
        }
    }
}

impl std::error::Error for SourceChangedSource {
    fn source(&self) -> Option<&(dyn std::error::Error + 'static)> {
        match self {
            Self::Core(error) => Some(error),
            Self::SourceMember(error) => Some(error),
            Self::Publication(error) => Some(error),
            Self::Insertion(error) => Some(error),
        }
    }
}

impl fmt::Display for Error {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Invalid(error) => error.fmt(formatter),
            Self::SourceMember(error) => error.fmt(formatter),
            Self::Plan(error) => error.fmt(formatter),
            Self::Publication(error) => error.fmt(formatter),
            Self::SourceChanged {
                expected,
                observed,
                progress,
                ..
            } => write!(
                formatter,
                "ODP source changed (expected {expected:?}, observed {observed:?}, {progress:?})"
            ),
            Self::PatchConflict { operation, reason } => {
                write!(formatter, "ODP patch conflict in {operation}: {reason}")
            },
        }
    }
}

impl std::error::Error for Error {
    fn source(&self) -> Option<&(dyn std::error::Error + 'static)> {
        match self {
            Self::Invalid(error) => Some(error),
            Self::SourceMember(error) => Some(error),
            Self::Plan(error) => Some(error),
            Self::Publication(error) => Some(error),
            Self::SourceChanged {
                source: Some(source),
                ..
            } => Some(source.as_ref()),
            Self::SourceChanged { source: None, .. } | Self::PatchConflict { .. } => None,
        }
    }
}

impl Error {
    /// Return sink progress when the common publication layer has observed it.
    #[must_use]
    pub const fn progress(&self) -> Option<SourceContentPublicationProgress> {
        match self {
            Self::Plan(error) => Some(error.progress()),
            Self::Publication(error) => Some(error.progress()),
            Self::Invalid(_) | Self::SourceMember(_) | Self::PatchConflict { .. } => None,
            Self::SourceChanged { progress, .. } => Some(*progress),
        }
    }
}

fn invalid(message: impl Into<String>) -> Error {
    Error::Invalid(CoreError::InvalidFormat(message.into()))
}

fn source_error(error: CoreError) -> Error {
    match error {
        error @ CoreError::SourceChanged { expected, observed } => Error::SourceChanged {
            expected,
            observed,
            progress: SourceContentPublicationProgress::Untouched,
            source: Some(Box::new(SourceChangedSource::Core(error))),
        },
        other => Error::Invalid(other),
    }
}

fn odf_error(error: OdfError) -> Error {
    match error {
        OdfError::Core(error) => source_error(error),
        OdfError::SourceContentPublication(error) => publication_error(*error),
        OdfError::SourceContentInsertion(error) => insertion_error(*error),
        other => Error::Invalid(CoreError::from(other)),
    }
}

fn publication_error(error: SourceContentPublicationError) -> Error {
    let progress = error.progress();
    match error {
        error @ SourceContentPublicationError::SourceChanged {
            expected, observed, ..
        }
        | error @ SourceContentPublicationError::Core(CoreError::SourceChanged {
            expected,
            observed,
        })
        | error @ SourceContentPublicationError::Source {
            source: CoreError::SourceChanged { expected, observed },
            ..
        } => Error::SourceChanged {
            expected,
            observed,
            progress,
            source: Some(Box::new(SourceChangedSource::Publication(error))),
        },
        other => Error::Plan(other),
    }
}

fn insertion_error(error: SourceContentInsertionError) -> Error {
    let progress = error.progress();
    let Some((expected, observed)) = insertion_source_changed(&error) else {
        return Error::Publication(error);
    };
    Error::SourceChanged {
        expected,
        observed,
        progress,
        source: Some(Box::new(SourceChangedSource::Insertion(error))),
    }
}

fn insertion_source_changed(
    error: &SourceContentInsertionError,
) -> Option<(SourceVersion, SourceVersion)> {
    match error {
        SourceContentInsertionError::Publication { source, .. } => {
            publication_source_changed(source)
        },
        SourceContentInsertionError::Replay { source, .. } => replay_source_changed(source),
        SourceContentInsertionError::Transport {
            publication,
            replay,
            ..
        } => publication_source_changed(publication).or_else(|| replay_source_changed(replay)),
        _ => None,
    }
}

fn publication_source_changed(
    error: &SourceContentPublicationError,
) -> Option<(SourceVersion, SourceVersion)> {
    match error {
        SourceContentPublicationError::SourceChanged {
            expected, observed, ..
        }
        | SourceContentPublicationError::Core(CoreError::SourceChanged { expected, observed })
        | SourceContentPublicationError::Source {
            source: CoreError::SourceChanged { expected, observed },
            ..
        } => Some((*expected, *observed)),
        _ => None,
    }
}

fn replay_source_changed(
    error: &soapberry_zip::ReplayPublicationError<SourceContentInsertionCallbackError>,
) -> Option<(SourceVersion, SourceVersion)> {
    use soapberry_zip::ReplayPublicationError;

    match error {
        ReplayPublicationError::Archive { callback_error, .. }
        | ReplayPublicationError::Limit { callback_error, .. }
        | ReplayPublicationError::Sink { callback_error, .. } => {
            callback_error.as_ref().and_then(callback_source_changed)
        },
        ReplayPublicationError::Callback { source, .. } => callback_source_changed(source),
        ReplayPublicationError::NonDeterministic { .. } => None,
        _ => None,
    }
}

fn callback_source_changed(
    error: &SourceContentInsertionCallbackError,
) -> Option<(SourceVersion, SourceVersion)> {
    match error {
        SourceContentInsertionCallbackError::Member(error) => match error.as_ref() {
            SourceMemberReaderError::Core {
                error,
                callback_error,
            } => core_source_changed(error)
                .or_else(|| callback_error.as_ref().and_then(callback_source_changed)),
            SourceMemberReaderError::Callback(error) => callback_source_changed(error),
            _ => None,
        },
        SourceContentInsertionCallbackError::Io(_)
        | SourceContentInsertionCallbackError::Cancelled
        | SourceContentInsertionCallbackError::Execution(_)
        | SourceContentInsertionCallbackError::Verification { .. } => None,
        _ => None,
    }
}

fn core_source_changed(error: &CoreError) -> Option<(SourceVersion, SourceVersion)> {
    match error {
        CoreError::SourceChanged { expected, observed } => Some((*expected, *observed)),
        _ => None,
    }
}

fn source_member_error(error: SourceMemberReaderError<CoreError>) -> Error {
    match error {
        error @ SourceMemberReaderError::Core {
            error: CoreError::SourceChanged { expected, observed },
            ..
        }
        | error @ SourceMemberReaderError::Callback(CoreError::SourceChanged {
            expected,
            observed,
        }) => Error::SourceChanged {
            expected,
            observed,
            progress: SourceContentPublicationProgress::Untouched,
            source: Some(Box::new(SourceChangedSource::SourceMember(error))),
        },
        other => Error::SourceMember(other),
    }
}

fn check_plan_options(options: &SourceContentPublicationOptions) -> std::result::Result<(), Error> {
    if options
        .cancellation()
        .is_some_and(CancellationToken::is_cancelled)
    {
        return Err(Error::Plan(SourceContentPublicationError::Cancelled {
            progress: SourceContentPublicationProgress::Untouched,
        }));
    }
    if let Some(execution) = options.execution_context() {
        execution
            .check()
            .map_err(map_execution_error)
            .map_err(Error::Plan)?;
    }
    Ok(())
}

fn map_execution_error(error: ExecutionError) -> SourceContentPublicationError {
    match error {
        ExecutionError::Cancelled => SourceContentPublicationError::Cancelled {
            progress: SourceContentPublicationProgress::Untouched,
        },
        source => SourceContentPublicationError::Execution {
            progress: SourceContentPublicationProgress::Untouched,
            source,
        },
    }
}

fn reserve_fragment_memory(
    options: &SourceContentPublicationOptions,
    maximum: usize,
) -> std::result::Result<Option<litchi_core::Reservation>, Error> {
    let Some(execution) = options.execution_context() else {
        return Ok(None);
    };
    let amount = u64::try_from(maximum).map_err(|_| {
        invalid("ODP generated page memory bound exceeds the execution accounting range")
    })?;
    execution
        .reserve(Resource::Memory, amount)
        .map(Some)
        .map_err(map_execution_error)
        .map_err(Error::Plan)
}

/// Bound the format-owned page generation and parser readback window.
///
/// The submitted strings remain live through generated-page readback and are charged by
/// their actual capacity, including spare caller-supplied capacity. The established
/// parser can have two decoded strings for each input channel live at once:
/// the current paragraph plus its retained shape, or the retained body shape
/// plus the slide body projection; the parser's decoded current token is the
/// third.  Each String is charged with geometric capacity slack.  The
/// generated fragment can overlap its old allocation during growth; while
/// readback runs, that allocation overlaps the wrapped document, one parser
/// token buffer, and the decoded current token.  The fixed term covers the
/// one-slide model vectors, namespace stack, and parser bookkeeping; all
/// source-sized buffers are charged separately.  The document wrapper is also
/// larger than the compact wrapper used by the later AuthoredXmlFragment
/// audit.
fn generated_readback_memory_bound(
    title: &str,
    body: &str,
    retained_text_capacity: usize,
    max_fragment_bytes: usize,
) -> Result<usize> {
    let submitted_text_bytes = title
        .len()
        .checked_add(body.len())
        .ok_or_else(|| CoreError::InvalidFormat("ODP submitted text size overflow".to_string()))?;
    let parser_text_bytes = submitted_text_bytes
        .checked_mul(GENERATED_READBACK_TEXT_GEOMETRIC_MULTIPLIER)
        .and_then(|value| value.checked_mul(GENERATED_READBACK_TEXT_PARALLEL_STRINGS))
        .ok_or_else(|| CoreError::InvalidFormat("ODP parser text bound overflow".to_string()))?;
    let text_memory = retained_text_capacity
        .checked_add(parser_text_bytes)
        .ok_or_else(|| CoreError::InvalidFormat("ODP text memory bound overflow".to_string()))?;
    let wrapper_bytes = GENERATED_DOCUMENT_PREFIX
        .len()
        .checked_add(GENERATED_DOCUMENT_SUFFIX.len())
        .ok_or_else(|| {
            CoreError::InvalidFormat("ODP generated document wrapper size overflow".to_string())
        })?;
    let fragment_growth = max_fragment_bytes
        .checked_mul(GENERATED_FRAGMENT_GROWTH_MULTIPLIER)
        .ok_or_else(|| {
            CoreError::InvalidFormat("ODP fragment growth bound overflow".to_string())
        })?;
    let readback_document = max_fragment_bytes
        .checked_add(wrapper_bytes)
        .ok_or_else(|| {
            CoreError::InvalidFormat("ODP readback document bound overflow".to_string())
        })?;
    let parser_buffer = max_fragment_bytes
        .checked_mul(GENERATED_READBACK_PARSER_BUFFER_MULTIPLIER)
        .ok_or_else(|| CoreError::InvalidFormat("ODP parser buffer bound overflow".to_string()))?;
    let decoded_token = max_fragment_bytes
        .checked_mul(GENERATED_READBACK_DECODED_TOKEN_MULTIPLIER)
        .ok_or_else(|| CoreError::InvalidFormat("ODP decoded token bound overflow".to_string()))?;
    text_memory
        .checked_add(fragment_growth)
        .and_then(|value| value.checked_add(readback_document))
        .and_then(|value| value.checked_add(parser_buffer))
        .and_then(|value| value.checked_add(decoded_token))
        .and_then(|value| value.checked_add(GENERATED_READBACK_FIXED_STATE_BYTES))
        .ok_or_else(|| {
            CoreError::InvalidFormat("ODP generated page memory bound overflow".to_string())
        })
}

/// Facts established by the bounded source scan before ZIP preparation.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SourceProof {
    source_version: SourceVersion,
    content_bytes: u64,
    insert_at: u64,
    slide_count: usize,
    page_name: String,
}

impl SourceProof {
    /// Source revision captured by the proof.
    #[must_use]
    pub const fn source_version(&self) -> SourceVersion {
        self.source_version
    }

    /// Verified decoded `content.xml` length.
    #[must_use]
    pub const fn content_bytes(&self) -> u64 {
        self.content_bytes
    }

    /// Exact decoded insertion offset, immediately after the final page.
    #[must_use]
    pub const fn insert_at(&self) -> u64 {
        self.insert_at
    }

    /// Number of direct pages in the source presentation.
    #[must_use]
    pub const fn slide_count(&self) -> usize {
        self.slide_count
    }

    /// Generated page name checked against every existing page during the
    /// candidate preflight pass.
    #[must_use]
    pub fn page_name(&self) -> &str {
        &self.page_name
    }
}

/// The result of a tail-append `commit`: the new source-backed
/// presentation, the patch and the diagnostics.
pub type Commit = litchi_core::edit::Commit<SourceBackedPresentation, Patch, Diagnostics>;

/// What a tail-append commit changed.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Diagnostics {
    changed: bool,
    proof: Option<SourceProof>,
}

impl Diagnostics {
    /// Whether the commit changed the presentation; an append always does.
    #[must_use]
    pub const fn changed(&self) -> bool {
        self.changed
    }

    /// The source proof of the committed append; `None` for a replayed
    /// patch.
    #[must_use]
    pub const fn proof(&self) -> Option<&SourceProof> {
        self.proof.as_ref()
    }
}

/// The exact, reversible change of a tail append: the presentations it was
/// committed between and the SHA-256 proofs of their `content.xml`.
///
/// It replays onto (a share of) the presentation it was committed on, or
/// onto an independently opened presentation with equal `content.xml`. Its
/// inverse removes the authored page through the same bounded replacement
/// proof; an unrelated content source is refused with [`Error::PatchConflict`].
#[derive(Clone)]
pub struct Patch {
    before: Arc<SourceBackedPresentation>,
    after: Arc<SourceBackedPresentation>,
    source_content_sha256: [u8; 32],
    target_content_sha256: [u8; 32],
    before_proof: SourceProof,
    after_proof: SourceProof,
    fragment: Arc<[u8]>,
    limits: Limits,
    options: SourceContentPublicationOptions,
    removes_source_range: bool,
}

impl Patch {
    /// Whether the patch changes nothing; an append always changes the
    /// presentation.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.source_content_sha256 == self.target_content_sha256
    }

    /// The patch that restores the before presentation.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            before: Arc::clone(&self.after),
            after: Arc::clone(&self.before),
            source_content_sha256: self.target_content_sha256,
            target_content_sha256: self.source_content_sha256,
            before_proof: self.after_proof.clone(),
            after_proof: self.before_proof.clone(),
            fragment: Arc::clone(&self.fragment),
            limits: self.limits,
            options: self.options.clone(),
            removes_source_range: !self.removes_source_range,
        }
    }

    /// SHA-256 of the `content.xml` the patch applies to.
    #[must_use]
    pub const fn source_content_sha256(&self) -> [u8; 32] {
        self.source_content_sha256
    }

    /// SHA-256 of the `content.xml` the patch produces.
    #[must_use]
    pub const fn target_content_sha256(&self) -> [u8; 32] {
        self.target_content_sha256
    }
}

impl fmt::Debug for Patch {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Patch")
            .field("noop", &self.is_noop())
            .finish_non_exhaustive()
    }
}

/// A staged one-slide append to a [`SourceBackedPresentation`], holding only
/// the bounded title and body. Dropping it discards it.
pub struct Edit<'a> {
    source: &'a SourceBackedPresentation,
    title: String,
    body: String,
    limits: Limits,
    options: SourceContentPublicationOptions,
}

impl fmt::Debug for Edit<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Edit")
            .field("title_bytes", &self.title.len())
            .field("body_bytes", &self.body.len())
            .field("limits", &self.limits)
            .finish_non_exhaustive()
    }
}

impl SourceBackedPresentation {
    /// Start appending one plain title/body slide after the last slide.
    ///
    /// # Errors
    ///
    /// Refuses text the append cannot represent or that exceeds the default
    /// [`Limits`].
    pub fn edit_tail_append(
        &self,
        title: impl Into<String>,
        body: impl Into<String>,
    ) -> std::result::Result<Edit<'_>, Error> {
        let title = title.into();
        let body = body.into();
        let limits = Limits::default();
        validate_text(&title, &body, limits.max_text_bytes)?;
        Ok(Edit {
            source: self,
            title,
            body,
            limits,
            options: SourceContentPublicationOptions::default(),
        })
    }

    /// Apply a tail-append patch to a new presentation; this one is
    /// unchanged.
    ///
    /// # Errors
    ///
    /// Returns [`Error::PatchConflict`] when the presentation has different
    /// `content.xml` bytes, or when its structure does not match the bounded
    /// source proof captured by the patch.
    pub fn apply_tail_append(&self, patch: &Patch) -> std::result::Result<Commit, Error> {
        if self.same_share(&patch.before) {
            self.check_source().map_err(source_error)?;
            return Ok(Commit::new(
                patch.after.share(),
                patch.clone(),
                Diagnostics {
                    changed: !patch.is_noop(),
                    proof: None,
                },
            ));
        }
        rebase_patch(self, patch)
    }
}

fn rebase_patch(
    recipient: &SourceBackedPresentation,
    patch: &Patch,
) -> std::result::Result<Commit, Error> {
    recipient.check_source().map_err(source_error)?;
    let base = recipient.effective_with_odf_error().map_err(odf_error)?;
    let expected = &patch.before_proof;
    let package = Arc::clone(&base.package);
    let facts = scan_or_cached_source(base, &package, patch.limits, &patch.options)?;
    if facts.content_bytes != expected.content_bytes
        || facts.insert_at != expected.insert_at
        || facts.slide_count != expected.slide_count
    {
        return Err(Error::PatchConflict {
            operation: "odp.tail_append.apply",
            reason: "the recipient content structure differs from the recorded base",
        });
    }

    let removes = patch.removes_source_range;
    let remove_length = if removes {
        u64::try_from(patch.fragment.len()).map_err(|_| {
            invalid("ODP tail-append patch fragment length exceeds the replay range")
        })?
    } else {
        0
    };
    let replacement_offset = if removes {
        expected
            .insert_at
            .checked_sub(remove_length)
            .ok_or(Error::PatchConflict {
                operation: "odp.tail_append.apply",
                reason: "the recorded inverse range precedes the recipient content",
            })?
    } else {
        expected.insert_at
    };
    let fragment = if removes {
        AuthoredXmlFragment::empty()
    } else {
        AuthoredXmlFragment::markup(patch.fragment.as_ref().to_vec()).map_err(Error::Invalid)?
    };
    let candidate_token_bytes = patch
        .limits
        .max_fragment_bytes
        .saturating_add(1024)
        .clamp(1, 64 * 1024 * 1024);
    let common = SourceContentInsertionPlan::prepare_replacement_with_validator(
        package,
        replacement_offset,
        remove_length,
        fragment,
        XmlStreamLimits::new(
            patch.limits.max_content_xml_bytes,
            256,
            4_000_000,
            1_000_000,
            candidate_token_bytes,
        )
        .map_err(Error::Invalid)?,
        &patch.options,
        |_event, _tracker| Ok(()),
    )
    .map_err(publication_error)?;
    if common.source_content_sha256() != patch.source_content_sha256
        || common.target_content_sha256() != patch.target_content_sha256
        || common.source_content_length() != expected.content_bytes
        || common.target_content_length() != patch.after_proof.content_bytes
    {
        return Err(Error::PatchConflict {
            operation: "odp.tail_append.apply",
            reason: "the recipient content bytes differ from the recorded base",
        });
    }

    base.remember_append_facts(&facts, &patch.options);

    let options = patch.options.clone();
    let package = base.deferred(common, options);
    let rebased = Patch {
        before: Arc::new(base.share()),
        after: Arc::new(package.share()),
        ..patch.clone()
    };
    recipient.check_source().map_err(source_error)?;
    Ok(Commit::new(
        package,
        rebased,
        Diagnostics {
            changed: !patch.is_noop(),
            proof: None,
        },
    ))
}

impl<'a> Edit<'a> {
    /// The append bounds.
    #[must_use]
    pub const fn limits(&self) -> Limits {
        self.limits
    }

    /// Replace the append bounds.
    ///
    /// # Errors
    ///
    /// Refuses when the staged text exceeds the new bounds; the edit keeps
    /// its previous bounds.
    pub fn set_limits(&mut self, limits: Limits) -> std::result::Result<(), Error> {
        validate_text(&self.title, &self.body, limits.max_text_bytes)?;
        self.limits = limits;
        Ok(())
    }

    /// Set the cancellation, execution and output policy of the source scan,
    /// the insertion plan and the committed presentation's
    /// [`write_to`](SourceBackedPresentation::write_to).
    pub fn set_options(&mut self, options: SourceContentPublicationOptions) {
        self.options = options;
    }

    /// Scan the source and prepare the bounded insertion: the source proof
    /// and the common raw-preserving ZIP plan. Nothing is written.
    ///
    /// # Errors
    ///
    /// Returns the typed refusal of the source, the authored page or the
    /// plan's limits and execution policy.
    pub fn prepare(self) -> std::result::Result<Plan<'a>, Error> {
        let Self {
            source,
            title,
            body,
            limits,
            options,
        } = self;
        let base = source.effective_with_odf_error().map_err(odf_error)?;
        let package = Arc::clone(&base.package);
        check_plan_options(&options)?;
        let source_version = package.source_version().map_err(source_error)?;
        let content_bytes = package
            .member_materialized_size(ODF_CONTENT)
            .map_err(source_error)?
            .ok_or_else(|| invalid("ODP package has no content.xml"))?;
        if content_bytes > limits.max_content_xml_bytes {
            return Err(invalid("ODP content.xml exceeds the source append limit"));
        }
        let first = scan_or_cached_source(base, &package, limits, &options)?;
        if first.slide_count >= limits.max_slides {
            return Err(invalid("ODP append would exceed the slide-count limit"));
        }
        let page_number = first
            .slide_count
            .checked_add(1)
            .ok_or_else(|| invalid("ODP page-name index overflow"))?;
        let page_name = format!("page{page_number}");
        let authored = {
            let retained_text_capacity = title
                .capacity()
                .checked_add(body.capacity())
                .ok_or_else(|| invalid("ODP retained text capacity overflow"))?;
            let readback_bound = generated_readback_memory_bound(
                &title,
                &body,
                retained_text_capacity,
                limits.max_fragment_bytes,
            )
            .map_err(Error::Invalid)?;
            let _memory = reserve_fragment_memory(&options, readback_bound)?;
            let fragment = build_plain_page(&page_name, &title, &body, limits.max_fragment_bytes)?;
            // Keep the reservation alive while the exact authored fragment is
            // audited by the established ODP parser.  The parser sees only
            // this bounded page wrapped in a small, locally namespace-bound
            // document; the source content remains cold and opaque.
            validate_generated_page_semantics(&fragment, &title, &body).map_err(Error::Invalid)?;
            let authored = AuthoredXmlFragment::markup(fragment).map_err(Error::Invalid)?;
            // Release submitted allocations while their reservation is still
            // held; only the authored fragment enters common planning.
            drop(title);
            drop(body);
            authored
        };
        let mut validator = CandidateValidator::new(first.slide_count, &page_name);
        let candidate_token_bytes = limits
            .max_fragment_bytes
            .saturating_add(1024)
            .clamp(1, 64 * 1024 * 1024);
        let common = SourceContentInsertionPlan::prepare_with_validator(
            package,
            first.insert_at,
            authored,
            XmlStreamLimits::new(
                limits.max_content_xml_bytes,
                256,
                4_000_000,
                1_000_000,
                candidate_token_bytes,
            )
            .map_err(Error::Invalid)?,
            &options,
            |event, tracker| validator.observe(event, tracker),
        )
        .map_err(publication_error)?;

        base.remember_append_facts(&first, &options);

        Ok(Plan {
            source,
            common,
            options,
            limits,
            proof: SourceProof {
                source_version,
                content_bytes,
                insert_at: first.insert_at,
                slide_count: first.slide_count,
                page_name,
            },
        })
    }

    /// [`Self::prepare`] and [`Plan::commit`].
    ///
    /// # Errors
    ///
    /// As [`Self::prepare`].
    pub fn commit(self) -> std::result::Result<Commit, Error> {
        self.prepare()?.commit()
    }
}

/// A fully source-proved one-slide append, ready to commit.
pub struct Plan<'a> {
    source: &'a SourceBackedPresentation,
    common: SourceContentInsertionPlan,
    options: SourceContentPublicationOptions,
    limits: Limits,
    proof: SourceProof,
}

impl fmt::Debug for Plan<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Plan")
            .field("proof", &self.proof)
            .finish_non_exhaustive()
    }
}

impl Plan<'_> {
    /// Return the ODP source proof.
    #[must_use]
    pub const fn proof(&self) -> &SourceProof {
        &self.proof
    }

    /// Return the common insertion plan for diagnostics.
    #[must_use]
    pub fn common_plan(&self) -> &SourceContentInsertionPlan {
        &self.common
    }

    /// Return the source archive version captured by the common plan.
    #[must_use]
    pub const fn source_version(&self) -> SourceVersion {
        self.common.source_version()
    }

    /// Return the source archive length captured by the common plan.
    #[must_use]
    pub const fn source_length(&self) -> u64 {
        self.common.source_length()
    }

    /// Return the decoded insertion offset captured by the common plan.
    #[must_use]
    pub const fn decoded_offset(&self) -> u64 {
        self.common.decoded_offset()
    }

    /// Return the source `content.xml` decoded length.
    #[must_use]
    pub const fn source_content_length(&self) -> u64 {
        self.common.source_content_length()
    }

    /// Return the candidate `content.xml` decoded length.
    #[must_use]
    pub const fn target_content_length(&self) -> u64 {
        self.common.target_content_length()
    }

    /// Return the source `content.xml` SHA-256 proof.
    #[must_use]
    pub const fn source_content_sha256(&self) -> [u8; 32] {
        self.common.source_content_sha256()
    }

    /// Return the candidate `content.xml` SHA-256 proof.
    #[must_use]
    pub const fn target_content_sha256(&self) -> [u8; 32] {
        self.common.target_content_sha256()
    }

    /// Commit the append to a new presentation over the same retained
    /// source. Nothing is written: writing the committed presentation
    /// streams the source with the page inserted into `content.xml` and
    /// every other member copied raw.
    ///
    /// # Errors
    ///
    /// Returns a source change of the presentation the edit was made on.
    pub fn commit(self) -> std::result::Result<Commit, Error> {
        let Self {
            source,
            common,
            options,
            limits,
            proof,
        } = self;
        let base = source.effective_with_odf_error().map_err(odf_error)?;
        let source_content_sha256 = common.source_content_sha256();
        let target_content_sha256 = common.target_content_sha256();
        let fragment: Arc<[u8]> = Arc::from(common.authored_fragment_bytes());
        let options_for_patch = options.clone();
        let package = base.deferred(common, options);
        let after_proof = SourceProof {
            source_version: proof.source_version,
            content_bytes: proof
                .content_bytes
                .checked_add(u64::try_from(fragment.len()).map_err(|_| {
                    Error::Invalid(CoreError::InvalidFormat(
                        "ODP tail-append fragment length exceeds u64".to_string(),
                    ))
                })?)
                .ok_or_else(|| {
                    Error::Invalid(CoreError::InvalidFormat(
                        "ODP tail-append target content length overflows u64".to_string(),
                    ))
                })?,
            insert_at: proof
                .insert_at
                .checked_add(u64::try_from(fragment.len()).map_err(|_| {
                    Error::Invalid(CoreError::InvalidFormat(
                        "ODP tail-append fragment length exceeds u64".to_string(),
                    ))
                })?)
                .ok_or_else(|| {
                    Error::Invalid(CoreError::InvalidFormat(
                        "ODP tail-append insertion offset overflows u64".to_string(),
                    ))
                })?,
            slide_count: proof.slide_count.checked_add(1).ok_or_else(|| {
                Error::Invalid(CoreError::InvalidFormat(
                    "ODP tail-append slide count overflows usize".to_string(),
                ))
            })?,
            page_name: proof.page_name.clone(),
        };
        let patch = Patch {
            before: Arc::new(source.share()),
            after: Arc::new(package.share()),
            source_content_sha256,
            target_content_sha256,
            before_proof: proof.clone(),
            after_proof,
            fragment,
            limits,
            options: options_for_patch,
            removes_source_range: false,
        };
        Ok(Commit::new(
            package,
            patch,
            Diagnostics {
                changed: true,
                proof: Some(proof),
            },
        ))
    }
}

#[derive(Debug, Default, Clone, PartialEq, Eq)]
pub(crate) struct ScanFacts {
    pub(crate) content_bytes: u64,
    pub(crate) insert_at: u64,
    pub(crate) slide_count: usize,
}

fn scan_source(
    package: &SourceBackedPackage,
    limits: Limits,
    options: &SourceContentPublicationOptions,
) -> std::result::Result<ScanFacts, Error> {
    let mut facts = ScanFacts {
        content_bytes: 0,
        insert_at: 0,
        slide_count: 0,
    };
    let stream_limits = XmlStreamLimits::new(
        limits.max_content_xml_bytes,
        256,
        4_000_000,
        1_000_000,
        64 * 1024,
    )
    .map_err(Error::Invalid)?;
    let mut state = OdpScanState::default();
    let report = scan_source_content(
        package,
        stream_limits,
        options,
        |event: &XmlStreamEvent<'_>, tracker: &BindingTracker| {
            state.observe(event, tracker, &mut facts, limits)
        },
    )
    .map_err(|error| match error {
        SourceContentScanError::Publication(error) => publication_error(error),
        SourceContentScanError::SourceMember(error) => source_member_error(error),
        _ => Error::Invalid(CoreError::InvalidFormat(
            "unknown source-content scan failure".to_string(),
        )),
    })?;
    facts.content_bytes = report.bytes();
    facts.slide_count = state.pages_seen;
    state.finish(&facts).map_err(Error::Invalid)?;
    Ok(facts)
}

fn scan_or_cached_source(
    base: &SourceBackedPresentation,
    package: &SourceBackedPackage,
    limits: Limits,
    options: &SourceContentPublicationOptions,
) -> std::result::Result<ScanFacts, Error> {
    if let Some(facts) = base
        .cached_append_facts(limits.max_content_xml_bytes, limits.max_slides, options)
        .map_err(Error::Invalid)?
    {
        return Ok(facts);
    }
    scan_source(package, limits, options)
}

#[derive(Debug, Default)]
struct OdpScanState {
    stack: Vec<ElementKind>,
    root_seen: bool,
    body_seen: bool,
    presentation_seen: bool,
    presentation_empty: bool,
    page_open: bool,
    pages_seen: usize,
    trailing_seen: bool,
    content_bytes: u64,
}

#[derive(Debug, Default, Clone, Copy, PartialEq, Eq)]
enum ElementKind {
    #[default]
    Other,
    Root,
    Body,
    Presentation,
    Page,
}

impl OdpScanState {
    fn observe(
        &mut self,
        event: &XmlStreamEvent<'_>,
        tracker: &BindingTracker,
        facts: &mut ScanFacts,
        limits: Limits,
    ) -> Result<()> {
        self.content_bytes = self.content_bytes.max(event.end());
        if self.content_bytes > limits.max_content_xml_bytes {
            return Err(candidate_invalid(
                "ODP content.xml exceeds the source append limit",
            ));
        }
        match event.event() {
            Event::Start(element) => {
                let depth = self.stack.len();
                let kind = self.classify_start(depth, element, tracker)?;
                if kind == ElementKind::Page {
                    if self.trailing_seen {
                        return Err(candidate_invalid(
                            "ODP presentation pages are not a contiguous direct prefix",
                        ));
                    }
                    if self.page_open {
                        return Err(candidate_invalid(
                            "ODP presentation contains a nested draw:page",
                        ));
                    }
                    if self.pages_seen >= limits.max_slides {
                        return Err(candidate_invalid(
                            "ODP presentation exceeds the slide-count limit",
                        ));
                    }
                    self.pages_seen = self
                        .pages_seen
                        .checked_add(1)
                        .ok_or_else(|| candidate_invalid("ODP slide-count overflow"))?;
                    self.page_open = true;
                } else if depth == 3
                    && self.presentation_seen
                    && self.pages_seen > 0
                    && !self.page_open
                {
                    self.trailing_seen = true;
                }
                self.stack.push(kind);
            },
            Event::Empty(element) => {
                let depth = self.stack.len();
                let kind = self.classify_start(depth, element, tracker)?;
                if kind == ElementKind::Page {
                    if self.trailing_seen {
                        return Err(candidate_invalid(
                            "ODP presentation pages are not a contiguous direct prefix",
                        ));
                    }
                    if self.pages_seen >= limits.max_slides {
                        return Err(candidate_invalid(
                            "ODP presentation exceeds the slide-count limit",
                        ));
                    }
                    self.pages_seen = self
                        .pages_seen
                        .checked_add(1)
                        .ok_or_else(|| candidate_invalid("ODP slide-count overflow"))?;
                    facts.insert_at = event.end();
                } else if depth == 3 && self.presentation_seen && self.pages_seen > 0 {
                    self.trailing_seen = true;
                }
                if kind == ElementKind::Presentation {
                    self.presentation_empty = true;
                }
            },
            Event::End(_) => {
                let kind = self
                    .stack
                    .pop()
                    .ok_or_else(|| candidate_invalid("ODP content.xml depth underflow"))?;
                match kind {
                    ElementKind::Page => {
                        if !self.page_open {
                            return Err(candidate_invalid("ODP presentation page state underflow"));
                        }
                        self.page_open = false;
                        facts.insert_at = event.end();
                    },
                    ElementKind::Presentation => {
                        if self.presentation_empty {
                            return Err(candidate_invalid(
                                "ODP append does not support self-closing office:presentation",
                            ));
                        }
                    },
                    _ => {},
                }
            },
            Event::Text(text)
                if self.stack.len() == 3 && self.presentation_seen && self.pages_seen > 0 =>
            {
                if !text.iter().all(u8::is_ascii_whitespace) {
                    self.trailing_seen = true;
                }
            },
            Event::Comment(_) | Event::PI(_) if self.stack.len() == 3 && self.pages_seen > 0 => {
                self.trailing_seen = true;
            },
            Event::CData(_) | Event::GeneralRef(_)
                if self.stack.len() == 3 && self.pages_seen > 0 =>
            {
                self.trailing_seen = true;
            },
            Event::DocType(_) => {
                return Err(candidate_invalid(
                    "ODP content.xml must not contain a doctype",
                ));
            },
            Event::Eof => {},
            _ => {},
        }
        Ok(())
    }

    fn classify_start(
        &mut self,
        depth: usize,
        element: &BytesStart<'_>,
        tracker: &BindingTracker,
    ) -> Result<ElementKind> {
        let (namespace, local) = tracker.resolve_element(element.name()).map_err(|error| {
            error.into_litchi_error_with_context(|| "invalid ODP namespace binding".to_string())
        })?;
        let is = |expected_namespace: &[u8], expected_local: &[u8]| {
            matches!(namespace, ResolveResult::Bound(Namespace(uri)) if uri == expected_namespace)
                && local.as_ref() == expected_local
        };
        let kind = if depth == 0 {
            if !is(OFFICE_NAMESPACE, b"document-content") {
                return Err(candidate_invalid(
                    "ODP content.xml root is not office:document-content",
                ));
            }
            if self.root_seen {
                return Err(candidate_invalid("ODP content.xml has duplicate root"));
            }
            self.root_seen = true;
            ElementKind::Root
        } else if depth == 1 && is(OFFICE_NAMESPACE, b"body") {
            if self.stack.len() != 1 || self.stack[0] != ElementKind::Root {
                return Err(candidate_invalid(
                    "ODP office:body is outside office:document-content",
                ));
            }
            if self.body_seen {
                return Err(candidate_invalid(
                    "ODP content.xml has duplicate office:body",
                ));
            }
            self.body_seen = true;
            ElementKind::Body
        } else if depth == 2 && is(OFFICE_NAMESPACE, b"presentation") {
            if self.stack.len() != 2
                || self.stack[0] != ElementKind::Root
                || self.stack[1] != ElementKind::Body
            {
                return Err(candidate_invalid(
                    "ODP office:presentation is outside office:body",
                ));
            }
            if self.presentation_seen {
                return Err(candidate_invalid(
                    "ODP content.xml has duplicate office:presentation",
                ));
            }
            self.presentation_seen = true;
            ElementKind::Presentation
        } else if is(DRAW_NAMESPACE, b"page") {
            if depth != 3
                || self.stack.len() != 3
                || self.stack[0] != ElementKind::Root
                || self.stack[1] != ElementKind::Body
                || self.stack[2] != ElementKind::Presentation
            {
                return Err(candidate_invalid(
                    "ODP draw:page must be a direct presentation child",
                ));
            }
            ElementKind::Page
        } else {
            ElementKind::Other
        };
        Ok(kind)
    }

    fn finish(&self, facts: &ScanFacts) -> Result<()> {
        if !self.root_seen || !self.body_seen || !self.presentation_seen {
            return Err(candidate_invalid(
                "ODP content.xml has no complete presentation body",
            ));
        }
        if self.presentation_empty || facts.slide_count == 0 {
            return Err(candidate_invalid(
                "ODP append requires a non-empty, non-self-closing presentation",
            ));
        }
        if !self.stack.is_empty() || self.page_open {
            return Err(candidate_invalid(
                "ODP content.xml has unterminated elements",
            ));
        }
        if facts.insert_at == 0 || facts.insert_at > facts.content_bytes {
            return Err(candidate_invalid(
                "ODP append insertion offset is outside content.xml",
            ));
        }
        Ok(())
    }
}

/// Candidate-side proof state for the one generated page.
///
/// The source scan proves the existing presentation and the common insertion
/// owner proves the complete byte-level splice.  This state adds the
/// format-owned candidate fact that exactly one generated direct page occurs
/// immediately after the source page prefix, before any trailing child.
#[derive(Debug)]
struct CandidateValidator {
    source_pages: usize,
    expected_name: String,
    stack: Vec<ElementKind>,
    root_seen: bool,
    body_seen: bool,
    presentation_seen: bool,
    presentation_empty: bool,
    page_open: bool,
    trailing_seen: bool,
    pages_seen: usize,
    expected_seen: bool,
}

impl CandidateValidator {
    fn new(source_pages: usize, expected_name: &str) -> Self {
        Self {
            source_pages,
            expected_name: expected_name.to_owned(),
            stack: Vec::new(),
            root_seen: false,
            body_seen: false,
            presentation_seen: false,
            presentation_empty: false,
            page_open: false,
            trailing_seen: false,
            pages_seen: 0,
            expected_seen: false,
        }
    }

    fn observe(&mut self, event: &XmlStreamEvent<'_>, tracker: &BindingTracker) -> Result<()> {
        match event.event() {
            Event::Start(element) => {
                let depth = self.stack.len();
                let kind = self.classify_start(depth, element, tracker)?;
                if kind == ElementKind::Page {
                    self.begin_page(element, tracker, event.decoder(), false)?;
                    self.page_open = true;
                } else if self.is_direct_presentation_child(depth)
                    && self.pages_seen > 0
                    && !self.page_open
                {
                    self.trailing_seen = true;
                }
                self.stack.push(kind);
            },
            Event::Empty(element) => {
                let depth = self.stack.len();
                let kind = self.classify_start(depth, element, tracker)?;
                if kind == ElementKind::Page {
                    self.begin_page(element, tracker, event.decoder(), true)?;
                } else if self.is_direct_presentation_child(depth) && self.pages_seen > 0 {
                    self.trailing_seen = true;
                }
                if kind == ElementKind::Presentation {
                    self.presentation_empty = true;
                }
            },
            Event::End(_) => {
                let kind = self
                    .stack
                    .pop()
                    .ok_or_else(|| candidate_invalid("ODP candidate depth underflow"))?;
                match kind {
                    ElementKind::Page => {
                        if !self.page_open {
                            return Err(candidate_invalid("ODP candidate page state underflow"));
                        }
                        self.page_open = false;
                    },
                    ElementKind::Presentation => {
                        if self.presentation_empty {
                            return Err(candidate_invalid(
                                "ODP candidate contains a self-closing presentation",
                            ));
                        }
                    },
                    _ => {},
                }
            },
            Event::Text(text) if self.is_direct_presentation_text() && self.pages_seen > 0 => {
                if !text.iter().all(u8::is_ascii_whitespace) {
                    self.trailing_seen = true;
                }
            },
            Event::Comment(_) | Event::PI(_)
                if self.is_direct_presentation_text() && self.pages_seen > 0 =>
            {
                self.trailing_seen = true;
            },
            Event::CData(_) | Event::GeneralRef(_)
                if self.is_direct_presentation_text() && self.pages_seen > 0 =>
            {
                self.trailing_seen = true;
            },
            Event::Eof => self.finish()?,
            _ => {},
        }
        Ok(())
    }

    fn begin_page(
        &mut self,
        element: &BytesStart<'_>,
        tracker: &BindingTracker,
        decoder: quick_xml::encoding::Decoder,
        empty: bool,
    ) -> Result<()> {
        if self.trailing_seen {
            return Err(candidate_invalid(
                "ODP candidate pages are not a contiguous direct prefix",
            ));
        }
        if self.page_open {
            return Err(candidate_invalid(
                "ODP candidate contains a nested draw:page",
            ));
        }
        let position = self.pages_seen;
        self.pages_seen = self
            .pages_seen
            .checked_add(1)
            .ok_or_else(|| candidate_invalid("ODP candidate slide-count overflow"))?;
        if page_name_matches_core(element, tracker, decoder, &self.expected_name)? {
            if self.expected_seen {
                return Err(candidate_invalid(
                    "ODP candidate contains the generated page more than once",
                ));
            }
            if position != self.source_pages {
                return Err(candidate_invalid(
                    "ODP generated page is not immediately after the source page prefix",
                ));
            }
            self.expected_seen = true;
            if empty {
                return Err(candidate_invalid(
                    "ODP generated page must be a non-empty draw:page",
                ));
            }
        }
        Ok(())
    }

    fn classify_start(
        &mut self,
        depth: usize,
        element: &BytesStart<'_>,
        tracker: &BindingTracker,
    ) -> Result<ElementKind> {
        let (namespace, local) = tracker.resolve_element(element.name()).map_err(|error| {
            error.into_litchi_error_with_context(|| {
                "invalid ODP candidate namespace binding".to_string()
            })
        })?;
        let is = |expected_namespace: &[u8], expected_local: &[u8]| {
            matches!(namespace, ResolveResult::Bound(Namespace(uri)) if uri == expected_namespace)
                && local.as_ref() == expected_local
        };
        if depth == 0 {
            if !is(OFFICE_NAMESPACE, b"document-content") {
                return Err(candidate_invalid(
                    "ODP candidate root is not office:document-content",
                ));
            }
            if self.root_seen {
                return Err(candidate_invalid("ODP candidate has duplicate root"));
            }
            self.root_seen = true;
            return Ok(ElementKind::Root);
        }
        if depth == 1 && is(OFFICE_NAMESPACE, b"body") {
            if self.stack.len() != 1 || self.stack[0] != ElementKind::Root {
                return Err(candidate_invalid(
                    "ODP candidate office:body is outside office:document-content",
                ));
            }
            if self.body_seen {
                return Err(candidate_invalid("ODP candidate has duplicate office:body"));
            }
            self.body_seen = true;
            return Ok(ElementKind::Body);
        }
        if depth == 2 && is(OFFICE_NAMESPACE, b"presentation") {
            if self.stack.len() != 2
                || self.stack[0] != ElementKind::Root
                || self.stack[1] != ElementKind::Body
            {
                return Err(candidate_invalid(
                    "ODP candidate office:presentation is outside office:body",
                ));
            }
            if self.presentation_seen {
                return Err(candidate_invalid(
                    "ODP candidate has duplicate office:presentation",
                ));
            }
            self.presentation_seen = true;
            return Ok(ElementKind::Presentation);
        }
        if is(DRAW_NAMESPACE, b"page") {
            if depth != 3
                || self.stack.len() != 3
                || self.stack[0] != ElementKind::Root
                || self.stack[1] != ElementKind::Body
                || self.stack[2] != ElementKind::Presentation
            {
                return Err(candidate_invalid(
                    "ODP candidate draw:page must be a direct presentation child",
                ));
            }
            return Ok(ElementKind::Page);
        }
        Ok(ElementKind::Other)
    }

    fn is_direct_presentation_child(&self, depth: usize) -> bool {
        depth == 3
            && self.presentation_seen
            && self.stack.len() == 3
            && self.stack[0] == ElementKind::Root
            && self.stack[1] == ElementKind::Body
            && self.stack[2] == ElementKind::Presentation
    }

    fn is_direct_presentation_text(&self) -> bool {
        self.stack.len() == 3
            && self.presentation_seen
            && self.stack[0] == ElementKind::Root
            && self.stack[1] == ElementKind::Body
            && self.stack[2] == ElementKind::Presentation
    }

    fn finish(&self) -> Result<()> {
        if !self.root_seen || !self.body_seen || !self.presentation_seen {
            return Err(candidate_invalid(
                "ODP candidate has no complete presentation body",
            ));
        }
        if self.presentation_empty || self.pages_seen == 0 {
            return Err(candidate_invalid(
                "ODP candidate requires a non-empty presentation",
            ));
        }
        if !self.stack.is_empty() || self.page_open {
            return Err(candidate_invalid("ODP candidate has unterminated elements"));
        }
        let expected_count = self
            .source_pages
            .checked_add(1)
            .ok_or_else(|| candidate_invalid("ODP candidate slide-count overflow"))?;
        if self.pages_seen != expected_count {
            return Err(candidate_invalid(
                "ODP candidate page count does not match the source plus one generated page",
            ));
        }
        if !self.expected_seen {
            return Err(candidate_invalid(
                "ODP candidate is missing the generated non-empty page",
            ));
        }
        Ok(())
    }
}

fn candidate_invalid(message: impl Into<String>) -> CoreError {
    CoreError::InvalidFormat(message.into())
}

fn page_name_matches_core(
    element: &BytesStart<'_>,
    tracker: &BindingTracker,
    decoder: quick_xml::encoding::Decoder,
    expected: &str,
) -> Result<bool> {
    for attribute in element.attributes().with_checks(false) {
        let attribute = attribute.map_err(|error| {
            CoreError::InvalidFormat(format!("invalid ODP draw:page attribute: {error}"))
        })?;
        let (namespace, local) = tracker.resolve_attribute(attribute.key).map_err(|error| {
            error.into_litchi_error_with_context(|| "invalid ODP attribute namespace".to_string())
        })?;
        if local.as_ref() == b"name"
            && matches!(namespace, ResolveResult::Bound(Namespace(uri)) if uri == DRAW_NAMESPACE)
        {
            let value = attribute
                .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
                .map_err(|error| {
                    CoreError::InvalidFormat(format!("invalid ODP draw:name: {error}"))
                })?;
            return Ok(value == expected);
        }
    }
    Ok(false)
}

/// Prove the generated page's semantic projection with the same ODP parser
/// used by `Presentation` and `SourceBackedPresentation`.  The fragment is
/// wrapped only for this bounded readback pass; the candidate scanner still
/// proves the exact namespace bindings, page position, count, and name in the
/// actual source splice.
fn validate_generated_page_semantics(fragment: &[u8], title: &str, body: &str) -> Result<()> {
    let length = GENERATED_DOCUMENT_PREFIX
        .len()
        .checked_add(fragment.len())
        .and_then(|length| length.checked_add(GENERATED_DOCUMENT_SUFFIX.len()))
        .ok_or_else(|| {
            CoreError::InvalidFormat("ODP generated document size overflow".to_string())
        })?;
    let mut document = Vec::new();
    document
        .try_reserve_exact(length)
        .map_err(|source| CoreError::Allocation {
            resource: "ODP generated semantic readback document",
            source,
        })?;
    document.extend_from_slice(GENERATED_DOCUMENT_PREFIX);
    document.extend_from_slice(fragment);
    document.extend_from_slice(GENERATED_DOCUMENT_SUFFIX);
    let document = std::str::from_utf8(&document).map_err(|error| {
        CoreError::InvalidFormat(format!("generated ODP page is not UTF-8: {error}"))
    })?;
    let slides = crate::codec::Parser::parse_slides_with_styles(document, None)?;
    let slide = slides.first().ok_or_else(|| {
        CoreError::InvalidFormat("ODP generated page semantic readback found no slide".to_string())
    })?;
    if slides.len() != 1 {
        return Err(CoreError::InvalidFormat(
            "ODP generated page semantic readback found more than one slide".to_string(),
        ));
    }
    if slide.title.as_deref().unwrap_or_default() != title {
        return Err(CoreError::InvalidFormat(
            "ODP generated title does not survive the established parser projection".to_string(),
        ));
    }
    if body.trim().is_empty() {
        if !slide.text.is_empty() {
            return Err(CoreError::InvalidFormat(
                "ODP generated whitespace-only body unexpectedly entered the slide text projection"
                    .to_string(),
            ));
        }
        let mut object_shape = None;
        for shape in &slide.shapes {
            if shape.presentation_class() != Some("object") {
                continue;
            }
            if object_shape.replace(shape).is_some() {
                return Err(CoreError::InvalidFormat(
                    "ODP generated whitespace-only body produced multiple object shapes"
                        .to_string(),
                ));
            }
        }
        if body.is_empty() {
            if object_shape.is_some() {
                return Err(CoreError::InvalidFormat(
                    "ODP generated empty body produced an object shape".to_string(),
                ));
            }
        } else {
            let shape = object_shape.ok_or_else(|| {
                CoreError::InvalidFormat(
                    "ODP generated whitespace-only body was not retained as an object shape"
                        .to_string(),
                )
            })?;
            if shape.text()? != body {
                return Err(CoreError::InvalidFormat(
                    "ODP generated whitespace-only body does not survive object-shape readback"
                        .to_string(),
                ));
            }
        }
    } else if slide.text != body {
        return Err(CoreError::InvalidFormat(
            "ODP generated body does not survive the established parser projection".to_string(),
        ));
    }
    Ok(())
}

fn validate_text(title: &str, body: &str, maximum: usize) -> std::result::Result<(), Error> {
    validate_text_core(title, body, maximum).map_err(Error::Invalid)
}

fn validate_text_core(title: &str, body: &str, maximum: usize) -> Result<()> {
    let bytes = title
        .len()
        .checked_add(body.len())
        .ok_or_else(|| CoreError::InvalidFormat("ODP slide text size overflow".to_string()))?;
    if bytes > maximum {
        return Err(CoreError::InvalidFormat(
            "ODP slide text exceeds the append limit".to_string(),
        ));
    }
    if title
        .chars()
        .chain(body.chars())
        .any(|character| character == '\r')
    {
        return Err(CoreError::InvalidFormat(
            "ODP slide text contains a carriage return; use '\\n' for paragraph breaks".to_string(),
        ));
    }
    if title.chars().chain(body.chars()).any(|character| {
        !matches!(
            character,
            '\u{9}' | '\u{A}' | '\u{D}' | '\u{20}'..='\u{D7FF}'
                | '\u{E000}'..='\u{FFFD}' | '\u{10000}'..='\u{10FFFF}'
        )
    }) {
        return Err(CoreError::InvalidFormat(
            "ODP slide text contains a character forbidden by XML 1.0".to_string(),
        ));
    }
    if title.chars().chain(body.chars()).any(|character| {
        character.is_whitespace() && !matches!(character, ' ' | '\t' | '\n' | '\r')
    }) {
        return Err(CoreError::InvalidFormat(
            "ODP slide text contains Unicode whitespace that the established parser normalizes; use spaces, tabs, or newlines".to_string(),
        ));
    }
    Ok(())
}

fn build_plain_page(
    name: &str,
    title: &str,
    body: &str,
    maximum: usize,
) -> std::result::Result<Vec<u8>, Error> {
    let mut output = Vec::new();
    append_limited(
        &mut output,
        br#"<draw:page xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:presentation="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" draw:name=""#,
        maximum,
    )?;
    append_escaped_text(&mut output, name, maximum)?;
    append_limited(&mut output, b"\">", maximum)?;
    if !title.is_empty() {
        append_limited(
            &mut output,
            br#"<draw:frame presentation:class="title" svg:width="25.199cm" svg:height="3.506cm" svg:x="1.4cm" svg:y="0.962cm"><draw:text-box>"#,
            maximum,
        )?;
        append_text_paragraphs(&mut output, title, maximum)?;
        append_limited(&mut output, b"</draw:text-box></draw:frame>", maximum)?;
    }
    if !body.is_empty() {
        let y = if title.is_empty() { "2.0cm" } else { "5.0cm" };
        append_limited(
            &mut output,
            br#"<draw:frame presentation:class="object" svg:width="25.199cm" svg:height="10cm" svg:x="1.4cm" svg:y=""#,
            maximum,
        )?;
        append_limited(&mut output, y.as_bytes(), maximum)?;
        append_limited(&mut output, b"\"><draw:text-box>", maximum)?;
        append_text_paragraphs(&mut output, body, maximum)?;
        append_limited(&mut output, b"</draw:text-box></draw:frame>", maximum)?;
    }
    append_limited(&mut output, b"</draw:page>", maximum)?;
    Ok(output)
}

fn append_limited(
    output: &mut Vec<u8>,
    bytes: &[u8],
    maximum: usize,
) -> std::result::Result<(), Error> {
    let required = output
        .len()
        .checked_add(bytes.len())
        .ok_or_else(|| invalid("generated ODP page length overflow"))?;
    if required > maximum {
        return Err(invalid("generated ODP page exceeds the fragment limit"));
    }
    output
        .try_reserve_exact(bytes.len())
        .map_err(|error| invalid(format!("ODP append fragment allocation failed: {error}")))?;
    output.extend_from_slice(bytes);
    Ok(())
}

fn append_escaped_text(
    output: &mut Vec<u8>,
    text: &str,
    maximum: usize,
) -> std::result::Result<(), Error> {
    for character in text.chars() {
        append_escaped_char(output, character, maximum)?;
    }
    Ok(())
}

fn append_escaped_char(
    output: &mut Vec<u8>,
    character: char,
    maximum: usize,
) -> std::result::Result<(), Error> {
    match character {
        '&' => append_limited(output, b"&amp;", maximum),
        '<' => append_limited(output, b"&lt;", maximum),
        '>' => append_limited(output, b"&gt;", maximum),
        '"' => append_limited(output, b"&quot;", maximum),
        '\'' => append_limited(output, b"&apos;", maximum),
        character => {
            let mut encoded = [0_u8; 4];
            let encoded = character.encode_utf8(&mut encoded);
            append_limited(output, encoded.as_bytes(), maximum)
        },
    }
}

fn append_text_paragraphs(
    output: &mut Vec<u8>,
    text: &str,
    maximum: usize,
) -> std::result::Result<(), Error> {
    for paragraph in text.split('\n') {
        append_limited(output, b"<text:p>", maximum)?;
        append_encoded_paragraph(output, paragraph, maximum)?;
        append_limited(output, b"</text:p>", maximum)?;
    }
    Ok(())
}

fn append_encoded_paragraph(
    output: &mut Vec<u8>,
    paragraph: &str,
    maximum: usize,
) -> std::result::Result<(), Error> {
    let mut characters = paragraph.chars().peekable();
    let mut encoded_any = false;
    while let Some(character) = characters.next() {
        match character {
            ' ' => {
                let mut count = 1usize;
                while characters.next_if_eq(&' ').is_some() {
                    count = count
                        .checked_add(1)
                        .ok_or_else(|| invalid("ODP generated space count overflow"))?;
                }
                if count == 1 && encoded_any && characters.peek().is_some() {
                    append_limited(output, b" ", maximum)?;
                } else if count == 1 {
                    append_limited(output, b"<text:s/>", maximum)?;
                } else {
                    let mut remaining = count;
                    while remaining != 0 {
                        let chunk_count = remaining.min(MAX_TEXT_SPACE_COUNT);
                        append_limited(output, b"<text:s text:c=\"", maximum)?;
                        let chunk = chunk_count.to_string();
                        append_limited(output, chunk.as_bytes(), maximum)?;
                        append_limited(output, b"\"/>", maximum)?;
                        remaining -= chunk_count;
                    }
                }
                encoded_any = true;
            },
            '\t' => {
                append_limited(output, b"<text:tab/>", maximum)?;
                encoded_any = true;
            },
            '\r' => {
                return Err(invalid(
                    "ODP slide text contains a carriage return; use '\\n' for paragraph breaks",
                ));
            },
            character => {
                append_escaped_char(output, character, maximum)?;
                encoded_any = true;
            },
        }
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;
    use litchi_core::{OwnedSource, ReadAt};
    use litchi_odf_common::Error as OdfError;
    use litchi_odf_common::constants::ODF_PRESENTATION;
    use litchi_odf_common::core::OwnedPackage;
    use soapberry_zip::office::StreamingArchiveWriter;
    use std::io;
    use std::sync::atomic::{AtomicU64, Ordering};

    use super::super::Presentation;

    const OFFICE_NAMESPACE_UTF8: &str = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
    const DRAW_NAMESPACE_UTF8: &str = "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0";
    const PRESENTATION_NAMESPACE_UTF8: &str =
        "urn:oasis:names:tc:opendocument:xmlns:presentation:1.0";

    fn source_package_bytes(content: &str) -> Vec<u8> {
        let manifest = format!(
            r#"<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"><manifest:file-entry manifest:full-path="/" manifest:media-type="{ODF_PRESENTATION}"/><manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/></manifest:manifest>"#,
        );
        let mut writer =
            StreamingArchiveWriter::new(soapberry_zip::office::StreamingArchiveLimits::default());
        writer
            .write_stored("mimetype", ODF_PRESENTATION.as_bytes())
            .unwrap();
        writer
            .write_stored(ODF_CONTENT, content.as_bytes())
            .unwrap();
        writer
            .write_stored("META-INF/manifest.xml", manifest.as_bytes())
            .unwrap();
        writer.finish_to_bytes().unwrap()
    }

    fn source_package(content: &str) -> Arc<dyn ReadAt> {
        Arc::new(OwnedSource::new(source_package_bytes(content)))
    }

    struct VersionedSource {
        bytes: Arc<Vec<u8>>,
        revision: AtomicU64,
    }

    impl VersionedSource {
        fn new(bytes: Vec<u8>) -> Self {
            Self {
                bytes: Arc::new(bytes),
                revision: AtomicU64::new(0),
            }
        }

        fn bump(&self) {
            self.revision.fetch_add(1, Ordering::Relaxed);
        }
    }

    impl ReadAt for VersionedSource {
        fn len(&self) -> io::Result<u64> {
            u64::try_from(self.bytes.len())
                .map_err(|_| io::Error::new(io::ErrorKind::InvalidInput, "source length overflow"))
        }

        fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
            let offset = usize::try_from(offset).map_err(|_| {
                io::Error::new(io::ErrorKind::InvalidInput, "source offset overflow")
            })?;
            if offset >= self.bytes.len() {
                return Ok(0);
            }
            let count = output.len().min(self.bytes.len() - offset);
            output[..count].copy_from_slice(&self.bytes[offset..offset + count]);
            Ok(count)
        }

        fn version(&self) -> io::Result<SourceVersion> {
            Ok(SourceVersion::new(
                0x4f44_505f_7461_696c,
                self.revision.load(Ordering::Relaxed),
            ))
        }
    }

    #[test]
    fn empty_presentation_refuses_with_typed_invalid_error() {
        let content = format!(
            r#"<office:document-content xmlns:office="{OFFICE_NAMESPACE_UTF8}"><office:body><office:presentation></office:presentation></office:body></office:document-content>"#,
        );
        let owner = SourceBackedPresentation::from_read_at(source_package(&content)).unwrap();
        let edit = owner.edit_tail_append("title", "body").unwrap();
        let error = edit.prepare().unwrap_err();
        assert!(matches!(
            error,
            Error::Invalid(CoreError::InvalidFormat(message))
                if message.contains("non-empty")
        ));
    }

    #[test]
    fn strict_open_keeps_the_legacy_body_probe_fallback_for_malformed_xml() {
        let content = format!(
            r#"<office:document-content xmlns:office="{OFFICE_NAMESPACE_UTF8}" xmlns:draw="{DRAW_NAMESPACE_UTF8}"><office:body><office:presentation><draw:page></office:presentation></office:body></office:document-content>"#,
        );
        let owner = SourceBackedPresentation::from_read_at(source_package(&content))
            .expect("the historical open probe admits a body before semantic parsing");
        assert!(
            owner.slides().is_err(),
            "semantic parsing still reports malformed XML"
        );
    }

    #[test]
    fn doctype_open_remains_accepted_but_append_keeps_its_refusal() {
        let content = format!(
            r#"<!DOCTYPE office:document-content><office:document-content xmlns:office="{OFFICE_NAMESPACE_UTF8}" xmlns:draw="{DRAW_NAMESPACE_UTF8}"><office:body><office:presentation><draw:page/></office:presentation></office:body></office:document-content>"#,
        );
        let owner = SourceBackedPresentation::from_read_at(source_package(&content)).unwrap();
        let error = owner
            .edit_tail_append("title", "body")
            .unwrap()
            .prepare()
            .unwrap_err();
        let message = format!("{error}");
        assert!(message.contains("DTD") || message.contains("DOCTYPE"));
    }

    #[test]
    fn self_closing_presentation_refuses_with_typed_invalid_error() {
        let content = format!(
            r#"<office:document-content xmlns:office="{OFFICE_NAMESPACE_UTF8}"><office:body><office:presentation/></office:body></office:document-content>"#,
        );
        let owner = SourceBackedPresentation::from_read_at(source_package(&content)).unwrap();
        let edit = owner.edit_tail_append("title", "body").unwrap();
        let error = edit.prepare().unwrap_err();
        assert!(matches!(
            error,
            Error::Invalid(CoreError::InvalidFormat(message))
                if message.contains("non-empty") || message.contains("self-closing")
        ));
    }

    #[test]
    fn candidate_validation_refuses_existing_generated_name_without_source_rescan() {
        let content = format!(
            r#"<office:document-content xmlns:office="{OFFICE_NAMESPACE_UTF8}" xmlns:draw="{DRAW_NAMESPACE_UTF8}"><office:body><office:presentation><draw:page draw:name="page2"/></office:presentation></office:body></office:document-content>"#,
        );
        let owner = SourceBackedPresentation::from_read_at(source_package(&content)).unwrap();
        let edit = owner.edit_tail_append("title", "body").unwrap();
        let error = edit.prepare().unwrap_err();
        assert!(matches!(
            error,
            Error::Plan(SourceContentPublicationError::Core(
                CoreError::InvalidFormat(message)
            )) if message.contains("generated page")
        ));
    }

    #[test]
    fn publication_inserts_before_native_trailing_presentation_child() {
        let content = format!(
            r#"<office:document-content xmlns:office="{OFFICE_NAMESPACE_UTF8}" xmlns:draw="{DRAW_NAMESPACE_UTF8}" xmlns:presentation="{PRESENTATION_NAMESPACE_UTF8}"><office:body><office:presentation><draw:page draw:name="page1"/><presentation:settings presentation:start-page="page1"/><!--tail-token--></office:presentation></office:body></office:document-content>"#,
        );
        let owner = SourceBackedPresentation::from_read_at(source_package(&content)).unwrap();
        let edit = owner.edit_tail_append("title", "body").unwrap();
        let commit = edit.commit().unwrap();
        let mut output = Vec::new();
        commit.package().write_to(&mut output).unwrap();

        let package = OwnedPackage::from_bytes(output).unwrap();
        let published = String::from_utf8(package.get_file(ODF_CONTENT).unwrap()).unwrap();
        let generated = published.find(r#"draw:name="page2""#).unwrap();
        let settings = published.find("<presentation:settings").unwrap();
        assert!(generated < settings);
        assert!(published.contains(
            r#"<presentation:settings presentation:start-page="page1"/><!--tail-token-->"#
        ));
    }

    #[test]
    fn generated_story_semantics_match_established_readback() {
        let content = format!(
            r#"<office:document-content xmlns:office="{OFFICE_NAMESPACE_UTF8}" xmlns:draw="{DRAW_NAMESPACE_UTF8}"><office:body><office:presentation><draw:page draw:name="page1"/></office:presentation></office:body></office:document-content>"#,
        );
        let title = "A & < B > \"C\" ' 😀\t\nnext";
        let body = "  body  \n第二行";
        let owner = SourceBackedPresentation::from_read_at(source_package(&content)).unwrap();
        let edit = owner.edit_tail_append(title, body).unwrap();
        let commit = edit.commit().unwrap();
        let mut output = Vec::new();
        commit.package().write_to(&mut output).unwrap();

        let source_backed =
            SourceBackedPresentation::from_read_at(Arc::new(OwnedSource::new(output.clone())))
                .unwrap();
        let source_slides = source_backed.slides().unwrap();
        let source_slide = source_slides.last().unwrap();
        assert_eq!(source_slide.title().unwrap(), Some(title));
        assert_eq!(source_slide.text().unwrap(), body);

        let presentation = Presentation::from_bytes(output).unwrap();
        let slides = presentation.slides().unwrap();
        let slide = slides.last().unwrap();
        assert_eq!(slide.title().unwrap(), Some(title));
        assert_eq!(slide.text().unwrap(), body);
    }

    #[test]
    fn whitespace_only_body_is_read_back_from_object_shape() {
        let content = format!(
            r#"<office:document-content xmlns:office="{OFFICE_NAMESPACE_UTF8}" xmlns:draw="{DRAW_NAMESPACE_UTF8}"><office:body><office:presentation><draw:page draw:name="page1"/></office:presentation></office:body></office:document-content>"#,
        );
        let body = " \t\n  ";
        let owner = SourceBackedPresentation::from_read_at(source_package(&content)).unwrap();
        let edit = owner.edit_tail_append("title", body).unwrap();
        let commit = edit.commit().unwrap();
        let mut output = Vec::new();
        commit.package().write_to(&mut output).unwrap();

        let source_backed =
            SourceBackedPresentation::from_read_at(Arc::new(OwnedSource::new(output.clone())))
                .unwrap();
        let source_slide = source_backed.slides().unwrap().pop().unwrap();
        assert!(source_slide.text().unwrap().is_empty());
        let source_objects: Vec<_> = source_slide
            .shapes()
            .unwrap()
            .iter()
            .filter(|shape| shape.presentation_class() == Some("object"))
            .collect();
        assert_eq!(source_objects.len(), 1);
        assert_eq!(source_objects[0].text().unwrap(), body);

        let presentation = Presentation::from_bytes(output).unwrap();
        let slide = presentation.slides().unwrap().pop().unwrap();
        assert!(slide.text().unwrap().is_empty());
        let objects: Vec<_> = slide
            .shapes()
            .unwrap()
            .iter()
            .filter(|shape| shape.presentation_class() == Some("object"))
            .collect();
        assert_eq!(objects.len(), 1);
        assert_eq!(objects[0].text().unwrap(), body);
    }

    #[test]
    fn unicode_whitespace_is_a_typed_refusal() {
        let content = format!(
            r#"<office:document-content xmlns:office="{OFFICE_NAMESPACE_UTF8}" xmlns:draw="{DRAW_NAMESPACE_UTF8}"><office:body><office:presentation><draw:page draw:name="page1"/></office:presentation></office:body></office:document-content>"#,
        );
        let owner = SourceBackedPresentation::from_read_at(source_package(&content)).unwrap();
        let error = owner.edit_tail_append("title\u{2003}", "body").unwrap_err();
        assert!(matches!(
            error,
            Error::Invalid(CoreError::InvalidFormat(message))
                if message.contains("Unicode whitespace")
        ));
    }

    #[test]
    fn large_space_runs_are_split_at_parser_limit() {
        let body = " ".repeat(MAX_TEXT_SPACE_COUNT + 1);
        let fragment = build_plain_page("page1", "", &body, MAX_FRAGMENT_BYTES).unwrap();
        let xml = std::str::from_utf8(&fragment).unwrap();
        assert_eq!(xml.matches(r#"<text:s text:c="#).count(), 2);
        assert!(xml.contains(r#"<text:s text:c="1000000"/>"#));
        assert!(!xml.contains(r#"text:c="1000001""#));
    }

    #[test]
    fn carriage_return_text_is_a_typed_refusal() {
        let content = format!(
            r#"<office:document-content xmlns:office="{OFFICE_NAMESPACE_UTF8}" xmlns:draw="{DRAW_NAMESPACE_UTF8}"><office:body><office:presentation><draw:page draw:name="page1"/></office:presentation></office:body></office:document-content>"#,
        );
        let owner = SourceBackedPresentation::from_read_at(source_package(&content)).unwrap();
        let error = owner
            .edit_tail_append("title\rwith carriage return", "body")
            .unwrap_err();
        assert!(matches!(
            error,
            Error::Invalid(CoreError::InvalidFormat(message))
                if message.contains("carriage return")
        ));
    }

    #[test]
    fn leading_presentation_declaration_does_not_block_page_tail_append() {
        let content = format!(
            r#"<office:document-content xmlns:office="{OFFICE_NAMESPACE_UTF8}" xmlns:draw="{DRAW_NAMESPACE_UTF8}" xmlns:presentation="{PRESENTATION_NAMESPACE_UTF8}"><office:body><office:presentation><presentation:settings presentation:start-page="page1"/><draw:page draw:name="page1"/></office:presentation></office:body></office:document-content>"#,
        );
        let owner = SourceBackedPresentation::from_read_at(source_package(&content)).unwrap();
        let edit = owner.edit_tail_append("title", "body").unwrap();
        let commit = edit.commit().unwrap();
        let mut output = Vec::new();
        commit.package().write_to(&mut output).unwrap();

        let package = OwnedPackage::from_bytes(output).unwrap();
        let published = String::from_utf8(package.get_file(ODF_CONTENT).unwrap()).unwrap();
        let settings = published.find("<presentation:settings").unwrap();
        let generated = published.find(r#"draw:name="page2""#).unwrap();
        assert!(settings < generated);
    }

    #[test]
    fn commit_is_deferred_reversible_and_chains() {
        let content = format!(
            r#"<office:document-content xmlns:office="{OFFICE_NAMESPACE_UTF8}" xmlns:draw="{DRAW_NAMESPACE_UTF8}"><office:body><office:presentation><draw:page draw:name="page1"/></office:presentation></office:body></office:document-content>"#,
        );
        let source = source_package(&content);
        let mut source_bytes = vec![0; usize::try_from(source.len().unwrap()).unwrap()];
        source.read_at(0, &mut source_bytes).unwrap();
        let owner = SourceBackedPresentation::from_read_at(source).unwrap();
        let plan = owner
            .edit_tail_append("first", "one")
            .unwrap()
            .prepare()
            .unwrap();
        assert_eq!(plan.proof().slide_count(), 1);
        assert_eq!(plan.proof().page_name(), "page2");
        let commit = plan.commit().unwrap();
        assert!(commit.diagnostics().changed());
        assert_eq!(commit.diagnostics().proof().unwrap().slide_count(), 1);
        assert!(!commit.patch().is_noop());

        // Reads materialize the committed output; the source is unchanged.
        assert_eq!(commit.package().slide_count().unwrap(), 2);
        assert_eq!(owner.slide_count().unwrap(), 1);

        // A second append on the committed presentation chains.
        let chained = commit
            .package()
            .edit_tail_append("second", "two")
            .unwrap()
            .commit()
            .unwrap();
        assert_eq!(chained.package().slide_count().unwrap(), 3);
        assert_eq!(
            chained
                .package()
                .slide(2)
                .unwrap()
                .unwrap()
                .title()
                .unwrap(),
            Some("second")
        );

        // The patch replays onto its source and its inverse restores it.
        let replayed = owner.apply_tail_append(commit.patch()).unwrap();
        let mut replayed_bytes = Vec::new();
        replayed.package().write_to(&mut replayed_bytes).unwrap();
        let mut committed_bytes = Vec::new();
        commit.package().write_to(&mut committed_bytes).unwrap();
        assert_eq!(replayed_bytes, committed_bytes);
        let restored = commit
            .package()
            .apply_tail_append(&commit.patch().inverse())
            .unwrap();
        let mut restored_bytes = Vec::new();
        restored.package().write_to(&mut restored_bytes).unwrap();
        assert_eq!(restored_bytes, source_bytes);
        assert!(matches!(
            chained.package().apply_tail_append(commit.patch()),
            Err(Error::PatchConflict { .. })
        ));
    }

    #[test]
    fn stale_source_is_reported_separately_from_patch_conflict() {
        let content = format!(
            r#"<office:document-content xmlns:office="{OFFICE_NAMESPACE_UTF8}" xmlns:draw="{DRAW_NAMESPACE_UTF8}"><office:body><office:presentation><draw:page draw:name="page1"/></office:presentation></office:body></office:document-content>"#,
        );
        let source = Arc::new(VersionedSource::new(source_package_bytes(&content)));
        let owner = SourceBackedPresentation::from_read_at(source.clone()).unwrap();
        let commit = owner
            .edit_tail_append("title", "body")
            .unwrap()
            .commit()
            .unwrap();

        source.bump();
        assert!(matches!(
            owner.apply_tail_append(commit.patch()),
            Err(Error::SourceChanged { .. })
        ));
        assert!(matches!(
            owner.edit_tail_append("later", "body").unwrap().prepare(),
            Err(Error::SourceChanged { .. })
        ));
    }

    #[cfg(any(unix, windows))]
    #[test]
    fn stale_save_retains_insertion_progress_and_source_diagnostic() {
        let content = format!(
            r#"<office:document-content xmlns:office="{OFFICE_NAMESPACE_UTF8}" xmlns:draw="{DRAW_NAMESPACE_UTF8}"><office:body><office:presentation><draw:page draw:name="page1"/></office:presentation></office:body></office:document-content>"#,
        );
        let source = Arc::new(VersionedSource::new(source_package_bytes(&content)));
        let owner = SourceBackedPresentation::from_read_at(source.clone()).unwrap();
        let committed = owner
            .edit_tail_append("title", "body")
            .unwrap()
            .commit()
            .unwrap();
        source.bump();

        let directory = tempfile::tempdir().unwrap();
        let error = committed
            .package()
            .save_with_durability(
                directory.path().join("stale.odp"),
                litchi_core::Durability::NoSync,
            )
            .unwrap_err();
        let OdfError::SourceContentInsertion(error) = error else {
            panic!("stale save must retain insertion error, got {error:?}");
        };
        assert_eq!(
            error.progress(),
            SourceContentPublicationProgress::Untouched
        );
        assert!(matches!(
            error.as_ref(),
            SourceContentInsertionError::Publication {
                source: SourceContentPublicationError::SourceChanged { .. },
                ..
            }
        ));
    }

    #[test]
    fn effective_odf_conversion_retains_replay_progress_and_callback_diagnostic() {
        use soapberry_zip::{ReplayPass, ReplayProgress, ReplayPublicationError};

        let expected = SourceVersion::new(0x301, 1);
        let observed = SourceVersion::new(0x301, 2);
        let insertion = SourceContentInsertionError::Replay {
            progress: SourceContentPublicationProgress::Complete { bytes: 29 },
            source: ReplayPublicationError::Callback {
                pass: ReplayPass::Emit,
                source: SourceContentInsertionCallbackError::Member(Box::new(
                    SourceMemberReaderError::Core {
                        error: CoreError::SourceChanged { expected, observed },
                        callback_error: Some(SourceContentInsertionCallbackError::Verification {
                            what: "callback verification failed",
                        }),
                    },
                )),
                progress: ReplayProgress::Complete { bytes: 29 },
            },
        };
        let error = odf_error(OdfError::from(insertion));

        assert_eq!(
            error.progress(),
            Some(SourceContentPublicationProgress::Complete { bytes: 29 })
        );
        assert!(matches!(
            &error,
            Error::SourceChanged {
                expected: actual_expected,
                observed: actual_observed,
                source: Some(source),
                ..
            } if *actual_expected == expected
                && *actual_observed == observed
                && matches!(
                    source.as_ref(),
                    SourceChangedSource::Insertion(SourceContentInsertionError::Replay {
                        source: ReplayPublicationError::Callback {
                            source: SourceContentInsertionCallbackError::Member(member),
                            ..
                        },
                        ..
                    }) if matches!(
                        member.as_ref(),
                        SourceMemberReaderError::Core {
                            callback_error: Some(
                                SourceContentInsertionCallbackError::Verification { what: message },
                            ),
                            ..
                        } if *message == "callback verification failed"
                    )
                )
        ));
    }

    #[test]
    fn stale_publication_preserves_before_and_after_progress() {
        let expected = SourceVersion::new(0x100, 1);
        let observed = SourceVersion::new(0x100, 2);
        let before = publication_error(SourceContentPublicationError::SourceChanged {
            expected,
            observed,
            progress: SourceContentPublicationProgress::Untouched,
        });
        assert_eq!(
            before.progress(),
            Some(SourceContentPublicationProgress::Untouched)
        );

        let after = publication_error(SourceContentPublicationError::Source {
            progress: SourceContentPublicationProgress::Complete { bytes: 17 },
            source: CoreError::SourceChanged { expected, observed },
        });
        assert_eq!(
            after.progress(),
            Some(SourceContentPublicationProgress::Complete { bytes: 17 })
        );

        let callback = source_member_error(SourceMemberReaderError::Callback(
            CoreError::SourceChanged { expected, observed },
        ));
        assert_eq!(
            callback.progress(),
            Some(SourceContentPublicationProgress::Untouched)
        );
    }

    #[test]
    fn stale_source_retains_core_diagnostic() {
        let expected = SourceVersion::new(0x300, 1);
        let observed = SourceVersion::new(0x300, 2);
        let error = source_error(CoreError::SourceChanged { expected, observed });

        assert!(matches!(
            &error,
            Error::SourceChanged {
                source: Some(source),
                ..
            } if matches!(
                source.as_ref(),
                SourceChangedSource::Core(CoreError::SourceChanged {
                    expected: retained_expected,
                    observed: retained_observed,
                }) if *retained_expected == expected && *retained_observed == observed
            )
        ));
        assert!(matches!(
            std::error::Error::source(&error)
                .and_then(|source| source.downcast_ref::<SourceChangedSource>()),
            Some(SourceChangedSource::Core(CoreError::SourceChanged { .. }))
        ));
    }

    #[test]
    fn stale_verification_retains_secondary_callback_failure() {
        let expected = SourceVersion::new(0x200, 1);
        let observed = SourceVersion::new(0x200, 2);
        let error = source_member_error(SourceMemberReaderError::Core {
            error: CoreError::SourceChanged { expected, observed },
            callback_error: Some(CoreError::InvalidFormat(
                "callback verification failed".to_string(),
            )),
        });

        assert_eq!(
            error.progress(),
            Some(SourceContentPublicationProgress::Untouched)
        );
        assert!(matches!(
            &error,
            Error::SourceChanged {
                source: Some(source),
                ..
            } if matches!(
                source.as_ref(),
                SourceChangedSource::SourceMember(SourceMemberReaderError::Core {
                    error: CoreError::SourceChanged { .. },
                    callback_error: Some(CoreError::InvalidFormat(message)),
                }) if message == "callback verification failed"
            )
        ));

        let retained = std::error::Error::source(&error)
            .and_then(|source| source.downcast_ref::<SourceChangedSource>())
            .expect("stale source keeps its typed reader diagnostic");
        assert!(matches!(
            retained,
            SourceChangedSource::SourceMember(SourceMemberReaderError::Core {
                callback_error: Some(CoreError::InvalidFormat(message)),
                ..
            }) if message == "callback verification failed"
        ));
    }

    #[test]
    fn equal_content_recipients_rebase_forward_and_inverse_tail_append() {
        let content = format!(
            r#"<office:document-content xmlns:office="{OFFICE_NAMESPACE_UTF8}" xmlns:draw="{DRAW_NAMESPACE_UTF8}"><office:body><office:presentation><draw:page draw:name="page1"/></office:presentation></office:body></office:document-content>"#,
        );
        let source = source_package(&content);
        let mut source_bytes = vec![0; usize::try_from(source.len().unwrap()).unwrap()];
        source.read_at(0, &mut source_bytes).unwrap();
        let owner = SourceBackedPresentation::from_read_at(Arc::new(OwnedSource::new(
            source_bytes.clone(),
        )))
        .unwrap();
        let independent = SourceBackedPresentation::from_read_at(Arc::new(OwnedSource::new(
            source_bytes.clone(),
        )))
        .unwrap();
        let commit = owner
            .edit_tail_append("title", "body")
            .unwrap()
            .commit()
            .unwrap();

        let rebased = independent.apply_tail_append(commit.patch()).unwrap();
        let mut rebased_bytes = Vec::new();
        rebased.package().write_to(&mut rebased_bytes).unwrap();
        let mut committed_bytes = Vec::new();
        commit.package().write_to(&mut committed_bytes).unwrap();
        assert_eq!(rebased_bytes, committed_bytes);

        let target_copy =
            SourceBackedPresentation::from_read_at(Arc::new(OwnedSource::new(rebased_bytes)))
                .unwrap();
        let restored = target_copy
            .apply_tail_append(&commit.patch().inverse())
            .unwrap();
        let mut restored_bytes = Vec::new();
        restored.package().write_to(&mut restored_bytes).unwrap();
        assert_eq!(restored_bytes, source_bytes);

        let wrong_content = content.replace("page1", "pageX");
        let wrong = SourceBackedPresentation::from_read_at(source_package(&wrong_content)).unwrap();
        assert!(matches!(
            wrong.apply_tail_append(commit.patch()),
            Err(Error::PatchConflict { .. })
        ));
    }
}
