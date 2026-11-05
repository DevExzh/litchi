//! Source-checked, lossless edits of text already owned by a PPT shape.
//!
//! Length-changing replacement is supported when the selected text body is a
//! `TextHeaderAtom` and one `TextCharsAtom` or `TextBytesAtom` followed by at
//! most one each of:
//!
//! - a `StyleTextPropAtom` with one paragraph run and one character run
//!   (optionally plus a one-character run for the paragraph mark);
//! - a `TextSpecialInfoAtom` with one run, or one text run plus a mark run;
//! - a `TextRulerAtom`, which holds no character position.
//!
//! The coverage counts are then updated with the text atom and every owning
//! container length. `[MS-PPT]` associates other formatting, interaction,
//! bookmark, metacharacter, and multi-run special-information records with
//! UTF-16 positions; their presence remains a typed dependency-closure
//! refusal for a changed length (same-length replacement stays available).
//!
//! The text of a title or body placeholder that `PowerPoint` keeps in the
//! slide list (its `ClientTextbox` holds only an `OutlineTextRefAtom`, MS-PPT
//! 2.9.78) is edited in place in the live `DocumentContainer`'s
//! `SlideListWithTextContainer` body the reference selects, under the same
//! rules, by the root transaction's batch publisher.
//!
//! The editing owners are crate-private since record 0357: the public text
//! edits are the [`crate::Package::edit`] verbs (`set_shape_text`,
//! `set_shape_texts`, `set_notes_text`, `set_run_format`, …), which drive
//! this module. It keeps public the vocabulary those verbs take
//! ([`Target`](crate::text_edit::Target),
//! [`RunFormat`](crate::text_edit::RunFormat)), the read views
//! ([`TextRuns`](crate::text_edit::TextRuns) and the source-backed
//! [`SourceSnapshot`](crate::text_edit::SourceSnapshot) reader) and the typed
//! [`Error`](crate::text_edit::Error)/[`Refusal`](crate::text_edit::Refusal).
//!
//! The source-backed same-length splice is a scoped family since record
//! 0357a: [`Package::edit_text`](crate::Package::edit_text) returns a
//! [`text_edit::Edit`](crate::text_edit::Edit) whose commit publishes a new
//! package over the composed overlay source without materializing it, and
//! [`Package::apply_text`](crate::Package::apply_text) replays its
//! [`text_edit::Patch`](crate::text_edit::Patch).
#![allow(
    dead_code,
    reason = "crate-private native text owners (record 0357): the whole-artifact transaction and the source-backed splice owner behind text_edit::Edit keep their full surface exercised by the unit tests"
)]

use std::collections::{BTreeMap, BTreeSet, btree_map::Entry};
use std::fmt;
use std::io::{Cursor, Write};
use std::sync::Arc;

pub use litchi_core::Position;

use crate::consts::RecordType;
use crate::package::{Error as PackageError, Result as PackageResult};
use crate::presentation::Presentation;
use crate::raw::Reader;
use crate::shapes::ShapeEnum;
use litchi_cfb::{
    ArtifactFingerprint, OleError, OverlayError, PublishReport, SameLengthStreamSplice,
    SharedOleFile, SharedOleFileLimits, StreamSpliceLimits, ValidatedOverlayPlan,
};
use litchi_core::{ReadAt, SourceVersion};
use litchi_ole_common::source_backed_overlay::SourceBackedOverlayPublisher;

mod family;
mod readback;
#[cfg(test)]
mod readback_tests;
pub(crate) mod runs;
mod source_range;

pub use family::{Commit, Diagnostics, Edit, LimitKind, Patch};
pub use runs::{CharacterRun, RunFormat, TextLink, TextRuns};

const PPT_HEADER_LEN: usize = 8;
const OFFICEART_SP_CONTAINER: u16 = 0xF004;
const OFFICEART_SP: u16 = 0xF00A;
const OFFICEART_CLIENT_TEXTBOX: u16 = 0xF00D;
const OFFICEART_CLIENT_ANCHOR: u16 = 0xF010;

/// Semantic identity of one existing text-bearing shape.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct Target {
    slide: Position,
    shape: Position,
}

impl Target {
    /// Creates a target from zero-based semantic positions in the immutable
    /// presentation and slide shape collections.
    #[must_use]
    pub const fn new(slide: Position, shape: Position) -> Self {
        Self { slide, shape }
    }

    /// The selected slide.
    #[must_use]
    pub const fn slide(self) -> Position {
        self.slide
    }

    /// The selected shape's zero-based position in its slide's source order.
    #[must_use]
    pub const fn shape(self) -> Position {
        self.shape
    }
}

/// A reason why a text edit cannot be published without rewriting unmodeled
/// dependencies.
#[non_exhaustive]
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum Refusal {
    /// The source package is signed, encrypted, or otherwise cannot be
    /// republished through the incremental PPT transaction owner.
    UnsupportedSource,
    /// No slide resolved from the public selector.
    SlideNotFound { position: Position },
    /// No shape exists at the selected source-order position.
    ShapeNotFound,
    /// The selected semantic shape position resolved to ambiguous native data.
    AmbiguousShape,
    /// The selected shape has no top-level `PowerPoint` host anchor.
    NoAnchor,
    /// The selected shape has several top-level host anchors.
    AmbiguousAnchor,
    /// The selected shape has no host `ClientTextbox` payload.
    NoTextbox,
    /// More than one host textbox payload would need a choice.
    AmbiguousTextbox,
    /// The textbox has no editable text atom.
    NoTextAtom,
    /// The textbox has several text atoms and their owner relationship is not
    /// modeled by this focused transaction.
    MultipleTextAtoms,
    /// The replacement cannot be represented by the original text atom's
    /// encoding while preserving its byte length.
    IncompatibleEncoding,
    /// The replacement changes UTF-16 position counts and would invalidate
    /// formatting, interaction, or special-information ranges.
    DependencyClosure,
    /// The source-backed splice contract cannot change the encoded text-atom
    /// length. The ordinary editor remains available for its narrower
    /// length-changing closure.
    LengthChange,
    /// The selected textbox contains a record whose relationship to the text
    /// atom is not modeled by the source-backed splice owner.
    UnsupportedDependency,
    /// The selected slide has no notes page to hold speaker-notes text.
    NoNotesPage,
}

impl fmt::Display for Refusal {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::UnsupportedSource => write!(
                formatter,
                "PPT source cannot be republished through the safe shape-text transaction"
            ),
            Self::SlideNotFound { position } => {
                write!(
                    formatter,
                    "PPT slide position {} was not found",
                    position.get()
                )
            },
            Self::ShapeNotFound => write!(formatter, "PPT shape position was not found"),
            Self::AmbiguousShape => write!(formatter, "PPT shape position is ambiguous"),
            Self::NoAnchor => write!(formatter, "PPT shape has no editable client anchor"),
            Self::AmbiguousAnchor => {
                write!(formatter, "PPT shape has multiple client anchors")
            },
            Self::NoTextbox => write!(formatter, "PPT shape has no ClientTextbox payload"),
            Self::AmbiguousTextbox => {
                write!(formatter, "PPT shape has multiple ClientTextbox payloads")
            },
            Self::NoTextAtom => write!(formatter, "PPT shape has no editable text atom"),
            Self::MultipleTextAtoms => write!(formatter, "PPT shape has multiple text atoms"),
            Self::IncompatibleEncoding => write!(
                formatter,
                "replacement text cannot preserve the PPT shape's encoded text atom"
            ),
            Self::DependencyClosure => write!(
                formatter,
                "replacement text changes the selected PPT shape's modeled dependency closure"
            ),
            Self::LengthChange => write!(
                formatter,
                "source-backed PPT shape-text replacement must preserve encoded atom length"
            ),
            Self::NoNotesPage => write!(formatter, "PPT slide has no notes page"),
            Self::UnsupportedDependency => write!(
                formatter,
                "selected PPT shape text has an unsupported dependency record"
            ),
        }
    }
}

impl Refusal {
    /// The shared refusal class of this reason (record 0351 §2.4 rule 2);
    /// unified callers classify by this kind alone.
    #[must_use]
    pub const fn kind(&self) -> litchi_core::edit::RefusalKind {
        use litchi_core::edit::RefusalKind;

        match self {
            Self::SlideNotFound { .. } | Self::ShapeNotFound => RefusalKind::NotFound,
            Self::AmbiguousShape
            | Self::NoAnchor
            | Self::AmbiguousAnchor
            | Self::NoTextbox
            | Self::AmbiguousTextbox
            | Self::NoTextAtom
            | Self::MultipleTextAtoms
            | Self::NoNotesPage => RefusalKind::UnsupportedTarget,
            Self::UnsupportedSource
            | Self::IncompatibleEncoding
            | Self::DependencyClosure
            | Self::LengthChange
            | Self::UnsupportedDependency => RefusalKind::UnsupportedContent,
        }
    }
}

/// Error returned by the focused text-edit transaction.
#[derive(Debug)]
#[non_exhaustive]
pub enum Error {
    /// The package could not be opened, parsed, or republished.
    Package(PackageError),
    /// The requested operation is not proven lossless for this source.
    Refused(Refusal),
    /// The immutable positional source or sequential overlay publisher
    /// rejected a source-backed operation.
    Source(OverlayError),
    /// A [`Patch`] was made from a package with another exact source
    /// artifact fingerprint (record 0357a).
    PatchConflict,
    /// A changed text splice would invalidate the package signature
    /// (record 0357a).
    Signed,
}

impl fmt::Display for Error {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Package(error) => error.fmt(formatter),
            Self::Refused(refusal) => refusal.fmt(formatter),
            Self::Source(error) => error.fmt(formatter),
            Self::PatchConflict => formatter
                .write_str("PPT text patch source does not match the exact package receiving it"),
            Self::Signed => {
                formatter.write_str("changed PPT text writes require explicit signature removal")
            },
        }
    }
}

impl std::error::Error for Error {
    fn source(&self) -> Option<&(dyn std::error::Error + 'static)> {
        match self {
            Self::Package(error) => Some(error),
            Self::Refused(_) | Self::PatchConflict | Self::Signed => None,
            Self::Source(error) => Some(error),
        }
    }
}

impl From<PackageError> for Error {
    fn from(error: PackageError) -> Self {
        Self::Package(error)
    }
}

impl From<OverlayError> for Error {
    fn from(error: OverlayError) -> Self {
        Self::Source(error)
    }
}

/// Result type for source-checked existing-shape text edits.
pub type Result<T> = std::result::Result<T, Error>;

/// Immutable whole-package snapshot used by [`Transaction`].
#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) struct Snapshot {
    bytes: Arc<[u8]>,
    source_limits: crate::SourceLimits,
}

impl Snapshot {
    /// Opens an exact PPT artifact with default package limits.
    ///
    /// # Errors
    ///
    /// Returns an error when the complete package cannot be opened.
    pub(crate) fn parse(bytes: &[u8]) -> Result<Self> {
        Self::from_bytes(bytes.to_vec())
    }

    /// Captures an owned exact PPT artifact after validating that it opens.
    ///
    /// # Errors
    ///
    /// Returns an error when the complete package cannot be opened.
    pub(crate) fn from_bytes(bytes: Vec<u8>) -> Result<Self> {
        let presentation = presentation(&bytes)?;
        Ok(Self {
            bytes: Arc::from(bytes.into_boxed_slice()),
            source_limits: presentation.source_limits(),
        })
    }

    pub(crate) fn from_shared_bytes(bytes: Arc<[u8]>) -> Result<Self> {
        let presentation = presentation_shared(bytes.clone())?;
        Ok(Self {
            bytes,
            source_limits: presentation.source_limits(),
        })
    }

    /// Exact bytes of the complete source or committed package artifact.
    #[must_use]
    pub(crate) fn bytes(&self) -> &[u8] {
        &self.bytes
    }

    /// Starts a source-checked transaction for one existing shape.
    ///
    /// # Errors
    ///
    /// Returns a typed refusal when the target cannot be resolved to one safe
    /// text atom.
    pub(crate) fn edit_text(&self, target: Target) -> Result<Transaction> {
        self.edit_text_in(target, None)
    }

    /// [`Self::edit_text`], resolving the target through `presentation`, the
    /// already opened projection of exactly these bytes, when one is given.
    pub(crate) fn edit_text_in(
        &self,
        target: Target,
        presentation: Option<&Presentation>,
    ) -> Result<Transaction> {
        // Resolve the semantic selector without surfacing its result yet. The
        // editor preflight below remains the first observable failure, while
        // keeping the two allocation-heavy owners sequential.
        let shape_target = match presentation {
            Some(presentation) => resolve_shape_target_at_with_view(presentation, target),
            None => presentation_shared(self.bytes.clone())
                .map_err(Error::from)
                .and_then(|presentation| {
                    resolve_shape_target_in_presentation_with_view(&presentation, target)
                }),
        };
        // The shared incremental owner rejects signed and encrypted CFB
        // envelopes before any candidate is staged. Surface that capability
        // boundary as a typed refusal instead of trying to normalize it.
        let editor =
            crate::embedded::object::Editor::open_records(self.bytes.to_vec(), self.source_limits)
                .map_err(|_error| Error::Refused(Refusal::UnsupportedSource))?;
        let (shape_target, source_view) = shape_target?;
        let mut resolved = resolve_with_shape_target(target, shape_target, &editor)?;
        resolved.source_view = source_view.map(Arc::new);
        Ok(Transaction {
            source: self.clone(),
            resolved,
            replacement: None,
        })
    }

    /// Applies one shared durable `litchi-ppt` shape-text operation.
    ///
    /// Application checks the format namespace, bounded semantic selector,
    /// exact artifact SHA-256, and expected source text before staging the
    /// ordinary source-checked transaction. A durable no-op patch returns this
    /// snapshot unchanged.
    ///
    /// # Errors
    ///
    /// Returns an error when the patch vocabulary is unsupported, a
    /// precondition conflicts, or the requested text edit is refused.
    pub(crate) fn apply_durable<Mode>(
        &self,
        patch: &litchi_core::patch::Patch<Mode>,
    ) -> Result<Self> {
        if patch.format() != "litchi-ppt" || !patch.blobs().is_empty() {
            return Err(invalid_durable_patch("unsupported format or blob bundle"));
        }
        let operations = patch.operations();
        if operations.is_empty() {
            return Ok(self.clone());
        }
        let [operation] = operations else {
            return Err(invalid_durable_patch(
                "shape-text patch must contain exactly one operation",
            ));
        };
        if operation.op != "shape-text.set" || operation.preconditions.len() != 2 {
            return Err(invalid_durable_patch(
                "unsupported shape-text operation vocabulary",
            ));
        }
        let target = parse_durable_target(&operation.target)?;
        let expected_hash = operation
            .preconditions
            .get("artifact_sha256")
            .and_then(serde_json::Value::as_str)
            .ok_or_else(|| invalid_durable_patch("missing artifact hash precondition"))?;
        let actual_hash = litchi_core::patch::BlobId::of(self.bytes()).as_hex();
        if expected_hash != actual_hash {
            return Err(PackageError::InvalidFormat(
                "PPT durable shape-text patch source artifact does not match".into(),
            )
            .into());
        }
        let expected_text = operation
            .preconditions
            .get("text")
            .and_then(serde_json::Value::as_str)
            .ok_or_else(|| invalid_durable_patch("missing text precondition"))?;
        let replacement = operation
            .value
            .as_str()
            .ok_or_else(|| invalid_durable_patch("shape-text value must be a string"))?;
        let mut transaction = self.edit_text(target)?;
        if transaction.text() != expected_text {
            return Err(PackageError::InvalidFormat(
                "PPT durable shape-text patch text precondition does not match".into(),
            )
            .into());
        }
        transaction.set_text(replacement)?;
        Ok(transaction.commit()?.snapshot)
    }
}

/// Options for the narrow immutable positional PPT text-splice owner.
///
/// The owner intentionally has no length-changing or topology-changing
/// fallback. `source_limits` bounds the semantic metadata and selected slide
/// reads; `splice_limits` bounds the common CFB physical plan.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct SourceBackedOptions {
    /// Finite PPT record and stream limits.
    pub source_limits: crate::SourceLimits,
    /// Finite CFB same-length splice limits.
    pub splice_limits: StreamSpliceLimits,
}

/// Immutable, source-backed legacy PPT snapshot.
///
/// The source allocation is retained behind `Arc<dyn ReadAt>` and is never
/// converted into an owned replacement artifact. Semantic reads fetch bounded
/// selector metadata ranges and one selected slide; the common publisher still
/// performs its full-artifact validation and fingerprint checks. Publication
/// is a checked same-length splice in the existing `PowerPoint Document`
/// stream. Embedded OLE storages are refused; the only accepted storage is the
/// canonical `PP97_DUALSTORAGE` wrapper for the two selected streams.
#[derive(Clone)]
pub struct SourceSnapshot {
    inner: Arc<SourceInner>,
}

struct SourceInner {
    source: Arc<dyn ReadAt>,
    shared: Arc<SharedOleFile>,
    version: SourceVersion,
    length: u64,
    fingerprint: ArtifactFingerprint,
    options: SourceBackedOptions,
}

impl fmt::Debug for SourceSnapshot {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("SourceSnapshot")
            .field("length", &self.inner.length)
            .field("version", &self.inner.version)
            .field("fingerprint", &self.inner.fingerprint)
            .finish_non_exhaustive()
    }
}

impl SourceSnapshot {
    /// Opens a positional source under the default finite limits.
    ///
    /// Opening validates the complete CFB index, protected-container policy,
    /// source version, and complete artifact fingerprint. It does not parse
    /// semantic slide payloads until a selector is requested.
    pub fn open(source: Arc<dyn ReadAt>) -> Result<Self> {
        Self::open_with_options(source, SourceBackedOptions::default())
    }

    /// Opens a positional source with explicit semantic and splice limits.
    pub fn open_with_options(
        source: Arc<dyn ReadAt>,
        options: SourceBackedOptions,
    ) -> Result<Self> {
        let shared_limits = shared_limits(options.source_limits)?;
        let length = source
            .len()
            .map_err(|error| Error::Source(OverlayError::Io(error)))?;
        if length > shared_limits.max_input_bytes() {
            return Err(PackageError::Limit {
                kind: crate::SourceLimitKind::InputBytes,
                observed: usize::try_from(length).unwrap_or(usize::MAX),
                maximum: options.source_limits.max_input_bytes(),
            }
            .into());
        }
        let publisher = SourceBackedOverlayPublisher::open(Arc::clone(&source), shared_limits)
            .map_err(Error::Source)?;
        let shared = publisher.shared();
        crate::package::preflight_shared_ole_file(&shared, options.source_limits)?;
        Self::from_publisher(source, length, &publisher, options)
    }

    /// Builds the snapshot over a positional CFB owner that is already open,
    /// such as the one a [`crate::SourceBackedPackage`] or a format detector
    /// holds, without indexing the container again.
    ///
    /// The checks are the ones [`Self::open_with_options`] runs after its own
    /// index: the protected-container refusal, the source version, and the
    /// complete-artifact fingerprint of an empty splice plan. The CFB index
    /// is the adopted one, so `options.source_limits` bounds the input length
    /// here but cannot re-run the index's directory and allocation-table
    /// ceilings; those are the ones `shared` was opened under.
    ///
    /// # Errors
    ///
    /// Returns an error when the source exceeds `options`, is protected, has
    /// changed since `shared` was opened, or cannot be fingerprinted.
    pub fn from_shared(shared: Arc<SharedOleFile>, options: SourceBackedOptions) -> Result<Self> {
        shared.source_version().map_err(map_ole_error)?;
        crate::package::preflight_shared_ole_file(&shared, options.source_limits)?;
        let shared_limits = shared_limits(options.source_limits)?;
        let source = shared.source_arc();
        let length = shared.file_size();
        if length > shared_limits.max_input_bytes() {
            return Err(PackageError::Limit {
                kind: crate::SourceLimitKind::InputBytes,
                observed: usize::try_from(length).unwrap_or(usize::MAX),
                maximum: options.source_limits.max_input_bytes(),
            }
            .into());
        }
        let publisher = SourceBackedOverlayPublisher::from_shared(shared).map_err(Error::Source)?;
        Self::from_publisher(source, length, &publisher, options)
    }

    /// The common tail of the two constructors: everything after the CFB
    /// owner exists and has passed the protected-container refusal.
    fn from_publisher(
        source: Arc<dyn ReadAt>,
        length: u64,
        publisher: &SourceBackedOverlayPublisher,
        options: SourceBackedOptions,
    ) -> Result<Self> {
        let version = publisher.source_version().map_err(Error::Source)?;

        // An empty splice still runs the common full-artifact fingerprint and
        // complete composed-CFB reopen. Retaining this plan's source digest
        // gives the semantic patch a stronger foreign/stale-source guard than
        // a version token alone.
        let identity_plan = publisher
            .plan_splices(Vec::new(), options.splice_limits)
            .map_err(Error::Source)?;
        let fingerprint = identity_plan.source_fingerprint();
        let shared = publisher.shared();
        ensure_version(&source, version)?;
        let observed_length = source
            .len()
            .map_err(|error| Error::Source(OverlayError::Io(error)))?;
        if observed_length != length {
            return Err(Error::Source(OverlayError::Unavailable {
                reason: format!(
                    "PPT source length changed from {length} to {observed_length} during open"
                ),
            }));
        }
        Ok(Self {
            inner: Arc::new(SourceInner {
                source,
                shared,
                version,
                length,
                fingerprint,
                options,
            }),
        })
    }

    /// Opaque source version captured at open.
    #[must_use]
    pub fn source_version(&self) -> SourceVersion {
        self.inner.version
    }

    /// Complete source-artifact fingerprint captured at open.
    #[must_use]
    pub fn fingerprint(&self) -> ArtifactFingerprint {
        self.inner.fingerprint
    }

    /// Complete source length captured at open.
    #[must_use]
    pub fn len(&self) -> u64 {
        self.inner.length
    }

    /// Whether the source has no bytes.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.inner.length == 0
    }

    /// Resolves one existing slide/shape text atom using semantic positions.
    ///
    /// The returned transaction admits only an equal encoded-length
    /// replacement. There is no fallback to the ordinary full-record editor.
    pub(crate) fn edit_text(&self, target: Target) -> Result<SourceTransaction> {
        let resolved = self.resolve_source_text(target)?;
        Ok(SourceTransaction {
            source: self.clone(),
            resolved,
            replacement: None,
        })
    }

    /// Reads one existing slide/shape text atom using semantic positions.
    ///
    /// The returned text is owned, so the source-backed selector does not
    /// expose a source lifetime or any physical CFB identity. Resolution uses
    /// the bounded positional path and refusal checks of the source-backed
    /// text owner.
    pub fn read_text(&self, target: Target) -> Result<String> {
        Ok(self.resolve_source_text(target)?.text)
    }

    fn resolve_source_text(&self, target: Target) -> Result<SourceResolved> {
        self.ensure_current()?;
        let shared = open_shared_source(self)?;
        reject_macro_components(&shared)?;
        let document_path = select_stream_path(&shared, &DOCUMENT_PATHS, "PowerPoint Document")?;
        let current_user_path = select_stream_path(&shared, &CURRENT_USER_PATHS, "Current User")?;
        ensure_matching_stream_topology(&document_path, &current_user_path)?;
        let resolved = source_range::resolve_source_target(
            &shared,
            &document_path,
            &current_user_path,
            target,
            self.inner.options.source_limits,
        )?;
        self.ensure_current()?;
        Ok(resolved)
    }

    fn ensure_current(&self) -> Result<()> {
        ensure_version(&self.inner.source, self.inner.version)?;
        let length = self
            .inner
            .source
            .len()
            .map_err(|error| Error::Source(OverlayError::Io(error)))?;
        if length != self.inner.length {
            return Err(Error::Source(OverlayError::Unavailable {
                reason: format!(
                    "PPT source length changed from {} to {length}",
                    self.inner.length
                ),
            }));
        }
        ensure_version(&self.inner.source, self.inner.version)
    }
}

/// One isolated equal-length text replacement staged against a positional
/// PPT source.
#[derive(Clone)]
pub(crate) struct SourceTransaction {
    source: SourceSnapshot,
    resolved: SourceResolved,
    replacement: Option<String>,
}

impl fmt::Debug for SourceTransaction {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("SourceTransaction")
            .field("target", &self.resolved.target)
            .field("encoded_length", &self.resolved.payload.len())
            .field("staged", &self.replacement.is_some())
            .finish()
    }
}

impl SourceTransaction {
    /// Exact semantic text currently read from the selected atom.
    #[must_use]
    pub(crate) fn text(&self) -> &str {
        &self.resolved.text
    }

    /// Selected semantic slide/shape position.
    #[must_use]
    pub(crate) const fn target(&self) -> Target {
        self.resolved.target
    }

    /// Stages a replacement only when its encoded atom length is unchanged.
    ///
    /// Failed validation leaves the transaction unchanged. The ordinary
    /// source-owning editor remains separate and is not used as a fallback.
    pub(crate) fn set_text(&mut self, value: impl Into<String>) -> Result<()> {
        let candidate = value.into();
        let encoded_length = encoded_replacement_len(&candidate, self.resolved.kind)?;
        if encoded_length != self.resolved.payload.len() {
            return Err(Error::Refused(Refusal::LengthChange));
        }
        self.replacement = Some(candidate);
        Ok(())
    }

    /// Publishes the bounded source-backed splice after complete candidate
    /// CFB reopen and selected-slide semantic readback.
    pub(crate) fn commit(self) -> Result<SourceBackedCommit> {
        let replacement = self
            .replacement
            .unwrap_or_else(|| self.resolved.text.clone());
        publish_source_text(self.source, self.resolved, replacement)
    }
}

/// Content-free diagnostics for one source-backed PPT text publication.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) struct SourceBackedDiagnostics {
    target: Target,
    changed_spans: usize,
    replacement_bytes: u64,
    source_bytes: u64,
    source_fingerprint: ArtifactFingerprint,
    target_fingerprint: ArtifactFingerprint,
    source_version: SourceVersion,
    target_version: SourceVersion,
}

impl SourceBackedDiagnostics {
    /// Semantic selector used by the operation.
    #[must_use]
    pub(crate) const fn target(self) -> Target {
        self.target
    }

    /// Number of physical CFB spans changed by the splice.
    #[must_use]
    pub(crate) const fn changed_spans(self) -> usize {
        self.changed_spans
    }

    /// Replacement bytes submitted to the CFB splice planner.
    #[must_use]
    pub(crate) const fn replacement_bytes(self) -> u64 {
        self.replacement_bytes
    }

    /// Complete source CFB length.
    #[must_use]
    pub(crate) const fn source_bytes(self) -> u64 {
        self.source_bytes
    }

    /// Exact source artifact fingerprint.
    #[must_use]
    pub(crate) const fn source_fingerprint(self) -> ArtifactFingerprint {
        self.source_fingerprint
    }

    /// Exact composed target artifact fingerprint.
    #[must_use]
    pub(crate) const fn target_fingerprint(self) -> ArtifactFingerprint {
        self.target_fingerprint
    }

    /// Source version checked before planning.
    #[must_use]
    pub(crate) const fn source_version(self) -> SourceVersion {
        self.source_version
    }

    /// Composed target version exposed by the positional overlay.
    #[must_use]
    pub(crate) const fn target_version(self) -> SourceVersion {
        self.target_version
    }
}

/// A published source-backed PPT target and its reusable overlay plan.
pub(crate) struct SourceBackedCommit {
    snapshot: SourceSnapshot,
    patch: SourceBackedPatch,
    plan: ValidatedOverlayPlan,
    diagnostics: SourceBackedDiagnostics,
}

impl fmt::Debug for SourceBackedCommit {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("SourceBackedCommit")
            .field("snapshot", &self.snapshot)
            .field("plan", &self.plan)
            .field("diagnostics", &self.diagnostics)
            .finish()
    }
}

impl SourceBackedCommit {
    /// Reopened target source snapshot.
    #[must_use]
    pub(crate) const fn snapshot(&self) -> &SourceSnapshot {
        &self.snapshot
    }

    /// Exact-source reversible patch.
    #[must_use]
    pub(crate) const fn patch(&self) -> &SourceBackedPatch {
        &self.patch
    }

    /// Common validated CFB overlay plan.
    #[must_use]
    pub(crate) const fn plan(&self) -> &ValidatedOverlayPlan {
        &self.plan
    }

    /// Content-free publication diagnostics.
    #[must_use]
    pub(crate) const fn diagnostics(&self) -> SourceBackedDiagnostics {
        self.diagnostics
    }

    /// Whether the selected replacement is an exact byte no-op.
    #[must_use]
    pub(crate) fn is_noop(&self) -> bool {
        self.plan.is_noop()
    }

    /// Streams the complete validated target to a sequential sink.
    pub(crate) fn write_to<W: Write>(&self, writer: &mut W) -> Result<PublishReport> {
        self.plan.write_to(writer).map_err(Error::Source)
    }

    /// Alias emphasizing the forward-only publication boundary.
    pub(crate) fn publish_to_stream<W: Write>(&self, writer: &mut W) -> Result<PublishReport> {
        self.write_to(writer)
    }

    /// Publishes through the common synced sibling-temp/atomic-rename path.
    pub(crate) fn save<P: AsRef<std::path::Path>>(&self, path: P) -> Result<PublishReport> {
        self.plan.save(path).map_err(Error::Source)
    }

    /// Publishes through the common sibling-temp/atomic-rename path at a
    /// caller-chosen [`Durability`](litchi_core::Durability); see
    /// [`litchi_cfb::ValidatedOverlayPlan::save_with_durability`].
    pub(crate) fn save_with_durability<P: AsRef<std::path::Path>>(
        &self,
        path: P,
        durability: litchi_core::Durability,
    ) -> Result<PublishReport> {
        self.plan
            .save_with_durability(path, durability)
            .map_err(Error::Source)
    }
}

/// A source-checked reversible equal-length PPT text splice.
#[derive(Clone)]
pub(crate) struct SourceBackedPatch {
    before: SourceSnapshot,
    after: SourceSnapshot,
    operation: Option<SourceOperation>,
}

impl fmt::Debug for SourceBackedPatch {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("SourceBackedPatch")
            .field("source_version", &self.before.source_version())
            .field("target_version", &self.after.source_version())
            .field("has_operation", &self.operation.is_some())
            .finish()
    }
}

impl SourceBackedPatch {
    /// Exact source snapshot required for forward application.
    #[must_use]
    pub(crate) const fn source(&self) -> &SourceSnapshot {
        &self.before
    }

    /// Exact target snapshot produced by forward application.
    #[must_use]
    pub(crate) const fn target(&self) -> &SourceSnapshot {
        &self.after
    }

    /// Exact source artifact fingerprint required for forward application.
    #[must_use]
    pub(crate) fn source_fingerprint(&self) -> ArtifactFingerprint {
        self.before.fingerprint()
    }

    /// Exact target artifact fingerprint produced by forward application.
    #[must_use]
    pub(crate) fn target_fingerprint(&self) -> ArtifactFingerprint {
        self.after.fingerprint()
    }

    /// Whether the patch changes no bytes.
    #[must_use]
    pub(crate) fn is_empty(&self) -> bool {
        self.operation.is_none()
    }

    /// Applies only to the exact source version and fingerprint captured by
    /// this patch. The candidate is reopened and semantically read back again.
    pub(crate) fn apply(&self, current: &SourceSnapshot) -> Result<SourceSnapshot> {
        if current.source_version() != self.before.source_version() {
            return Err(Error::Source(OverlayError::SourceChanged {
                expected: self.before.source_version(),
                observed: current.source_version(),
            }));
        }
        if current.fingerprint() != self.before.fingerprint() {
            return Err(Error::Source(OverlayError::SourceFingerprintChanged {
                expected: self.before.fingerprint(),
                observed: current.fingerprint(),
            }));
        }
        let Some(operation) = self.operation.clone() else {
            current.ensure_identity()?;
            return Ok(current.clone());
        };
        let commit = publish_source_operation(current.clone(), operation)?;
        if commit.snapshot.fingerprint() != self.after.fingerprint() {
            return Err(Error::Source(OverlayError::TargetFingerprintChanged {
                expected: self.after.fingerprint(),
                observed: commit.snapshot.fingerprint(),
            }));
        }
        Ok(commit.snapshot)
    }

    /// Returns the exact-source-checked inverse operation.
    #[must_use]
    pub(crate) fn inverse(&self) -> Self {
        Self {
            before: self.after.clone(),
            after: self.before.clone(),
            operation: self.operation.as_ref().map(SourceOperation::inverse),
        }
    }
}

#[derive(Clone)]
struct SourceOperation {
    target: Target,
    document_path: Vec<String>,
    current_user_path: Vec<String>,
    atom_offset: u64,
    expected: Arc<[u8]>,
    replacement: Arc<[u8]>,
    before_text: String,
    after_text: String,
    slide_persist_id: u32,
    atom_kind: TextKind,
}

impl SourceOperation {
    fn inverse(&self) -> Self {
        Self {
            target: self.target,
            document_path: self.document_path.clone(),
            current_user_path: self.current_user_path.clone(),
            atom_offset: self.atom_offset,
            expected: Arc::clone(&self.replacement),
            replacement: Arc::clone(&self.expected),
            before_text: self.after_text.clone(),
            after_text: self.before_text.clone(),
            slide_persist_id: self.slide_persist_id,
            atom_kind: self.atom_kind,
        }
    }
}

fn publish_source_text(
    source: SourceSnapshot,
    resolved: SourceResolved,
    replacement: String,
) -> Result<SourceBackedCommit> {
    // The source text is already known to be representable. Avoid even a
    // replacement-encoding allocation for an exact semantic no-op, retaining
    // the source allocation and identity all the way through publication.
    if replacement == resolved.text {
        return publish_source_noop(source, resolved.target);
    }

    // Compute representability and encoded size without allocating. The
    // splice byte limit must be checked before the replacement buffer is
    // materialized, so hostile callers cannot turn a rejected over-limit edit
    // into an avoidable allocation.
    let encoded_length = encoded_replacement_len(&replacement, resolved.kind)?;
    if encoded_length != resolved.payload.len() {
        return Err(Error::Refused(Refusal::LengthChange));
    }
    let max_splice_bytes = source.inner.options.splice_limits.max_splice_bytes();
    let encoded_length_u64 = u64::try_from(encoded_length).map_err(|_error| {
        PackageError::ResourceLimit("PPT encoded replacement length does not fit u64".into())
    })?;
    if encoded_length_u64 > max_splice_bytes {
        return Err(Error::Source(OverlayError::Unavailable {
            reason: format!(
                "PPT replacement length {encoded_length_u64} exceeds splice limit {max_splice_bytes}"
            ),
        }));
    }
    let encoded = encode_replacement(&replacement, resolved.kind)?;
    if encoded.len() != resolved.payload.len() {
        return Err(PackageError::Corrupted(
            "PPT source-backed replacement size preflight disagrees with encoding".into(),
        )
        .into());
    }

    let atom_offset = u64::try_from(resolved.atom_offset).map_err(|_error| {
        PackageError::ResourceLimit("PPT source-backed atom offset does not fit u64".into())
    })?;
    let operation = SourceOperation {
        target: resolved.target,
        document_path: resolved.document_path.clone(),
        current_user_path: resolved.current_user_path.clone(),
        atom_offset,
        expected: Arc::clone(&resolved.payload),
        replacement: Arc::from(encoded),
        before_text: resolved.text.clone(),
        after_text: replacement,
        slide_persist_id: resolved.slide_persist_id,
        atom_kind: resolved.kind,
    };
    publish_source_operation(source, operation)
}

fn publish_source_noop(source: SourceSnapshot, target: Target) -> Result<SourceBackedCommit> {
    source.ensure_current()?;
    let shared_limits = shared_limits(source.inner.options.source_limits)?;
    let publisher =
        SourceBackedOverlayPublisher::open(Arc::clone(&source.inner.source), shared_limits)
            .map_err(Error::Source)?;
    let plan = publisher
        .plan_splices(Vec::new(), source.inner.options.splice_limits)
        .map_err(Error::Source)?;
    if plan.source_fingerprint() != source.fingerprint() {
        return Err(Error::Source(OverlayError::SourceFingerprintChanged {
            expected: source.fingerprint(),
            observed: plan.source_fingerprint(),
        }));
    }
    let diagnostics = SourceBackedDiagnostics {
        target,
        changed_spans: plan.changed_spans(),
        replacement_bytes: 0,
        source_bytes: source.len(),
        source_fingerprint: plan.source_fingerprint(),
        target_fingerprint: plan.target_fingerprint(),
        source_version: source.source_version(),
        target_version: source.source_version(),
    };
    let patch = SourceBackedPatch {
        before: source.clone(),
        after: source.clone(),
        operation: None,
    };
    Ok(SourceBackedCommit {
        snapshot: source,
        patch,
        plan,
        diagnostics,
    })
}

fn publish_source_operation(
    source: SourceSnapshot,
    operation: SourceOperation,
) -> Result<SourceBackedCommit> {
    source.ensure_current()?;
    let shared_limits = shared_limits(source.inner.options.source_limits)?;
    let publisher =
        SourceBackedOverlayPublisher::open(Arc::clone(&source.inner.source), shared_limits)
            .map_err(Error::Source)?;
    let observed_version = publisher.source_version().map_err(Error::Source)?;
    if observed_version != source.source_version() {
        return Err(Error::Source(OverlayError::SourceChanged {
            expected: source.source_version(),
            observed: observed_version,
        }));
    }
    let (plan, candidate) = publisher.plan_splices_with_owner(
        vec![SameLengthStreamSplice::new(
            operation.document_path.clone(),
            operation.atom_offset,
            Arc::clone(&operation.expected),
            Arc::clone(&operation.replacement),
        )],
        source.inner.options.splice_limits,
        |composed| -> Result<(Arc<dyn ReadAt>, SharedOleFile, SourceVersion)> {
            let candidate: Arc<dyn ReadAt> = Arc::new(composed.clone());
            let candidate_shared = SharedOleFile::open(Arc::clone(&candidate), shared_limits)
                .map_err(map_ole_error)?;
            reject_macro_components(&candidate_shared)?;
            let candidate_resolved = source_range::resolve_source_target(
                &candidate_shared,
                &operation.document_path,
                &operation.current_user_path,
                operation.target,
                source.inner.options.source_limits,
            )?;
            verify_source_readback(&operation, &candidate_resolved)?;
            let target_version = candidate
                .version()
                .map_err(|error| Error::Source(OverlayError::Io(error)))?;
            Ok((candidate, candidate_shared, target_version))
        },
    )?;
    if plan.source_fingerprint() != source.fingerprint() {
        return Err(Error::Source(OverlayError::SourceFingerprintChanged {
            expected: source.fingerprint(),
            observed: plan.source_fingerprint(),
        }));
    }

    let (candidate, candidate_shared, target_version) = candidate.ok_or_else(|| {
        Error::Source(OverlayError::Unavailable {
            reason: "changed PPT splice did not produce an owner candidate".into(),
        })
    })?;
    let target = SourceSnapshot {
        inner: Arc::new(SourceInner {
            source: candidate,
            shared: Arc::new(candidate_shared),
            version: target_version,
            length: source.len(),
            fingerprint: plan.target_fingerprint(),
            options: source.inner.options,
        }),
    };
    let diagnostics = SourceBackedDiagnostics {
        target: operation.target,
        changed_spans: plan.changed_spans(),
        replacement_bytes: u64::try_from(operation.replacement.len()).map_err(|_error| {
            PackageError::ResourceLimit("PPT replacement length does not fit u64".into())
        })?,
        source_bytes: source.len(),
        source_fingerprint: plan.source_fingerprint(),
        target_fingerprint: plan.target_fingerprint(),
        source_version: source.source_version(),
        target_version,
    };
    let patch = SourceBackedPatch {
        before: source,
        after: target.clone(),
        operation: Some(operation),
    };
    Ok(SourceBackedCommit {
        snapshot: target,
        patch,
        plan,
        diagnostics,
    })
}

fn verify_source_readback(operation: &SourceOperation, resolved: &SourceResolved) -> Result<()> {
    let atom_offset = u64::try_from(resolved.atom_offset).map_err(|_error| {
        PackageError::ResourceLimit("PPT source-backed atom offset does not fit u64".into())
    })?;
    if resolved.slide_persist_id != operation.slide_persist_id
        || atom_offset != operation.atom_offset
        || resolved.kind != operation.atom_kind
        || resolved.payload.len() != operation.replacement.len()
        || resolved.payload.as_ref() != operation.replacement.as_ref()
        || resolved.text != operation.after_text
    {
        return Err(PackageError::Corrupted(
            "published PPT source-backed shape text failed semantic readback".into(),
        )
        .into());
    }
    Ok(())
}

fn shared_limits(source_limits: crate::SourceLimits) -> Result<SharedOleFileLimits> {
    let maximum = u64::try_from(source_limits.max_package_bytes).unwrap_or(u64::MAX);
    SharedOleFileLimits::new(maximum).map_err(map_ole_error)
}

fn ensure_version(source: &Arc<dyn ReadAt>, expected: SourceVersion) -> Result<()> {
    let observed = source
        .version()
        .map_err(|error| Error::Source(OverlayError::Io(error)))?;
    if observed != expected {
        return Err(Error::Source(OverlayError::SourceChanged {
            expected,
            observed,
        }));
    }
    Ok(())
}

fn open_shared_source(source: &SourceSnapshot) -> Result<Arc<SharedOleFile>> {
    Ok(Arc::clone(&source.inner.shared))
}

impl SourceSnapshot {
    fn ensure_identity(&self) -> Result<()> {
        self.ensure_current()?;
        let shared_limits = shared_limits(self.inner.options.source_limits)?;
        let publisher =
            SourceBackedOverlayPublisher::open(Arc::clone(&self.inner.source), shared_limits)
                .map_err(Error::Source)?;
        let plan = publisher
            .plan_splices(Vec::new(), self.inner.options.splice_limits)
            .map_err(Error::Source)?;
        if plan.source_fingerprint() != self.inner.fingerprint {
            return Err(Error::Source(OverlayError::SourceFingerprintChanged {
                expected: self.inner.fingerprint,
                observed: plan.source_fingerprint(),
            }));
        }
        Ok(())
    }
}

fn map_ole_error(error: OleError) -> Error {
    match error {
        OleError::SourceChanged { expected, observed } => {
            Error::Source(OverlayError::SourceChanged { expected, observed })
        },
        other => Error::Package(PackageError::Ole(other)),
    }
}

fn select_stream_path(
    shared: &SharedOleFile,
    candidates: &[&[&str]],
    label: &str,
) -> Result<Vec<String>> {
    let mut selected = None;
    for candidate in candidates {
        match shared.stream_len(candidate) {
            Ok(_) => {
                if selected.is_some() {
                    return Err(Error::Refused(Refusal::UnsupportedDependency));
                }
                selected = Some(
                    candidate
                        .iter()
                        .map(|component| (*component).to_string())
                        .collect(),
                );
            },
            Err(OleError::StreamNotFound) => {},
            Err(error) => return Err(map_ole_error(error)),
        }
    }
    selected.ok_or_else(|| PackageError::StreamNotFound(label.to_string()).into())
}

fn ensure_matching_stream_topology(document: &[String], current_user: &[String]) -> Result<()> {
    let Some((_document_leaf, document_parent)) = document.split_last() else {
        return Err(Error::Refused(Refusal::UnsupportedDependency));
    };
    let Some((_current_user_leaf, current_user_parent)) = current_user.split_last() else {
        return Err(Error::Refused(Refusal::UnsupportedDependency));
    };
    if document_parent != current_user_parent {
        return Err(Error::Refused(Refusal::UnsupportedDependency));
    }
    Ok(())
}

#[cfg(test)]
#[allow(dead_code, reason = "retained as the owned differential oracle")]
fn read_bounded_stream(
    shared: &SharedOleFile,
    path: &[String],
    limits: crate::SourceLimits,
    _label: &str,
) -> Result<Vec<u8>> {
    let refs: Vec<_> = path.iter().map(String::as_str).collect();
    let length = shared.stream_len(&refs).map_err(map_ole_error)?;
    let length_usize = usize::try_from(length).map_err(|_error| PackageError::Limit {
        kind: crate::SourceLimitKind::EntryBytes,
        observed: usize::MAX,
        maximum: limits.max_input_bytes,
    })?;
    if length_usize > limits.max_input_bytes {
        return Err(PackageError::Limit {
            kind: crate::SourceLimitKind::EntryBytes,
            observed: length_usize,
            maximum: limits.max_input_bytes,
        }
        .into());
    }
    let mut bytes = Vec::new();
    bytes
        .try_reserve_exact(length_usize)
        .map_err(|_error| PackageError::AllocationFailed("PPT source-backed stream"))?;
    bytes.resize(length_usize, 0);
    shared
        .read_stream_range(&refs, 0, &mut bytes)
        .map_err(map_ole_error)?;
    Ok(bytes)
}

fn reject_macro_components(shared: &SharedOleFile) -> Result<()> {
    let mut dual_storages = 0usize;
    for entry in shared.directory_entries() {
        if is_macro_storage_name(&entry.name) {
            return Err(Error::Refused(Refusal::UnsupportedSource));
        }
        // The source-backed tranche has no embedded-object/storage rewrite
        // closure. The root and the one legacy dual-layout storage are the
        // only accepted storage topology; every other storage is an explicit
        // conservative refusal, including `ObjectPool` and nested OLE data.
        if entry.entry_type == litchi_cfb::consts::STGTY_STORAGE {
            if entry.name.eq_ignore_ascii_case("PP97_DUALSTORAGE") {
                dual_storages = dual_storages.saturating_add(1);
                if dual_storages > 1 {
                    return Err(Error::Refused(Refusal::UnsupportedDependency));
                }
            } else {
                return Err(Error::Refused(Refusal::UnsupportedDependency));
            }
        }
    }
    Ok(())
}

fn is_macro_storage_name(name: &str) -> bool {
    ["VBA", "VBAProject", "VbaProjectStg", "_VBA_PROJECT"]
        .iter()
        .any(|marker| name.eq_ignore_ascii_case(marker))
}

#[cfg(test)]
fn reject_macro_records(
    document: &[u8],
    document_offset: usize,
    limits: crate::SourceLimits,
) -> Result<()> {
    let (record, _consumed) =
        crate::Record::parse_strict_with_limits(document, document_offset, limits)?;
    if record.record_type != RecordType::Document || record.version != 0x0F || record.instance != 0
    {
        return Err(PackageError::Corrupted(
            "live PPT DocumentContainer has an invalid macro-scan header".into(),
        )
        .into());
    }
    inspect_live_macro_records(&record, false, false, &mut LiveMacroOwners::default())
}

#[derive(Default)]
struct LiveMacroOwners {
    doc_info_lists: usize,
    vba_infos: usize,
}

fn inspect_live_macro_records(
    record: &crate::Record,
    in_doc_info_list: bool,
    in_vba_info: bool,
    owners: &mut LiveMacroOwners,
) -> Result<()> {
    if record.record_type_raw == RecordType::DocInfoList as u16 {
        owners.doc_info_lists = owners.doc_info_lists.saturating_add(1);
        if owners.doc_info_lists > 1 {
            return Err(PackageError::Corrupted(
                "live PPT DocumentContainer has multiple DocInfoList owners".into(),
            )
            .into());
        }
    }
    match record.record_type_raw {
        value if value == RecordType::VBAInfo as u16 && in_doc_info_list => {
            owners.vba_infos = owners.vba_infos.saturating_add(1);
            if owners.vba_infos > 1 {
                return Err(PackageError::Corrupted(
                    "live PPT DocInfoList has multiple VBAInfo owners".into(),
                )
                .into());
            }
            let [atom] = record.children.as_slice() else {
                return Err(PackageError::Corrupted(
                    "live PPT VBAInfo container has invalid child ownership".into(),
                )
                .into());
            };
            if record.version != 0x0F
                || record.instance != 1
                || record.data.len() != 20
                || atom.record_type_raw != RecordType::VBAInfoAtom as u16
                || atom.version != 2
                || atom.instance != 0
                || atom.data.len() != 12
            {
                return Err(PackageError::Corrupted(
                    "live PPT VBAInfo metadata has an invalid record shape".into(),
                )
                .into());
            }
            let persist_id = macro_info_u32(&atom.data, 0)?;
            let has_macros = match macro_info_u32(&atom.data, 4)? {
                0 => false,
                1 => true,
                _ => {
                    return Err(PackageError::Corrupted(
                        "live PPT VBAInfoAtom has an invalid macro flag".into(),
                    )
                    .into());
                },
            };
            if macro_info_u32(&atom.data, 8)? != 2 || (has_macros && persist_id == 0) {
                return Err(PackageError::Corrupted(
                    "live PPT VBAInfoAtom has inconsistent metadata".into(),
                )
                .into());
            }
            if persist_id != 0 || has_macros {
                return Err(Error::Refused(Refusal::UnsupportedSource));
            }
            return Ok(());
        },
        value if value == RecordType::VBAInfoAtom as u16 && (in_doc_info_list || in_vba_info) => {
            if record.version != 2 || record.instance != 0 || record.data.len() != 12 {
                return Err(PackageError::Corrupted(
                    "live PPT VBAInfoAtom has an invalid record header".into(),
                )
                .into());
            }
            return Err(Error::Refused(Refusal::UnsupportedSource));
        },
        _ => {},
    }
    let child_doc_info = record.record_type_raw == RecordType::DocInfoList as u16;
    let child_vba_info = record.record_type_raw == RecordType::VBAInfo as u16;
    for child in &record.children {
        inspect_live_macro_records(child, child_doc_info, child_vba_info, owners)?;
    }
    Ok(())
}

fn macro_info_u32(data: &[u8], offset: usize) -> Result<u32> {
    let bytes: [u8; 4] = data
        .get(offset..offset + 4)
        .ok_or_else(|| PackageError::Corrupted("live PPT VBAInfoAtom is truncated".into()))?
        .try_into()
        .map_err(|_error| {
            PackageError::Corrupted("live PPT VBAInfoAtom field has an invalid width".into())
        })?;
    Ok(u32::from_le_bytes(bytes))
}

const DOCUMENT_PATHS: [&[&str]; 2] = [
    &["PowerPoint Document"],
    &["PP97_DUALSTORAGE", "PowerPoint Document"],
];
const CURRENT_USER_PATHS: [&[&str]; 2] = [&["Current User"], &["PP97_DUALSTORAGE", "Current User"]];

#[derive(Debug, Clone)]
struct SourceResolved {
    target: Target,
    document_path: Vec<String>,
    current_user_path: Vec<String>,
    slide_persist_id: u32,
    atom_offset: usize,
    kind: TextKind,
    payload: Arc<[u8]>,
    text: String,
}

#[cfg(test)]
#[allow(dead_code, reason = "retained as the owned differential oracle")]
fn resolve_source_target_owned(
    shared: &SharedOleFile,
    document_path: &[String],
    current_user_path: &[String],
    target: Target,
    limits: crate::SourceLimits,
) -> Result<SourceResolved> {
    let document_refs: Vec<_> = document_path.iter().map(String::as_str).collect();
    let current_user_refs: Vec<_> = current_user_path.iter().map(String::as_str).collect();
    let document_length = shared.stream_len(&document_refs).map_err(map_ole_error)?;
    let current_user_length = shared
        .stream_len(&current_user_refs)
        .map_err(map_ole_error)?;
    let aggregate_length =
        document_length
            .checked_add(current_user_length)
            .ok_or(PackageError::Limit {
                kind: crate::SourceLimitKind::AggregateBytes,
                observed: usize::MAX,
                maximum: limits.max_aggregate_input_bytes,
            })?;
    let aggregate_limit = u64::try_from(limits.max_aggregate_input_bytes).unwrap_or(u64::MAX);
    if aggregate_length > aggregate_limit {
        return Err(PackageError::Limit {
            kind: crate::SourceLimitKind::AggregateBytes,
            observed: usize::try_from(aggregate_length).unwrap_or(usize::MAX),
            maximum: limits.max_aggregate_input_bytes,
        }
        .into());
    }
    // Shared so the selected slide's record payloads span this buffer instead
    // of being copied out of it; every reader below still sees `&[u8]`.
    let document = Arc::new(read_bounded_stream(
        shared,
        document_path,
        limits,
        "PowerPoint Document",
    )?);
    let current_user = read_bounded_stream(shared, current_user_path, limits, "Current User")?;
    let current = crate::CurrentUser::parse_with_limits(&current_user, limits)?;
    if current.is_encrypted() {
        return Err(Error::Refused(Refusal::UnsupportedSource));
    }
    let mapping = crate::embedded::object::Editor::inspect_live_mapping(&document, &current_user)?;
    // The edit path resolves slides exactly as the read path does, applying
    // the same slide-list repairs without enumerating them.
    let directory = crate::slide::SlideDirectory::build_with_limits(
        &document,
        &current_user,
        &mapping,
        limits,
        &mut crate::leniency::ToleranceReport::default(),
    )?;
    // The writer emits a valid no-macro VBAInfo sentinel in the live
    // DocumentContainer. Inspect only this structural record tree; do not
    // reinterpret opaque payload bytes or historical DocumentContainers as
    // nested VBA metadata.
    reject_macro_records(&document, directory.document_offset(), limits)?;
    let slide_position = target.slide.get();
    let entry = directory
        .entries()
        .get(slide_position)
        .ok_or(Error::Refused(Refusal::SlideNotFound {
            position: target.slide,
        }))?;
    let factory =
        crate::slide::SlideFactory::new_shared_with_limits(&document, &mapping, &directory, limits);
    let slide_data = factory.parse_slide(entry.persist_id())?;
    let slide_offset = slide_data.offset;
    let slide = crate::Slide::from_slide_data(slide_data, slide_position + 1);
    let shape = slide
        .shapes()?
        .get(target.shape.get())
        .ok_or(Error::Refused(Refusal::ShapeNotFound))?;
    let shape_id = native_shape_id(shape);
    let atom = inspect_source_text_atom(&document, slide_offset, shape_id, limits)?;
    Ok(SourceResolved {
        target,
        document_path: document_path.to_vec(),
        current_user_path: current_user_path.to_vec(),
        slide_persist_id: entry.persist_id(),
        atom_offset: atom.offset,
        kind: atom.kind,
        payload: atom.payload,
        text: atom.text,
    })
}

#[derive(Debug)]
struct SourceTextAtom {
    offset: usize,
    kind: TextKind,
    payload: Arc<[u8]>,
    text: String,
}

#[cfg(test)]
#[allow(dead_code, reason = "retained as the owned differential oracle")]
fn inspect_source_text_atom(
    document: &[u8],
    slide_offset: usize,
    shape_id: u32,
    limits: crate::SourceLimits,
) -> Result<SourceTextAtom> {
    let (slide, consumed) =
        crate::Record::parse_strict_with_limits(document, slide_offset, limits)?;
    let slide_end = slide_offset
        .checked_add(consumed)
        .ok_or_else(|| PackageError::Corrupted("PPT slide range overflows".into()))?;
    let slide_bytes = document
        .get(slide_offset..slide_end)
        .ok_or_else(|| PackageError::Corrupted("PPT slide range exceeds Document".into()))?;

    inspect_source_text_atom_parts(slide_bytes, slide_offset, &slide, shape_id)
}

fn inspect_source_text_atom_parts(
    slide_bytes: &[u8],
    slide_offset: usize,
    slide: &crate::Record,
    shape_id: u32,
) -> Result<SourceTextAtom> {
    if slide.record_type != RecordType::Slide {
        return Err(
            PackageError::Corrupted("selected persist record is not a Slide".into()).into(),
        );
    }
    let mut drawings = 0usize;
    let mut matches = Vec::new();
    find_source_drawing(
        slide_bytes,
        slide_offset,
        shape_id,
        0,
        &mut drawings,
        &mut matches,
    )?;
    if drawings != 1 {
        return Err(PackageError::Corrupted(
            "selected slide has ambiguous PPDrawing ownership".into(),
        )
        .into());
    }
    match matches.len() {
        0 => Err(Error::Refused(Refusal::ShapeNotFound)),
        1 => matches
            .pop()
            .ok_or_else(|| PackageError::Corrupted("selected text atom disappeared".into()).into()),
        _ => Err(Error::Refused(Refusal::AmbiguousShape)),
    }
}

fn find_source_drawing(
    record: &[u8],
    record_offset: usize,
    shape_id: u32,
    depth: usize,
    drawings: &mut usize,
    matches: &mut Vec<SourceTextAtom>,
) -> Result<()> {
    if depth >= 128 {
        return Err(PackageError::Corrupted(
            "PPT source-backed record nesting exceeds limit".into(),
        )
        .into());
    }
    if record_type(record)? == RecordType::PPDrawing as u16 {
        validate_source_ppdrawing_header(record)?;
        *drawings = drawings.saturating_add(1);
        visit_source_officeart_stream(
            drawing_payload(record)?,
            record_offset
                .checked_add(PPT_HEADER_LEN)
                .ok_or_else(|| PackageError::Corrupted("PPDrawing offset overflows".into()))?,
            shape_id,
            matches,
        )?;
        return Ok(());
    }
    if record_version(record)? != 0xF {
        return Ok(());
    }
    for child in children_with_offsets(record, record_offset)? {
        let (offset, child) = child?;
        find_source_drawing(child, offset, shape_id, depth + 1, drawings, matches)?;
    }
    Ok(())
}

fn visit_source_officeart_stream(
    payload: &[u8],
    payload_offset: usize,
    shape_id: u32,
    matches: &mut Vec<SourceTextAtom>,
) -> Result<()> {
    let roots = ChildrenWithOffsets {
        payload,
        payload_offset,
        offset: 0,
    };
    for root in roots {
        let (offset, record) = root?;
        visit_source_officeart(record, offset, shape_id, 0, matches)?;
    }
    Ok(())
}

fn visit_source_officeart(
    record: &[u8],
    record_offset: usize,
    shape_id: u32,
    depth: usize,
    matches: &mut Vec<SourceTextAtom>,
) -> Result<()> {
    if depth >= 128 {
        return Err(PackageError::Corrupted(
            "OfficeArt source-backed nesting exceeds limit".into(),
        )
        .into());
    }
    let source_record_type = record_type(record)?;
    if record_version(record)? == 0x0F {
        validate_source_officeart_container(record)?;
    }
    if source_record_type == OFFICEART_SP_CONTAINER {
        validate_source_officeart_container(record)?;
    }
    if source_record_type == OFFICEART_SP_CONTAINER && shape_id_of(record)? == Some(shape_id) {
        let mut textboxes = Vec::new();
        for child in children_with_offsets(record, record_offset)? {
            let (offset, child) = child?;
            if record_type(child)? == OFFICEART_CLIENT_TEXTBOX {
                validate_source_officeart_container(child)?;
                textboxes.push((offset, child));
            }
        }
        match textboxes.as_slice() {
            [] => return Err(Error::Refused(Refusal::NoTextbox)),
            [(textbox_offset, textbox)] => {
                matches.push(inspect_source_textbox(textbox, *textbox_offset)?);
            },
            _ => return Err(Error::Refused(Refusal::AmbiguousTextbox)),
        }
        return Ok(());
    }
    if record_version(record)? != 0xF {
        return Ok(());
    }
    for child in children_with_offsets(record, record_offset)? {
        let (offset, child) = child?;
        visit_source_officeart(child, offset, shape_id, depth + 1, matches)?;
    }
    Ok(())
}

fn validate_source_ppdrawing_header(record: &[u8]) -> Result<()> {
    if record_version(record)? != 0x0F || record_instance(record)? != 0 {
        return Err(PackageError::Corrupted(
            "source-backed PPDrawing has an invalid container header".into(),
        )
        .into());
    }
    Ok(())
}

fn validate_source_officeart_container(record: &[u8]) -> Result<()> {
    if record_version(record)? != 0x0F || record_instance(record)? != 0 {
        return Err(PackageError::Corrupted(
            "source-backed OfficeArt container has an invalid container header".into(),
        )
        .into());
    }
    Ok(())
}

fn inspect_source_textbox(textbox: &[u8], textbox_offset: usize) -> Result<SourceTextAtom> {
    let mut atoms = Vec::new();
    let mut child_types = Vec::new();
    let mut style_payload = None;
    for child in children_with_offsets(textbox, textbox_offset)? {
        let (offset, child) = child?;
        let child_type = record_type(child)?;
        child_types.push(child_type);
        if !matches!(
            child_type,
            value if value == RecordType::TextHeaderAtom as u16
                || value == RecordType::TextBytesAtom as u16
                || value == RecordType::TextCharsAtom as u16
                || value == RecordType::StyleTextPropAtom as u16
        ) {
            return Err(Error::Refused(Refusal::UnsupportedDependency));
        }
        if record_version(child)? != 0 || record_instance(child)? != 0 {
            return Err(PackageError::Corrupted(
                "source-backed textbox atom has an invalid record header".into(),
            )
            .into());
        }
        if child_type == RecordType::TextHeaderAtom as u16 {
            let payload = drawing_payload(child)?;
            if payload.len() != 4 {
                return Err(PackageError::Corrupted(
                    "source-backed TextHeaderAtom has an invalid payload length".into(),
                )
                .into());
            }
            let bytes: &[u8; 4] = payload.try_into().map_err(|_error| {
                PackageError::Corrupted(
                    "source-backed TextHeaderAtom has an invalid payload width".into(),
                )
            })?;
            crate::TextType::parse(u32::from_le_bytes(*bytes))?;
        }
        if child_type == RecordType::StyleTextPropAtom as u16 {
            style_payload = Some(drawing_payload(child)?);
        }
        if child_type == RecordType::TextBytesAtom as u16
            || child_type == RecordType::TextCharsAtom as u16
        {
            let kind = if child_type == RecordType::TextBytesAtom as u16 {
                TextKind::Bytes
            } else {
                TextKind::Chars
            };
            let payload = drawing_payload(child)?;
            let text = match kind {
                TextKind::Bytes => payload.iter().map(|byte| char::from(*byte)).collect(),
                TextKind::Chars => decode_utf16(payload)?,
            };
            if text.contains('\0') {
                return Err(Error::Refused(Refusal::IncompatibleEncoding));
            }
            atoms.push(SourceTextAtom {
                offset: offset
                    .checked_add(PPT_HEADER_LEN)
                    .ok_or_else(|| PackageError::Corrupted("text atom offset overflows".into()))?,
                kind,
                payload: Arc::from(payload),
                text,
            });
        }
    }
    let supported_order = matches!(
        child_types.as_slice(),
        [header, text]
            if *header == RecordType::TextHeaderAtom as u16
                && (*text == RecordType::TextBytesAtom as u16
                    || *text == RecordType::TextCharsAtom as u16)
    ) || matches!(
        child_types.as_slice(),
        [header, text, style]
            if *header == RecordType::TextHeaderAtom as u16
                && (*text == RecordType::TextBytesAtom as u16
                    || *text == RecordType::TextCharsAtom as u16)
                && *style == RecordType::StyleTextPropAtom as u16
    );
    if !supported_order {
        return Err(Error::Refused(Refusal::UnsupportedDependency));
    }
    if let Some(style_payload) = style_payload {
        let atom = atoms.first().ok_or(Error::Refused(Refusal::NoTextAtom))?;
        let text_length = text_units(atom.kind, atom.payload.len())?;
        crate::StyleTextPropAtom::parse(style_payload, text_length)?;
    }
    match atoms.len() {
        0 => Err(Error::Refused(Refusal::NoTextAtom)),
        1 => atoms
            .pop()
            .ok_or_else(|| PackageError::Corrupted("selected text atom disappeared".into()).into()),
        _ => Err(Error::Refused(Refusal::MultipleTextAtoms)),
    }
}

fn children_with_offsets<'a>(
    record: &'a [u8],
    record_offset: usize,
) -> PackageResult<ChildrenWithOffsets<'a>> {
    let payload = drawing_payload(record)?;
    let payload_offset = record_offset
        .checked_add(PPT_HEADER_LEN)
        .ok_or_else(|| PackageError::Corrupted("PPT child payload offset overflows".into()))?;
    Ok(ChildrenWithOffsets {
        payload,
        payload_offset,
        offset: 0,
    })
}

struct ChildrenWithOffsets<'a> {
    payload: &'a [u8],
    payload_offset: usize,
    offset: usize,
}

impl<'a> Iterator for ChildrenWithOffsets<'a> {
    type Item = PackageResult<(usize, &'a [u8])>;

    fn next(&mut self) -> Option<Self::Item> {
        if self.offset == self.payload.len() {
            return None;
        }
        let start = self.offset;
        let record = self.payload.get(start..);
        let result = record.and_then(|tail| {
            drawing_payload(tail)
                .ok()
                .and_then(|payload| PPT_HEADER_LEN.checked_add(payload.len()))
                .and_then(|length| tail.get(..length).map(|child| (length, child)))
        });
        let Some((length, child)) = result else {
            self.offset = self.payload.len();
            return Some(Err(PackageError::Corrupted(
                "truncated PPT source-backed child record".into(),
            )));
        };
        self.offset = start.saturating_add(length);
        let absolute = match self.payload_offset.checked_add(start) {
            Some(value) => value,
            None => {
                return Some(Err(PackageError::Corrupted(
                    "PPT source-backed child offset overflows".into(),
                )));
            },
        };
        Some(Ok((absolute, child)))
    }
}

fn parse_durable_target(value: &str) -> Result<Target> {
    let (slide_text, shape_text) = value
        .strip_prefix("slide:")
        .and_then(|suffix| suffix.split_once("/shape:"))
        .ok_or_else(|| invalid_durable_patch("invalid shape-text target"))?;
    if slide_text.is_empty()
        || shape_text.is_empty()
        || !slide_text.bytes().all(|byte| byte.is_ascii_digit())
        || !shape_text.bytes().all(|byte| byte.is_ascii_digit())
    {
        return Err(invalid_durable_patch("invalid shape-text target"));
    }
    let slide_position = slide_text
        .parse::<usize>()
        .map_err(|_error| invalid_durable_patch("slide position exceeds this platform"))?;
    let shape_position = shape_text
        .parse::<usize>()
        .map_err(|_error| invalid_durable_patch("shape position exceeds this platform"))?;
    Ok(Target::new(
        Position::new(slide_position),
        Position::new(shape_position),
    ))
}

fn invalid_durable_patch(message: &str) -> Error {
    PackageError::InvalidFormat(format!("invalid PPT durable patch: {message}")).into()
}

/// One isolated text replacement staged against an immutable package.
#[derive(Debug, Clone)]
pub(crate) struct Transaction {
    source: Snapshot,
    resolved: Resolved,
    replacement: Option<String>,
}

impl Transaction {
    /// The exact text currently stored in the selected shape.
    #[must_use]
    pub(crate) fn text(&self) -> &str {
        &self.resolved.text
    }

    /// The selected semantic target.
    #[must_use]
    pub(crate) const fn target(&self) -> Target {
        self.resolved.target
    }

    /// Stages a replacement after proving that the dependent text ranges stay
    /// valid. Failed validation leaves the staged candidate untouched.
    ///
    /// # Errors
    ///
    /// Returns a typed refusal when the replacement has an incompatible
    /// encoding or when its changed length has unmodeled dependent ranges.
    pub(crate) fn set_text(&mut self, value: impl Into<String>) -> Result<()> {
        let candidate = value.into();
        let encoded = encode_replacement(&candidate, self.resolved.kind)?;
        check_replacement(
            self.resolved.kind,
            &self.resolved.payload,
            &encoded,
            self.resolved.can_resize,
            &self.resolved.pins,
        )?;
        self.replacement = Some(candidate);
        Ok(())
    }

    /// Publishes the replacement atomically and returns a reversible,
    /// exact-source-checked patch. Equal text reuses the source allocation.
    ///
    /// # Errors
    ///
    /// Returns an error if package publication or source-checked readback
    /// fails, or a typed refusal for an unsafe replacement.
    pub(crate) fn commit(self) -> Result<RecordCommit> {
        self.commit_verified().map(|(commit, _)| commit)
    }

    /// [`Self::commit`], also returning the shapes a parse of the published
    /// slide returns when the readback proved them incrementally (record
    /// 0264).
    pub(crate) fn commit_verified(
        self,
    ) -> Result<(RecordCommit, Option<(u32, crate::slide::PatchedShapes)>)> {
        let replacement = self
            .replacement
            .unwrap_or_else(|| self.resolved.text.clone());
        if replacement == self.resolved.text {
            let patch = RecordPatch::new(self.source.clone(), self.source.clone(), None);
            return Ok((
                RecordCommit {
                    snapshot: self.source,
                    patch,
                    replaced_slide_persist_id: None,
                },
                None,
            ));
        }

        let encoded = encode_replacement(&replacement, self.resolved.kind)?;
        check_replacement(
            self.resolved.kind,
            &self.resolved.payload,
            &encoded,
            self.resolved.can_resize,
            &self.resolved.pins,
        )?;
        let target_slide = rewrite_slide(
            &self.resolved.slide_record,
            self.resolved.native_shape_id,
            self.resolved.kind,
            &self.resolved.payload,
            &encoded,
        )?;
        let mut editor = crate::embedded::object::Editor::open_records(
            self.source.bytes.clone().to_vec(),
            self.source.source_limits,
        )?;
        let live = editor.persisted_record(self.resolved.slide_persist_id)?;
        if live != self.resolved.slide_record.as_ref() {
            return Err(PackageError::Corrupted(
                "PPT shape-text transaction source changed before publication".into(),
            )
            .into());
        }
        editor.replace_persisted_record(self.resolved.slide_persist_id, target_slide)?;
        let bytes = editor.finish()?;
        // `Snapshot::from_bytes` validates that the candidate opens; that
        // open also serves the incremental readback (record 0264).
        let bytes: Arc<[u8]> = Arc::from(bytes.into_boxed_slice());
        let opened = presentation_shared(bytes.clone())?;
        let snapshot = Snapshot {
            bytes,
            source_limits: opened.source_limits(),
        };
        let (readback, verified) =
            resolve_published(&snapshot.bytes, &self.resolved, &replacement, &opened)?;
        drop(opened);
        if readback.text != replacement {
            return Err(PackageError::Corrupted(
                "published PPT shape text did not round-trip through the selected source shape"
                    .into(),
            )
            .into());
        }
        let change = Change {
            target: self.resolved.target,
            before_text: self.resolved.text,
            after_text: replacement,
        };
        let patch = RecordPatch::new(self.source, snapshot.clone(), Some(change));
        Ok((
            RecordCommit {
                snapshot,
                patch,
                replaced_slide_persist_id: Some(self.resolved.slide_persist_id),
            },
            verified,
        ))
    }
}

/// A published snapshot and its reversible source-checked patch.
#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) struct RecordCommit {
    snapshot: Snapshot,
    patch: RecordPatch,
    replaced_slide_persist_id: Option<u32>,
}

impl RecordCommit {
    /// The immutable committed package.
    #[must_use]
    pub(crate) const fn snapshot(&self) -> &Snapshot {
        &self.snapshot
    }

    /// The source-checked patch that produced this snapshot.
    #[must_use]
    pub(crate) const fn patch(&self) -> &RecordPatch {
        &self.patch
    }

    /// Consumes the commit into its snapshot and patch.
    #[must_use]
    pub(crate) fn into_parts(self) -> (Snapshot, RecordPatch) {
        (self.snapshot, self.patch)
    }

    pub(crate) fn into_root_publication(self) -> Option<RootPublication> {
        let replaced_slide_persist_id = self.replaced_slide_persist_id?;
        Some(RootPublication {
            source: self.patch.before.bytes,
            output: self.snapshot.bytes,
            replaced_slide_persist_id,
        })
    }
}

pub(crate) struct RootPublication {
    source: Arc<[u8]>,
    output: Arc<[u8]>,
    replaced_slide_persist_id: u32,
}

/// One source-resolved shape-text replacement returned to the root owner.
///
/// Native identities remain private to this adapter. The root transaction
/// records only the semantic target and exact before/after text in its patch.
#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) struct RootShapeTextChange {
    target: Target,
    before: String,
    after: String,
    slide_persist_id: u32,
}

impl RootShapeTextChange {
    pub(crate) const fn target(&self) -> Target {
        self.target
    }

    pub(crate) fn before(&self) -> &str {
        &self.before
    }

    pub(crate) fn after(&self) -> &str {
        &self.after
    }

    pub(crate) const fn slide_persist_id(&self) -> u32 {
        self.slide_persist_id
    }
}

/// One append-only publication containing replacements for several persisted
/// slide records.
pub(crate) struct RootBatchPublication {
    source: Arc<[u8]>,
    output: Arc<[u8]>,
    changes: Vec<RootShapeTextChange>,
    /// Whether the publication replaced the live `DocumentContainer`, which
    /// owns slide-list (outline) text bodies.
    document_changed: bool,
    /// Shapes that parsing each edited slide of `output` returns, proven by
    /// the incremental readback (record 0264); empty when not proven or not
    /// requested.
    verified: readback::VerifiedSlides,
}

impl RootBatchPublication {
    /// Take the proven shapes of the edited slides of the output.
    pub(crate) fn take_verified_shapes(&mut self) -> readback::VerifiedSlides {
        std::mem::take(&mut self.verified)
    }

    pub(crate) fn source(&self) -> &[u8] {
        &self.source
    }

    pub(crate) const fn document_changed(&self) -> bool {
        self.document_changed
    }

    pub(crate) fn changes(&self) -> &[RootShapeTextChange] {
        &self.changes
    }

    pub(crate) fn into_parts(self) -> (Arc<[u8]>, Vec<RootShapeTextChange>) {
        (self.output, self.changes)
    }
}

impl RootPublication {
    pub(crate) fn source(&self) -> &[u8] {
        &self.source
    }

    pub(crate) const fn replaced_slide_persist_id(&self) -> u32 {
        self.replaced_slide_persist_id
    }

    pub(crate) fn into_output(self) -> Arc<[u8]> {
        self.output
    }
}

/// A reversible in-memory patch authorized by exact package bytes.
#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) struct RecordPatch {
    before: Snapshot,
    after: Snapshot,
    change: Option<Change>,
}

impl RecordPatch {
    fn new(before: Snapshot, after: Snapshot, change: Option<Change>) -> Self {
        Self {
            before,
            after,
            change,
        }
    }

    /// Exact source bytes required for forward application.
    #[must_use]
    pub(crate) fn before(&self) -> &[u8] {
        self.before.bytes()
    }

    /// Exact target bytes produced by forward application.
    #[must_use]
    pub(crate) fn after(&self) -> &[u8] {
        self.after.bytes()
    }

    /// Whether the patch changes no bytes.
    #[must_use]
    pub(crate) fn is_empty(&self) -> bool {
        self.before.bytes() == self.after.bytes()
    }

    /// Applies only to the exact source artifact used to create this patch.
    ///
    /// # Errors
    ///
    /// Returns an error if `current` is not that exact source artifact.
    pub(crate) fn apply(&self, current: &Snapshot) -> Result<Snapshot> {
        if current.bytes() != self.before.bytes() {
            return Err(PackageError::InvalidFormat(
                "PPT shape-text patch source does not match its base artifact".into(),
            )
            .into());
        }
        Ok(self.after.clone())
    }

    /// Returns the exact-source-checked inverse patch.
    #[must_use]
    pub(crate) fn inverse(&self) -> Self {
        Self::new(
            self.after.clone(),
            self.before.clone(),
            self.change.as_ref().map(Change::inverse),
        )
    }

    /// Converts this patch to the shared versioned deterministic-JSON patch
    /// vocabulary.
    ///
    /// The durable operation uses semantic slide/shape positions and retains
    /// both the expected text and exact artifact SHA-256 as preconditions. It
    /// contains no native shape or persist identifier.
    ///
    /// # Errors
    ///
    /// Returns an error when caller-selected wire limits cannot represent the
    /// operation or its authored text.
    pub(crate) fn to_durable(
        &self,
        limits: litchi_core::patch::PatchLimits,
    ) -> std::result::Result<
        litchi_core::patch::Patch<litchi_core::patch::Reversible>,
        litchi_core::patch::PatchError,
    > {
        use litchi_core::patch::{BlobBundle, ReversibleOperation};

        let operations = self
            .change
            .as_ref()
            .map(|change| {
                let forward = durable_operation(
                    limits,
                    change.target,
                    self.before.bytes(),
                    &change.before_text,
                    &change.after_text,
                )?;
                let inverse = durable_operation(
                    limits,
                    change.target,
                    self.after.bytes(),
                    &change.after_text,
                    &change.before_text,
                )?;
                Ok(ReversibleOperation::new(forward, inverse))
            })
            .transpose()?
            .into_iter();
        litchi_core::patch::Patch::<litchi_core::patch::Reversible>::new(
            limits,
            "litchi-ppt",
            operations,
            BlobBundle::new(limits.blobs()),
            BlobBundle::new(limits.blobs()),
        )
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
struct Change {
    target: Target,
    before_text: String,
    after_text: String,
}

impl Change {
    fn inverse(&self) -> Self {
        Self {
            target: self.target,
            before_text: self.after_text.clone(),
            after_text: self.before_text.clone(),
        }
    }
}

fn durable_operation(
    limits: litchi_core::patch::PatchLimits,
    target: Target,
    source: &[u8],
    before_text: &str,
    after_text: &str,
) -> std::result::Result<litchi_core::patch::PatchOperation, litchi_core::patch::PatchError> {
    let mut preconditions = BTreeMap::new();
    preconditions.insert(
        "artifact_sha256".to_string(),
        serde_json::Value::String(litchi_core::patch::BlobId::of(source).as_hex()),
    );
    preconditions.insert(
        "text".to_string(),
        serde_json::Value::String(before_text.to_string()),
    );
    litchi_core::patch::PatchOperation::new(
        limits,
        "shape-text.set",
        format!(
            "slide:{}/shape:{}",
            target.slide().get(),
            target.shape().get()
        ),
        preconditions,
        serde_json::Value::String(after_text.to_string()),
    )
}

#[derive(Debug, Clone)]
pub(crate) struct AnchorRewrite {
    pub(crate) bytes: Vec<u8>,
    pub(crate) before: crate::Anchor,
    pub(crate) after: crate::Anchor,
}

struct ShapeLocation {
    persist_id: u32,
    shape_id: u32,
    slide_record: Vec<u8>,
}

#[derive(Debug, Clone)]
struct Resolved {
    target: Target,
    slide_persist_id: u32,
    native_shape_id: u32,
    slide_record: Arc<[u8]>,
    kind: TextKind,
    payload: Vec<u8>,
    text: String,
    can_resize: bool,
    /// The text positions of the body's metacharacter fields (record 0304).
    pins: Vec<usize>,
    /// The source parse of the target's slide, which the commit readback
    /// verifies the publication against incrementally (record 0264).
    source_view: Option<Arc<readback::SourceSlideView>>,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum TextKind {
    Bytes,
    Chars,
}

fn presentation(bytes: &[u8]) -> PackageResult<Presentation> {
    let mut package = Reader::from_reader(Cursor::new(bytes.to_vec()))?;
    package.presentation()
}

fn presentation_with_options(
    bytes: &[u8],
    options: crate::package::OpenOptions<'_>,
) -> PackageResult<Presentation> {
    let mut package = Reader::from_reader_with_options(Cursor::new(bytes), options)?;
    package.presentation_with_options(options)
}

fn presentation_shared(bytes: Arc<[u8]>) -> PackageResult<Presentation> {
    let mut package = Reader::from_reader(Cursor::new(bytes))?;
    package.presentation()
}

fn resolve(bytes: &[u8], target: Target) -> Result<Resolved> {
    let location = resolve_shape(bytes, target)?;
    resolved_from_location(target, location)
}

fn resolve_with_shape_target(
    target: Target,
    shape_target: (u32, u32),
    editor: &crate::embedded::object::Editor,
) -> Result<Resolved> {
    let location = resolve_shape_record(shape_target.0, shape_target.1, editor)?;
    resolved_from_location(target, location)
}

fn resolved_from_location(target: Target, location: ShapeLocation) -> Result<Resolved> {
    resolved_from_shared_location(
        target,
        location.persist_id,
        location.shape_id,
        Arc::from(location.slide_record.into_boxed_slice()),
    )
}

fn resolved_from_shared_location(
    target: Target,
    slide_persist_id: u32,
    native_shape_id: u32,
    slide_record: Arc<[u8]>,
) -> Result<Resolved> {
    let text = inspect_slide(&slide_record, native_shape_id)?;
    Ok(Resolved {
        target,
        slide_persist_id,
        native_shape_id,
        slide_record,
        kind: text.kind,
        payload: text.payload,
        text: text.text,
        can_resize: text.can_resize,
        pins: text.pins,
        source_view: None,
    })
}

fn resolve_shape(bytes: &[u8], target: Target) -> Result<ShapeLocation> {
    let presentation = presentation(bytes)?;
    resolve_shape_in_presentation(bytes, target, &presentation)
}

fn resolve_shape_in_presentation(
    bytes: &[u8],
    target: Target,
    presentation: &Presentation,
) -> Result<ShapeLocation> {
    let source_limits = presentation.source_limits();
    let (persist_id, shape_id) = resolve_shape_target_in_presentation(presentation, target)?;
    let editor = crate::embedded::object::Editor::open_records(bytes.to_vec(), source_limits)?;
    resolve_shape_record(persist_id, shape_id, &editor)
}

/// Reads one shape's text without opening the mutation-capable persisted
/// record editor.  Signed packages use this path to distinguish an exact
/// semantic no-op from a changed request while retaining the signature CFB
/// components byte-for-byte.
pub(crate) fn inspect_shape_text_readonly(
    bytes: &[u8],
    target: Target,
    options: crate::package::OpenOptions<'_>,
) -> Result<String> {
    let location = resolve_shape_readonly(bytes, target, options)?;
    match inspect_slide(&location.slide_record, location.shape_id) {
        Ok(atom) => Ok(atom.text),
        Err(Error::Refused(Refusal::NoTextAtom)) => {
            let Some(body) = outline_reference(&location.slide_record, location.shape_id)? else {
                return Err(Error::Refused(Refusal::NoTextAtom));
            };
            let presentation = presentation_with_options(bytes, options)?;
            let document = presentation.live_document_bytes()?;
            Ok(inspect_outline_body(&document, location.persist_id, body)?.text)
        },
        Err(error) => Err(error),
    }
}

/// Reads one shape's anchor without opening the mutation-capable persisted
/// record editor.  This is deliberately a read-only companion to
/// [`inspect_shape_anchor`], whose normal path also prepares a publication.
pub(crate) fn inspect_shape_anchor_readonly(
    bytes: &[u8],
    target: Target,
    options: crate::package::OpenOptions<'_>,
) -> Result<crate::Anchor> {
    let location = resolve_shape_readonly(bytes, target, options)?;
    inspect_anchor(&location.slide_record, location.shape_id)
}

fn resolve_shape_readonly(
    bytes: &[u8],
    target: Target,
    options: crate::package::OpenOptions<'_>,
) -> Result<ShapeLocation> {
    let presentation = presentation_with_options(bytes, options)?;
    let (persist_id, shape_id) = resolve_shape_target_in_presentation(&presentation, target)?;
    let factory = crate::slide::SlideFactory::new_shared_with_limits(
        &presentation.inner.powerpoint_document,
        &presentation.inner.persist_mapping,
        &presentation.inner.slide_directory,
        presentation.source_limits(),
    );
    let slide = factory.parse_slide(persist_id)?;
    let document = slide.doc_data();
    let tail = document.get(slide.offset..).ok_or_else(|| {
        PackageError::Corrupted("PPT selected slide offset exceeds its document stream".into())
    })?;
    let (_, consumed) = crate::Record::parse(tail, 0, &presentation.source_limits())?;
    let slide_end = slide
        .offset
        .checked_add(consumed)
        .ok_or_else(|| PackageError::Corrupted("PPT selected slide length overflows".into()))?;
    let slide_record = document
        .get(slide.offset..slide_end)
        .ok_or_else(|| {
            PackageError::Corrupted("PPT selected slide exceeds its document stream".into())
        })?
        .to_vec();
    Ok(ShapeLocation {
        persist_id,
        shape_id,
        slide_record,
    })
}

fn resolve_shape_target(bytes: &[u8], target: Target) -> Result<(u32, u32)> {
    let presentation = presentation(bytes)?;
    resolve_shape_target_in_presentation(&presentation, target)
}

/// Resolve `target` in a published candidate, as [`resolve`] does, proving
/// its slide's parse incrementally from the source parse when possible
/// (record 0264). [`resolve`] parses the whole slide to find the target's
/// persist and shape identifiers; when [`readback::readback_incremental`]
/// accepts, those identifiers are proven to be the ones that parse yields,
/// and every later step is the same.
fn resolve_published(
    bytes: &[u8],
    resolved: &Resolved,
    replacement: &str,
    opened: &Presentation,
) -> Result<(Resolved, Option<(u32, crate::slide::PatchedShapes)>)> {
    if let Some(view) = &resolved.source_view
        && let Some((shape_target, patched)) =
            incremental_published_target(opened, resolved, view, replacement)
    {
        let editor =
            crate::embedded::object::Editor::open_records(bytes.to_vec(), opened.source_limits())?;
        let location = resolve_shape_record(shape_target.0, shape_target.1, &editor)?;
        return Ok((
            resolved_from_location(resolved.target, location)?,
            Some(patched),
        ));
    }
    Ok((resolve(bytes, resolved.target)?, None))
}

/// `opened` is a default-policy open of exactly the candidate, the same open
/// `resolve_shape_target` makes.
fn incremental_published_target(
    opened: &Presentation,
    resolved: &Resolved,
    view: &Arc<readback::SourceSlideView>,
    replacement: &str,
) -> Option<((u32, u32), (u32, crate::slide::PatchedShapes))> {
    let slides = opened.slides().ok()?;
    let targets = [readback::ReadbackTarget {
        target: resolved.target,
        slide_persist_id: resolved.slide_persist_id,
        native_shape_id: resolved.native_shape_id,
        replacement,
    }];
    let sources = BTreeMap::from([(resolved.slide_persist_id, Arc::clone(view))]);
    let (persist_id, patched) =
        readback::readback_incremental(&slides, &targets, &sources)?.pop()?;
    Some((
        (persist_id, resolved.native_shape_id),
        (persist_id, patched),
    ))
}

/// Resolves one target through a single selected slide of an opened
/// projection, with the same refusals as
/// [`resolve_shape_target_in_presentation`], also capturing the slide's
/// parse (record 0264).
fn resolve_shape_target_at_with_view(
    presentation: &Presentation,
    target: Target,
) -> Result<((u32, u32), Option<readback::SourceSlideView>)> {
    let slide = presentation
        .slide_at(target.slide.get())?
        .ok_or(Error::Refused(Refusal::SlideNotFound {
            position: target.slide,
        }))?;
    let shape = slide
        .shapes()?
        .get(target.shape.get())
        .ok_or(Error::Refused(Refusal::ShapeNotFound))?;
    let shape_target = (slide.persist_id(), native_shape_id(shape));
    Ok((
        shape_target,
        readback::SourceSlideView::capture(&slide).ok(),
    ))
}

/// [`resolve_shape_target_in_presentation`], also capturing the selected
/// slide's parse.
fn resolve_shape_target_in_presentation_with_view(
    presentation: &Presentation,
    target: Target,
) -> Result<((u32, u32), Option<readback::SourceSlideView>)> {
    let slides = presentation.slides()?;
    let slide = slides
        .get(target.slide.get())
        .ok_or(Error::Refused(Refusal::SlideNotFound {
            position: target.slide,
        }))?;
    let shape = slide
        .shapes()?
        .get(target.shape.get())
        .ok_or(Error::Refused(Refusal::ShapeNotFound))?;
    let shape_target = (slide.persist_id(), native_shape_id(shape));
    Ok((shape_target, readback::SourceSlideView::capture(slide).ok()))
}

fn resolve_shape_target_in_presentation(
    presentation: &Presentation,
    target: Target,
) -> Result<(u32, u32)> {
    let slides = presentation.slides()?;
    let slide = slides
        .get(target.slide.get())
        .ok_or(Error::Refused(Refusal::SlideNotFound {
            position: target.slide,
        }))?;
    let shape = slide
        .shapes()?
        .get(target.shape.get())
        .ok_or(Error::Refused(Refusal::ShapeNotFound))?;
    Ok((slide.persist_id(), native_shape_id(shape)))
}

#[derive(Debug)]
struct BatchTextTarget {
    target: Target,
    replacement: String,
    slide_persist_id: u32,
    native_shape_id: u32,
    slide_record: Arc<[u8]>,
}

#[derive(Debug)]
struct PreparedBatchText {
    resolved: Resolved,
    replacement: String,
    encoded: Vec<u8>,
    /// The slide-list text body the shape's `OutlineTextRefAtom` selects, when
    /// the shape's text lives in the live document instead of its drawing.
    outline_body: Option<usize>,
}

/// Replaces a bounded caller-preflighted set of semantic shape targets through
/// one persisted-record editor and one full candidate reopen.
///
/// Callers retain the hard operation-count bound. This owner resolves every
/// target and validates every encoding/dependency closure before staging any
/// record in the editor. Targets sharing a persisted slide are indexed and
/// rewritten in one drawing pass, while each rewrite retains the original
/// atom payload as its exact source precondition.
pub(crate) fn replace_shape_texts_batch(
    source: Arc<[u8]>,
    replacements: &[(Target, &str)],
    max_output_bytes: usize,
    opened: Option<&Presentation>,
) -> Result<Option<RootBatchPublication>> {
    let replacement_bytes =
        replacements
            .iter()
            .try_fold(0usize, |total, (_target, replacement)| {
                total.checked_add(replacement.len()).ok_or_else(|| {
                    PackageError::ResourceLimit(
                        "PPT shape-text batch replacement byte count overflow".into(),
                    )
                })
            })?;
    require_batch_text_budget(replacement_bytes, 0, max_output_bytes)?;

    // An already opened projection of exactly `source` (the root edit's
    // untouched package) shares its parsed shapes; otherwise open one.
    let presentation = match opened {
        Some(presentation) => presentation.clone(),
        None => presentation_shared(source.clone())?,
    };
    let mut editor = crate::embedded::object::Editor::open_records_arc_with_limit(
        source.clone(),
        max_output_bytes,
        presentation.source_limits(),
    )?;
    let document_persist_id = presentation.slide_directory().document_persist_id();
    let slides = presentation.slides()?;
    let mut targets = Vec::with_capacity(replacements.len());
    let mut original_slides = BTreeMap::<u32, Arc<[u8]>>::new();
    let mut requested_shapes = BTreeMap::<u32, BTreeSet<u32>>::new();
    let mut source_views = BTreeMap::<u32, Arc<readback::SourceSlideView>>::new();
    let mut deferred_resolution_error = None;
    // Resolve the semantic slide and shape positions once per request, but
    // defer text inspection until each persisted slide's drawing has been
    // indexed.  This keeps the public positional/error order while avoiding a
    // complete drawing walk for every target on a large slide.
    for (target, replacement) in replacements {
        let Some(slide) = slides.get(target.slide.get()) else {
            deferred_resolution_error = Some(Error::Refused(Refusal::SlideNotFound {
                position: target.slide,
            }));
            break;
        };
        let shapes = match slide.shapes() {
            Ok(shapes) => shapes,
            Err(error) => {
                deferred_resolution_error = Some(error.into());
                break;
            },
        };
        let Some(shape) = shapes.get(target.shape.get()) else {
            deferred_resolution_error = Some(Error::Refused(Refusal::ShapeNotFound));
            break;
        };
        let slide_persist_id = slide.persist_id();
        let native_shape_id = native_shape_id(shape);
        if let Entry::Vacant(entry) = source_views.entry(slide_persist_id)
            && let Ok(view) = readback::SourceSlideView::capture(slide)
        {
            // The incremental readback starts from this parse (record 0264).
            entry.insert(Arc::new(view));
        }
        let slide_record = match original_slides.get(&slide_persist_id) {
            Some(record) => record.clone(),
            None => {
                let record: Arc<[u8]> = match editor.persisted_record(slide_persist_id) {
                    Ok(record) => Arc::from(record.into_boxed_slice()),
                    Err(error) => {
                        deferred_resolution_error = Some(error.into());
                        break;
                    },
                };
                original_slides.insert(slide_persist_id, record.clone());
                record
            },
        };
        targets.push(BatchTextTarget {
            target: *target,
            replacement: (*replacement).to_string(),
            slide_persist_id,
            native_shape_id,
            slide_record,
        });
        requested_shapes
            .entry(slide_persist_id)
            .or_default()
            .insert(native_shape_id);
    }
    drop(slides);
    drop(presentation);

    let mut indexed_shapes = BTreeMap::new();
    let mut index_failures = BTreeSet::new();
    let mut prepared = Vec::with_capacity(targets.len());
    let mut encoded_bytes = 0usize;
    let mut inspected_shapes = BTreeMap::<(u32, u32), (TextAtom, Option<usize>)>::new();
    let mut live_document: Option<Vec<u8>> = None;

    // Complete text/dependency resolution and replacement validation before
    // the editor receives its first staged mutation.  The index was built by
    // one framing walk per drawing; this loop retains the caller's request
    // order for all typed shape/text errors.
    for target in targets {
        if !index_failures.contains(&target.slide_persist_id)
            && let Entry::Vacant(entry) = indexed_shapes.entry(target.slide_persist_id)
        {
            let shape_ids = requested_shapes
                .get(&target.slide_persist_id)
                .ok_or_else(|| {
                    PackageError::Corrupted(
                        "PPT batch shape-text indexing lost its requested identities".into(),
                    )
                })?;
            let slide_record = original_slides
                .get(&target.slide_persist_id)
                .ok_or_else(|| {
                    PackageError::Corrupted(
                        "PPT batch shape-text indexing lost its original slide record".into(),
                    )
                })?;
            match index_slide_text_targets(slide_record, shape_ids) {
                Ok(index) => {
                    entry.insert(index);
                },
                Err(_error) => {
                    // A malformed OfficeArt subtree takes the slow error
                    // path below, which replays the old one-target
                    // inspector in caller order and therefore preserves
                    // its exact first-error precedence.
                    index_failures.insert(target.slide_persist_id);
                },
            }
        }
        let (text, outline_body) = match inspected_shapes
            .get(&(target.slide_persist_id, target.native_shape_id))
        {
            Some(inspected) => inspected.clone(),
            None => {
                let inspected = if index_failures.contains(&target.slide_persist_id) {
                    inspect_slide(&target.slide_record, target.native_shape_id)
                } else {
                    let index = indexed_shapes
                        .get(&target.slide_persist_id)
                        .ok_or_else(|| {
                            PackageError::Corrupted(
                                "PPT batch shape-text validation lost its drawing index".into(),
                            )
                        })?;
                    inspect_indexed_shape_text(index, target.native_shape_id)
                };
                let inspected = match inspected {
                    Ok(text) => (text, None),
                    Err(Error::Refused(Refusal::NoTextAtom)) => {
                        // A placeholder whose text lives in the slide
                        // list holds only an `OutlineTextRefAtom`
                        // (MS-PPT 2.9.78); its body is edited in the live
                        // document instead of the drawing.
                        let Some(body) =
                            outline_reference(&target.slide_record, target.native_shape_id)?
                        else {
                            return Err(Error::Refused(Refusal::NoTextAtom));
                        };
                        if live_document.is_none() {
                            live_document = Some(editor.persisted_record(document_persist_id)?);
                        }
                        let document = live_document.as_deref().ok_or_else(|| {
                            PackageError::Corrupted(
                                "PPT outline text lost its live document".into(),
                            )
                        })?;
                        (
                            inspect_outline_body(document, target.slide_persist_id, body)?,
                            Some(body),
                        )
                    },
                    Err(error) => return Err(error),
                };
                inspected_shapes.insert(
                    (target.slide_persist_id, target.native_shape_id),
                    inspected.clone(),
                );
                inspected
            },
        };
        let resolved = Resolved {
            target: target.target,
            slide_persist_id: target.slide_persist_id,
            native_shape_id: target.native_shape_id,
            slide_record: target.slide_record,
            kind: text.kind,
            payload: text.payload,
            text: text.text,
            can_resize: text.can_resize,
            pins: text.pins,
            source_view: None,
        };
        let encoded_len = encoded_replacement_len(&target.replacement, resolved.kind)?;
        encoded_bytes = encoded_bytes.checked_add(encoded_len).ok_or_else(|| {
            PackageError::ResourceLimit("PPT shape-text batch encoded byte count overflow".into())
        })?;
        require_batch_text_budget(replacement_bytes, encoded_bytes, max_output_bytes)?;
        let encoded = encode_replacement(&target.replacement, resolved.kind)?;
        if encoded.len() != encoded_len {
            return Err(PackageError::Corrupted(
                "PPT shape-text batch replacement size preflight disagrees with encoding".into(),
            )
            .into());
        }
        check_replacement(
            resolved.kind,
            &resolved.payload,
            &encoded,
            resolved.can_resize,
            &resolved.pins,
        )?;
        if target.replacement != resolved.text {
            prepared.push(PreparedBatchText {
                resolved,
                replacement: target.replacement,
                encoded,
                outline_body,
            });
        }
    }

    if let Some(error) = deferred_resolution_error {
        return Err(error);
    }

    if prepared.is_empty() {
        return Ok(None);
    }

    // Input order is retained in the semantic change list. Each native slide
    // drawing is rewritten exactly once after all of its shape replacements
    // succeed; the rewrite itself visits the OfficeArt tree once and applies
    // every selected textbox while preserving wire order.
    let mut grouped = BTreeMap::<u32, Vec<ShapeTextRewrite<'_>>>::new();
    let mut outline = BTreeMap::<(u32, usize), ShapeTextRewrite<'_>>::new();
    for prepared in &prepared {
        if let Some(body) = prepared.outline_body {
            let rewrite = ShapeTextRewrite {
                shape_id: prepared.resolved.native_shape_id,
                kind: prepared.resolved.kind,
                before: &prepared.resolved.payload,
                after: &prepared.encoded,
            };
            if outline
                .insert((prepared.resolved.slide_persist_id, body), rewrite)
                .is_some()
            {
                // Two shapes referencing one body would publish two texts
                // into one record.
                return Err(Error::Refused(Refusal::AmbiguousTextbox));
            }
            continue;
        }
        grouped
            .entry(prepared.resolved.slide_persist_id)
            .or_default()
            .push(ShapeTextRewrite {
                shape_id: prepared.resolved.native_shape_id,
                kind: prepared.resolved.kind,
                before: &prepared.resolved.payload,
                after: &prepared.encoded,
            });
    }
    let mut staged_slides = BTreeMap::<u32, Vec<u8>>::new();
    for (slide_persist_id, replacements) in grouped {
        let original = original_slides.get(&slide_persist_id).ok_or_else(|| {
            PackageError::Corrupted(
                "PPT batch shape-text rewrite lost its original slide record".into(),
            )
        })?;
        staged_slides.insert(
            slide_persist_id,
            rewrite_slide_grouped(original, &replacements)?,
        );
    }

    for (persist_id, staged) in staged_slides {
        let original = original_slides.get(&persist_id).ok_or_else(|| {
            PackageError::Corrupted(
                "PPT batch shape-text publication lost its original slide record".into(),
            )
        })?;
        if editor.persisted_record(persist_id)?.as_slice() != original.as_ref() {
            return Err(PackageError::Corrupted(
                "PPT batch shape-text transaction source changed before publication".into(),
            )
            .into());
        }
        editor.replace_persisted_record(persist_id, staged)?;
    }
    let document_changed = !outline.is_empty();
    if document_changed {
        let original = live_document.as_deref().ok_or_else(|| {
            PackageError::Corrupted("PPT outline text publication lost its live document".into())
        })?;
        let staged = rewrite_outline_bodies(original, &outline)?;
        if editor.persisted_record(document_persist_id)?.as_slice() != original {
            return Err(PackageError::Corrupted(
                "PPT outline text transaction source changed before publication".into(),
            )
            .into());
        }
        editor.replace_persisted_record(document_persist_id, staged)?;
    }

    let candidate = editor.finish()?;
    crate::font::validate_unrelated_streams(&source, &candidate)?;
    let output: Arc<[u8]> = Arc::from(candidate.into_boxed_slice());

    // One complete reopen validates the publication and supplies all semantic
    // readbacks without reopening the package once per shape. The edited
    // slides are verified incrementally against their source parse when that
    // is provably equivalent to parsing them again (record 0264).
    let reopened = presentation_shared(output.clone())?;
    let reopened_slides = reopened.slides()?;
    let targets = prepared
        .iter()
        .map(|prepared| readback::ReadbackTarget {
            target: prepared.resolved.target,
            slide_persist_id: prepared.resolved.slide_persist_id,
            native_shape_id: prepared.resolved.native_shape_id,
            replacement: &prepared.replacement,
        })
        .collect::<Vec<_>>();
    let verified = if document_changed {
        // Outline bodies live in the republished document; parse in full.
        readback::readback_full(&reopened_slides, &targets)?;
        Vec::new()
    } else {
        readback::readback(&reopened_slides, &targets, &source_views)?
    };
    drop(targets);
    let changes = prepared
        .into_iter()
        .map(|prepared| RootShapeTextChange {
            target: prepared.resolved.target,
            before: prepared.resolved.text,
            after: prepared.replacement,
            slide_persist_id: prepared.resolved.slide_persist_id,
        })
        .collect();

    Ok(Some(RootBatchPublication {
        source,
        output,
        changes,
        document_changed,
        verified,
    }))
}

/// Publishes a batch through a test hook that may corrupt each staged slide
/// record, then returns both verifiers' verdicts on the published bytes
/// (record 0264 differential tests).
#[cfg(test)]
pub(crate) fn differential_readback_for_test(
    source: &[u8],
    replacements: &[(Target, &str)],
    corrupt: &dyn Fn(u32, &mut Vec<u8>),
) -> Result<DifferentialVerdicts> {
    let source: Arc<[u8]> = Arc::from(source);
    let presentation = presentation_shared(source.clone())?;
    let source_limits = presentation.source_limits();
    let slides = presentation.slides()?;
    let mut editor = crate::embedded::object::Editor::open_records_arc_with_limit(
        source.clone(),
        usize::MAX,
        source_limits,
    )?;
    let mut views = BTreeMap::new();
    let mut grouped = BTreeMap::<u32, Vec<ShapeTextRewrite<'_>>>::new();
    let mut originals = BTreeMap::new();
    let mut prepared = Vec::new();
    for (target, replacement) in replacements {
        let slide =
            slides
                .get(target.slide.get())
                .ok_or(Error::Refused(Refusal::SlideNotFound {
                    position: target.slide,
                }))?;
        let shape = slide
            .shapes()?
            .get(target.shape.get())
            .ok_or(Error::Refused(Refusal::ShapeNotFound))?;
        let persist_id = slide.persist_id();
        let shape_id = native_shape_id(shape);
        if let Entry::Vacant(entry) = views.entry(persist_id) {
            entry.insert(Arc::new(readback::SourceSlideView::capture(slide)?));
        }
        if let Entry::Vacant(entry) = originals.entry(persist_id) {
            entry.insert(editor.persisted_record(persist_id)?);
        }
        let record = originals
            .get(&persist_id)
            .ok_or(Error::Refused(Refusal::ShapeNotFound))?;
        let text = inspect_slide(record, shape_id)?;
        prepared.push((
            *target,
            persist_id,
            shape_id,
            text,
            encode_replacement_owned(replacement)?,
        ));
    }
    for (_, persist_id, shape_id, text, encoded) in &prepared {
        let encoded = match text.kind {
            TextKind::Chars => &encoded.0,
            TextKind::Bytes => encoded
                .1
                .as_ref()
                .ok_or(Error::Refused(Refusal::DependencyClosure))?,
        };
        grouped
            .entry(*persist_id)
            .or_default()
            .push(ShapeTextRewrite {
                shape_id: *shape_id,
                kind: text.kind,
                before: &text.payload,
                after: encoded,
            });
    }
    for (persist_id, rewrites) in grouped {
        let original = originals
            .get(&persist_id)
            .ok_or(Error::Refused(Refusal::ShapeNotFound))?;
        let mut staged = rewrite_slide_grouped(original, &rewrites)?;
        corrupt(persist_id, &mut staged);
        editor.replace_persisted_record(persist_id, staged)?;
    }
    let output: Arc<[u8]> = Arc::from(editor.finish()?.into_boxed_slice());
    // A candidate that does not reopen is refused by both verifiers, which
    // share the reopen.
    let reopened = presentation_shared(output.clone());
    let reopened_slides = reopened.as_ref().map(Presentation::slides);
    let Ok(Ok(reopened_slides)) = reopened_slides else {
        return Ok(DifferentialVerdicts {
            incremental: None,
            full_accepted: false,
            reopened: false,
            single: None,
            persist_ids: Vec::new(),
            parsed: Vec::new(),
        });
    };
    let targets = prepared
        .iter()
        .zip(replacements)
        .map(
            |((target, persist_id, shape_id, _, _), (_, replacement))| readback::ReadbackTarget {
                target: *target,
                slide_persist_id: *persist_id,
                native_shape_id: *shape_id,
                replacement,
            },
        )
        .collect::<Vec<_>>();
    let incremental = readback::readback_incremental(&reopened_slides, &targets, &views);
    // The single-target commit readback on the same candidate: the proven
    // resolution against the reference `resolve`.
    let single = match replacements {
        [(target, replacement)] => {
            let transaction = Snapshot::from_shared_bytes(source.clone())?.edit_text(*target)?;
            let opened = presentation_shared(output.clone())?;
            let proven = resolve_published(&output, &transaction.resolved, replacement, &opened);
            let reference = resolve(&output, *target);
            let summary = |resolved: &Resolved| {
                (
                    resolved.slide_persist_id,
                    resolved.native_shape_id,
                    resolved.payload.clone(),
                    resolved.text.clone(),
                )
            };
            Some(SingleVerdicts {
                proven: proven
                    .as_ref()
                    .map(|(resolved, _)| summary(resolved))
                    .map_err(|error| format!("{error:?}")),
                used_proof: proven
                    .as_ref()
                    .is_ok_and(|(_, verified)| verified.is_some()),
                reference: reference
                    .as_ref()
                    .map(summary)
                    .map_err(|error| format!("{error:?}")),
            })
        },
        _ => None,
    };
    let full = readback::readback_full(&reopened_slides, &targets);
    let parsed = reopened_slides
        .iter()
        .map(|slide| {
            slide
                .shapes()
                .map(|shapes| format!("{shapes:?}"))
                .map_err(|_| ())
        })
        .collect();
    Ok(DifferentialVerdicts {
        incremental: incremental.map(|verified| {
            verified
                .into_iter()
                .map(|(persist_id, shapes)| (persist_id, format!("{:?}", &*shapes.materialize())))
                .collect()
        }),
        full_accepted: full.is_ok(),
        reopened: true,
        single,
        persist_ids: reopened_slides
            .iter()
            .map(|slide| slide.persist_id())
            .collect(),
        parsed,
    })
}

/// Both verifiers' verdicts on one published candidate.
#[cfg(test)]
#[derive(Debug)]
pub(crate) struct DifferentialVerdicts {
    /// The incremental verifier's proven shapes (debug form) per edited
    /// slide, or `None` when it could not prove acceptance.
    pub(crate) incremental: Option<Vec<(u32, String)>>,
    /// Whether the reference verifier accepted.
    pub(crate) full_accepted: bool,
    /// Whether the candidate reopened and listed its slides (the step both
    /// verifiers share).
    pub(crate) reopened: bool,
    /// For a one-target batch, the single-commit readback on the candidate.
    pub(crate) single: Option<SingleVerdicts>,
    /// Persist identifiers of the reopened slides, in order.
    pub(crate) persist_ids: Vec<u32>,
    /// A full parse of every reopened slide's shapes (debug form).
    pub(crate) parsed: Vec<std::result::Result<String, ()>>,
}

/// A resolved readback: persist id, shape id, text payload, text.
#[cfg(test)]
pub(crate) type ResolvedSummary = (u32, u32, Vec<u8>, String);

/// The single-target commit readback, proven and reference.
#[cfg(test)]
#[derive(Debug)]
pub(crate) struct SingleVerdicts {
    /// `resolve_published` (incremental when provable).
    pub(crate) proven: std::result::Result<ResolvedSummary, String>,
    /// Whether `resolve_published` used the incremental proof.
    pub(crate) used_proof: bool,
    /// The reference `resolve` (full parse).
    pub(crate) reference: std::result::Result<ResolvedSummary, String>,
}

#[cfg(test)]
fn encode_replacement_owned(value: &str) -> Result<(Vec<u8>, Option<Vec<u8>>)> {
    Ok((
        encode_replacement(value, TextKind::Chars)?,
        encode_replacement(value, TextKind::Bytes).ok(),
    ))
}

/// One published speaker-notes text replacement.
pub(crate) struct RootNotesPublication {
    source: Arc<[u8]>,
    output: Arc<[u8]>,
    notes_persist_id: u32,
    before: String,
}

impl RootNotesPublication {
    pub(crate) fn source(&self) -> &[u8] {
        &self.source
    }

    pub(crate) const fn notes_persist_id(&self) -> u32 {
        self.notes_persist_id
    }

    pub(crate) fn before(&self) -> &str {
        &self.before
    }

    pub(crate) fn into_output(self) -> Arc<[u8]> {
        self.output
    }
}

/// Reads the speaker-notes text of the slide at a source position.
pub(crate) fn inspect_notes_text_readonly(
    bytes: &[u8],
    slide: Position,
    options: crate::package::OpenOptions<'_>,
) -> Result<String> {
    let presentation = presentation_with_options(bytes, options)?;
    let selected = presentation
        .slide_at(slide.get())?
        .ok_or(Error::Refused(Refusal::SlideNotFound { position: slide }))?;
    let notes = selected
        .speaker_notes()?
        .ok_or(Error::Refused(Refusal::NoNotesPage))?;
    Ok(notes.text()?.to_owned())
}

/// Replaces the speaker-notes text of the slide at a source position.
///
/// The notes page must hold exactly one body placeholder
/// (`PlaceholderKind::NotesBody`), whose text body is rewritten under the
/// same dependency-closure rules as slide shape text; every other record of
/// the notes page and every other persist object is copied exactly. Equal
/// text returns `Ok(None)` without publishing.
pub(crate) fn replace_notes_text(
    source: Arc<[u8]>,
    slide: Position,
    replacement: &str,
    max_output_bytes: usize,
) -> Result<Option<RootNotesPublication>> {
    require_batch_text_budget(replacement.len(), 0, max_output_bytes)?;
    let presentation = presentation_shared(source.clone())?;
    let selected = presentation
        .slide_at(slide.get())?
        .ok_or(Error::Refused(Refusal::SlideNotFound { position: slide }))?;
    let notes = selected
        .speaker_notes()?
        .ok_or(Error::Refused(Refusal::NoNotesPage))?;
    let before = notes.text()?.to_owned();
    if before == replacement {
        return Ok(None);
    }
    let notes_persist_id = notes.persist_id();
    let source_limits = presentation.source_limits();
    let shape_id = match notes.body_shape_ids()?.as_slice() {
        [] => return Err(Error::Refused(Refusal::NoTextbox)),
        [shape_id] => *shape_id,
        _ => return Err(Error::Refused(Refusal::AmbiguousTextbox)),
    };
    drop(selected);
    drop(presentation);

    let mut editor = crate::embedded::object::Editor::open_records_arc_with_limit(
        source.clone(),
        max_output_bytes,
        source_limits,
    )?;
    let original = editor.persisted_record(notes_persist_id)?;
    if record_type(&original)? != RecordType::Notes as u16 {
        return Err(PackageError::Corrupted(
            "PPT notes persist object is not a NotesContainer".into(),
        )
        .into());
    }
    let atom = inspect_slide(&original, shape_id)?;
    let encoded_len = encoded_replacement_len(replacement, atom.kind)?;
    require_batch_text_budget(replacement.len(), encoded_len, max_output_bytes)?;
    let encoded = encode_replacement(replacement, atom.kind)?;
    check_replacement(
        atom.kind,
        &atom.payload,
        &encoded,
        atom.can_resize,
        &atom.pins,
    )?;
    let staged = rewrite_slide_grouped(
        &original,
        &[ShapeTextRewrite {
            shape_id,
            kind: atom.kind,
            before: &atom.payload,
            after: &encoded,
        }],
    )?;
    editor.replace_persisted_record(notes_persist_id, staged)?;
    let candidate = editor.finish()?;
    crate::font::validate_unrelated_streams(&source, &candidate)?;
    let output: Arc<[u8]> = Arc::from(candidate.into_boxed_slice());

    let reopened = presentation_shared(output.clone())?;
    let readback = reopened
        .slide_at(slide.get())?
        .ok_or(Error::Refused(Refusal::SlideNotFound { position: slide }))?;
    let notes = readback
        .speaker_notes()?
        .ok_or(Error::Refused(Refusal::NoNotesPage))?;
    if notes.persist_id() != notes_persist_id || notes.text()? != replacement {
        return Err(PackageError::Corrupted(
            "published PPT notes text did not round-trip through its notes page".into(),
        )
        .into());
    }
    Ok(Some(RootNotesPublication {
        source,
        output,
        notes_persist_id,
        before,
    }))
}

/// Rewrite the text of the shape with `OfficeArt` ID `shape_id` under the
/// ordinary text-closure rules; `None` when the text is already equal.
pub(crate) fn rewrite_shape_text_by_id(
    slide: &[u8],
    shape_id: u32,
    replacement: &str,
) -> Result<Option<Vec<u8>>> {
    let atom = match inspect_slide(slide, shape_id) {
        Ok(atom) => atom,
        Err(Error::Refused(Refusal::NoTextAtom)) => {
            return insert_text_atom(slide, shape_id, replacement);
        },
        Err(error) => return Err(error),
    };
    if atom.text == replacement {
        return Ok(None);
    }
    let encoded = encode_replacement(replacement, atom.kind)?;
    check_replacement(
        atom.kind,
        &atom.payload,
        &encoded,
        atom.can_resize,
        &atom.pins,
    )?;
    rewrite_slide(slide, shape_id, atom.kind, &atom.payload, &encoded).map(Some)
}

/// Give an empty text body its first text atom (record 0323).
///
/// The shape's `ClientTextbox` must hold a `TextHeaderAtom` followed only by
/// at most one each of `StyleTextPropAtom`, `TextSpecialInfoAtom` and
/// `TextRulerAtom`: the records [`resize_dependency_closure_is_modeled`]
/// admits. A `TextCharsAtom` is inserted after the header and the style and
/// special-information runs grow from the paragraph mark alone to cover the
/// new text. An empty replacement keeps the body.
fn insert_text_atom(slide: &[u8], shape_id: u32, replacement: &str) -> Result<Option<Vec<u8>>> {
    if replacement.is_empty() {
        return Ok(None);
    }
    let encoded = encode_replacement(replacement, TextKind::Chars)?;
    let units = text_units(TextKind::Chars, encoded.len())?;
    let mut drawing_count = 0usize;
    let mut matches = 0usize;
    let rewritten = rewrite_ppt_record(slide, 0, &mut |record| {
        if record_type(record)? != RecordType::PPDrawing as u16 {
            return Ok(None);
        }
        drawing_count += 1;
        let data = insert_into_officeart(
            drawing_payload(record)?,
            shape_id,
            &encoded,
            units,
            0,
            &mut matches,
        )?;
        rebuild(record, &data).map(Some).map_err(Into::into)
    })?;
    if drawing_count != 1 {
        return Err(PackageError::Corrupted(
            "selected slide has ambiguous PPDrawing ownership".into(),
        )
        .into());
    }
    match matches {
        0 => Err(Error::Refused(Refusal::ShapeNotFound)),
        1 => rewritten
            .map(Some)
            .ok_or(Error::Refused(Refusal::ShapeNotFound)),
        _ => Err(Error::Refused(Refusal::AmbiguousShape)),
    }
}

fn insert_into_officeart(
    record: &[u8],
    shape_id: u32,
    encoded: &[u8],
    units: usize,
    depth: usize,
    matches: &mut usize,
) -> Result<Vec<u8>> {
    if depth >= 128 {
        return Err(
            PackageError::Corrupted("OfficeArt nesting exceeds text-edit limit".into()).into(),
        );
    }
    if record_type(record)? == OFFICEART_SP_CONTAINER && shape_id_of(record)? == Some(shape_id) {
        *matches = matches.saturating_add(1);
        let mut textboxes = 0usize;
        let mut data = Vec::with_capacity(record.len().saturating_add(encoded.len() + 16));
        for child in children(record)? {
            let child = child?;
            if record_type(child)? == OFFICEART_CLIENT_TEXTBOX {
                textboxes += 1;
                data.extend_from_slice(&insert_into_textbox(child, encoded, units)?);
            } else {
                data.extend_from_slice(child);
            }
        }
        return match textboxes {
            0 => Err(Error::Refused(Refusal::NoTextbox)),
            1 => rebuild(record, &data).map_err(Into::into),
            _ => Err(Error::Refused(Refusal::AmbiguousTextbox)),
        };
    }
    if record_version(record)? != 0xF {
        return Ok(record.to_vec());
    }
    let mut data = Vec::with_capacity(drawing_payload(record)?.len());
    for child in children(record)? {
        data.extend_from_slice(&insert_into_officeart(
            child?,
            shape_id,
            encoded,
            units,
            depth + 1,
            matches,
        )?);
    }
    rebuild(record, &data).map_err(Into::into)
}

fn insert_into_textbox(textbox: &[u8], encoded: &[u8], units: usize) -> Result<Vec<u8>> {
    let records = children(textbox)?.collect::<PackageResult<Vec<_>>>()?;
    let [header, rest @ ..] = records.as_slice() else {
        return Err(Error::Refused(Refusal::NoTextAtom));
    };
    if record_type(header)? != RecordType::TextHeaderAtom as u16 {
        return Err(Error::Refused(Refusal::NoTextAtom));
    }
    let resizable = [
        RecordType::StyleTextPropAtom as u16,
        RecordType::TextSpecInfoAtom as u16,
        RecordType::TextRulerAtom as u16,
    ];
    let mut seen = [false; 3];
    let mut data = Vec::with_capacity(textbox.len().saturating_add(encoded.len() + 8));
    data.extend_from_slice(header);
    let length = u32::try_from(encoded.len())
        .map_err(|_err| PackageError::ResourceLimit("PPT text atom exceeds u32".into()))?;
    data.extend_from_slice(&0_u16.to_le_bytes());
    data.extend_from_slice(&(RecordType::TextCharsAtom as u16).to_le_bytes());
    data.extend_from_slice(&length.to_le_bytes());
    data.extend_from_slice(encoded);
    for record in rest {
        let kind = record_type(record)?;
        let Some(slot) = resizable.iter().position(|candidate| *candidate == kind) else {
            // Another record may hold a character position.
            return Err(Error::Refused(Refusal::NoTextAtom));
        };
        if std::mem::replace(&mut seen[slot], true) {
            return Err(Error::Refused(Refusal::NoTextAtom));
        }
        match slot {
            0 => data.extend_from_slice(
                &crate::StyleTextPropAtom::resize_single_run_record(record, 0, units)?
                    .ok_or(Error::Refused(Refusal::DependencyClosure))?,
            ),
            1 => data.extend_from_slice(
                &crate::TextSpecialInfoRuns::resize_single_run_record(record, 0, units)?
                    .ok_or(Error::Refused(Refusal::DependencyClosure))?,
            ),
            _ => data.extend_from_slice(record),
        }
    }
    rebuild(textbox, &data).map_err(Into::into)
}

fn require_batch_text_budget(
    replacement_bytes: usize,
    encoded_bytes: usize,
    max_bytes: usize,
) -> Result<()> {
    let retained_bytes = replacement_bytes
        .checked_add(encoded_bytes)
        .ok_or_else(|| {
            PackageError::ResourceLimit("PPT shape-text batch retained byte count overflow".into())
        })?;
    if retained_bytes > max_bytes {
        return Err(PackageError::ResourceLimit(format!(
            "PPT shape-text batch retained bytes {retained_bytes} exceeds limit {max_bytes}"
        ))
        .into());
    }
    Ok(())
}

fn resolve_shape_record(
    persist_id: u32,
    shape_id: u32,
    editor: &crate::embedded::object::Editor,
) -> Result<ShapeLocation> {
    let slide_record = editor.persisted_record(persist_id)?;
    Ok(ShapeLocation {
        persist_id,
        shape_id,
        slide_record,
    })
}

pub(crate) fn replace_shape_anchor(
    bytes: &[u8],
    target: Target,
    replacement: crate::Anchor,
) -> Result<AnchorRewrite> {
    let source_presentation = presentation(bytes)?;
    let location = resolve_shape_in_presentation(bytes, target, &source_presentation)?;
    let before = inspect_anchor(&location.slide_record, location.shape_id)?;
    if before == replacement {
        return Ok(AnchorRewrite {
            bytes: bytes.to_vec(),
            before,
            after: replacement,
        });
    }
    let rewritten = rewrite_anchor(
        &location.slide_record,
        location.shape_id,
        &before.to_bytes(),
        &replacement.to_bytes(),
    )?;
    let mut editor = crate::embedded::object::Editor::open_records(
        bytes.to_vec(),
        source_presentation.source_limits(),
    )?;
    if editor.persisted_record(location.persist_id)? != location.slide_record {
        return Err(PackageError::Corrupted(
            "PPT shape-anchor source changed before publication".into(),
        )
        .into());
    }
    editor.replace_persisted_record(location.persist_id, rewritten)?;
    let candidate = editor.finish()?;
    crate::font::validate_unrelated_streams(bytes, &candidate)?;
    let _ = presentation(&candidate)?;
    if inspect_shape_anchor(&candidate, target)? != replacement {
        return Err(PackageError::Corrupted(
            "published PPT shape anchor did not round-trip".into(),
        )
        .into());
    }
    Ok(AnchorRewrite {
        bytes: candidate,
        before,
        after: replacement,
    })
}

pub(crate) fn inspect_shape_anchor(bytes: &[u8], target: Target) -> Result<crate::Anchor> {
    let location = resolve_shape(bytes, target)?;
    inspect_anchor(&location.slide_record, location.shape_id)
}

#[cfg(test)]
pub(crate) fn shape_text_can_resize(bytes: &[u8], target: Target) -> Result<bool> {
    resolve(bytes, target).map(|resolved| resolved.can_resize)
}

#[cfg(test)]
pub(crate) fn add_unmodeled_text_dependency_for_test(
    bytes: &[u8],
    target: Target,
) -> Result<Vec<u8>> {
    fn inject_officeart(
        record: &[u8],
        shape_id: u32,
        depth: usize,
        matches: &mut usize,
    ) -> Result<Vec<u8>> {
        if depth >= 128 {
            return Err(PackageError::Corrupted(
                "OfficeArt nesting exceeds text-edit test limit".into(),
            )
            .into());
        }
        if record_type(record)? == OFFICEART_SP_CONTAINER && shape_id_of(record)? == Some(shape_id)
        {
            *matches += 1;
            let mut textboxes = 0usize;
            let mut data = Vec::with_capacity(drawing_payload(record)?.len() + PPT_HEADER_LEN);
            for child_result in children(record)? {
                let child = child_result?;
                if record_type(child)? == OFFICEART_CLIENT_TEXTBOX {
                    textboxes += 1;
                    let mut textbox_data = drawing_payload(child)?.to_vec();
                    textbox_data.extend_from_slice(&0_u16.to_le_bytes());
                    // Unknown inert child: it is preserved by the reader but
                    // deliberately makes the text-position dependency closure
                    // ineligible for length-changing edits.
                    textbox_data.extend_from_slice(&0x7ffe_u16.to_le_bytes());
                    textbox_data.extend_from_slice(&0_u32.to_le_bytes());
                    data.extend_from_slice(&rebuild(child, &textbox_data)?);
                } else {
                    data.extend_from_slice(child);
                }
            }
            if textboxes != 1 {
                return Err(Error::Refused(Refusal::AmbiguousTextbox));
            }
            return rebuild(record, &data).map_err(Into::into);
        }
        if record_version(record)? != 0xF {
            return Ok(record.to_vec());
        }
        let mut data = Vec::with_capacity(drawing_payload(record)?.len() + PPT_HEADER_LEN);
        for child_result in children(record)? {
            data.extend_from_slice(&inject_officeart(
                child_result?,
                shape_id,
                depth + 1,
                matches,
            )?);
        }
        rebuild(record, &data).map_err(Into::into)
    }

    let presentation = presentation(bytes)?;
    let location = resolve_shape_in_presentation(bytes, target, &presentation)?;
    let mut drawings = 0usize;
    let rewritten = rewrite_ppt_record(&location.slide_record, 0, &mut |record| {
        if record_type(record)? != RecordType::PPDrawing as u16 {
            return Ok(None);
        }
        drawings += 1;
        let mut matches = 0usize;
        let data = inject_officeart(drawing_payload(record)?, location.shape_id, 0, &mut matches)?;
        if matches != 1 {
            return Err(Error::Refused(Refusal::AmbiguousShape));
        }
        rebuild(record, &data).map(Some).map_err(Into::into)
    })?
    .ok_or(Error::Refused(Refusal::ShapeNotFound))?;
    if drawings != 1 {
        return Err(PackageError::Corrupted(
            "selected slide has ambiguous PPDrawing ownership".into(),
        )
        .into());
    }
    let mut editor = crate::embedded::object::Editor::open_records(
        bytes.to_vec(),
        presentation.source_limits(),
    )?;
    editor.replace_persisted_record(location.persist_id, rewritten)?;
    editor.finish().map_err(Into::into)
}

pub(crate) fn native_shape_id(selected: &ShapeEnum<'_>) -> u32 {
    match selected {
        ShapeEnum::TextBox(textbox) => textbox.source_shape_id(),
        ShapeEnum::Placeholder(placeholder) => placeholder.source_shape_id(),
        ShapeEnum::AutoShape(auto_shape) => auto_shape.source_shape_id(),
        ShapeEnum::Picture(picture) => picture.properties().id,
        ShapeEnum::Table(table) => table.id(),
        ShapeEnum::Group(group) => group.id(),
        ShapeEnum::Line(line) => line.id(),
    }
}

#[derive(Debug, Clone)]
struct TextAtom {
    kind: TextKind,
    payload: Vec<u8>,
    text: String,
    can_resize: bool,
    /// The text positions of the body's metacharacter fields (record 0304).
    pins: Vec<usize>,
}

#[derive(Debug)]
struct IndexedShapeText<'a> {
    containers: Vec<&'a [u8]>,
}

/// Index only the requested OfficeArt shape containers in one drawing walk.
///
/// The index deliberately records raw containers instead of eagerly parsing
/// their text.  This lets the caller inspect targets in the original request
/// order, retaining the single-target error precedence while ensuring that
/// an unrelated textbox is never validated merely because it shares a
/// drawing with a requested shape.
fn index_slide_text_targets<'a>(
    slide: &'a [u8],
    requested_shape_ids: &BTreeSet<u32>,
) -> Result<BTreeMap<u32, IndexedShapeText<'a>>> {
    let (_, consumed) = crate::Record::parse_strict(slide, 0)?;
    if consumed != slide.len() {
        return Err(PackageError::Corrupted("selected slide has trailing bytes".into()).into());
    }
    let drawing = find_ppt_record(slide, RecordType::PPDrawing as u16, 0)?
        .ok_or(Error::Refused(Refusal::ShapeNotFound))?;
    let mut index = BTreeMap::new();
    let mut active_shape_ids = requested_shape_ids.clone();
    index_officeart_shape_targets(
        drawing_payload(drawing)?,
        &mut active_shape_ids,
        0,
        &mut index,
    )?;
    Ok(index)
}

fn index_officeart_shape_targets<'a>(
    record: &'a [u8],
    active_shape_ids: &mut BTreeSet<u32>,
    depth: usize,
    index: &mut BTreeMap<u32, IndexedShapeText<'a>>,
) -> Result<()> {
    if depth >= 128 {
        return Err(PackageError::Corrupted(
            "OfficeArt nesting exceeds shape-inspection limit".into(),
        )
        .into());
    }
    if record_type(record)? == OFFICEART_SP_CONTAINER {
        if let Some(shape_id) = shape_id_of(record)?
            && active_shape_ids.contains(&shape_id)
        {
            index
                .entry(shape_id)
                .or_insert_with(|| IndexedShapeText {
                    containers: Vec::new(),
                })
                .containers
                .push(record);
            // A single-target visit stops at its matching SP container.  If
            // this drawing has other requested identities, those target
            // visits continue through the selected container and therefore
            // still require its nested OfficeArt framing to be checked.  The
            // active set models those independent visits without rescanning
            // the drawing once per target.  Removing the matched identity
            // also prevents nested containers with that identity from being
            // misclassified as additional semantic matches.
            if active_shape_ids.len() == 1 {
                return Ok(());
            }
            active_shape_ids.remove(&shape_id);
            let result = (|| {
                if active_shape_ids.is_empty() {
                    return Ok(());
                }
                if record_version(record)? == 0xF {
                    for child_result in children(record)? {
                        index_officeart_shape_targets(
                            child_result?,
                            active_shape_ids,
                            depth + 1,
                            index,
                        )?;
                    }
                }
                Ok(())
            })();
            // Every independent target scan stops at its own matching
            // container, so restore this identity before returning even when
            // nested framing reported an error.
            active_shape_ids.insert(shape_id);
            return result;
        }
    }
    if record_version(record)? == 0xF {
        for child_result in children(record)? {
            index_officeart_shape_targets(child_result?, active_shape_ids, depth + 1, index)?;
        }
    }
    Ok(())
}

fn inspect_indexed_shape_text(
    index: &BTreeMap<u32, IndexedShapeText<'_>>,
    shape_id: u32,
) -> Result<TextAtom> {
    let indexed = index
        .get(&shape_id)
        .ok_or(Error::Refused(Refusal::ShapeNotFound))?;
    let mut matches = 0usize;
    let mut text = None;
    for container in &indexed.containers {
        matches = matches.saturating_add(1);
        let atom = inspect_shape_container_text(container)?;
        if text.replace(atom).is_some() {
            return Err(Error::Refused(Refusal::AmbiguousShape));
        }
    }
    match matches {
        0 => Err(Error::Refused(Refusal::ShapeNotFound)),
        1 => text.ok_or(Error::Refused(Refusal::NoTextbox)),
        _ => Err(Error::Refused(Refusal::AmbiguousShape)),
    }
}

fn inspect_shape_container_text(container: &[u8]) -> Result<TextAtom> {
    let mut textbox = None;
    for child_result in children(container)? {
        let child = child_result?;
        if record_type(child)? == OFFICEART_CLIENT_TEXTBOX && textbox.replace(child).is_some() {
            return Err(Error::Refused(Refusal::AmbiguousTextbox));
        }
    }
    let selected_textbox = textbox.ok_or(Error::Refused(Refusal::NoTextbox))?;
    inspect_textbox(selected_textbox)
}

fn inspect_slide(slide: &[u8], shape_id: u32) -> Result<TextAtom> {
    let (_, consumed) = crate::Record::parse_strict(slide, 0)?;
    if consumed != slide.len() {
        return Err(PackageError::Corrupted("selected slide has trailing bytes".into()).into());
    }
    let drawing = find_ppt_record(slide, RecordType::PPDrawing as u16, 0)?
        .ok_or(Error::Refused(Refusal::ShapeNotFound))?;
    inspect_drawing(drawing, shape_id)
}

fn inspect_anchor(slide: &[u8], shape_id: u32) -> Result<crate::Anchor> {
    let (_, consumed) = crate::Record::parse_strict(slide, 0)?;
    if consumed != slide.len() {
        return Err(PackageError::Corrupted("selected slide has trailing bytes".into()).into());
    }
    let drawing = find_ppt_record(slide, RecordType::PPDrawing as u16, 0)?
        .ok_or(Error::Refused(Refusal::ShapeNotFound))?;
    let mut matches = 0usize;
    let mut anchor = None;
    visit_shape_container(
        drawing_payload(drawing)?,
        shape_id,
        0,
        &mut matches,
        &mut |container| {
            for child_result in children(container)? {
                let child = child_result?;
                if record_type(child)? == OFFICEART_CLIENT_ANCHOR
                    && anchor
                        .replace(crate::Anchor::parse(
                            child,
                            &crate::client_anchor::Limits::default(),
                        )?)
                        .is_some()
                {
                    return Err(Error::Refused(Refusal::AmbiguousAnchor));
                }
            }
            Ok(())
        },
    )?;
    match matches {
        0 => Err(Error::Refused(Refusal::ShapeNotFound)),
        1 => anchor.ok_or(Error::Refused(Refusal::NoAnchor)),
        _ => Err(Error::Refused(Refusal::AmbiguousShape)),
    }
}

fn find_ppt_record(record: &[u8], target: u16, depth: usize) -> PackageResult<Option<&[u8]>> {
    if depth >= 128 {
        return Err(PackageError::Corrupted(
            "PPT record nesting exceeds text-edit limit".into(),
        ));
    }
    if record_type(record)? == target {
        return Ok(Some(record));
    }
    if record_version(record)? != 0xF {
        return Ok(None);
    }
    let mut found = None;
    for child_result in children(record)? {
        let child = child_result?;
        if let Some(candidate) = find_ppt_record(child, target, depth + 1)? {
            if found.is_some() {
                return Err(PackageError::Corrupted(
                    "selected slide has multiple PPDrawing records".into(),
                ));
            }
            found = Some(candidate);
        }
    }
    Ok(found)
}

fn inspect_drawing(drawing: &[u8], shape_id: u32) -> Result<TextAtom> {
    let mut matches = 0usize;
    let mut text = None;
    visit_officeart(
        drawing_payload(drawing)?,
        shape_id,
        0,
        &mut matches,
        &mut |textbox| {
            let atom = inspect_textbox(textbox)?;
            text = Some(atom);
            Ok(())
        },
    )?;
    match matches {
        0 => Err(Error::Refused(Refusal::ShapeNotFound)),
        1 => text.ok_or(Error::Refused(Refusal::NoTextbox)),
        _ => Err(Error::Refused(Refusal::AmbiguousShape)),
    }
}

fn inspect_textbox(textbox: &[u8]) -> Result<TextAtom> {
    inspect_text_records(children(textbox)?)
}

/// Inspects the records of one text body: a `ClientTextbox` payload or a
/// slide-list text body, which share the same record vocabulary.
fn inspect_text_records<'a>(
    records: impl Iterator<Item = PackageResult<&'a [u8]>>,
) -> Result<TextAtom> {
    let mut atom = None;
    let mut child_types = Vec::new();
    let mut style_record = None;
    let mut special_info_record = None;
    let mut pins = Vec::new();
    for child_result in records {
        let child = child_result?;
        let child_type = record_type(child)?;
        child_types.push(child_type);
        if is_metacharacter(child_type) {
            let position = drawing_payload(child)?
                .first_chunk::<4>()
                .map(|bytes| u32::from_le_bytes(*bytes))
                .ok_or_else(|| PackageError::Corrupted("metacharacter atom is truncated".into()))?;
            pins.push(usize::try_from(position).unwrap_or(usize::MAX));
        }
        if child_type == RecordType::StyleTextPropAtom as u16
            && style_record.replace(child).is_some()
        {
            return Err(Error::Refused(Refusal::DependencyClosure));
        }
        if child_type == RecordType::TextSpecInfoAtom as u16 {
            special_info_record = Some(child);
        }
        let atom_kind = match child_type {
            value if value == RecordType::TextBytesAtom as u16 => Some(TextKind::Bytes),
            value if value == RecordType::TextCharsAtom as u16 => Some(TextKind::Chars),
            _ => None,
        };
        let Some(kind) = atom_kind else { continue };
        if atom.is_some() {
            return Err(Error::Refused(Refusal::MultipleTextAtoms));
        }
        let payload = drawing_payload(child)?.to_vec();
        let text = match kind {
            TextKind::Bytes => payload.iter().map(|byte| char::from(*byte)).collect(),
            TextKind::Chars => decode_utf16(&payload)?,
        };
        atom = Some(TextAtom {
            kind,
            payload,
            text,
            can_resize: false,
            pins: Vec::new(),
        });
    }
    let mut text_atom = atom.ok_or(Error::Refused(Refusal::NoTextAtom))?;
    text_atom.pins = pins;
    text_atom.can_resize = resize_dependency_closure_is_modeled(
        &child_types,
        style_record,
        special_info_record,
        text_units(text_atom.kind, text_atom.payload.len())?,
    )?;
    Ok(text_atom)
}

/// Refuse a replacement whose length change has unmodeled dependent ranges,
/// or (record 0304) that writes another character at a metacharacter
/// field's position: the field (`*` for a date, time, slide number,
/// header or footer) would become plain text without a word.
fn check_replacement(
    kind: TextKind,
    payload: &[u8],
    encoded: &[u8],
    can_resize: bool,
    pins: &[usize],
) -> Result<()> {
    if encoded.len() != payload.len() && !can_resize {
        return Err(Error::Refused(Refusal::DependencyClosure));
    }
    let width = match kind {
        TextKind::Bytes => 1,
        TextKind::Chars => 2,
    };
    for &pin in pins {
        let unit = pin
            .checked_mul(width)
            .and_then(|start| Some(start..start.checked_add(width)?));
        let same = unit.is_some_and(|unit| {
            payload.get(unit.clone()).is_some() && payload.get(unit.clone()) == encoded.get(unit)
        });
        if !same {
            return Err(Error::Refused(Refusal::DependencyClosure));
        }
    }
    Ok(())
}

fn is_metacharacter(record: u16) -> bool {
    [
        RecordType::SlideNumberMCAtom,
        RecordType::DateTimeMCAtom,
        RecordType::GenericDateMCAtom,
        RecordType::HeaderMCAtom,
        RecordType::FooterMCAtom,
        RecordType::RtfDateTimeMCAtom,
    ]
    .iter()
    .any(|kind| *kind as u16 == record)
}

/// Whether a length-changing replacement keeps every dependent range of the
/// text body valid.
///
/// The body must be a `TextHeaderAtom` followed by its one text atom, then at
/// most one each of:
///
/// - `StyleTextPropAtom` with one paragraph run and one character run, whose
///   two coverage counts are resized;
/// - `TextSpecialInfoAtom` with one run (language and spelling), whose
///   coverage count is resized;
/// - `TextRulerAtom`, which holds indents and tab stops but no character
///   position, so it stays exact.
///
/// Any other record (interactions, bookmarks, metacharacter fields, several
/// formatting runs) associates data with character positions that a whole
/// replacement would invalidate; the body then keeps the same-length-only
/// contract and a changed length is a typed dependency-closure refusal.
fn resize_dependency_closure_is_modeled(
    child_types: &[u16],
    style_record: Option<&[u8]>,
    special_info_record: Option<&[u8]>,
    old_text_units: usize,
) -> Result<bool> {
    let text_header = RecordType::TextHeaderAtom as u16;
    let text_bytes = RecordType::TextBytesAtom as u16;
    let text_chars = RecordType::TextCharsAtom as u16;
    let [header, text, rest @ ..] = child_types else {
        return Ok(false);
    };
    if *header != text_header || (*text != text_bytes && *text != text_chars) {
        return Ok(false);
    }
    let resizable = [
        RecordType::StyleTextPropAtom as u16,
        RecordType::TextSpecInfoAtom as u16,
        RecordType::TextRulerAtom as u16,
    ];
    let mut seen = [false; 3];
    for child_type in rest {
        let Some(slot) = resizable.iter().position(|kind| kind == child_type) else {
            return Ok(false);
        };
        if std::mem::replace(&mut seen[slot], true) {
            return Ok(false);
        }
    }
    if let Some(style_bytes) = style_record
        && crate::StyleTextPropAtom::resize_single_run_record(
            style_bytes,
            old_text_units,
            old_text_units,
        )?
        .is_none()
    {
        return Ok(false);
    }
    if let Some(special_info_bytes) = special_info_record
        && crate::TextSpecialInfoRuns::resize_single_run_record(
            special_info_bytes,
            old_text_units,
            old_text_units,
        )?
        .is_none()
    {
        return Ok(false);
    }
    Ok(true)
}

fn rewrite_slide(
    slide: &[u8],
    shape_id: u32,
    kind: TextKind,
    before: &[u8],
    after: &[u8],
) -> Result<Vec<u8>> {
    let mut drawing_count = 0usize;
    let rewritten = rewrite_ppt_record(slide, 0, &mut |record| {
        if record_type(record)? != RecordType::PPDrawing as u16 {
            return Ok(None);
        }
        drawing_count += 1;
        let mut shapes = 0usize;
        let drawing = rewrite_drawing(record, shape_id, kind, before, after, &mut shapes)?;
        if shapes == 0 {
            return Ok(None);
        }
        if shapes > 1 {
            return Err(Error::Refused(Refusal::AmbiguousShape));
        }
        Ok(Some(drawing))
    })?;
    if drawing_count != 1 {
        return Err(PackageError::Corrupted(
            "selected slide has ambiguous PPDrawing ownership".into(),
        )
        .into());
    }
    rewritten.ok_or(Error::Refused(Refusal::ShapeNotFound))
}

#[derive(Debug, Clone, Copy)]
struct ShapeTextRewrite<'a> {
    shape_id: u32,
    kind: TextKind,
    before: &'a [u8],
    after: &'a [u8],
}

fn rewrite_slide_grouped(slide: &[u8], replacements: &[ShapeTextRewrite<'_>]) -> Result<Vec<u8>> {
    let mut drawing_count = 0usize;
    let mut match_counts = BTreeMap::<u32, usize>::new();
    for replacement in replacements {
        match_counts.entry(replacement.shape_id).or_insert(0);
    }
    let rewritten = rewrite_ppt_record(slide, 0, &mut |record| {
        if record_type(record)? != RecordType::PPDrawing as u16 {
            return Ok(None);
        }
        drawing_count += 1;
        let drawing = rewrite_drawing_grouped(record, replacements, &mut match_counts)?;
        Ok(Some(drawing))
    })?;
    if drawing_count != 1 {
        return Err(PackageError::Corrupted(
            "selected slide has ambiguous PPDrawing ownership".into(),
        )
        .into());
    }
    for (shape_id, matches) in match_counts {
        match matches {
            0 => return Err(Error::Refused(Refusal::ShapeNotFound)),
            1 => {},
            _ => return Err(Error::Refused(Refusal::AmbiguousShape)),
        }
        // Keep the binding above explicit: the count is checked for every
        // requested identity even when the drawing visitor found another
        // target first in wire order.
        let _ = shape_id;
    }
    rewritten.ok_or(Error::Refused(Refusal::ShapeNotFound))
}

fn rewrite_drawing_grouped(
    drawing: &[u8],
    replacements: &[ShapeTextRewrite<'_>],
    match_counts: &mut BTreeMap<u32, usize>,
) -> Result<Vec<u8>> {
    let mut replacements_by_shape = BTreeMap::<u32, Vec<ShapeTextRewrite<'_>>>::new();
    for replacement in replacements {
        replacements_by_shape
            .entry(replacement.shape_id)
            .or_default()
            .push(*replacement);
    }
    let data = rewrite_officeart_grouped(
        drawing_payload(drawing)?,
        &replacements_by_shape,
        match_counts,
        0,
    )?;
    rebuild(drawing, &data).map_err(Into::into)
}

fn rewrite_officeart_grouped(
    record: &[u8],
    replacements: &BTreeMap<u32, Vec<ShapeTextRewrite<'_>>>,
    match_counts: &mut BTreeMap<u32, usize>,
    depth: usize,
) -> Result<Vec<u8>> {
    let mut output = Vec::with_capacity(record.len());
    rewrite_officeart_grouped_into(record, replacements, match_counts, depth, &mut output)?;
    Ok(output)
}

/// Appends the rewritten record to `output`. Every container is framed again
/// in place (its length patched after its children), so one drawing pass
/// writes one buffer instead of allocating a copy per record (record 0264);
/// the visited records, checks, errors and bytes are those of the
/// per-record rebuild.
fn rewrite_officeart_grouped_into(
    record: &[u8],
    replacements: &BTreeMap<u32, Vec<ShapeTextRewrite<'_>>>,
    match_counts: &mut BTreeMap<u32, usize>,
    depth: usize,
    output: &mut Vec<u8>,
) -> Result<()> {
    if depth >= 128 {
        return Err(
            PackageError::Corrupted("OfficeArt nesting exceeds text-edit limit".into()).into(),
        );
    }
    if record_type(record)? == OFFICEART_SP_CONTAINER {
        if let Some(shape_id) = shape_id_of(record)?
            && let Some(shape_replacements) = replacements.get(&shape_id)
        {
            let count = match_counts.get_mut(&shape_id).ok_or_else(|| {
                PackageError::Corrupted("PPT grouped rewrite lost shape identity".into())
            })?;
            *count = count.saturating_add(1);
            output.extend_from_slice(&rewrite_shape_container_grouped(
                record,
                shape_replacements,
            )?);
            return Ok(());
        }
    }
    if record_version(record)? != 0xF {
        output.extend_from_slice(record);
        return Ok(());
    }
    // `drawing_payload` proves the 8-byte header and the payload extent.
    let payload = drawing_payload(record)?;
    let start = output.len();
    output.extend_from_slice(&record[..PPT_HEADER_LEN]);
    output.reserve(payload.len());
    for child_result in children(record)? {
        rewrite_officeart_grouped_into(
            child_result?,
            replacements,
            match_counts,
            depth + 1,
            output,
        )?;
    }
    let data_length = u32::try_from(output.len() - start - PPT_HEADER_LEN)
        .map_err(|_err| PackageError::ResourceLimit("PPT text-edit record exceeds u32".into()))?;
    output[start + 4..start + PPT_HEADER_LEN].copy_from_slice(&data_length.to_le_bytes());
    Ok(())
}

fn rewrite_shape_container_grouped(
    record: &[u8],
    replacements: &[ShapeTextRewrite<'_>],
) -> Result<Vec<u8>> {
    let mut textbox_count = 0usize;
    let mut data = Vec::with_capacity(drawing_payload(record)?.len());
    for child_result in children(record)? {
        let child = child_result?;
        if record_type(child)? != OFFICEART_CLIENT_TEXTBOX {
            data.extend_from_slice(child);
            continue;
        }
        textbox_count += 1;
        let mut rewritten = child.to_vec();
        for replacement in replacements {
            rewritten = rewrite_textbox(
                &rewritten,
                replacement.kind,
                replacement.before,
                replacement.after,
            )?;
        }
        data.extend_from_slice(&rewritten);
    }
    if textbox_count == 0 {
        return Err(Error::Refused(Refusal::NoTextbox));
    }
    if textbox_count > 1 {
        return Err(Error::Refused(Refusal::AmbiguousTextbox));
    }
    rebuild(record, &data).map_err(Into::into)
}

fn rewrite_anchor(slide: &[u8], shape_id: u32, before: &[u8], after: &[u8]) -> Result<Vec<u8>> {
    let mut drawing_count = 0usize;
    let rewritten = rewrite_ppt_record(slide, 0, &mut |record| {
        if record_type(record)? != RecordType::PPDrawing as u16 {
            return Ok(None);
        }
        drawing_count += 1;
        let mut shapes = 0usize;
        let drawing = rewrite_anchor_drawing(record, shape_id, before, after, &mut shapes)?;
        match shapes {
            0 => Ok(None),
            1 => Ok(Some(drawing)),
            _ => Err(Error::Refused(Refusal::AmbiguousShape)),
        }
    })?;
    if drawing_count != 1 {
        return Err(PackageError::Corrupted(
            "selected slide has ambiguous PPDrawing ownership".into(),
        )
        .into());
    }
    rewritten.ok_or(Error::Refused(Refusal::ShapeNotFound))
}

fn rewrite_anchor_drawing(
    drawing: &[u8],
    shape_id: u32,
    before: &[u8],
    after: &[u8],
    matches: &mut usize,
) -> Result<Vec<u8>> {
    let data = rewrite_anchor_officeart(
        drawing_payload(drawing)?,
        shape_id,
        before,
        after,
        0,
        matches,
    )?;
    rebuild(drawing, &data).map_err(Into::into)
}

fn rewrite_anchor_officeart(
    record: &[u8],
    shape_id: u32,
    before: &[u8],
    after: &[u8],
    depth: usize,
    matches: &mut usize,
) -> Result<Vec<u8>> {
    if depth >= 128 {
        return Err(
            PackageError::Corrupted("OfficeArt nesting exceeds shape-anchor limit".into()).into(),
        );
    }
    if record_type(record)? == OFFICEART_SP_CONTAINER && shape_id_of(record)? == Some(shape_id) {
        *matches = matches.saturating_add(1);
        let mut anchors = 0usize;
        let mut data = Vec::with_capacity(drawing_payload(record)?.len());
        for child_result in children(record)? {
            let child = child_result?;
            if record_type(child)? == OFFICEART_CLIENT_ANCHOR {
                anchors += 1;
                if child != before {
                    return Err(PackageError::Corrupted(
                        "selected shape anchor changed before publication".into(),
                    )
                    .into());
                }
                data.extend_from_slice(after);
            } else {
                data.extend_from_slice(child);
            }
        }
        return match anchors {
            0 => Err(Error::Refused(Refusal::NoAnchor)),
            1 => rebuild(record, &data).map_err(Into::into),
            _ => Err(Error::Refused(Refusal::AmbiguousAnchor)),
        };
    }
    if record_version(record)? != 0xF {
        return Ok(record.to_vec());
    }
    let mut data = Vec::with_capacity(drawing_payload(record)?.len());
    for child_result in children(record)? {
        data.extend_from_slice(&rewrite_anchor_officeart(
            child_result?,
            shape_id,
            before,
            after,
            depth + 1,
            matches,
        )?);
    }
    rebuild(record, &data).map_err(Into::into)
}

fn rewrite_ppt_record(
    record: &[u8],
    depth: usize,
    visit: &mut impl FnMut(&[u8]) -> Result<Option<Vec<u8>>>,
) -> Result<Option<Vec<u8>>> {
    if depth >= 128 {
        return Err(
            PackageError::Corrupted("PPT record nesting exceeds text-edit limit".into()).into(),
        );
    }
    if let Some(replacement) = visit(record)? {
        return Ok(Some(replacement));
    }
    if record_version(record)? != 0xF {
        return Ok(None);
    }
    let mut changed = false;
    let mut data = Vec::with_capacity(drawing_payload(record)?.len());
    for child_result in children(record)? {
        let child = child_result?;
        if let Some(replacement) = rewrite_ppt_record(child, depth + 1, visit)? {
            changed = true;
            data.extend_from_slice(&replacement);
        } else {
            data.extend_from_slice(child);
        }
    }
    if !changed {
        return Ok(None);
    }
    rebuild(record, &data).map(Some).map_err(Into::into)
}

fn rewrite_drawing(
    drawing: &[u8],
    shape_id: u32,
    kind: TextKind,
    before: &[u8],
    after: &[u8],
    matches: &mut usize,
) -> Result<Vec<u8>> {
    let data = rewrite_officeart(
        drawing_payload(drawing)?,
        shape_id,
        kind,
        before,
        after,
        0,
        matches,
    )?;
    rebuild(drawing, &data).map_err(Into::into)
}

fn rewrite_officeart(
    record: &[u8],
    shape_id: u32,
    kind: TextKind,
    before: &[u8],
    after: &[u8],
    depth: usize,
    matches: &mut usize,
) -> Result<Vec<u8>> {
    let mut output = Vec::with_capacity(record.len());
    let target = SingleRewrite {
        shape_id,
        kind,
        before,
        after,
    };
    rewrite_officeart_into(record, &target, depth, matches, &mut output)?;
    Ok(output)
}

struct SingleRewrite<'a> {
    shape_id: u32,
    kind: TextKind,
    before: &'a [u8],
    after: &'a [u8],
}

/// Appends the rewritten record to `output`, framing each container in
/// place; the same visits, checks, errors and bytes as a per-record rebuild
/// (record 0264, as [`rewrite_officeart_grouped_into`]).
fn rewrite_officeart_into(
    record: &[u8],
    target: &SingleRewrite<'_>,
    depth: usize,
    matches: &mut usize,
    output: &mut Vec<u8>,
) -> Result<()> {
    if depth >= 128 {
        return Err(
            PackageError::Corrupted("OfficeArt nesting exceeds text-edit limit".into()).into(),
        );
    }
    if record_type(record)? == OFFICEART_SP_CONTAINER
        && shape_id_of(record)? == Some(target.shape_id)
    {
        *matches = matches.saturating_add(1);
        output.extend_from_slice(&rewrite_shape_container(
            record,
            target.kind,
            target.before,
            target.after,
        )?);
        return Ok(());
    }
    if record_version(record)? != 0xF {
        output.extend_from_slice(record);
        return Ok(());
    }
    // `drawing_payload` proves the 8-byte header and the payload extent.
    let payload = drawing_payload(record)?;
    let start = output.len();
    output.extend_from_slice(&record[..PPT_HEADER_LEN]);
    output.reserve(payload.len());
    for child_result in children(record)? {
        rewrite_officeart_into(child_result?, target, depth + 1, matches, output)?;
    }
    let data_length = u32::try_from(output.len() - start - PPT_HEADER_LEN)
        .map_err(|_err| PackageError::ResourceLimit("PPT text-edit record exceeds u32".into()))?;
    output[start + 4..start + PPT_HEADER_LEN].copy_from_slice(&data_length.to_le_bytes());
    Ok(())
}

fn rewrite_shape_container(
    record: &[u8],
    kind: TextKind,
    before: &[u8],
    after: &[u8],
) -> Result<Vec<u8>> {
    let mut textbox_count = 0usize;
    let mut replaced = false;
    let mut data = Vec::with_capacity(drawing_payload(record)?.len());
    for child_result in children(record)? {
        let child = child_result?;
        if record_type(child)? != OFFICEART_CLIENT_TEXTBOX {
            data.extend_from_slice(child);
            continue;
        }
        textbox_count += 1;
        let rewritten = rewrite_textbox(child, kind, before, after)?;
        replaced = true;
        data.extend_from_slice(&rewritten);
    }
    if textbox_count == 0 {
        return Err(Error::Refused(Refusal::NoTextbox));
    }
    if textbox_count > 1 {
        return Err(Error::Refused(Refusal::AmbiguousTextbox));
    }
    if !replaced {
        return Err(Error::Refused(Refusal::NoTextAtom));
    }
    rebuild(record, &data).map_err(Into::into)
}

fn rewrite_textbox(textbox: &[u8], kind: TextKind, before: &[u8], after: &[u8]) -> Result<Vec<u8>> {
    let data = rewrite_text_records(children(textbox)?, kind, before, after)?;
    rebuild(textbox, &data).map_err(Into::into)
}

/// Rewrites the text atom of one text body and resizes the dependent ranges
/// [`resize_dependency_closure_is_modeled`] admits, returning the body's
/// re-encoded records. Every other record is copied exactly.
fn rewrite_text_records<'a>(
    records: impl Iterator<Item = PackageResult<&'a [u8]>>,
    kind: TextKind,
    before: &[u8],
    after: &[u8],
) -> Result<Vec<u8>> {
    let mut count = 0usize;
    let mut data = Vec::with_capacity(after.len().saturating_add(64));
    let old_text_units = text_units(kind, before.len())?;
    let new_text_units = text_units(kind, after.len())?;
    for child_result in records {
        let child = child_result?;
        let record_type = record_type(child)?;
        let is_target = match kind {
            TextKind::Bytes => record_type == RecordType::TextBytesAtom as u16,
            TextKind::Chars => record_type == RecordType::TextCharsAtom as u16,
        };
        if is_target {
            count += 1;
            if drawing_payload(child)? != before {
                return Err(PackageError::Corrupted(
                    "selected text atom changed before publication".into(),
                )
                .into());
            }
            let rewritten = rebuild(child, after)?;
            data.extend_from_slice(&rewritten);
        } else if record_type == RecordType::StyleTextPropAtom as u16
            && old_text_units != new_text_units
        {
            let rewritten = crate::StyleTextPropAtom::resize_single_run_record(
                child,
                old_text_units,
                new_text_units,
            )?
            .ok_or(Error::Refused(Refusal::DependencyClosure))?;
            data.extend_from_slice(&rewritten);
        } else if record_type == RecordType::TextSpecInfoAtom as u16
            && old_text_units != new_text_units
        {
            let rewritten = crate::TextSpecialInfoRuns::resize_single_run_record(
                child,
                old_text_units,
                new_text_units,
            )?
            .ok_or(Error::Refused(Refusal::DependencyClosure))?;
            data.extend_from_slice(&rewritten);
        } else {
            data.extend_from_slice(child);
        }
    }
    match count {
        0 => Err(Error::Refused(Refusal::NoTextAtom)),
        1 => Ok(data),
        _ => Err(Error::Refused(Refusal::MultipleTextAtoms)),
    }
}

fn visit_officeart(
    record: &[u8],
    shape_id: u32,
    depth: usize,
    matches: &mut usize,
    visit: &mut impl FnMut(&[u8]) -> Result<()>,
) -> Result<()> {
    if depth >= 128 {
        return Err(
            PackageError::Corrupted("OfficeArt nesting exceeds text-edit limit".into()).into(),
        );
    }
    if record_type(record)? == OFFICEART_SP_CONTAINER && shape_id_of(record)? == Some(shape_id) {
        *matches = matches.saturating_add(1);
        let mut textbox = None;
        for child_result in children(record)? {
            let child = child_result?;
            if record_type(child)? == OFFICEART_CLIENT_TEXTBOX && textbox.replace(child).is_some() {
                return Err(Error::Refused(Refusal::AmbiguousTextbox));
            }
        }
        let selected_textbox = textbox.ok_or(Error::Refused(Refusal::NoTextbox))?;
        visit(selected_textbox)?;
        return Ok(());
    }
    if record_version(record)? == 0xF {
        for child_result in children(record)? {
            let child = child_result?;
            visit_officeart(child, shape_id, depth + 1, matches, visit)?;
        }
    }
    Ok(())
}

fn visit_shape_container(
    record: &[u8],
    shape_id: u32,
    depth: usize,
    matches: &mut usize,
    visit: &mut impl FnMut(&[u8]) -> Result<()>,
) -> Result<()> {
    if depth >= 128 {
        return Err(PackageError::Corrupted(
            "OfficeArt nesting exceeds shape-inspection limit".into(),
        )
        .into());
    }
    if record_type(record)? == OFFICEART_SP_CONTAINER && shape_id_of(record)? == Some(shape_id) {
        *matches = matches.saturating_add(1);
        visit(record)?;
        return Ok(());
    }
    if record_version(record)? == 0xF {
        for child_result in children(record)? {
            visit_shape_container(child_result?, shape_id, depth + 1, matches, visit)?;
        }
    }
    Ok(())
}

pub(crate) fn shape_id_of(record: &[u8]) -> PackageResult<Option<u32>> {
    for child_result in children(record)? {
        let child = child_result?;
        if record_type(child)? == OFFICEART_SP {
            let payload = drawing_payload(child)?;
            let bytes: &[u8; 4] = payload
                .get(..4)
                .and_then(|value| value.try_into().ok())
                .ok_or_else(|| {
                    PackageError::Corrupted("OfficeArt Sp record has no shape identity".into())
                })?;
            return Ok(Some(u32::from_le_bytes(*bytes)));
        }
    }
    Ok(None)
}

/// The slide-list body index an outline-backed shape's `ClientTextbox`
/// selects, or `None` when the shape is not outline-backed.
///
/// The textbox must hold exactly one `OutlineTextRefAtom` and no text atom;
/// a `TextRulerAtom` beside it is allowed. A negative index selects nothing
/// and is reported as a missing text atom.
fn outline_reference(slide: &[u8], shape_id: u32) -> Result<Option<usize>> {
    let (_, consumed) = crate::Record::parse_strict(slide, 0)?;
    if consumed != slide.len() {
        return Err(PackageError::Corrupted("selected slide has trailing bytes".into()).into());
    }
    let drawing = find_ppt_record(slide, RecordType::PPDrawing as u16, 0)?
        .ok_or(Error::Refused(Refusal::ShapeNotFound))?;
    let mut matches = 0usize;
    let mut reference = None;
    visit_officeart(
        drawing_payload(drawing)?,
        shape_id,
        0,
        &mut matches,
        &mut |textbox| {
            let mut found = None;
            for child_result in children(textbox)? {
                let child = child_result?;
                let child_type = record_type(child)?;
                if child_type == RecordType::OutlineTextRefAtom as u16 {
                    if found.is_some() {
                        return Ok(());
                    }
                    let payload: [u8; 4] = drawing_payload(child)?.try_into().map_err(|_err| {
                        PackageError::Corrupted("malformed OutlineTextRefAtom".into())
                    })?;
                    found = Some(i32::from_le_bytes(payload));
                } else if child_type != RecordType::TextRulerAtom as u16 {
                    return Ok(());
                }
            }
            reference = found.and_then(|index| usize::try_from(index).ok());
            Ok(())
        },
    )?;
    match matches {
        0 => Err(Error::Refused(Refusal::ShapeNotFound)),
        1 => Ok(reference),
        _ => Err(Error::Refused(Refusal::AmbiguousShape)),
    }
}

/// Visits the records of every text body of each slide in the live
/// document's slide `SlideListWithTextContainer` (instance 0).
///
/// Bodies are numbered per slide in slide-list order, as an
/// `OutlineTextRefAtom` counts them; records before a slide's first
/// `TextHeaderAtom` belong to no body. `visit` receives the slide persist
/// identifier, the body index and the body's records.
fn for_each_outline_body<'a>(
    slide_list: &'a [u8],
    mut visit: impl FnMut(Option<(u32, usize)>, &[&'a [u8]]) -> Result<()>,
) -> Result<()> {
    let mut slide = None;
    let mut body = None::<usize>;
    let mut segment: Vec<&'a [u8]> = Vec::new();
    for child_result in children(slide_list)? {
        let child = child_result?;
        let child_type = record_type(child)?;
        let starts_segment = child_type == RecordType::SlidePersistAtom as u16
            || child_type == RecordType::TextHeaderAtom as u16;
        if starts_segment {
            let key = slide.zip(body);
            visit(key, &segment)?;
            segment.clear();
            if child_type == RecordType::SlidePersistAtom as u16 {
                let payload = drawing_payload(child)?;
                let persist_id: [u8; 4] = payload
                    .get(..4)
                    .and_then(|value| value.try_into().ok())
                    .ok_or_else(|| PackageError::Corrupted("malformed SlidePersistAtom".into()))?;
                slide = Some(u32::from_le_bytes(persist_id));
                body = None;
            } else {
                body = Some(body.map_or(0, |index| index.saturating_add(1)));
            }
        }
        segment.push(child);
    }
    visit(slide.zip(body), &segment)
}

/// The live document's slide `SlideListWithTextContainer`.
fn slide_list_of(document: &[u8]) -> Result<&[u8]> {
    if record_type(document)? != RecordType::Document as u16 {
        return Err(
            PackageError::Corrupted("live PPT record is not a DocumentContainer".into()).into(),
        );
    }
    let mut found = None;
    for child_result in children(document)? {
        let child = child_result?;
        if record_type(child)? == RecordType::SlideListWithText as u16
            && record_instance(child)? == 0
            && found.replace(child).is_some()
        {
            return Err(
                PackageError::Corrupted("PPT document has several slide lists".into()).into(),
            );
        }
    }
    found.ok_or(Error::Refused(Refusal::NoTextAtom))
}

/// Inspects the slide-list text body an outline-backed shape selects.
fn inspect_outline_body(document: &[u8], slide_persist_id: u32, body: usize) -> Result<TextAtom> {
    let mut selected = None;
    for_each_outline_body(slide_list_of(document)?, |key, records| {
        if key == Some((slide_persist_id, body)) {
            if selected.is_some() {
                return Err(Error::Refused(Refusal::AmbiguousTextbox));
            }
            selected = Some(inspect_text_records(
                records.iter().map(|record| Ok(*record)),
            )?);
        }
        Ok(())
    })?;
    selected.ok_or(Error::Refused(Refusal::NoTextAtom))
}

/// Rewrites selected slide-list text bodies of the live document; every
/// other record of the document is copied exactly.
fn rewrite_outline_bodies(
    document: &[u8],
    rewrites: &BTreeMap<(u32, usize), ShapeTextRewrite<'_>>,
) -> Result<Vec<u8>> {
    let slide_list = slide_list_of(document)?;
    let mut list_data = Vec::with_capacity(drawing_payload(slide_list)?.len().saturating_add(64));
    let mut applied = 0usize;
    for_each_outline_body(slide_list, |key, records| {
        match key.and_then(|key| rewrites.get(&key)) {
            Some(rewrite) => {
                let body = rewrite_text_records(
                    records.iter().map(|record| Ok(*record)),
                    rewrite.kind,
                    rewrite.before,
                    rewrite.after,
                )?;
                list_data.extend_from_slice(&body);
                applied = applied.saturating_add(1);
            },
            None => {
                for record in records {
                    list_data.extend_from_slice(record);
                }
            },
        }
        Ok(())
    })?;
    if applied != rewrites.len() {
        return Err(PackageError::Corrupted(
            "PPT outline text rewrite did not find every selected body".into(),
        )
        .into());
    }
    let rebuilt_list = rebuild(slide_list, &list_data)?;
    let mut document_data = Vec::with_capacity(drawing_payload(document)?.len().saturating_add(64));
    for child_result in children(document)? {
        let child = child_result?;
        if std::ptr::eq(child.as_ptr(), slide_list.as_ptr()) {
            document_data.extend_from_slice(&rebuilt_list);
        } else {
            document_data.extend_from_slice(child);
        }
    }
    rebuild(document, &document_data).map_err(Into::into)
}

fn encode_replacement(value: &str, kind: TextKind) -> Result<Vec<u8>> {
    if value.contains('\0') {
        return Err(Error::Refused(Refusal::IncompatibleEncoding));
    }
    let bytes = match kind {
        TextKind::Bytes => value
            .chars()
            .map(|character| {
                u8::try_from(u32::from(character))
                    .map_err(|_err| Error::Refused(Refusal::IncompatibleEncoding))
            })
            .collect::<Result<Vec<_>>>()?,
        TextKind::Chars => value.encode_utf16().flat_map(u16::to_le_bytes).collect(),
    };
    Ok(bytes)
}

fn encoded_replacement_len(value: &str, kind: TextKind) -> Result<usize> {
    if value.contains('\0') {
        return Err(Error::Refused(Refusal::IncompatibleEncoding));
    }
    match kind {
        TextKind::Bytes => value.chars().try_fold(0usize, |length, character| {
            u8::try_from(u32::from(character))
                .map_err(|_err| Error::Refused(Refusal::IncompatibleEncoding))?;
            length.checked_add(1).ok_or_else(|| {
                PackageError::ResourceLimit(
                    "PPT shape-text replacement encoded length overflow".into(),
                )
                .into()
            })
        }),
        TextKind::Chars => value.encode_utf16().count().checked_mul(2).ok_or_else(|| {
            PackageError::ResourceLimit("PPT shape-text replacement encoded length overflow".into())
                .into()
        }),
    }
}

fn text_units(kind: TextKind, byte_length: usize) -> Result<usize> {
    match kind {
        TextKind::Bytes => Ok(byte_length),
        TextKind::Chars if byte_length.is_multiple_of(2) => Ok(byte_length / 2),
        TextKind::Chars => Err(Error::Refused(Refusal::IncompatibleEncoding)),
    }
}

pub(crate) fn decode_utf16(bytes: &[u8]) -> Result<String> {
    if !bytes.len().is_multiple_of(2) {
        return Err(Error::Refused(Refusal::IncompatibleEncoding));
    }
    let units = bytes
        .as_chunks::<2>()
        .0
        .iter()
        .map(|pair| u16::from_le_bytes([pair[0], pair[1]]))
        .collect::<Vec<_>>();
    String::from_utf16(&units).map_err(|_err| Error::Refused(Refusal::IncompatibleEncoding))
}

pub(crate) fn record_type(record: &[u8]) -> PackageResult<u16> {
    let bytes: &[u8; 2] = record
        .get(2..4)
        .and_then(|value| value.try_into().ok())
        .ok_or_else(|| PackageError::Corrupted("truncated text-edit record header".into()))?;
    Ok(u16::from_le_bytes(*bytes))
}

pub(crate) fn record_version(record: &[u8]) -> PackageResult<u16> {
    let bytes: &[u8; 2] = record
        .get(..2)
        .and_then(|value| value.try_into().ok())
        .ok_or_else(|| PackageError::Corrupted("truncated text-edit record header".into()))?;
    Ok(u16::from_le_bytes(*bytes) & 0x000F)
}

pub(crate) fn record_instance(record: &[u8]) -> PackageResult<u16> {
    let bytes: &[u8; 2] = record
        .get(..2)
        .and_then(|value| value.try_into().ok())
        .ok_or_else(|| PackageError::Corrupted("truncated text-edit record header".into()))?;
    Ok(u16::from_le_bytes(*bytes) >> 4)
}

pub(crate) fn drawing_payload(record: &[u8]) -> PackageResult<&[u8]> {
    let bytes: &[u8; 4] = record
        .get(4..8)
        .and_then(|value| value.try_into().ok())
        .ok_or_else(|| PackageError::Corrupted("truncated text-edit record length".into()))?;
    let length = usize::try_from(u32::from_le_bytes(*bytes)).map_err(|_err| {
        PackageError::Corrupted("text-edit record length exceeds this platform".into())
    })?;
    let end = PPT_HEADER_LEN
        .checked_add(length)
        .ok_or_else(|| PackageError::Corrupted("text-edit record length overflows".into()))?;
    record
        .get(PPT_HEADER_LEN..end)
        .ok_or_else(|| PackageError::Corrupted("truncated text-edit record payload".into()))
}

pub(crate) fn children(record: &[u8]) -> PackageResult<impl Iterator<Item = PackageResult<&[u8]>>> {
    let payload = drawing_payload(record)?;
    Ok(ChildRecords { payload, offset: 0 })
}

struct ChildRecords<'a> {
    payload: &'a [u8],
    offset: usize,
}

impl<'a> Iterator for ChildRecords<'a> {
    type Item = PackageResult<&'a [u8]>;

    fn next(&mut self) -> Option<Self::Item> {
        if self.offset == self.payload.len() {
            return None;
        }
        let start = self.offset;
        let record = match drawing_payload(self.payload.get(start..)?) {
            Ok(payload) => match PPT_HEADER_LEN.checked_add(payload.len()) {
                Some(length) => self.payload.get(start..start + length),
                None => None,
            },
            Err(error) => return Some(Err(error)),
        };
        if let Some(child_record) = record {
            self.offset = start + child_record.len();
            Some(Ok(child_record))
        } else {
            self.offset = self.payload.len();
            Some(Err(PackageError::Corrupted(
                "truncated text-edit child record".into(),
            )))
        }
    }
}

pub(crate) fn rebuild(record: &[u8], data: &[u8]) -> PackageResult<Vec<u8>> {
    let _ = drawing_payload(record)?;
    let data_length = u32::try_from(data.len())
        .map_err(|_err| PackageError::ResourceLimit("PPT text-edit record exceeds u32".into()))?;
    let mut output = Vec::with_capacity(record.len());
    output.extend_from_slice(&record[..4]);
    output.extend_from_slice(&data_length.to_le_bytes());
    output.extend_from_slice(data);
    Ok(output)
}

#[cfg(test)]
#[allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]
mod tests {
    use super::{Error, Position, Refusal, Snapshot, Target};
    use crate::raw::Reader;
    use crate::writer::Builder;
    use litchi_cfb::{
        OleFile, OutputProgress, OverlayError, SharedOleFile, SharedOleFileLimits,
        StreamSpliceLimits,
    };
    use litchi_core::{OwnedSource, ReadAt, SourceVersion};
    use std::collections::BTreeSet;
    use std::io::{self, Cursor, Write};
    use std::sync::{
        Arc, Mutex,
        atomic::{AtomicBool, AtomicU64, AtomicUsize, Ordering},
    };

    fn fixture(text: &str) -> Vec<u8> {
        let mut writer = Builder::new();
        let slide = writer.add_slide().unwrap();
        writer.add_textbox(slide, 10, 10, 240, 40, text).unwrap();
        let mut output = Cursor::new(Vec::new());
        crate::raw::encode(&mut writer, &mut output).unwrap();
        output.into_inner()
    }

    fn target() -> Target {
        Target::new(Position::new(0), Position::new(0))
    }

    fn generated_fixture(slide_count: usize, shapes_per_slide: usize) -> Vec<u8> {
        let mut writer = Builder::new();
        for slide_index in 0..slide_count {
            let slide = writer.add_slide().unwrap();
            for shape_index in 0..shapes_per_slide {
                let x = 10 + i32::try_from(shape_index).unwrap() * 260;
                let text = format!("slide-{slide_index:02}-shape-{shape_index:02}");
                writer.add_textbox(slide, x, 10, 240, 40, &text).unwrap();
            }
        }
        let mut output = Cursor::new(Vec::new());
        crate::raw::encode(&mut writer, &mut output).unwrap();
        output.into_inner()
    }

    #[test]
    fn grouped_shape_rewrite_matches_sequential_wire_output() {
        let source = generated_fixture(1, 4);
        let presentation = super::presentation(&source).unwrap();
        let slide = &presentation.slides().unwrap()[0];
        let slide_record = super::resolve(&source, Target::new(Position::new(0), Position::new(0)))
            .unwrap()
            .slide_record;
        let mut plans = Vec::new();
        let mut sequential = slide_record.to_vec();

        for (shape_position, shape) in slide.shapes().unwrap().iter().enumerate() {
            let target = Target::new(Position::new(0), Position::new(shape_position));
            let resolved = super::resolve(&source, target).unwrap();
            let mut after = resolved.payload.clone();
            assert!(!after.is_empty());
            after[0] = b'X';
            sequential = super::rewrite_slide(
                &sequential,
                super::native_shape_id(shape),
                resolved.kind,
                &resolved.payload,
                &after,
            )
            .unwrap();
            plans.push((
                super::native_shape_id(shape),
                resolved.kind,
                resolved.payload,
                after,
            ));
        }

        let grouped_plans = plans
            .iter()
            .map(|(shape_id, kind, before, after)| super::ShapeTextRewrite {
                shape_id: *shape_id,
                kind: *kind,
                before,
                after,
            })
            .collect::<Vec<_>>();
        let grouped = super::rewrite_slide_grouped(&slide_record, &grouped_plans).unwrap();
        assert_eq!(grouped, sequential);
    }

    #[derive(Debug, Default)]
    struct ReadMetrics {
        calls: usize,
        bytes: usize,
        max_request: usize,
    }

    struct CountingSource {
        bytes: Arc<Vec<u8>>,
        metrics: Arc<Mutex<ReadMetrics>>,
        revision: Arc<AtomicU64>,
    }

    impl CountingSource {
        fn new(bytes: Vec<u8>) -> (Self, Arc<Mutex<ReadMetrics>>) {
            let metrics = Arc::new(Mutex::new(ReadMetrics::default()));
            let source = Self {
                bytes: Arc::new(bytes),
                metrics: Arc::clone(&metrics),
                revision: Arc::new(AtomicU64::new(0)),
            };
            (source, metrics)
        }

        fn reset(metrics: &Arc<Mutex<ReadMetrics>>) {
            *metrics.lock().unwrap() = ReadMetrics::default();
        }
    }

    impl ReadAt for CountingSource {
        fn len(&self) -> io::Result<u64> {
            u64::try_from(self.bytes.len())
                .map_err(|_error| io::Error::other("counting source length overflow"))
        }

        fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
            let start = usize::try_from(offset).unwrap_or(usize::MAX);
            let available = self.bytes.get(start..).unwrap_or_default();
            let count = available.len().min(output.len());
            output[..count].copy_from_slice(&available[..count]);
            let mut metrics = self.metrics.lock().unwrap();
            metrics.calls = metrics.calls.saturating_add(1);
            metrics.bytes = metrics.bytes.saturating_add(count);
            metrics.max_request = metrics.max_request.max(output.len());
            Ok(count)
        }

        fn version(&self) -> io::Result<SourceVersion> {
            Ok(SourceVersion::new(
                0x5050_5443,
                self.revision.load(Ordering::Acquire),
            ))
        }
    }

    /// Hostile adapter that flips one tail byte after an exact read ordinal
    /// while keeping a stable version token.
    struct ScheduledMutationSource {
        bytes: Mutex<Vec<u8>>,
        reads: Arc<AtomicUsize>,
        request_sizes: Mutex<Vec<usize>>,
        mutate_after_read: usize,
        fired: Arc<AtomicBool>,
    }

    impl ScheduledMutationSource {
        fn new(bytes: Vec<u8>, mutate_after_read: usize) -> Self {
            Self {
                bytes: Mutex::new(bytes),
                reads: Arc::new(AtomicUsize::new(0)),
                request_sizes: Mutex::new(Vec::new()),
                mutate_after_read,
                fired: Arc::new(AtomicBool::new(false)),
            }
        }

        fn current_bytes(&self) -> Vec<u8> {
            self.bytes.lock().unwrap().clone()
        }

        /// Ordinal of the first read that requests the complete artifact,
        /// which is the first chunk of the first complete identity pass.
        fn first_identity_read(&self, length: usize) -> usize {
            self.request_sizes
                .lock()
                .unwrap()
                .iter()
                .position(|size| *size == length)
                .expect("a complete-artifact fingerprint request")
                + 1
        }
    }

    impl ReadAt for ScheduledMutationSource {
        fn len(&self) -> io::Result<u64> {
            Ok(self.bytes.lock().unwrap().len() as u64)
        }

        fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
            let call = self.reads.fetch_add(1, Ordering::SeqCst) + 1;
            self.request_sizes.lock().unwrap().push(output.len());
            let mut bytes = self.bytes.lock().unwrap();
            let start = usize::try_from(offset).unwrap_or(usize::MAX);
            let Some(available) = bytes.get(start..) else {
                return Ok(0);
            };
            let count = available.len().min(output.len());
            output[..count].copy_from_slice(&available[..count]);
            if call == self.mutate_after_read
                && !self.fired.swap(true, Ordering::SeqCst)
                && let Some(last) = bytes.last_mut()
            {
                *last ^= 0x01;
            }
            Ok(count)
        }

        fn version(&self) -> io::Result<SourceVersion> {
            Ok(SourceVersion::new(0x5050_0589, 0))
        }
    }

    /// Change 0589: the text-edit snapshot's identity fence still refuses a
    /// stable-token mutation placed in every gap between its fence points.
    ///
    /// `SourceSnapshot::open` is one empty-splice identity plan: two complete
    /// source reads, formerly four SHA-256 passes and now two, because with no
    /// span the composed target *is* the source. Eliding the duplicate hasher
    /// must not move a fence point, so once the first complete identity pass
    /// has started, every later read of the open must be refused — except the
    /// very last, which no finite number of passes can bracket and which the
    /// commit fence catches instead.
    #[test]
    fn a_stable_token_mutation_after_every_open_read_is_still_refused() {
        let bytes = fixture("abc");
        let clean = Arc::new(ScheduledMutationSource::new(bytes.clone(), usize::MAX));
        let clean_source: Arc<dyn ReadAt> = clean.clone();
        super::SourceSnapshot::open(clean_source).expect("clean source-backed PPT fixture");
        let reads = clean.reads.load(Ordering::SeqCst);
        let first_identity = clean.first_identity_read(bytes.len());
        assert!(
            reads > first_identity,
            "an open hashes after its first pass"
        );

        let mut refused = 0_usize;
        for ordinal in 1..=reads {
            let hostile = Arc::new(ScheduledMutationSource::new(bytes.clone(), ordinal));
            let source: Arc<dyn ReadAt> = hostile.clone();
            let outcome = super::SourceSnapshot::open(source);
            assert!(
                hostile.fired.load(Ordering::SeqCst),
                "read {ordinal} never fired"
            );
            match outcome {
                Err(Error::Source(OverlayError::SourceFingerprintChanged { .. })) => {
                    assert!(ordinal >= first_identity);
                    refused += 1;
                },
                Err(other) => panic!("read {ordinal} produced an unexpected error: {other:?}"),
                Ok(snapshot) => {
                    let settled = super::SourceSnapshot::open(Arc::new(OwnedSource::new(
                        hostile.current_bytes(),
                    )))
                    .expect("mutated artifact reopens");
                    if ordinal < first_identity {
                        // The mutation preceded the first identity pass, so the
                        // snapshot is bound to exactly the bytes that exist.
                        assert_eq!(snapshot.fingerprint(), settled.fingerprint());
                    } else {
                        assert_eq!(
                            ordinal, reads,
                            "read {ordinal} opened over an artifact that changed inside the fence"
                        );
                        assert_ne!(snapshot.fingerprint(), settled.fingerprint());
                        // The commit fence re-plans the identity and must refuse.
                        let mut transaction = snapshot.edit_text(target()).expect("first shape");
                        let replacement = transaction.text().to_owned();
                        transaction.set_text(replacement).expect("same width");
                        assert!(matches!(
                            transaction.commit(),
                            Err(Error::Source(OverlayError::SourceFingerprintChanged { .. }))
                        ));
                    }
                },
            }
        }
        assert_eq!(refused, reads - first_identity);
    }

    /// Change 0589: an empty-splice identity plan reports one digest for both
    /// identities, and it is the same digest the owned path computes.
    #[test]
    fn the_identity_plan_reports_one_digest_for_both_identities() {
        let bytes = fixture("abc");
        let shared = SharedOleFile::open(
            Arc::new(OwnedSource::new(bytes.clone())),
            SharedOleFileLimits::default(),
        )
        .unwrap();
        let plan = shared
            .plan_same_length_stream_splices(Vec::new(), StreamSpliceLimits::default())
            .unwrap();
        assert!(plan.is_noop());
        assert_eq!(plan.source_fingerprint(), plan.target_fingerprint());

        let owned = super::SourceSnapshot::open(Arc::new(OwnedSource::new(bytes.clone()))).unwrap();
        assert_eq!(owned.fingerprint(), plan.source_fingerprint());

        let generic: Arc<dyn ReadAt> = Arc::new(ScheduledMutationSource::new(bytes, usize::MAX));
        let generic = super::SourceSnapshot::open(generic).unwrap();
        assert_eq!(generic.fingerprint(), plan.source_fingerprint());
    }

    fn resolve_owned_and_range(
        bytes: Vec<u8>,
        target: Target,
    ) -> (
        super::SourceResolved,
        super::SourceResolved,
        Arc<Mutex<ReadMetrics>>,
    ) {
        let (counting, metrics) = CountingSource::new(bytes);
        let source: Arc<dyn ReadAt> = Arc::new(counting);
        let shared =
            SharedOleFile::open(Arc::clone(&source), SharedOleFileLimits::default()).unwrap();
        let document_path =
            super::select_stream_path(&shared, &super::DOCUMENT_PATHS, "PowerPoint Document")
                .unwrap();
        let current_user_path =
            super::select_stream_path(&shared, &super::CURRENT_USER_PATHS, "Current User").unwrap();
        let limits = crate::SourceLimits::default();
        let owned = super::resolve_source_target_owned(
            &shared,
            &document_path,
            &current_user_path,
            target,
            limits,
        )
        .unwrap();
        CountingSource::reset(&metrics);
        let ranged = super::source_range::resolve_source_target(
            &shared,
            &document_path,
            &current_user_path,
            target,
            limits,
        )
        .unwrap();
        (owned, ranged, metrics)
    }

    fn live_document_bounds(source: &[u8]) -> (Vec<u8>, usize, usize) {
        let mut ole = OleFile::open(
            Cursor::new(source.to_vec()),
            litchi_cfb::OleFileLimits::default(),
        )
        .unwrap();
        let document = ole.open_stream(&["PowerPoint Document"]).unwrap();
        let current_user = ole.open_stream(&["Current User"]).unwrap();
        let directory = crate::slide::SlideDirectory::build(
            &document,
            &current_user,
            &crate::embedded::object::Editor::inspect_live_mapping(&document, &current_user)
                .unwrap(),
        )
        .unwrap();
        let offset = directory.document_offset();
        let (_record, consumed) = crate::Record::parse_strict_with_limits(
            &document,
            offset,
            crate::SourceLimits::default(),
        )
        .unwrap();
        (document, offset, offset + consumed)
    }

    fn record_children_offsets(data: &[u8], record_offset: usize) -> Vec<usize> {
        let payload_len = u32::from_le_bytes(
            data[record_offset + 4..record_offset + 8]
                .try_into()
                .unwrap(),
        ) as usize;
        let payload_start = record_offset + 8;
        let payload_end = payload_start + payload_len;
        let mut children = Vec::new();
        let mut offset = payload_start;
        while offset < payload_end {
            children.push(offset);
            let child_len =
                u32::from_le_bytes(data[offset + 4..offset + 8].try_into().unwrap()) as usize;
            offset += 8 + child_len;
        }
        assert_eq!(offset, payload_end);
        children
    }

    fn record_type_at(data: &[u8], offset: usize) -> u16 {
        u16::from_le_bytes(data[offset + 2..offset + 4].try_into().unwrap())
    }

    fn mutate_document_stream(
        source: Vec<u8>,
        mutate: impl FnOnce(&mut Vec<u8>, usize, usize),
    ) -> Vec<u8> {
        let (mut document, document_offset, document_end) = live_document_bounds(&source);
        mutate(&mut document, document_offset, document_end);
        with_stream(source, &["PowerPoint Document"], &document)
    }

    fn find_nested_record_offset(data: &[u8], record_offset: usize, wanted: u16) -> Option<usize> {
        if record_type_at(data, record_offset) == wanted {
            return Some(record_offset);
        }
        let record_type = crate::RecordType::from(record_type_at(data, record_offset));
        if !crate::Record::is_container_record(record_type) {
            return None;
        }
        record_children_offsets(data, record_offset)
            .into_iter()
            .find_map(|child| find_nested_record_offset(data, child, wanted))
    }

    fn raw_record(version: u16, instance: u16, record_type: u16, payload: &[u8]) -> Vec<u8> {
        let packed = (version & 0x000F) | ((instance & 0x0FFF) << 4);
        let length = u32::try_from(payload.len()).unwrap();
        let mut bytes = Vec::with_capacity(8 + payload.len());
        bytes.extend_from_slice(&packed.to_le_bytes());
        bytes.extend_from_slice(&record_type.to_le_bytes());
        bytes.extend_from_slice(&length.to_le_bytes());
        bytes.extend_from_slice(payload);
        bytes
    }

    fn current_user_stream(edit_offset: u32) -> Vec<u8> {
        let mut stream = Vec::with_capacity(32);
        stream.extend_from_slice(&[0, 0, 0xF6, 0x0F]);
        stream.extend_from_slice(&24_u32.to_le_bytes());
        stream.extend_from_slice(&20_u32.to_le_bytes());
        stream.extend_from_slice(&0xE391_C05F_u32.to_le_bytes());
        stream.extend_from_slice(&edit_offset.to_le_bytes());
        stream.extend_from_slice(&0_u16.to_le_bytes());
        stream.extend_from_slice(&0x03F4_u16.to_le_bytes());
        stream.extend_from_slice(&[3, 0, 0, 0]);
        stream.extend_from_slice(&8_u32.to_le_bytes());
        stream
    }

    fn synthetic_history_source(
        historical_version: u16,
        historical_instance: u16,
        historical_payload_len: usize,
        historical_directory_offset: Option<u32>,
    ) -> Vec<u8> {
        use litchi_cfb::OleWriter;

        let historical_directory = raw_record(0, 0, 6002, &[]);
        let historical_offset = historical_directory.len();
        let mut historical_payload = vec![0_u8; historical_payload_len];
        if historical_payload_len >= 16 {
            historical_payload[12..16].copy_from_slice(
                &historical_directory_offset
                    .unwrap_or_else(|| u32::try_from(historical_offset).unwrap())
                    .to_le_bytes(),
            );
        }
        let historical = raw_record(
            historical_version,
            historical_instance,
            4085,
            &historical_payload,
        );
        let current_directory_offset = historical_offset + historical.len();
        let current_directory = raw_record(0, 0, 6002, &[]);
        let current_edit_offset = current_directory_offset + current_directory.len();
        let mut current_payload = vec![0_u8; 28];
        current_payload[12..16].copy_from_slice(
            &u32::try_from(current_directory_offset)
                .unwrap()
                .to_le_bytes(),
        );
        current_payload[16..20].copy_from_slice(&1_u32.to_le_bytes());
        current_payload[20..24].copy_from_slice(&1_u32.to_le_bytes());
        current_payload[8..12]
            .copy_from_slice(&u32::try_from(historical_offset).unwrap().to_le_bytes());
        let current_edit = raw_record(0, 0, 4085, &current_payload);
        let document = [
            historical_directory,
            historical,
            current_directory,
            current_edit,
        ]
        .concat();

        let mut writer = OleWriter::new();
        writer
            .create_stream(&["PowerPoint Document"], &document)
            .unwrap();
        writer
            .create_stream(
                &["Current User"],
                &current_user_stream(u32::try_from(current_edit_offset).unwrap()),
            )
            .unwrap();
        let mut output = Cursor::new(Vec::new());
        writer.write_to(&mut output).unwrap();
        output.into_inner()
    }

    fn resolve_synthetic_history(source: Vec<u8>) -> Result<(), Error> {
        let shared = SharedOleFile::open(
            Arc::new(OwnedSource::new(source)),
            SharedOleFileLimits::default(),
        )
        .unwrap();
        super::source_range::resolve_source_target(
            &shared,
            &["PowerPoint Document".to_string()],
            &["Current User".to_string()],
            target(),
            crate::SourceLimits::default(),
        )
        .map(|_| ())
    }

    fn with_storage(source: Vec<u8>, path: &[&str]) -> Vec<u8> {
        use litchi_cfb::{OleFile, OleWriter};

        let mut source_file =
            OleFile::open(Cursor::new(source), litchi_cfb::OleFileLimits::default()).unwrap();
        let streams = source_file
            .list_streams()
            .into_iter()
            .map(|stream_path| {
                let refs = stream_path.iter().map(String::as_str).collect::<Vec<_>>();
                let data = source_file.open_stream(&refs).unwrap();
                (stream_path, data)
            })
            .collect::<Vec<_>>();
        let mut writer = OleWriter::new();
        for (stream_path, data) in streams {
            let refs = stream_path.iter().map(String::as_str).collect::<Vec<_>>();
            writer.create_stream(&refs, &data).unwrap();
        }
        writer.create_storage(path).unwrap();
        let mut output = Cursor::new(Vec::new());
        writer.write_to(&mut output).unwrap();
        output.into_inner()
    }

    fn with_stream(source: Vec<u8>, path: &[&str], data: &[u8]) -> Vec<u8> {
        use litchi_cfb::{OleFile, OleWriter};

        let mut source_file =
            OleFile::open(Cursor::new(source), litchi_cfb::OleFileLimits::default()).unwrap();
        let streams = source_file
            .list_streams()
            .into_iter()
            .map(|stream_path| {
                let refs = stream_path.iter().map(String::as_str).collect::<Vec<_>>();
                let data = source_file.open_stream(&refs).unwrap();
                (stream_path, data)
            })
            .collect::<Vec<_>>();
        let mut writer = OleWriter::new();
        for (stream_path, stream_data) in streams {
            let refs = stream_path.iter().map(String::as_str).collect::<Vec<_>>();
            writer.create_stream(&refs, &stream_data).unwrap();
        }
        writer.create_stream(path, data).unwrap();
        let mut output = Cursor::new(Vec::new());
        writer.write_to(&mut output).unwrap();
        output.into_inner()
    }

    fn relocate_root_stream_to_dual(source: Vec<u8>, name: &str) -> Vec<u8> {
        use litchi_cfb::{OleFile, OleWriter};

        let mut source_file =
            OleFile::open(Cursor::new(source), litchi_cfb::OleFileLimits::default()).unwrap();
        let streams = source_file
            .list_streams()
            .into_iter()
            .map(|stream_path| {
                let refs = stream_path.iter().map(String::as_str).collect::<Vec<_>>();
                let data = source_file.open_stream(&refs).unwrap();
                (stream_path, data)
            })
            .collect::<Vec<_>>();
        let mut relocated = None;
        let mut writer = OleWriter::new();
        for (stream_path, data) in streams {
            if stream_path.len() == 1 && stream_path[0] == name {
                relocated = Some(data);
            } else {
                let refs = stream_path.iter().map(String::as_str).collect::<Vec<_>>();
                writer.create_stream(&refs, &data).unwrap();
            }
        }
        writer.create_storage(&["PP97_DUALSTORAGE"]).unwrap();
        writer
            .create_stream(
                &["PP97_DUALSTORAGE", name],
                &relocated.expect("fixture root stream must exist"),
            )
            .unwrap();
        let mut output = Cursor::new(Vec::new());
        writer.write_to(&mut output).unwrap();
        output.into_inner()
    }

    fn source_officeart_shape(shape_id: u32, text: &[u8]) -> Vec<u8> {
        let shape = raw_record(2, 0, super::OFFICEART_SP, &shape_id.to_le_bytes());
        let header = raw_record(
            0,
            0,
            crate::RecordType::TextHeaderAtom as u16,
            &0_u32.to_le_bytes(),
        );
        let text = raw_record(0, 0, crate::RecordType::TextBytesAtom as u16, text);
        let textbox = raw_record(
            0x0F,
            0,
            super::OFFICEART_CLIENT_TEXTBOX,
            &[header, text].concat(),
        );
        raw_record(
            0x0F,
            0,
            super::OFFICEART_SP_CONTAINER,
            &[shape, textbox].concat(),
        )
    }

    fn find_officeart_shape_offset(
        data: &[u8],
        record_offset: usize,
        shape_id: u32,
    ) -> Option<usize> {
        fn find(data: &[u8], record_offset: usize, shape_id: u32, depth: usize) -> Option<usize> {
            assert!(depth < 128, "test OfficeArt fixture exceeded its bound");
            if record_type_at(data, record_offset) == super::OFFICEART_SP_CONTAINER {
                let shape = record_children_offsets(data, record_offset)
                    .into_iter()
                    .find(|&child| record_type_at(data, child) == super::OFFICEART_SP)
                    .and_then(|child| {
                        let payload_len =
                            u32::from_le_bytes(data[child + 4..child + 8].try_into().unwrap())
                                as usize;
                        let payload = &data[child + 8..child + 8 + payload_len];
                        payload
                            .get(..4)
                            .and_then(|value| value.try_into().ok())
                            .map(u32::from_le_bytes)
                    });
                if shape == Some(shape_id) {
                    return Some(record_offset);
                }
            }
            let version =
                u16::from_le_bytes(data[record_offset..record_offset + 2].try_into().unwrap())
                    & 0x000F;
            if version != 0x000F {
                return None;
            }
            record_children_offsets(data, record_offset)
                .into_iter()
                .find_map(|child| find(data, child, shape_id, depth + 1))
        }

        find(data, record_offset, shape_id, 0)
    }

    fn malformed_officeart_child() -> Vec<u8> {
        // The unknown version-15 container has one child header whose
        // declared payload is longer than the bounded container payload.
        let mut malformed_inner = Vec::with_capacity(8);
        malformed_inner.extend_from_slice(&0x000F_u16.to_le_bytes());
        malformed_inner.extend_from_slice(&0x7FFD_u16.to_le_bytes());
        malformed_inner.extend_from_slice(&1_u32.to_le_bytes());
        raw_record(0x0F, 0, 0x7FFE, &malformed_inner)
    }

    fn mutate_batch_corruption_slide(
        slide: &[u8],
        early_shape_id: Option<u32>,
        malformed_shape_id: u32,
    ) -> super::Result<Vec<u8>> {
        fn rewrite_textbox_header(textbox: &[u8]) -> super::Result<Vec<u8>> {
            let mut data = Vec::with_capacity(super::drawing_payload(textbox)?.len());
            for child_result in super::children(textbox)? {
                let child = child_result?;
                if super::record_type(child)? == crate::RecordType::TextHeaderAtom as u16 {
                    let mut changed = child.to_vec();
                    changed[2..4]
                        .copy_from_slice(&(crate::RecordType::TextBytesAtom as u16).to_le_bytes());
                    data.extend_from_slice(&changed);
                } else {
                    data.extend_from_slice(child);
                }
            }
            super::rebuild(textbox, &data).map_err(Into::into)
        }

        fn rewrite_officeart(
            record: &[u8],
            early_shape_id: Option<u32>,
            malformed_shape_id: u32,
            malformed_child: &[u8],
            depth: usize,
        ) -> super::Result<Vec<u8>> {
            assert!(depth < 128, "test OfficeArt fixture exceeded its bound");
            if super::record_type(record)? == super::OFFICEART_SP_CONTAINER {
                let shape_id = super::shape_id_of(record)?;
                let mut data = Vec::with_capacity(
                    super::drawing_payload(record)?.len() + malformed_child.len(),
                );
                for child_result in super::children(record)? {
                    let child = child_result?;
                    if shape_id == early_shape_id
                        && super::record_type(child)? == super::OFFICEART_CLIENT_TEXTBOX
                    {
                        data.extend_from_slice(&rewrite_textbox_header(child)?);
                    } else {
                        data.extend_from_slice(child);
                    }
                }
                if shape_id == Some(malformed_shape_id) {
                    data.extend_from_slice(malformed_child);
                }
                return super::rebuild(record, &data).map_err(Into::into);
            }
            if super::record_version(record)? != 0xF {
                return Ok(record.to_vec());
            }
            let mut data = Vec::with_capacity(super::drawing_payload(record)?.len());
            for child_result in super::children(record)? {
                data.extend_from_slice(&rewrite_officeart(
                    child_result?,
                    early_shape_id,
                    malformed_shape_id,
                    malformed_child,
                    depth + 1,
                )?);
            }
            super::rebuild(record, &data).map_err(Into::into)
        }

        let malformed_child = malformed_officeart_child();
        let mut drawings = 0usize;
        let rewritten = super::rewrite_ppt_record(slide, 0, &mut |record| {
            if super::record_type(record)? != crate::RecordType::PPDrawing as u16 {
                return Ok(None);
            }
            drawings += 1;
            let data = rewrite_officeart(
                super::drawing_payload(record)?,
                early_shape_id,
                malformed_shape_id,
                &malformed_child,
                0,
            )?;
            super::rebuild(record, &data).map(Some).map_err(Into::into)
        })?;
        assert_eq!(drawings, 1);
        rewritten.ok_or(Error::Refused(Refusal::ShapeNotFound))
    }

    fn batch_corruption_fixture(early_shape_error: bool) -> (Vec<u8>, u32, u32) {
        let source = generated_fixture(1, 2);
        let first =
            super::resolve(&source, Target::new(Position::new(0), Position::new(0))).unwrap();
        let second =
            super::resolve(&source, Target::new(Position::new(0), Position::new(1))).unwrap();
        let drawing =
            find_nested_record_offset(&first.slide_record, 0, crate::RecordType::PPDrawing as u16)
                .unwrap();
        assert!(
            find_officeart_shape_offset(&first.slide_record, drawing, first.native_shape_id)
                .is_some()
        );
        assert!(
            find_officeart_shape_offset(&first.slide_record, drawing, second.native_shape_id)
                .is_some()
        );
        let slide = mutate_batch_corruption_slide(
            &first.slide_record,
            early_shape_error.then_some(first.native_shape_id),
            second.native_shape_id,
        )
        .unwrap();
        (slide, first.native_shape_id, second.native_shape_id)
    }

    struct PartialSink {
        bytes: Vec<u8>,
        limit: usize,
    }

    impl Write for PartialSink {
        fn write(&mut self, input: &[u8]) -> io::Result<usize> {
            if self.bytes.len() >= self.limit {
                return Err(io::Error::other("intentional partial sink failure"));
            }
            let count = (self.limit - self.bytes.len()).min(input.len());
            self.bytes.extend_from_slice(&input[..count]);
            Ok(count)
        }

        fn flush(&mut self) -> io::Result<()> {
            Ok(())
        }
    }

    fn durable_limits() -> litchi_core::patch::PatchLimits {
        litchi_core::patch::PatchLimits::new(
            litchi_core::patch::BlobLimits::new(0, 0, 0),
            1024 * 1024,
            1,
            8,
            256 * 1024,
            512 * 1024,
        )
    }

    #[test]
    fn batch_text_budget_accepts_exact_limit_and_rejects_one_over_and_overflow() {
        assert!(super::require_batch_text_budget(5, 7, 12).is_ok());
        assert!(matches!(
            super::require_batch_text_budget(5, 8, 12),
            Err(Error::Package(crate::package::Error::ResourceLimit(message)))
                if message.contains("13") && message.contains("12")
        ));
        assert!(matches!(
            super::require_batch_text_budget(usize::MAX, 1, usize::MAX),
            Err(Error::Package(crate::package::Error::ResourceLimit(message)))
                if message.contains("overflow")
        ));
    }

    #[test]
    fn batch_preserves_early_text_error_before_later_shape_resolution_error() {
        let source = Arc::from(fixture("abc").into_boxed_slice());
        let missing_shape = Target::new(Position::new(0), Position::new(1));
        let error = super::replace_shape_texts_batch(
            source,
            &[(target(), "\0"), (missing_shape, "replacement")],
            usize::MAX,
            None,
        )
        .err()
        .expect("a missing later shape must leave the earlier text error observable");
        assert!(matches!(
            error,
            Error::Refused(Refusal::IncompatibleEncoding)
        ));
    }

    #[test]
    fn batch_index_replays_legacy_error_for_malformed_requested_sibling() {
        let (slide, first_shape_id, second_shape_id) = batch_corruption_fixture(false);
        let requested = [first_shape_id, second_shape_id]
            .into_iter()
            .collect::<BTreeSet<_>>();

        let legacy = super::inspect_slide(&slide, first_shape_id)
            .expect_err("legacy first-target inspection must reject malformed sibling framing");
        assert!(matches!(
            &legacy,
            Error::Package(crate::package::Error::Corrupted(message))
                if message.contains("truncated")
        ));

        let indexed = super::index_slide_text_targets(&slide, &requested);
        assert!(matches!(
            &indexed,
            Err(Error::Package(crate::package::Error::Corrupted(message)))
                if message.contains("truncated")
        ));
        let batch = match indexed {
            Err(_index_error) => super::inspect_slide(&slide, first_shape_id),
            Ok(index) => super::inspect_indexed_shape_text(&index, first_shape_id),
        }
        .expect_err("the batch admission path must reject");
        assert_eq!(format!("{batch:?}"), format!("{legacy:?}"));
    }

    #[test]
    fn batch_index_preserves_earlier_target_error_before_sibling_framing_error() {
        let (slide, first_shape_id, second_shape_id) = batch_corruption_fixture(true);
        let requested = [first_shape_id, second_shape_id]
            .into_iter()
            .collect::<BTreeSet<_>>();

        let legacy = super::inspect_slide(&slide, first_shape_id)
            .expect_err("legacy first-target inspection must reject the malformed textbox");
        assert!(matches!(legacy, Error::Refused(Refusal::MultipleTextAtoms)));

        let indexed = super::index_slide_text_targets(&slide, &requested);
        assert!(matches!(
            &indexed,
            Err(Error::Package(crate::package::Error::Corrupted(message)))
                if message.contains("truncated")
        ));
        let batch = match indexed {
            Err(_index_error) => super::inspect_slide(&slide, first_shape_id),
            Ok(index) => super::inspect_indexed_shape_text(&index, first_shape_id),
        }
        .expect_err("the batch admission path must reject");
        assert!(matches!(batch, Error::Refused(Refusal::MultipleTextAtoms)));
    }

    #[test]
    fn source_checked_text_edit_round_trips_and_reverses() {
        let source = Snapshot::from_bytes(fixture("abc")).unwrap();
        let target = target();
        let source_slide = super::resolve(source.bytes(), target).unwrap();
        let mut edit = source.edit_text(target).unwrap();
        assert_eq!(edit.text(), "abc");
        edit.set_text("xyz").unwrap();
        let commit = edit.commit().unwrap();
        let committed_slide = super::resolve(commit.snapshot().bytes(), target).unwrap();
        assert_eq!(
            source_slide.slide_record.len(),
            committed_slide.slide_record.len()
        );
        let changed = source_slide
            .slide_record
            .iter()
            .zip(committed_slide.slide_record.iter())
            .filter(|(before, after)| before != after)
            .count();
        assert_eq!(changed, 3);
        let mut package =
            Reader::from_reader(Cursor::new(commit.snapshot().bytes().to_vec())).unwrap();
        let presentation = package.presentation().unwrap();
        assert_eq!(
            presentation.slides().unwrap()[0].shapes().unwrap()[0]
                .text()
                .unwrap(),
            "xyz"
        );
        let undone = commit.patch().inverse().apply(commit.snapshot()).unwrap();
        assert_eq!(undone.bytes(), source.bytes());
        assert!(commit.patch().apply(&undone).is_ok());
    }

    #[test]
    fn source_backed_splice_reopens_semantics_and_preserves_unrelated_stream() {
        let bytes = with_stream(fixture("abc"), &["UnrelatedPayload"], b"opaque bytes");
        let source = Arc::new(OwnedSource::new(bytes.clone()));
        let snapshot = super::SourceSnapshot::open(source).unwrap();
        let mut transaction = snapshot.edit_text(target()).unwrap();
        transaction.set_text("xyz").unwrap();
        let commit = transaction.commit().unwrap();

        assert_eq!(commit.snapshot().edit_text(target()).unwrap().text(), "xyz");
        assert_eq!(commit.snapshot().read_text(target()).unwrap(), "xyz");
        assert_eq!(commit.snapshot().read_text(target()).unwrap(), "xyz");
        assert_eq!(commit.diagnostics().replacement_bytes(), 3);
        assert!(!commit.is_noop());

        let mut output = Vec::new();
        let report = commit.write_to(&mut output).unwrap();
        assert_eq!(report.output_bytes(), u64::try_from(bytes.len()).unwrap());

        let mut original =
            OleFile::open(Cursor::new(bytes), litchi_cfb::OleFileLimits::default()).unwrap();
        let mut published =
            OleFile::open(Cursor::new(output), litchi_cfb::OleFileLimits::default()).unwrap();
        assert_eq!(
            published.open_stream(&["UnrelatedPayload"]).unwrap(),
            original.open_stream(&["UnrelatedPayload"]).unwrap()
        );
        assert_eq!(
            published
                .open_stream(&["PowerPoint Document"])
                .unwrap()
                .len(),
            original
                .open_stream(&["PowerPoint Document"])
                .unwrap()
                .len()
        );

        let restored = commit.patch().inverse().apply(commit.snapshot()).unwrap();
        assert_eq!(restored.edit_text(target()).unwrap().text(), "abc");
    }

    #[test]
    fn source_backed_noop_reuses_source_identity_and_is_exact() {
        let bytes = fixture("abc");
        let snapshot =
            super::SourceSnapshot::open(Arc::new(OwnedSource::new(bytes.clone()))).unwrap();
        let commit = snapshot.edit_text(target()).unwrap().commit().unwrap();

        assert!(commit.is_noop());
        assert!(commit.patch().is_empty());
        assert!(Arc::ptr_eq(&snapshot.inner, &commit.snapshot().inner));

        let mut output = Vec::new();
        commit.write_to(&mut output).unwrap();
        assert_eq!(output, bytes);
        let applied = commit.patch().apply(&snapshot).unwrap();
        assert!(Arc::ptr_eq(&snapshot.inner, &applied.inner));
    }

    #[test]
    fn source_backed_transaction_refuses_length_changes_without_staging() {
        let snapshot =
            super::SourceSnapshot::open(Arc::new(OwnedSource::new(fixture("abc")))).unwrap();
        let mut transaction = snapshot.edit_text(target()).unwrap();
        assert!(matches!(
            transaction.set_text("longer"),
            Err(Error::Refused(Refusal::LengthChange))
        ));
        assert_eq!(transaction.text(), "abc");
        let commit = transaction.commit().unwrap();
        assert!(commit.is_noop());
    }

    #[test]
    fn source_backed_splice_reports_partial_sink_progress() {
        let snapshot =
            super::SourceSnapshot::open(Arc::new(OwnedSource::new(fixture("abc")))).unwrap();
        let mut transaction = snapshot.edit_text(target()).unwrap();
        transaction.set_text("xyz").unwrap();
        let commit = transaction.commit().unwrap();
        let mut sink = PartialSink {
            bytes: Vec::new(),
            limit: 11,
        };
        let error = commit.write_to(&mut sink).unwrap_err();
        assert!(matches!(
            error,
            Error::Source(OverlayError::IncompleteOutput {
                progress: OutputProgress::Prefix { accepted, .. },
                ..
            }) if accepted == 11
        ));
        assert_eq!(sink.bytes.len(), 11);
    }

    #[test]
    fn source_backed_limits_accept_exact_splice_budget_and_reject_one_over() {
        let bytes = fixture("abc");
        let exact_source = Arc::new(OwnedSource::new(bytes.clone()));
        let exact_options = super::SourceBackedOptions {
            splice_limits: StreamSpliceLimits::new(1, 1, 3, 65_536, 1024).unwrap(),
            ..super::SourceBackedOptions::default()
        };
        let exact = super::SourceSnapshot::open_with_options(exact_source, exact_options).unwrap();
        let mut edit = exact.edit_text(target()).unwrap();
        edit.set_text("xyz").unwrap();
        assert!(!edit.commit().unwrap().is_noop());

        let over_source = Arc::new(OwnedSource::new(bytes));
        let over_options = super::SourceBackedOptions {
            splice_limits: StreamSpliceLimits::new(1, 1, 2, 65_536, 1024).unwrap(),
            ..super::SourceBackedOptions::default()
        };
        let over = super::SourceSnapshot::open_with_options(over_source, over_options).unwrap();
        let mut edit = over.edit_text(target()).unwrap();
        edit.set_text("xyz").unwrap();
        assert!(matches!(
            edit.commit(),
            Err(Error::Source(OverlayError::Unavailable { .. }))
        ));

        let limited_source = Arc::new(OwnedSource::new(fixture("abc")));
        let limited_options = super::SourceBackedOptions {
            source_limits: crate::SourceLimits {
                max_aggregate_input_bytes: 1,
                ..crate::SourceLimits::default()
            },
            ..super::SourceBackedOptions::default()
        };
        assert!(matches!(
            super::SourceSnapshot::open_with_options(limited_source, limited_options),
            Err(Error::Package(crate::package::Error::Limit {
                kind: crate::SourceLimitKind::AggregateBytes,
                maximum: 1,
                ..
            }))
        ));
    }

    #[test]
    fn source_range_resolver_matches_owned_oracle_and_reads_selected_slide_ranges() {
        let bytes = generated_fixture(6, 4);
        let target = Target::new(Position::new(5), Position::new(3));
        let (owned, ranged, metrics) = resolve_owned_and_range(bytes.clone(), target);

        assert_eq!(owned.target, ranged.target);
        assert_eq!(owned.slide_persist_id, ranged.slide_persist_id);
        assert_eq!(owned.atom_offset, ranged.atom_offset);
        assert_eq!(owned.kind, ranged.kind);
        assert_eq!(owned.payload, ranged.payload);
        assert_eq!(owned.text, ranged.text);

        let mut ole =
            OleFile::open(Cursor::new(bytes), litchi_cfb::OleFileLimits::default()).unwrap();
        let document = ole.open_stream(&["PowerPoint Document"]).unwrap();
        let metrics = metrics.lock().unwrap();
        assert!(metrics.calls > 0);
        assert!(metrics.bytes < document.len());
        assert!(metrics.max_request < document.len());
    }

    #[test]
    fn source_range_resolver_matches_owned_oracle_for_every_shape() {
        let bytes = generated_fixture(6, 4);
        let mut package = Reader::from_reader(Cursor::new(bytes.clone())).unwrap();
        let presentation = package.presentation().unwrap();

        for (slide_index, slide) in presentation.slides().unwrap().iter().enumerate() {
            for (shape_index, shape) in slide.shapes().unwrap().iter().enumerate() {
                let target = Target::new(Position::new(slide_index), Position::new(shape_index));
                let (owned, ranged, _metrics) = resolve_owned_and_range(bytes.clone(), target);
                assert_eq!(ranged.target, target);
                assert_eq!(ranged.slide_persist_id, owned.slide_persist_id);
                assert_eq!(ranged.atom_offset, owned.atom_offset);
                assert_eq!(ranged.kind, owned.kind);
                assert_eq!(ranged.payload, owned.payload);
                assert_eq!(ranged.text, owned.text);
                assert_eq!(ranged.text, shape.text().unwrap());
            }
        }
    }

    #[test]
    fn source_backed_read_text_matches_eager_shape_text() {
        let bytes = generated_fixture(6, 4);
        let source =
            super::SourceSnapshot::open(Arc::new(OwnedSource::new(bytes.clone()))).unwrap();

        let mut package = Reader::from_reader(Cursor::new(bytes)).unwrap();
        let presentation = package.presentation().unwrap();
        let slides = presentation.slides().unwrap();
        for (slide_index, slide) in slides.iter().enumerate() {
            for (shape_index, shape) in slide.shapes().unwrap().iter().enumerate() {
                let target = Target::new(Position::new(slide_index), Position::new(shape_index));
                assert_eq!(
                    source.read_text(target).unwrap(),
                    shape.text().unwrap(),
                    "source-backed text differs at slide {slide_index}, shape {shape_index}"
                );
            }
        }
    }

    #[test]
    fn adopted_shared_owner_matches_a_fresh_positional_open() {
        let bytes = generated_fixture(6, 4);
        let opened =
            super::SourceSnapshot::open(Arc::new(OwnedSource::new(bytes.clone()))).unwrap();
        let ole = OleFile::open(
            Cursor::new(bytes.clone()),
            litchi_cfb::OleFileLimits::default(),
        )
        .unwrap();
        let shared = Arc::new(
            SharedOleFile::from_owned_ole_file(
                ole,
                SourceVersion::new(0x5050_5431, 0),
                SharedOleFileLimits::default(),
            )
            .unwrap(),
        );
        let adopted = super::SourceSnapshot::from_shared(
            Arc::clone(&shared),
            super::SourceBackedOptions::default(),
        )
        .unwrap();
        assert!(Arc::ptr_eq(&adopted.inner.shared, &shared));
        assert_eq!(adopted.fingerprint(), opened.fingerprint());
        assert_eq!(adopted.len(), opened.len());
        let target = Target::new(Position::new(5), Position::new(3));
        assert_eq!(
            adopted.read_text(target).unwrap(),
            opened.read_text(target).unwrap()
        );

        let package =
            crate::SourceBackedPackage::from_read_at(Arc::new(OwnedSource::new(bytes))).unwrap();
        let from_package = package.text_edit_snapshot().unwrap();
        assert_eq!(from_package.fingerprint(), opened.fingerprint());
        assert_eq!(
            from_package.read_text(target).unwrap(),
            opened.read_text(target).unwrap()
        );

        let limited = super::SourceBackedOptions {
            source_limits: crate::SourceLimits {
                max_package_bytes: 1,
                ..crate::SourceLimits::default()
            },
            ..super::SourceBackedOptions::default()
        };
        assert!(matches!(
            super::SourceSnapshot::from_shared(shared, limited),
            Err(Error::Package(crate::package::Error::Limit {
                kind: crate::SourceLimitKind::InputBytes,
                maximum: 1,
                ..
            }))
        ));
    }

    const fn assert_send_sync<T: Send + Sync>() {}
    // Risk R3 of unified-ops record 0003: the unified `Ppt` handle retains
    // these owners behind an `Arc`, so they must stay thread-safe.
    const _: () = {
        assert_send_sync::<super::SourceSnapshot>();
        assert_send_sync::<Snapshot>();
        assert_send_sync::<crate::SourceBackedPackage>();
        assert_send_sync::<crate::presentation::Presentation>();
    };

    #[test]
    fn repeated_source_reads_reuse_validated_cfb_index_and_reject_mutation() {
        let bytes = generated_fixture(6, 4);
        let (legacy_counting, legacy_metrics) = CountingSource::new(bytes.clone());
        let legacy_source: Arc<dyn ReadAt> = Arc::new(legacy_counting);
        let legacy_snapshot = super::SourceSnapshot::open(Arc::clone(&legacy_source)).unwrap();

        // This is the former per-query ownership shape: every selector opens
        // and validates a new SharedOleFile before resolving its bounded
        // ranges. Keep it in the regression so the I/O reduction remains
        // auditable without relying on a second worktree or a benchmark.
        CountingSource::reset(&legacy_metrics);
        let mut legacy_text = None;
        for _ in 0..2 {
            legacy_snapshot.ensure_current().unwrap();
            let shared = SharedOleFile::open(
                Arc::clone(&legacy_source),
                super::shared_limits(crate::SourceLimits::default()).unwrap(),
            )
            .unwrap();
            super::reject_macro_components(&shared).unwrap();
            let document_path =
                super::select_stream_path(&shared, &super::DOCUMENT_PATHS, "PowerPoint Document")
                    .unwrap();
            let current_user_path =
                super::select_stream_path(&shared, &super::CURRENT_USER_PATHS, "Current User")
                    .unwrap();
            super::ensure_matching_stream_topology(&document_path, &current_user_path).unwrap();
            let resolved = super::source_range::resolve_source_target(
                &shared,
                &document_path,
                &current_user_path,
                Target::new(Position::new(5), Position::new(3)),
                crate::SourceLimits::default(),
            )
            .unwrap();
            if let Some(expected) = &legacy_text {
                assert_eq!(&resolved.text, expected);
            } else {
                legacy_text = Some(resolved.text);
            }
            legacy_snapshot.ensure_current().unwrap();
        }
        let legacy_metrics = {
            let metrics = legacy_metrics.lock().unwrap();
            (metrics.calls, metrics.bytes)
        };

        let (counting, metrics) = CountingSource::new(bytes);
        let revision = Arc::clone(&counting.revision);
        let source = super::SourceSnapshot::open(Arc::new(counting)).unwrap();

        // Opening performs the one-time CFB catalog/FAT validation. The
        // semantic selector should then issue only its bounded stream-range
        // reads, with no second CFB index parse on a repeated query.
        CountingSource::reset(&metrics);
        let first = source
            .read_text(Target::new(Position::new(5), Position::new(3)))
            .unwrap();
        let first_metrics = {
            let metrics = metrics.lock().unwrap();
            (metrics.calls, metrics.bytes)
        };
        let second = source
            .read_text(Target::new(Position::new(5), Position::new(3)))
            .unwrap();
        let second_metrics = {
            let metrics = metrics.lock().unwrap();
            (metrics.calls, metrics.bytes)
        };
        assert_eq!(first, second);
        assert!(first_metrics.0 > 0);
        assert_eq!(second_metrics.0, first_metrics.0.saturating_mul(2));
        assert_eq!(second_metrics.1, first_metrics.1.saturating_mul(2));
        assert_eq!(legacy_text.as_deref(), Some(first.as_str()));
        assert_eq!(legacy_metrics, (74, 8_310));
        assert_eq!(second_metrics, (66, 3_190));
        assert!(legacy_metrics.0 > second_metrics.0);
        assert!(legacy_metrics.1 > second_metrics.1);

        revision.fetch_add(1, Ordering::AcqRel);
        assert!(matches!(
            source.read_text(Target::new(Position::new(5), Position::new(3))),
            Err(Error::Source(OverlayError::SourceChanged { .. }))
        ));
    }

    #[test]
    fn source_backed_read_text_preserves_stale_and_refusal_gates() {
        let (counting, _metrics) = CountingSource::new(fixture("abc"));
        let revision = Arc::clone(&counting.revision);
        let stale = super::SourceSnapshot::open(Arc::new(counting)).unwrap();
        revision.fetch_add(1, Ordering::AcqRel);
        assert!(matches!(
            stale.read_text(target()),
            Err(Error::Source(OverlayError::SourceChanged { .. }))
        ));

        let macro_source = super::SourceSnapshot::open(Arc::new(OwnedSource::new(with_stream(
            fixture("abc"),
            &["VBA"],
            b"macro bytes",
        ))))
        .unwrap();
        assert!(matches!(
            macro_source.read_text(target()),
            Err(Error::Refused(Refusal::UnsupportedSource))
        ));

        let topology_source = super::SourceSnapshot::open(Arc::new(OwnedSource::new(
            relocate_root_stream_to_dual(fixture("abc"), "Current User"),
        )))
        .unwrap();
        assert!(matches!(
            topology_source.read_text(target()),
            Err(Error::Refused(Refusal::UnsupportedDependency))
        ));

        let limited = super::SourceSnapshot::open_with_options(
            Arc::new(OwnedSource::new(fixture("abc"))),
            super::SourceBackedOptions {
                source_limits: crate::SourceLimits {
                    max_records: 1,
                    ..crate::SourceLimits::default()
                },
                ..super::SourceBackedOptions::default()
            },
        )
        .unwrap();
        assert!(matches!(
            limited.read_text(target()),
            Err(Error::Package(crate::package::Error::Limit {
                kind: crate::SourceLimitKind::Events,
                maximum: 1,
                ..
            }))
        ));

        let valid =
            super::SourceSnapshot::open(Arc::new(OwnedSource::new(fixture("abc")))).unwrap();
        assert!(matches!(
            valid.read_text(Target::new(Position::new(1), Position::new(0))),
            Err(Error::Refused(Refusal::SlideNotFound { position }))
                if position == Position::new(1)
        ));
    }

    #[test]
    fn source_range_rejects_hostile_macro_duplicate_trailing_and_cycle_inputs() {
        let macro_source = mutate_document_stream(fixture("abc"), |document, root, _end| {
            let doc_info = record_children_offsets(document, root)
                .into_iter()
                .find(|&offset| record_type_at(document, offset) == 2000)
                .unwrap();
            let vba_info = record_children_offsets(document, doc_info)
                .into_iter()
                .find(|&offset| record_type_at(document, offset) == 1023)
                .unwrap();
            let atom = record_children_offsets(document, vba_info)
                .into_iter()
                .find(|&offset| record_type_at(document, offset) == 1024)
                .unwrap();
            document[atom + 8..atom + 12].copy_from_slice(&41_u32.to_le_bytes());
            document[atom + 12..atom + 16].copy_from_slice(&1_u32.to_le_bytes());
        });
        let macro_snapshot =
            super::SourceSnapshot::open(Arc::new(OwnedSource::new(macro_source))).unwrap();
        assert!(matches!(
            macro_snapshot.edit_text(target()),
            Err(Error::Refused(Refusal::UnsupportedSource))
        ));

        let duplicate_source = mutate_document_stream(fixture("abc"), |document, root, _end| {
            let doc_info = record_children_offsets(document, root)
                .into_iter()
                .find(|&offset| record_type_at(document, offset) == 2000)
                .unwrap();
            let vba_info = record_children_offsets(document, doc_info)
                .into_iter()
                .find(|&offset| record_type_at(document, offset) == 1023)
                .unwrap();
            document[doc_info + 2..doc_info + 4]
                .copy_from_slice(&(crate::RecordType::SlideListWithText as u16).to_le_bytes());
            document[vba_info + 2..vba_info + 4]
                .copy_from_slice(&(crate::RecordType::SlideViewInfo as u16).to_le_bytes());
        });
        let duplicate_snapshot =
            super::SourceSnapshot::open(Arc::new(OwnedSource::new(duplicate_source))).unwrap();
        let duplicate_result = duplicate_snapshot.edit_text(target());
        assert!(matches!(
            duplicate_result,
            Err(Error::Package(crate::package::Error::Corrupted(message)))
                if message.contains("duplicate presentation SlideListWithTextContainer")
        ));

        let trailing_source = mutate_document_stream(fixture("abc"), |document, _root, _end| {
            let slide = (0..document.len().saturating_sub(8))
                .find(|&offset| record_type_at(document, offset) == crate::RecordType::Slide as u16)
                .unwrap();
            let drawing =
                find_nested_record_offset(document, slide, crate::RecordType::PPDrawing as u16)
                    .unwrap();
            let payload_len =
                u32::from_le_bytes(document[drawing + 4..drawing + 8].try_into().unwrap()) as usize;
            assert!(payload_len > 8);
            document[drawing + 4..drawing + 8]
                .copy_from_slice(&(u32::try_from(payload_len - 1).unwrap()).to_le_bytes());
        });
        let trailing_snapshot =
            super::SourceSnapshot::open(Arc::new(OwnedSource::new(trailing_source))).unwrap();
        let trailing_result = trailing_snapshot.edit_text(target());
        assert!(matches!(
            trailing_result,
            Err(Error::Package(crate::package::Error::Corrupted(message)))
                if message.contains("truncated")
                    || message.contains("trailing")
                    || message.contains("extends beyond")
        ));

        let source = fixture("abc");
        let mut ole = OleFile::open(
            Cursor::new(source.clone()),
            litchi_cfb::OleFileLimits::default(),
        )
        .unwrap();
        let current_user = ole.open_stream(&["Current User"]).unwrap();
        let edit_offset = usize::try_from(
            crate::CurrentUser::parse(&current_user, &crate::SourceLimits::default())
                .unwrap()
                .current_edit_offset(),
        )
        .unwrap();
        let cycle_source = mutate_document_stream(source, move |document, _root, _end| {
            document[edit_offset + 16..edit_offset + 20]
                .copy_from_slice(&(u32::try_from(edit_offset).unwrap()).to_le_bytes());
        });
        let cycle_snapshot =
            super::SourceSnapshot::open(Arc::new(OwnedSource::new(cycle_source))).unwrap();
        let cycle_result = cycle_snapshot.edit_text(target());
        assert!(matches!(
            cycle_result,
            Err(Error::Package(crate::package::Error::Corrupted(message)))
                if message.contains("cyclic or excessive UserEdit chain")
        ));
    }

    #[test]
    fn source_range_rejects_forged_historical_headers_and_directory_topology() {
        for (version, instance, payload_len) in [(1, 0, 28), (0, 1, 28), (0, 0, 27), (0, 0, 33)] {
            let result = resolve_synthetic_history(synthetic_history_source(
                version,
                instance,
                payload_len,
                None,
            ));
            assert!(matches!(
                result,
                Err(Error::Package(crate::package::Error::Corrupted(message)))
                    if message.contains("historical UserEditAtom header")
            ));
        }

        let result = resolve_synthetic_history(synthetic_history_source(0, 0, 28, Some(8)));
        assert!(matches!(
            result,
            Err(Error::Package(crate::package::Error::Corrupted(message)))
                if message.contains("PersistDirectoryAtom does not precede")
        ));

        let source = fixture("abc");
        let mut ole = OleFile::open(
            Cursor::new(source.clone()),
            litchi_cfb::OleFileLimits::default(),
        )
        .unwrap();
        let document = ole.open_stream(&["PowerPoint Document"]).unwrap();
        let current_user = ole.open_stream(&["Current User"]).unwrap();
        let edit_offset = usize::try_from(
            crate::CurrentUser::parse(&current_user, &crate::SourceLimits::default())
                .unwrap()
                .current_edit_offset(),
        )
        .unwrap();
        let directory_offset = usize::try_from(u32::from_le_bytes(
            document[edit_offset + 20..edit_offset + 24]
                .try_into()
                .unwrap(),
        ))
        .unwrap();
        for packed_header in [1_u16, 0x0010_u16] {
            let mut forged = document.clone();
            forged[directory_offset..directory_offset + 2]
                .copy_from_slice(&packed_header.to_le_bytes());
            let forged_source = with_stream(source.clone(), &["PowerPoint Document"], &forged);
            let snapshot =
                super::SourceSnapshot::open(Arc::new(OwnedSource::new(forged_source))).unwrap();
            assert!(matches!(
                snapshot.edit_text(target()),
                Err(Error::Package(crate::package::Error::Corrupted(message)))
                    if message.contains("invalid PersistDirectoryAtom header or shape")
            ));
        }
        let mut forged = document;
        forged[directory_offset + 4..directory_offset + 8].copy_from_slice(&2_u32.to_le_bytes());
        let forged_source = with_stream(source, &["PowerPoint Document"], &forged);
        let snapshot =
            super::SourceSnapshot::open(Arc::new(OwnedSource::new(forged_source))).unwrap();
        assert!(matches!(
            snapshot.edit_text(target()),
            Err(Error::Package(crate::package::Error::Corrupted(message)))
                if message.contains("invalid PersistDirectoryAtom header or shape")
        ));
    }

    #[test]
    fn source_backed_rejects_macros_and_protected_components() {
        for name in ["VBA", "VBAProject", "VbaProjectStg", "_VBA_PROJECT"] {
            let macro_source = super::SourceSnapshot::open(Arc::new(OwnedSource::new(
                with_stream(fixture("abc"), &[name], b"macro bytes"),
            )))
            .unwrap();
            assert!(matches!(
                macro_source.edit_text(target()),
                Err(Error::Refused(Refusal::UnsupportedSource))
            ));
        }

        let mut macro_atom_data = [0_u8; 12];
        macro_atom_data[0..4].copy_from_slice(&41_u32.to_le_bytes());
        macro_atom_data[4..8].copy_from_slice(&1_u32.to_le_bytes());
        macro_atom_data[8..12].copy_from_slice(&2_u32.to_le_bytes());
        let macro_atom = raw_record(
            2,
            0,
            crate::RecordType::VBAInfoAtom as u16,
            &macro_atom_data,
        );
        let macro_info = raw_record(0x0F, 1, crate::RecordType::VBAInfo as u16, &macro_atom);
        let doc_info = raw_record(0x0F, 0, crate::RecordType::DocInfoList as u16, &macro_info);
        let document = raw_record(0x0F, 0, crate::RecordType::Document as u16, &doc_info);
        assert!(matches!(
            super::reject_macro_records(&document, 0, crate::SourceLimits::default()),
            Err(Error::Refused(Refusal::UnsupportedSource))
        ));

        let mut empty_atom_data = [0_u8; 12];
        empty_atom_data[8..12].copy_from_slice(&2_u32.to_le_bytes());
        let empty_atom = raw_record(
            2,
            0,
            crate::RecordType::VBAInfoAtom as u16,
            &empty_atom_data,
        );
        let empty_info = raw_record(0x0F, 1, crate::RecordType::VBAInfo as u16, &empty_atom);
        let duplicate_info = raw_record(
            0x0F,
            0,
            crate::RecordType::DocInfoList as u16,
            &[empty_info.clone(), empty_info.clone()].concat(),
        );
        let document = raw_record(0x0F, 0, crate::RecordType::Document as u16, &duplicate_info);
        assert!(matches!(
            super::reject_macro_records(&document, 0, crate::SourceLimits::default()),
            Err(Error::Package(crate::package::Error::Corrupted(message)))
                if message.contains("multiple VBAInfo")
        ));

        let doc_info = raw_record(0x0F, 0, crate::RecordType::DocInfoList as u16, &empty_info);
        let document = raw_record(
            0x0F,
            0,
            crate::RecordType::Document as u16,
            &[doc_info.clone(), doc_info].concat(),
        );
        assert!(matches!(
            super::reject_macro_records(&document, 0, crate::SourceLimits::default()),
            Err(Error::Package(crate::package::Error::Corrupted(message)))
                if message.contains("multiple DocInfoList")
        ));

        let orphan_atom = raw_record(
            2,
            0,
            crate::RecordType::VBAInfoAtom as u16,
            &macro_atom_data,
        );
        let doc_info = raw_record(0x0F, 0, crate::RecordType::DocInfoList as u16, &orphan_atom);
        let document = raw_record(0x0F, 0, crate::RecordType::Document as u16, &doc_info);
        assert!(matches!(
            super::reject_macro_records(&document, 0, crate::SourceLimits::default()),
            Err(Error::Refused(Refusal::UnsupportedSource))
        ));

        let embedded_storage = super::SourceSnapshot::open(Arc::new(OwnedSource::new(
            with_storage(fixture("abc"), &["ObjectPool"]),
        )))
        .unwrap();
        assert!(matches!(
            embedded_storage.edit_text(target()),
            Err(Error::Refused(Refusal::UnsupportedDependency))
        ));

        let nested_dual_storage = super::SourceSnapshot::open(Arc::new(OwnedSource::new(
            with_storage(fixture("abc"), &["PP97_DUALSTORAGE", "PP97_DUALSTORAGE"]),
        )))
        .unwrap();
        assert!(matches!(
            nested_dual_storage.edit_text(target()),
            Err(Error::Refused(Refusal::UnsupportedDependency))
        ));

        assert!(matches!(
            super::SourceSnapshot::open(Arc::new(OwnedSource::new(with_storage(
                fixture("abc"),
                &["_SIGNATURES"],
            )))),
            Err(Error::Source(OverlayError::Ole(_)))
        ));

        let dependency =
            super::add_unmodeled_text_dependency_for_test(&fixture("abc"), target()).unwrap();
        let dependency_source =
            super::SourceSnapshot::open(Arc::new(OwnedSource::new(dependency))).unwrap();
        assert!(matches!(
            dependency_source.edit_text(target()),
            Err(Error::Refused(Refusal::UnsupportedDependency))
        ));
    }

    #[test]
    fn source_backed_rejects_invalid_text_header_enum() {
        let header = raw_record(
            0,
            0,
            crate::RecordType::TextHeaderAtom as u16,
            &3_u32.to_le_bytes(),
        );
        let text = raw_record(0, 0, crate::RecordType::TextBytesAtom as u16, b"abc");
        let textbox = raw_record(
            0x0F,
            0,
            super::OFFICEART_CLIENT_TEXTBOX,
            &[header, text].concat(),
        );
        assert!(matches!(
            super::inspect_source_textbox(&textbox, 0),
            Err(Error::Package(crate::package::Error::Corrupted(message)))
                if message.contains("TextTypeEnum")
        ));
    }

    #[test]
    fn source_backed_rejects_invalid_ppdrawing_and_officeart_headers() {
        for (version, instance) in [(0, 0), (0x0F, 1)] {
            let ppdrawing = raw_record(version, instance, crate::RecordType::PPDrawing as u16, &[]);
            let mut drawings = 0;
            let mut matches = Vec::new();
            assert!(matches!(
                super::find_source_drawing(
                    &ppdrawing,
                    0,
                    1,
                    0,
                    &mut drawings,
                    &mut matches,
                ),
                Err(Error::Package(crate::package::Error::Corrupted(message)))
                    if message.contains("PPDrawing")
            ));

            let shape = raw_record(2, 0, super::OFFICEART_SP, &1_u32.to_le_bytes());
            let container = raw_record(version, instance, super::OFFICEART_SP_CONTAINER, &shape);
            let mut matches = Vec::new();
            assert!(matches!(
                super::visit_source_officeart(&container, 0, 1, 0, &mut matches),
                Err(Error::Package(crate::package::Error::Corrupted(message)))
                    if message.contains("OfficeArt")
            ));
        }

        let shape = raw_record(2, 0, super::OFFICEART_SP, &1_u32.to_le_bytes());
        let textbox = raw_record(0, 0, super::OFFICEART_CLIENT_TEXTBOX, &[]);
        let container = raw_record(
            0x0F,
            0,
            super::OFFICEART_SP_CONTAINER,
            &[shape, textbox].concat(),
        );
        let mut matches = Vec::new();
        assert!(matches!(
            super::visit_source_officeart(&container, 0, 1, 0, &mut matches),
            Err(Error::Package(crate::package::Error::Corrupted(message)))
                if message.contains("OfficeArt")
        ));
    }

    #[test]
    fn source_backed_scans_every_officeart_root_and_rejects_trailing_bytes() {
        let root = source_officeart_shape(7, b"abc");
        let ppdrawing = raw_record(
            0x0F,
            0,
            crate::RecordType::PPDrawing as u16,
            &[root.clone(), root.clone()].concat(),
        );
        let mut drawings = 0;
        let mut matches = Vec::new();
        super::find_source_drawing(&ppdrawing, 0, 7, 0, &mut drawings, &mut matches).unwrap();
        assert_eq!(drawings, 1);
        assert_eq!(matches.len(), 2);

        let mut malformed = root;
        malformed.extend_from_slice(&[0_u8; 7]);
        assert!(matches!(
            super::visit_source_officeart_stream(&malformed, 0, 7, &mut Vec::new()),
            Err(Error::Package(crate::package::Error::Corrupted(message)))
                if message.contains("truncated")
        ));
    }

    #[test]
    fn source_backed_rejects_cross_topology_stream_pairs() {
        let source = relocate_root_stream_to_dual(fixture("abc"), "Current User");
        let snapshot = super::SourceSnapshot::open(Arc::new(OwnedSource::new(source))).unwrap();
        assert!(matches!(
            snapshot.edit_text(target()),
            Err(Error::Refused(Refusal::UnsupportedDependency))
        ));
    }

    #[test]
    fn supplied_shape_target_resolution_matches_standalone_resolution() {
        let bytes = fixture("abc");
        let target = target();
        let standalone = super::resolve(&bytes, target).unwrap();
        let shape_target = super::resolve_shape_target(&bytes, target).unwrap();
        let editor = crate::embedded::object::Editor::open_records(
            bytes.clone(),
            super::presentation(&bytes).unwrap().source_limits(),
        )
        .unwrap();
        let supplied = super::resolve_with_shape_target(target, shape_target, &editor).unwrap();

        assert_eq!(supplied.target, standalone.target);
        assert_eq!(supplied.slide_persist_id, standalone.slide_persist_id);
        assert_eq!(supplied.native_shape_id, standalone.native_shape_id);
        assert_eq!(supplied.slide_record, standalone.slide_record);
        assert_eq!(supplied.kind, standalone.kind);
        assert_eq!(supplied.payload, standalone.payload);
        assert_eq!(supplied.text, standalone.text);
        assert_eq!(supplied.can_resize, standalone.can_resize);
    }

    #[test]
    fn protected_source_refusal_precedes_target_resolution() {
        let source = Snapshot::from_bytes(with_storage(fixture("abc"), &["_SIGNATURES"]))
            .expect("signed PPT remains readable");
        let out_of_range = Target::new(Position::new(1), Position::new(0));

        assert!(matches!(
            source.edit_text(out_of_range),
            Err(Error::Refused(Refusal::UnsupportedSource))
        ));
    }

    #[test]
    fn length_changing_plain_text_edit_round_trips_and_reverses() {
        let source = Snapshot::from_bytes(fixture("abc")).unwrap();
        let target = target();
        let source_slide = super::resolve(source.bytes(), target).unwrap();
        let mut edit = source.edit_text(target).unwrap();
        edit.set_text("a longer plain text box").unwrap();
        let commit = edit.commit().unwrap();
        let committed_slide = super::resolve(commit.snapshot().bytes(), target).unwrap();
        assert_eq!(committed_slide.text, "a longer plain text box");
        assert!(committed_slide.slide_record.len() > source_slide.slide_record.len());
        let undone = commit.patch().inverse().apply(commit.snapshot()).unwrap();
        assert_eq!(undone.bytes(), source.bytes());
    }

    #[test]
    fn length_changing_utf16_text_updates_record_and_style_coverage() {
        let source = Snapshot::from_bytes(fixture("初稿")).unwrap();
        let target = target();
        let mut edit = source.edit_text(target).unwrap();
        edit.set_text("修订后的演示文稿\u{1f34a}").unwrap();
        let commit = edit.commit().unwrap();
        assert_eq!(
            super::resolve(commit.snapshot().bytes(), target)
                .unwrap()
                .text,
            "修订后的演示文稿\u{1f34a}"
        );
    }

    #[test]
    fn null_text_and_unmodeled_ranges_are_typed_refusals() {
        let source = Snapshot::from_bytes(fixture("abc")).unwrap();
        let mut edit = source.edit_text(target()).unwrap();
        assert!(matches!(
            edit.set_text("a\0c"),
            Err(Error::Refused(Refusal::IncompatibleEncoding))
        ));

        assert!(
            !super::resize_dependency_closure_is_modeled(
                &[
                    crate::RecordType::TextHeaderAtom as u16,
                    crate::RecordType::TextBytesAtom as u16,
                    crate::RecordType::StyleTextPropAtom as u16,
                    crate::RecordType::TextInteractiveInfoAtom as u16,
                ],
                None,
                None,
                3,
            )
            .unwrap()
        );
        // A repeated position-free record is not one modeled body either.
        assert!(
            !super::resize_dependency_closure_is_modeled(
                &[
                    crate::RecordType::TextHeaderAtom as u16,
                    crate::RecordType::TextBytesAtom as u16,
                    crate::RecordType::TextRulerAtom as u16,
                    crate::RecordType::TextRulerAtom as u16,
                ],
                None,
                None,
                3,
            )
            .unwrap()
        );
        assert!(
            super::resize_dependency_closure_is_modeled(
                &[
                    crate::RecordType::TextHeaderAtom as u16,
                    crate::RecordType::TextBytesAtom as u16,
                    crate::RecordType::TextRulerAtom as u16,
                ],
                None,
                None,
                3,
            )
            .unwrap()
        );
    }

    #[test]
    fn patch_rejects_a_different_source() {
        let source = Snapshot::from_bytes(fixture("abc")).unwrap();
        let target = target();
        let mut edit = source.edit_text(target).unwrap();
        edit.set_text("xyz").unwrap();
        let patch = edit.commit().unwrap().patch().clone();
        let other = Snapshot::from_bytes(fixture("def")).unwrap();
        assert!(patch.apply(&other).is_err());
    }

    #[test]
    fn durable_patch_round_trips_as_deterministic_json_and_applies() {
        let source = Snapshot::from_bytes(fixture("abc")).unwrap();
        let mut edit = source.edit_text(target()).unwrap();
        edit.set_text("a durable replacement").unwrap();
        let commit = edit.commit().unwrap();
        let durable = commit.patch().to_durable(durable_limits()).unwrap();
        let json = durable.to_deterministic_json().unwrap();
        let decoded =
            litchi_core::patch::Patch::<litchi_core::patch::Reversible>::from_deterministic_json(
                &json,
                durable_limits(),
            )
            .unwrap();

        let applied = source.apply_durable(&decoded).unwrap();
        assert_eq!(applied.bytes(), commit.snapshot().bytes());
        assert_eq!(
            super::resolve(applied.bytes(), target()).unwrap().text,
            "a durable replacement"
        );

        let restored = applied.apply_durable(&decoded.inverse()).unwrap();
        assert_eq!(
            super::resolve(restored.bytes(), target()).unwrap().text,
            "abc"
        );
        assert!(
            Snapshot::from_bytes(fixture("other"))
                .unwrap()
                .apply_durable(&decoded)
                .is_err()
        );
    }

    #[test]
    fn semantic_positions_are_checked_against_the_source() {
        let source = Snapshot::from_bytes(fixture("abc")).unwrap();
        assert!(matches!(
            source.edit_text(Target::new(Position::new(1), Position::new(0))),
            Err(Error::Refused(Refusal::SlideNotFound { position }))
                if position == Position::new(1)
        ));
        assert!(matches!(
            source.edit_text(Target::new(Position::new(0), Position::new(1))),
            Err(Error::Refused(Refusal::ShapeNotFound))
        ));
    }
}
