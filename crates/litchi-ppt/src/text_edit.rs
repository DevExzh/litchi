//! Source-checked, lossless edits of text already owned by a PPT shape.
//!
//! Length-changing replacement is supported when the selected
//! `TextCharsAtom` or `TextBytesAtom` has either no style record or a single
//! paragraph-formatting run and a single character-formatting run. The two
//! style coverage counts are then updated with the text atom and every owning
//! container length. `[MS-PPT]` associates other formatting, interaction,
//! bookmark, metacharacter, and special-information records with UTF-16
//! positions; their presence remains a typed dependency-closure refusal.

use std::collections::BTreeMap;
use std::fmt;
use std::io::{Cursor, Write};
use std::sync::Arc;

pub use litchi_core::Position;

use crate::consts::RecordType;
use crate::package::{Error as PackageError, Package, Result as PackageResult};
use crate::presentation::Presentation;
use crate::shapes::ShapeEnum;
use litchi_cfb::{
    ArtifactFingerprint, OleError, OverlayError, PublishReport, SameLengthStreamSplice,
    SharedOleFile, SharedOleFileLimits, StreamSpliceLimits, ValidatedOverlayPlan,
};
use litchi_core::{ReadAt, SourceVersion};
use litchi_ole_common::source_backed_overlay::SourceBackedOverlayPublisher;

mod source_range;

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
            Self::UnsupportedDependency => write!(
                formatter,
                "selected PPT shape text has an unsupported dependency record"
            ),
        }
    }
}

/// Error returned by the focused text-edit transaction.
#[derive(Debug)]
pub enum Error {
    /// The package could not be opened, parsed, or republished.
    Package(PackageError),
    /// The requested operation is not proven lossless for this source.
    Refused(Refusal),
    /// The immutable positional source or sequential overlay publisher
    /// rejected a source-backed operation.
    Source(OverlayError),
}

impl fmt::Display for Error {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Package(error) => error.fmt(formatter),
            Self::Refused(refusal) => refusal.fmt(formatter),
            Self::Source(error) => error.fmt(formatter),
        }
    }
}

impl std::error::Error for Error {
    fn source(&self) -> Option<&(dyn std::error::Error + 'static)> {
        match self {
            Self::Package(error) => Some(error),
            Self::Refused(_) => None,
            Self::Source(error) => Some(error),
        }
    }
}

impl From<PackageError> for Error {
    fn from(error: PackageError) -> Self {
        Self::Package(error)
    }
}

/// Result type for source-checked existing-shape text edits.
pub type Result<T> = std::result::Result<T, Error>;

/// Immutable whole-package snapshot used by [`Transaction`].
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Snapshot {
    bytes: Arc<[u8]>,
}

impl Snapshot {
    /// Opens an exact PPT artifact with default package limits.
    ///
    /// # Errors
    ///
    /// Returns an error when the complete package cannot be opened.
    pub fn parse(bytes: &[u8]) -> Result<Self> {
        Self::from_bytes(bytes.to_vec())
    }

    /// Captures an owned exact PPT artifact after validating that it opens.
    ///
    /// # Errors
    ///
    /// Returns an error when the complete package cannot be opened.
    pub fn from_bytes(bytes: Vec<u8>) -> Result<Self> {
        let _ = presentation(&bytes)?;
        Ok(Self {
            bytes: Arc::from(bytes.into_boxed_slice()),
        })
    }

    pub(crate) fn from_shared_bytes(bytes: Arc<[u8]>) -> Result<Self> {
        let _ = presentation_shared(bytes.clone())?;
        Ok(Self { bytes })
    }

    /// Exact bytes of the complete source or committed package artifact.
    #[must_use]
    pub fn bytes(&self) -> &[u8] {
        &self.bytes
    }

    /// Starts a source-checked transaction for one existing shape.
    ///
    /// # Errors
    ///
    /// Returns a typed refusal when the target cannot be resolved to one safe
    /// text atom.
    pub fn edit_text(&self, target: Target) -> Result<Transaction> {
        // Resolve the semantic selector without surfacing its result yet. The
        // editor preflight below remains the first observable failure, while
        // keeping the two allocation-heavy owners sequential.
        let shape_target = resolve_shape_target(&self.bytes, target);
        // The shared incremental owner rejects signed and encrypted CFB
        // envelopes before any candidate is staged. Surface that capability
        // boundary as a typed refusal instead of trying to normalize it.
        let editor = crate::embedded::object::Editor::open_records(self.bytes.to_vec())
            .map_err(|_error| Error::Refused(Refusal::UnsupportedSource))?;
        let resolved = resolve_with_shape_target(target, shape_target?, &editor)?;
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
    pub fn apply_durable<Mode>(&self, patch: &litchi_core::patch::Patch<Mode>) -> Result<Self> {
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
/// fallback. `record_limits` bounds the semantic metadata and selected slide
/// reads; `splice_limits` bounds the common CFB physical plan.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct SourceBackedOptions {
    /// Finite PPT record and stream limits.
    pub record_limits: crate::RecordLimits,
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

    /// Opens a positional source with explicit finite limits.
    pub fn open_with_limits(
        source: Arc<dyn ReadAt>,
        record_limits: crate::RecordLimits,
    ) -> Result<Self> {
        Self::open_with_options(
            source,
            SourceBackedOptions {
                record_limits,
                ..SourceBackedOptions::default()
            },
        )
    }

    /// Opens a positional source with explicit semantic and splice limits.
    pub fn open_with_options(
        source: Arc<dyn ReadAt>,
        options: SourceBackedOptions,
    ) -> Result<Self> {
        let shared_limits = shared_limits(options.record_limits)?;
        let length = source
            .len()
            .map_err(|error| Error::Source(OverlayError::Io(error)))?;
        if length > shared_limits.max_input_bytes() {
            return Err(PackageError::ResourceLimit(format!(
                "PPT source length {length} exceeds limit {}",
                shared_limits.max_input_bytes()
            ))
            .into());
        }
        let publisher =
            SourceBackedOverlayPublisher::open_with_limits(Arc::clone(&source), shared_limits)
                .map_err(Error::Source)?;
        let version = publisher.source_version().map_err(Error::Source)?;

        // An empty splice still runs the common full-artifact fingerprint and
        // complete composed-CFB reopen. Retaining this plan's source digest
        // gives the semantic patch a stronger foreign/stale-source guard than
        // a version token alone.
        let identity_plan = publisher
            .plan_splices(Vec::new(), options.splice_limits)
            .map_err(Error::Source)?;
        let fingerprint = identity_plan.source_fingerprint();
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
    pub fn edit_text(&self, target: Target) -> Result<SourceTransaction> {
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
            self.inner.options.record_limits,
        )?;
        self.ensure_current()?;
        Ok(SourceTransaction {
            source: self.clone(),
            resolved,
            replacement: None,
        })
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
pub struct SourceTransaction {
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
    pub fn text(&self) -> &str {
        &self.resolved.text
    }

    /// Selected semantic slide/shape position.
    #[must_use]
    pub const fn target(&self) -> Target {
        self.resolved.target
    }

    /// Stages a replacement only when its encoded atom length is unchanged.
    ///
    /// Failed validation leaves the transaction unchanged. The ordinary
    /// source-owning editor remains separate and is not used as a fallback.
    pub fn set_text(&mut self, value: impl Into<String>) -> Result<()> {
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
    pub fn commit(self) -> Result<SourceBackedCommit> {
        let replacement = self
            .replacement
            .unwrap_or_else(|| self.resolved.text.clone());
        publish_source_text(self.source, self.resolved, replacement)
    }

    /// Discards this transaction and returns its immutable source snapshot.
    #[must_use]
    pub fn rollback(self) -> SourceSnapshot {
        self.source
    }
}

/// Content-free diagnostics for one source-backed PPT text publication.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct SourceBackedDiagnostics {
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
    pub const fn target(self) -> Target {
        self.target
    }

    /// Number of physical CFB spans changed by the splice.
    #[must_use]
    pub const fn changed_spans(self) -> usize {
        self.changed_spans
    }

    /// Replacement bytes submitted to the CFB splice planner.
    #[must_use]
    pub const fn replacement_bytes(self) -> u64 {
        self.replacement_bytes
    }

    /// Complete source CFB length.
    #[must_use]
    pub const fn source_bytes(self) -> u64 {
        self.source_bytes
    }

    /// Exact source artifact fingerprint.
    #[must_use]
    pub const fn source_fingerprint(self) -> ArtifactFingerprint {
        self.source_fingerprint
    }

    /// Exact composed target artifact fingerprint.
    #[must_use]
    pub const fn target_fingerprint(self) -> ArtifactFingerprint {
        self.target_fingerprint
    }

    /// Source version checked before planning.
    #[must_use]
    pub const fn source_version(self) -> SourceVersion {
        self.source_version
    }

    /// Composed target version exposed by the positional overlay.
    #[must_use]
    pub const fn target_version(self) -> SourceVersion {
        self.target_version
    }
}

/// A published source-backed PPT target and its reusable overlay plan.
pub struct SourceBackedCommit {
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
    pub const fn snapshot(&self) -> &SourceSnapshot {
        &self.snapshot
    }

    /// Exact-source reversible patch.
    #[must_use]
    pub const fn patch(&self) -> &SourceBackedPatch {
        &self.patch
    }

    /// Common validated CFB overlay plan.
    #[must_use]
    pub const fn plan(&self) -> &ValidatedOverlayPlan {
        &self.plan
    }

    /// Content-free publication diagnostics.
    #[must_use]
    pub const fn diagnostics(&self) -> SourceBackedDiagnostics {
        self.diagnostics
    }

    /// Whether the selected replacement is an exact byte no-op.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.plan.is_noop()
    }

    /// Streams the complete validated target to a sequential sink.
    pub fn write_to<W: Write>(&self, writer: &mut W) -> Result<PublishReport> {
        self.plan.write_to(writer).map_err(Error::Source)
    }

    /// Alias emphasizing the forward-only publication boundary.
    pub fn publish_to_stream<W: Write>(&self, writer: &mut W) -> Result<PublishReport> {
        self.write_to(writer)
    }

    /// Publishes through the common synced sibling-temp/atomic-rename path.
    pub fn save<P: AsRef<std::path::Path>>(&self, path: P) -> Result<PublishReport> {
        self.plan.save(path).map_err(Error::Source)
    }
}

/// A source-checked reversible equal-length PPT text splice.
#[derive(Clone)]
pub struct SourceBackedPatch {
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
    pub const fn source(&self) -> &SourceSnapshot {
        &self.before
    }

    /// Exact target snapshot produced by forward application.
    #[must_use]
    pub const fn target(&self) -> &SourceSnapshot {
        &self.after
    }

    /// Exact source artifact fingerprint required for forward application.
    #[must_use]
    pub fn source_fingerprint(&self) -> ArtifactFingerprint {
        self.before.fingerprint()
    }

    /// Exact target artifact fingerprint produced by forward application.
    #[must_use]
    pub fn target_fingerprint(&self) -> ArtifactFingerprint {
        self.after.fingerprint()
    }

    /// Whether the patch changes no bytes.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.operation.is_none()
    }

    /// Applies only to the exact source version and fingerprint captured by
    /// this patch. The candidate is reopened and semantically read back again.
    pub fn apply(&self, current: &SourceSnapshot) -> Result<SourceSnapshot> {
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
    pub fn inverse(&self) -> Self {
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
    let shared_limits = shared_limits(source.inner.options.record_limits)?;
    let publisher = SourceBackedOverlayPublisher::open_with_limits(
        Arc::clone(&source.inner.source),
        shared_limits,
    )
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
    let shared_limits = shared_limits(source.inner.options.record_limits)?;
    let publisher = SourceBackedOverlayPublisher::open_with_limits(
        Arc::clone(&source.inner.source),
        shared_limits,
    )
    .map_err(Error::Source)?;
    let observed_version = publisher.source_version().map_err(Error::Source)?;
    if observed_version != source.source_version() {
        return Err(Error::Source(OverlayError::SourceChanged {
            expected: source.source_version(),
            observed: observed_version,
        }));
    }
    let plan = publisher
        .plan_splices(
            vec![SameLengthStreamSplice::new(
                operation.document_path.clone(),
                operation.atom_offset,
                Arc::clone(&operation.expected),
                Arc::clone(&operation.replacement),
            )],
            source.inner.options.splice_limits,
        )
        .map_err(Error::Source)?;
    if plan.source_fingerprint() != source.fingerprint() {
        return Err(Error::Source(OverlayError::SourceFingerprintChanged {
            expected: source.fingerprint(),
            observed: plan.source_fingerprint(),
        }));
    }

    let composed = plan.composed_source().map_err(Error::Source)?;
    let candidate: Arc<dyn ReadAt> = Arc::new(composed);
    let candidate_shared = SharedOleFile::open_with_limits(Arc::clone(&candidate), shared_limits)
        .map_err(map_ole_error)?;
    reject_macro_components(&candidate_shared)?;
    let candidate_resolved = source_range::resolve_source_target(
        &candidate_shared,
        &operation.document_path,
        &operation.current_user_path,
        operation.target,
        source.inner.options.record_limits,
    )?;
    verify_source_readback(&operation, &candidate_resolved)?;
    let target_version = candidate
        .version()
        .map_err(|error| Error::Source(OverlayError::Io(error)))?;
    let target = SourceSnapshot {
        inner: Arc::new(SourceInner {
            source: candidate,
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

fn shared_limits(record_limits: crate::RecordLimits) -> Result<SharedOleFileLimits> {
    let maximum = u64::try_from(record_limits.max_package_bytes).map_err(|_error| {
        PackageError::ResourceLimit("PPT package limit does not fit u64".into())
    })?;
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

fn open_shared_source(source: &SourceSnapshot) -> Result<SharedOleFile> {
    let limits = shared_limits(source.inner.options.record_limits)?;
    SharedOleFile::open_with_limits(Arc::clone(&source.inner.source), limits).map_err(map_ole_error)
}

impl SourceSnapshot {
    fn ensure_identity(&self) -> Result<()> {
        self.ensure_current()?;
        let shared_limits = shared_limits(self.inner.options.record_limits)?;
        let publisher = SourceBackedOverlayPublisher::open_with_limits(
            Arc::clone(&self.inner.source),
            shared_limits,
        )
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
    limits: crate::RecordLimits,
    label: &str,
) -> Result<Vec<u8>> {
    let refs: Vec<_> = path.iter().map(String::as_str).collect();
    let length = shared.stream_len(&refs).map_err(map_ole_error)?;
    let length_usize = usize::try_from(length).map_err(|_error| {
        PackageError::ResourceLimit(format!("{label} stream size exceeds this platform"))
    })?;
    if length_usize > limits.max_input_bytes {
        return Err(PackageError::ResourceLimit(format!(
            "{label} stream size {length_usize} exceeds limit {}",
            limits.max_input_bytes
        ))
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
    limits: crate::RecordLimits,
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
    limits: crate::RecordLimits,
) -> Result<SourceResolved> {
    let document_refs: Vec<_> = document_path.iter().map(String::as_str).collect();
    let current_user_refs: Vec<_> = current_user_path.iter().map(String::as_str).collect();
    let document_length = shared.stream_len(&document_refs).map_err(map_ole_error)?;
    let current_user_length = shared
        .stream_len(&current_user_refs)
        .map_err(map_ole_error)?;
    let aggregate_length = document_length
        .checked_add(current_user_length)
        .ok_or_else(|| PackageError::ResourceLimit("PPT metadata stream sizes overflow".into()))?;
    let aggregate_limit = u64::try_from(limits.max_aggregate_input_bytes).map_err(|_error| {
        PackageError::ResourceLimit("PPT aggregate input limit does not fit u64".into())
    })?;
    if aggregate_length > aggregate_limit {
        return Err(PackageError::ResourceLimit(format!(
            "PPT Document and Current User streams total {aggregate_length} bytes, exceeding limit {}",
            limits.max_aggregate_input_bytes
        ))
        .into());
    }
    let document = read_bounded_stream(shared, document_path, limits, "PowerPoint Document")?;
    let current_user = read_bounded_stream(shared, current_user_path, limits, "Current User")?;
    let current = crate::CurrentUser::parse_with_limits(&current_user, limits)?;
    if current.is_encrypted() {
        return Err(Error::Refused(Refusal::UnsupportedSource));
    }
    let mapping = crate::embedded::object::Editor::inspect_live_mapping(&document, &current_user)?;
    let directory = crate::slide::SlideDirectory::build_with_limits(
        &document,
        &current_user,
        &mapping,
        limits,
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
        crate::slide::SlideFactory::new_with_limits(&document, &mapping, &directory, limits);
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
    limits: crate::RecordLimits,
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
pub struct Transaction {
    source: Snapshot,
    resolved: Resolved,
    replacement: Option<String>,
}

impl Transaction {
    /// The exact text currently stored in the selected shape.
    #[must_use]
    pub fn text(&self) -> &str {
        &self.resolved.text
    }

    /// The selected semantic target.
    #[must_use]
    pub const fn target(&self) -> Target {
        self.resolved.target
    }

    /// Stages a replacement after proving that the dependent text ranges stay
    /// valid. Failed validation leaves the staged candidate untouched.
    ///
    /// # Errors
    ///
    /// Returns a typed refusal when the replacement has an incompatible
    /// encoding or when its changed length has unmodeled dependent ranges.
    pub fn set_text(&mut self, value: impl Into<String>) -> Result<()> {
        let candidate = value.into();
        let encoded = encode_replacement(&candidate, self.resolved.kind)?;
        if encoded.len() != self.resolved.payload.len() && !self.resolved.can_resize {
            return Err(Error::Refused(Refusal::DependencyClosure));
        }
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
    pub fn commit(self) -> Result<Commit> {
        let replacement = self
            .replacement
            .unwrap_or_else(|| self.resolved.text.clone());
        if replacement == self.resolved.text {
            let patch = Patch::new(self.source.clone(), self.source.clone(), None);
            return Ok(Commit {
                snapshot: self.source,
                patch,
                replaced_slide_persist_id: None,
            });
        }

        let encoded = encode_replacement(&replacement, self.resolved.kind)?;
        if encoded.len() != self.resolved.payload.len() && !self.resolved.can_resize {
            return Err(Error::Refused(Refusal::DependencyClosure));
        }
        let target_slide = rewrite_slide(
            &self.resolved.slide_record,
            self.resolved.native_shape_id,
            self.resolved.kind,
            &self.resolved.payload,
            &encoded,
        )?;
        let mut editor =
            crate::embedded::object::Editor::open_records(self.source.bytes.clone().to_vec())?;
        let live = editor.persisted_record(self.resolved.slide_persist_id)?;
        if live != self.resolved.slide_record.as_ref() {
            return Err(PackageError::Corrupted(
                "PPT shape-text transaction source changed before publication".into(),
            )
            .into());
        }
        editor.replace_persisted_record(self.resolved.slide_persist_id, target_slide)?;
        let bytes = editor.finish()?;
        let snapshot = Snapshot::from_bytes(bytes)?;
        let readback = resolve(&snapshot.bytes, self.resolved.target)?;
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
        let patch = Patch::new(self.source, snapshot.clone(), Some(change));
        Ok(Commit {
            snapshot,
            patch,
            replaced_slide_persist_id: Some(self.resolved.slide_persist_id),
        })
    }

    /// Discards this candidate without changing the source snapshot.
    #[must_use]
    pub fn rollback(self) -> Snapshot {
        self.source
    }
}

/// A published snapshot and its reversible source-checked patch.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Commit {
    snapshot: Snapshot,
    patch: Patch,
    replaced_slide_persist_id: Option<u32>,
}

impl Commit {
    /// The immutable committed package.
    #[must_use]
    pub const fn snapshot(&self) -> &Snapshot {
        &self.snapshot
    }

    /// The source-checked patch that produced this snapshot.
    #[must_use]
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Consumes the commit into its snapshot and patch.
    #[must_use]
    pub fn into_parts(self) -> (Snapshot, Patch) {
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
}

impl RootBatchPublication {
    pub(crate) fn source(&self) -> &[u8] {
        &self.source
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
pub struct Patch {
    before: Snapshot,
    after: Snapshot,
    change: Option<Change>,
}

impl Patch {
    fn new(before: Snapshot, after: Snapshot, change: Option<Change>) -> Self {
        Self {
            before,
            after,
            change,
        }
    }

    /// Exact source bytes required for forward application.
    #[must_use]
    pub fn before(&self) -> &[u8] {
        self.before.bytes()
    }

    /// Exact target bytes produced by forward application.
    #[must_use]
    pub fn after(&self) -> &[u8] {
        self.after.bytes()
    }

    /// Whether the patch changes no bytes.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.before.bytes() == self.after.bytes()
    }

    /// Applies only to the exact source artifact used to create this patch.
    ///
    /// # Errors
    ///
    /// Returns an error if `current` is not that exact source artifact.
    pub fn apply(&self, current: &Snapshot) -> Result<Snapshot> {
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
    pub fn inverse(&self) -> Self {
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
    pub fn to_durable(
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
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum TextKind {
    Bytes,
    Chars,
}

fn presentation(bytes: &[u8]) -> PackageResult<Presentation> {
    let mut package = Package::from_reader(Cursor::new(bytes.to_vec()))?;
    package.presentation()
}

fn presentation_shared(bytes: Arc<[u8]>) -> PackageResult<Presentation> {
    let mut package = Package::from_reader(Cursor::new(bytes))?;
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
    })
}

fn resolve_shape(bytes: &[u8], target: Target) -> Result<ShapeLocation> {
    let (persist_id, shape_id) = resolve_shape_target(bytes, target)?;
    let editor = crate::embedded::object::Editor::open_records(bytes.to_vec())?;
    resolve_shape_record(persist_id, shape_id, &editor)
}

fn resolve_shape_target(bytes: &[u8], target: Target) -> Result<(u32, u32)> {
    let presentation = presentation(bytes)?;
    resolve_shape_target_in_presentation(&presentation, target)
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

/// Replaces a bounded caller-preflighted set of semantic shape targets through
/// one persisted-record editor and one full candidate reopen.
///
/// Callers retain the hard operation-count bound. This owner resolves every
/// target and validates every encoding/dependency closure before staging any
/// record in the editor. Distinct shapes on one slide are rewritten against
/// the progressively staged slide record, while each rewrite retains the
/// original atom payload as its exact source precondition.
pub(crate) fn replace_shape_texts_batch(
    source: Arc<[u8]>,
    replacements: &[(Target, &str)],
    max_output_bytes: usize,
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

    let mut editor = crate::embedded::object::Editor::open_records_arc_with_limit(
        source.clone(),
        max_output_bytes,
    )?;
    let presentation = presentation_shared(source.clone())?;
    let slides = presentation.slides()?;
    let mut prepared = Vec::with_capacity(replacements.len());
    let mut original_slides = BTreeMap::<u32, Arc<[u8]>>::new();
    let mut encoded_bytes = 0usize;

    // Complete semantic resolution and replacement validation before the
    // editor receives its first staged mutation.
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
        let slide_record = match original_slides.get(&persist_id) {
            Some(record) => record.clone(),
            None => {
                let record: Arc<[u8]> =
                    Arc::from(editor.persisted_record(persist_id)?.into_boxed_slice());
                original_slides.insert(persist_id, record.clone());
                record
            },
        };
        let resolved = resolved_from_shared_location(*target, persist_id, shape_id, slide_record)?;
        let encoded_len = encoded_replacement_len(replacement, resolved.kind)?;
        encoded_bytes = encoded_bytes.checked_add(encoded_len).ok_or_else(|| {
            PackageError::ResourceLimit("PPT shape-text batch encoded byte count overflow".into())
        })?;
        require_batch_text_budget(replacement_bytes, encoded_bytes, max_output_bytes)?;
        let encoded = encode_replacement(replacement, resolved.kind)?;
        if encoded.len() != encoded_len {
            return Err(PackageError::Corrupted(
                "PPT shape-text batch replacement size preflight disagrees with encoding".into(),
            )
            .into());
        }
        if encoded.len() != resolved.payload.len() && !resolved.can_resize {
            return Err(Error::Refused(Refusal::DependencyClosure));
        }
        if *replacement != resolved.text {
            prepared.push((resolved, (*replacement).to_string(), encoded));
        }
    }
    drop(slides);
    drop(presentation);

    if prepared.is_empty() {
        return Ok(None);
    }

    // Input order is retained in the semantic change list. Each native slide
    // record is progressively rewritten in memory and staged in the editor
    // exactly once after all of its shape replacements succeed.
    let mut staged_slides = BTreeMap::<u32, Vec<u8>>::new();
    for (resolved, _replacement, encoded) in &prepared {
        let staged = staged_slides
            .entry(resolved.slide_persist_id)
            .or_insert_with(|| resolved.slide_record.to_vec());
        *staged = rewrite_slide(
            staged,
            resolved.native_shape_id,
            resolved.kind,
            &resolved.payload,
            encoded,
        )?;
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

    let candidate = editor.finish()?;
    crate::font::validate_unrelated_streams(&source, &candidate)?;
    let output: Arc<[u8]> = Arc::from(candidate.into_boxed_slice());

    // One complete reopen validates the publication and supplies all semantic
    // readbacks without reopening the package once per shape.
    let reopened = presentation_shared(output.clone())?;
    let reopened_slides = reopened.slides()?;
    let mut changes = Vec::with_capacity(prepared.len());
    for (resolved, replacement, _encoded) in prepared {
        let slide = reopened_slides
            .get(resolved.target.slide.get())
            .ok_or(Error::Refused(Refusal::SlideNotFound {
                position: resolved.target.slide,
            }))?;
        if slide.persist_id() != resolved.slide_persist_id {
            return Err(PackageError::Corrupted(
                "published PPT batch shape text resolved a different slide".into(),
            )
            .into());
        }
        let shape = slide
            .shapes()?
            .get(resolved.target.shape.get())
            .ok_or(Error::Refused(Refusal::ShapeNotFound))?;
        if shape.text()? != replacement {
            return Err(PackageError::Corrupted(
                "published PPT batch shape text did not round-trip through its source shape".into(),
            )
            .into());
        }
        changes.push(RootShapeTextChange {
            target: resolved.target,
            before: resolved.text,
            after: replacement,
            slide_persist_id: resolved.slide_persist_id,
        });
    }

    Ok(Some(RootBatchPublication {
        source,
        output,
        changes,
    }))
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
    let location = resolve_shape(bytes, target)?;
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
    let mut editor = crate::embedded::object::Editor::open_records(bytes.to_vec())?;
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

    let location = resolve_shape(bytes, target)?;
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
    let mut editor = crate::embedded::object::Editor::open_records(bytes.to_vec())?;
    editor.replace_persisted_record(location.persist_id, rewritten)?;
    editor.finish().map_err(Into::into)
}

fn native_shape_id(selected: &ShapeEnum<'_>) -> u32 {
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

#[derive(Debug)]
struct TextAtom {
    kind: TextKind,
    payload: Vec<u8>,
    text: String,
    can_resize: bool,
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
                    && anchor.replace(crate::Anchor::parse(child)?).is_some()
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
    let mut atom = None;
    let mut child_types = Vec::new();
    let mut style_record = None;
    for child_result in children(textbox)? {
        let child = child_result?;
        let child_type = record_type(child)?;
        child_types.push(child_type);
        if child_type == RecordType::StyleTextPropAtom as u16
            && style_record.replace(child).is_some()
        {
            return Err(Error::Refused(Refusal::DependencyClosure));
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
        });
    }
    let mut text_atom = atom.ok_or(Error::Refused(Refusal::NoTextAtom))?;
    text_atom.can_resize = resize_dependency_closure_is_modeled(
        &child_types,
        style_record,
        text_units(text_atom.kind, text_atom.payload.len())?,
    )?;
    Ok(text_atom)
}

fn resize_dependency_closure_is_modeled(
    child_types: &[u16],
    style_record: Option<&[u8]>,
    old_text_units: usize,
) -> Result<bool> {
    let text_header = RecordType::TextHeaderAtom as u16;
    let text_bytes = RecordType::TextBytesAtom as u16;
    let text_chars = RecordType::TextCharsAtom as u16;
    let style = RecordType::StyleTextPropAtom as u16;
    let supported_shape = matches!(
        child_types,
        [header, text]
            if *header == text_header && (*text == text_bytes || *text == text_chars)
    ) || matches!(
        child_types,
        [header, text, style_type]
            if *header == text_header
                && (*text == text_bytes || *text == text_chars)
                && *style_type == style
    );
    if !supported_shape {
        return Ok(false);
    }
    let Some(style_bytes) = style_record else {
        return Ok(true);
    };
    Ok(crate::StyleTextPropAtom::resize_single_run_record(
        style_bytes,
        old_text_units,
        old_text_units,
    )?
    .is_some())
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
    if depth >= 128 {
        return Err(
            PackageError::Corrupted("OfficeArt nesting exceeds text-edit limit".into()).into(),
        );
    }
    if record_type(record)? == OFFICEART_SP_CONTAINER && shape_id_of(record)? == Some(shape_id) {
        *matches = matches.saturating_add(1);
        return rewrite_shape_container(record, kind, before, after);
    }
    if record_version(record)? != 0xF {
        return Ok(record.to_vec());
    }
    let mut data = Vec::with_capacity(drawing_payload(record)?.len());
    for child_result in children(record)? {
        let child = child_result?;
        data.extend_from_slice(&rewrite_officeart(
            child,
            shape_id,
            kind,
            before,
            after,
            depth + 1,
            matches,
        )?);
    }
    rebuild(record, &data).map_err(Into::into)
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
    let mut count = 0usize;
    let mut data = Vec::with_capacity(drawing_payload(textbox)?.len());
    let old_text_units = text_units(kind, before.len())?;
    let new_text_units = text_units(kind, after.len())?;
    for child_result in children(textbox)? {
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
        } else {
            data.extend_from_slice(child);
        }
    }
    match count {
        0 => Err(Error::Refused(Refusal::NoTextAtom)),
        1 => rebuild(textbox, &data).map_err(Into::into),
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

fn shape_id_of(record: &[u8]) -> PackageResult<Option<u32>> {
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

fn decode_utf16(bytes: &[u8]) -> Result<String> {
    if !bytes.len().is_multiple_of(2) {
        return Err(Error::Refused(Refusal::IncompatibleEncoding));
    }
    let units = bytes
        .chunks_exact(2)
        .map(|pair| u16::from_le_bytes([pair[0], pair[1]]))
        .collect::<Vec<_>>();
    String::from_utf16(&units).map_err(|_err| Error::Refused(Refusal::IncompatibleEncoding))
}

fn record_type(record: &[u8]) -> PackageResult<u16> {
    let bytes: &[u8; 2] = record
        .get(2..4)
        .and_then(|value| value.try_into().ok())
        .ok_or_else(|| PackageError::Corrupted("truncated text-edit record header".into()))?;
    Ok(u16::from_le_bytes(*bytes))
}

fn record_version(record: &[u8]) -> PackageResult<u16> {
    let bytes: &[u8; 2] = record
        .get(..2)
        .and_then(|value| value.try_into().ok())
        .ok_or_else(|| PackageError::Corrupted("truncated text-edit record header".into()))?;
    Ok(u16::from_le_bytes(*bytes) & 0x000F)
}

fn record_instance(record: &[u8]) -> PackageResult<u16> {
    let bytes: &[u8; 2] = record
        .get(..2)
        .and_then(|value| value.try_into().ok())
        .ok_or_else(|| PackageError::Corrupted("truncated text-edit record header".into()))?;
    Ok(u16::from_le_bytes(*bytes) >> 4)
}

fn drawing_payload(record: &[u8]) -> PackageResult<&[u8]> {
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

fn children(record: &[u8]) -> PackageResult<impl Iterator<Item = PackageResult<&[u8]>>> {
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

fn rebuild(record: &[u8], data: &[u8]) -> PackageResult<Vec<u8>> {
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
    use crate::Package;
    use crate::writer::Writer;
    use litchi_cfb::{OleFile, OutputProgress, OverlayError, SharedOleFile, StreamSpliceLimits};
    use litchi_core::{OwnedSource, ReadAt, SourceVersion};
    use std::io::{self, Cursor, Write};
    use std::sync::{Arc, Mutex};

    fn fixture(text: &str) -> Vec<u8> {
        let mut writer = Writer::new();
        let slide = writer.add_slide().unwrap();
        writer.add_textbox(slide, 10, 10, 240, 40, text).unwrap();
        let mut output = Cursor::new(Vec::new());
        writer.write_to(&mut output).unwrap();
        output.into_inner()
    }

    fn target() -> Target {
        Target::new(Position::new(0), Position::new(0))
    }

    fn generated_fixture(slide_count: usize, shapes_per_slide: usize) -> Vec<u8> {
        let mut writer = Writer::new();
        for slide_index in 0..slide_count {
            let slide = writer.add_slide().unwrap();
            for shape_index in 0..shapes_per_slide {
                let x = 10 + i32::try_from(shape_index).unwrap() * 260;
                let text = format!("slide-{slide_index:02}-shape-{shape_index:02}");
                writer.add_textbox(slide, x, 10, 240, 40, &text).unwrap();
            }
        }
        let mut output = Cursor::new(Vec::new());
        writer.write_to(&mut output).unwrap();
        output.into_inner()
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
        version: SourceVersion,
    }

    impl CountingSource {
        fn new(bytes: Vec<u8>) -> (Self, Arc<Mutex<ReadMetrics>>) {
            let metrics = Arc::new(Mutex::new(ReadMetrics::default()));
            let source = Self {
                bytes: Arc::new(bytes),
                metrics: Arc::clone(&metrics),
                version: SourceVersion::new(0x5050_5443, 0),
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
            Ok(self.version)
        }
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
        let shared = SharedOleFile::open(Arc::clone(&source)).unwrap();
        let document_path =
            super::select_stream_path(&shared, &super::DOCUMENT_PATHS, "PowerPoint Document")
                .unwrap();
        let current_user_path =
            super::select_stream_path(&shared, &super::CURRENT_USER_PATHS, "Current User").unwrap();
        let limits = crate::RecordLimits::default();
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
        let mut ole = OleFile::open(Cursor::new(source.to_vec())).unwrap();
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
            crate::RecordLimits::default(),
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
        let shared = SharedOleFile::open(Arc::new(OwnedSource::new(source))).unwrap();
        super::source_range::resolve_source_target(
            &shared,
            &["PowerPoint Document".to_string()],
            &["Current User".to_string()],
            target(),
            crate::RecordLimits::default(),
        )
        .map(|_| ())
    }

    fn with_storage(source: Vec<u8>, path: &[&str]) -> Vec<u8> {
        use litchi_cfb::{OleFile, OleWriter};

        let mut source_file = OleFile::open(Cursor::new(source)).unwrap();
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

        let mut source_file = OleFile::open(Cursor::new(source)).unwrap();
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

        let mut source_file = OleFile::open(Cursor::new(source)).unwrap();
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
            Package::from_reader(Cursor::new(commit.snapshot().bytes().to_vec())).unwrap();
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
        assert_eq!(commit.diagnostics().replacement_bytes(), 3);
        assert!(!commit.is_noop());

        let mut output = Vec::new();
        let report = commit.write_to(&mut output).unwrap();
        assert_eq!(report.bytes(), u64::try_from(bytes.len()).unwrap());

        let mut original = OleFile::open(Cursor::new(bytes)).unwrap();
        let mut published = OleFile::open(Cursor::new(output)).unwrap();
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
            record_limits: crate::RecordLimits {
                max_aggregate_input_bytes: 1,
                ..crate::RecordLimits::default()
            },
            ..super::SourceBackedOptions::default()
        };
        let limited =
            super::SourceSnapshot::open_with_options(limited_source, limited_options).unwrap();
        assert!(matches!(
            limited.edit_text(target()),
            Err(Error::Package(crate::package::Error::ResourceLimit(_)))
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

        let mut ole = OleFile::open(Cursor::new(bytes)).unwrap();
        let document = ole.open_stream(&["PowerPoint Document"]).unwrap();
        let metrics = metrics.lock().unwrap();
        assert!(metrics.calls > 0);
        assert!(metrics.bytes < document.len());
        assert!(metrics.max_request < document.len());
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
        let mut ole = OleFile::open(Cursor::new(source.clone())).unwrap();
        let current_user = ole.open_stream(&["Current User"]).unwrap();
        let edit_offset = usize::try_from(
            crate::CurrentUser::parse(&current_user)
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
        let mut ole = OleFile::open(Cursor::new(source.clone())).unwrap();
        let document = ole.open_stream(&["PowerPoint Document"]).unwrap();
        let current_user = ole.open_stream(&["Current User"]).unwrap();
        let edit_offset = usize::try_from(
            crate::CurrentUser::parse(&current_user)
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
            super::reject_macro_records(&document, 0, crate::RecordLimits::default()),
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
            super::reject_macro_records(&document, 0, crate::RecordLimits::default()),
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
            super::reject_macro_records(&document, 0, crate::RecordLimits::default()),
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
            super::reject_macro_records(&document, 0, crate::RecordLimits::default()),
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
        let editor = crate::embedded::object::Editor::open_records(bytes.clone()).unwrap();
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
