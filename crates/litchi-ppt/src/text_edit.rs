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
use std::io::Cursor;
use std::sync::Arc;

pub use litchi_core::Position;

use crate::consts::RecordType;
use crate::package::{Error as PackageError, Package, Result as PackageResult};
use crate::presentation::Presentation;
use crate::shapes::ShapeEnum;

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
}

impl fmt::Display for Error {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Package(error) => error.fmt(formatter),
            Self::Refused(refusal) => refusal.fmt(formatter),
        }
    }
}

impl std::error::Error for Error {
    fn source(&self) -> Option<&(dyn std::error::Error + 'static)> {
        match self {
            Self::Package(error) => Some(error),
            Self::Refused(_) => None,
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
    use std::io::Cursor;

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
